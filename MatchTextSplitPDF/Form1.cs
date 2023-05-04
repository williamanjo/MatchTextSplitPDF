using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;
using iTextSharp.text.pdf.parser;
using System.Threading.Tasks;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace MatchTextSplitPDF
{
    public partial class Form1 : Form
    {
        private CancellationTokenSource cts = new CancellationTokenSource();

        public Form1()
        {
            InitializeComponent();
            OpenExcel.Select();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void OpenExcel_Click(object sender, EventArgs e)
        {
            comboBoxExcelSheetNames.Items.Clear();
            if (openExcelDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                ExcelPath.Text = openExcelDialog.FileName;
                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(@"" + ExcelPath.Text)))
                {
                    foreach (var worksheet in xlPackage.Workbook.Worksheets)
                    {
                        comboBoxExcelSheetNames.Items.Add(worksheet.Name);
                    }
                }
                comboBoxExcelSheetNames.SelectedIndex = 0;
            }
            if (ExcelPath.Text != "" && PdfPath.Text != "" && FolderSavePath.Text != "") { Start.Enabled = true; }

        }

        private void comboBoxExcelSheetNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filePath = ExcelPath.Text;
            string sheetName = comboBoxExcelSheetNames.SelectedItem.ToString();

            comboBox1.Items.Clear();
            comboBox2.Items.Clear();

            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(filePath)))
            {

                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[sheetName];
                var columnsUsed = worksheet.Cells["1:1"].Where(c => c.Value != null).Select(c => c.Start.Column);

                foreach (var columnNumber in columnsUsed)
                {
                    string columnLetter = ExcelCellAddress.GetColumnLetter(columnNumber);
                    comboBox1.Items.Add(columnLetter);
                }
                comboBox2.Items.AddRange(comboBox1.Items.Cast<object>().ToArray());

                if (comboBox1.Items.Count >= 2)
                {
                    comboBox1.SelectedIndex = 0;
                    comboBox2.SelectedIndex = 1;
                }
                else if(comboBox1.Items.Count == 1)
                {
                    comboBox1.SelectedIndex = 0;
                }
            }
        }

        private void SplitInFolders_CheckedChanged(object sender, EventArgs e)
        {
            if (SplitInFolders.Checked)
            {

                comboBox2.Visible = true; ColunaFoldersLabel.Visible = true;

            }
            else
            {

                comboBox2.Visible = false; ColunaFoldersLabel.Visible = false;
            }
        }

        private void OpenPDF_Click(object sender, EventArgs e)
        {
            if (openPdfDialog.ShowDialog() == DialogResult.OK) { 
                PdfPath.Text = openPdfDialog.FileName;  
            }
            if (ExcelPath.Text != "" && PdfPath.Text != "" && FolderSavePath.Text != "") { Start.Enabled = true; }
        }

        private void OpenSaveFolder_Click(object sender, EventArgs e)
        {
            using (CommonOpenFileDialog dialog = new CommonOpenFileDialog())
            {

                dialog.IsFolderPicker = true;
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    FolderSavePath.Text = dialog.FileName;
                }
            }
            if (ExcelPath.Text != "" && PdfPath.Text != "" && FolderSavePath.Text != "") { Start.Enabled = true; }
        }
        private void BlockUI(bool block)
        {

            OpenExcel.Enabled = block;
            comboBoxExcelSheetNames.Enabled = block;
            comboBox1.Enabled = block;
            comboBox2.Enabled = block;
            checkBox1.Enabled = block;
            SplitInFolders.Enabled = block;
            OpenPDF.Enabled = block;
            OpenSaveFolder.Enabled = block;



        }

        private async void Start_Click(object sender, EventArgs e)
        {
            if (Start.Text != "Cancelar")
            {
                WriteLog("-----------------Starting----------------", Log);
                BlockUI(false);
                Start.Text = "Cancelar";

                try
                {
                    using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(@"" + ExcelPath.Text)))
                    {
                        var cancellationTokenSource = new CancellationTokenSource();
                        var cancellationToken = cancellationTokenSource.Token;

                        var myWorksheet = xlPackage.Workbook.Worksheets[comboBoxExcelSheetNames.SelectedIndex];
                        var totalRows = myWorksheet.Dimension.End.Row;
                        var totalColumns = myWorksheet.Dimension.End.Column;

                        progressBar1.Maximum = checkBox1.Checked ? totalRows - 2 : totalRows - 1;
                        progressBar1.Value = 0;

                        // Funções de atualização da UI.
                        Action<int> updateProgress = value =>
                        {
                            progressBar1.Invoke((Action)(() => progressBar1.Value = value));
                        };

                        Action<string> writeLog = log =>
                        {
                            WriteLog(log, Log);
                        };

                        int comboBox1SelectedIndex = comboBox1.SelectedIndex + 1;
                        int comboBox2SelectedIndex = comboBox2.SelectedIndex + 1;
                        bool splitInFoldersChecked = SplitInFolders.Checked;
                        string pdfPath = PdfPath.Text;
                        string folderSavePath = FolderSavePath.Text;


                        await Task.Run(() =>
                        {
                            DoSearch(myWorksheet, cancellationToken, updateProgress, writeLog, comboBox1SelectedIndex, comboBox2SelectedIndex, splitInFoldersChecked, pdfPath, folderSavePath).Wait();
                        }, cancellationToken);
                    }
                }
                catch (Exception ex)
                {
                    WriteLog(ex.Message, Log);
                }
                finally
                {
                    BlockUI(true);
                    Start.Text = "Start";
                    if (OpenOnFinal.Checked && !cts.IsCancellationRequested ) { System.Diagnostics.Process.Start("explorer.exe", FolderSavePath.Text); }
                }
            }
            else
            {
                WriteLog("Cancelando.....", Log);
                Start.Text = "Start";
                cts.Cancel();

            }

        }


        private async Task DoSearch(ExcelWorksheet myWorksheet, CancellationToken cancellationToken, Action<int> updateProgress, Action<string> writeLog, int comboBox1SelectedIndex, int comboBox2SelectedIndex, bool splitInFoldersChecked, string pdfPath, string folderSavePath)
        {
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(@"" + ExcelPath.Text)))
            {

                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;

                int rowNum = (checkBox1.Checked ? 2 : 1);


                await Task.Run(() =>
            {
                try
                {
                    while (rowNum <= totalRows && !cancellationToken.IsCancellationRequested)
                    {
                        cts.Token.ThrowIfCancellationRequested();
                        if (Convert.ToString(myWorksheet.Cells[rowNum, comboBox1SelectedIndex].Value) != "-" && myWorksheet.Cells[rowNum, comboBox1SelectedIndex].Value != null && Convert.ToString(myWorksheet.Cells[rowNum, comboBox1SelectedIndex].Value) != "#N/D")
                        {
                            var item = new ExcelData
                            {
                                SearchText = (Convert.ToString(myWorksheet.Cells[rowNum, comboBox1SelectedIndex].Value)),
                                Folder = splitInFoldersChecked ? (Convert.ToString(myWorksheet.Cells[rowNum, comboBox2SelectedIndex].Value)) : null
                            };
                            WriteLog(Environment.NewLine + $"Procurando ({item.SearchText}) em {System.IO.Path.GetFileName(pdfPath)}", Log);

                            var Pag = SearchPdfText(pdfPath, item.SearchText);

                            if (Pag >= 1)
                            {
                                WriteLog($" Encontrado na Pagina {{{Pag}}}", Log);

                                var Folder = splitInFoldersChecked ? (folderSavePath + @"\" + item.Folder) : folderSavePath;

                                CreateFolder(Folder);

                                WriteLog($"Extraindo Pagina para {Folder + @"\" + item.SearchText + ".pdf"}", Log);
                                WriteLog(Environment.NewLine + "----------------------------------------", Log);
                                ExtractPage(pdfPath, Folder + @"\" + item.SearchText + ".pdf", Pag);
                            }
                            else
                            {
                                WriteLog($" Não encontrado ", Log);
                            }
                        }
                        updateProgress(rowNum - (checkBox1.Checked ? 2 : 1));
                        rowNum++;
                    }
                }
                catch (System.OperationCanceledException)
                {

                    ResetCancelationToken();
                    WriteLog("Cancelado.....", Log);
                }
            }, cts.Token);

            }
        }
        private void ResetCancelationToken()
        {
            cts?.Dispose();
            cts = new CancellationTokenSource();
        }

        private void CreateFolder(string path)
        {

            if (!Directory.Exists(path)) // verifica se a pasta não existe
            {
                Directory.CreateDirectory(path); // cria a pasta
            }

        }


        private int SearchPdfText(string pdfFilePath, string searchText)
        {
            using (PdfReader reader = new PdfReader(pdfFilePath))
            {
                for (int pageNum = 1; pageNum <= reader.NumberOfPages; pageNum++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentPageText = PdfTextExtractor.GetTextFromPage(reader, pageNum, strategy);
                    if (currentPageText.Contains(searchText))
                    {
                        return pageNum;
                    }
                }
            }
            return -1;
        }

        public void ExtractPage(string sourcePdfPath, string outputPdfPath, int pageNumber)
        {
            PdfReader reader = new PdfReader(sourcePdfPath);

            // check if the page number is valid
            if (pageNumber < 1 || pageNumber > reader.NumberOfPages)
            {
                throw new ArgumentOutOfRangeException(nameof(pageNumber), "Page number is out of range.");
            }

            // create a new document with the selected page
            Document document = new Document(reader.GetPageSizeWithRotation(pageNumber));
            PdfCopy copy = new PdfCopy(document, new FileStream(outputPdfPath, FileMode.Create));
            document.Open();
            copy.AddPage(copy.GetImportedPage(reader, pageNumber));

            // save the new document
            document.Close();
            reader.Close();
        }

        public void WriteLog(string log, System.Windows.Forms.TextBox tb)
        {
            tb.Invoke((Action)(() => tb.AppendText(log + Environment.NewLine)));
        }

    }
}
