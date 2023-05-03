using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Windows.Forms;
using iTextSharp.text.pdf.parser;
using System.Collections.Generic;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using Microsoft.WindowsAPICodePack.Dialogs;


namespace MatchTextSplitPDF
{
    public partial class Form1 : Form
    {
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
            //if (textBox1.Text != "" && textBox4.Text != "" && textBox2.Text != "") { Start.Enabled = true; }

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

        }

        private void OpenSaveFolder_Click(object sender, EventArgs e)
        {
            using (CommonOpenFileDialog dialog = new CommonOpenFileDialog()) {

                dialog.IsFolderPicker = true;
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    FolderSavePath.Text =  dialog.FileName;
                }
            }
            
        }
    }
}
