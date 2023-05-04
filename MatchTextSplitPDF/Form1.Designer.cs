namespace MatchTextSplitPDF
{
    partial class Form1
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ColunaFoldersLabel = new System.Windows.Forms.Label();
            this.ColunaSearch = new System.Windows.Forms.Label();
            this.SplitInFolders = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBoxExcelSheetNames = new System.Windows.Forms.ComboBox();
            this.OpenExcel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.ExcelPath = new System.Windows.Forms.TextBox();
            this.openExcelDialog = new System.Windows.Forms.OpenFileDialog();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.OpenOnFinal = new System.Windows.Forms.CheckBox();
            this.OpenSaveFolder = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.FolderSavePath = new System.Windows.Forms.TextBox();
            this.OpenPDF = new System.Windows.Forms.Button();
            this.PDFLabel = new System.Windows.Forms.Label();
            this.PdfPath = new System.Windows.Forms.TextBox();
            this.openPdfDialog = new System.Windows.Forms.OpenFileDialog();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.Log = new System.Windows.Forms.TextBox();
            this.Start = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ColunaFoldersLabel);
            this.groupBox1.Controls.Add(this.ColunaSearch);
            this.groupBox1.Controls.Add(this.SplitInFolders);
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Controls.Add(this.comboBox2);
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.comboBoxExcelSheetNames);
            this.groupBox1.Controls.Add(this.OpenExcel);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.ExcelPath);
            this.groupBox1.Location = new System.Drawing.Point(13, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(702, 107);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "ExcelData";
            // 
            // ColunaFoldersLabel
            // 
            this.ColunaFoldersLabel.AutoSize = true;
            this.ColunaFoldersLabel.Location = new System.Drawing.Point(326, 50);
            this.ColunaFoldersLabel.Name = "ColunaFoldersLabel";
            this.ColunaFoldersLabel.Size = new System.Drawing.Size(77, 13);
            this.ColunaFoldersLabel.TabIndex = 10;
            this.ColunaFoldersLabel.Text = "ColunaFolders:";
            this.ColunaFoldersLabel.Visible = false;
            // 
            // ColunaSearch
            // 
            this.ColunaSearch.AutoSize = true;
            this.ColunaSearch.Location = new System.Drawing.Point(187, 50);
            this.ColunaSearch.Name = "ColunaSearch";
            this.ColunaSearch.Size = new System.Drawing.Size(77, 13);
            this.ColunaSearch.TabIndex = 9;
            this.ColunaSearch.Text = "ColunaSearch:";
            // 
            // SplitInFolders
            // 
            this.SplitInFolders.AutoSize = true;
            this.SplitInFolders.Location = new System.Drawing.Point(190, 75);
            this.SplitInFolders.Name = "SplitInFolders";
            this.SplitInFolders.Size = new System.Drawing.Size(117, 17);
            this.SplitInFolders.TabIndex = 8;
            this.SplitInFolders.Text = "Separar em pastas ";
            this.SplitInFolders.UseVisualStyleBackColor = true;
            this.SplitInFolders.CheckedChanged += new System.EventHandler(this.SplitInFolders_CheckedChanged);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(71, 75);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(101, 17);
            this.checkBox1.TabIndex = 7;
            this.checkBox1.Text = "Tem Cabeçalho";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(409, 46);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(36, 21);
            this.comboBox2.TabIndex = 6;
            this.comboBox2.Visible = false;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(272, 46);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(36, 21);
            this.comboBox1.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Sheet :";
            // 
            // comboBoxExcelSheetNames
            // 
            this.comboBoxExcelSheetNames.FormattingEnabled = true;
            this.comboBoxExcelSheetNames.Location = new System.Drawing.Point(71, 46);
            this.comboBoxExcelSheetNames.Name = "comboBoxExcelSheetNames";
            this.comboBoxExcelSheetNames.Size = new System.Drawing.Size(90, 21);
            this.comboBoxExcelSheetNames.TabIndex = 3;
            this.comboBoxExcelSheetNames.SelectedIndexChanged += new System.EventHandler(this.comboBoxExcelSheetNames_SelectedIndexChanged);
            // 
            // OpenExcel
            // 
            this.OpenExcel.Location = new System.Drawing.Point(612, 18);
            this.OpenExcel.Name = "OpenExcel";
            this.OpenExcel.Size = new System.Drawing.Size(75, 23);
            this.OpenExcel.TabIndex = 2;
            this.OpenExcel.Text = "Open";
            this.OpenExcel.UseVisualStyleBackColor = true;
            this.OpenExcel.Click += new System.EventHandler(this.OpenExcel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Excel :";
            // 
            // ExcelPath
            // 
            this.ExcelPath.Location = new System.Drawing.Point(71, 20);
            this.ExcelPath.Name = "ExcelPath";
            this.ExcelPath.ReadOnly = true;
            this.ExcelPath.Size = new System.Drawing.Size(535, 20);
            this.ExcelPath.TabIndex = 0;
            // 
            // openExcelDialog
            // 
            this.openExcelDialog.Filter = "Planilhas Excel (.xls, .xlsx, .xlsm)|.xls;*.xlsx;*.xlsm";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.OpenOnFinal);
            this.groupBox2.Controls.Add(this.OpenSaveFolder);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.FolderSavePath);
            this.groupBox2.Controls.Add(this.OpenPDF);
            this.groupBox2.Controls.Add(this.PDFLabel);
            this.groupBox2.Controls.Add(this.PdfPath);
            this.groupBox2.Location = new System.Drawing.Point(13, 127);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(702, 121);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "PdfData";
            // 
            // OpenOnFinal
            // 
            this.OpenOnFinal.AutoSize = true;
            this.OpenOnFinal.Location = new System.Drawing.Point(89, 95);
            this.OpenOnFinal.Name = "OpenOnFinal";
            this.OpenOnFinal.Size = new System.Drawing.Size(103, 17);
            this.OpenOnFinal.TabIndex = 9;
            this.OpenOnFinal.Text = "Abrir ao Concluir";
            this.OpenOnFinal.UseVisualStyleBackColor = true;
            // 
            // OpenSaveFolder
            // 
            this.OpenSaveFolder.Location = new System.Drawing.Point(612, 67);
            this.OpenSaveFolder.Name = "OpenSaveFolder";
            this.OpenSaveFolder.Size = new System.Drawing.Size(75, 23);
            this.OpenSaveFolder.TabIndex = 8;
            this.OpenSaveFolder.Text = "Open";
            this.OpenSaveFolder.UseVisualStyleBackColor = true;
            this.OpenSaveFolder.Click += new System.EventHandler(this.OpenSaveFolder_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(23, 72);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Salvar em :";
            // 
            // FolderSavePath
            // 
            this.FolderSavePath.Location = new System.Drawing.Point(89, 69);
            this.FolderSavePath.Name = "FolderSavePath";
            this.FolderSavePath.ReadOnly = true;
            this.FolderSavePath.Size = new System.Drawing.Size(517, 20);
            this.FolderSavePath.TabIndex = 6;
            // 
            // OpenPDF
            // 
            this.OpenPDF.Location = new System.Drawing.Point(612, 27);
            this.OpenPDF.Name = "OpenPDF";
            this.OpenPDF.Size = new System.Drawing.Size(75, 23);
            this.OpenPDF.TabIndex = 5;
            this.OpenPDF.Text = "Open";
            this.OpenPDF.UseVisualStyleBackColor = true;
            this.OpenPDF.Click += new System.EventHandler(this.OpenPDF_Click);
            // 
            // PDFLabel
            // 
            this.PDFLabel.AutoSize = true;
            this.PDFLabel.Location = new System.Drawing.Point(28, 33);
            this.PDFLabel.Name = "PDFLabel";
            this.PDFLabel.Size = new System.Drawing.Size(34, 13);
            this.PDFLabel.TabIndex = 4;
            this.PDFLabel.Text = "PDF :";
            // 
            // PdfPath
            // 
            this.PdfPath.Location = new System.Drawing.Point(71, 29);
            this.PdfPath.Name = "PdfPath";
            this.PdfPath.ReadOnly = true;
            this.PdfPath.Size = new System.Drawing.Size(535, 20);
            this.PdfPath.TabIndex = 3;
            // 
            // openPdfDialog
            // 
            this.openPdfDialog.FileName = "openPdfDialog";
            this.openPdfDialog.Filter = "Pdf (*.pdf)| *.pdf;";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.Log);
            this.groupBox3.Location = new System.Drawing.Point(12, 296);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(702, 211);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Log";
            // 
            // Log
            // 
            this.Log.Enabled = false;
            this.Log.Location = new System.Drawing.Point(6, 19);
            this.Log.Multiline = true;
            this.Log.Name = "Log";
            this.Log.ReadOnly = true;
            this.Log.Size = new System.Drawing.Size(690, 179);
            this.Log.TabIndex = 0;
            // 
            // Start
            // 
            this.Start.Enabled = false;
            this.Start.Location = new System.Drawing.Point(11, 262);
            this.Start.Name = "Start";
            this.Start.Size = new System.Drawing.Size(703, 23);
            this.Start.TabIndex = 3;
            this.Start.Text = "Start";
            this.Start.UseVisualStyleBackColor = true;
            this.Start.Click += new System.EventHandler(this.Start_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.progressBar1);
            this.groupBox4.Location = new System.Drawing.Point(12, 513);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(703, 65);
            this.groupBox4.TabIndex = 4;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Progress";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(16, 27);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(672, 23);
            this.progressBar1.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(639, 583);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(75, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Version : 1.0.0";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(727, 604);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.Start);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MatchTextSplitPDF";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox ExcelPath;
        private System.Windows.Forms.Button OpenExcel;
        private System.Windows.Forms.ComboBox comboBoxExcelSheetNames;
        private System.Windows.Forms.OpenFileDialog openExcelDialog;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label ColunaSearch;
        private System.Windows.Forms.CheckBox SplitInFolders;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label ColunaFoldersLabel;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button OpenPDF;
        private System.Windows.Forms.Label PDFLabel;
        private System.Windows.Forms.TextBox PdfPath;
        private System.Windows.Forms.Button OpenSaveFolder;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox FolderSavePath;
        private System.Windows.Forms.OpenFileDialog openPdfDialog;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button Start;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox Log;
        private System.Windows.Forms.CheckBox OpenOnFinal;
    }
}

