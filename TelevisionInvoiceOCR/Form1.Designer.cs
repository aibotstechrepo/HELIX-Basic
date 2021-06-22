namespace TelevisionInvoiceOCR
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label8 = new System.Windows.Forms.Label();
            this.btnClearField = new System.Windows.Forms.Button();
            this.LblStatus = new System.Windows.Forms.Label();
            this.btnPDFFileBrowse = new System.Windows.Forms.Button();
            this.btnFolderBrowse = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txtOutputFolder = new System.Windows.Forms.TextBox();
            this.txtPDFFile = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnExtract = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.txtExcelName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.cbOutputType = new System.Windows.Forms.ComboBox();
            this.winPDFViewer = new AxAcroPDFLib.AxAcroPDF();
            this.label5 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.progressBar2 = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.winPDFViewer)).BeginInit();
            this.SuspendLayout();
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.label8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label8.Font = new System.Drawing.Font("Cambria", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label8.Location = new System.Drawing.Point(16, 9);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(716, 45);
            this.label8.TabIndex = 30;
            this.label8.Text = "INTELLIGENT DOCUMENT EXTRACTOR";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnClearField
            // 
            this.btnClearField.BackColor = System.Drawing.Color.Silver;
            this.btnClearField.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClearField.Location = new System.Drawing.Point(413, 226);
            this.btnClearField.Name = "btnClearField";
            this.btnClearField.Size = new System.Drawing.Size(130, 35);
            this.btnClearField.TabIndex = 42;
            this.btnClearField.Text = "CLEAR";
            this.btnClearField.UseVisualStyleBackColor = false;
            this.btnClearField.Click += new System.EventHandler(this.button4_Click);
            // 
            // LblStatus
            // 
            this.LblStatus.AutoSize = true;
            this.LblStatus.BackColor = System.Drawing.Color.Transparent;
            this.LblStatus.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblStatus.ForeColor = System.Drawing.Color.Black;
            this.LblStatus.Location = new System.Drawing.Point(85, 267);
            this.LblStatus.Name = "LblStatus";
            this.LblStatus.Size = new System.Drawing.Size(64, 15);
            this.LblStatus.TabIndex = 41;
            this.LblStatus.Text = "Completed";
            this.LblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnPDFFileBrowse
            // 
            this.btnPDFFileBrowse.BackColor = System.Drawing.Color.Silver;
            this.btnPDFFileBrowse.Font = new System.Drawing.Font("Cambria", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPDFFileBrowse.Location = new System.Drawing.Point(631, 93);
            this.btnPDFFileBrowse.Name = "btnPDFFileBrowse";
            this.btnPDFFileBrowse.Size = new System.Drawing.Size(101, 26);
            this.btnPDFFileBrowse.TabIndex = 37;
            this.btnPDFFileBrowse.Text = "Browse";
            this.btnPDFFileBrowse.UseVisualStyleBackColor = false;
            this.btnPDFFileBrowse.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // btnFolderBrowse
            // 
            this.btnFolderBrowse.BackColor = System.Drawing.Color.Silver;
            this.btnFolderBrowse.Font = new System.Drawing.Font("Cambria", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFolderBrowse.Location = new System.Drawing.Point(630, 134);
            this.btnFolderBrowse.Name = "btnFolderBrowse";
            this.btnFolderBrowse.Size = new System.Drawing.Size(102, 26);
            this.btnFolderBrowse.TabIndex = 36;
            this.btnFolderBrowse.Text = "Browse";
            this.btnFolderBrowse.UseVisualStyleBackColor = false;
            this.btnFolderBrowse.Click += new System.EventHandler(this.button2_Click);
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Cambria", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(12, 257);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(83, 30);
            this.label3.TabIndex = 38;
            this.label3.Text = " Status :";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtOutputFolder
            // 
            this.txtOutputFolder.Font = new System.Drawing.Font("Courier New", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOutputFolder.Location = new System.Drawing.Point(183, 134);
            this.txtOutputFolder.Name = "txtOutputFolder";
            this.txtOutputFolder.Size = new System.Drawing.Size(423, 26);
            this.txtOutputFolder.TabIndex = 35;
            this.txtOutputFolder.Text = "a";
            // 
            // txtPDFFile
            // 
            this.txtPDFFile.Font = new System.Drawing.Font("Courier New", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPDFFile.Location = new System.Drawing.Point(183, 92);
            this.txtPDFFile.Name = "txtPDFFile";
            this.txtPDFFile.Size = new System.Drawing.Size(424, 26);
            this.txtPDFFile.TabIndex = 34;
            this.txtPDFFile.Text = "a";
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Cambria", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(12, 133);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(166, 27);
            this.label2.TabIndex = 33;
            this.label2.Text = "Output Location   :";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Cambria", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(12, 93);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(165, 22);
            this.label1.TabIndex = 32;
            this.label1.Text = "PDF File                   : ";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnExtract
            // 
            this.btnExtract.BackColor = System.Drawing.Color.Silver;
            this.btnExtract.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExtract.Location = new System.Drawing.Point(246, 226);
            this.btnExtract.Name = "btnExtract";
            this.btnExtract.Size = new System.Drawing.Size(130, 35);
            this.btnExtract.TabIndex = 31;
            this.btnExtract.Text = "EXTRACT";
            this.btnExtract.UseVisualStyleBackColor = false;
            this.btnExtract.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // txtExcelName
            // 
            this.txtExcelName.Font = new System.Drawing.Font("Courier New", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtExcelName.Location = new System.Drawing.Point(182, 175);
            this.txtExcelName.Name = "txtExcelName";
            this.txtExcelName.Size = new System.Drawing.Size(298, 26);
            this.txtExcelName.TabIndex = 46;
            this.txtExcelName.Text = "a";
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Cambria", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(12, 172);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(166, 27);
            this.label4.TabIndex = 45;
            this.label4.Text = "Excel name             :";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cbOutputType
            // 
            this.cbOutputType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOutputType.Font = new System.Drawing.Font("Courier New", 12F);
            this.cbOutputType.IntegralHeight = false;
            this.cbOutputType.ItemHeight = 18;
            this.cbOutputType.Items.AddRange(new object[] {
            ".XLSX",
            ".XLS"});
            this.cbOutputType.Location = new System.Drawing.Point(486, 175);
            this.cbOutputType.Name = "cbOutputType";
            this.cbOutputType.Size = new System.Drawing.Size(121, 26);
            this.cbOutputType.TabIndex = 47;
            // 
            // winPDFViewer
            // 
            this.winPDFViewer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.winPDFViewer.Enabled = true;
            this.winPDFViewer.Location = new System.Drawing.Point(779, 7);
            this.winPDFViewer.Name = "winPDFViewer";
            this.winPDFViewer.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("winPDFViewer.OcxState")));
            this.winPDFViewer.Size = new System.Drawing.Size(720, 304);
            this.winPDFViewer.TabIndex = 48;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.Font = new System.Drawing.Font("Cambria", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(708, 282);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(34, 40);
            this.label5.TabIndex = 44;
            this.label5.Text = "%";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.Font = new System.Drawing.Font("Cambria", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.Black;
            this.label7.Location = new System.Drawing.Point(658, 282);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(44, 40);
            this.label7.TabIndex = 49;
            this.label7.Text = "100";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // progressBar2
            // 
            this.progressBar2.Location = new System.Drawing.Point(12, 290);
            this.progressBar2.Name = "progressBar2";
            this.progressBar2.Size = new System.Drawing.Size(640, 23);
            this.progressBar2.TabIndex = 50;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::Helix_Basic.Properties.Resources.ocr_blur_background;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1508, 323);
            this.Controls.Add(this.progressBar2);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.winPDFViewer);
            this.Controls.Add(this.cbOutputType);
            this.Controls.Add(this.txtExcelName);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnClearField);
            this.Controls.Add(this.LblStatus);
            this.Controls.Add(this.btnPDFFileBrowse);
            this.Controls.Add(this.btnFolderBrowse);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtOutputFolder);
            this.Controls.Add(this.txtPDFFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExtract);
            this.Controls.Add(this.label8);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(1524, 362);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AIBOTS OCR";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.winPDFViewer)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label8;
        public static System.Windows.Forms.Label lblProgress1;
        private System.Windows.Forms.Button btnClearField;
        private System.Windows.Forms.Label LblStatus;
        private System.Windows.Forms.Button btnPDFFileBrowse;
        private System.Windows.Forms.Button btnFolderBrowse;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtOutputFolder;
        private System.Windows.Forms.TextBox txtPDFFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnExtract;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TextBox txtExcelName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.ComboBox cbOutputType;
        private AxAcroPDFLib.AxAcroPDF winPDFViewer;
        public static System.Windows.Forms.Label label6;
        public static System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ProgressBar progressBar2;
    }
}

