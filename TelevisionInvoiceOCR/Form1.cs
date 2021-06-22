using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//KEY:  Y6goL-0onH0-ceTTst-OBbiK-AavnJ

namespace TelevisionInvoiceOCR
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        TempProcess a = new TempProcess();
         
        private void button1_Click_1(object sender, EventArgs e)
        {

            LblStatus.Text = "Extracting...";
            bool clear = false;
            if(txtPDFFile.Text!="")
            {
                if(txtOutputFolder.Text!="")
                {
                    if(txtExcelName.Text!="")
                    {
                        clear = true;
                    }
                    else
                    {
                        MessageBox.Show("Please enter the Output Excel file Name to Create", "Output Excel Name Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        txtExcelName.Focus();
                    }
                }
                else
                {
                    MessageBox.Show("Please enter the Output location to save the file", "Output Location Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtOutputFolder.Focus();
                }
            }
            else
            {
                MessageBox.Show("Please enter the PDF file name to Extract", "PDF File Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtPDFFile.Focus();
            } 
            if(clear==true)
            {
                //MessageBox.Show("All Details are correct","AIBOTS IOCR", MessageBoxButtons.OK, MessageBoxIcon.Information);
                progressBar2.Value = 5;

                a.Extract("", @txtOutputFolder.Text + "\\" + txtExcelName.Text + ".xlsx");
                MessageBox.Show("Data Extraction completed.", "AIBOTS IOCR", MessageBoxButtons.OK, MessageBoxIcon.Information);
                progressBar2.Value = 100;
                progressBar2.Value = 0;
                label7.Text = "0";
                LblStatus.Text = "Ready. Browse PDF to Extract";
            }


        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            txtPDFFile.Text = ""; 

            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Browse PDF";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "PDF file (*.pdf)|*.pdf";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txtPDFFile.Text = fdlg.FileName;
                    this.Size = new Size(1524, 362);
                    winPDFViewer.src = txtPDFFile.Text;
                }
                catch
                {
                    ;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK) 
            {
                txtOutputFolder.Text = folderDlg.SelectedPath;
                Environment.SpecialFolder root = folderDlg.RootFolder;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cbOutputType.SelectedIndex = 0;
            txtExcelName.Clear();
            txtOutputFolder.Clear();
            txtPDFFile.Clear();
            label7.Text = "0";
            LblStatus.Text = "Ready. Browse PDF to Extract";
            this.Size = new Size(762, 362);
            //762, 362

        }

        private void button4_Click(object sender, EventArgs e)
        {
            txtExcelName.Clear();
            txtOutputFolder.Clear();
            txtPDFFile.Clear();
            label7.Text = "0";
            LblStatus.Text = "Ready. Browse PDF to Extract";
            this.Size = new Size(762, 362);
        }
    }
}
