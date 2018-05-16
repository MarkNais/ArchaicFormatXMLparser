using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;

namespace XMLParse1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Removes the file from a directory-path string.
        /// </summary>
        public static string DirFromFullPath(string path)
        {
            return path.Substring(0, path.LastIndexOf("\\"));
        }

        private void btn_Remind_Header_Click(object sender, EventArgs e)
        {
            ExcelHeaderReuirements formReminder = new ExcelHeaderReuirements();
            formReminder.ShowDialog();
        }

        private void btn_Browse_Click(object sender, EventArgs e)
        {
            OpenFileDialog fbd = new OpenFileDialog();
            DialogResult dial = fbd.ShowDialog();
            if (dial == DialogResult.OK)
            {
                if (fbd.CheckPathExists == false) MessageBox.Show("Path doesn't exist!");
                else tbPath.Text = fbd.FileName;
            }
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            if (bw_XML.IsBusy) bw_XML.CancelAsync();
            if (bw_Excel.IsBusy) bw_Excel.CancelAsync();
            btn_Run_XML.Enabled = true;
            btn_Create_Excel.Enabled = true;
        }

        private void btn_Run_XML_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbPath.Text))  MessageBox.Show("Select a valid file.");
            else
            {
                string fileExt = Path.GetExtension(tbPath.Text).ToLower();
                if (fileExt != ".xls" && fileExt != ".xlsx") MessageBox.Show("File's extension is not XLS or XLSX!");
                else
                {
                    btn_Run_XML.Enabled = false;
                    btn_Cancel.Enabled = true;
                    bw_XML.RunWorkerAsync();
                }
            }
        }

        private void btn_Create_Excel_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbPath.Text)) MessageBox.Show("Select a valid file.");
            else
            {
                string fileExt = Path.GetExtension(tbPath.Text).ToLower();
                if (fileExt != ".xml" && fileExt != ".txt") MessageBox.Show("File's extension is not XML or TXT!");
                else
                {
                    btn_Create_Excel.Enabled = false;
                    btn_Cancel.Enabled = true;
                    bw_Excel.RunWorkerAsync();
                }
            }
        }

        private void bw_XML_DoWork(object sender, DoWorkEventArgs e)
        {
            string targetFile = tbPath.Text;
            string outFile = DirFromFullPath(targetFile) + "//" + tbOutputName.Text + ".txt";
            ApptusXMLencode.transformExceltoXML(targetFile, outFile, !chk_Concise_XML.Checked, !chk_XML_Min_Data.Checked);
        }

        private void bw_XML_RunComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            btn_Run_XML.Enabled = true;
        }

        private void bw_Excel_DoWork(object sender, DoWorkEventArgs e)
        {
            string targetFile = tbPath.Text;
            string outFile = DirFromFullPath(targetFile) + "//" + tbOutputName.Text + ".xlsx";
            ApttusExcelEncode.transformXMLtoExcel(targetFile, outFile);
        }

        private void bw_Excel_RunComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            btn_Create_Excel.Enabled = true;
        }
    }
}
