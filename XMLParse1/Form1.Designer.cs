namespace XMLParse1
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
            this.tbPath = new System.Windows.Forms.TextBox();
            this.btn_Run_XML = new System.Windows.Forms.Button();
            this.btn_Browse = new System.Windows.Forms.Button();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.btn_Remind_Header = new System.Windows.Forms.Button();
            this.chk_Concise_XML = new System.Windows.Forms.CheckBox();
            this.chk_XML_Min_Data = new System.Windows.Forms.CheckBox();
            this.group_XML = new System.Windows.Forms.GroupBox();
            this.group_Excel = new System.Windows.Forms.GroupBox();
            this.btn_Create_Excel = new System.Windows.Forms.Button();
            this.bw_XML = new System.ComponentModel.BackgroundWorker();
            this.bw_Excel = new System.ComponentModel.BackgroundWorker();
            this.lbl_Input_Path = new System.Windows.Forms.Label();
            this.lbl_Output_Name = new System.Windows.Forms.Label();
            this.tbOutputName = new System.Windows.Forms.TextBox();
            this.group_XML.SuspendLayout();
            this.group_Excel.SuspendLayout();
            this.SuspendLayout();
            // 
            // tbPath
            // 
            this.tbPath.Location = new System.Drawing.Point(78, 6);
            this.tbPath.Name = "tbPath";
            this.tbPath.Size = new System.Drawing.Size(405, 20);
            this.tbPath.TabIndex = 0;
            this.tbPath.Text = "C:\\Users\\MarkN\\Documents\\BUILDING CONSTRUCTION PRODUCTS.xlsx";
            // 
            // btn_Run_XML
            // 
            this.btn_Run_XML.Location = new System.Drawing.Point(36, 71);
            this.btn_Run_XML.Name = "btn_Run_XML";
            this.btn_Run_XML.Size = new System.Drawing.Size(75, 23);
            this.btn_Run_XML.TabIndex = 3;
            this.btn_Run_XML.Text = "Create XML";
            this.btn_Run_XML.UseVisualStyleBackColor = true;
            this.btn_Run_XML.Click += new System.EventHandler(this.btn_Run_XML_Click);
            // 
            // btn_Browse
            // 
            this.btn_Browse.Location = new System.Drawing.Point(489, 4);
            this.btn_Browse.Name = "btn_Browse";
            this.btn_Browse.Size = new System.Drawing.Size(75, 23);
            this.btn_Browse.TabIndex = 1;
            this.btn_Browse.Text = "Browse";
            this.btn_Browse.UseVisualStyleBackColor = true;
            this.btn_Browse.Click += new System.EventHandler(this.btn_Browse_Click);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Enabled = false;
            this.btn_Cancel.Location = new System.Drawing.Point(240, 201);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(75, 23);
            this.btn_Cancel.TabIndex = 6;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // btn_Remind_Header
            // 
            this.btn_Remind_Header.Location = new System.Drawing.Point(352, 30);
            this.btn_Remind_Header.Name = "btn_Remind_Header";
            this.btn_Remind_Header.Size = new System.Drawing.Size(209, 23);
            this.btn_Remind_Header.TabIndex = 7;
            this.btn_Remind_Header.Text = "Check Excel Header Requirements";
            this.btn_Remind_Header.UseVisualStyleBackColor = true;
            this.btn_Remind_Header.Click += new System.EventHandler(this.btn_Remind_Header_Click);
            // 
            // chk_Concise_XML
            // 
            this.chk_Concise_XML.AutoSize = true;
            this.chk_Concise_XML.Location = new System.Drawing.Point(19, 19);
            this.chk_Concise_XML.Name = "chk_Concise_XML";
            this.chk_Concise_XML.Size = new System.Drawing.Size(115, 17);
            this.chk_Concise_XML.TabIndex = 8;
            this.chk_Concise_XML.Text = "\"Concise\" XML file";
            this.chk_Concise_XML.UseVisualStyleBackColor = true;
            // 
            // chk_XML_Min_Data
            // 
            this.chk_XML_Min_Data.AutoSize = true;
            this.chk_XML_Min_Data.Location = new System.Drawing.Point(19, 42);
            this.chk_XML_Min_Data.Name = "chk_XML_Min_Data";
            this.chk_XML_Min_Data.Size = new System.Drawing.Size(118, 17);
            this.chk_XML_Min_Data.TabIndex = 9;
            this.chk_XML_Min_Data.Text = "Bare Minimum Data";
            this.chk_XML_Min_Data.UseVisualStyleBackColor = true;
            // 
            // group_XML
            // 
            this.group_XML.Controls.Add(this.chk_XML_Min_Data);
            this.group_XML.Controls.Add(this.chk_Concise_XML);
            this.group_XML.Controls.Add(this.btn_Run_XML);
            this.group_XML.Location = new System.Drawing.Point(46, 70);
            this.group_XML.Name = "group_XML";
            this.group_XML.Size = new System.Drawing.Size(146, 116);
            this.group_XML.TabIndex = 10;
            this.group_XML.TabStop = false;
            this.group_XML.Text = "Excel To XML";
            // 
            // group_Excel
            // 
            this.group_Excel.Controls.Add(this.btn_Create_Excel);
            this.group_Excel.Location = new System.Drawing.Point(367, 70);
            this.group_Excel.Name = "group_Excel";
            this.group_Excel.Size = new System.Drawing.Size(146, 116);
            this.group_Excel.TabIndex = 11;
            this.group_Excel.TabStop = false;
            this.group_Excel.Text = "XML To Excel";
            // 
            // btn_Create_Excel
            // 
            this.btn_Create_Excel.Location = new System.Drawing.Point(36, 69);
            this.btn_Create_Excel.Name = "btn_Create_Excel";
            this.btn_Create_Excel.Size = new System.Drawing.Size(75, 23);
            this.btn_Create_Excel.TabIndex = 0;
            this.btn_Create_Excel.Text = "Create Excel";
            this.btn_Create_Excel.UseVisualStyleBackColor = true;
            this.btn_Create_Excel.Click += new System.EventHandler(this.btn_Create_Excel_Click);
            // 
            // bw_XML
            // 
            this.bw_XML.WorkerSupportsCancellation = true;
            this.bw_XML.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bw_XML_DoWork);
            this.bw_XML.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bw_XML_RunComplete);
            // 
            // bw_Excel
            // 
            this.bw_Excel.WorkerSupportsCancellation = true;
            this.bw_Excel.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bw_Excel_DoWork);
            this.bw_Excel.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bw_Excel_RunComplete);
            // 
            // lbl_Input_Path
            // 
            this.lbl_Input_Path.AutoSize = true;
            this.lbl_Input_Path.Location = new System.Drawing.Point(12, 9);
            this.lbl_Input_Path.Name = "lbl_Input_Path";
            this.lbl_Input_Path.Size = new System.Drawing.Size(50, 13);
            this.lbl_Input_Path.TabIndex = 12;
            this.lbl_Input_Path.Text = "Input file:";
            // 
            // lbl_Output_Name
            // 
            this.lbl_Output_Name.AutoSize = true;
            this.lbl_Output_Name.Location = new System.Drawing.Point(12, 35);
            this.lbl_Output_Name.Name = "lbl_Output_Name";
            this.lbl_Output_Name.Size = new System.Drawing.Size(153, 13);
            this.lbl_Output_Name.TabIndex = 13;
            this.lbl_Output_Name.Text = "Output filename (no extension):";
            // 
            // tbOutputName
            // 
            this.tbOutputName.Location = new System.Drawing.Point(171, 32);
            this.tbOutputName.Name = "tbOutputName";
            this.tbOutputName.Size = new System.Drawing.Size(143, 20);
            this.tbOutputName.TabIndex = 14;
            this.tbOutputName.Text = "TestFile";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(573, 232);
            this.Controls.Add(this.tbOutputName);
            this.Controls.Add(this.lbl_Output_Name);
            this.Controls.Add(this.lbl_Input_Path);
            this.Controls.Add(this.group_Excel);
            this.Controls.Add(this.group_XML);
            this.Controls.Add(this.btn_Remind_Header);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_Browse);
            this.Controls.Add(this.tbPath);
            this.Name = "Form1";
            this.Text = "Form1";
            this.group_XML.ResumeLayout(false);
            this.group_XML.PerformLayout();
            this.group_Excel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbPath;
        private System.Windows.Forms.Button btn_Run_XML;
        private System.Windows.Forms.Button btn_Browse;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Button btn_Remind_Header;
        private System.Windows.Forms.CheckBox chk_Concise_XML;
        private System.Windows.Forms.CheckBox chk_XML_Min_Data;
        private System.Windows.Forms.GroupBox group_XML;
        private System.Windows.Forms.GroupBox group_Excel;
        private System.ComponentModel.BackgroundWorker bw_XML;
        private System.ComponentModel.BackgroundWorker bw_Excel;
        private System.Windows.Forms.Button btn_Create_Excel;
        private System.Windows.Forms.Label lbl_Input_Path;
        private System.Windows.Forms.Label lbl_Output_Name;
        private System.Windows.Forms.TextBox tbOutputName;
    }
}

