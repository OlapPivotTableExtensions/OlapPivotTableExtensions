namespace OlapPivotTableExtensions
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabCalcs = new System.Windows.Forms.TabPage();
            this.linkHelp = new System.Windows.Forms.LinkLabel();
            this.btnDeleteCalc = new System.Windows.Forms.Button();
            this.btnAddCalc = new System.Windows.Forms.Button();
            this.txtCalcFormula = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.comboCalcName = new System.Windows.Forms.ComboBox();
            this.tabLibrary = new System.Windows.Forms.TabPage();
            this.radDelete = new System.Windows.Forms.RadioButton();
            this.radioExport = new System.Windows.Forms.RadioButton();
            this.btnExportFilePath = new System.Windows.Forms.Button();
            this.txtExportFilePath = new System.Windows.Forms.TextBox();
            this.radImport = new System.Windows.Forms.RadioButton();
            this.listImportExportCalcs = new System.Windows.Forms.CheckedListBox();
            this.btnImportExportExecute = new System.Windows.Forms.Button();
            this.lblSelectCalcs = new System.Windows.Forms.Label();
            this.btnImportFilePath = new System.Windows.Forms.Button();
            this.txtImportFilePath = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tabMDX = new System.Windows.Forms.TabPage();
            this.txtMDX = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tabAbout = new System.Windows.Forms.TabPage();
            this.linkCodeplex = new System.Windows.Forms.LinkLabel();
            this.label5 = new System.Windows.Forms.Label();
            this.lblVersion = new System.Windows.Forms.Label();
            this.tabDefaults = new System.Windows.Forms.TabPage();
            this.chkShowCalcMembers = new System.Windows.Forms.CheckBox();
            this.btnSaveDefaults = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.tabControl.SuspendLayout();
            this.tabCalcs.SuspendLayout();
            this.tabLibrary.SuspendLayout();
            this.tabMDX.SuspendLayout();
            this.tabAbout.SuspendLayout();
            this.tabDefaults.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl
            // 
            this.tabControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl.Controls.Add(this.tabCalcs);
            this.tabControl.Controls.Add(this.tabLibrary);
            this.tabControl.Controls.Add(this.tabMDX);
            this.tabControl.Controls.Add(this.tabDefaults);
            this.tabControl.Controls.Add(this.tabAbout);
            this.tabControl.Location = new System.Drawing.Point(12, 12);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(447, 331);
            this.tabControl.TabIndex = 0;
            this.tabControl.SelectedIndexChanged += new System.EventHandler(this.tabControl_SelectedIndexChanged);
            // 
            // tabCalcs
            // 
            this.tabCalcs.Controls.Add(this.linkHelp);
            this.tabCalcs.Controls.Add(this.btnDeleteCalc);
            this.tabCalcs.Controls.Add(this.btnAddCalc);
            this.tabCalcs.Controls.Add(this.txtCalcFormula);
            this.tabCalcs.Controls.Add(this.label2);
            this.tabCalcs.Controls.Add(this.label1);
            this.tabCalcs.Controls.Add(this.comboCalcName);
            this.tabCalcs.Location = new System.Drawing.Point(4, 22);
            this.tabCalcs.Name = "tabCalcs";
            this.tabCalcs.Padding = new System.Windows.Forms.Padding(3);
            this.tabCalcs.Size = new System.Drawing.Size(439, 305);
            this.tabCalcs.TabIndex = 0;
            this.tabCalcs.Text = "Calculations";
            this.tabCalcs.UseVisualStyleBackColor = true;
            // 
            // linkHelp
            // 
            this.linkHelp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.linkHelp.AutoSize = true;
            this.linkHelp.Location = new System.Drawing.Point(404, 65);
            this.linkHelp.Name = "linkHelp";
            this.linkHelp.Size = new System.Drawing.Size(29, 13);
            this.linkHelp.TabIndex = 6;
            this.linkHelp.TabStop = true;
            this.linkHelp.Text = "Help";
            this.linkHelp.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.linkHelp.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkHelp_LinkClicked);
            // 
            // btnDeleteCalc
            // 
            this.btnDeleteCalc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDeleteCalc.Enabled = false;
            this.btnDeleteCalc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDeleteCalc.Location = new System.Drawing.Point(166, 276);
            this.btnDeleteCalc.Name = "btnDeleteCalc";
            this.btnDeleteCalc.Size = new System.Drawing.Size(139, 23);
            this.btnDeleteCalc.TabIndex = 5;
            this.btnDeleteCalc.Text = "Delete from PivotTable";
            this.btnDeleteCalc.UseVisualStyleBackColor = true;
            this.btnDeleteCalc.Click += new System.EventHandler(this.btnDeleteCalc_Click);
            // 
            // btnAddCalc
            // 
            this.btnAddCalc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAddCalc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAddCalc.Location = new System.Drawing.Point(311, 276);
            this.btnAddCalc.Name = "btnAddCalc";
            this.btnAddCalc.Size = new System.Drawing.Size(122, 23);
            this.btnAddCalc.TabIndex = 4;
            this.btnAddCalc.Text = "Add to PivotTable";
            this.btnAddCalc.UseVisualStyleBackColor = true;
            this.btnAddCalc.Click += new System.EventHandler(this.btnAddCalc_Click);
            // 
            // txtCalcFormula
            // 
            this.txtCalcFormula.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCalcFormula.Location = new System.Drawing.Point(7, 82);
            this.txtCalcFormula.Multiline = true;
            this.txtCalcFormula.Name = "txtCalcFormula";
            this.txtCalcFormula.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtCalcFormula.Size = new System.Drawing.Size(426, 187);
            this.txtCalcFormula.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(294, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Formula: (e.g. [Measures].[Sales Amount] - [Measures].[Cost])";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(192, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Calculation Name: (e.g. My Calculation)";
            // 
            // comboCalcName
            // 
            this.comboCalcName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.comboCalcName.FormattingEnabled = true;
            this.comboCalcName.Location = new System.Drawing.Point(7, 28);
            this.comboCalcName.Name = "comboCalcName";
            this.comboCalcName.Size = new System.Drawing.Size(426, 21);
            this.comboCalcName.TabIndex = 1;
            this.comboCalcName.TextChanged += new System.EventHandler(this.comboCalcName_TextChanged);
            // 
            // tabLibrary
            // 
            this.tabLibrary.Controls.Add(this.radDelete);
            this.tabLibrary.Controls.Add(this.radioExport);
            this.tabLibrary.Controls.Add(this.btnExportFilePath);
            this.tabLibrary.Controls.Add(this.txtExportFilePath);
            this.tabLibrary.Controls.Add(this.radImport);
            this.tabLibrary.Controls.Add(this.listImportExportCalcs);
            this.tabLibrary.Controls.Add(this.btnImportExportExecute);
            this.tabLibrary.Controls.Add(this.lblSelectCalcs);
            this.tabLibrary.Controls.Add(this.btnImportFilePath);
            this.tabLibrary.Controls.Add(this.txtImportFilePath);
            this.tabLibrary.Controls.Add(this.label4);
            this.tabLibrary.Location = new System.Drawing.Point(4, 22);
            this.tabLibrary.Name = "tabLibrary";
            this.tabLibrary.Padding = new System.Windows.Forms.Padding(3);
            this.tabLibrary.Size = new System.Drawing.Size(439, 305);
            this.tabLibrary.TabIndex = 3;
            this.tabLibrary.Text = "Library";
            this.tabLibrary.UseVisualStyleBackColor = true;
            // 
            // radDelete
            // 
            this.radDelete.AutoSize = true;
            this.radDelete.Location = new System.Drawing.Point(10, 70);
            this.radDelete.Name = "radDelete";
            this.radDelete.Size = new System.Drawing.Size(173, 17);
            this.radDelete.TabIndex = 12;
            this.radDelete.Text = "Delete Calculations from Library";
            this.radDelete.UseVisualStyleBackColor = true;
            this.radDelete.CheckedChanged += new System.EventHandler(this.radDelete_CheckedChanged);
            // 
            // radioExport
            // 
            this.radioExport.AutoSize = true;
            this.radioExport.Location = new System.Drawing.Point(10, 47);
            this.radioExport.Name = "radioExport";
            this.radioExport.Size = new System.Drawing.Size(67, 17);
            this.radioExport.TabIndex = 11;
            this.radioExport.Text = "Export to";
            this.radioExport.UseVisualStyleBackColor = true;
            this.radioExport.CheckedChanged += new System.EventHandler(this.radioExport_CheckedChanged);
            // 
            // btnExportFilePath
            // 
            this.btnExportFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportFilePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportFilePath.Location = new System.Drawing.Point(397, 47);
            this.btnExportFilePath.Name = "btnExportFilePath";
            this.btnExportFilePath.Size = new System.Drawing.Size(28, 20);
            this.btnExportFilePath.TabIndex = 10;
            this.btnExportFilePath.Text = "...";
            this.btnExportFilePath.UseVisualStyleBackColor = true;
            this.btnExportFilePath.Click += new System.EventHandler(this.btnExportFilePath_Click);
            // 
            // txtExportFilePath
            // 
            this.txtExportFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtExportFilePath.Location = new System.Drawing.Point(93, 47);
            this.txtExportFilePath.Name = "txtExportFilePath";
            this.txtExportFilePath.ReadOnly = true;
            this.txtExportFilePath.Size = new System.Drawing.Size(298, 20);
            this.txtExportFilePath.TabIndex = 9;
            // 
            // radImport
            // 
            this.radImport.AutoSize = true;
            this.radImport.Checked = true;
            this.radImport.Location = new System.Drawing.Point(10, 24);
            this.radImport.Name = "radImport";
            this.radImport.Size = new System.Drawing.Size(77, 17);
            this.radImport.TabIndex = 8;
            this.radImport.TabStop = true;
            this.radImport.Text = "Import from";
            this.radImport.UseVisualStyleBackColor = true;
            this.radImport.CheckedChanged += new System.EventHandler(this.radImport_CheckedChanged);
            // 
            // listImportExportCalcs
            // 
            this.listImportExportCalcs.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listImportExportCalcs.CheckOnClick = true;
            this.listImportExportCalcs.Enabled = false;
            this.listImportExportCalcs.FormattingEnabled = true;
            this.listImportExportCalcs.Location = new System.Drawing.Point(11, 111);
            this.listImportExportCalcs.Name = "listImportExportCalcs";
            this.listImportExportCalcs.ScrollAlwaysVisible = true;
            this.listImportExportCalcs.Size = new System.Drawing.Size(414, 154);
            this.listImportExportCalcs.TabIndex = 7;
            this.listImportExportCalcs.ThreeDCheckBoxes = true;
            // 
            // btnImportExportExecute
            // 
            this.btnImportExportExecute.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnImportExportExecute.Enabled = false;
            this.btnImportExportExecute.Location = new System.Drawing.Point(349, 272);
            this.btnImportExportExecute.Name = "btnImportExportExecute";
            this.btnImportExportExecute.Size = new System.Drawing.Size(75, 23);
            this.btnImportExportExecute.TabIndex = 5;
            this.btnImportExportExecute.Text = "Execute";
            this.btnImportExportExecute.UseVisualStyleBackColor = true;
            this.btnImportExportExecute.Click += new System.EventHandler(this.btnImportExportExecute_Click);
            // 
            // lblSelectCalcs
            // 
            this.lblSelectCalcs.AutoSize = true;
            this.lblSelectCalcs.Location = new System.Drawing.Point(8, 94);
            this.lblSelectCalcs.Name = "lblSelectCalcs";
            this.lblSelectCalcs.Size = new System.Drawing.Size(215, 13);
            this.lblSelectCalcs.TabIndex = 4;
            this.lblSelectCalcs.Text = "Select Calculations to Import/Export/Delete:";
            // 
            // btnImportFilePath
            // 
            this.btnImportFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnImportFilePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImportFilePath.Location = new System.Drawing.Point(397, 24);
            this.btnImportFilePath.Name = "btnImportFilePath";
            this.btnImportFilePath.Size = new System.Drawing.Size(28, 20);
            this.btnImportFilePath.TabIndex = 2;
            this.btnImportFilePath.Text = "...";
            this.btnImportFilePath.UseVisualStyleBackColor = true;
            this.btnImportFilePath.Click += new System.EventHandler(this.btnImportFilePath_Click);
            // 
            // txtImportFilePath
            // 
            this.txtImportFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtImportFilePath.Location = new System.Drawing.Point(93, 24);
            this.txtImportFilePath.Name = "txtImportFilePath";
            this.txtImportFilePath.ReadOnly = true;
            this.txtImportFilePath.Size = new System.Drawing.Size(298, 20);
            this.txtImportFilePath.TabIndex = 1;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 7);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(158, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Calculation Library Maintenance";
            // 
            // tabMDX
            // 
            this.tabMDX.Controls.Add(this.txtMDX);
            this.tabMDX.Controls.Add(this.label3);
            this.tabMDX.Location = new System.Drawing.Point(4, 22);
            this.tabMDX.Name = "tabMDX";
            this.tabMDX.Padding = new System.Windows.Forms.Padding(3);
            this.tabMDX.Size = new System.Drawing.Size(439, 305);
            this.tabMDX.TabIndex = 1;
            this.tabMDX.Text = "MDX";
            this.tabMDX.UseVisualStyleBackColor = true;
            // 
            // txtMDX
            // 
            this.txtMDX.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtMDX.Location = new System.Drawing.Point(10, 24);
            this.txtMDX.MaxLength = 1000000;
            this.txtMDX.Multiline = true;
            this.txtMDX.Name = "txtMDX";
            this.txtMDX.ReadOnly = true;
            this.txtMDX.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtMDX.Size = new System.Drawing.Size(423, 272);
            this.txtMDX.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 7);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "MDX Query:";
            // 
            // tabAbout
            // 
            this.tabAbout.Controls.Add(this.linkCodeplex);
            this.tabAbout.Controls.Add(this.label5);
            this.tabAbout.Controls.Add(this.lblVersion);
            this.tabAbout.Location = new System.Drawing.Point(4, 22);
            this.tabAbout.Name = "tabAbout";
            this.tabAbout.Padding = new System.Windows.Forms.Padding(3);
            this.tabAbout.Size = new System.Drawing.Size(439, 305);
            this.tabAbout.TabIndex = 2;
            this.tabAbout.Text = "About";
            this.tabAbout.UseVisualStyleBackColor = true;
            // 
            // linkCodeplex
            // 
            this.linkCodeplex.AutoSize = true;
            this.linkCodeplex.Cursor = System.Windows.Forms.Cursors.Hand;
            this.linkCodeplex.Location = new System.Drawing.Point(12, 64);
            this.linkCodeplex.Name = "linkCodeplex";
            this.linkCodeplex.Size = new System.Drawing.Size(242, 13);
            this.linkCodeplex.TabIndex = 2;
            this.linkCodeplex.TabStop = true;
            this.linkCodeplex.Text = "http://www.codeplex.com/OlapPivotTableExtend";
            this.linkCodeplex.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkCodeplex_LinkClicked);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(11, 47);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(275, 13);
            this.label5.TabIndex = 1;
            this.label5.Text = "View documentation and report bugs and suggestions at:";
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVersion.Location = new System.Drawing.Point(10, 16);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(220, 13);
            this.lblVersion.TabIndex = 0;
            this.lblVersion.Text = "OLAP PivotTable Extensions v0.0.0.0";
            // 
            // tabDefaults
            // 
            this.tabDefaults.Controls.Add(this.label6);
            this.tabDefaults.Controls.Add(this.btnSaveDefaults);
            this.tabDefaults.Controls.Add(this.chkShowCalcMembers);
            this.tabDefaults.Location = new System.Drawing.Point(4, 22);
            this.tabDefaults.Name = "tabDefaults";
            this.tabDefaults.Size = new System.Drawing.Size(439, 305);
            this.tabDefaults.TabIndex = 4;
            this.tabDefaults.Text = "Defaults";
            this.tabDefaults.UseVisualStyleBackColor = true;
            // 
            // chkShowCalcMembers
            // 
            this.chkShowCalcMembers.AutoSize = true;
            this.chkShowCalcMembers.CheckAlign = System.Drawing.ContentAlignment.TopLeft;
            this.chkShowCalcMembers.Location = new System.Drawing.Point(13, 34);
            this.chkShowCalcMembers.Name = "chkShowCalcMembers";
            this.chkShowCalcMembers.Size = new System.Drawing.Size(335, 17);
            this.chkShowCalcMembers.TabIndex = 0;
            this.chkShowCalcMembers.Text = "Turn on \"Show calculated members from OLAP server\" by default";
            this.chkShowCalcMembers.UseVisualStyleBackColor = true;
            // 
            // btnSaveDefaults
            // 
            this.btnSaveDefaults.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSaveDefaults.Location = new System.Drawing.Point(351, 270);
            this.btnSaveDefaults.Name = "btnSaveDefaults";
            this.btnSaveDefaults.Size = new System.Drawing.Size(75, 23);
            this.btnSaveDefaults.TabIndex = 1;
            this.btnSaveDefaults.Text = "Save";
            this.btnSaveDefaults.UseVisualStyleBackColor = true;
            this.btnSaveDefaults.Click += new System.EventHandler(this.btnSaveDefaults_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(10, 15);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(171, 13);
            this.label6.TabIndex = 2;
            this.label6.Text = "For new OLAP PivotTables...";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 355);
            this.Controls.Add(this.tabControl);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(400, 300);
            this.Name = "MainForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "OLAP PivotTable Extensions";
            this.tabControl.ResumeLayout(false);
            this.tabCalcs.ResumeLayout(false);
            this.tabCalcs.PerformLayout();
            this.tabLibrary.ResumeLayout(false);
            this.tabLibrary.PerformLayout();
            this.tabMDX.ResumeLayout(false);
            this.tabMDX.PerformLayout();
            this.tabAbout.ResumeLayout(false);
            this.tabAbout.PerformLayout();
            this.tabDefaults.ResumeLayout(false);
            this.tabDefaults.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabCalcs;
        private System.Windows.Forms.TabPage tabMDX;
        private System.Windows.Forms.ComboBox comboCalcName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtCalcFormula;
        private System.Windows.Forms.Button btnAddCalc;
        private System.Windows.Forms.Button btnDeleteCalc;
        private System.Windows.Forms.TextBox txtMDX;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.LinkLabel linkHelp;
        private System.Windows.Forms.TabPage tabAbout;
        private System.Windows.Forms.Label lblVersion;
        private System.Windows.Forms.LinkLabel linkCodeplex;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TabPage tabLibrary;
        private System.Windows.Forms.Button btnImportFilePath;
        private System.Windows.Forms.TextBox txtImportFilePath;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblSelectCalcs;
        private System.Windows.Forms.Button btnImportExportExecute;
        private System.Windows.Forms.CheckedListBox listImportExportCalcs;
        private System.Windows.Forms.RadioButton radImport;
        private System.Windows.Forms.RadioButton radioExport;
        private System.Windows.Forms.Button btnExportFilePath;
        private System.Windows.Forms.TextBox txtExportFilePath;
        private System.Windows.Forms.RadioButton radDelete;
        private System.Windows.Forms.TabPage tabDefaults;
        private System.Windows.Forms.CheckBox chkShowCalcMembers;
        private System.Windows.Forms.Button btnSaveDefaults;
        private System.Windows.Forms.Label label6;
    }
}