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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
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
            this.chkFormatMDX = new System.Windows.Forms.CheckBox();
            this.richTextBoxMDX = new System.Windows.Forms.RichTextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tabSearch = new System.Windows.Forms.TabPage();
            this.chkAddToCurrentFilters = new System.Windows.Forms.CheckBox();
            this.btnCancelSearch = new System.Windows.Forms.Button();
            this.prgSearch = new System.Windows.Forms.ProgressBar();
            this.lblNoSearchMatches = new System.Windows.Forms.Label();
            this.btnSearchAdd = new System.Windows.Forms.Button();
            this.chkMemberProperties = new System.Windows.Forms.CheckBox();
            this.chkExactMatch = new System.Windows.Forms.CheckBox();
            this.cmbLookIn = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.btnFind = new System.Windows.Forms.Button();
            this.lblSearchFor = new System.Windows.Forms.Label();
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.lblSearchError = new System.Windows.Forms.Label();
            this.dataGridSearchResults = new System.Windows.Forms.DataGridView();
            this.colCheck = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFolder = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDesc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cubeSearchMatchBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tabFilterList = new System.Windows.Forms.TabPage();
            this.btnFilterListShowCurrentFilters = new System.Windows.Forms.Button();
            this.btnCancelFilterList = new System.Windows.Forms.Button();
            this.progressFilterList = new System.Windows.Forms.ProgressBar();
            this.lblFilterListError = new System.Windows.Forms.Label();
            this.btnFilterList = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.cmbFilterListLookIn = new System.Windows.Forms.ComboBox();
            this.txtFilterList = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.tabDefaults = new System.Windows.Forms.TabPage();
            this.chkRefreshDataWhenOpeningTheFile = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.btnSaveDefaults = new System.Windows.Forms.Button();
            this.chkShowCalcMembers = new System.Windows.Forms.CheckBox();
            this.tabAbout = new System.Windows.Forms.TabPage();
            this.btnUpgradeOnRefresh = new System.Windows.Forms.Button();
            this.lblExcelUILanguage = new System.Windows.Forms.Label();
            this.lblUpgradePivotTableInstructions = new System.Windows.Forms.Label();
            this.linkUpgradePivotTable = new System.Windows.Forms.LinkLabel();
            this.lblPivotTableVersion = new System.Windows.Forms.Label();
            this.linkCodeplex = new System.Windows.Forms.LinkLabel();
            this.label5 = new System.Windows.Forms.Label();
            this.lblVersion = new System.Windows.Forms.Label();
            this.tooltip = new System.Windows.Forms.ToolTip(this.components);
            this.lblFormattingMdxQuery = new System.Windows.Forms.Label();
            this.tabControl.SuspendLayout();
            this.tabCalcs.SuspendLayout();
            this.tabLibrary.SuspendLayout();
            this.tabMDX.SuspendLayout();
            this.tabSearch.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridSearchResults)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cubeSearchMatchBindingSource)).BeginInit();
            this.tabFilterList.SuspendLayout();
            this.tabDefaults.SuspendLayout();
            this.tabAbout.SuspendLayout();
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
            this.tabControl.Controls.Add(this.tabSearch);
            this.tabControl.Controls.Add(this.tabFilterList);
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
            this.txtCalcFormula.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
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
            this.tabMDX.Controls.Add(this.lblFormattingMdxQuery);
            this.tabMDX.Controls.Add(this.chkFormatMDX);
            this.tabMDX.Controls.Add(this.richTextBoxMDX);
            this.tabMDX.Controls.Add(this.label3);
            this.tabMDX.Location = new System.Drawing.Point(4, 22);
            this.tabMDX.Name = "tabMDX";
            this.tabMDX.Padding = new System.Windows.Forms.Padding(3);
            this.tabMDX.Size = new System.Drawing.Size(439, 305);
            this.tabMDX.TabIndex = 1;
            this.tabMDX.Text = "MDX";
            this.tabMDX.UseVisualStyleBackColor = true;
            // 
            // chkFormatMDX
            // 
            this.chkFormatMDX.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.chkFormatMDX.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkFormatMDX.Location = new System.Drawing.Point(209, 9);
            this.chkFormatMDX.Name = "chkFormatMDX";
            this.chkFormatMDX.Size = new System.Drawing.Size(221, 17);
            this.chkFormatMDX.TabIndex = 5;
            this.chkFormatMDX.Text = "Format MDX query using web service?";
            this.chkFormatMDX.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkFormatMDX.UseVisualStyleBackColor = true;
            this.chkFormatMDX.CheckedChanged += new System.EventHandler(this.chkFormatMDX_CheckedChanged);
            // 
            // richTextBoxMDX
            // 
            this.richTextBoxMDX.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.richTextBoxMDX.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBoxMDX.Location = new System.Drawing.Point(10, 32);
            this.richTextBoxMDX.Name = "richTextBoxMDX";
            this.richTextBoxMDX.ReadOnly = true;
            this.richTextBoxMDX.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.richTextBoxMDX.Size = new System.Drawing.Size(420, 263);
            this.richTextBoxMDX.TabIndex = 4;
            this.richTextBoxMDX.Text = "";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(7, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "MDX Query:";
            // 
            // tabSearch
            // 
            this.tabSearch.Controls.Add(this.chkAddToCurrentFilters);
            this.tabSearch.Controls.Add(this.btnCancelSearch);
            this.tabSearch.Controls.Add(this.prgSearch);
            this.tabSearch.Controls.Add(this.lblNoSearchMatches);
            this.tabSearch.Controls.Add(this.btnSearchAdd);
            this.tabSearch.Controls.Add(this.chkMemberProperties);
            this.tabSearch.Controls.Add(this.chkExactMatch);
            this.tabSearch.Controls.Add(this.cmbLookIn);
            this.tabSearch.Controls.Add(this.label7);
            this.tabSearch.Controls.Add(this.btnFind);
            this.tabSearch.Controls.Add(this.lblSearchFor);
            this.tabSearch.Controls.Add(this.txtSearch);
            this.tabSearch.Controls.Add(this.lblSearchError);
            this.tabSearch.Controls.Add(this.dataGridSearchResults);
            this.tabSearch.Location = new System.Drawing.Point(4, 22);
            this.tabSearch.Name = "tabSearch";
            this.tabSearch.Padding = new System.Windows.Forms.Padding(3);
            this.tabSearch.Size = new System.Drawing.Size(439, 305);
            this.tabSearch.TabIndex = 5;
            this.tabSearch.Text = "Search";
            this.tabSearch.UseVisualStyleBackColor = true;
            // 
            // chkAddToCurrentFilters
            // 
            this.chkAddToCurrentFilters.AutoSize = true;
            this.chkAddToCurrentFilters.Location = new System.Drawing.Point(304, 95);
            this.chkAddToCurrentFilters.Name = "chkAddToCurrentFilters";
            this.chkAddToCurrentFilters.Size = new System.Drawing.Size(120, 17);
            this.chkAddToCurrentFilters.TabIndex = 13;
            this.chkAddToCurrentFilters.Text = "Add to current filters";
            this.chkAddToCurrentFilters.UseVisualStyleBackColor = true;
            // 
            // btnCancelSearch
            // 
            this.btnCancelSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnCancelSearch.BackColor = System.Drawing.SystemColors.Control;
            this.btnCancelSearch.Location = new System.Drawing.Point(10, 276);
            this.btnCancelSearch.Name = "btnCancelSearch";
            this.btnCancelSearch.Size = new System.Drawing.Size(60, 23);
            this.btnCancelSearch.TabIndex = 11;
            this.btnCancelSearch.Text = "Cancel";
            this.btnCancelSearch.UseVisualStyleBackColor = false;
            this.btnCancelSearch.Visible = false;
            this.btnCancelSearch.Click += new System.EventHandler(this.btnCancelSearch_Click);
            // 
            // prgSearch
            // 
            this.prgSearch.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.prgSearch.Location = new System.Drawing.Point(79, 279);
            this.prgSearch.Name = "prgSearch";
            this.prgSearch.Size = new System.Drawing.Size(206, 18);
            this.prgSearch.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.prgSearch.TabIndex = 10;
            this.prgSearch.Visible = false;
            // 
            // lblNoSearchMatches
            // 
            this.lblNoSearchMatches.AutoSize = true;
            this.lblNoSearchMatches.BackColor = System.Drawing.SystemColors.Window;
            this.lblNoSearchMatches.ForeColor = System.Drawing.Color.Red;
            this.lblNoSearchMatches.Location = new System.Drawing.Point(16, 143);
            this.lblNoSearchMatches.Name = "lblNoSearchMatches";
            this.lblNoSearchMatches.Size = new System.Drawing.Size(94, 13);
            this.lblNoSearchMatches.TabIndex = 9;
            this.lblNoSearchMatches.Text = "No matches found";
            this.lblNoSearchMatches.Visible = false;
            // 
            // btnSearchAdd
            // 
            this.btnSearchAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSearchAdd.Enabled = false;
            this.btnSearchAdd.Location = new System.Drawing.Point(309, 276);
            this.btnSearchAdd.Name = "btnSearchAdd";
            this.btnSearchAdd.Size = new System.Drawing.Size(119, 23);
            this.btnSearchAdd.TabIndex = 8;
            this.btnSearchAdd.Text = "Add to PivotTable";
            this.btnSearchAdd.UseVisualStyleBackColor = true;
            this.btnSearchAdd.Click += new System.EventHandler(this.btnSearchAdd_Click);
            // 
            // chkMemberProperties
            // 
            this.chkMemberProperties.AutoSize = true;
            this.chkMemberProperties.Checked = true;
            this.chkMemberProperties.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkMemberProperties.Location = new System.Drawing.Point(126, 95);
            this.chkMemberProperties.Name = "chkMemberProperties";
            this.chkMemberProperties.Size = new System.Drawing.Size(149, 17);
            this.chkMemberProperties.TabIndex = 6;
            this.chkMemberProperties.Text = "Search member properties";
            this.chkMemberProperties.UseVisualStyleBackColor = true;
            // 
            // chkExactMatch
            // 
            this.chkExactMatch.AutoSize = true;
            this.chkExactMatch.Location = new System.Drawing.Point(10, 95);
            this.chkExactMatch.Name = "chkExactMatch";
            this.chkExactMatch.Size = new System.Drawing.Size(85, 17);
            this.chkExactMatch.TabIndex = 5;
            this.chkExactMatch.Text = "Exact match";
            this.chkExactMatch.UseVisualStyleBackColor = true;
            // 
            // cmbLookIn
            // 
            this.cmbLookIn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbLookIn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLookIn.FormattingEnabled = true;
            this.cmbLookIn.Items.AddRange(new object[] {
            "Field list",
            "Dimension data"});
            this.cmbLookIn.Location = new System.Drawing.Point(10, 68);
            this.cmbLookIn.MaxDropDownItems = 10;
            this.cmbLookIn.Name = "cmbLookIn";
            this.cmbLookIn.Size = new System.Drawing.Size(418, 21);
            this.cmbLookIn.TabIndex = 3;
            this.cmbLookIn.SelectedIndexChanged += new System.EventHandler(this.cmbLookIn_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(7, 52);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(45, 13);
            this.label7.TabIndex = 4;
            this.label7.Text = "Look in:";
            // 
            // btnFind
            // 
            this.btnFind.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFind.Location = new System.Drawing.Point(355, 22);
            this.btnFind.Name = "btnFind";
            this.btnFind.Size = new System.Drawing.Size(73, 23);
            this.btnFind.TabIndex = 2;
            this.btnFind.Text = "Find Next";
            this.btnFind.UseVisualStyleBackColor = true;
            this.btnFind.Click += new System.EventHandler(this.btnFind_Click);
            // 
            // lblSearchFor
            // 
            this.lblSearchFor.AutoSize = true;
            this.lblSearchFor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSearchFor.Location = new System.Drawing.Point(7, 7);
            this.lblSearchFor.Name = "lblSearchFor";
            this.lblSearchFor.Size = new System.Drawing.Size(56, 13);
            this.lblSearchFor.TabIndex = 1;
            this.lblSearchFor.Text = "Find what:";
            // 
            // txtSearch
            // 
            this.txtSearch.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSearch.Location = new System.Drawing.Point(10, 24);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(339, 20);
            this.txtSearch.TabIndex = 0;
            this.txtSearch.Leave += new System.EventHandler(this.txtSearch_Leave);
            this.txtSearch.Enter += new System.EventHandler(this.txtSearch_Enter);
            // 
            // lblSearchError
            // 
            this.lblSearchError.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSearchError.AutoEllipsis = true;
            this.lblSearchError.BackColor = System.Drawing.Color.Transparent;
            this.lblSearchError.ForeColor = System.Drawing.Color.Red;
            this.lblSearchError.Location = new System.Drawing.Point(7, 279);
            this.lblSearchError.Name = "lblSearchError";
            this.lblSearchError.Size = new System.Drawing.Size(296, 18);
            this.lblSearchError.TabIndex = 12;
            this.lblSearchError.Text = "Error: Text here";
            this.lblSearchError.Visible = false;
            // 
            // dataGridSearchResults
            // 
            this.dataGridSearchResults.AllowUserToAddRows = false;
            this.dataGridSearchResults.AllowUserToDeleteRows = false;
            this.dataGridSearchResults.AllowUserToResizeRows = false;
            this.dataGridSearchResults.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridSearchResults.AutoGenerateColumns = false;
            this.dataGridSearchResults.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dataGridSearchResults.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridSearchResults.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridSearchResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridSearchResults.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colCheck,
            this.colName,
            this.colType,
            this.colFolder,
            this.colDesc});
            this.dataGridSearchResults.DataSource = this.cubeSearchMatchBindingSource;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridSearchResults.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridSearchResults.Enabled = false;
            this.dataGridSearchResults.Location = new System.Drawing.Point(10, 118);
            this.dataGridSearchResults.MultiSelect = false;
            this.dataGridSearchResults.Name = "dataGridSearchResults";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridSearchResults.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridSearchResults.RowHeadersVisible = false;
            this.dataGridSearchResults.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridSearchResults.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridSearchResults.Size = new System.Drawing.Size(418, 150);
            this.dataGridSearchResults.TabIndex = 7;
            this.dataGridSearchResults.TabStop = false;
            this.dataGridSearchResults.CellMouseLeave += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridSearchResults_CellMouseLeave);
            this.dataGridSearchResults.CellMouseEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridSearchResults_CellMouseEnter);
            this.dataGridSearchResults.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridSearchResults_CellClick);
            // 
            // colCheck
            // 
            this.colCheck.DataPropertyName = "Checked";
            this.colCheck.Frozen = true;
            this.colCheck.HeaderText = "Add";
            this.colCheck.MinimumWidth = 30;
            this.colCheck.Name = "colCheck";
            this.colCheck.Width = 30;
            // 
            // colName
            // 
            this.colName.DataPropertyName = "Name";
            this.colName.HeaderText = "Name";
            this.colName.Name = "colName";
            this.colName.ReadOnly = true;
            this.colName.Width = 140;
            // 
            // colType
            // 
            this.colType.DataPropertyName = "Type";
            this.colType.HeaderText = "Type";
            this.colType.MinimumWidth = 30;
            this.colType.Name = "colType";
            this.colType.ReadOnly = true;
            this.colType.Width = 60;
            // 
            // colFolder
            // 
            this.colFolder.DataPropertyName = "Folder";
            this.colFolder.HeaderText = "Folder";
            this.colFolder.Name = "colFolder";
            this.colFolder.ReadOnly = true;
            this.colFolder.Width = 120;
            // 
            // colDesc
            // 
            this.colDesc.DataPropertyName = "Description";
            this.colDesc.HeaderText = "Description";
            this.colDesc.Name = "colDesc";
            this.colDesc.ReadOnly = true;
            this.colDesc.Width = 300;
            // 
            // cubeSearchMatchBindingSource
            // 
            this.cubeSearchMatchBindingSource.DataSource = typeof(OlapPivotTableExtensions.CubeSearcher.CubeSearchMatch);
            // 
            // tabFilterList
            // 
            this.tabFilterList.Controls.Add(this.btnFilterListShowCurrentFilters);
            this.tabFilterList.Controls.Add(this.btnCancelFilterList);
            this.tabFilterList.Controls.Add(this.progressFilterList);
            this.tabFilterList.Controls.Add(this.lblFilterListError);
            this.tabFilterList.Controls.Add(this.btnFilterList);
            this.tabFilterList.Controls.Add(this.label9);
            this.tabFilterList.Controls.Add(this.cmbFilterListLookIn);
            this.tabFilterList.Controls.Add(this.txtFilterList);
            this.tabFilterList.Controls.Add(this.label8);
            this.tabFilterList.Location = new System.Drawing.Point(4, 22);
            this.tabFilterList.Name = "tabFilterList";
            this.tabFilterList.Padding = new System.Windows.Forms.Padding(3);
            this.tabFilterList.Size = new System.Drawing.Size(439, 305);
            this.tabFilterList.TabIndex = 6;
            this.tabFilterList.Text = "Filter List";
            this.tabFilterList.ToolTipText = "Paste in a list to filter your PivotTable";
            this.tabFilterList.UseVisualStyleBackColor = true;
            // 
            // btnFilterListShowCurrentFilters
            // 
            this.btnFilterListShowCurrentFilters.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFilterListShowCurrentFilters.Location = new System.Drawing.Point(313, 24);
            this.btnFilterListShowCurrentFilters.Name = "btnFilterListShowCurrentFilters";
            this.btnFilterListShowCurrentFilters.Size = new System.Drawing.Size(119, 23);
            this.btnFilterListShowCurrentFilters.TabIndex = 16;
            this.btnFilterListShowCurrentFilters.Text = "Show Current Filters";
            this.btnFilterListShowCurrentFilters.UseVisualStyleBackColor = true;
            this.btnFilterListShowCurrentFilters.Click += new System.EventHandler(this.btnFilterListShowCurrentFilters_Click);
            // 
            // btnCancelFilterList
            // 
            this.btnCancelFilterList.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnCancelFilterList.BackColor = System.Drawing.SystemColors.Control;
            this.btnCancelFilterList.Location = new System.Drawing.Point(9, 276);
            this.btnCancelFilterList.Name = "btnCancelFilterList";
            this.btnCancelFilterList.Size = new System.Drawing.Size(60, 23);
            this.btnCancelFilterList.TabIndex = 14;
            this.btnCancelFilterList.Text = "Cancel";
            this.btnCancelFilterList.UseVisualStyleBackColor = false;
            this.btnCancelFilterList.Visible = false;
            this.btnCancelFilterList.Click += new System.EventHandler(this.btnCancelFilterList_Click);
            // 
            // progressFilterList
            // 
            this.progressFilterList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.progressFilterList.Location = new System.Drawing.Point(78, 279);
            this.progressFilterList.Name = "progressFilterList";
            this.progressFilterList.Size = new System.Drawing.Size(206, 18);
            this.progressFilterList.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressFilterList.TabIndex = 13;
            this.progressFilterList.Visible = false;
            // 
            // lblFilterListError
            // 
            this.lblFilterListError.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lblFilterListError.AutoEllipsis = true;
            this.lblFilterListError.BackColor = System.Drawing.Color.Transparent;
            this.lblFilterListError.ForeColor = System.Drawing.Color.Red;
            this.lblFilterListError.Location = new System.Drawing.Point(6, 279);
            this.lblFilterListError.Name = "lblFilterListError";
            this.lblFilterListError.Size = new System.Drawing.Size(296, 18);
            this.lblFilterListError.TabIndex = 15;
            this.lblFilterListError.Text = "Error: Text here";
            this.lblFilterListError.Visible = false;
            // 
            // btnFilterList
            // 
            this.btnFilterList.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFilterList.Location = new System.Drawing.Point(309, 276);
            this.btnFilterList.Name = "btnFilterList";
            this.btnFilterList.Size = new System.Drawing.Size(119, 23);
            this.btnFilterList.TabIndex = 9;
            this.btnFilterList.Text = "Filter PivotTable";
            this.btnFilterList.UseVisualStyleBackColor = true;
            this.btnFilterList.Click += new System.EventHandler(this.btnFilterList_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(6, 59);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(204, 13);
            this.label9.TabIndex = 5;
            this.label9.Text = "To the members with the following names:";
            // 
            // cmbFilterListLookIn
            // 
            this.cmbFilterListLookIn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFilterListLookIn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFilterListLookIn.FormattingEnabled = true;
            this.cmbFilterListLookIn.Location = new System.Drawing.Point(9, 24);
            this.cmbFilterListLookIn.MaxDropDownItems = 10;
            this.cmbFilterListLookIn.Name = "cmbFilterListLookIn";
            this.cmbFilterListLookIn.Size = new System.Drawing.Size(293, 21);
            this.cmbFilterListLookIn.TabIndex = 4;
            // 
            // txtFilterList
            // 
            this.txtFilterList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtFilterList.BackColor = System.Drawing.SystemColors.Window;
            this.txtFilterList.Location = new System.Drawing.Point(9, 76);
            this.txtFilterList.MaxLength = 1000000;
            this.txtFilterList.Multiline = true;
            this.txtFilterList.Name = "txtFilterList";
            this.txtFilterList.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtFilterList.Size = new System.Drawing.Size(423, 192);
            this.txtFilterList.TabIndex = 3;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 8);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(78, 13);
            this.label8.TabIndex = 2;
            this.label8.Text = "Filter hierarchy:";
            // 
            // tabDefaults
            // 
            this.tabDefaults.Controls.Add(this.chkRefreshDataWhenOpeningTheFile);
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
            // chkRefreshDataWhenOpeningTheFile
            // 
            this.chkRefreshDataWhenOpeningTheFile.AutoSize = true;
            this.chkRefreshDataWhenOpeningTheFile.CheckAlign = System.Drawing.ContentAlignment.TopLeft;
            this.chkRefreshDataWhenOpeningTheFile.Location = new System.Drawing.Point(13, 57);
            this.chkRefreshDataWhenOpeningTheFile.Name = "chkRefreshDataWhenOpeningTheFile";
            this.chkRefreshDataWhenOpeningTheFile.Size = new System.Drawing.Size(290, 17);
            this.chkRefreshDataWhenOpeningTheFile.TabIndex = 3;
            this.chkRefreshDataWhenOpeningTheFile.Text = "Turn on \"Refresh data when opening the file\" by default";
            this.chkRefreshDataWhenOpeningTheFile.UseVisualStyleBackColor = true;
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
            // tabAbout
            // 
            this.tabAbout.Controls.Add(this.btnUpgradeOnRefresh);
            this.tabAbout.Controls.Add(this.lblExcelUILanguage);
            this.tabAbout.Controls.Add(this.lblUpgradePivotTableInstructions);
            this.tabAbout.Controls.Add(this.linkUpgradePivotTable);
            this.tabAbout.Controls.Add(this.lblPivotTableVersion);
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
            // btnUpgradeOnRefresh
            // 
            this.btnUpgradeOnRefresh.Location = new System.Drawing.Point(15, 224);
            this.btnUpgradeOnRefresh.Name = "btnUpgradeOnRefresh";
            this.btnUpgradeOnRefresh.Size = new System.Drawing.Size(192, 23);
            this.btnUpgradeOnRefresh.TabIndex = 7;
            this.btnUpgradeOnRefresh.Text = "Set UpgradeOnRefresh to True";
            this.btnUpgradeOnRefresh.UseVisualStyleBackColor = true;
            this.btnUpgradeOnRefresh.Click += new System.EventHandler(this.btnUpgradeOnRefresh_Click);
            // 
            // lblExcelUILanguage
            // 
            this.lblExcelUILanguage.AutoSize = true;
            this.lblExcelUILanguage.Location = new System.Drawing.Point(12, 95);
            this.lblExcelUILanguage.Name = "lblExcelUILanguage";
            this.lblExcelUILanguage.Size = new System.Drawing.Size(101, 13);
            this.lblExcelUILanguage.TabIndex = 6;
            this.lblExcelUILanguage.Text = "Excel UI Language:";
            // 
            // lblUpgradePivotTableInstructions
            // 
            this.lblUpgradePivotTableInstructions.AutoSize = true;
            this.lblUpgradePivotTableInstructions.Location = new System.Drawing.Point(11, 207);
            this.lblUpgradePivotTableInstructions.Name = "lblUpgradePivotTableInstructions";
            this.lblUpgradePivotTableInstructions.Size = new System.Drawing.Size(262, 13);
            this.lblUpgradePivotTableInstructions.TabIndex = 5;
            this.lblUpgradePivotTableInstructions.Text = "To upgrade, save as .xlsx then refresh the PivotTable.";
            // 
            // linkUpgradePivotTable
            // 
            this.linkUpgradePivotTable.AutoSize = true;
            this.linkUpgradePivotTable.Cursor = System.Windows.Forms.Cursors.Hand;
            this.linkUpgradePivotTable.Location = new System.Drawing.Point(12, 188);
            this.linkUpgradePivotTable.Name = "linkUpgradePivotTable";
            this.linkUpgradePivotTable.Size = new System.Drawing.Size(182, 13);
            this.linkUpgradePivotTable.TabIndex = 4;
            this.linkUpgradePivotTable.TabStop = true;
            this.linkUpgradePivotTable.Text = "How to Upgrade PivotTable Versions";
            this.linkUpgradePivotTable.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkUpgradePivotTable_LinkClicked);
            // 
            // lblPivotTableVersion
            // 
            this.lblPivotTableVersion.AutoSize = true;
            this.lblPivotTableVersion.Location = new System.Drawing.Point(11, 171);
            this.lblPivotTableVersion.Name = "lblPivotTableVersion";
            this.lblPivotTableVersion.Size = new System.Drawing.Size(161, 13);
            this.lblPivotTableVersion.TabIndex = 3;
            this.lblPivotTableVersion.Text = "Version of This PivotTable: 2007";
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
            // tooltip
            // 
            this.tooltip.AutoPopDelay = 5000;
            this.tooltip.InitialDelay = 500;
            this.tooltip.ReshowDelay = 100;
            this.tooltip.ShowAlways = true;
            // 
            // lblFormattingMdxQuery
            // 
            this.lblFormattingMdxQuery.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblFormattingMdxQuery.AutoSize = true;
            this.lblFormattingMdxQuery.Location = new System.Drawing.Point(10, 281);
            this.lblFormattingMdxQuery.Name = "lblFormattingMdxQuery";
            this.lblFormattingMdxQuery.Size = new System.Drawing.Size(175, 13);
            this.lblFormattingMdxQuery.TabIndex = 6;
            this.lblFormattingMdxQuery.Text = "Formatting MDX query in progress...";
            this.lblFormattingMdxQuery.Visible = false;
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
            this.Activated += new System.EventHandler(this.MainForm_Activated);
            this.tabControl.ResumeLayout(false);
            this.tabCalcs.ResumeLayout(false);
            this.tabCalcs.PerformLayout();
            this.tabLibrary.ResumeLayout(false);
            this.tabLibrary.PerformLayout();
            this.tabMDX.ResumeLayout(false);
            this.tabMDX.PerformLayout();
            this.tabSearch.ResumeLayout(false);
            this.tabSearch.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridSearchResults)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cubeSearchMatchBindingSource)).EndInit();
            this.tabFilterList.ResumeLayout(false);
            this.tabFilterList.PerformLayout();
            this.tabDefaults.ResumeLayout(false);
            this.tabDefaults.PerformLayout();
            this.tabAbout.ResumeLayout(false);
            this.tabAbout.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tabMDX;
        private System.Windows.Forms.ComboBox comboCalcName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtCalcFormula;
        private System.Windows.Forms.Button btnAddCalc;
        private System.Windows.Forms.Button btnDeleteCalc;
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
        private System.Windows.Forms.Button btnFind;
        private System.Windows.Forms.Label lblSearchFor;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.ComboBox cmbLookIn;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.CheckBox chkExactMatch;
        private System.Windows.Forms.CheckBox chkMemberProperties;
        private System.Windows.Forms.Button btnSearchAdd;
        private System.Windows.Forms.Label lblNoSearchMatches;
        private System.Windows.Forms.ProgressBar prgSearch;
        private System.Windows.Forms.Button btnCancelSearch;
        private System.Windows.Forms.Label lblSearchError;
        private System.Windows.Forms.TabControl tabControl;
        public System.Windows.Forms.TabPage tabSearch;
        public System.Windows.Forms.TabPage tabCalcs;
        private System.Windows.Forms.DataGridView dataGridSearchResults;
        private System.Windows.Forms.BindingSource cubeSearchMatchBindingSource;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colCheck;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colType;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFolder;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDesc;
        private System.Windows.Forms.TabPage tabFilterList;
        private System.Windows.Forms.ComboBox cmbFilterListLookIn;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btnFilterList;
        private System.Windows.Forms.Button btnCancelFilterList;
        private System.Windows.Forms.ProgressBar progressFilterList;
        private System.Windows.Forms.Label lblFilterListError;
        private System.Windows.Forms.CheckBox chkRefreshDataWhenOpeningTheFile;
        private System.Windows.Forms.TextBox txtFilterList;
        private System.Windows.Forms.CheckBox chkAddToCurrentFilters;
        private System.Windows.Forms.Label lblPivotTableVersion;
        private System.Windows.Forms.LinkLabel linkUpgradePivotTable;
        private System.Windows.Forms.Label lblUpgradePivotTableInstructions;
        private System.Windows.Forms.Button btnFilterListShowCurrentFilters;
        private System.Windows.Forms.Label lblExcelUILanguage;
        private System.Windows.Forms.Button btnUpgradeOnRefresh;
        private System.Windows.Forms.RichTextBox richTextBoxMDX;
        private System.Windows.Forms.CheckBox chkFormatMDX;
        private System.Windows.Forms.ToolTip tooltip;
        private System.Windows.Forms.Label lblFormattingMdxQuery;
    }
}