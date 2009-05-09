using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.AnalysisServices.AdomdClient;

namespace OlapPivotTableExtensions
{
    public partial class MainForm : Form
    {
        private Excel.PivotTable pvt;
        private Excel.Application application;
        private CalculationsLibrary library;
        private AdomdConnection connCube;
        private CubeDef cube;
        private CubeSearcher searcher;
        public bool AddInWorking = false;
        private BackgroundWorker workerFilterList;

        private int _LibraryComboDividerItemIndex = int.MaxValue;


        public MainForm(Excel.Application app)
        {
            InitializeComponent();

            try
            {
                System.Reflection.AssemblyFileVersionAttribute attrVersion = (System.Reflection.AssemblyFileVersionAttribute)typeof(MainForm).Assembly.GetCustomAttributes(typeof(System.Reflection.AssemblyFileVersionAttribute), true)[0];
                lblVersion.Text = "OLAP PivotTable Extensions v" + attrVersion.Version;

                application = app;
                pvt = app.ActiveCell.PivotTable;

                library = new CalculationsLibrary();
                library.Load();

                FillCalcsDropdown();

                tabControl.SelectedTab = tabCalcs;

                chkShowCalcMembers.Checked = ThisAddIn.ShowCalcMembersByDefault;
                chkRefreshDataWhenOpeningTheFile.Checked = ThisAddIn.RefreshDataByDefault;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "OLAP PivotTable Extensions");
                this.Visible = false;
                this.Close();
            }
        }

        private void SetMDX()
        {
            StringBuilder sMdxQuery = new StringBuilder(pvt.MDX);

            //add (session) calculated members to the query so that you can run it from SSMS
            if (pvt.CalculatedMembers.Count > 0)
            {
                StringBuilder sCalcs = new StringBuilder();
                foreach (Excel.CalculatedMember calc in pvt.CalculatedMembers)
                {
                    sCalcs.AppendFormat("MEMBER {0} as {1}\r\n", calc.Name, calc.Formula.Replace("\r\n","\r").Replace("\r","\r\n")); //normalize the line breaks which have been turned into \r to workaround an Excel Services bug
                }
                if (sMdxQuery.ToString().StartsWith("with", StringComparison.CurrentCultureIgnoreCase))
                {
                    sMdxQuery.Insert(5, sCalcs.ToString());
                }
                else
                {
                    sCalcs.Insert(0, "WITH\r\n");
                    sMdxQuery.Insert(0, sCalcs.ToString());
                }
            }

            txtMDX.Text = sMdxQuery.ToString();
            txtMDX.SelectionStart = 0;
            txtMDX.SelectionLength = sMdxQuery.Length;
            txtMDX.Focus();
        }

        private void btnDeleteCalc_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.CalculatedMember oCalcMember = GetCalculatedMember(comboCalcName.Text);
                if (oCalcMember != null)
                {
                    oCalcMember.Delete();
                    pvt.RefreshTable();
                }
                FillCalcsDropdown();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "OLAP PivotTable Extensions");
            }
        }

        private void btnAddCalc_Click(object sender, EventArgs e)
        {
            try
            {
                bool bMeasure = true;
                string sName = comboCalcName.Text;
                string sFormula = txtCalcFormula.Text;
                if (!sName.StartsWith("[") && !sName.StartsWith("[Measures].", StringComparison.CurrentCultureIgnoreCase))
                {
                    sName = "[Measures].[" + sName.Replace("]", "]]") + "]";
                }
                else if (sName.StartsWith("[") && !sName.StartsWith("[Measures].", StringComparison.CurrentCultureIgnoreCase))
                {
                    bMeasure = false;
                }

                try
                {
                    library.AddCalculation(sName, sFormula);
                    library.Save();
                }
                catch (Exception ex)
                {
                    throw new Exception("There was a problem saving this calculation to the library at " + CalculationsLibrary.LibraryPath + ". " + ex.Message, ex);
                }

                Excel.CalculatedMember oCalcMember = GetCalculatedMember(sName);
                if (oCalcMember != null)
                    oCalcMember.Delete();

                try
                {
                    //replace the line breaks in the formula we save to the PivotTable to workaround a bug in Excel Services: http://www.codeplex.com/OlapPivotTableExtend/Thread/View.aspx?ThreadId=41697
                    oCalcMember = pvt.CalculatedMembers.Add(sName, sFormula.Replace("\r\n","\r"), System.Reflection.Missing.Value, Excel.XlCalculatedMemberType.xlCalculatedMember);
                    if (bMeasure)
                    {
                        pvt.RefreshTable();
                        pvt.CubeFields.get_Item(sName).Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                    }
                    else
                    {
                        pvt.ViewCalculatedMembers = true;
                        pvt.RefreshTable();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("There was a problem creating the calculation:\r\n" + ex.Message, "OLAP PivotTable Extensions");
                }

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an unexpected error creating the calculation:\r\n" + ex.Message, "OLAP PivotTable Extensions");
            }
        }

        private void FillCalcsDropdown()
        {
            comboCalcName.Items.Clear();
            List<string> listCalcs = new List<string>();
            foreach (Excel.CalculatedMember calc in pvt.CalculatedMembers)
            {
                listCalcs.Add(calc.Name);
            }
            listCalcs.Sort();

            foreach (string calc in listCalcs)
            {
                comboCalcName.Items.Add(calc);
            }

            comboCalcName.Items.Add(string.Empty);
            if (library.Calculations.Length > 0)
            {
                _LibraryComboDividerItemIndex = comboCalcName.Items.Add("---CALCULATIONS LIBRARY---");

                foreach (CalculationsLibrary.Calculation c in library.Calculations)
                {
                    comboCalcName.Items.Add(c.Name);
                }
            }

            comboCalcName.Text = string.Empty;
            comboCalcName.Focus();
            txtCalcFormula.Text = string.Empty;
        }

        //returns the calc member if it exists
        private Excel.CalculatedMember GetCalculatedMember(string sName)
        {
            try
            {
                return pvt.CalculatedMembers.get_Item(sName);
            }
            catch
            {
                return null;
            }
        }

        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl.SelectedTab == tabMDX)
            {
                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    SetMDX();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("There was a problem capturing the MDX query for this PivotTable.\r\n" + ex.Message, "OLAP PivotTable Extensions");
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
            else if (tabControl.SelectedTab == tabSearch)
            {
                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    SetupSearchTab();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("There was a problem setting up the search tab.\r\n" + ex.Message, "OLAP PivotTable Extensions");
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
            else if (tabControl.SelectedTab == tabFilterList)
            {
                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    SetupFilterListTab("");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("There was a problem setting up the Filter List tab.\r\n" + ex.Message, "OLAP PivotTable Extensions");
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        public void SetupFilterListTab(string SelectedLookIn)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                tabControl.SelectedTab = tabFilterList;

                //fill in the "Look in" dropdown with the dimension hierarchies in the PivotTable
                cmbFilterListLookIn.SuspendLayout();
                cmbFilterListLookIn.Items.Clear();

                foreach (Excel.CubeField f in pvt.CubeFields)
                {
                    if (f.Orientation != Excel.XlPivotFieldOrientation.xlHidden && f.CubeFieldType == Excel.XlCubeFieldType.xlHierarchy) //not named sets since you can't filter them, and not measures
                    {
                        cmbFilterListLookIn.Items.Add(f.Name);
                    }
                }

                if (!string.IsNullOrEmpty(SelectedLookIn))
                {
                    cmbFilterListLookIn.SelectedItem = SelectedLookIn;
                }

                if (!IsExcel2007OrHigherPivotTableVersion())
                {
                    lblFilterListError.Text = "Upgrade PivotTable to Excel 2007 to use Filter List";
                    lblFilterListError.Visible = true;
                    btnFilterList.Enabled = false;
                    txtFilterList.Enabled = false;
                }

                cmbFilterListLookIn.ResumeLayout();
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        public void SetupSearchTab(string SelectedLookIn)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                tabControl.SelectedTab = tabSearch;
                SetupSearchTab();
                cmbLookIn.SelectedItem = SelectedLookIn;
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void SetupSearchTab()
        {
            if (cmbLookIn.SelectedItem == null)
            {
                cmbLookIn.SelectedIndex = 0;
                cmbLookIn_SelectedIndexChanged(null, null);
            }

            txtSearch.Focus();
            txtSearch.SelectAll();

            Application.DoEvents();

            ConnectAdomdClientCube();
        }

        private void ConnectAdomdClientCube()
        {
            Microsoft.Office.Interop.Excel.PivotCache cache = pvt.PivotCache();
            if (!cache.IsConnected)
                cache.MakeConnection();

            ADODB.Connection connADO = cache.ADOConnection as ADODB.Connection;
            if (connADO == null) throw new Exception("Could not cast PivotCache.ADOConnection to ADODB.Connection.");

            string sConnectionString = connADO.ConnectionString;

            Excel.OLEDBConnection connOLEDB = cache.WorkbookConnection.OLEDBConnection;

            //figure out current locale
            if (connOLEDB.RetrieveInOfficeUILang && !sConnectionString.ToLower().Contains("language identifier=") && !sConnectionString.ToLower().Contains("localeidentifier="))
                sConnectionString += ";LocaleIdentifier=" + this.application.LanguageSettings.get_LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI);

            sConnectionString += ";Application Name=" + lblVersion.Text;

            //look for impersonation info so we can mimic what Excel does
            bool bImpersonate = false;
            string sUsername = null;
            string sDomain = null;
            string sPassword = null;
            ConnectionStringParser connParser = new ConnectionStringParser(sConnectionString);
            if ((connParser.ContainsKey("User Id") || connParser.ContainsKey("Uid")) && connParser.ContainsKey("Password"))
            {
                bImpersonate = true;
                if (connParser.ContainsKey("User Id"))
                    sUsername = connParser["User Id"];
                else
                    sUsername = connParser["Uid"];
                sPassword = connParser["Password"];
                int iSlashIndex = sUsername.IndexOf('\\');
                if (iSlashIndex >= 0)
                {
                    sDomain = sUsername.Split('\\')[0];
                    sUsername = sUsername.Split('\\')[1];
                }
                else
                {
                    sDomain = connParser["Data Source"];
                    if (sDomain.ToLower().StartsWith("http://") || sDomain.ToLower().StartsWith("https://"))
                    {
                        MessageBox.Show("Please specify the domain name for the username in the connection string. Syntax: User Id=DOMAIN\\Username", "OLAP PivotTable Extensions");
                        this.Visible = false;
                        this.Close();
                        return;
                    }
                }
            }

            if (connCube == null)
            {
                if (!IsExcel2007OrHigherPivotTableVersion())
                {
                    lblSearchError.Text = "Upgrade PivotTable to Excel 2007 for full support";
                    lblSearchError.Visible = true;
                }

                if (bImpersonate)
                {
                    using (new Impersonator(sUsername, sDomain, sPassword))
                    {
                        try
                        {
                            connCube = new AdomdConnection(sConnectionString);
                            connCube.Open();
                        }
                        catch (ArgumentException ex)
                        {
                            //may be that you can't use Integrated Security=SSPI with an HTTP or HTTPS connection... try to workaround that
                            if (sConnectionString.ToLower().IndexOf("data source=http") >= 0 && sConnectionString.ToLower().IndexOf("integrated security=sspi;") >= 0)
                            {
                                sConnectionString = sConnectionString.Remove(sConnectionString.ToLower().IndexOf("integrated security=sspi;"), "integrated security=sspi;".Length);
                                connCube = new AdomdConnection(sConnectionString);
                                connCube.Open();
                            }
                            else
                            {
                                throw ex;
                            }
                        }
                    }
                }
                else
                {
                    try
                    {
                        connCube = new AdomdConnection(sConnectionString);
                        connCube.Open();
                    }
                    catch (ArgumentException ex)
                    {
                        //may be that you can't use Integrated Security=SSPI with an HTTP or HTTPS connection... try to workaround that
                        if (sConnectionString.ToLower().IndexOf("data source=http") >= 0 && sConnectionString.ToLower().IndexOf("integrated security=sspi;") >= 0)
                        {
                            sConnectionString = sConnectionString.Remove(sConnectionString.ToLower().IndexOf("integrated security=sspi;"), "integrated security=sspi;".Length);
                            connCube = new AdomdConnection(sConnectionString);
                            connCube.Open();
                        }
                        else
                        {
                            throw ex;
                        }
                    }
                }

                cube = connCube.Cubes.Find(Convert.ToString(cache.CommandText));
                if (cube == null)
                {
                    throw new Exception("Could not find cube [" + Convert.ToString(cache.CommandText) + "]");
                }

                //fill in the "Look in" dropdown with the dimension hierarchies in the PivotTable
                cmbLookIn.SuspendLayout();
                while (cmbLookIn.Items.Count > 2)
                    cmbLookIn.Items.RemoveAt(2);
                foreach (Excel.CubeField f in pvt.CubeFields)
                {
                    if (f.CubeFieldType == Excel.XlCubeFieldType.xlHierarchy) //not named sets since you can't filter them, and not measures since they are returned in the field list search
                    {
                        if (f.Orientation != Excel.XlPivotFieldOrientation.xlHidden)
                        {
                            cmbLookIn.Items.Add(f.Name);
                        }
                    }
                }
                cmbLookIn.ResumeLayout();
            }
            else
            {
                try
                {
                    int iCubesCount = connCube.Cubes.Count; //hitting this property will help AdomdClient detect if the connection has been dropped
                }
                catch (AdomdConnectionException) { } //we expect this exception if the connection has been dropped
                catch (Exception) { }

                if (connCube.State != ConnectionState.Open)
                {
                    if (searcher != null)
                    {
                        MessageBox.Show("The connection was dropped. The search results are now invalid and cannot be used. Please close the OLAP PivotTable Extensions window.", "OLAP PivotTable Extensions");
                        return;
                    }

                    if (bImpersonate)
                    {
                        using (new Impersonator(sUsername, sDomain, sPassword))
                        {
                            connCube.Open();
                        }
                    }
                    else
                    {
                        connCube.Open();
                    }

                    cube = connCube.Cubes.Find(Convert.ToString(cache.CommandText));
                    if (cube == null)
                    {
                        throw new Exception("Could not find cube [" + Convert.ToString(cache.CommandText) + "]");
                    }
                }
            }

        }

        private void comboCalcName_TextChanged(object sender, EventArgs e)
        {
            if (comboCalcName.SelectedIndex == _LibraryComboDividerItemIndex)
            {
                comboCalcName.Text = string.Empty;
                btnDeleteCalc.Enabled = false;
            }
            else if (comboCalcName.SelectedIndex > _LibraryComboDividerItemIndex)
            {
                CalculationsLibrary.Calculation c = library.GetCalculation(comboCalcName.Text);
                if (c != null)
                {
                    txtCalcFormula.Text = c.Formula;
                }
                btnDeleteCalc.Enabled = false;
            }
            else
            {
                Excel.CalculatedMember oCalcMember = GetCalculatedMember(comboCalcName.Text);
                if (oCalcMember != null)
                {
                    txtCalcFormula.Text = oCalcMember.Formula.Replace("\r\n", "\r").Replace("\r", "\r\n"); //normalize the line breaks which have been turned into \r to workaround an Excel Services bug
                    btnDeleteCalc.Enabled = true;
                }
                else
                {
                    btnDeleteCalc.Enabled = false;
                }
            }
        }

        private void linkCodeplex_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.codeplex.com/OlapPivotTableExtend");
        }

        private void linkHelp_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.codeplex.com/OlapPivotTableExtend/Wiki/View.aspx?title=Calculations%20Help");
        }

        private void radioExport_CheckedChanged(object sender, EventArgs e)
        {
            if (radioExport.Checked)
            {
                listImportExportCalcs.Items.Clear();
                foreach (CalculationsLibrary.Calculation c in library.Calculations)
                {
                    listImportExportCalcs.Items.Add(c.Name, true);
                }

                listImportExportCalcs.Enabled = true;
                btnImportExportExecute.Enabled = true;
            }
        }

        private void radDelete_CheckedChanged(object sender, EventArgs e)
        {
            if (radDelete.Checked)
            {
                listImportExportCalcs.Items.Clear();
                foreach (CalculationsLibrary.Calculation c in library.Calculations)
                {
                    listImportExportCalcs.Items.Add(c.Name, false);
                }

                listImportExportCalcs.Enabled = true;
                btnImportExportExecute.Enabled = true;
            }
        }

        private void btnImportExportExecute_Click(object sender, EventArgs e)
        {
            try
            {
                if (radImport.Checked)
                {
                    CalculationsLibrary libraryImportExport = new CalculationsLibrary();
                    libraryImportExport.Load(txtImportFilePath.Text);
                    foreach (CalculationsLibrary.Calculation c in libraryImportExport.Calculations)
                    {
                        if (listImportExportCalcs.CheckedItems.Contains(c.Name))
                        {
                            library.AddCalculation(c.Name, c.Formula);
                        }
                    }
                    library.Save();
                }
                else if (radioExport.Checked)
                {
                    CalculationsLibrary libraryImportExport = new CalculationsLibrary();
                    List<CalculationsLibrary.Calculation> calcs = new List<CalculationsLibrary.Calculation>();
                    foreach (CalculationsLibrary.Calculation c in library.Calculations)
                    {
                        if (listImportExportCalcs.CheckedItems.Contains(c.Name))
                        {
                            calcs.Add(c);
                        }
                    }
                    libraryImportExport.Calculations = calcs.ToArray();
                    libraryImportExport.Save(txtExportFilePath.Text);
                    MessageBox.Show("Export completed successfully.", "OLAP PivotTable Extensions");
                    return;
                }
                else if (radDelete.Checked)
                {
                    foreach (CalculationsLibrary.Calculation c in library.Calculations)
                    {
                        if (listImportExportCalcs.CheckedItems.Contains(c.Name))
                        {
                            library.DeleteCalculation(c.Name);
                        }
                    }
                    library.Save();
                }

                radImport.Checked = true;
                listImportExportCalcs.Items.Clear();

                FillCalcsDropdown();
                tabControl.SelectedTab = tabCalcs;
                comboCalcName.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "OLAP PivotTable Extensions");
            }
        }

        private void btnImportFilePath_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Title = "Choose Calculation Library To Import...";
                dlg.Filter = "Calculation Library (*.xml)|*.xml";
                dlg.CheckFileExists = true;
                dlg.Multiselect = false;
                dlg.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    txtImportFilePath.Text = dlg.FileName;
                    radImport_CheckedChanged(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "OLAP PivotTable Extensions");
            }
        }

        private void btnExportFilePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Export Calculations To...";
            dlg.Filter = "Calculation Library (*.xml)|*.xml";
            dlg.CheckFileExists = false;
            dlg.Multiselect = false;
            dlg.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                this.txtExportFilePath.Text = dlg.FileName;
            }
        }

        private void radImport_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(txtImportFilePath.Text))
                {
                    CalculationsLibrary libraryImportExport = new CalculationsLibrary();
                    libraryImportExport.Load(txtImportFilePath.Text);
                    listImportExportCalcs.Items.Clear();
                    foreach (CalculationsLibrary.Calculation c in libraryImportExport.Calculations)
                    {
                        listImportExportCalcs.Items.Add(c.Name, true);
                    }

                    listImportExportCalcs.Enabled = true;
                    btnImportExportExecute.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was a problem loading that XML file: " + ex.Message, "OLAP PivotTable Extensions");
            }
        }

        private void btnSaveDefaults_Click(object sender, EventArgs e)
        {
            try
            {
                ThisAddIn.ShowCalcMembersByDefault = chkShowCalcMembers.Checked;
                ThisAddIn.RefreshDataByDefault = this.chkRefreshDataWhenOpeningTheFile.Checked;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "OLAP PivotTable Extensions");
            }
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtSearch.Text == string.Empty && !chkExactMatch.Checked && cmbLookIn.SelectedIndex == 1)
                {
                    MessageBox.Show("Please enter a search term", "OLAP PivotTable Extensions");
                    return;
                }

                dataGridSearchResults.Enabled = true;
                colName.SortMode = DataGridViewColumnSortMode.NotSortable;
                colType.SortMode = DataGridViewColumnSortMode.NotSortable;
                colFolder.SortMode = DataGridViewColumnSortMode.NotSortable;
                colDesc.SortMode = DataGridViewColumnSortMode.NotSortable;

                lblNoSearchMatches.Visible = false;
                lblSearchError.Visible = false;

                btnFind.Enabled = false;
                btnCancelSearch.Visible = true;
                btnSearchAdd.Enabled = false;

                Application.DoEvents();

                prgSearch.Value = 0;
                prgSearch.Visible = true;

                CubeSearcher.CubeSearchScope scope = CubeSearcher.CubeSearchScope.FieldList;
                if (cmbLookIn.SelectedIndex >= 1)
                    scope = CubeSearcher.CubeSearchScope.DimensionData;

                string sSearchOnlyDimension = null;
                if (cmbLookIn.SelectedIndex > 1)
                    sSearchOnlyDimension = Convert.ToString(cmbLookIn.SelectedItem);

                //TODO: bold any search results which are already in the PivotTable... may need to pass in a delegate to update the CubeSearchMatch object to have a reference to the parent class

                searcher = new CubeSearcher(cube, scope, txtSearch.Text, chkExactMatch.Checked, chkMemberProperties.Checked, sSearchOnlyDimension, dataGridSearchResults);
                searcher.ProgressChanged += new ProgressChangedEventHandler(searcher_ProgressChanged);
                searcher.SearchAsync();

                cubeSearchMatchBindingSource.DataSource = searcher.Matches;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during search: " + ex.Message, "OLAP PivotTable Extensions");
            }
        }

        private void dataGridSearchResults_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 0 && e.RowIndex >= 0 && searcher != null && !searcher.Complete)
                {
                    //the control doesn't do too well while we're constantly setting a new DataSource... this code worksaround that problem so you can check/uncheck while the search is still going on
                    searcher.Matches[e.RowIndex].Checked = !searcher.Matches[e.RowIndex].Checked;
                    dataGridSearchResults.InvalidateCell(e.ColumnIndex, e.RowIndex);
                    dataGridSearchResults.Refresh();
                }
            }
            catch { }
        }

        private delegate void searcher_ProgressChanged_Delegate(object sender, ProgressChangedEventArgs e);
        private void searcher_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try
            {
                if (prgSearch.InvokeRequired)
                {
                    //avoid the "cross-thread operation not valid" error message
                    prgSearch.BeginInvoke(new searcher_ProgressChanged_Delegate(searcher_ProgressChanged), new object[] { sender, e });
                }
                else
                {
                    prgSearch.Value = e.ProgressPercentage;

                    if (searcher.Complete)
                    {
                        lblNoSearchMatches.Visible = (searcher.Matches.Count == 0);
                        btnSearchAdd.Enabled = (searcher.Matches.Count > 0);
                        btnFind.Enabled = true;

                        prgSearch.Visible = false;
                        btnCancelSearch.Visible = false;

                        colName.SortMode = DataGridViewColumnSortMode.Automatic;
                        colType.SortMode = DataGridViewColumnSortMode.Automatic;
                        colFolder.SortMode = DataGridViewColumnSortMode.Automatic;
                        colDesc.SortMode = DataGridViewColumnSortMode.Automatic;

                        if (!string.IsNullOrEmpty(searcher.Error))
                        {
                            lblSearchError.Text = searcher.Error;
                            lblSearchError.Visible = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during update of search progress: " + ex.Message, "OLAP PivotTable Extensions");
            }
        }

        private void txtSearch_Enter(object sender, EventArgs e)
        {
            try
            {
                this.AcceptButton = btnFind;
            }
            catch { }
        }

        private void txtSearch_Leave(object sender, EventArgs e)
        {
            try
            {
                this.AcceptButton = null;
            }
            catch { }
        }

        private void btnSearchAdd_Click(object sender, EventArgs e)
        {
            //TODO: future: when you've selected multiple members from a hierarchy that's going to be a filter, decide how to handle that better... test this
            string sSearchFor = null;
            try
            {
                AddInWorking = true;

                Dictionary<string, NamedSet> hierarchiesInPivotTableAsNamedSet = new Dictionary<string, NamedSet>(StringComparer.CurrentCultureIgnoreCase);
                foreach (Excel.CubeField cf in pvt.CubeFields)
                {
                    if (cf.CubeFieldType == Excel.XlCubeFieldType.xlSet && cf.Orientation != Excel.XlPivotFieldOrientation.xlHidden)
                    {
                        NamedSet ns = cube.NamedSets.Find(cf.Name.Substring(1, cf.Name.Length - 2));
                        if (ns != null)
                        {
                            hierarchiesInPivotTableAsNamedSet.Add(Convert.ToString(ns.Properties["DIMENSIONS"].Value), ns);
                        }
                    }
                }

                foreach (CubeSearcher.CubeSearchMatch item in searcher.Matches)
                {
                    if (!item.Checked) continue;
                    if (item.IsFieldListField)
                    {
                        Excel.CubeField field;
                        if (item.InnerObject is Dimension)
                        {
                            Dimension d = (Dimension)item.InnerObject;
                            sSearchFor = d.UniqueName;
                            string sDefaultHierarchy = Convert.ToString(d.Properties["DEFAULT_HIERARCHY"].Value);
                            sSearchFor = sDefaultHierarchy;
                            field = pvt.CubeFields.get_Item(sSearchFor);
                        }
                        else if (item.InnerObject is Level)
                        {
                            Level l = (Level)item.InnerObject;
                            sSearchFor = l.ParentHierarchy.UniqueName;
                            field = pvt.CubeFields.get_Item(sSearchFor);
                        }
                        else if (item.InnerObject is Kpi)
                        {
                            Kpi k = (Kpi)item.InnerObject;
                            sSearchFor = k.Name;
                            PivotTableKpiUtility.AddKpiToPivotTable(k, pvt);
                            continue;
                        }
                        else if (item.InnerObject is NamedSet)
                        {
                            NamedSet s = (NamedSet)item.InnerObject;
                            sSearchFor = s.Name;
                            field = pvt.CubeFields.get_Item("[" + s.Name + "]");
                            field.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                            continue;
                        }
                        else
                        {
                            sSearchFor = item.UniqueName;
                            field = pvt.CubeFields.get_Item(sSearchFor);
                        }

                        if (field.Orientation == Excel.XlPivotFieldOrientation.xlHidden)
                        {
                            if (item.InnerObject is Measure)
                                field.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                            else if (item.MemberProperty != null)
                                field.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                            else
                                field.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                        }
                        if (item.MemberProperty != null && item.InnerObject is Level)
                        {
                            Level l = (Level)item.InnerObject;
                            try
                            {
                                //try to add to all levels... this works for balanced hierarchies but fails on unbalanced ones: http://msdn.microsoft.com/en-us/library/bb209584.aspx
                                field.AddMemberPropertyField(l.ParentHierarchy.UniqueName + ".[" + item.MemberProperty.Name + "]", System.Type.Missing);
                            }
                            catch (Exception ex)
                            {
                                try
                                {
                                    //try to add to just this level... this works on unbalanced hierarchies
                                    field.AddMemberPropertyField(item.MemberProperty.UniqueName, System.Type.Missing);
                                }
                                catch
                                {
                                    //if neither succeeded, then raise the error
                                    throw ex;
                                }
                            }
                        }
                    }
                    else //it's a member
                    {
                        if (item.IsCalculated)
                        {
                            //make sure the PivotTable has "show calculated members" on
                            pvt.ViewCalculatedMembers = true;
                        }

                        Member m = (Member)item.InnerObject;
                        Excel.CubeField field;
                        sSearchFor = m.Caption + " (" + m.UniqueName + ")";
                        if (hierarchiesInPivotTableAsNamedSet.ContainsKey(m.ParentLevel.ParentHierarchy.UniqueName))
                        {
                            NamedSet ns = hierarchiesInPivotTableAsNamedSet[m.ParentLevel.ParentHierarchy.UniqueName];
                            field = pvt.CubeFields.get_Item("[" + ns.Name + "]");
                            if (string.Compare("[" + ns.Name + "]", Convert.ToString(cmbLookIn.SelectedItem)) != 0)
                            {
                                //TODO: future... see if it's in that set so you don't have to prompt
                                MessageBox.Show("The named set [" + Convert.ToString(ns.Properties["SET_CAPTION"].Value) + "] containing " + m.ParentLevel.ParentHierarchy.UniqueName + " is in the PivotTable, so [" + m.Caption + "] will not show up in the PivotTable unless it is in that set.", "OLAP PivotTable Extensions");
                                continue;
                            }
                        }
                        else
                        {
                            field = pvt.CubeFields.get_Item(m.ParentLevel.ParentHierarchy.UniqueName);
                        }
                        if (field.Orientation == Excel.XlPivotFieldOrientation.xlHidden)
                        {
                            if (item.MemberProperty != null)
                                field.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                            else
                                field.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                        }
                        if (field.Orientation == Excel.XlPivotFieldOrientation.xlPageField)
                        {
                            field.CurrentPageName = m.UniqueName;
                        }
                        else
                        {
                            try
                            {
                                field.CreatePivotFields(); //Excel apparently doesn't always have the levels loaded, so this loads them
                            }
                            catch { }

                            //TODO: future... clear other filters (like value filters) so that the member you're wanting will show up?
                            EnsureMemberIsVisible(field, m, true);
                            if (item.MemberProperty != null)
                            {
                                try
                                {
                                    //try to add to all levels... this works for balanced hierarchies but fails on unbalanced ones: http://msdn.microsoft.com/en-us/library/bb209584.aspx
                                    field.AddMemberPropertyField(m.ParentLevel.ParentHierarchy.UniqueName + ".[" + item.MemberProperty.Name + "]", System.Type.Missing);
                                }
                                catch (Exception ex)
                                {
                                    try
                                    {
                                        //try to add to just this level... this works on unbalanced hierarchies
                                        field.AddMemberPropertyField(item.MemberProperty.UniqueName, System.Type.Missing);
                                    }
                                    catch
                                    {
                                        //if neither succeeded, then raise the error
                                        throw ex;
                                    }
                                }
                            }
                        }
                    }
                }

                this.Visible = false;
                this.Close();
            }
            catch (Exception ex)
            {
                if (string.IsNullOrEmpty(sSearchFor))
                    MessageBox.Show("Problem adding to PivotTable: " + ex.Message, "OLAP PivotTable Extensions");
                else
                    MessageBox.Show("Problem adding " + sSearchFor + " to PivotTable: " + ex.Message, "OLAP PivotTable Extensions");
            }
            finally
            {
                AddInWorking = false;
            }
        }

        private void EnsureMemberIsVisible(Excel.CubeField field, Member m, bool showInAxis)
        {
            if (field.CubeFieldType != Excel.XlCubeFieldType.xlHierarchy) return;

            //ensure parents are visible
            if (m.Parent != null && m.Parent.ParentLevel.LevelType != LevelTypeEnum.All)
                EnsureMemberIsVisible(field, m.Parent, false);

            Excel.PivotField pivotField = (Excel.PivotField)field.PivotFields.Item(m.ParentLevel.UniqueName);

            if (IsExcel2007OrHigherPivotTableVersion())
            {
                //the PivotField.Hidden and PivotField.VisibleItemsList properties weren't added until Excel 2007 version PivotTables... not sure what the old equivalent is... oh well... this EnsureMemberIsVisible function still works unless you've filtered out that member explictly
                try
                {
                    if (showInAxis && !pivotField.ShowingInAxis)
                        pivotField.Hidden = false;

                    System.Array arrOldVisibleItems = (System.Array)pivotField.VisibleItemsList;
                    List<object> listNewVisibleItems = new List<object>();
                    bool bFound = false;
                    foreach (object o in arrOldVisibleItems)
                    {
                        listNewVisibleItems.Add(o);
                        if (Convert.ToString(o) == m.UniqueName)
                        {
                            bFound = true;
                        }
                    }
                    if (!bFound)
                    {
                        if (!(listNewVisibleItems.Count == 1 && string.IsNullOrEmpty(Convert.ToString(listNewVisibleItems[0]))))
                        {
                            //this level is filtered, so add this member to this level's filters
                            listNewVisibleItems.Add(m.UniqueName);
                            System.Array arrNewVisibleItems = listNewVisibleItems.ToArray();
                            pivotField.VisibleItemsList = arrNewVisibleItems;
                        }
                    }
                }
                catch { } //not sure why it failed... oh well
            }

            //now expand all parents to get to this
            //don't expand the found member itself
            if (!showInAxis && pivotField.ShowingInAxis)
            {
                foreach (Excel.PivotItem pivotItem in (Excel.PivotItems)pivotField.PivotItems(System.Type.Missing))
                {
                    if (pivotItem.Value == m.UniqueName)
                    {
                        try //can't always drilldown
                        {
                            pivotItem.DrilledDown = true;
                        }
                        catch { }
                    }
                }
                //TODO: future: if you can't find the member, then see if the filter or grouping should be cleared
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            try
            {
                btnCancelSearch_Click(null, null);

                try
                {
                    if (connCube != null && connCube.State != ConnectionState.Closed)
                        connCube.Close();
                }
                catch { }

                base.OnClosed(e);
            }
            catch { }
        }

        private void btnCancelSearch_Click(object sender, EventArgs e)
        {
            try
            {
                if (searcher != null)
                {
                    searcher.Cancel();

                    lblNoSearchMatches.Visible = (searcher.Matches.Count == 0);
                    btnSearchAdd.Enabled = (searcher.Matches.Count > 0);
                    btnFind.Enabled = true;

                    prgSearch.Visible = false;
                    btnCancelSearch.Visible = false;

                    if (!string.IsNullOrEmpty(searcher.Error))
                    {
                        lblSearchError.Text = searcher.Error;
                        lblSearchError.Visible = true;
                    }
                }
            }
            catch { }
        }

        private bool IsExcel2007OrHigherPivotTableVersion()
        {
            return ((int)pvt.Version >= (int)Excel.XlPivotTableVersionList.xlPivotTableVersion12);
        }

        private void cmbLookIn_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                chkMemberProperties.Checked = (cmbLookIn.SelectedIndex != 1);
            }
            catch { }
        }

        private void MainForm_Activated(object sender, EventArgs e)
        {
            try
            {
                if (tabControl.SelectedTab == tabSearch && !txtSearch.Focused && !cmbLookIn.Focused)
                {
                    txtSearch.Focus();
                    txtSearch.SelectAll();
                }
                else if (tabControl.SelectedTab == tabCalcs && !comboCalcName.Focused && !txtCalcFormula.Focused)
                {
                    comboCalcName.Focus();
                }
            }
            catch { }
        }

        //change the cursor for the header cells which can be sorted
        private void dataGridSearchResults_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (searcher != null && searcher.Complete && e.RowIndex == -1 && e.ColumnIndex > 0)
                {
                    dataGridSearchResults.Cursor = Cursors.Hand;
                }
            }
            catch { }
        }

        private void dataGridSearchResults_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridSearchResults.Cursor = Cursors.Default;
            }
            catch { }
        }

        //TODO: future: let them filter fields not in the PivotTable
        private void btnFilterList_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbFilterListLookIn.SelectedIndex < 0)
                {
                    MessageBox.Show("Choose a field to filter first.");
                    return;
                }
                if (txtFilterList.Text.Length == 0)
                {
                    MessageBox.Show("Paste in a list of items to set the filter to first.");
                    return;
                }

                progressFilterList.Visible = true;
                progressFilterList.Value = 0;

                btnCancelFilterList.Visible = true;
                btnFilterList.Enabled = false;
                txtFilterList.ReadOnly = true;

                FilterListWorkerArgs args = new FilterListWorkerArgs();
                args.Lines = txtFilterList.Lines;
                args.LookIn = Convert.ToString(cmbFilterListLookIn.SelectedItem);

                workerFilterList = new BackgroundWorker();
                workerFilterList.DoWork += new DoWorkEventHandler(workerFilterList_DoWork);
                workerFilterList.WorkerSupportsCancellation = true;
                workerFilterList.RunWorkerAsync(args);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        private class FilterListWorkerArgs
        {
            public string[] Lines;
            public string LookIn;
        }

        void workerFilterList_DoWork(object sender, DoWorkEventArgs e)
        {
            List<string> listMembersNotFound = new List<string>();

            try
            {
                AddInWorking = true;
                ConnectAdomdClientCube();

                if (e.Cancel) return;

                FilterListWorkerArgs args = (FilterListWorkerArgs)e.Argument;

                ////////////////////////////////////////////////////////////////////
                // SEARCH FOR MEMBERS
                ////////////////////////////////////////////////////////////////////
                AdomdCommand cmd = new AdomdCommand();
                cmd.Connection = cube.ParentConnection;

                StringBuilder sFoundMemberUniqueNames = new StringBuilder();

                Dictionary<string, List<object>> dictLevelsOfFoundMembers = new Dictionary<string, List<object>>();

                int iNumLinesFinished = 0;
                foreach (string sLine in args.Lines)
                {
                    if (e.Cancel) return;
                    if (!string.IsNullOrEmpty(sLine.Trim()))
                    {

                        AdomdRestrictionCollection restrictions = new AdomdRestrictionCollection();
                        restrictions.Add(new AdomdRestriction("CATALOG_NAME", cube.ParentConnection.Database));
                        restrictions.Add(new AdomdRestriction("CUBE_NAME", cube.Name));
                        restrictions.Add(new AdomdRestriction("HIERARCHY_UNIQUE_NAME", args.LookIn));
                        restrictions.Add(new AdomdRestriction("MEMBER_NAME", sLine.Trim()));
                        System.Data.DataTable tblExactMatchMembers = cube.ParentConnection.GetSchemaDataSet("MDSCHEMA_MEMBERS", restrictions).Tables[0];

                        if (tblExactMatchMembers.Rows.Count > 0)
                        {
                            foreach (System.Data.DataRow row in tblExactMatchMembers.Rows)
                            {
                                if (!dictLevelsOfFoundMembers.ContainsKey(Convert.ToString(row["LEVEL_UNIQUE_NAME"])))
                                    dictLevelsOfFoundMembers.Add(Convert.ToString(row["LEVEL_UNIQUE_NAME"]), new List<object>());
                                dictLevelsOfFoundMembers[Convert.ToString(row["LEVEL_UNIQUE_NAME"])].Add(Convert.ToString(row["MEMBER_UNIQUE_NAME"]));
                            }
                        }
                        else
                        {
                            listMembersNotFound.Add(sLine.Trim());
                        }
                    }

                    SetFilterListProgress((int)(90 * (++iNumLinesFinished) / args.Lines.Length), true, null);
                }

                Excel.CubeField field = pvt.CubeFields.get_Item(args.LookIn);
                field.CreatePivotFields();

                foreach (string sLevelUniqueName in dictLevelsOfFoundMembers.Keys)
                {
                    Excel.PivotField pivotField = (Excel.PivotField)field.PivotFields.Item(sLevelUniqueName);
                    if (field.Orientation == Excel.XlPivotFieldOrientation.xlPageField)
                    {
                        field.EnableMultiplePageItems = true;
                    }
                    System.Array arrNewVisibleItems = dictLevelsOfFoundMembers[sLevelUniqueName].ToArray();
                    pivotField.VisibleItemsList = arrNewVisibleItems;
                    if (field.Orientation == Excel.XlPivotFieldOrientation.xlHidden)
                    {
                        field.Orientation = Excel.XlPivotFieldOrientation.xlRowField; //if it's not in the PivotTable, then add it to rows
                    }
                    pivotField.ClearValueFilters();
                    pivotField.ClearLabelFilters();
                }

                SetFilterListProgress(100, false, listMembersNotFound.ToArray());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);

                SetFilterListProgress(0, false, listMembersNotFound.ToArray());
            }
            finally
            {
                AddInWorking = false;
            }
        }

        private void btnCancelFilterList_Click(object sender, EventArgs e)
        {
            try
            {
                if (workerFilterList != null)
                {
                    workerFilterList.CancelAsync();
                }
            }
            catch { }
        }

        private delegate void SetFilterListProgress_Delegate(int iProgress, bool bVisible, string[] arrMembersNotFound);
        private void SetFilterListProgress(int iProgress, bool bVisible, string[] arrMembersNotFound)
        {
            if (progressFilterList.InvokeRequired)
            {
                //avoid the "cross-thread operation not valid" error message
                progressFilterList.BeginInvoke(new SetFilterListProgress_Delegate(SetFilterListProgress), new object[] { iProgress, bVisible, arrMembersNotFound });
            }
            else
            {
                progressFilterList.Value = iProgress;
                progressFilterList.Visible = bVisible;
                btnCancelFilterList.Visible = bVisible;

                if (iProgress == 100)
                {
                    btnCancelFilterList.Visible = false;
                    btnFilterList.Enabled = true;

                    if (arrMembersNotFound.Length == 0)
                    {
                        this.Close();
                    }
                    else
                    {
                        txtFilterList.ReadOnly = false;
                        string sError = "The following members were not found.\r\n";
                        if (arrMembersNotFound.Length > 10) sError += " (Showing first 10)\r\n";
                        sError += "\r\n" + string.Join("\r\n", arrMembersNotFound);
                        MessageBox.Show(sError);
                    }
                }
            }
        }
    }
}
