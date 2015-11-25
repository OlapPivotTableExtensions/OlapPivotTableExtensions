using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using OlapPivotTableExtensions.AdomdClientWrappers;
using LevelTypeEnum = Microsoft.AnalysisServices.AdomdClient.LevelTypeEnum;


namespace OlapPivotTableExtensions
{
    public partial class MainForm : Form
    {
        private Excel.PivotTable pvt;
        private Excel.Application application;
        private CalculationsLibrary library;
        internal AdomdConnection connCube;
        private CubeDef cube;
        internal string cubeName;
        private CubeSearcher searcher;
        public bool AddInWorking = false;
        private BackgroundWorker workerFilterList;
        private int xlPivotTableVersion14 = 4; //since we're using the Excel 2007 object model, we can't see the Excel 2010 version
        private int xlPivotTableVersion15 = 5; //since we're using the Excel 2007 object model, we can't see the Excel 2013 version
        private int xlPivotTableVersion16 = 6; //since we're using the Excel 2007 object model, we can't see the Excel 2013 version
        internal static int xlConnectionTypeMODEL = 7; //since we're using the Excel 2007 object model, we can't see the Excel 2013 connection types
        private int xlCalculatedMeasure = 2; //since we're using the Excel 2007 object model, we can't see the new Excel 2013 calc measure type

        private int _LibraryComboDividerItemIndex = int.MaxValue;

        private BackgroundWorker workerFormatMDX;

        private DataGridViewCheckBoxColumnHeaderCell colHeaderSearchCheckAll;
        private bool _SearchLookInShowAllHierarchies = false;
        private bool SearchWasCancelled = false;

        private bool bImpersonate = false;
        private string sUsername = null;
        private string sDomain = null;
        private string sPassword = null;

        public MainForm(Excel.Application app)
        {
            InitializeComponent();

            try
            {
                application = app;

                string sLanguage = "Windows Language: " + Connect.OriginalLanguage;
                sLanguage += "\r\nWindows UI Language: " + System.Globalization.CultureInfo.InstalledUICulture.EnglishName;

                bool bCachedSupportedLanguageConfig = IsSupportedLanguageConfiguration;

                //set the culture to the Excel UI language, then leave it until this form is closed
                SetCulture(app);

                System.Globalization.CultureInfo nciExcelUI = new System.Globalization.CultureInfo(app.LanguageSettings.get_LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI));
                sLanguage += "\r\nExcel UI Language: " + nciExcelUI.EnglishName;

                try
                {
                    System.Globalization.CultureInfo nciInstall = new System.Globalization.CultureInfo(app.LanguageSettings.get_LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDInstall));
                    sLanguage += "\r\nExcel Install Language: " + nciInstall.EnglishName;
                }
                catch { }

                lblExcelUILanguage.Text = sLanguage;

                System.Reflection.AssemblyFileVersionAttribute attrVersion = (System.Reflection.AssemblyFileVersionAttribute)typeof(MainForm).Assembly.GetCustomAttributes(typeof(System.Reflection.AssemblyFileVersionAttribute), true)[0];
                lblVersion.Text = "OLAP PivotTable Extensions v" + attrVersion.Version;

#if VSTO40
                lblVersion.Text += (Environment.Is64BitProcess ? " (64-bit)" : " (32-bit)");
#elif X64
                lblVersion.Text += " (64-bit)";
#else
                lblVersion.Text += " (32-bit)";
#endif

                pvt = app.ActiveCell.PivotTable;

                library = new CalculationsLibrary();
                library.Load();

                FillCalcsDropdown();

                if (!bCachedSupportedLanguageConfig)
                {
                    MessageBox.Show("You are not running a supported language configuration!\r\n\r\nClick on the \"UNSUPPORTED LANGUAGE CONFIGURATION!\" link on the About tab for details on how to resolve this problem.", "OLAP PivotTable Extensions");
                    tabControl.SelectedTab = tabAbout;
                    linkUnsupportedLanguageConfiguration.Visible = true;
                    tooltip.SetToolTip(linkUnsupportedLanguageConfiguration, "If you want to use OLAP PivotTable Extensions without errors, you must do one of the following:\r\n* Install the Office Language Pack for " + Connect.OriginalLanguage + "\r\n* Change the Windows Regional settings to a language for which you have an Office Language Pack installed\r\n* Check \"Retrieve data and errors in the Office display language when available\" on this PivotTable connection\r\n* Include LocaleIdentifier on the connection string\r\n\r\nClick for more instructions");
                }
                else
                {
                    tabControl.SelectedTab = tabCalcs;
                    linkUnsupportedLanguageConfiguration.Visible = false;
                }

                chkShowCalcMembers.Checked = Connect.ShowCalcMembersByDefault;
                chkRefreshDataWhenOpeningTheFile.Checked = Connect.RefreshDataByDefault;

                chkFormatMDX.Enabled = false; //signals to checked event not to format the MDX right now
                chkFormatMDX.Checked = Connect.FormatMdx;
                chkFormatMDX.Enabled = true;


                lblPivotTableVersion.Text = "Version of This PivotTable: " + GetPivotTableVersion();

                if (string.Compare(GetPivotTableVersion(), GetExcelVersion()) >= 0)
                {
                    linkUpgradePivotTable.Visible = false;
                    lblUpgradePivotTableInstructions.Visible = false;
                    btnUpgradeOnRefresh.Visible = false;
                }
                else
                {
                    if (this.application.ActiveWorkbook.FileFormat == Excel.XlFileFormat.xlOpenXMLWorkbook //if it's xlsx
                        || this.application.ActiveWorkbook.FileFormat == Excel.XlFileFormat.xlExcel12) //if it's xlsb
                    {
                        if (pvt.PivotCache().UpgradeOnRefresh)
                        {
                            lblUpgradePivotTableInstructions.Text = "To upgrade, refresh the PivotTable.";
                            btnUpgradeOnRefresh.Visible = false;
                        }
                        else
                        {
                            lblUpgradePivotTableInstructions.Text = "To upgrade, click the UpgradeOnRefresh button, then refresh the PivotTable.";
                            btnUpgradeOnRefresh.Visible = true;
                        }
                    }
                    else
                    {
                        lblUpgradePivotTableInstructions.Text = "To upgrade, save as .xlsx then refresh the PivotTable.";
                        btnUpgradeOnRefresh.Visible = false;
                    }
                }

                if (!Connect.IsOledbConnection(application.ActiveCell.PivotTable))
                {
                    //MDX calcs don't appear to be supported on ExcelDataModel pivots
                    tabControl.Controls.Remove(tabCalcs);
                    tabControl.Controls.Remove(tabLibrary);
                    tabControl_SelectedIndexChanged(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
                this.Visible = false;
                this.Close();
            }
        }

        //be sure to set the culture back when the form is closed
        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            ResetCulture(application);
        }

        private static bool _ShouldRunSetCulture = true;
        private bool? _IsSupportedLanguageConfiguration = null;
        private bool _RetrieveInOfficeUILang = true;
        private bool IsSupportedLanguageConfiguration
        {
            get
            {
                if (_IsSupportedLanguageConfiguration != null) return (bool)_IsSupportedLanguageConfiguration;

                bool bReceivedOldFormatError = false;
                Excel.PivotTable pvtLocal = null;
                try
                {
                    //try this without setting the culture... if their Windows regional settings language isn't a language for which they have an Office language pack, this should blow up with the "old format" error
                    pvtLocal = application.ActiveCell.PivotTable;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    bReceivedOldFormatError = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("There was an unexpected error checking your language configuration:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
                    _IsSupportedLanguageConfiguration = true;
                    return true;
                }

                bool bConnectionStringContainsLCID = false;
                try
                {
                    //now get the pvt object using the Excel UI culture
                    SetCulture(application);
                    pvtLocal = application.ActiveCell.PivotTable;

                    Microsoft.Office.Interop.Excel.PivotCache cache = pvtLocal.PivotCache();
                    Excel.WorkbookConnection workbookConn = cache.WorkbookConnection;
                    if (Connect.IsOledbConnection(pvtLocal))
                    {
                        Excel.OLEDBConnection connOLEDB = workbookConn.OLEDBConnection;
                        _RetrieveInOfficeUILang = connOLEDB.RetrieveInOfficeUILang;

                        string sConnectionString = Convert.ToString(connOLEDB.Connection); //not the same as the connection string we use for AdomdClient since it won't contain the password, but it's good enough for this and doesn't require connecting

                        if (sConnectionString.ToLower().Contains("language identifier=")
                            || sConnectionString.ToLower().Contains("localeidentifier=")
                            || sConnectionString.ToLower().Contains("locale identifier=") //note, Locale Identifier doesn't often show up unless it's inside Extended Properties. So OLAP PivotTable Extensions can't use it... but it does work for Excel
                        )
                        {
                            bConnectionStringContainsLCID = true;
                        }
                    }
                }
                catch (Exception exInner)
                {
                    MessageBox.Show("ERROR FIGURING OUT WHETHER IT'S AN INVALID CONFIGURATION! " + exInner.Message + " - " + exInner.GetType().FullName + "\r\n" + exInner.StackTrace);
                    _IsSupportedLanguageConfiguration = false;
                    return false;
                }
                finally
                {
                    ResetCulture(application);
                }

                if (bReceivedOldFormatError) //don't have language pack installed
                {
                    if (!_RetrieveInOfficeUILang)
                    {
                        _ShouldRunSetCulture = true;
                        if (!bConnectionStringContainsLCID)
                        {
                            _IsSupportedLanguageConfiguration = false;
                            return false;
                        }
                        else
                        {
                            _IsSupportedLanguageConfiguration = true;
                            return true;
                        }
                    }
                    else
                    {
                        _ShouldRunSetCulture = true;
                        _IsSupportedLanguageConfiguration = true;
                        return true;
                    }
                }
                else //have language pack installed
                {
                    if (!_RetrieveInOfficeUILang)
                    {
                        _ShouldRunSetCulture = false;
                        _IsSupportedLanguageConfiguration = true;
                        return true;
                    }
                    else
                    {
                        _ShouldRunSetCulture = true;
                        _IsSupportedLanguageConfiguration = true;
                        return true;
                    }
                }


            }
        }

        private static Dictionary<int, int> _dictSetCultureDepth = new Dictionary<int, int>();

        //fix for the "old format or invalid type library" error on non-english locales
        public static void SetCulture(Excel.Application app)
        {
            if (!_ShouldRunSetCulture) return;

            System.Globalization.CultureInfo nci =
            new System.Globalization.CultureInfo(
            app.LanguageSettings.get_LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI));

            System.Threading.Thread.CurrentThread.CurrentCulture = nci;

            //cache the set culture depth
            if (_dictSetCultureDepth.ContainsKey(System.Threading.Thread.CurrentThread.ManagedThreadId))
                _dictSetCultureDepth[System.Threading.Thread.CurrentThread.ManagedThreadId]++;
            else
                _dictSetCultureDepth.Add(System.Threading.Thread.CurrentThread.ManagedThreadId, 1);
        }

        //fix for the LocaleIdentifier error on drillthrough
        public static void ResetCulture(Excel.Application app)
        {
            if (!_ShouldRunSetCulture) return;

            if (!_dictSetCultureDepth.ContainsKey(System.Threading.Thread.CurrentThread.ManagedThreadId)) return;

            //if two SetCulture calls are made before the first ResetCulture call is made, we should skip it until we get to the final ResetCulture call, otherwise it's reset to prematurely
            if (_dictSetCultureDepth[System.Threading.Thread.CurrentThread.ManagedThreadId] <= 1)
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = Connect.OriginalCultureInfo;
            }

            if (_dictSetCultureDepth[System.Threading.Thread.CurrentThread.ManagedThreadId] > 0) //ResetCulture should never be called when this = 0, but this is just to make sure
            {
                _dictSetCultureDepth[System.Threading.Thread.CurrentThread.ManagedThreadId]--;
            }
        }

        private void SetMDX()
        {
            //if this isn't a supported language configuration, still try to help them be able to see the MDX by resetting the culture, grabbing the MDX, then setting it again
            if (!IsSupportedLanguageConfiguration) ResetCulture(application);

            try
            {
                StringBuilder sMdxQuery = new StringBuilder(pvt.MDX);

                AddCalculatedMembersToMdxQuery(sMdxQuery);

                richTextBoxMDX.Text = sMdxQuery.ToString();
                richTextBoxMDX.SelectionStart = 0;
                richTextBoxMDX.SelectionLength = sMdxQuery.Length;
                richTextBoxMDX.Focus();
                richTextBoxMDX.ScrollToCaret();

                if (Connect.FormatMdx)
                {
                    InitiateFormatMDX(sMdxQuery.ToString());
                }

                tooltip.SetToolTip(chkFormatMDX, "Checking this box will send your MDX query over the internet to this web service:\r\nhttp://formatmdx.azurewebsites.net/formatter.asmx");
            }
            finally
            {
                //if this isn't a supported language configuration, still try to help them be able to see the MDX by using a reset culture above, but a set culture here
                if (!IsSupportedLanguageConfiguration) SetCulture(application);
            }
        }

        internal void AddCalculatedMembersToMdxQuery(StringBuilder sMdxQuery)
        {
            //add (session) calculated members to the query so that you can run it from SSMS
            if (pvt.CalculatedMembers.Count > 0)
            {
                StringBuilder sCalcs = new StringBuilder();
                foreach (Excel.CalculatedMember calc in pvt.CalculatedMembers)
                {
                    if (calc.Type == Excel.XlCalculatedMemberType.xlCalculatedSet)
                        sCalcs.Append("SET ");
                    else
                        sCalcs.Append("MEMBER ");
                    sCalcs.AppendFormat("{0} as {1}\r\n", calc.Name, calc.Formula.Replace("\r\n", "\r").Replace("\r", "\r\n")); //normalize the line breaks which have been turned into \r to workaround an Excel Services bug
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
                MessageBox.Show(AddOledbErrorToException(ex, false), "OLAP PivotTable Extensions");
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
                    int iCalcType = (int)Excel.XlCalculatedMemberType.xlCalculatedMember;
                    if (bMeasure && string.Compare(GetExcelVersion(), "2013") >= 0)
                    {
                        iCalcType = xlCalculatedMeasure;
                    }
                    oCalcMember = pvt.CalculatedMembers.Add(sName, sFormula.Replace("\r\n", "\r"), System.Reflection.Missing.Value, iCalcType);
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
                    MessageBox.Show("There was a problem creating the calculation:\r\n" + AddOledbErrorToException(ex, false), "OLAP PivotTable Extensions");
                }

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an unexpected error creating the calculation:\r\n" + AddOledbErrorToException(ex, false), "OLAP PivotTable Extensions");
            }
        }

        private void FillCalcsDropdown()
        {
            comboCalcName.Items.Clear();
            List<string> listCalcs = new List<string>();
            foreach (Excel.CalculatedMember calc in pvt.CalculatedMembers)
            {
                if (calc.Type == Excel.XlCalculatedMemberType.xlCalculatedMember || (int)calc.Type == xlCalculatedMeasure)
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

        private string AddOledbErrorToException(Exception ex, bool includeStackTrace)
        {
            string sErrors = string.Empty;
            try
            {
                if (application.OLEDBErrors != null)
                {
                    foreach (Excel.OLEDBError err in application.OLEDBErrors)
                    {
                        if (sErrors.Length > 0) sErrors += "\r\n";
                        sErrors += err.ErrorString;
                    }
                }
            }
            catch { }
            if (sErrors.Length > 0) sErrors += "\r\n";
            sErrors += ex.Message;
            if (includeStackTrace) sErrors += "\r\n" + ex.StackTrace;
            return sErrors;
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
                    MessageBox.Show("There was a problem capturing the MDX query for this PivotTable.\r\n" + AddOledbErrorToException(ex, false), "OLAP PivotTable Extensions");
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
                    MessageBox.Show("There was a problem setting up the search tab.\r\n" + AddOledbErrorToException(ex, true), "OLAP PivotTable Extensions");
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
                    MessageBox.Show("There was a problem setting up the Filter List tab.\r\n" + AddOledbErrorToException(ex, false), "OLAP PivotTable Extensions");
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
                    btnFilterListShowCurrentFilters.Enabled = false;
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
                if (Connect.SearchMeasuresOnlyDefault)
                {
                    cmbLookIn.SelectedIndex = 1;
                }
                else
                {
                    cmbLookIn.SelectedIndex = 0;
                }
                cmbLookIn_SelectedIndexChanged(null, null);
            }

            if (_IsSingleSearch)
            {
                txtSearch.Focus();
                txtSearch.SelectAll();
            }
            else
            {
                richTextBoxSearch.Focus();
                richTextBoxSearch.SelectAll();
            }

            //setup the check all header
            if (colHeaderSearchCheckAll == null)
            {
                colHeaderSearchCheckAll = new DataGridViewCheckBoxColumnHeaderCell();
                colHeaderSearchCheckAll.Value = string.Empty;
                colCheck.HeaderCell = colHeaderSearchCheckAll;
                colCheck.HeaderCell.ToolTipText = "Check the box to add this search result to your PivotTable.";
            }

            Application.DoEvents();

            ConnectAdomdClientCube();

            CubeSearcher.SetupSearchOptimizationsAsync(cube, bImpersonate, sUsername, sDomain, sPassword);
        }

        internal void ConnectAdomdClientCube()
        {
            Microsoft.Office.Interop.Excel.PivotCache cache = pvt.PivotCache();
            if (!cache.IsConnected)
                cache.MakeConnection();

            ADODB.Connection connADO = cache.ADOConnection as ADODB.Connection;
            if (connADO == null) throw new Exception("Could not cast PivotCache.ADOConnection to ADODB.Connection.");

            string sConnectionString = connADO.ConnectionString;

            bool bIsExcel15Model = false;
            if (cache.WorkbookConnection.Type == Excel.XlConnectionType.xlConnectionTypeOLEDB)
            {
                Excel.OLEDBConnection connOLEDB = cache.WorkbookConnection.OLEDBConnection;

                //figure out current locale
                if (connOLEDB.RetrieveInOfficeUILang
                    && !sConnectionString.ToLower().Contains("language identifier=")
                    && !sConnectionString.ToLower().Contains("localeidentifier=")
                    && !sConnectionString.ToLower().Contains("locale identifier=") //note, Locale Identifier doesn't often show up. So OLAP PivotTable Extensions can't use it... but it does work for Excel
                )
                {
                    sConnectionString += ";LocaleIdentifier=" + this.application.LanguageSettings.get_LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI);
                }
            }
            else if ((int)cache.WorkbookConnection.Type == xlConnectionTypeMODEL)
            {
                bIsExcel15Model = true;
            }

            sConnectionString += ";Application Name=" + lblVersion.Text;


            //look for impersonation info so we can mimic what Excel does
            bImpersonate = false;
            sUsername = null;
            sDomain = null;
            sPassword = null;
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

            //support PowerPivot in-process cube
            if (connParser["Data Source"] != null && connParser["Data Source"].ToLower() == "$embedded$" && connParser["Location"] == null)
            {
                //DISCLAIMER: The ability to connect to PowerPivot from OLAP PivotTable Extensions is using unsupported APIs and as such the behaviour may change or stop working without notice in future releases. This functionality is provided on an "as-is" basis.
                sConnectionString += ";Location=" + this.application.ActiveWorkbook.FullName;
            }

            //remove the Data Source Version connection string parameter as it will cause an error in AdomdClient... work item 23022
            if (connParser.ContainsKey("Data Source Version"))
            {
                sConnectionString = sConnectionString.Replace("Data Source Version=" + connParser["Data Source Version"], string.Empty);
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
                            connCube = new AdomdConnection(sConnectionString, AdomdType.AnalysisServices);
                            connCube.Open();
                        }
                        catch (ArgumentException ex)
                        {
                            //may be that you can't use Integrated Security=SSPI with an HTTP or HTTPS connection... try to workaround that
                            if (sConnectionString.ToLower().IndexOf("data source=http") >= 0 && sConnectionString.ToLower().IndexOf("integrated security=sspi;") >= 0)
                            {
                                sConnectionString = sConnectionString.Remove(sConnectionString.ToLower().IndexOf("integrated security=sspi;"), "integrated security=sspi;".Length);
                                connCube = new AdomdConnection(sConnectionString, AdomdType.AnalysisServices);
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
                        if (!bIsExcel15Model)
                        {
                            connCube = new AdomdConnection(sConnectionString, AdomdType.AnalysisServices);
                        }
                        else
                        {
                            connCube = new AdomdConnection(sConnectionString, AdomdType.Excel);
                        }
                        connCube.Open();
                    }
                    catch (ArgumentException ex)
                    {
                        //may be that you can't use Integrated Security=SSPI with an HTTP or HTTPS connection... try to workaround that
                        if (!bIsExcel15Model && sConnectionString.ToLower().IndexOf("data source=http") >= 0 && sConnectionString.ToLower().IndexOf("integrated security=sspi;") >= 0)
                        {
                            sConnectionString = sConnectionString.Remove(sConnectionString.ToLower().IndexOf("integrated security=sspi;"), "integrated security=sspi;".Length);
                            connCube = new AdomdConnection(sConnectionString, AdomdType.AnalysisServices);
                            connCube.Open();
                        }
                        else
                        {
                            MessageBox.Show(AddOledbErrorToException(ex, false) + "\r\n" + sConnectionString);
                            throw;
                        }
                    }
                    catch (Exception ex)
                    {
                        if (connCube != null && connCube.UnderlyingConnection != null)
                        {
                            MessageBox.Show(AddOledbErrorToException(ex, false) + "\r\n" + sConnectionString + "\r\n" + connCube.ClientVersion + "\r\n" + connCube.UnderlyingConnection.GetType().Assembly.Location + "\r\n" + ex.StackTrace);
                        }
                        else
                        {
                            MessageBox.Show(AddOledbErrorToException(ex, false) + "\r\n" + sConnectionString + "\r\n" + ex.StackTrace);
                        }
                        throw;
                    }
                }

                cubeName = Convert.ToString(cache.CommandText);
                cube = connCube.Cubes.Find(cubeName);
                if (cube == null)
                {
                    throw new Exception("Could not find cube [" + Convert.ToString(cache.CommandText) + "]");
                }

                FillSearchLookInDropdown();
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

        private const string LOOK_IN_SHOW_ALL_HIERARCHIES = "+ Show All Hierarchies";
        private void FillSearchLookInDropdown()
        {
            //fill in the "Look in" dropdown with the dimension hierarchies in the PivotTable
            cmbLookIn.SuspendLayout();
            while (cmbLookIn.Items.Count > 3)
                cmbLookIn.Items.RemoveAt(3);

            if (!_SearchLookInShowAllHierarchies)
            {
                List<Excel.CubeField> listHierarchies = new List<Excel.CubeField>();
                foreach (Excel.CubeField f in pvt.CubeFields)
                {
                    if (f.CubeFieldType == Excel.XlCubeFieldType.xlHierarchy) //not named sets since you can't filter them, and not measures since they are returned in the field list search
                    {
                        if (f.Orientation != Excel.XlPivotFieldOrientation.xlHidden)
                        {
                            listHierarchies.Add(f);
                        }
                    }
                }

                listHierarchies.Sort(delegate(Excel.CubeField x, Excel.CubeField y) { return x.Name.CompareTo(y.Name); });

                foreach (Excel.CubeField f in listHierarchies)
                {
                    cmbLookIn.Items.Add(f.Name);
                    Excel.PivotFields pfs = null;
                    try
                    {
                        pfs = f.PivotFields;
                    }
                    catch
                    {
                        f.CreatePivotFields(); //this connects to SSAS, so avoid it if not necessary
                        pfs = f.PivotFields;
                    }
                    List<string> listLevels = new List<string>();
                    foreach (Excel.PivotField pf in pfs)
                    {
                        if (!pf.IsMemberProperty)
                        {
                            listLevels.Add("  " + pf.Name);
                        }
                    }
                    if (listLevels.Count > 1)
                    {
                        cmbLookIn.Items.AddRange(listLevels.ToArray());
                    }

                }

                cmbLookIn.Items.Add(LOOK_IN_SHOW_ALL_HIERARCHIES);
            }
            else
            {
                //we're showing all fields, so we have to use ADOMD.NET to get the list of hierarchies because looping through CubeFields will show fields that are hidden: http://support.microsoft.com/kb/931386
                List<Hierarchy> listHierarchies = new List<Hierarchy>();
                foreach (Dimension d in cube.Dimensions)
                {
                    foreach (Hierarchy h in d.Hierarchies)
                    {
                        if (h.UniqueName == "[Measures]") continue;
                        listHierarchies.Add(h);
                    }
                }
                listHierarchies.Sort(delegate(Hierarchy x, Hierarchy y) { return x.UniqueName.CompareTo(y.UniqueName); });
                foreach (Hierarchy f in listHierarchies)
                {
                    cmbLookIn.Items.Add(f.UniqueName);
                    List<string> listLevels = new List<string>();
                    foreach (Level l in f.Levels)
                    {
                        if (l.LevelType == LevelTypeEnum.All) continue;
                        listLevels.Add("  " + l.UniqueName);
                    }
                    if (listLevels.Count > 1)
                    {
                        cmbLookIn.Items.AddRange(listLevels.ToArray());
                    }

                }
            }

            cmbLookIn.ResumeLayout();
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

        private void linkUpgradePivotTable_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://office.microsoft.com/en-us/excel-help/working-with-different-pivottable-formats-in-office-excel-HA010167298.aspx");
        }

        private void linkUnsupportedLanguageConfiguration_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://olappivottableextend.codeplex.com/wikipage?title=Unsupported%20Language%20Configuration");
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
                Connect.ShowCalcMembersByDefault = chkShowCalcMembers.Checked;
                Connect.RefreshDataByDefault = this.chkRefreshDataWhenOpeningTheFile.Checked;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "OLAP PivotTable Extensions");
            }
        }

        private string SearchText
        {
            get
            {
                if (_IsSingleSearch)
                    return txtSearch.Text.Trim();
                else
                    return richTextBoxSearch.Text.Trim();
            }
        }

        private void RemoveBlankLinesAndTrimLinesAndRemoveDuplicates(List<string> lines)
        {
            List<string> listNew = new List<string>();
            
            for (int i = 0; i < lines.Count; i++)
            {
                string sLineTrimmed = lines[i].Trim();
                bool bAlreadyInList = (listNew.FindIndex(x => x.Equals(sLineTrimmed, StringComparison.CurrentCultureIgnoreCase)) != -1); //case insensitive Contains

                if (lines[i].Trim().Length == 0 || bAlreadyInList)
                {
                    lines.RemoveAt(i);
                    i--;
                }
                else
                {
                    lines[i] = sLineTrimmed;
                    listNew.Add(sLineTrimmed);
                }
            }
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            try
            {
                if (SearchText == string.Empty && !chkExactMatch.Checked && cmbLookIn.SelectedIndex == 2)
                {
                    MessageBox.Show("Please enter a search term", "OLAP PivotTable Extensions");
                    return;
                }

                dataGridSearchResults.Enabled = true;
                colName.SortMode = DataGridViewColumnSortMode.NotSortable;
                colType.SortMode = DataGridViewColumnSortMode.NotSortable;
                colFolder.SortMode = DataGridViewColumnSortMode.NotSortable;
                colDesc.SortMode = DataGridViewColumnSortMode.NotSortable;

                colHeaderSearchCheckAll.CheckAll = false;

                lblNoSearchMatches.Visible = false;
                lblSearchError.Visible = false;
                lblSearchTermsHighlighted.Visible = false;

                //if the find button which we're about to disable has focus, then switch focus to the search box
                if (btnFind.Focused)
                {
                    if (_IsSingleSearch)
                        txtSearch.Focus();
                    else
                        richTextBoxSearch.Focus();
                }

                btnFind.Enabled = false;
                btnCancelSearch.Visible = true;
                btnSearchAdd.Enabled = false;
                btnMultiSearch.Enabled = false;
                SearchWasCancelled = false;

                //remove formatting by assigning to .Text
                //remove blank lines as they will cause extra matches in the search
                List<string> listSearchTermLines = new List<string>(richTextBoxSearch.Text.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries));
                RemoveBlankLinesAndTrimLinesAndRemoveDuplicates(listSearchTermLines);
                richTextBoxSearch.Text = string.Join("\r\n", listSearchTermLines.ToArray());

                Application.DoEvents();

                prgSearch.Value = 0;
                prgSearch.Visible = true;

                lblSearchMatchesCount.Visible = false;

                CubeSearcher.CubeSearchScope scope = CubeSearcher.CubeSearchScope.FieldList;
                if (cmbLookIn.SelectedIndex == 1)
                    scope = CubeSearcher.CubeSearchScope.MeasuresCaptionOnly;
                else if (cmbLookIn.SelectedIndex >= 2)
                    scope = CubeSearcher.CubeSearchScope.DimensionData;

                string sSearchOnlyDimension = null;
                if (cmbLookIn.SelectedIndex > 2)
                    sSearchOnlyDimension = Convert.ToString(cmbLookIn.SelectedItem);

                //TODO: bold any search results which are already in the PivotTable... may need to pass in a delegate to update the CubeSearchMatch object to have a reference to the parent class

                searcher = new CubeSearcher(cube, scope, SearchText, chkExactMatch.Checked, chkMemberProperties.Checked, sSearchOnlyDimension, dataGridSearchResults, bImpersonate, sUsername, sDomain, sPassword);
                searcher.ProgressChanged += new ProgressChangedEventHandler(searcher_ProgressChanged);
                searcher.SearchAsync();

                cubeSearchMatchBindingSource.DataSource = searcher.Matches;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during search: " + AddOledbErrorToException(ex, false), "OLAP PivotTable Extensions");
            }
        }

        private void dataGridSearchResults_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 0 && searcher != null)
                {
                    if (e.RowIndex >= 0 && !searcher.Complete)
                    {
                        //the control doesn't do too well while we're constantly setting a new DataSource... this code worksaround that problem so you can check/uncheck while the search is still going on
                        searcher.Matches[e.RowIndex].Checked = !searcher.Matches[e.RowIndex].Checked;
                        dataGridSearchResults.InvalidateCell(e.ColumnIndex, e.RowIndex);
                        //dataGridSearchResults.Refresh();
                    }
                    else if (e.RowIndex < 0) //it's the header cell so check all
                    {
                        DataGridViewCheckBoxColumnHeaderCell header = (DataGridViewCheckBoxColumnHeaderCell)colCheck.HeaderCell;
                        for (int iRow = 0; iRow < searcher.Matches.Count; iRow++)
                        {
                            dataGridSearchResults.Rows[iRow].Cells[e.ColumnIndex].Value = !header.CheckAll;
                            //searcher.Matches[iRow].Checked = header.CheckAll;
                            dataGridSearchResults.InvalidateCell(e.ColumnIndex, iRow);
                        }
                        //dataGridSearchResults.Refresh();
                    }
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
                    prgSearch.Invoke(new searcher_ProgressChanged_Delegate(searcher_ProgressChanged), new object[] { sender, e });
                }
                else
                {
                    prgSearch.Value = e.ProgressPercentage;

                    if (searcher.Matches.Count > 0)
                    {
                        string sMatchCount = "";
                        if (_IsSingleSearch)
                        {
                            if (searcher.Matches.Count == 1)
                                sMatchCount = searcher.Matches.Count.ToString("N0") + " match";
                            else
                                sMatchCount = searcher.Matches.Count.ToString("N0") + " matches";
                        }
                        else
                        {
                            sMatchCount = searcher.SearchTermCount.ToString("N0") + " search term";
                            if (searcher.SearchTermCount > 1) sMatchCount += "s";

                            if (searcher.Matches.Count == 1)
                                sMatchCount += " and " + searcher.Matches.Count.ToString("N0") + " match";
                            else
                                sMatchCount += " and " + searcher.Matches.Count.ToString("N0") + " matches";
                        }
                        lblSearchMatchesCount.Text = sMatchCount;
                        lblSearchMatchesCount.Visible = true;
                    }

                    if (searcher.Complete)
                    {
                        lblNoSearchMatches.Visible = (searcher.Matches.Count == 0);
                        btnSearchAdd.Enabled = (searcher.Matches.Count > 0);
                        btnFind.Enabled = true;
                        btnMultiSearch.Enabled = true;

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

                        if (!_IsSingleSearch && !SearchWasCancelled)
                        {
                            //highlight which search terms were found

                            //remove formatting by assigning to .Text
                            //remove blank lines as they will cause it to think there's a search term not matching
                            List<string> listSearchTermLines = new List<string>(richTextBoxSearch.Text.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries));
                            RemoveBlankLinesAndTrimLinesAndRemoveDuplicates(listSearchTermLines);
                            richTextBoxSearch.Text = string.Join("\r\n", listSearchTermLines.ToArray());
                            
                            //split Rtf into lines
                            string sRtf = richTextBoxSearch.Rtf.Replace("\r\n","\n").Replace("\r","\n").Replace("\n","\r\n");
                            List<string> listRtfLines = new List<string>(sRtf.Split(new string[] { "\r\n" }, StringSplitOptions.None));
                            
                            //add the font color table to line 2
                            //first color (cf1) is default or black... second color (cf2) is red
                            listRtfLines.Insert(1, @"{\colortbl;\red0\green0\blue0;\red255\green0\blue0;}");


                            int iLines = richTextBoxSearch.Lines.Length;
                            int iRtfLine = 0;
                            int iLine = 0;
                            StringBuilder sb = new StringBuilder(richTextBoxSearch.Rtf.Length + 10);
                            foreach (string sRtfLine in listRtfLines)
                            {
                                if (sRtfLine.StartsWith("{") || (sRtfLine.StartsWith(@"\") && !sRtfLine.StartsWith(@"\\") && sRtfLine != @"\par" && sRtfLine.Split(new char[] { ' ', '\t' }).Length == 1))
                                {
                                    //this line doesn't have any content so don't count it
                                }
                                else if (iLine < iLines)
                                {
                                    string sLine = richTextBoxSearch.Lines[iLine].Trim();
                                    bool bFound = false;
                                    foreach (string sSearchTermMatch in searcher.MatchedSearchTerms)
                                    {
                                        if (sSearchTermMatch.Trim() == sLine)
                                        {
                                            bFound = true;
                                            break;
                                        }
                                    }
                                    if (bFound)
                                    {
                                        sb.AppendLine(@"\cf1"); //default color
                                    }
                                    else
                                    {
                                        sb.AppendLine(@"\cf2"); //red color
                                        lblSearchTermsHighlighted.Visible = true;
                                    }
                                    iRtfLine++;
                                    iLine++;
                                }

                                sb.AppendLine(sRtfLine);

                                if (iLine == iLines) //if we just inserted the last line of text...
                                {
                                    sb.AppendLine(@"\cf1\par"); //if they type in a new line, it should be black
                                    iLine++; //make sure this block doesn't run again
                                }
                            }

                            richTextBoxSearch.Rtf = sb.ToString();

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
                if (_IsSingleSearch)
                {
                    this.AcceptButton = btnFind;
                }
                else
                {
                    this.AcceptButton = null;
                }
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
                bool bMatchChecked = false;
                foreach (CubeSearcher.CubeSearchMatch item in searcher.Matches)
                {
                    if (item.Checked)
                    {
                        bMatchChecked = true;
                        break;
                    }
                }
                if (!bMatchChecked)
                {
                    MessageBox.Show("Please check any matches you wish to add to the PivotTable first.", "OLAP PivotTable Extensions");
                    return;
                }

                AddInWorking = true;

                pvt.ManualUpdate = true; //TODO... this doesn't appear to be working

                //TODO: handle if they have searched and found the All member (using the server side search or exact match search)... currently not working

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

                Dictionary<string, int> dictMatchedCubeFields = new Dictionary<string, int>();
                Dictionary<string, bool> dictMatchedCubeFieldsIsMemberProperty = new Dictionary<string, bool>();
                Dictionary<string, List<object>> dictLevelsOfFoundMembers = new Dictionary<string, List<object>>();

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
                                catch (Exception exInner)
                                {
                                    //if neither succeeded, then raise the error
                                    throw new Exception("Failed adding member property " + item.MemberProperty.UniqueName + " to screen. Errors were " + AddOledbErrorToException(ex, false) + " and " + exInner.Message, ex);
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

                        if (hierarchiesInPivotTableAsNamedSet.ContainsKey(m.ParentLevel.ParentHierarchy.UniqueName))
                        {
                            NamedSet ns = hierarchiesInPivotTableAsNamedSet[m.ParentLevel.ParentHierarchy.UniqueName];
                            if (string.Compare("[" + ns.Name + "]", Convert.ToString(cmbLookIn.SelectedItem)) != 0)
                            {
                                continue;
                            }
                        }
                        
                        if (dictMatchedCubeFields.ContainsKey(m.ParentLevel.ParentHierarchy.UniqueName))
                        {
                            dictMatchedCubeFields[m.ParentLevel.ParentHierarchy.UniqueName]++;
                        }
                        else
                        {
                            dictMatchedCubeFields.Add(m.ParentLevel.ParentHierarchy.UniqueName, 1);
                            dictMatchedCubeFieldsIsMemberProperty.Add(m.ParentLevel.ParentHierarchy.UniqueName, false);
                        }
                        if (item.MemberProperty != null)
                            dictMatchedCubeFieldsIsMemberProperty[m.ParentLevel.ParentHierarchy.UniqueName] = true;

                        if (!dictLevelsOfFoundMembers.ContainsKey(m.ParentLevel.UniqueName))
                            dictLevelsOfFoundMembers.Add(m.ParentLevel.UniqueName, new List<object>());
                        dictLevelsOfFoundMembers[m.ParentLevel.UniqueName].Add(m.UniqueName);

                    }
                }

                foreach (string sCubeField in dictMatchedCubeFields.Keys)
                {
                    Excel.CubeField field;
                    if (hierarchiesInPivotTableAsNamedSet.ContainsKey(sCubeField))
                    {
                        NamedSet ns = hierarchiesInPivotTableAsNamedSet[sCubeField];
                        field = pvt.CubeFields.get_Item("[" + ns.Name + "]");
                        if (string.Compare("[" + ns.Name + "]", Convert.ToString(cmbLookIn.SelectedItem)) != 0)
                        {
                            //TODO: future... see if it's in that set so you don't have to prompt
                            MessageBox.Show("The named set [" + Convert.ToString(ns.Properties["SET_CAPTION"].Value) + "] containing " + sCubeField + " is in the PivotTable, so members will not show up in the PivotTable unless they are in that set.", "OLAP PivotTable Extensions");
                            continue;
                        }
                    }
                    else
                    {
                        field = pvt.CubeFields.get_Item(sCubeField);
                    }


                    try
                    {
                        field.CreatePivotFields(); //Excel apparently doesn't always have the levels loaded, so this loads them
                    }
                    catch { } 
                    
                    if (field.Orientation == Excel.XlPivotFieldOrientation.xlHidden)
                    {
                        if (dictMatchedCubeFieldsIsMemberProperty[sCubeField])
                            field.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                        else
                            field.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                    }

                    pvt.ManualUpdate = true;


                    field.IncludeNewItemsInFilter = false; //if this is set to true, they essentially wanted to show everything but what was specifically unchecked. With Filter List, we're doing the reverse... showing only what's spefically checked

                    int iLevel = 0;
                    int iMaxLevelWithMembers = 0;
                    Excel.PivotField firstPivotField = null;
                    foreach (Excel.PivotField pivotField in field.PivotFields)
                    {
                        if (pivotField.IsMemberProperty) continue;
                        if (firstPivotField == null) firstPivotField = pivotField;
                        iLevel++;
                        if (dictLevelsOfFoundMembers.ContainsKey(pivotField.Name))
                        {
                            iMaxLevelWithMembers = iLevel;

                            if (field.Orientation == Excel.XlPivotFieldOrientation.xlPageField)
                            {
                                pvt.ManualUpdate = true;
                                field.EnableMultiplePageItems = true;
                            }

                            List<object> listVisibleItems = new List<object>();
                            List<object> listExistingVisibleItems = new List<object>();
                            //must first cast to an object in .NET 4 (Excel 2016)
                            foreach (object o in (System.Array)(object)pivotField.VisibleItemsList)
                            {
                                if (chkAddToCurrentFilters.Checked)
                                    listExistingVisibleItems.Add(o);
                            }

                            foreach (object oNew in dictLevelsOfFoundMembers[pivotField.Name])
                            {
                                listVisibleItems.Add(oNew);

                                for (int i = 0; i < listExistingVisibleItems.Count; i++)
                                {
                                    if (Convert.ToString(listExistingVisibleItems[i]) == Convert.ToString(oNew))
                                    {
                                        listExistingVisibleItems.RemoveAt(i);
                                        i--;
                                    }
                                }
                            }

                            listVisibleItems.AddRange(listExistingVisibleItems);

                            System.Array arrNewVisibleItems = listVisibleItems.ToArray();
                            pivotField.VisibleItemsList = arrNewVisibleItems;
                        }
                        else
                        {
                            if (((System.Array)(object)pivotField.VisibleItemsList).Length > 1
                                || (((System.Array)(object)pivotField.VisibleItemsList).Length == 1 && !string.IsNullOrEmpty(Convert.ToString(((System.Array)(object)pivotField.VisibleItemsList).GetValue(1)))))
                            {
                                //pivotField.ClearManualFilter(); //got an error trying to do this, so guess I have to set VisibleItemsList
                                List<object> listVisibleItems = new List<object>();
                                listVisibleItems.Add(string.Empty);
                                System.Array arrNewVisibleItems = listVisibleItems.ToArray();
                                pivotField.VisibleItemsList = arrNewVisibleItems;
                            }
                        }
                        if (pivotField.PivotFilters.Count > 0)
                        {
                            pivotField.ClearValueFilters();
                            pivotField.ClearLabelFilters();
                        }
                    }

                    if (field.Orientation != Excel.XlPivotFieldOrientation.xlPageField)
                    {
                        //now expand any levels above the lowest level we have
                        //this works now because we're filtering what's on rows to the search results
                        int iCounter = 0;
                        foreach (Excel.PivotField pivotField in field.PivotFields)
                        {
                            if (pivotField.IsMemberProperty) continue;
                            iCounter++;
                            if (iCounter < iMaxLevelWithMembers
                            && !pivotField.Hidden //don't drilldown a level that's already hidden
                            && iCounter < iLevel) //don't drill down the last level
                            {
                                pvt.ManualUpdate = true;
                                pivotField.DrilledDown = true;
                            }
                        }
                    }
                }


                foreach (CubeSearcher.CubeSearchMatch item in searcher.Matches)
                {
                    if (!item.Checked) continue;
                    if (!item.IsFieldListField) 
                    {
                        Member m = (Member)item.InnerObject;
                        Excel.CubeField field;
                        if (hierarchiesInPivotTableAsNamedSet.ContainsKey(m.ParentLevel.ParentHierarchy.UniqueName))
                        {
                            NamedSet ns = hierarchiesInPivotTableAsNamedSet[m.ParentLevel.ParentHierarchy.UniqueName];
                            field = pvt.CubeFields.get_Item("[" + ns.Name + "]");
                            if (string.Compare("[" + ns.Name + "]", Convert.ToString(cmbLookIn.SelectedItem)) != 0)
                            {
                                continue;
                            }
                        }
                        else
                        {
                            field = pvt.CubeFields.get_Item(m.ParentLevel.ParentHierarchy.UniqueName);
                        }

                        sSearchFor = m.Caption + " (" + m.UniqueName + ")";

                        if (field.Orientation != Excel.XlPivotFieldOrientation.xlPageField)
                        {
                            //EnsureMemberIsVisible(field, m, true); //DrillDown above takes care of this since we're now filtering to the search results
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
                                    catch (Exception exInner)
                                    {
                                        //if neither succeeded, then raise the error
                                        throw new Exception("Failed adding member property " + item.MemberProperty.UniqueName + " to screen. Errors were " + AddOledbErrorToException(ex, false) + " and " + exInner.Message, ex);
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
                    MessageBox.Show("Problem adding to PivotTable: " + AddOledbErrorToException(ex, true), "OLAP PivotTable Extensions");
                else
                    MessageBox.Show("Problem adding " + sSearchFor + " to PivotTable: " + AddOledbErrorToException(ex, true), "OLAP PivotTable Extensions");
            }
            finally
            {
                AddInWorking = false;
                pvt.ManualUpdate = false;
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

                    //old code that isn't needed anymore now that we're having the Search feature filter fields on an axis
                    //System.Array arrOldVisibleItems = (System.Array)(object)pivotField.VisibleItemsList;
                    //List<object> listNewVisibleItems = new List<object>();
                    //bool bFound = false;
                    //foreach (object o in arrOldVisibleItems)
                    //{
                    //    listNewVisibleItems.Add(o);
                    //    if (Convert.ToString(o) == m.UniqueName)
                    //    {
                    //        bFound = true;
                    //    }
                    //}
                    //if (!bFound)
                    //{
                    //    if (!(listNewVisibleItems.Count == 1 && string.IsNullOrEmpty(Convert.ToString(listNewVisibleItems[0]))))
                    //    {
                    //        //this level is filtered, so add this member to this level's filters
                    //        listNewVisibleItems.Add(m.UniqueName);
                    //        System.Array arrNewVisibleItems = listNewVisibleItems.ToArray();
                    //        pivotField.VisibleItemsList = arrNewVisibleItems;
                    //    }
                    //}
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
                    SearchWasCancelled = true;

                    lblNoSearchMatches.Visible = (searcher.Matches.Count == 0);
                    btnSearchAdd.Enabled = (searcher.Matches.Count > 0);
                    btnFind.Enabled = true;
                    btnMultiSearch.Enabled = true;

                    prgSearch.Visible = false;
                    btnCancelSearch.Visible = false;

                    if (!string.IsNullOrEmpty(searcher.Error))
                    {
                        lblSearchError.Text = searcher.Error;
                        lblSearchError.Visible = true;
                    }
                }
            }
            catch
            {
            }
        }

        internal bool IsExcel2007OrHigherPivotTableVersion()
        {
            return ((int)pvt.Version >= (int)Excel.XlPivotTableVersionList.xlPivotTableVersion12);
        }

        private string GetPivotTableVersion()
        {
            if (pvt.Version == Excel.XlPivotTableVersionList.xlPivotTableVersion2000)
                return "2000";
            else if (pvt.Version == Excel.XlPivotTableVersionList.xlPivotTableVersion10)
                return "2002";
            else if (pvt.Version == Excel.XlPivotTableVersionList.xlPivotTableVersion11)
                return "2003";
            else if (pvt.Version == Excel.XlPivotTableVersionList.xlPivotTableVersion12)
                return "2007";
            else if ((int)pvt.Version == xlPivotTableVersion14) //since we're using the Excel 2007 object model, the Excel 2010 version isn't visible
                return "2010";
            else if ((int)pvt.Version == xlPivotTableVersion15) //since we're using the Excel 2007 object model, the Excel 2013 version isn't visible
                return "2013";
            else if ((int)pvt.Version == xlPivotTableVersion16) //since we're using the Excel 2007 object model, the Excel 2013 version isn't visible
                return "2016";
            else
                return pvt.Version.ToString();
        }

        private string GetExcelVersion()
        {
            int iVersion = (int)decimal.Parse(application.Version, System.Globalization.CultureInfo.InvariantCulture.NumberFormat);
            if (iVersion == 12)
                return "2007";
            else if (iVersion == 14)
                return "2010";
            else if (iVersion == 15)
                return "2013";
            else if (iVersion == 16)
                return "2016";
            else
                return "Unknown";
        }

        private void cmbLookIn_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                chkMemberProperties.Checked = (cmbLookIn.SelectedIndex == 0);
            }
            catch { }
            try
            {
                if (cmbLookIn.SelectedIndex == 0)
                {
                    Connect.SearchMeasuresOnlyDefault = false;
                }
                else if (cmbLookIn.SelectedIndex == 1)
                {
                    Connect.SearchMeasuresOnlyDefault = true;
                }
            }
            catch { }
            try
            {
                if (Convert.ToString(cmbLookIn.SelectedItem) == LOOK_IN_SHOW_ALL_HIERARCHIES)
                {
                    try
                    {
                        _SearchLookInShowAllHierarchies = true;
                        FillSearchLookInDropdown();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);
                    }
                }
            }
            catch { }
        }

        private void MainForm_Activated(object sender, EventArgs e)
        {
            try
            {
                if (tabControl.SelectedTab == tabSearch)
                {
                    if (_IsSingleSearch && !txtSearch.Focused && !cmbLookIn.Focused)
                    {
                        txtSearch.Focus();
                        txtSearch.SelectAll();
                    }
                    else if (!_IsSingleSearch && !richTextBoxSearch.Focused && !cmbLookIn.Focused)
                    {
                        richTextBoxSearch.Focus();
                        richTextBoxSearch.SelectAll();
                    }
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
                btnFilterListShowCurrentFilters.Enabled = false;
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

                SetCulture(application);

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
                        restrictions.Add(new AdomdRestriction("MEMBER_CAPTION", sLine.Trim()));
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

                    SetFilterListProgress((int)(90 * (++iNumLinesFinished) / args.Lines.Length), true, null, true);
                }

                Excel.CubeField field = pvt.CubeFields.get_Item(args.LookIn);
                field.CreatePivotFields();
                field.IncludeNewItemsInFilter = false; //if this is set to true, they essentially wanted to show everything but what was specifically unchecked. With Filter List, we're doing the reverse... showing only what's spefically checked

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

                SetFilterListProgress(100, false, listMembersNotFound.ToArray(), true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);

                SetFilterListProgress(0, false, listMembersNotFound.ToArray(), true);
            }
            finally
            {
                ResetCulture(application);
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

        private delegate void SetFilterListProgress_Delegate(int iProgress, bool bVisible, string[] arrMembersNotFound, bool bCloseIfSuccessful);
        private void SetFilterListProgress(int iProgress, bool bVisible, string[] arrMembersNotFound, bool bCloseIfSuccessful)
        {
            if (progressFilterList.InvokeRequired)
            {
                //avoid the "cross-thread operation not valid" error message
                progressFilterList.Invoke(new SetFilterListProgress_Delegate(SetFilterListProgress), new object[] { iProgress, bVisible, arrMembersNotFound, bCloseIfSuccessful });
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
                    btnFilterListShowCurrentFilters.Enabled = true;
                    txtFilterList.ReadOnly = false;

                    if (arrMembersNotFound.Length == 0)
                    {
                        if (bCloseIfSuccessful)
                            this.Close();
                    }
                    else
                    {
                        string sError = "The following members were not found.\r\n";
                        if (arrMembersNotFound.Length > 10) sError += " (Showing first 10)\r\n";
                        sError += "\r\n" + string.Join("\r\n", arrMembersNotFound);
                        MessageBox.Show(sError);
                    }
                }
            }
        }

        private void btnFilterListShowCurrentFilters_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbFilterListLookIn.SelectedIndex < 0)
                {
                    MessageBox.Show("Choose a field first.");
                    return;
                }

                FilterListWorkerArgs args = new FilterListWorkerArgs();
                args.Lines = GetSelectedMemberUniqueNames(Convert.ToString(this.cmbFilterListLookIn.SelectedItem));
                args.LookIn = Convert.ToString(cmbFilterListLookIn.SelectedItem);

                if (args.Lines.Length == 0)
                    return;

                progressFilterList.Visible = true;
                progressFilterList.Value = 0;

                btnCancelFilterList.Visible = true;
                btnFilterList.Enabled = false;
                btnFilterListShowCurrentFilters.Enabled = false;
                txtFilterList.ReadOnly = true;

                workerFilterList = new BackgroundWorker();
                workerFilterList.DoWork += new DoWorkEventHandler(workerFilterList_ShowCurrentFilters_DoWork);
                workerFilterList.WorkerSupportsCancellation = true;
                workerFilterList.RunWorkerAsync(args);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        private string[] GetSelectedMemberUniqueNames(string CubeFieldName)
        {
            List<string> selectedMembers = new List<string>();
            Excel.CubeField field = pvt.CubeFields.get_Item(CubeFieldName);
            if (field.IncludeNewItemsInFilter == false)
            {
                field.CreatePivotFields();

                foreach (Excel.PivotField pivotField in field.PivotFields)
                {
                    if (!pivotField.IsMemberProperty)
                    {
                        System.Array arrNewVisibleItems = (System.Array)(object)pivotField.VisibleItemsList;
                        foreach (string sMember in arrNewVisibleItems)
                        {
                            if (!string.IsNullOrEmpty(sMember))
                            {
                                selectedMembers.Add(sMember);
                            }
                        }
                    }
                }
            }
            return selectedMembers.ToArray();
        }

        void workerFilterList_ShowCurrentFilters_DoWork(object sender, DoWorkEventArgs e)
        {
            List<string> listMembersNotFound = new List<string>();

            try
            {
                AddInWorking = true;

                SetCulture(application);

                ConnectAdomdClientCube();

                if (e.Cancel) return;

                FilterListWorkerArgs args = (FilterListWorkerArgs)e.Argument;

                AdomdCommand cmd = new AdomdCommand();
                cmd.Connection = cube.ParentConnection;

                StringBuilder sFoundMemberCaptions = new StringBuilder();

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
                        restrictions.Add(new AdomdRestriction("MEMBER_UNIQUE_NAME", sLine.Trim()));
                        System.Data.DataTable tblExactMatchMembers = cube.ParentConnection.GetSchemaDataSet("MDSCHEMA_MEMBERS", restrictions).Tables[0];

                        if (tblExactMatchMembers.Rows.Count > 0)
                        {
                            foreach (System.Data.DataRow row in tblExactMatchMembers.Rows)
                            {
                                sFoundMemberCaptions.Append(Convert.ToString(row["MEMBER_CAPTION"])).AppendLine();
                            }
                        }
                    }

                    SetFilterListProgress((int)(90 * (++iNumLinesFinished) / args.Lines.Length), true, null, false);
                }

                txtFilterList.Text = sFoundMemberCaptions.ToString();

                SetFilterListProgress(100, false, listMembersNotFound.ToArray(), false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);

                SetFilterListProgress(0, false, listMembersNotFound.ToArray(), false);
            }
            finally
            {
                ResetCulture(application);
                AddInWorking = false;
            }
        }

        private void btnUpgradeOnRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                pvt.PivotCache().UpgradeOnRefresh = true;
                btnUpgradeOnRefresh.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(AddOledbErrorToException(ex, true), "OLAP PivotTable Extensions");
            }
        }

        private void InitiateFormatMDX(string MDX)
        {
            workerFormatMDX = new BackgroundWorker();
            workerFormatMDX.DoWork += new DoWorkEventHandler(workerFormatMDX_DoWork);
            workerFormatMDX.RunWorkerAsync(MDX);

            if (!lblFormattingMdxQuery.Visible)
                richTextBoxMDX.Height -= (lblFormattingMdxQuery.Height + 5);
            lblFormattingMdxQuery.ForeColor = System.Drawing.Color.Black;
            lblFormattingMdxQuery.Text = "Formatting MDX query in progress...";
            tooltip.SetToolTip(lblFormattingMdxQuery, "Calling web service...");
            lblFormattingMdxQuery.Visible = true;
        }

        void workerFormatMDX_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                com.msftlabs.formatmdx.Formatter formatter = new com.msftlabs.formatmdx.Formatter();
                com.msftlabs.formatmdx.Settings settings = new com.msftlabs.formatmdx.Settings();
                settings.AdjustCase = false;
                settings.CommaPlacement = com.msftlabs.formatmdx.CommaPlacementEnum.BegginingOfLine;
                settings.OpenBraceAfterFunctionOrSubselectOnNewLine = false;
                settings.SpacesPerIdent = 1;
                settings.TabAsIdent = false;
                formatter.Proxy = System.Net.WebRequest.GetSystemWebProxy(); //use current IE proxy settings
                string sMdxRtf = formatter.FormatAsRtfWithSettings(e.Argument.ToString(), settings);
                SetFormattedMDX(sMdxRtf, null);
            }
            catch (Exception ex)
            {
                SetFormattedMDX(null, ex);
            }
        }

        private void chkFormatMDX_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkFormatMDX.Enabled)
                {
                    Connect.FormatMdx = chkFormatMDX.Checked;
                    if (Connect.FormatMdx)
                    {
                        if (string.IsNullOrEmpty(richTextBoxMDX.Text))
                            tabControl_SelectedIndexChanged(null, null);
                        else
                            InitiateFormatMDX(richTextBoxMDX.Text);
                    }
                }
            }
            catch (Exception exInner)
            {
                MessageBox.Show(exInner.Message + "\r\n" + exInner.StackTrace);
            }
        }

        private delegate void SetFormattedMDX_Delegate(string MDX, Exception ex);
        private void SetFormattedMDX(string MDX, Exception ex)
        {
            try
            {
                if (richTextBoxMDX.InvokeRequired)
                {
                    //avoid the "cross-thread operation not valid" error message
                    richTextBoxMDX.Invoke(new SetFormattedMDX_Delegate(SetFormattedMDX), new object[] { MDX, ex });
                }
                else
                {
                    if (ex == null)
                    {
                        richTextBoxMDX.Rtf = MDX;
                        richTextBoxMDX.SelectionStart = 0;
                        richTextBoxMDX.SelectionLength = richTextBoxMDX.Text.Length;
                        richTextBoxMDX.Focus();
                        richTextBoxMDX.ScrollToCaret();

                        lblFormattingMdxQuery.Visible = false;
                        richTextBoxMDX.Height += lblFormattingMdxQuery.Height + 5;
                    }
                    else
                    {
                        lblFormattingMdxQuery.ForeColor = System.Drawing.Color.Red;
                        lblFormattingMdxQuery.Text = "An error occurred formatting MDX query. Mouse over to see error.";
                        com.msftlabs.formatmdx.Formatter formatter = new com.msftlabs.formatmdx.Formatter();
                        tooltip.SetToolTip(lblFormattingMdxQuery, "Problem formatting MDX using the " + formatter.Url + " web service. Error was:\r\n\r\n" + ex.Message + "\r\n" + ex.StackTrace);
                    }
                }
            }
            catch (Exception exInner)
            {
                MessageBox.Show(exInner.Message + "\r\n" + exInner.StackTrace);
            }
        }

        bool _IsSingleSearch = true;
        private void btnMultiSearch_Click(object sender, EventArgs e)
        {
            try
            {
                int iMove = 120;
                if (!_IsSingleSearch)
                {
                    iMove = -iMove;
                    txtSearch.Visible = true;
                    txtSearch.Focus();
                    txtSearch.SelectAll();
                    richTextBoxSearch.Visible = false;
                    txtSearch.Text = string.Empty;
                    btnMultiSearch.Text = "↓";
                    tooltip.SetToolTip(this.btnMultiSearch, "Allow entering multiple search terms");
                }
                else
                {
                    richTextBoxSearch.Location = txtSearch.Location;
                    richTextBoxSearch.Size = txtSearch.Size;
                    richTextBoxSearch.Text = txtSearch.Text;
                    richTextBoxSearch.Visible = true;
                    richTextBoxSearch.Focus();
                    richTextBoxSearch.SelectAll();
                    txtSearch.Visible = false;
                    btnMultiSearch.Text = "↑";
                    tooltip.SetToolTip(this.btnMultiSearch, "Search for one search term at a time");
                    txtSearch_Leave(null, null); //turn off Enter causing a search to start
                }

                this.Height += iMove;
                Size minSizeNew = new Size(MainForm.ActiveForm.MinimumSize.Width, MainForm.ActiveForm.MinimumSize.Height);
                this.MinimumSize = minSizeNew;

                lblNoSearchMatches.Top += iMove;
                dataGridSearchResults.Top += iMove;
                dataGridSearchResults.Height += -iMove;
                chkAddToCurrentFilters.Top += iMove;
                chkExactMatch.Top += iMove;
                cmbLookIn.Top += iMove;
                lblLookIn.Top += iMove;

                chkMemberProperties.Top += iMove;

                txtSearch.Height += iMove;
                richTextBoxSearch.Height += iMove;
                _IsSingleSearch = !_IsSingleSearch;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        private void richTextBoxSearch_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control && e.KeyCode == Keys.V)
                {
                    //paste in as unformatted text
                    string s = (string)Clipboard.GetData("Text");
                    if (s != null)
                        richTextBoxSearch.SelectedText = s;

                    // cancel the actual paste
                    e.Handled = true;
                }
            }
            catch { }
        }

    }
}
