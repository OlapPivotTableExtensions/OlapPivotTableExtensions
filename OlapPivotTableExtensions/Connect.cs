using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OlapPivotTableExtensions
{
    public partial class Connect
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                OriginalLanguage = System.Threading.Thread.CurrentThread.CurrentCulture.EnglishName;
                OriginalCultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;

                Application.SheetBeforeRightClick += m_xlAppEvents_xlSheetBeforeRightClick;
                Application.SheetPivotTableUpdate += m_xlAppEvents_xlSheetPivotTableUpdate;

                //m_xlAppEvents = new xlEvents();
                //m_xlAppEvents.DisableEventsIfEmbedded = true;
                //???? m_xlAppEvents.SetupConnection(Application);

                //the following code works around an issue that is surfaced by typical event handling
                //http://olappivottableextend.codeplex.com/discussions/271174
                //typical event handling: Application.SheetBeforeRightClick += new Microsoft.Office.Interop.Excel.AppEvents_SheetBeforeRightClickEventHandler(Application_SheetBeforeRightClick);
                //m_xlAppEvents.xlSheetBeforeRightClick += new xlEvents.DSheetBeforeRightClick(m_xlAppEvents_xlSheetBeforeRightClick);
                //m_xlAppEvents.xlSheetPivotTableUpdate += new xlEvents.DSheetPivotTableUpdate(m_xlAppEvents_xlSheetPivotTableUpdate);

                try
                {
                    MainForm.SetCulture(Application);

                    ExcelVersion = (int)decimal.Parse(Application.Version, System.Globalization.CultureInfo.InvariantCulture.NumberFormat);
                    if (ExcelVersion >= 15)
                    {
                        IsSingleDocumentInterface = true;
                    }
                }
                finally
                {
                    MainForm.ResetCulture(Application);
                }

                if (IsSingleDocumentInterface)
                {
                    //m_xlAppEvents.xlWindowActivate += new xlEvents.DWindowActivate(m_xlAppEvents_xlWindowActivate);
                    Application.WindowActivate += m_xlAppEvents_xlWindowActivate;
                }

                CreateOlapPivotTableExtensionsMenu();

                AppDomain currentDomain = AppDomain.CurrentDomain;
                currentDomain.AssemblyResolve += new ResolveEventHandler(currentDomain_AssemblyResolve);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during startup of OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
        }
        
        private void m_xlAppEvents_xlWindowActivate(Excel._Workbook oWB, Excel.Window oWn)
        {
            try
            {
                MainForm.SetCulture(Application);
                CreateOlapPivotTableExtensionsMenu();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during WindowActivate:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
            finally
            {
                MainForm.ResetCulture(Application);
            }
        }

        //the Microsoft.Excel.AdomdClient.dll used for Excel Data Models in Excel 15 isn't in any of the paths .NET looks for assemblies in... so we have to catch the AssemblyResolve event and manually load that assembly
        private static AdomdClientWrappers.ExcelAdoMdConnections _helper = new AdomdClientWrappers.ExcelAdoMdConnections();
        System.Reflection.Assembly currentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("AssemblyResolve: " + args.Name);
                if (args.Name.Contains("Microsoft.Excel.AdomdClient"))
                {
                    return _helper.ExcelAdomdClientAssembly;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during AssemblyResolve in OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
                return null;
            }
        }

        /// <summary>
        ///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
        ///     Receives notification that the Add-in is being unloaded.
        /// </summary>
        /// <param term='disconnectMode'>
        ///      Describes how the Add-in is being unloaded.
        /// </param>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        //public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        //{
        //    DeleteOlapPivotTableExtensionsMenu();
        //}

        /// <summary>
        ///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
        ///      Receives notification that the collection of Add-ins has changed.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnAddInsUpdate(ref System.Array custom)
        {
        }

        /// <summary>
        ///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application has completed loading.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref System.Array custom)
        {
        }

        /// <summary>
        ///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application is being unloaded.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref System.Array custom)
        {
            try
            {
                m_xlAppEvents.RemoveConnection();
                m_xlAppEvents = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Application);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during close of OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
        }

        //private Excel.Application Application;
        private object addInInstance;
        private xlEvents m_xlAppEvents;

        public static string OriginalLanguage = "";
        public static System.Globalization.CultureInfo OriginalCultureInfo;

        private int ExcelVersion;
        private bool IsSingleDocumentInterface;
        private bool IsEmbedded = false;

        private const string REGISTRY_BASE_PATH = "SOFTWARE\\OLAP PivotTable Extensions";
        private const string REGISTRY_PATH_SHOW_CALC_MEMBERS_BY_DEFAULT = "ShowCalcMembersByDefault";
        private const string REGISTRY_PATH_REFRESH_DATA_BY_DEFAULT = "RefreshDataByDefault";
        private const string REGISTRY_PATH_SEARCH_MEASURES_ONLY_DEFAULT = "SearchMeasuresOnlyByDefault";
        private const string REGISTRY_PATH_FORMAT_MDX = "FormatMDX";
        //private global::System.Object missing = global::System.Type.Missing;


        private const string MENU_TAG = "OLAP PivotTable Extensions";
        private const string PIVOTTABLE_CONTEXT_MENU = "PivotTable Context Menu";
        private Office.CommandBarButton cmdMenuItem = null;
        private Office.CommandBarButton cmdSearchMenuItem = null;
        private Office.CommandBarButton cmdFilterListMenuItem = null;
        private Office.CommandBarButton cmdChooseFieldsMenuItem = null;
        private Office.CommandBarPopup cmdShowPropertyAsCaptionMenuItem = null;
        private Office.CommandBarButton cmdClearPivotTableCacheMenuItem = null;
        private Office.CommandBarButton cmdErrorMenuItem = null;
        private Office.CommandBarButton cmdDisableAutoRefresh = null;

        private MainForm frm;

        private static bool? _ShowCalcMembersByDefaultCached = null;
        public static bool ShowCalcMembersByDefault
        {
            get
            {
                if (_ShowCalcMembersByDefaultCached == null)
                {
                    Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(REGISTRY_BASE_PATH);
                    _ShowCalcMembersByDefaultCached = ((int)regKey.GetValue(REGISTRY_PATH_SHOW_CALC_MEMBERS_BY_DEFAULT, 0) == 1) ? true : false;
                    regKey.Close();
                }
                return (bool)_ShowCalcMembersByDefaultCached;
            }
            set
            {
                Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(REGISTRY_BASE_PATH);
                regKey.SetValue(REGISTRY_PATH_SHOW_CALC_MEMBERS_BY_DEFAULT, value, Microsoft.Win32.RegistryValueKind.DWord);
                regKey.Close();
                _ShowCalcMembersByDefaultCached = value;
            }
        }

        private static bool? _RefreshDataByDefaultCached = null;
        public static bool RefreshDataByDefault
        {
            get
            {
                if (_RefreshDataByDefaultCached == null)
                {
                    Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(REGISTRY_BASE_PATH);
                    _RefreshDataByDefaultCached = ((int)regKey.GetValue(REGISTRY_PATH_REFRESH_DATA_BY_DEFAULT, 0) == 1) ? true : false;
                    regKey.Close();
                }
                return (bool)_RefreshDataByDefaultCached;
            }
            set
            {
                Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(REGISTRY_BASE_PATH);
                regKey.SetValue(REGISTRY_PATH_REFRESH_DATA_BY_DEFAULT, value, Microsoft.Win32.RegistryValueKind.DWord);
                regKey.Close();
                _RefreshDataByDefaultCached = value;
            }
        }

        private static bool? _FormatMdxCached = null;
        public static bool FormatMdx
        {
            get
            {
                if (_FormatMdxCached == null)
                {
                    Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(REGISTRY_BASE_PATH);
                    _FormatMdxCached = ((int)regKey.GetValue(REGISTRY_PATH_FORMAT_MDX, 0) == 1) ? true : false;
                    regKey.Close();
                }
                return (bool)_FormatMdxCached;
            }
            set
            {
                Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(REGISTRY_BASE_PATH);
                regKey.SetValue(REGISTRY_PATH_FORMAT_MDX, value, Microsoft.Win32.RegistryValueKind.DWord);
                regKey.Close();
                _FormatMdxCached = value;
            }
        }

        private static bool? _SearchMeasuresOnlyDefault = null;
        public static bool SearchMeasuresOnlyDefault
        {
            get
            {
                if (_SearchMeasuresOnlyDefault == null)
                {
                    Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(REGISTRY_BASE_PATH);
                    _SearchMeasuresOnlyDefault = ((int)regKey.GetValue(REGISTRY_PATH_SEARCH_MEASURES_ONLY_DEFAULT, 0) == 1) ? true : false;
                    regKey.Close();
                }
                return (bool)_SearchMeasuresOnlyDefault;
            }
            set
            {
                Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(REGISTRY_BASE_PATH);
                regKey.SetValue(REGISTRY_PATH_SEARCH_MEASURES_ONLY_DEFAULT, value, Microsoft.Win32.RegistryValueKind.DWord);
                regKey.Close();
                _SearchMeasuresOnlyDefault = value;
            }
        }

        private void CreateOlapPivotTableExtensionsMenu()
        {
            try
            {
                DeleteOlapPivotTableExtensionsMenu();

                //if this is an embedded Excel document in a Word or PowerPoint document, then detect this and don't create menus
                try
                {
                    Excel._Workbook wb = (Excel._Workbook)Application.ActiveWorkbook;
#if VSTO40
                    IsEmbedded = wb.IsInplace;
#else
                    IsEmbedded = this.m_xlAppEvents.IsEmbedded(ref wb);
                    this.m_xlAppEvents.ComRelease(wb);
#endif
                }
                catch { }

                if (IsEmbedded) return;

                Office.CommandBar ptcon = Application.CommandBars[PIVOTTABLE_CONTEXT_MENU];

                cmdSearchMenuItem = (Office.CommandBarButton)ptcon.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);
                cmdSearchMenuItem.Caption = "Search...";
                cmdSearchMenuItem.FaceId = 1733;
                cmdSearchMenuItem.Tag = MENU_TAG;
                cmdSearchMenuItem.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdSearchMenuItem_Click);


                cmdClearPivotTableCacheMenuItem = (Office.CommandBarButton)ptcon.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);
                cmdClearPivotTableCacheMenuItem.Caption = "Clear PivotTable Cache";
                cmdClearPivotTableCacheMenuItem.FaceId = 47;
                cmdClearPivotTableCacheMenuItem.Tag = MENU_TAG;
                cmdClearPivotTableCacheMenuItem.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdClearPivotTableCacheMenuItem_Click);


                Office.CommandBarPopup popupFilter = null;
                try
                {
                    //find the Filter sub-menu under the PivotTable context menu by ID 31404
                    popupFilter = (Office.CommandBarPopup)Application.CommandBars.FindControl(Office.MsoControlType.msoControlPopup, 31404, missing, missing);
                }
                catch { }
                if (popupFilter != null)
                    cmdFilterListMenuItem = (Office.CommandBarButton)popupFilter.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);
                else
                    cmdFilterListMenuItem = (Office.CommandBarButton)ptcon.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);

                cmdFilterListMenuItem.Caption = "Filter List...";
                cmdFilterListMenuItem.FaceId = 517;
                cmdFilterListMenuItem.Tag = MENU_TAG;
                cmdFilterListMenuItem.BeginGroup = true;
                cmdFilterListMenuItem.Visible = false;
                cmdFilterListMenuItem.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdFilterListMenuItem_Click);


                Office.CommandBarPopup popupShowHideFields = null;
                try
                {
                    //find the Show/Hide Fields sub-menu under the PivotTable context menu by ID 31406
                    popupShowHideFields = (Office.CommandBarPopup)Application.CommandBars.FindControl(Office.MsoControlType.msoControlPopup, 31406, missing, missing);
                }
                catch { }
                if (popupShowHideFields != null)
                    cmdChooseFieldsMenuItem = (Office.CommandBarButton)popupShowHideFields.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);
                else
                    cmdChooseFieldsMenuItem = (Office.CommandBarButton)ptcon.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);

                cmdChooseFieldsMenuItem.Caption = "Choose Fields to Show...";
                cmdChooseFieldsMenuItem.FaceId = 222;
                cmdChooseFieldsMenuItem.Tag = MENU_TAG;
                cmdChooseFieldsMenuItem.BeginGroup = true;
                cmdChooseFieldsMenuItem.Visible = false;
                cmdChooseFieldsMenuItem.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdChooseFieldsMenuItem_Click);


                cmdDisableAutoRefresh = (Office.CommandBarButton)ptcon.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);
                cmdDisableAutoRefresh.Caption = "Disable Auto Refresh";
                cmdDisableAutoRefresh.FaceId = 1919;
                cmdDisableAutoRefresh.Tag = MENU_TAG;
                cmdDisableAutoRefresh.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdDisableAutoRefresh_Click);

                cmdMenuItem = (Office.CommandBarButton)ptcon.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);
                cmdMenuItem.Caption = "OLAP PivotTable Extensions...";
                cmdMenuItem.FaceId = 1122;
                cmdMenuItem.Tag = MENU_TAG;
                cmdMenuItem.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdMenuItem_Click);

                //foreach (object btn in ptcon.Controls)
                //{
                //    if (btn is Office.CommandBarButton)
                //    {
                //        Office.CommandBarButton mybtn = (Office.CommandBarButton)btn;
                //        System.Diagnostics.Debug.WriteLine(mybtn.Caption + " - " + mybtn.Id);
                //    }
                //    else if (btn is Office.CommandBarPopup)
                //    {
                //        Office.CommandBarPopup mybtn = (Office.CommandBarPopup)btn;
                //        System.Diagnostics.Debug.WriteLine(mybtn.Caption + " - " + mybtn.Id);
                //    }
                //}

                object popupAdditionalActionsIndex = System.Reflection.Missing.Value;
                try
                {
                    //find the Additional Actions sub-menu under the PivotTable context menu by ID 31595
                    Office.CommandBarPopup popup = (Office.CommandBarPopup)Application.CommandBars.FindControl(Office.MsoControlType.msoControlPopup, 31595, missing, missing);
                    popupAdditionalActionsIndex = popup.Index - 2; //not sure why -2 works
                }
                catch { }

                //add this button before the Additional Actions button
                cmdShowPropertyAsCaptionMenuItem = (Office.CommandBarPopup)ptcon.Controls.Add(Office.MsoControlType.msoControlPopup, System.Reflection.Missing.Value, System.Reflection.Missing.Value, popupAdditionalActionsIndex, true);
                cmdShowPropertyAsCaptionMenuItem.Caption = "Show Property as Caption";
                cmdShowPropertyAsCaptionMenuItem.Tag = MENU_TAG;

                cmdErrorMenuItem = (Office.CommandBarButton)ptcon.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);
                cmdErrorMenuItem.Caption = "View Error...";
                cmdErrorMenuItem.FaceId = 463;
                cmdErrorMenuItem.Tag = MENU_TAG;
                cmdErrorMenuItem.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdErrorMenuItem_Click);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during startup of OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
        }

        void cmdChooseFieldsMenuItem_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                if (Ctrl.Tag != cmdChooseFieldsMenuItem.Tag || Ctrl.Caption != cmdChooseFieldsMenuItem.Caption || Ctrl.FaceId != cmdChooseFieldsMenuItem.FaceId)
                    return;

                Excel.CubeField cf = Application.ActiveCell.PivotCell.PivotField.CubeField;

                LevelChooserForm frm = new LevelChooserForm(cf, Application.ActiveCell.PivotTable);
                frm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem showing Level Chooser:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
                return;
            }

        }

        void cmdClearPivotTableCacheMenuItem_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                if (Ctrl.Tag != cmdClearPivotTableCacheMenuItem.Tag || Ctrl.Caption != cmdClearPivotTableCacheMenuItem.Caption || Ctrl.FaceId != cmdClearPivotTableCacheMenuItem.FaceId)
                    return;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during startup of OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
                return;
            }

            string sErrorLocation = "";

            try
            {
                MainForm.SetCulture(Application);

                Excel.PivotTable pvt = Application.ActiveCell.PivotTable;
                Microsoft.Office.Interop.Excel.PivotCache cache = pvt.PivotCache();

                if (!IsOledbConnection(pvt))
                {
                    MessageBox.Show("Clear PivotTable Cache is not supported on this connection!", "OLAP PivotTable Extensions");
                    return;
                }

                sErrorLocation = "Initial MakeConnection";
                cache.WorkbookConnection.OLEDBConnection.MaintainConnection = true;
                if (!cache.IsConnected)
                    cache.MakeConnection();

                ADODB.Connection connADO = cache.ADOConnection as ADODB.Connection;
                if (connADO == null) throw new Exception("Could not cast PivotCache.ADOConnection to ADODB.Connection.");

                sErrorLocation = "Caching old connection string info";
                string sConnectionFile = cache.WorkbookConnection.OLEDBConnection.SourceConnectionFile;
                bool bUseConnectionFile = cache.WorkbookConnection.OLEDBConnection.AlwaysUseConnectionFile;
                string sConnectionString = connADO.ConnectionString;
                Excel.WorkbookConnection connOld = cache.WorkbookConnection;

                if (cache.WorkbookConnection.OLEDBConnection.CommandType != Excel.XlCmdType.xlCmdCube)
                    throw new Exception("Connection command type is not Cube. This functionality is not supported in this scenario.");

                sErrorLocation = "Determining number of PivotTables sharing connection";
                int iPivotTablesSharingConnection = 0;
                foreach (Excel.PivotCache otherCache in Application.ActiveWorkbook.PivotCaches())
                {
                    if (connOld.Name == otherCache.WorkbookConnection.Name)
                    {
                        iPivotTablesSharingConnection++;
                    }
                }
                if (iPivotTablesSharingConnection > 1)
                {
                    if (MessageBox.Show("There are multiple PivotTables using this same connection. For this feature to work, this PivotTable must be on its own connection.\r\n\r\nWould you like to move this PivotTable to a new connection?", "OLAP PivotTable Extensions", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        int iSuffix = 1;
                        bool bSuffixTaken = true;
                        while (bSuffixTaken)
                        {
                            bSuffixTaken = false;
                            foreach (Excel.WorkbookConnection otherConn in Application.ActiveWorkbook.Connections)
                            {
                                if (connOld.Name + iSuffix == otherConn.Name)
                                {
                                    bSuffixTaken = true;
                                    iSuffix++;
                                    break;
                                }
                            }
                        }

                        System.Collections.Generic.Dictionary<string, string> dictCalculatedMembers = new System.Collections.Generic.Dictionary<string, string>();
                        System.Collections.Generic.Dictionary<string, string> dictCalculatedSets = new System.Collections.Generic.Dictionary<string, string>();
                        System.Collections.Generic.Dictionary<string, Excel.XlPivotFieldOrientation> dictCalculatedMemberOrientations = new System.Collections.Generic.Dictionary<string, Excel.XlPivotFieldOrientation>();
                        foreach (Excel.CalculatedMember memb in pvt.CalculatedMembers)
                        {
                            if (memb.Type == Excel.XlCalculatedMemberType.xlCalculatedMember)
                            {
                                dictCalculatedMembers.Add(memb.Name, memb.Formula);
                                dictCalculatedMemberOrientations.Add(memb.Name, pvt.CubeFields[memb.Name].Orientation);
                            }
                            else
                            {
                                dictCalculatedSets.Add(memb.Name, memb.Formula);
                                dictCalculatedMemberOrientations.Add(memb.Name, pvt.CubeFields[memb.Name].Orientation);
                            }
                        }

                        if (dictCalculatedMemberOrientations.Count > 0)
                        {
                            MessageBox.Show("Due to a bug in Excel, calculated members and sets are removed from the PivotTable during changing the connection. OLAP PivotTable Extensions will attempt to restore them. After clearing this PivotTable's cache completes, ensure that the fields are still in the right location and order. If they are not right, fixing their order and location then rerunning this command should fix the problem.", "OLAP PivotTable Extensions");
                        }

                        Excel.WorkbookConnection connNew = Application.ActiveWorkbook.Connections.Add(cache.WorkbookConnection.Name + iSuffix, "", "OLEDB;" + sConnectionString + ";Cube=" + connOld.OLEDBConnection.CommandText, connOld.OLEDBConnection.CommandText, connOld.OLEDBConnection.CommandType);
                        connNew.OLEDBConnection.MaintainConnection = true;
                        connNew.OLEDBConnection.RefreshOnFileOpen = connOld.OLEDBConnection.RefreshOnFileOpen;
                        connNew.OLEDBConnection.RefreshPeriod = connOld.OLEDBConnection.RefreshPeriod;
                        connNew.OLEDBConnection.RetrieveInOfficeUILang = connOld.OLEDBConnection.RetrieveInOfficeUILang;
                        connNew.OLEDBConnection.RobustConnect = connOld.OLEDBConnection.RobustConnect;
                        connNew.OLEDBConnection.ServerCredentialsMethod = connOld.OLEDBConnection.ServerCredentialsMethod;
                        connNew.OLEDBConnection.ServerFillColor = connOld.OLEDBConnection.ServerFillColor;
                        connNew.OLEDBConnection.ServerFontStyle = connOld.OLEDBConnection.ServerFontStyle;
                        connNew.OLEDBConnection.ServerNumberFormat = connOld.OLEDBConnection.ServerNumberFormat;
                        connNew.OLEDBConnection.ServerSSOApplicationID = connOld.OLEDBConnection.ServerSSOApplicationID;
                        connNew.OLEDBConnection.ServerTextColor = connOld.OLEDBConnection.ServerTextColor;
                        connNew.OLEDBConnection.SourceConnectionFile = connOld.OLEDBConnection.SourceConnectionFile;
                        connNew.OLEDBConnection.AlwaysUseConnectionFile = connOld.OLEDBConnection.AlwaysUseConnectionFile;
                        pvt.ChangeConnection(connNew);
                        cache = pvt.PivotCache();
                        pvt.PivotCache().MakeConnection();

                        string sCurrentCalcName = null;
                        try
                        {
                            foreach (string sName in dictCalculatedMemberOrientations.Keys)
                            {
                                sCurrentCalcName = sName;
                                if (dictCalculatedMembers.ContainsKey(sName))
                                {
                                    Excel.CalculatedMember memb = pvt.CalculatedMembers.Add(sName, dictCalculatedMembers[sName], System.Reflection.Missing.Value, Excel.XlCalculatedMemberType.xlCalculatedMember);
                                }
                                else
                                {
                                    Excel.CalculatedMember memb = pvt.CalculatedMembers.Add(sName, dictCalculatedSets[sName], System.Reflection.Missing.Value, Excel.XlCalculatedMemberType.xlCalculatedSet);
                                }
                            }

                            sCurrentCalcName = null;
                            pvt.RefreshTable();

                            foreach (string sName in dictCalculatedMemberOrientations.Keys)
                            {
                                sCurrentCalcName = sName;
                                pvt.CubeFields.get_Item(sName).Orientation = dictCalculatedMemberOrientations[sName];
                            }
                        }
                        catch (Exception ex)
                        {
                            if (sCurrentCalcName != null)
                                throw new Exception("Problem adding " + sCurrentCalcName + " to the PivotTable. Error was: " + ex.Message, ex);
                            else
                                throw new Exception("Problem adding calculated members/sets to the PivotTable. Error was: " + ex.Message, ex);
                        }

                        connADO = cache.ADOConnection as ADODB.Connection;
                        if (connADO == null) throw new Exception("Could not cast PivotCache.ADOConnection to ADODB.Connection.");
                    }
                    else
                    {
                        return;
                    }
                }

                sErrorLocation = "Capturing Cube property in connection string";
                string sCubeInConnectionString = null;
                try
                {
                    sCubeInConnectionString = Convert.ToString(connADO.Properties["Cube"].Value);
                }
                catch { }

                if (!string.IsNullOrEmpty(sCubeInConnectionString))
                {
                    if (string.Compare(sCubeInConnectionString, Convert.ToString(cache.WorkbookConnection.OLEDBConnection.CommandText), true) != 0)
                    {
                        throw new Exception("The connection string contains Cube=" + sCubeInConnectionString + " but the command text is " + Convert.ToString(cache.WorkbookConnection.OLEDBConnection.CommandText));
                    }
                }

                //find the last measure in the PivotTable. This will be the field we remove then add back to cause the PivotTable to requery the cube without calling Refresh which recreates the connection
                sErrorLocation = "Finding last measure in PivotTable";
                int iMaxPos = -1;
                Excel.CubeField fieldMeasure = null;
                Excel.CubeField fieldFallbackMeasure = null;
                foreach (Excel.CubeField field in pvt.CubeFields)
                {
                    if (field.Orientation == Excel.XlPivotFieldOrientation.xlDataField)
                    {
                        if (field.Position > iMaxPos)
                        {
                            iMaxPos = field.Position;
                            fieldMeasure = field;
                        }
                    }
                    else if (fieldFallbackMeasure == null && field.DragToData)
                    {
                        fieldFallbackMeasure = field;
                    }
                    else if (field.Orientation == Excel.XlPivotFieldOrientation.xlColumnField || field.Orientation == Excel.XlPivotFieldOrientation.xlRowField)
                    {
                        try
                        {
                            field.CreatePivotFields();
                        }
                        catch { }

                        //accumulate the visible items from all levels in case some items from multiple levels are checked
                        foreach (Excel.PivotField pf in field.PivotFields)
                        {
                            if (pf.IsMemberProperty) continue;

                        }
                    }
                }

                System.Collections.Generic.Dictionary<Excel.PivotField, int> dictPivotFieldSortOrder = new System.Collections.Generic.Dictionary<Excel.PivotField, int>();
                if (fieldMeasure != null)
                {
                    sErrorLocation = "Finding sort order for pivot fields";
                    foreach (Excel.CubeField field in pvt.CubeFields)
                    {
                        if (field.Orientation == Excel.XlPivotFieldOrientation.xlColumnField || field.Orientation == Excel.XlPivotFieldOrientation.xlRowField)
                        {
                            try
                            {
                                field.CreatePivotFields();
                            }
                            catch { }

                            try
                            {
                                //accumulate the visible items from all levels in case some items from multiple levels are checked
                                foreach (Excel.PivotField pf in field.PivotFields)
                                {
                                    if (pf.IsMemberProperty) continue;
                                    if (pf.AutoSortField == fieldMeasure.Name)
                                    {
                                        //we are sorting by the field I'm about to remove... need to save the sort settings and recreate the sort
                                        if (pf.AutoSortPivotLine.LineType != Excel.XlPivotLineType.xlPivotLineGrandTotal)
                                        {
                                            MessageBox.Show("Field " + pf.Name + " is sorting by measure " + pf.AutoSortField + " but not by the grand total. It will be sorted by the grand total after clearing the PivotTable cache since the other sort data won't exist.");
                                        }
                                        dictPivotFieldSortOrder.Add(pf, pf.AutoSortOrder);
                                    }
                                }
                            }
                            catch { } //if it fails, oh well... we'll only lose sorting
                        }
                    }
                }

                if (string.IsNullOrEmpty(sCubeInConnectionString))
                {
                    sErrorLocation = "Setting connection string";
                    cache.WorkbookConnection.OLEDBConnection.Connection = "OLEDB;" + sConnectionString + ";Cube=" + cache.WorkbookConnection.OLEDBConnection.CommandText;
                    sErrorLocation = "Setting OLEDBConnection.AlwaysUseConnectionFile";
                    cache.WorkbookConnection.OLEDBConnection.AlwaysUseConnectionFile = false;
                    cache.MakeConnection();

                    connADO = pvt.PivotCache().ADOConnection as ADODB.Connection;
                    if (connADO == null) throw new Exception("Could not cast PivotCache.ADOConnection to ADODB.Connection.");
                }

                try
                {
                    sErrorLocation = "Setting AllMembers = null";
                    object iRecords = null;

                    connADO.Execute("[Measures].AllMembers = null;", out iRecords, (int)ADODB.CommandTypeEnum.adCmdText);

                    //remove and re-add a measure to cause the PivotTable to requery the nulled out cube
                    if (fieldMeasure != null) // && iDataFieldCount > 1)
                    {
                        string sMeasureName = fieldMeasure.Name;
                        fieldMeasure.Orientation = Excel.XlPivotFieldOrientation.xlHidden;

                        fieldMeasure = pvt.CubeFields[sMeasureName];
                        fieldMeasure.Orientation = Excel.XlPivotFieldOrientation.xlDataField;

                        //restore the sort order as best we can
                        foreach (Excel.PivotField pf in dictPivotFieldSortOrder.Keys)
                        {
                            pf.AutoSort(dictPivotFieldSortOrder[pf], fieldMeasure.Name);
                        }
                    }
                    else
                    {
                        fieldFallbackMeasure.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                        fieldFallbackMeasure.Orientation = Excel.XlPivotFieldOrientation.xlHidden;
                    }

                }
                finally
                {
                    try
                    {
                        if (string.IsNullOrEmpty(sCubeInConnectionString))
                        {
                            cache.WorkbookConnection.OLEDBConnection.Connection = "OLEDB;" + sConnectionString;
                            cache.WorkbookConnection.OLEDBConnection.SourceConnectionFile = sConnectionFile;
                            cache.WorkbookConnection.OLEDBConnection.AlwaysUseConnectionFile = bUseConnectionFile;
                        }
                    }
                    catch { }
                }

            }
            catch (Exception ex)
            {
                string sDebugObjectInfo = "";
                try
                {
                    sDebugObjectInfo += GetPropertiesFromObject(typeof(Excel.OLEDBConnection), Application.ActiveCell.PivotTable.PivotCache().WorkbookConnection.OLEDBConnection);
                }
                catch { }

                try
                {
                    sDebugObjectInfo += GetPropertiesFromObject(typeof(Excel.PivotTable), Application.ActiveCell.PivotTable);
                }
                catch { }

                try
                {
                    sDebugObjectInfo += GetPropertiesFromObject(typeof(Excel.PivotCache), Application.ActiveCell.PivotTable.PivotCache());
                }
                catch { }

                MessageBox.Show("Problem during Clear PivotTable Cache:\r\n" + ex.Message + "\r\n" + ex.StackTrace + "\r\n\r\nAt task: " + sErrorLocation + "\r\n" + sDebugObjectInfo, "OLAP PivotTable Extensions");
            }
            finally
            {
                MainForm.ResetCulture(Application);
            }
        }

        private static string GetPropertiesFromObject(Type t, object o)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine().AppendLine().Append(t.FullName).AppendLine(" properties:");
            try
            {
                foreach (var prop in t.GetProperties())
                {
                    try
                    {
                        sb.AppendFormat("{0}={1}", prop.Name, prop.GetValue(o, null)).AppendLine();
                    }
                    catch { }
                }
            }
            catch { }
            return sb.ToString();
        }

        public static bool IsOlapPivotTable(Excel.PivotTable pvt)
        {
            try
            {
                if (pvt == null)
                    return false;
                Excel.PivotCache cache = pvt.PivotCache();
                return cache.OLAP
                    && cache.WorkbookConnection != null; //catches the situation when the connection for a PivotTable has been deleted
            }
            catch
            {
                return false;
            }
        }

        public static bool IsOledbConnection(Excel.PivotTable pvt)
        {
            try
            {
                if (pvt == null)
                    return false;
                Excel.PivotCache cache = pvt.PivotCache();
                return (cache.WorkbookConnection.Type == Excel.XlConnectionType.xlConnectionTypeOLEDB);
            }
            catch
            {
                return false;
            }
        }

        public static string GetOlapPivotTableHierarchy(Excel.PivotCell cell)
        {
            try
            {
                if (IsOlapPivotTable(cell.PivotTable))
                {
                    if (cell.PivotCellType == Excel.XlPivotCellType.xlPivotCellPageFieldItem || cell.PivotCellType == Excel.XlPivotCellType.xlPivotCellPivotField || cell.PivotCellType == Excel.XlPivotCellType.xlPivotCellPivotItem)
                    {
                        Excel.CubeField field = cell.PivotField.CubeField;
                        if (field.CubeFieldType == Excel.XlCubeFieldType.xlHierarchy) //not named sets since you can't filter them
                        {
                            return field.Name;
                        }
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        void m_xlAppEvents_xlSheetBeforeRightClick(object Sh, Microsoft.Office.Interop.Excel.Range Target, ref bool Cancel)
        {
            try
            {
                MainForm.SetCulture(Application);

                if (IsSingleDocumentInterface) //if it's Excel 2013, then delete and readd the menu items or else they will not properly show up and work in any window but the first document opened
                {
                    CreateOlapPivotTableExtensionsMenu();
                }

                if (IsOlapPivotTable(Application.ActiveCell.PivotTable))
                {
                    cmdMenuItem.Visible = true;
                    string sSelectedHierarchy = GetOlapPivotTableHierarchy(Application.ActiveCell.PivotCell);
                    cmdSearchMenuItem.Visible = !string.IsNullOrEmpty(sSelectedHierarchy);
                    cmdFilterListMenuItem.Visible = !string.IsNullOrEmpty(sSelectedHierarchy);
                    cmdChooseFieldsMenuItem.Visible = !string.IsNullOrEmpty(sSelectedHierarchy);
                    cmdClearPivotTableCacheMenuItem.Visible = IsOledbConnection(Application.ActiveCell.PivotTable);
                    SetupShowPropertyAsCaption();
                    SetupShowErrorMenu(Target);
                    SetupShowDisableAutoRefreshMenu();
                }
                else
                {
                    cmdMenuItem.Visible = false;
                    cmdSearchMenuItem.Visible = false;
                    cmdFilterListMenuItem.Visible = false;
                    cmdChooseFieldsMenuItem.Visible = false;
                    cmdShowPropertyAsCaptionMenuItem.Visible = false;
                    cmdClearPivotTableCacheMenuItem.Visible = false;
                    cmdErrorMenuItem.Visible = false;
                    cmdDisableAutoRefresh.Visible = false;
                }
            }
            catch
            {
                cmdMenuItem.Visible = true;
            }
            finally
            {
                MainForm.ResetCulture(Application);
            }
        }

        void SetupShowErrorMenu(Microsoft.Office.Interop.Excel.Range Target)
        {
            try
            {
                MainForm.SetCulture(Application);

                if (Target.Cells.Count == 1 && Convert.ToString(Target.Cells.Text) == "#VALUE!" && Target.PivotCell.PivotCellType == Excel.XlPivotCellType.xlPivotCellValue)
                {
                    cmdErrorMenuItem.Visible = true;
                }
                else
                {
                    cmdErrorMenuItem.Visible = false;
                }
            }
            catch
            {
                cmdErrorMenuItem.Visible = false;
            }
            finally
            {
                MainForm.ResetCulture(Application);
            }
        }

        void SetupShowDisableAutoRefreshMenu()
        {
            try
            {
                MainForm.SetCulture(Application);

                Excel.PivotCache pc = Application.ActiveCell.PivotTable.PivotCache();
                if (!PivotCacheIsDataModel(pc))
                {
                    cmdDisableAutoRefresh.Visible = false;
                }
                else if (pc.EnableRefresh)
                {
                    cmdDisableAutoRefresh.Caption = "Disable Auto Refresh";
                    cmdDisableAutoRefresh.Visible = true;
                    cmdDisableAutoRefresh.FaceId = 1919;
                }
                else
                {
                    cmdDisableAutoRefresh.Caption = "Enable Auto Refresh";
                    cmdDisableAutoRefresh.Visible = true;
                    cmdDisableAutoRefresh.FaceId = 1759;
                }
            }
            catch
            {
                cmdDisableAutoRefresh.Visible = false;
            }
            finally
            {
                MainForm.ResetCulture(Application);
            }
        }

        void SetupShowPropertyAsCaption()
        {
            try
            {
                if (Application.ActiveCell.PivotCell.PivotCellType == Excel.XlPivotCellType.xlPivotCellPivotItem
                    && !Application.ActiveCell.PivotCell.PivotField.IsMemberProperty)
                {
                    cmdShowPropertyAsCaptionMenuItem.Visible = true;
                    foreach (Office.CommandBarButton btn in cmdShowPropertyAsCaptionMenuItem.Controls)
                    {
                        btn.Delete(System.Reflection.Missing.Value);
                    }

                    bool bAddSeparator = false;
                    if (Application.ActiveCell.PivotCell.PivotField.UseMemberPropertyAsCaption)
                    {
                        Office.CommandBarButton btnStopUsingMemberPropertyAsCaption = (Office.CommandBarButton)cmdShowPropertyAsCaptionMenuItem.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);
                        btnStopUsingMemberPropertyAsCaption.Caption = "Reset Caption";
                        btnStopUsingMemberPropertyAsCaption.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnStopUsingMemberPropertyAsCaption_Click);

                        bAddSeparator = true;
                    }

                    bool bHasProperties = false;
                    foreach (Excel.PivotField memberProperty in Application.ActiveCell.PivotCell.PivotField.CubeField.PivotFields)
                    {
                        if (memberProperty.IsMemberProperty && memberProperty.Name.StartsWith(Application.ActiveCell.PivotCell.PivotField.Name))
                        {
                            Office.CommandBarButton btn = (Office.CommandBarButton)cmdShowPropertyAsCaptionMenuItem.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);
                            btn.Caption = memberProperty.Caption;
                            btn.Parameter = memberProperty.Name;
                            btn.BeginGroup = bAddSeparator;
                            bAddSeparator = false;
                            btn.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnShowPropertyAsCaption_Click);
                            bHasProperties = true;
                        }
                    }

                    if (!bHasProperties)
                    {
                        Office.CommandBarButton btn = (Office.CommandBarButton)cmdShowPropertyAsCaptionMenuItem.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);
                        btn.Caption = "(No Properties Retrieved)";
                        btn.Enabled = false;
                    }
                }
                else
                {
                    cmdShowPropertyAsCaptionMenuItem.Visible = false;
                }
            }
            catch
            {
                //swallow this error since if you have Defer Layout Update checked and right click on a hierarchy, it will give an error... but there's no way to detect that Defer Layout Update (PivotTable.ManualUpdate) is checked?!???
            }
        }

        void btnShowPropertyAsCaption_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                MainForm.SetCulture(Application);

                Application.ActiveCell.PivotCell.PivotField.MemberPropertyCaption = Ctrl.Parameter;
                Application.ActiveCell.PivotCell.PivotField.UseMemberPropertyAsCaption = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem using member property as caption:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
            finally
            {
                MainForm.ResetCulture(Application);
            }
        }

        void btnStopUsingMemberPropertyAsCaption_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                MainForm.SetCulture(Application);
                Application.ActiveCell.PivotCell.PivotField.UseMemberPropertyAsCaption = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem using resetting caption:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
            finally
            {
                MainForm.ResetCulture(Application);
            }
        }

        void m_xlAppEvents_xlSheetPivotTableUpdate(object Sh, Microsoft.Office.Interop.Excel.PivotTable Target)
        {
            if (frm != null && frm.AddInWorking) return; //short circuit if we're in the middle of changing the PivotTable with the add-in

            try
            {
                MainForm.SetCulture(Application);

                if (!IsOlapPivotTable(Target)) return;

                foreach (Excel.CubeField field in Target.CubeFields)
                {
                    if (field.Orientation != Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlHidden)
                    {
                        //this PivotTable isn't blank
                        return;
                    }
                }

                if (ShowCalcMembersByDefault && !Target.ViewCalculatedMembers)
                {
                    Target.ViewCalculatedMembers = true;
                }

                if (RefreshDataByDefault
                    && !Target.PivotCache().RefreshOnFileOpen
                    && IsOledbConnection(Target)) //don't cause the Excel data model pivots to refresh since that will reconnect to SQL
                {
                    Target.PivotCache().RefreshOnFileOpen = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during update of OLAP PivotTable:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
            finally
            {
                MainForm.ResetCulture(Application);
            }
        }

        void cmdMenuItem_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                if (Ctrl.Tag != cmdMenuItem.Tag || Ctrl.Caption != cmdMenuItem.Caption || Ctrl.FaceId != cmdMenuItem.FaceId)
                    return;

                frm = new MainForm(Application);
                frm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
        }

        void cmdSearchMenuItem_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                if (Ctrl.Tag != cmdSearchMenuItem.Tag || Ctrl.Caption != cmdSearchMenuItem.Caption || Ctrl.FaceId != cmdSearchMenuItem.FaceId)
                    return;

                frm = new MainForm(Application);
                string sSelectedHierarchy = GetOlapPivotTableHierarchy(Application.ActiveCell.PivotCell);
                frm.SetupSearchTab(sSelectedHierarchy);
                frm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
        }

        void cmdErrorMenuItem_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            //MainForm.SetCulture(Application); //don't need to call this here since it will be called in the MainForm constructor... but FormClosing won't be called since we never open the form... so we'll have to call it in the finally manually

            System.Text.StringBuilder sMdxQuery = new System.Text.StringBuilder();
            try
            {
                if (Ctrl.Tag != cmdErrorMenuItem.Tag || Ctrl.Caption != cmdErrorMenuItem.Caption || Ctrl.FaceId != cmdErrorMenuItem.FaceId)
                    return;

                frm = new MainForm(Application);
                bool bIsExcel2007OrHigherPivotTable = frm.IsExcel2007OrHigherPivotTableVersion();

                System.Text.StringBuilder sWhere = new System.Text.StringBuilder();
                System.Text.StringBuilder sSubselect = new System.Text.StringBuilder();

                //Excel 2010 only... much the easiest way for single-select filters... but for multi-select filters we'd have to parse the MDX... so go with the approach below which works on earlier versions of Excel
                //sMDX = (string)Application.ActiveCell.PivotCell.GetType().InvokeMember("MDX", System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Instance, null, Application.ActiveCell.PivotCell, null); //PivotCell.MDX is an Excel 2010 feature

                //get slicer filters
                System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>> dictFilters = GetSlicerFilters(Application.ActiveCell.PivotTable);

                //accumulate the row/column items. If you are drilled down through a hierarchy, both the top level and next level will appear as ColumnItems, so we want the most granular
                System.Collections.Generic.Dictionary<string, string> dictAxisItems = new System.Collections.Generic.Dictionary<string, string>();
                foreach (Excel.PivotItem pi in Application.ActiveCell.PivotCell.ColumnItems)
                {
                    if (dictAxisItems.ContainsKey(pi.Parent.CubeField.Name))
                        dictAxisItems[pi.Parent.CubeField.Name] = pi.Value;
                    else
                        dictAxisItems.Add(pi.Parent.CubeField.Name, pi.Value);
                }

                foreach (Excel.PivotItem pi in Application.ActiveCell.PivotCell.RowItems)
                {
                    if (dictAxisItems.ContainsKey(pi.Parent.CubeField.Name))
                        dictAxisItems[pi.Parent.CubeField.Name] = pi.Value;
                    else
                        dictAxisItems.Add(pi.Parent.CubeField.Name, pi.Value);
                }

                foreach (string sKey in dictAxisItems.Keys)
                {
                    string sValue = dictAxisItems[sKey];
                    if (dictFilters.ContainsKey(sKey))
                    {
                        dictFilters.Remove(sKey);
                    }
                    if (sWhere.Length > 0) sWhere.Append(",");
                    sWhere.Append(sValue).AppendLine();
                }

                //find all PivotTable filter fields
                foreach (Excel.CubeField cf in Application.ActiveCell.PivotTable.CubeFields)
                {
                    if (cf.Orientation == Excel.XlPivotFieldOrientation.xlPageField)
                    {
                        try
                        {
                            cf.CreatePivotFields();
                        }
                        catch { }

                        if (dictFilters.ContainsKey(cf.Name))
                        {
                            dictFilters.Remove(cf.Name);
                        }

                        System.Collections.Generic.List<string> listVisibleItems = new System.Collections.Generic.List<string>();
                        if (!cf.EnableMultiplePageItems)
                        {
                            listVisibleItems.Add(cf.CurrentPageName);
                        }
                        else
                        {
                            //accumulate the visible items from all levels in case some items from multiple levels are checked
                            foreach (Excel.PivotField pf in cf.PivotFields)
                            {
                                if (pf.IsMemberProperty) continue;
                                System.Array arrVisibleItems;
                                if (bIsExcel2007OrHigherPivotTable)
                                    arrVisibleItems = (System.Array)(object)pf.VisibleItemsList; //new to Excel 2007, so use CurrentPageList instead for older version PivotTables?
                                else
                                    arrVisibleItems = (System.Array)(object)pf.CurrentPageList;

                                foreach (string s in arrVisibleItems)
                                {
                                    if (string.IsNullOrEmpty(s)) continue;
                                    listVisibleItems.Add(s);
                                }
                            }
                        }

                        dictFilters.Add(cf.Name, listVisibleItems);
                    }
                }

                //add the filters to the where or subselect
                foreach (System.Collections.Generic.List<string> listVisibleItems in dictFilters.Values)
                {
                    if (listVisibleItems.Count == 1)
                    {
                        foreach (string s in listVisibleItems)
                        {
                            if (sWhere.Length > 0) sWhere.Append(",");
                            sWhere.Append(s).AppendLine();
                        }
                    }
                    else
                    {
                        if (sSubselect.Length > 0) sSubselect.Append("*");
                        sSubselect.Append(" {");
                        for (int i = 0; i < listVisibleItems.Count; i++)
                        {
                            if (i > 0) sSubselect.Append(",");
                            sSubselect.Append(listVisibleItems[i]);
                        }
                        sSubselect.Append("}").AppendLine();
                    }
                }


                //get the current measure
                if (sWhere.Length > 0) sWhere.Append(",");
                sWhere.Append(Application.ActiveCell.PivotCell.DataField.Value).AppendLine();



                frm.ConnectAdomdClientCube();
                AdomdClientWrappers.AdomdCommand cmd = new AdomdClientWrappers.AdomdCommand();
                cmd.Connection = frm.connCube;

                sMdxQuery.Append("select (" + sWhere.ToString() + ") on 0 from ");
                if (sSubselect.Length == 0)
                {
                    sMdxQuery.Append("[" + frm.cubeName + "] CELL PROPERTIES VALUE");
                }
                else
                {
                    sMdxQuery.Append("(select ").Append(sSubselect.ToString()).Append(" on 0 from [" + frm.cubeName + "]) CELL PROPERTIES VALUE");
                }

                frm.AddCalculatedMembersToMdxQuery(sMdxQuery);

                cmd.CommandText = sMdxQuery.ToString();

                AdomdClientWrappers.CellSet cellset = cmd.ExecuteCellSet();

                try
                {
                    object val = cellset.Cells[0].Value;
                }
                catch (Exception ex)
                {
                    if (ex.GetType().Name == "AdomdErrorResponseException")
                    {
                        MessageBox.Show("The error message behind #VALUE! is:\r\n\r\n" + ex.Message, "OLAP PivotTable Extensions");
                        return;
                    }
                    else
                    {
                        throw;
                    }
                }

                throw new Exception("Unable to reproduce the #VALUE! error for this cell with MDX query:\r\n\r\n" + sMdxQuery.ToString());
            }
            catch (Exception ex)
            {
                if (sMdxQuery.Length > 0)
                {
                    MessageBox.Show("Problem determining the error message behind this cell. MDX query for this cell was:\r\n\r\n" + sMdxQuery.ToString() + "\r\n\r\nError was: \r\n\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
                }
                else
                {
                    MessageBox.Show("Unable to determine the error message behind this cell. Error was: \r\n\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
                }
            }
            finally
            {
                try
                {
                    if (frm != null) frm.connCube.Close();
                }
                catch { }
                MainForm.ResetCulture(Application);
            }
        }

        System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>> GetSlicerFilters(Excel.PivotTable pivot)
        {
            System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>> dict = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>>();
            if (ExcelVersion >= 14)
            {
                //do all of this with InvokeMember since we're still using the Excel 2007 object model
                object oSlicers = pivot.GetType().InvokeMember("Slicers", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty, null, pivot, null);
                System.Collections.IEnumerable slicers = (System.Collections.IEnumerable)oSlicers;
                foreach (object oSlicer in slicers)
                {
                    object oSlicerCache = oSlicer.GetType().InvokeMember("SlicerCache", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty, null, oSlicer, null);
                    string sSourceName = (string)oSlicerCache.GetType().InvokeMember("SourceName", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty, null, oSlicerCache, null);
                    if (dict.ContainsKey(sSourceName)) continue; //if a hierarchy, the same SlicerCache will be seen multiple times
                    System.Array arrVisible = (System.Array)(object)oSlicerCache.GetType().InvokeMember("VisibleSlicerItemsList", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty, null, oSlicerCache, null);
                    System.Collections.Generic.List<string> list = new System.Collections.Generic.List<string>();
                    foreach (string sItem in arrVisible)
                    {
                        list.Add(sItem);
                    }
                    dict.Add(sSourceName, list);
                }
            }
            return dict;
        }

        private const string TEMP_MODEL_FLAT_FILE_CONNECTION_NAME = "OLAP PivotTable Extensions Temp Connection";
        private Excel.XlCalculation _OriginalCalculationMode = Excel.XlCalculation.xlCalculationAutomatic;

        ////doesn't work since it opens the Power Pivot DLLs in the wrong AppDomain
        //private void InitializeModel()
        //{
        //    try
        //    {
        //        MessageBox.Show("starting!");
        //        Application.ActiveWorkbook.Model.Initialize();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);
        //        throw ex;
        //    }
        //}

        void cmdDisableAutoRefresh_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            MainForm.SetCulture(Application); //don't need to call this here since it will be called in the MainForm constructor... but FormClosing won't be called since we never open the form... so we'll have to call it in the finally manually

            try
            {
                if (Ctrl.Tag != cmdDisableAutoRefresh.Tag || Ctrl.Caption != cmdDisableAutoRefresh.Caption || Ctrl.FaceId != this.cmdDisableAutoRefresh.FaceId)
                    return;

#if VSTO40
                //it appears adding a connection before the data model is loaded fails in Excel 2016
                //Application.CommandBars.ExecuteMso("DataModelManage"); //this opens the Power Pivot window in the wrong AppDomain
                //Application.ActiveWorkbook.Model.Initialize(); //this initializes the appropriate libraries in the wrong AppDomain
                //haven't figured out how to initialize Power Pivot in the proper AppDomain but did figure out how to detect it's not initialized yet:
                //mscoree.ICorRuntimeHost host = new mscoree.CorRuntimeHost();
                //object oDefaultAppDomain;
                //host.GetDefaultDomain(out oDefaultAppDomain);
                //AppDomain appD = (AppDomain)oDefaultAppDomain;

                ////now that we have the default AppDomain (where Power Pivot is supposed to be loaded, not the OlapPivotTableExtensions app domain) launch a new class which can loop through the assemblies loaded already in that default app domain
                //System.Runtime.Remoting.ObjectHandle handle = Activator.CreateInstanceFrom(appD, typeof(PowerPivotLaunchedChecker).Assembly.ManifestModule.FullyQualifiedName, typeof(PowerPivotLaunchedChecker).FullName);
                //PowerPivotLaunchedChecker newDomainInstance = (PowerPivotLaunchedChecker)handle.Unwrap();

                //check all app domains. it appears PowerPivot moved to a separate app domain
                //PowerPivotLaunchedChecker checker = new PowerPivotLaunchedChecker();
                bool bIsPowerPivotLoaded = false; // checker.IsPowerPivotLoaded;
                foreach (AppDomain appD in PowerPivotLaunchedChecker.GetProcessAppDomains())
                {
                    try
                    {
                        System.Runtime.Remoting.ObjectHandle handle = Activator.CreateInstanceFrom(appD, typeof(PowerPivotLaunchedChecker).Assembly.ManifestModule.FullyQualifiedName, typeof(PowerPivotLaunchedChecker).FullName);
                        PowerPivotLaunchedChecker newDomainInstance = (PowerPivotLaunchedChecker)handle.Unwrap();
                        bIsPowerPivotLoaded = newDomainInstance.IsPowerPivotLoaded;
                    }
                    catch { }
                    if (bIsPowerPivotLoaded) break;
                }
                if (!bIsPowerPivotLoaded)
                {
                    MessageBox.Show("First open the Power Pivot window. If you continue to get this message, restart Excel and try again.", "OLAP PivotTable Extensions");
                    return;
                }
#endif

                bool bEnableRefresh = Application.ActiveCell.PivotTable.PivotCache().EnableRefresh;
                Excel.WorkbookConnection connTemp = null;
                if (!bEnableRefresh)
                {
                    //if we are about to re-enable refresh, make a quick model change (adding a simple flat file connection which we will delete in a second... deleting it will cause the pivots to refresh)
                    string sTempDir = System.IO.Path.GetTempPath();
                    string sPath = sTempDir + @"\" + TEMP_MODEL_FLAT_FILE_CONNECTION_NAME + ".txt";
                    System.IO.File.WriteAllText(sPath, "col1\r\n1"); //just some sample contents to load
                    Excel.Connections conns = Application.ActiveWorkbook.Connections;
                    foreach (Excel.WorkbookConnection conn in conns)
                    {
                        if (conn.Name == TEMP_MODEL_FLAT_FILE_CONNECTION_NAME)
                        {
                            connTemp = conn;
                            break;
                        }
                    }
                    if (connTemp == null)
                    {
                        int iExcelVersionNumber = (int)decimal.Parse(Application.Version, System.Globalization.CultureInfo.InvariantCulture.NumberFormat);
                        //because we're using the Excel 2007 object model, use reflection to call this Excel 2013 method
                        //conns.Add2 TEMP_MODEL_FLAT_FILE_CONNECTION_NAME,"","OLEDB;Provider=Microsoft.ACE.OLEDB.15.0;Data Source=C:\Users\ggalloway\AppData\Local\Temp\;Persist Security Info=false;Extended Properties=""Text;HDR=Yes;FMT=CSVDelimited"";", "OLAP PivotTable Extensions Temp Connection#txt", xlCmdTable, True, False
                        string sConnectionString = "OLEDB;Provider=Microsoft.ACE.OLEDB.15.0;Data Source=" + sTempDir + ";Persist Security Info=false;Extended Properties=\"Text;HDR=Yes;FMT=CSVDelimited\";";
                        if (iExcelVersionNumber >= 16)
                        {
                            sConnectionString = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sTempDir + ";Persist Security Info=false;Extended Properties=\"Text;HDR=Yes;FMT=CSVDelimited\";Jet OLEDB:Registry Path=Software\\Microsoft\\Office\\16.0\\PowerPivot\\ACE\\";
                        }
                        //MessageBox.Show(sConnectionString); //TODO
                        connTemp = (Excel.WorkbookConnection)conns.GetType().InvokeMember("Add2", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.InvokeMethod, null, conns, new object[] { TEMP_MODEL_FLAT_FILE_CONNECTION_NAME, "This is a temporary connection used by OLAP PivotTable Extensions to trigger a quick refresh of the PivotTable field list. Feel free to delete.", sConnectionString, "OLAP PivotTable Extensions Temp Connection#txt", Excel.XlCmdType.xlCmdTable, true, false });
                    }
                }

                //enable/disable PivotCaches
                foreach (Excel.PivotCache pc in Application.ActiveWorkbook.PivotCaches())
                {
                    if (PivotCacheIsDataModel(pc))
                    {
                        if (bEnableRefresh)
                        {
                            pc.EnableRefresh = false;
                        }
                        else
                        {
                            pc.EnableRefresh = true;
                        }
                    }
                }

                //enable/disable DAX query tables
                foreach (Excel.WorkbookConnection conn in Application.ActiveWorkbook.Connections)
                {
                    if ((int)conn.Type == MainForm.xlConnectionTypeMODEL)
                    {
                        try
                        {
                            conn.OLEDBConnection.EnableRefresh = !bEnableRefresh; //this statement will fail for the ThisWorkbookDataModel connection but will succeed for DAX query tables... if we want to avoid this error in the future, we may have to check whether ModelConnection.CommandType = xlCmdCube (which means it's ThisWorkbookDataModel) or ModelConnection.CommandType = xlCmdDAX (or maybe xlCmdTable, too?) which means it's a DAX query table
                        }
                        catch { }
                    }
                }

                //set the calculation mode to manual when disabling auto refresh so that CUBEVALUE formulas don't refresh
                if (bEnableRefresh)
                {
                    //save the current calculation mode before setting it to manual
                    _OriginalCalculationMode = Application.Calculation;
                }
                Application.Calculation = (!bEnableRefresh ? Excel.XlCalculation.xlCalculationAutomatic : Excel.XlCalculation.xlCalculationManual);

                if (!bEnableRefresh)
                {
                    //Power Pivot window, if open, would throw an unhandled exception unless I paused for it to catch up with the new temporary flat file
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(1000);
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(1000);
                    System.Windows.Forms.Application.DoEvents();

                    connTemp.Delete(); //delete the temporary flat file connection to trigger a refresh of the field list in the PivotTables without refreshing the SQL data sources
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem enabling or disabling auto refresh. Error was: \r\n\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
            finally
            {
                MainForm.ResetCulture(Application);
            }
        }

        private bool PivotCacheIsDataModel(Excel.PivotCache pc)
        {
            return pc.OLAP && pc.WorkbookConnection != null && (int)pc.WorkbookConnection.Type == MainForm.xlConnectionTypeMODEL;
        }

        void cmdFilterListMenuItem_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                if (Ctrl.Tag != cmdFilterListMenuItem.Tag || Ctrl.Caption != cmdFilterListMenuItem.Caption || Ctrl.FaceId != cmdFilterListMenuItem.FaceId)
                    return;

                frm = new MainForm(Application);
                string sSelectedHierarchy = GetOlapPivotTableHierarchy(Application.ActiveCell.PivotCell);
                frm.SetupFilterListTab(sSelectedHierarchy);
                frm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
        }

        private void DeleteOlapPivotTableExtensionsMenu()
        {
            try
            {
                if (cmdSearchMenuItem != null) cmdSearchMenuItem.Click -= new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdSearchMenuItem_Click);
                if (cmdClearPivotTableCacheMenuItem != null) cmdClearPivotTableCacheMenuItem.Click -= new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdClearPivotTableCacheMenuItem_Click);
                if (cmdFilterListMenuItem != null) cmdFilterListMenuItem.Click -= new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdFilterListMenuItem_Click);
                if (cmdMenuItem != null) cmdMenuItem.Click -= new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdMenuItem_Click);
                if (cmdErrorMenuItem != null) cmdErrorMenuItem.Click -= new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdErrorMenuItem_Click);
                if (cmdDisableAutoRefresh != null) cmdDisableAutoRefresh.Click -= new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdDisableAutoRefresh_Click);
                if (cmdChooseFieldsMenuItem != null) cmdChooseFieldsMenuItem.Click -= new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdChooseFieldsMenuItem_Click);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error on unhook events: " + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }

            try
            {
                if (Application.CommandBars == null) return;

                Office.CommandBar ptcon = Application.CommandBars[PIVOTTABLE_CONTEXT_MENU];
                foreach (Office.CommandBarControl btn in ptcon.Controls)
                {
                    if (btn is Office.CommandBarPopup)
                    {
                        try
                        {
                            foreach (Office.CommandBarControl btn2 in ((Office.CommandBarPopup)btn).Controls)
                            {
                                if (btn2.Tag == MENU_TAG || btn.Tag == MENU_TAG)
                                {
                                    btn2.Delete(System.Reflection.Missing.Value);
                                }
                            }
                        }
                        catch { }
                    }
                    if (btn.Tag == MENU_TAG)
                    {
                        btn.Delete(System.Reflection.Missing.Value);
                    }
                }
            }
            catch { }
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            DeleteOlapPivotTableExtensionsMenu();
        }

#region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
#endregion
    }
}
