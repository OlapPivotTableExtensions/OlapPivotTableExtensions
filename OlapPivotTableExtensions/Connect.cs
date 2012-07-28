using System;
using Extensibility;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace OlapPivotTableExtensions
{
    /// <summary>
	///   The object for implementing an Add-in.
	/// </summary>
	/// <seealso class='IDTExtensibility2' />
	[GuidAttribute("DD16A145-E2F0-40B9-9993-5018BA8B6FF3"), ProgId("OlapPivotTableExtensions.Connect")]
	public class Connect : Object, Extensibility.IDTExtensibility2
	{
		/// <summary>
		///		Implements the constructor for the Add-in object.
		///		Place your initialization code within this method.
		/// </summary>
		public Connect()
		{
		}

		/// <summary>
		///      Implements the OnConnection method of the IDTExtensibility2 interface.
		///      Receives notification that the Add-in is being loaded.
		/// </summary>
		/// <param term='application'>
		///      Root object of the host application.
		/// </param>
		/// <param term='connectMode'>
		///      Describes how the Add-in is being loaded.
		/// </param>
		/// <param term='addInInst'>
		///      Object representing this Add-in.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
		{
            try
            {
                Application = (Excel.Application)application;
                addInInstance = addInInst;

                OriginalLanguage = System.Threading.Thread.CurrentThread.CurrentCulture.EnglishName;
                OriginalCultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;

                m_xlAppEvents = new xlEvents();
                m_xlAppEvents.DisableEventsIfEmbedded = true;
                m_xlAppEvents.SetupConnection(Application);

                //the following code works around an issue that is surfaced by typical event handling
                //http://olappivottableextend.codeplex.com/discussions/271174
                //typical event handling: Application.SheetBeforeRightClick += new Microsoft.Office.Interop.Excel.AppEvents_SheetBeforeRightClickEventHandler(Application_SheetBeforeRightClick);
                m_xlAppEvents.xlSheetBeforeRightClick += new xlEvents.DSheetBeforeRightClick(m_xlAppEvents_xlSheetBeforeRightClick);
                m_xlAppEvents.xlSheetPivotTableUpdate += new xlEvents.DSheetPivotTableUpdate(m_xlAppEvents_xlSheetPivotTableUpdate);

                ExcelVersion = (int)decimal.Parse(Application.Version, System.Globalization.CultureInfo.InvariantCulture.NumberFormat);
                if (ExcelVersion >= 15)
                {
                    IsSingleDocumentInterface = true;
                }

                if (IsSingleDocumentInterface)
                {
                    m_xlAppEvents.xlWindowActivate += new xlEvents.DWindowActivate(m_xlAppEvents_xlWindowActivate);
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
		public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
		{
            DeleteOlapPivotTableExtensionsMenu();
        }

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
		
		private Excel.Application Application;
		private object addInInstance;
        private xlEvents m_xlAppEvents;

        public static string OriginalLanguage = "";
        public static System.Globalization.CultureInfo OriginalCultureInfo;

        private int ExcelVersion;
        private bool IsSingleDocumentInterface;

        private const string REGISTRY_BASE_PATH = "SOFTWARE\\OLAP PivotTable Extensions";
        private const string REGISTRY_PATH_SHOW_CALC_MEMBERS_BY_DEFAULT = "ShowCalcMembersByDefault";
        private const string REGISTRY_PATH_REFRESH_DATA_BY_DEFAULT = "RefreshDataByDefault";
        private const string REGISTRY_PATH_FORMAT_MDX = "FormatMDX";
        private global::System.Object missing = global::System.Type.Missing;


        private const string MENU_TAG = "OLAP PivotTable Extensions";
        private const string PIVOTTABLE_CONTEXT_MENU = "PivotTable Context Menu";
        private Office.CommandBarButton cmdMenuItem = null;
        private Office.CommandBarButton cmdSearchMenuItem = null;
        private Office.CommandBarButton cmdFilterListMenuItem = null;
        private Office.CommandBarPopup cmdShowPropertyAsCaptionMenuItem = null;
        private Office.CommandBarButton cmdClearPivotTableCacheMenuItem = null;

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

        private void CreateOlapPivotTableExtensionsMenu()
        {
            try
            {
                DeleteOlapPivotTableExtensionsMenu();

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
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during startup of OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
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

                cache.WorkbookConnection.OLEDBConnection.MaintainConnection = true;
                if (!cache.IsConnected)
                    cache.MakeConnection();

                ADODB.Connection connADO = cache.ADOConnection as ADODB.Connection;
                if (connADO == null) throw new Exception("Could not cast PivotCache.ADOConnection to ADODB.Connection.");

                string sConnectionFile = cache.WorkbookConnection.OLEDBConnection.SourceConnectionFile;
                bool bUseConnectionFile = cache.WorkbookConnection.OLEDBConnection.AlwaysUseConnectionFile;
                string sConnectionString = connADO.ConnectionString;
                Excel.WorkbookConnection connOld = cache.WorkbookConnection;

                if (cache.WorkbookConnection.OLEDBConnection.CommandType != Excel.XlCmdType.xlCmdCube)
                    throw new Exception("Connection command type is not Cube. This functionality is not supported in this scenario.");

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
                }

                if (string.IsNullOrEmpty(sCubeInConnectionString))
                {
                    cache.WorkbookConnection.OLEDBConnection.Connection = "OLEDB;" + sConnectionString + ";Cube=" + cache.WorkbookConnection.OLEDBConnection.CommandText;
                    cache.WorkbookConnection.OLEDBConnection.AlwaysUseConnectionFile = false;
                    cache.MakeConnection();

                    connADO = pvt.PivotCache().ADOConnection as ADODB.Connection;
                    if (connADO == null) throw new Exception("Could not cast PivotCache.ADOConnection to ADODB.Connection.");
                }

                try
                {
                    object iRecords = null;

                    connADO.Execute("[Measures].AllMembers = null;", out iRecords, (int)ADODB.CommandTypeEnum.adCmdText);

                    //remove and re-add a measure to cause the PivotTable to requery the nulled out cube
                    if (fieldMeasure != null) // && iDataFieldCount > 1)
                    {
                        string sMeasureName = fieldMeasure.Name;
                        fieldMeasure.Orientation = Excel.XlPivotFieldOrientation.xlHidden;

                        fieldMeasure = pvt.CubeFields[sMeasureName];
                        fieldMeasure.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                    }
                    else
                    {
                        fieldFallbackMeasure.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                        fieldFallbackMeasure.Orientation = Excel.XlPivotFieldOrientation.xlHidden;
                    }

                }
                finally
                {
                    if (string.IsNullOrEmpty(sCubeInConnectionString))
                    {
                        cache.WorkbookConnection.OLEDBConnection.Connection = "OLEDB;" + sConnectionString;
                        cache.WorkbookConnection.OLEDBConnection.SourceConnectionFile = sConnectionFile;
                        cache.WorkbookConnection.OLEDBConnection.AlwaysUseConnectionFile = bUseConnectionFile;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during startup of OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
            finally
            {
                MainForm.ResetCulture(Application);
            }
        }

        public static bool IsOlapPivotTable(Excel.PivotTable pvt)
        {
            try
            {
                if (pvt == null)
                    return false;
                Excel.PivotCache cache = pvt.PivotCache();
                return cache.OLAP;
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
                    cmdClearPivotTableCacheMenuItem.Visible = IsOledbConnection(Application.ActiveCell.PivotTable);
                    SetupShowPropertyAsCaption();
                }
                else
                {
                    cmdMenuItem.Visible = false;
                    cmdSearchMenuItem.Visible = false;
                    cmdFilterListMenuItem.Visible = false;
                    cmdShowPropertyAsCaptionMenuItem.Visible = false;
                    cmdClearPivotTableCacheMenuItem.Visible = false;
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
            catch (Exception ex)
            {
                MessageBox.Show("Problem setting up Show Property as Caption:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
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
                if (Ctrl.Tag != cmdMenuItem.Tag || Ctrl.Caption!= cmdMenuItem.Caption || Ctrl.FaceId != cmdMenuItem.FaceId)
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


	}
}