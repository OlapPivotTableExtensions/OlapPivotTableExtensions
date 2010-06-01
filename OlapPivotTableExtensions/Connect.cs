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

                CreateOlapPivotTableExtensionsMenu();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during startup of OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
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
		}
		
		private Excel.Application Application;
		private object addInInstance;





        private const string REGISTRY_BASE_PATH = "SOFTWARE\\OLAP PivotTable Extensions";
        private const string REGISTRY_PATH_SHOW_CALC_MEMBERS_BY_DEFAULT = "ShowCalcMembersByDefault";
        private const string REGISTRY_PATH_REFRESH_DATA_BY_DEFAULT = "RefreshDataByDefault";
        private global::System.Object missing = global::System.Type.Missing;


        private const string MENU_TAG = "OLAP PivotTable Extensions";
        private const string PIVOTTABLE_CONTEXT_MENU = "PivotTable Context Menu";
        private Office.CommandBarButton cmdMenuItem = null;
        private Office.CommandBarButton cmdSearchMenuItem = null;
        private Office.CommandBarButton cmdFilterListMenuItem = null;

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

                Application.SheetBeforeRightClick += new Microsoft.Office.Interop.Excel.AppEvents_SheetBeforeRightClickEventHandler(Application_SheetBeforeRightClick);
                Application.SheetPivotTableUpdate += new Microsoft.Office.Interop.Excel.AppEvents_SheetPivotTableUpdateEventHandler(Application_SheetPivotTableUpdate);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during startup of OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
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

        void Application_SheetBeforeRightClick(object Sh, Microsoft.Office.Interop.Excel.Range Target, ref bool Cancel)
        {
            try
            {
                if (IsOlapPivotTable(Application.ActiveCell.PivotTable))
                {
                    cmdMenuItem.Visible = true;
                    string sSelectedHierarchy = GetOlapPivotTableHierarchy(Application.ActiveCell.PivotCell);
                    cmdSearchMenuItem.Visible = !string.IsNullOrEmpty(sSelectedHierarchy);
                    cmdFilterListMenuItem.Visible = !string.IsNullOrEmpty(sSelectedHierarchy);
                }
                else
                {
                    cmdMenuItem.Visible = false;
                    cmdSearchMenuItem.Visible = false;
                    cmdFilterListMenuItem.Visible = false;
                }
            }
            catch
            {
                cmdMenuItem.Visible = true;
            }
        }

        void Application_SheetPivotTableUpdate(object Sh, Microsoft.Office.Interop.Excel.PivotTable Target)
        {
            try
            {
                if (frm != null && frm.AddInWorking) return; //short circuit if we're in the middle of changing the PivotTable with the add-in
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

                if (RefreshDataByDefault && !Target.PivotCache().RefreshOnFileOpen)
                {
                    Target.PivotCache().RefreshOnFileOpen = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during update of OLAP PivotTable:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
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
                Office.CommandBar ptcon = Application.CommandBars[PIVOTTABLE_CONTEXT_MENU];
                foreach (Office.CommandBarControl btn in ptcon.Controls)
                {
                    if (btn.Tag == MENU_TAG)
                    {
                        btn.Delete(System.Reflection.Missing.Value);
                    }
                    if (btn is Office.CommandBarPopup)
                    {
                        foreach (Office.CommandBarControl btn2 in ((Office.CommandBarPopup)btn).Controls)
                        {
                            if (btn2.Tag == MENU_TAG)
                            {
                                btn2.Delete(System.Reflection.Missing.Value);
                            }
                        }
                    }
                }
            }
            catch { }
        }        

	}
}