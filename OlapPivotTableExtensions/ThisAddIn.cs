using System;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace OlapPivotTableExtensions
{
    public partial class ThisAddIn
    {
        private const string REGISTRY_BASE_PATH = "SOFTWARE\\OLAP PivotTable Extensions";
        private const string REGISTRY_PATH_SHOW_CALC_MEMBERS_BY_DEFAULT = "ShowCalcMembersByDefault";
        private const string REGISTRY_PATH_REFRESH_DATA_BY_DEFAULT = "RefreshDataByDefault";

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
                if (Ctrl.InstanceId != cmdMenuItem.InstanceId)
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
                if (Ctrl.InstanceId != cmdSearchMenuItem.InstanceId)
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
                if (Ctrl.InstanceId != cmdFilterListMenuItem.InstanceId)
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
        
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                #region VSTO generated code

                this.Application = (Excel.Application)Microsoft.Office.Tools.Excel.ExcelLocale1033Proxy.Wrap(typeof(Excel.Application), this.Application);

                #endregion

                CreateOlapPivotTableExtensionsMenu();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during startup of OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
            }
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
