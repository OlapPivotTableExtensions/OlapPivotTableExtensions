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

        private const string MENU_CAPTION = "OLAP PivotTable Extensions...";
        private Office.CommandBarButton cmdMenuItem = null;

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

        private void CreateOlapPivotTableExtensionsMenu()
        {
            try
            {
                DeleteOlapPivotTableExtensionsMenu();

                Office.CommandBar ptcon = Application.CommandBars["PivotTable Context Menu"];
                cmdMenuItem = (Office.CommandBarButton)ptcon.Controls.Add(Office.MsoControlType.msoControlButton, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true);
                cmdMenuItem.Caption = MENU_CAPTION;
                cmdMenuItem.FaceId = 1122;

                cmdMenuItem.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(cmdMenuItem_Click);

                Application.SheetBeforeRightClick += new Microsoft.Office.Interop.Excel.AppEvents_SheetBeforeRightClickEventHandler(Application_SheetBeforeRightClick);
                Application.SheetPivotTableUpdate += new Microsoft.Office.Interop.Excel.AppEvents_SheetPivotTableUpdateEventHandler(Application_SheetPivotTableUpdate);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem during startup of OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace);
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

        void Application_SheetBeforeRightClick(object Sh, Microsoft.Office.Interop.Excel.Range Target, ref bool Cancel)
        {
            try
            {
                if (IsOlapPivotTable(Application.ActiveCell.PivotTable))
                    cmdMenuItem.Visible = true;
                else
                    cmdMenuItem.Visible = false;
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void cmdMenuItem_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                MainForm frm = new MainForm(Application);
                frm.ShowDialog();
            }
            catch { }
        }

        private void DeleteOlapPivotTableExtensionsMenu()
        {
            try
            {
                Office.CommandBar ptcon = Application.CommandBars["PivotTable Context Menu"];
                foreach (Office.CommandBarControl btn in ptcon.Controls)
                {
                    if (btn.Caption == MENU_CAPTION)
                    {
                        btn.Delete(System.Reflection.Missing.Value);
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
                MessageBox.Show("Problem during startup of OLAP PivotTable Extensions:\r\n" + ex.Message + "\r\n" + ex.StackTrace);
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
