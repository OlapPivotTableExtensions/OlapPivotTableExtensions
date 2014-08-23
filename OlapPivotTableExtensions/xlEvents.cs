/***************************************************************************
 *  Copyright ©2008; Microsoft Corporation. All rights reserved.
 *  Written by Microsoft Office Developer Support
 * 
 *  This code is provided as a sample. It is not a formal
 *  product and has not been fully tested. Use it
 *  for educational purposes only.
 *
 *  THIS CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
 *  EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
 *  WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
 *
 ***************************************************************************/

/*
 * Note: This code is from http://blogs.msdn.com/b/vsofficedeveloper/archive/2008/04/11/excel-ole-embedding-errors-with-managed-addin.aspx
 * It is intended to workaround this issue: http://olappivottableextend.codeplex.com/discussions/271174
 */
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace OlapPivotTableExtensions
{
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
	GuidAttribute("00024413-0000-0000-C000-000000000046")]
	public interface DExcel12AppEvents
	{
        [DispId(0x0000061D)]
        void NewWorkbook(Excel._Workbook oWB);
        [DispId(0x00000616)]
        void SheetSelectionChange([MarshalAs(UnmanagedType.IDispatch)] object oSheet, Excel.Range oTarget);
        [DispId(0x00000617)]
        void SheetBeforeDoubleClick([MarshalAs(UnmanagedType.IDispatch)] object oSheet, Excel.Range oTarget, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x00000618)]
        void SheetBeforeRightClick([MarshalAs(UnmanagedType.IDispatch)] object oSheet, Excel.Range oTarget, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x00000619)]
        void SheetActivate([MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        [DispId(0x0000061A)]
        void SheetDeactivate([MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        [DispId(0x0000061B)]
        void SheetCalculate([MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        [DispId(0x0000061C)]
        void SheetChange([MarshalAs(UnmanagedType.IDispatch)] object oSheet, Excel.Range oTarget);
		[DispId(0x0000061F)] void WorkbookOpen(Excel._Workbook oWB);
        [DispId(0x00000620)]
        void WorkbookActivate(Excel._Workbook oWB);
        [DispId(0x00000621)]
        void WorkbookDeactivate(Excel._Workbook oWB);
        [DispId(0x00000622)]
        void WorkbookBeforeClose(Excel._Workbook oWB, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x00000623)]
        void WorkbookBeforeSave(Excel._Workbook oWB, [MarshalAs(UnmanagedType.VariantBool)]  bool SaveUI, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x00000624)]
        void WorkbookBeforePrint(Excel._Workbook oWB, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x00000625)]
        void WorkbookNewSheet(Excel._Workbook oWB, [MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        [DispId(0x00000626)]
        void WorkbookAddinInstall(Excel._Workbook oWB);
        [DispId(0x00000627)]
        void WorkbookAddinUninstall(Excel._Workbook oWB);
        [DispId(0x00000612)]
        void WindowResize(Excel._Workbook oWB, Excel.Window oWn);
        [DispId(0x00000614)]
        void WindowActivate(Excel._Workbook oWB, Excel.Window oWn);
        [DispId(0x00000615)]
        void WindowDeactivate(Excel._Workbook oWB, Excel.Window oWn);
        [DispId(0x0000073E)]
        void SheetFollowHyperlink([MarshalAs(UnmanagedType.IDispatch)] object oSheet, Excel.Hyperlink oTarget);
        [DispId(0x0000086D)]
        void SheetPivotTableUpdate([MarshalAs(UnmanagedType.IDispatch)] object oSheet, Excel.PivotTable oTarget);
        [DispId(0x00000870)]
        void WorkbookPivotTableCloseConnection(Excel._Workbook oWB, Excel.PivotTable oTarget);
        [DispId(0x00000871)]
        void WorkbookPivotTableOpenConnection(Excel._Workbook oWB, Excel.PivotTable oTarget);
        [DispId(0x000008F1)]
        void WorkbookSync(Excel._Workbook oWB, Office.MsoSyncEventType SyncType);
        [DispId(0x000008F2)]
        void WorkbookBeforeXmlImport(Excel._Workbook oWB, Excel.XmlMap oMap, string sUrl, [MarshalAs(UnmanagedType.VariantBool)]  bool IsRefresh, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x000008F3)]
        void WorkbookAfterXmlImport(Excel._Workbook oWB, Excel.XmlMap oMap, [MarshalAs(UnmanagedType.VariantBool)]  bool IsRefresh, Excel.XlXmlImportResult Result);
        [DispId(0x000008F4)]
        void WorkbookBeforeXmlExport(Excel._Workbook oWB, Excel.XmlMap oMap, string sUrl, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x000008F5)]
        void WorkbookAfterXmlExport(Excel._Workbook oWB, Excel.XmlMap oMap, string sUrl, Excel.XlXmlExportResult Result);
        [DispId(0x00000A33)]
        void WorkbookRowsetComplete(Excel._Workbook oWB, string sDesciption, string sSheet, [MarshalAs(UnmanagedType.VariantBool)]  bool Success);
        [DispId(0x00000A34)]
        void AfterCalculate();
	}

	
	public class xlEvents: IDisposable,DExcel12AppEvents
	{
		private IConnectionPoint m_oConnectionPoint;
		private int m_Cookie;
		private bool m_DisableEventsIfEmbedded;


        #region Events

        //NewWorkbook
        public delegate void DNewWorkbook(Excel._Workbook oWB);
        public event DNewWorkbook xlNewWorkbook;

        //SheetSelectionChange
        public delegate void DSheetSelectionChange([MarshalAs(UnmanagedType.IDispatch)] object oSheet, Excel.Range oTarget);
        public event DSheetSelectionChange xlSheetSelectionChange;

        //SheetBeforeDoubleClick
        public delegate void DSheetBeforeDoubleClick([MarshalAs(UnmanagedType.IDispatch)] object oSheet, Excel.Range oTarget, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        public event DSheetBeforeDoubleClick xlSheetBeforeDoubleClick;

        //SheetBeforeRightClick
        public delegate void DSheetBeforeRightClick([MarshalAs(UnmanagedType.IDispatch)] object oSheet, Excel.Range oTarget, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        public event DSheetBeforeRightClick xlSheetBeforeRightClick;
     
        //SheetActivate
        public delegate void DSheetActivate([MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        public event DSheetActivate xlSheetActivate;

        //SheetDeactivate
        public delegate void DSheetDeactivate([MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        public event DSheetDeactivate xlSheetDeactivate;

        //SheetCalculate
        public delegate void DSheetCalculate([MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        public event DSheetCalculate xlSheetCalculate;

        //SheetChange
        public delegate void DSheetChange([MarshalAs(UnmanagedType.IDispatch)] object oSheet, Excel.Range oTarget);
        public event DSheetChange xlSheetChange;

        //WorkbookOpen
        public delegate void DWorkbookOpen(Excel._Workbook oWB);
        public event DWorkbookOpen xlWorkbookOpen;

        //WorkbookActivate
        public delegate void DWorkbookActivate(Excel._Workbook oWB);
        public event DWorkbookActivate xlWorkbookActivate;

        //WorkbookDeactivate
        public delegate void DWorkbookDeactivate(Excel._Workbook oWB);
        public event DWorkbookDeactivate xlWorkbookDeactivate;

        //WorkbookBeforeClose
        public delegate void DWorkbookBeforeClose(Excel._Workbook oWB, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        public event DWorkbookBeforeClose xlWorkbookBeforeClose;

        //WorkbookBeforeSave
        public delegate void DWorkbookBeforeSave(Excel._Workbook oWB, [MarshalAs(UnmanagedType.VariantBool)]  bool SaveUI, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        public event DWorkbookBeforeSave xlWorkbookBeforeSave;

        //WorkbookBeforePrint
        public delegate void DWorkbookBeforePrint(Excel._Workbook oWB, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        public event DWorkbookBeforePrint xlWorkbookBeforePrint;

        //WorkbookNewSheet
        public delegate void DWorkbookNewSheet(Excel._Workbook oWB, [MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        public event DWorkbookNewSheet xlWorkbookNewSheet;

        //WorkbookAddinInstall
        public delegate void DWorkbookAddinInstall(Excel._Workbook oWB);
        public event DWorkbookAddinInstall xlWorkbookAddinInstall;

        //WorkbookAddinUninstall
        public delegate void DWorkbookAddinUninstall(Excel._Workbook oWB);
        public event DWorkbookAddinUninstall xlWorkbookAddinUninstall;

        //WindowResize
        public delegate void DWindowResize(Excel._Workbook oWB, Excel.Window oWn);
        public event DWindowResize xlWindowResize;

        //WindowActivate
        public delegate void DWindowActivate(Excel._Workbook oWB, Excel.Window oWn);
        public event DWindowActivate xlWindowActivate;

        //WindowDeactivate
        public delegate void DWindowDeactivate(Excel._Workbook oWB, Excel.Window oWn);
        public event DWindowDeactivate xlWindowDeactivate;

        //SheetFollowHyperlink
        public delegate void DSheetFollowHyperlink([MarshalAs(UnmanagedType.IDispatch)] object oSheet, Excel.Hyperlink oTarget);
        public event DSheetFollowHyperlink xlSheetFollowHyperlink;

        //SheetPivotTableUpdate
        public delegate void DSheetPivotTableUpdate([MarshalAs(UnmanagedType.IDispatch)] object oSheet,  Excel.PivotTable oTarget);
        public event DSheetPivotTableUpdate xlSheetPivotTableUpdate;

        //WorkbookPivotTableCloseConnection
        public delegate void DWorkbookPivotTableCloseConnection(Excel._Workbook oWB,  Excel.PivotTable oTarget);
        public event DWorkbookPivotTableCloseConnection xlWorkbookPivotTableCloseConnection;

        //WorkbookPivotTableOpenConnection
        public delegate void DWorkbookPivotTableOpenConnection(Excel._Workbook oWB,  Excel.PivotTable oTarget);
        public event DWorkbookPivotTableOpenConnection xlWorkbookPivotTableOpenConnection;

        //WorkbookSync
        public delegate void DWorkbookSync(Excel._Workbook oWB,  Office.MsoSyncEventType SyncType);
        public event DWorkbookSync xlWorkbookSync;

        //WorkbookBeforeXmlImport
        public delegate void DWorkbookBeforeXmlImport(Excel._Workbook oWB,  Excel.XmlMap oMap, string sUrl, [MarshalAs(UnmanagedType.VariantBool)]  bool IsRefresh, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        public event DWorkbookBeforeXmlImport xlWorkbookBeforeXmlImport;

        //WorkbookAfterXmlImport
        public delegate void DWorkbookAfterXmlImport(Excel._Workbook oWB,  Excel.XmlMap oMap, [MarshalAs(UnmanagedType.VariantBool)]  bool IsRefresh,  Excel.XlXmlImportResult Result);
        public event DWorkbookAfterXmlImport xlWorkbookAfterXmlImport;

        //WorkbookBeforeXmlExport
        public delegate void DWorkbookBeforeXmlExport(Excel._Workbook oWB, Excel.XmlMap oMap, string sUrl, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        public event DWorkbookBeforeXmlExport xlWorkbookBeforeXmlExport;

        //WorkbookAfterXmlExport
        public delegate void DWorkbookAfterXmlExport(Excel._Workbook oWB,  Excel.XmlMap oMap, string sUrl,  Excel.XlXmlExportResult Result);
        public event DWorkbookAfterXmlExport xlWorkbookAfterXmlExport;

        //WorkbookRowsetComplete
        public delegate void DWorkbookRowsetComplete(Excel._Workbook oWB,  string sDesciption, string sSheet, [MarshalAs(UnmanagedType.VariantBool)]  bool Success);
        public event DWorkbookRowsetComplete xlWorkbookRowsetComplete;

        //AfterCalculate
        public delegate void DAfterCalculate();
        public event DAfterCalculate xlAfterCalculate;

        #endregion

        #region Properties
        public bool DisableEventsIfEmbedded
		{
			get
			{
				return m_DisableEventsIfEmbedded;
			}
			set
			{
				m_DisableEventsIfEmbedded = value;
			}
		}
		#endregion

		public xlEvents()
		{
			m_oConnectionPoint = null;
			m_Cookie = 0;
            m_DisableEventsIfEmbedded = true;
		}

		public void SetupConnection(Excel.Application app)
		{
			if (m_Cookie != 0) return;

			//GUID of the DIID_ApplicationEvents dispinterface
			Guid guid = new Guid("00024413-0000-0000-C000-000000000046");

			//QI for IConnectionPointContainer
			IConnectionPointContainer oConnPointContainer = 
				(IConnectionPointContainer)app;

			//Find the connection point and then advise
			oConnPointContainer.FindConnectionPoint(ref guid, 
				out m_oConnectionPoint);
			m_oConnectionPoint.Advise(this, out m_Cookie);
		}

		public void RemoveConnection()
		{
			if (m_Cookie != 0)
			{
				m_oConnectionPoint.Unadvise(m_Cookie);
				ComRelease(m_oConnectionPoint);
				m_oConnectionPoint = null;
				m_Cookie = 0;
			}
		}

		public void ComRelease(object o)
		{
			try
			{
				Marshal.ReleaseComObject(o);
			}
			catch{}
			finally
			{
				o = null;
			}
		}

		public bool IsEmbedded(ref Excel._Workbook oWB)
		{
			try
			{
				string sName = oWB.Name;

				//Return true if we are editing the file inplace.
				if (oWB.IsInplace)
					return true;

				//Return true if the workbook name = object
				//or if the name contains the text "Workbook in" 
				//or "Chart in"
				if(sName=="Object")
					return true;

				if(sName.IndexOf("Workbook in") != -1 || 
					sName.IndexOf("Chart in") != -1)
					return true;

				//Return true if the Path is empty
				//Note that this may not be a valid
				//test if calling the method from an
				//event such as the NewWorkbook event
				//where the Path would be expected to be
				//an empty string
                //if(oWB.Path == "")
                //    return true;

				//Lastly check the Container property
				//of the workbook.  If the property is
				//not set this will return an error.
				object o = oWB.Container;
				if (o != null) 
				{
					ComRelease(o);
					return true;
				}
			}
			catch{}
			return false;
		}
	
        #region IDisposable Members

		public void Dispose()
		{
            RemoveConnection();
		}

		#endregion

		#region DExcel12AppEvents Members

        [DispId(0x0000061D)]
        public void NewWorkbook(Microsoft.Office.Interop.Excel._Workbook oWB)
        {
            if (xlNewWorkbook != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlNewWorkbook(oWB);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
        }

        [DispId(0x00000616)]
        public void SheetSelectionChange(object oSheet, Microsoft.Office.Interop.Excel.Range oTarget)
        {
            if (xlSheetSelectionChange != null)
                xlSheetSelectionChange(oSheet,oTarget);

            //Release any COM objects passed into the event
            ComRelease(oSheet);
            ComRelease(oTarget);
        }

        [DispId(0x00000617)]
        public void SheetBeforeDoubleClick(object oSheet, Microsoft.Office.Interop.Excel.Range oTarget, ref bool Cancel)
        {
            if (xlSheetBeforeDoubleClick != null)
                xlSheetBeforeDoubleClick(oSheet,oTarget,ref Cancel);

            //Release any COM objects passed into the event
            ComRelease(oSheet);
            ComRelease(oTarget);
        }

        [DispId(0x00000618)]
        public void SheetBeforeRightClick(object oSheet, Microsoft.Office.Interop.Excel.Range oTarget, ref bool Cancel)
        {
            if (xlSheetBeforeRightClick != null)
                xlSheetBeforeRightClick(oSheet,oTarget,ref Cancel);

            //Release any COM objects passed into the event
            ComRelease(oSheet);
            ComRelease(oTarget);
        }

        [DispId(0x00000619)]
        public void SheetActivate(object oSheet)
        {
            if (xlSheetActivate != null)
                xlSheetActivate(oSheet);

            //Release any COM objects passed into the event
            ComRelease(oSheet);
        }

        [DispId(0x0000061A)]
        public void SheetDeactivate(object oSheet)
        {
            if (xlSheetDeactivate != null)
                xlSheetDeactivate(oSheet);

            //Release any COM objects passed into the event
            ComRelease(oSheet);
        }

        [DispId(0x0000061B)]
        public void SheetCalculate(object oSheet)
        {
            if (xlSheetCalculate != null)
                xlSheetCalculate(oSheet);

            //Release any COM objects passed into the event
            ComRelease(oSheet);
        }

        [DispId(0x0000061C)]
        public void SheetChange(object oSheet, Microsoft.Office.Interop.Excel.Range oTarget)
        {
            if (xlSheetChange != null)
                xlSheetChange(oSheet,oTarget);

            //Release any COM objects passed into the event
            ComRelease(oSheet);
            ComRelease(oTarget);
        }

		[DispId(0x0000061F)]
		public void WorkbookOpen(Microsoft.Office.Interop.Excel._Workbook oWB)
		{
			
			if(xlWorkbookOpen != null)
			{
				if(!m_DisableEventsIfEmbedded ||
					(m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
				{
					xlWorkbookOpen(oWB);
				}
			}

			//Release any COM objects passed into the event
			ComRelease(oWB);
		}

        [DispId(0x00000620)]
        public void WorkbookActivate(Microsoft.Office.Interop.Excel._Workbook oWB)
        {
            if (xlWorkbookActivate != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookActivate(oWB);
                }
            }
            //Release any COM objects passed into the event
            ComRelease(oWB);
        }

        [DispId(0x00000621)]
        public void WorkbookDeactivate(Microsoft.Office.Interop.Excel._Workbook oWB)
        {
            if (xlWorkbookDeactivate != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookDeactivate(oWB);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
        }

        [DispId(0x00000622)]
        public void WorkbookBeforeClose(Microsoft.Office.Interop.Excel._Workbook oWB, ref bool Cancel)
        {
            if (xlWorkbookBeforeClose != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookBeforeClose(oWB,ref Cancel);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
        }

        [DispId(0x00000623)]
        public void WorkbookBeforeSave(Microsoft.Office.Interop.Excel._Workbook oWB, bool SaveUI, ref bool Cancel)
        {
            if (xlWorkbookBeforeSave != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookBeforeSave(oWB,SaveUI,ref Cancel);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
        }

        [DispId(0x00000624)]
        public void WorkbookBeforePrint(Microsoft.Office.Interop.Excel._Workbook oWB, ref bool Cancel)
        {
            if (xlWorkbookBeforePrint != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookBeforePrint(oWB,ref Cancel);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
        }

        [DispId(0x00000625)]
        public void WorkbookNewSheet(Microsoft.Office.Interop.Excel._Workbook oWB, object oSheet)
        {
            if (xlWorkbookNewSheet != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookNewSheet(oWB,oSheet);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
            ComRelease(oSheet);
        }

        [DispId(0x00000626)]
        public void WorkbookAddinInstall(Microsoft.Office.Interop.Excel._Workbook oWB)
        {
            if (xlWorkbookAddinInstall != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookAddinInstall(oWB);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
        }

        [DispId(0x00000627)]
        public void WorkbookAddinUninstall(Microsoft.Office.Interop.Excel._Workbook oWB)
        {
            if (xlWorkbookAddinUninstall != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookAddinUninstall(oWB);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
        }

        [DispId(0x00000612)]
        public void WindowResize(Microsoft.Office.Interop.Excel._Workbook oWB, Microsoft.Office.Interop.Excel.Window oWn)
        {
            if (xlWindowResize != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWindowResize(oWB,oWn);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
            ComRelease(oWn);
        }

        [DispId(0x00000614)]
        public void WindowActivate(Microsoft.Office.Interop.Excel._Workbook oWB, Microsoft.Office.Interop.Excel.Window oWn)
        {
            if (xlWindowActivate != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWindowActivate(oWB,oWn);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
            ComRelease(oWn);
        }

        [DispId(0x00000615)]
        public void WindowDeactivate(Microsoft.Office.Interop.Excel._Workbook oWB, Microsoft.Office.Interop.Excel.Window oWn)
        {
            if (xlWindowDeactivate != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWindowDeactivate(oWB,oWn);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
            ComRelease(oWn);
        }

        [DispId(0x0000073E)]
        public void SheetFollowHyperlink(object oSheet, Microsoft.Office.Interop.Excel.Hyperlink oTarget)
        {
            if (xlSheetFollowHyperlink != null)
                xlSheetFollowHyperlink(oSheet, oTarget);

            //Release any COM objects passed into the event
            ComRelease(oSheet);
            ComRelease(oTarget);
        }

        [DispId(0x0000086D)]
        public void SheetPivotTableUpdate(object oSheet, Microsoft.Office.Interop.Excel.PivotTable oTarget)
        {
            if (xlSheetPivotTableUpdate != null)
                xlSheetPivotTableUpdate(oSheet, oTarget);

            //Release any COM objects passed into the event
            ComRelease(oSheet);
            ComRelease(oTarget);
        }

        [DispId(0x00000870)]
        public void WorkbookPivotTableCloseConnection(Microsoft.Office.Interop.Excel._Workbook oWB, Microsoft.Office.Interop.Excel.PivotTable oTarget)
        {
            if (xlWorkbookPivotTableCloseConnection != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookPivotTableCloseConnection(oWB,oTarget);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
            ComRelease(oTarget);
        }

        [DispId(0x00000871)]
        public void WorkbookPivotTableOpenConnection(Microsoft.Office.Interop.Excel._Workbook oWB, Microsoft.Office.Interop.Excel.PivotTable oTarget)
        {
            if (xlWorkbookPivotTableOpenConnection != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookPivotTableOpenConnection(oWB,oTarget);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
            ComRelease(oTarget);
        }

        [DispId(0x000008F1)]
        public void WorkbookSync(Microsoft.Office.Interop.Excel._Workbook oWB, Microsoft.Office.Core.MsoSyncEventType SyncType)
        {
            if (xlWorkbookSync != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookSync(oWB,SyncType);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
        }

        [DispId(0x000008F2)]
        public void WorkbookBeforeXmlImport(Microsoft.Office.Interop.Excel._Workbook oWB, Microsoft.Office.Interop.Excel.XmlMap oMap, string sUrl, bool IsRefresh, ref bool Cancel)
        {
            if (xlWorkbookBeforeXmlImport != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookBeforeXmlImport(oWB,oMap,sUrl,IsRefresh,ref Cancel);
                }
            }
            //Release any COM objects passed into the event
            ComRelease(oWB);
            ComRelease(oMap);
        }

        [DispId(0x000008F3)]
        public void WorkbookAfterXmlImport(Microsoft.Office.Interop.Excel._Workbook oWB, Microsoft.Office.Interop.Excel.XmlMap oMap, bool IsRefresh, Microsoft.Office.Interop.Excel.XlXmlImportResult Result)
        {
            if (xlWorkbookAfterXmlImport != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookAfterXmlImport(oWB,oMap,IsRefresh,Result);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
            ComRelease(oMap);
        }

        [DispId(0x000008F4)]
        public void WorkbookBeforeXmlExport(Microsoft.Office.Interop.Excel._Workbook oWB, Microsoft.Office.Interop.Excel.XmlMap oMap, string sUrl, ref bool Cancel)
        {
            if (xlWorkbookBeforeXmlExport != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookBeforeXmlExport(oWB,oMap,sUrl,ref Cancel);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
            ComRelease(oMap);
        }

        [DispId(0x000008F5)]
        public void WorkbookAfterXmlExport(Microsoft.Office.Interop.Excel._Workbook oWB, Microsoft.Office.Interop.Excel.XmlMap oMap, string sUrl, Microsoft.Office.Interop.Excel.XlXmlExportResult Result)
        {
            if (xlWorkbookAfterXmlExport != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookAfterXmlExport(oWB,oMap,sUrl,Result);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
            ComRelease(oMap);
        }

        [DispId(0x00000A33)]
        public void WorkbookRowsetComplete(Microsoft.Office.Interop.Excel._Workbook oWB, string sDesciption, string sSheet, bool Success)
        {
            if (xlWorkbookRowsetComplete != null)
            {
                if (!m_DisableEventsIfEmbedded ||
                    (m_DisableEventsIfEmbedded && !IsEmbedded(ref oWB)))
                {
                    xlWorkbookRowsetComplete(oWB,sDesciption,sSheet,Success);
                }
            }

            //Release any COM objects passed into the event
            ComRelease(oWB);
        }

        [DispId(0x00000A34)]
        public void AfterCalculate()
        {
            if (xlAfterCalculate != null)
                xlAfterCalculate();
        }

		#endregion
	}
}
