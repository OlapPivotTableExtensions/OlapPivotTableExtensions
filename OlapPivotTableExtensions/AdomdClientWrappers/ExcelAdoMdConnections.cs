using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace OlapPivotTableExtensions.AdomdClientWrappers
{
    //Microsoft.Excel.AdomdClient.dll path logic from Microsoft.ReportingServices.AdHoc.Excel.Client.ExcelAdoMdConnections
    //Microsoft.Excel.AdomdClient.dll assembly loading improved over that approach
    internal class ExcelAdoMdConnections
    {
        internal delegate void VoidDelegate();
        internal delegate T ReturnDelegate<T>();

        private Assembly m_excelAdomdClientAssembly;
        private string m_excelAdomdClientAssemblyPath;

        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        private static extern uint GetModuleFileName([In] IntPtr hModule, [Out] StringBuilder lpFilename, [In, MarshalAs(UnmanagedType.U4)] int nSize);
        [DllImport("Kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        protected string RetrieveAdomdClientAssemblyPath()
        {
            IntPtr moduleHandle = GetModuleHandle("msmdlocal_xl.dll");
            if (moduleHandle == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                throw new Win32Exception(error);
            }
            StringBuilder lpFilename = new StringBuilder(0x400);
            if (GetModuleFileName(moduleHandle, lpFilename, lpFilename.Capacity) == 0)
            {
                int num3 = Marshal.GetLastWin32Error();
                throw new Win32Exception(num3);
            }
            string directoryName = Path.GetDirectoryName(lpFilename.ToString());
            return Path.Combine(directoryName, "Microsoft.Excel.AdomdClient.dll");
        }

        internal Assembly ExcelAdomdClientAssembly
        {
            get
            {
                if (this.m_excelAdomdClientAssembly == null)
                {
                    this.m_excelAdomdClientAssembly = Assembly.LoadFrom(this.ExcelAdomdClientAssemblyPath);
                }
                return this.m_excelAdomdClientAssembly;
            }
        }

        protected string ExcelAdomdClientAssemblyPath
        {
            get
            {
                if (this.m_excelAdomdClientAssemblyPath == null)
                {
                    this.m_excelAdomdClientAssemblyPath = this.RetrieveAdomdClientAssemblyPath();
                }
                return this.m_excelAdomdClientAssemblyPath;
            }
        }

    }
}
