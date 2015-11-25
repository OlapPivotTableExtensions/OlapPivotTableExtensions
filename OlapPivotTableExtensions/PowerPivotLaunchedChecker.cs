using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OlapPivotTableExtensions
{
    /// <summary>
    /// This class is launched in the default AppDomain (not the OLAP PivotTable Extensions AppDomain) and looks at the currently loaded assemblies to see if a Power Pivot (Excel Data Model) assembly is loaded yet.
    /// If Power Pivot hasn't been launched yet, then enabling auto refresh will fail in Excel 2016
    /// </summary>
    [Serializable]
    public class PowerPivotLaunchedChecker : MarshalByRefObject
    {
        public bool IsPowerPivotLoaded = false;
        public PowerPivotLaunchedChecker()
        {
            AppDomain appD = AppDomain.CurrentDomain;
            bool bIsPowerPivotLoaded = false;
            foreach (System.Reflection.Assembly ass in appD.GetAssemblies())
            {
                try
                {
                    if (ass.FullName.StartsWith("Microsoft.Office.Excel.DataModel"))
                    {
                        bIsPowerPivotLoaded = true;
                        break;
                    }
                }
                catch { }
            }
            IsPowerPivotLoaded = bIsPowerPivotLoaded;
        }
    }

}
