extern alias ExcelAdomdClientReference;

using System;
using System.Collections.Generic;
using System.Text;
using AsAdomdClient = Microsoft.AnalysisServices.AdomdClient;
using ExcelAdomdClient = ExcelAdomdClientReference::Microsoft.AnalysisServices.AdomdClient;

namespace OlapPivotTableExtensions.AdomdClientWrappers
{
    public class Cell
    {
        private AsAdomdClient.Cell _obj;
        private ExcelAdomdClient.Cell _objExcel;

        public Cell(AsAdomdClient.Cell obj)
        {
            _obj = obj;
        }
        public Cell(ExcelAdomdClient.Cell obj)
        {
            _objExcel = obj;
        }

        public object Value
        {
            get
            {
                if (_obj != null)
                {
                    return _obj.Value;
                }
                else
                {
                    ExcelAdoMdConnections.ReturnDelegate<object> f = delegate
                    {
                        return _objExcel.Value;
                    };
                    return f();
                }
            }
        }

    }
}
