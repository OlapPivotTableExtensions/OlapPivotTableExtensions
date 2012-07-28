extern alias ExcelAdomdClientReference;

using System;
using System.Collections.Generic;
using System.Text;
using AsAdomdClient = Microsoft.AnalysisServices.AdomdClient;
using ExcelAdomdClient = ExcelAdomdClientReference::Microsoft.AnalysisServices.AdomdClient;

namespace OlapPivotTableExtensions.AdomdClientWrappers
{
    public class AdomdCommand
    {
        private AsAdomdClient.AdomdCommand _obj;
        private ExcelAdomdClient.AdomdCommand _objExcel;

        public AdomdCommand() { }
        public AdomdCommand(AsAdomdClient.AdomdCommand obj)
        {
            _obj = obj;
        }
        public AdomdCommand(ExcelAdomdClient.AdomdCommand obj)
        {
            _objExcel = obj;
        }

        public AdomdConnection Connection
        {
            get
            {
                if (_obj != null)
                {
                    return new AdomdConnection(_obj.Connection);
                }
                else
                {
                    ExcelAdoMdConnections.ReturnDelegate<AdomdConnection> f = delegate
                    {
                        return new AdomdConnection(_objExcel.Connection);
                    };
                    return f();
                }
            }

            set
            {
                if (value.Type == AdomdType.AnalysisServices)
                {
                    if (_obj == null)
                        _obj = new AsAdomdClient.AdomdCommand();
                    _obj.Connection = (AsAdomdClient.AdomdConnection)value.UnderlyingConnection;
                }
                else
                {
                    ExcelAdoMdConnections.VoidDelegate f = delegate
                    {
                        if (_objExcel == null)
                            _objExcel = new ExcelAdomdClient.AdomdCommand();
                        _objExcel.Connection = (ExcelAdomdClient.AdomdConnection)value.UnderlyingConnection;
                    };
                    f();
                }
            }
        }

        public string CommandText
        {
            get
            {
                if (_obj != null)
                {
                    return _obj.CommandText;
                }
                else
                {
                    ExcelAdoMdConnections.ReturnDelegate<string> f = delegate
                    {
                        return _objExcel.CommandText;
                    };
                    return f();
                }
            }

            set
            {
                if (_obj != null)
                {
                    _obj.CommandText = value;
                }
                else
                {
                    ExcelAdoMdConnections.VoidDelegate f = delegate
                    {
                        _objExcel.CommandText = value;
                    };
                    f();
                }
            }
        }

        public void Cancel()
        {
            if (_obj != null)
            {
                _obj.Cancel();
            }
            else
            {
                ExcelAdoMdConnections.VoidDelegate f = delegate
                {
                    _objExcel.Cancel();
                };
                f();
            }
        }

        AdomdParameterCollection _parameters = new AdomdParameterCollection();
        public AdomdParameterCollection Parameters
        {
            get
            {
                return _parameters;
            }
        }

        public CellSet ExecuteCellSet()
        {
            if (_obj != null)
            {
                _obj.Parameters.Clear();
                foreach (AdomdParameter param in _parameters)
                {
                    _obj.Parameters.Add(new AsAdomdClient.AdomdParameter(param.Name, param.Value));
                }
                return new CellSet(_obj.ExecuteCellSet());
            }
            else
            {
                ExcelAdoMdConnections.ReturnDelegate<CellSet> f = delegate
                {
                    _objExcel.Parameters.Clear();
                    foreach (AdomdParameter param in _parameters)
                    {
                        _objExcel.Parameters.Add(new ExcelAdomdClient.AdomdParameter(param.Name, param.Value));
                    }
                    return new CellSet(_objExcel.ExecuteCellSet());
                };
                return f();
            }
        }
    }

}
