extern alias ExcelAdomdClientReference;

using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using AsAdomdClient = Microsoft.AnalysisServices.AdomdClient;
using ExcelAdomdClient = ExcelAdomdClientReference::Microsoft.AnalysisServices.AdomdClient;

namespace OlapPivotTableExtensions.AdomdClientWrappers
{
    public class AdomdDataReader
    {
        private AdomdType _type;
        private AsAdomdClient.AdomdDataReader _reader;
        private ExcelAdomdClient.AdomdDataReader _readerExcel;

        public AdomdDataReader(AsAdomdClient.AdomdDataReader obj)
        {
            _type = AdomdType.AnalysisServices;
            _reader = obj;
        }
        public AdomdDataReader(ExcelAdomdClient.AdomdDataReader obj)
        {
            _type = AdomdType.Excel;
            _readerExcel = obj;
        }

        internal AdomdType Type
        {
            get { return _type; }
        }

        public int FieldCount
        {
            get
            {
                if (_type == AdomdType.AnalysisServices)
                {
                    return _reader.FieldCount;
                }
                else
                {
                    ExcelAdoMdConnections.ReturnDelegate<int> f = delegate
                    {
                        return _readerExcel.FieldCount;
                    };
                    return f();
                }
            }
        }

        public object this[string index]
        {
            get
            {
                if (_type == AdomdType.AnalysisServices)
                {
                    return _reader[index];
                }
                else
                {
                    ExcelAdoMdConnections.ReturnDelegate<object> f = delegate
                    {
                        return _readerExcel[index];
                    };
                    return f();
                }
            }
        }

        public object this[int index]
        {
            get
            {
                if (_type == AdomdType.AnalysisServices)
                {
                    return _reader[index];
                }
                else
                {
                    ExcelAdoMdConnections.ReturnDelegate<object> f = delegate
                    {
                        return _readerExcel[index];
                    };
                    return f();
                }
            }
        }

        public string GetName(int ordinal)
        {
            if (_type == AdomdType.AnalysisServices)
            {
                return _reader.GetName(ordinal);
            }
            else
            {
                ExcelAdoMdConnections.ReturnDelegate<string> f = delegate
                {
                    return _readerExcel.GetName(ordinal);
                };
                return f();
            }
        }

        public bool Read()
        {
            if (_type == AdomdType.AnalysisServices)
            {
                return _reader.Read();
            }
            else
            {
                ExcelAdoMdConnections.ReturnDelegate<bool> f = delegate
                {
                    return _readerExcel.Read();
                };
                return f();
            }
        }

        public void Close()
        {
            if (_type == AdomdType.AnalysisServices)
            {
                _reader.Close();
            }
            else
            {
                ExcelAdoMdConnections.VoidDelegate f = delegate
                {
                    _readerExcel.Close();
                };
                f();
            }
        }
    }
}
