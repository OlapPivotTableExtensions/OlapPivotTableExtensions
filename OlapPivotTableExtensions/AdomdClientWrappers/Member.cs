extern alias ExcelAdomdClientReference;

using System;
using System.Collections.Generic;
using System.Text;
using AsAdomdClient = Microsoft.AnalysisServices.AdomdClient;
using ExcelAdomdClient = ExcelAdomdClientReference::Microsoft.AnalysisServices.AdomdClient;

namespace OlapPivotTableExtensions.AdomdClientWrappers
{
    public class Member
    {
        private AsAdomdClient.Member _obj;
        private ExcelAdomdClient.Member _objExcel;

        public Member(AsAdomdClient.Member obj)
        {
            _obj = obj;
        }
        public Member(ExcelAdomdClient.Member obj)
        {
            _objExcel = obj;
        }

        public static bool operator ==(Member a, Member b)
        {
            // If both are null, or both are same instance, return true.
            if (System.Object.ReferenceEquals(a, b))
            {
                return true;
            }

            // If one is null, but not both, return false.
            if (((object)a == null) || ((object)b == null))
            {
                return false;
            }

            // Return true if the fields match:
            return a.UniqueName == b.UniqueName;
        }

        public static bool operator !=(Member a, Member b)
        {
            return !(a == b);
        }

        public string Caption
        {
            get
            {
                if (_obj != null)
                {
                    return _obj.Caption;
                }
                else
                {
                    ExcelAdoMdConnections.ReturnDelegate<string> f = delegate
                    {
                        return _objExcel.Caption;
                    };
                    return f();
                }
            }
        }

        public string UniqueName
        {
            get
            {
                if (_obj != null)
                {
                    return _obj.UniqueName;
                }
                else
                {
                    ExcelAdoMdConnections.ReturnDelegate<string> f = delegate
                    {
                        return _objExcel.UniqueName;
                    };
                    return f();
                }
            }
        }

        public Level ParentLevel
        {
            get
            {
                if (_obj != null)
                {
                    return new Level(_obj.ParentLevel);
                }
                else
                {
                    ExcelAdoMdConnections.ReturnDelegate<Level> f = delegate
                    {
                        return new Level(_objExcel.ParentLevel);
                    };
                    return f();
                }
            }
        }

        public Member Parent
        {
            get
            {
                if (_obj != null)
                {
                    if (_obj.Parent == null)
                        return null;
                    else
                        return new Member(_obj.Parent);
                }
                else
                {
                    ExcelAdoMdConnections.ReturnDelegate<Member> f = delegate
                    {
                        if (_objExcel.Parent == null)
                            return null;
                        else
                            return new Member(_objExcel.Parent);
                    };
                    return f();
                }
            }
        }

        public MemberPropertyCollection MemberProperties
        {
            get
            {
                if (_obj != null)
                {
                    MemberPropertyCollection coll = new MemberPropertyCollection();
                    foreach (AsAdomdClient.MemberProperty member in _obj.MemberProperties)
                    {
                        coll.Add(new MemberProperty(member));
                    }
                    return coll;
                }
                else
                {
                    ExcelAdoMdConnections.ReturnDelegate<MemberPropertyCollection> f = delegate
                    {
                        MemberPropertyCollection coll = new MemberPropertyCollection();
                        foreach (ExcelAdomdClient.MemberProperty member in _objExcel.MemberProperties)
                        {
                            coll.Add(new MemberProperty(member));
                        }
                        return coll;
                    };
                    return f();
                }
            }
        }
    }

    public class MemberCollection : List<Member>
    {
    }

    public class MemberFilter
    {
    }
}
