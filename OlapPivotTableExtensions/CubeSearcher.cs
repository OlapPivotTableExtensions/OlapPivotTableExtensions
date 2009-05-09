using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.AnalysisServices.AdomdClient;
using System.ComponentModel;

namespace OlapPivotTableExtensions
{
    /// <summary>
    /// A generic class that uses AdomdClient to search a cube. Searches metadata and dimension members. Has no dependencies on PivotTables.
    /// </summary>
    public class CubeSearcher
    {
        private CubeSearchScope _scope;
        private CubeDef _cube;
        private string _searchString;
        private bool _exactMatch;
        private bool _searchMemberProperties = false;
        private string _searchOnly;
        private int _totalTaskCount;
        private int _completedTaskCount;
        private BackgroundWorker _thread;
        private Exception _error;
        private SortableList<CubeSearchMatch> _listMatches;
        private AdomdCommand cmd;
        private bool _Complete = false;
        private System.Windows.Forms.Control _checkInvokeRequired;
        private static Dictionary<string, string> _measureGroupCaptions;

        public CubeSearcher(CubeDef Cube, CubeSearchScope Scope, string SearchString, bool ExactMatch, bool SearchMemberProperties, string SearchOnly, System.Windows.Forms.Control ConsumingControl)
        {
            _cube = Cube;
            _scope = Scope;
            _searchString = SearchString;
            _exactMatch = ExactMatch;
            _searchMemberProperties = SearchMemberProperties;
            _searchOnly = SearchOnly;
            _checkInvokeRequired = ConsumingControl;
        }

        public SortableList<CubeSearchMatch> Matches
        {
            get { return _listMatches; }
        }

        public void SearchAsync()
        {
            _listMatches = new SortableList<CubeSearchMatch>();

            _error = null;
            _completedTaskCount = 0;
            _totalTaskCount = 1;
            _measureGroupCaptions = new Dictionary<string, string>();

            _thread = new BackgroundWorker();
            _thread.WorkerSupportsCancellation = true;
            _thread.DoWork += new DoWorkEventHandler(_thread_DoWork);
            _thread.RunWorkerAsync();
        }

        public void Cancel()
        {
            try
            {
                _thread.CancelAsync();
                if (cmd != null)
                    cmd.Cancel();
            }
            catch { }
        }

        public string Error
        {
            get
            {
                if (_error != null)
                    return _error.Message;
                else
                    return null;
            }
        }

        public bool Complete
        {
            get { return _Complete; }
        }

        private delegate void AddMatch_Delegate(CubeSearchMatch match);
        private void AddMatch(CubeSearchMatch match)
        {
            try
            {
                if (_checkInvokeRequired != null && _checkInvokeRequired.InvokeRequired)
                {
                    //avoid the "cross-thread operation not valid" error message
                    //since a control is using this list as a BindingSource, we have to update the list this way
                    _checkInvokeRequired.BeginInvoke(new AddMatch_Delegate(AddMatch), new object[] { match });
                }
                else
                {
                    _listMatches.Add(match);
                }
            }
            catch (Exception ex)
            {
                _error = ex;
            }
        }

        /// <summary>
        /// Gets properly translated caption for measure group
        /// </summary>
        /// <param name="DatabaseName"></param>
        /// <param name="CubeName"></param>
        /// <param name="MeasureGroupName"></param>
        /// <returns></returns>
        private static string GetMeasureGroupCaption(string DatabaseName, string CubeName, string MeasureGroupName)
        {
            string sKey = DatabaseName + "|" + CubeName + "|" + MeasureGroupName;
            if (_measureGroupCaptions.ContainsKey(sKey))
                return _measureGroupCaptions[sKey];
            return null;
        }

        public event ProgressChangedEventHandler ProgressChanged;

        private void _thread_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                _Complete = false;
                _totalTaskCount = 0;

                if (_cube.ParentConnection.State != System.Data.ConnectionState.Open)
                    _cube.ParentConnection.Open();

                if (string.IsNullOrEmpty(_searchOnly))
                {
                    foreach (Dimension d in _cube.Dimensions)
                        _totalTaskCount += d.Hierarchies.Count;
                }
                else
                {
                    _totalTaskCount = 1;
                }

                System.Data.DataTable tblProperties = new System.Data.DataTable();
                if (_searchMemberProperties)
                {
                    //get all member properties for the entire cube
                    AdomdRestrictionCollection restrictions = new AdomdRestrictionCollection();
                    restrictions.Add(new AdomdRestriction("CATALOG_NAME", _cube.ParentConnection.Database));
                    restrictions.Add(new AdomdRestriction("CUBE_NAME", _cube.Name));
                    restrictions.Add(new AdomdRestriction("PROPERTY_TYPE", 1));
                    tblProperties = _cube.ParentConnection.GetSchemaDataSet("MDSCHEMA_PROPERTIES", restrictions).Tables[0];
                }

                if (_scope == CubeSearchScope.FieldList)
                {
                    ////////////////////////////////////////////////////////////////////
                    // SEARCH FIELD LIST
                    ////////////////////////////////////////////////////////////////////

                    if (GetSSASServerVersion() >= 2005)
                    {
                        //build a list of measure groups and their captions so the "Folder" property of a match will be correctly translated
                        AdomdRestrictionCollection restrictions = new AdomdRestrictionCollection();
                        restrictions.Add(new AdomdRestriction("CATALOG_NAME", _cube.ParentConnection.Database));
                        restrictions.Add(new AdomdRestriction("CUBE_NAME", _cube.Name));
                        foreach (System.Data.DataRow r in _cube.ParentConnection.GetSchemaDataSet("MDSCHEMA_MEASUREGROUPS", restrictions).Tables[0].Rows)
                        {
                            _measureGroupCaptions[Convert.ToString(r["CATALOG_NAME"]) + "|" + Convert.ToString(r["CUBE_NAME"]) + "|" + Convert.ToString(r["MEASUREGROUP_NAME"])] = Convert.ToString(r["MEASUREGROUP_CAPTION"]);
                        }
                    }

                    _totalTaskCount++; //for measures search task
                    foreach (Measure m in _cube.Measures)
                    {
                        if (_thread.CancellationPending) return;
                        if (_exactMatch)
                        {
                            if (string.Compare(m.Caption, _searchString, true) == 0)
                            {
                                AddMatch(new CubeSearchMatch(m));
                            }
                        }
                        else
                        {
                            if (m.Caption.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                AddMatch(new CubeSearchMatch(m));
                            }
                            else if (m.Description.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                AddMatch(new CubeSearchMatch(m));
                            }
                            else if (m.DisplayFolder.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                AddMatch(new CubeSearchMatch(m));
                            }
                            else
                            {
                                //search the measure group caption
                                string sMeasureGroup = Convert.ToString(m.Properties["MEASUREGROUP_NAME"].Value);
                                if (!string.IsNullOrEmpty(sMeasureGroup))
                                {
                                    string sMeasureGroupCaption = CubeSearcher.GetMeasureGroupCaption(m.ParentCube.ParentConnection.Database, m.ParentCube.Name, sMeasureGroup);
                                    sMeasureGroup = (sMeasureGroupCaption != null ? sMeasureGroupCaption : sMeasureGroup);
                                    if (sMeasureGroup.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                    {
                                        AddMatch(new CubeSearchMatch(m));
                                    }
                                }
                            }
                        }
                    }
                    _completedTaskCount++;
                    ProgressChanged.Invoke(this, new ProgressChangedEventArgs((int)(100 * _completedTaskCount / ((double)_totalTaskCount)), null));

                    foreach (Dimension d in _cube.Dimensions)
                    {
                        if (_exactMatch)
                        {
                            if (string.Compare(d.Caption, _searchString, true) == 0)
                            {
                                AddMatch(new CubeSearchMatch(d));
                            }
                        }
                        else
                        {
                            if (d.Caption.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                AddMatch(new CubeSearchMatch(d));
                            }
                            else if (d.Description.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                AddMatch(new CubeSearchMatch(d));
                            }
                        }
                        foreach (Hierarchy h in d.Hierarchies)
                        {
                            if (_thread.CancellationPending) return;
                            if (_exactMatch)
                            {
                                if (string.Compare(h.Caption, _searchString, true) == 0)
                                {
                                    AddMatch(new CubeSearchMatch(h));
                                }
                            }
                            else
                            {
                                if (h.Caption.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                {
                                    AddMatch(new CubeSearchMatch(h));
                                }
                                else if (h.Description.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                {
                                    AddMatch(new CubeSearchMatch(h));
                                }
                                else if (h.DisplayFolder.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                {
                                    AddMatch(new CubeSearchMatch(h));
                                }
                            }
                            foreach (Level l in h.Levels)
                            {
                                if (h.HierarchyOrigin != HierarchyOrigin.AttributeHierarchy)
                                {
                                    if (_exactMatch)
                                    {
                                        if (string.Compare(l.Caption, _searchString, true) == 0)
                                        {
                                            AddMatch(new CubeSearchMatch(l));
                                        }
                                    }
                                    else
                                    {
                                        if (l.Caption.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            AddMatch(new CubeSearchMatch(l));
                                        }
                                        else if (l.Description.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            AddMatch(new CubeSearchMatch(l));
                                        }
                                    }
                                }

                                if (l.LevelType != LevelTypeEnum.All || h.HierarchyOrigin == HierarchyOrigin.ParentChildHierarchy)
                                {
                                    //search member properties
                                    foreach (System.Data.DataRow row in tblProperties.Rows)
                                    {
                                        string sLevelUniqueName = Convert.ToString(row["LEVEL_UNIQUE_NAME"]);
                                        if (sLevelUniqueName != l.UniqueName) continue;

                                        string sPropertyName = Convert.ToString(row["PROPERTY_NAME"]);
                                        string sPropertyCaption = Convert.ToString(row["PROPERTY_CAPTION"]);
                                        string sDescription = Convert.ToString(row["DESCRIPTION"]);
                                        bool bIsMatch = false;
                                        if (_exactMatch)
                                        {
                                            if (string.Compare(sPropertyCaption, _searchString, true) == 0)
                                            {
                                                bIsMatch = true;
                                            }
                                        }
                                        else
                                        {
                                            if (sPropertyCaption.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                            {
                                                bIsMatch = true;
                                            }
                                            else if (sDescription.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                            {
                                                bIsMatch = true;
                                            }
                                        }
                                        if (bIsMatch)
                                        {
                                            //need to retrieve a MemberProperty object, so find one member
                                            foreach (Member m in l.GetMembers(0, 1, new string[] { l.UniqueName + ".[" + sPropertyName + "]" }, new MemberFilter[] { }))
                                            {
                                                AddMatch(new CubeSearchMatch(l, m.MemberProperties[sPropertyName], sDescription));
                                            }
                                        }
                                    }
                                }
                            }

                            _completedTaskCount++;
                            ProgressChanged.Invoke(this, new ProgressChangedEventArgs((int)(100 * _completedTaskCount / ((double)_totalTaskCount)), null));
                        }
                    }

                    //search KPIs
                    foreach (Kpi k in _cube.Kpis)
                    {
                        if (_exactMatch)
                        {
                            if (string.Compare(k.Caption, _searchString, true) == 0)
                            {
                                AddMatch(new CubeSearchMatch(k));
                            }
                        }
                        else
                        {
                            if (k.Caption.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                AddMatch(new CubeSearchMatch(k));
                            }
                            else if (k.Description.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                AddMatch(new CubeSearchMatch(k));
                            }
                            else if (k.DisplayFolder.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                AddMatch(new CubeSearchMatch(k));
                            }
                        }
                    }

                    //search named sets
                    foreach (NamedSet s in _cube.NamedSets)
                    {
                        string sDimensions = Convert.ToString(s.Properties["DIMENSIONS"].Value);
                        if (sDimensions.Contains("],["))
                            continue; //only dimensions with single hierarchies are shown in Excel

                        string sSetCaption = Convert.ToString(s.Properties["SET_CAPTION"].Value);
                        string sDisplayFolder = Convert.ToString(s.Properties["SET_DISPLAY_FOLDER"].Value);
                        if (string.IsNullOrEmpty(sDisplayFolder)) sDisplayFolder = "Sets";

                        if (_exactMatch)
                        {
                            if (string.Compare(sSetCaption, _searchString, true) == 0)
                            {
                                AddMatch(new CubeSearchMatch(s));
                            }
                        }
                        else
                        {
                            if (sSetCaption.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                AddMatch(new CubeSearchMatch(s));
                            }
                            else if (s.Description.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                AddMatch(new CubeSearchMatch(s));
                            }
                            else if (sDisplayFolder.IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                AddMatch(new CubeSearchMatch(s));
                            }
                        }
                    }
                }
                else
                {
                    ////////////////////////////////////////////////////////////////////
                    // SEARCH FOR MEMBERS
                    ////////////////////////////////////////////////////////////////////
                    cmd = new AdomdCommand();
                    cmd.Connection = _cube.ParentConnection;

                    if (_thread.CancellationPending) return;

                    //do quick full cube search for exact match in any dimension... this code uses the name hash index and is very fast (except for ROLAP dimensions)
                    //even if not looking for an exact match, run this code every time because it is so much faster than Filter(AllMembers) function
                    AdomdRestrictionCollection restrictions = new AdomdRestrictionCollection();
                    restrictions.Add(new AdomdRestriction("CATALOG_NAME", _cube.ParentConnection.Database));
                    restrictions.Add(new AdomdRestriction("CUBE_NAME", _cube.Name));
                    if (!string.IsNullOrEmpty(_searchOnly))
                        restrictions.Add(new AdomdRestriction("HIERARCHY_UNIQUE_NAME", _searchOnly));
                    restrictions.Add(new AdomdRestriction("MEMBER_NAME", _searchString));
                    System.Data.DataTable tblExactMatchMembers = _cube.ParentConnection.GetSchemaDataSet("MDSCHEMA_MEMBERS", restrictions).Tables[0];

                    List<string> listFoundMemberUniqueNames = new List<string>();
                    Dictionary<Hierarchy, List<Member>> dictFoundHierarchyMembers = new Dictionary<Hierarchy,List<Member>>();
                    Dictionary<string, string> dictFoundHierarchyMembersString = new Dictionary<string,string>();
                    foreach (System.Data.DataRow row in tblExactMatchMembers.Rows)
                    {
                        string sHier = Convert.ToString(row["HIERARCHY_UNIQUE_NAME"]);
                        if (sHier.ToLower().StartsWith("[measures]")) continue;
                        string sMemb = Convert.ToString(row["MEMBER_UNIQUE_NAME"]);
                        if (!dictFoundHierarchyMembersString.ContainsKey(sHier))
                            dictFoundHierarchyMembersString.Add(sHier, sMemb);
                        else if (dictFoundHierarchyMembersString[sHier].StartsWith("{"))
                            dictFoundHierarchyMembersString[sHier].Insert(dictFoundHierarchyMembersString[sHier].Length - 1, sMemb);
                        else
                            dictFoundHierarchyMembersString[sHier] = "{" + dictFoundHierarchyMembersString[sHier] + "," + sMemb + "}";
                    }

                    CellSet cs = null;
                    foreach (string sSet in dictFoundHierarchyMembersString.Values)
                    {
                        cmd.CommandText = "select {} on 0, " + sSet + " dimension properties Member_Type on 1 from [" + _cube.Name + "]";

                        if (_thread.CancellationPending) return;

                        cs = cmd.ExecuteCellSet();

                        if (_thread.CancellationPending) return;
                        
                        foreach (Position p in cs.Axes[1].Positions)
                        {
                            foreach (Member m in p.Members)
                            {
                                if (dictFoundHierarchyMembers.ContainsKey(m.ParentLevel.ParentHierarchy))
                                    dictFoundHierarchyMembers[m.ParentLevel.ParentHierarchy].Add(m);
                                else
                                    dictFoundHierarchyMembers.Add(m.ParentLevel.ParentHierarchy, new List<Member>(new Member[] { m }));
                                listFoundMemberUniqueNames.Add(m.UniqueName);
                                AddMatch(new CubeSearchMatch(m));
                            }
                        }
                        ProgressChanged.Invoke(this, new ProgressChangedEventArgs(1, null));
                    }

                    if (_thread.CancellationPending) return;
                    ProgressChanged.Invoke(this, new ProgressChangedEventArgs(1, null));

                    Hierarchy hierSearchOnly = null;

                    //now do traditional search by looping all hierarchies and executing one MDX query per
                    //put dimensions into structure that can be sorted by dimension size
                    //sort dimensions based on cardinality
                    List<Dimension> listDimensions = new List<Dimension>(_cube.Dimensions.Count);
                    if (string.IsNullOrEmpty(_searchOnly))
                    {
                        foreach (Dimension d in _cube.Dimensions)
                            listDimensions.Add(d);
                        listDimensions.Sort(delegate(Dimension x, Dimension y) { return ((uint)x.Properties["DIMENSION_CARDINALITY"].Value).CompareTo((uint)y.Properties["DIMENSION_CARDINALITY"].Value); });
                    }
                    else
                    {
                        foreach (Dimension d in _cube.Dimensions)
                        {
                            foreach (Hierarchy h in d.Hierarchies)
                            {
                                if (string.Compare(h.UniqueName, _searchOnly, true) == 0)
                                {
                                    hierSearchOnly = h;
                                    listDimensions.Add(d);
                                    break;
                                }
                            }
                        }
                    }

                    ProgressChanged.Invoke(this, new ProgressChangedEventArgs(2, null));

                    cmd.Parameters.Clear();
                    cmd.Parameters.Add(new AdomdParameter("SearchString", _searchString.ToLower())); //prevent "SQL-injection"

                    if (!_exactMatch || _searchMemberProperties)
                    {
                        //search each hierarchy... start with the smallest dimensions
                        foreach (Dimension d in listDimensions)
                        {
                            foreach (Hierarchy h in d.Hierarchies)
                            {
                                if (_thread.CancellationPending) return;
                                if (hierSearchOnly != null && h != hierSearchOnly) continue;

                                //TODO: future... test a named set with a bracket in the name... currently Excel 2007 doesn't support them because it doesn't escape the name right
                                if (_searchMemberProperties)
                                {
                                    foreach (Level l in h.Levels)
                                    {
                                        string sLevelMembers = l.UniqueName + ".AllMembers";

                                        List<string> listProperties = new List<string>();
                                        string sFilterPropertiesMDX = string.Empty;
                                        string sDimensionPropertiesClause = "Member_Type";
                                        foreach (System.Data.DataRow row in tblProperties.Rows)
                                        {
                                            string sLevelUniqueName = Convert.ToString(row["LEVEL_UNIQUE_NAME"]);
                                            if (sLevelUniqueName != l.UniqueName) continue;

                                            string sPropertyName = Convert.ToString(row["PROPERTY_NAME"]);
                                            if (sDimensionPropertiesClause.Length > 0) sDimensionPropertiesClause += ", ";
                                            sDimensionPropertiesClause += sLevelUniqueName + ".[" + sPropertyName + "]";
                                            if (!listProperties.Contains(sPropertyName))
                                            {
                                                if (_exactMatch)
                                                    sFilterPropertiesMDX += " or LCase(" + h.UniqueName + ".CurrentMember.Properties(\"" + sPropertyName + "\")) = @SearchString";
                                                else
                                                    sFilterPropertiesMDX += " or InStr(LCase(" + h.UniqueName + ".CurrentMember.Properties(\"" + sPropertyName + "\")), @SearchString) > 0";
                                                listProperties.Add(sPropertyName);
                                            }
                                        }
                                        if (sDimensionPropertiesClause.Length > 0)
                                            sDimensionPropertiesClause = "dimension properties " + sDimensionPropertiesClause;

                                        if (_exactMatch)
                                            cmd.CommandText = "select {} on 0, Filter(" + sLevelMembers + ", LCase(" + h.UniqueName + ".CurrentMember.Member_Caption) = @SearchString" + sFilterPropertiesMDX + ") " + sDimensionPropertiesClause + " on 1 from [" + _cube.Name + "]";
                                        else
                                            cmd.CommandText = "select {} on 0, Filter(" + sLevelMembers + ", InStr(LCase(" + h.UniqueName + ".CurrentMember.Member_Caption), @SearchString) > 0" + sFilterPropertiesMDX + ") " + sDimensionPropertiesClause + " on 1 from [" + _cube.Name + "]";

                                        if (_thread.CancellationPending) return;

                                        cs = cmd.ExecuteCellSet();

                                        if (_thread.CancellationPending) return;

                                        if (cs != null && cs.Axes.Count > 1)
                                        {
                                            foreach (Position p in cs.Axes[1].Positions)
                                            {
                                                foreach (Member m in p.Members)
                                                {
                                                    if (!listFoundMemberUniqueNames.Contains(m.UniqueName))
                                                    {
                                                        //the quick full cube search above didn't find this member
                                                        bool bFound = false;
                                                        foreach (MemberProperty mp in m.MemberProperties)
                                                        {
                                                            if (mp.Name == "MEMBER_TYPE") continue;
                                                            if (_exactMatch)
                                                            {
                                                                if (string.Compare(Convert.ToString(mp.Value), _searchString, true) == 0)
                                                                {
                                                                    AddMatch(new CubeSearchMatch(m, mp));
                                                                    bFound = true;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (Convert.ToString(mp.Value).IndexOf(_searchString, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                                                {
                                                                    AddMatch(new CubeSearchMatch(m, mp));
                                                                    bFound = true;
                                                                }
                                                            }
                                                        }
                                                        if (!bFound)
                                                            AddMatch(new CubeSearchMatch(m));
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    string sHierarchyMembers = h.UniqueName + ".AllMembers";

                                    if (_exactMatch)
                                        cmd.CommandText = "select {} on 0, Filter(" + sHierarchyMembers + ", LCase(" + h.UniqueName + ".CurrentMember.Member_Caption) = @SearchString) dimension properties Member_Type on 1 from [" + _cube.Name + "]";
                                    else
                                        cmd.CommandText = "select {} on 0, Filter(" + sHierarchyMembers + ", InStr(LCase(" + h.UniqueName + ".CurrentMember.Member_Caption), @SearchString) > 0) dimension properties Member_Type on 1 from [" + _cube.Name + "]";

                                    if (_thread.CancellationPending) return;

                                    cs = cmd.ExecuteCellSet();

                                    if (_thread.CancellationPending) return;

                                    if (cs != null && cs.Axes.Count > 1)
                                    {
                                        foreach (Position p in cs.Axes[1].Positions)
                                        {
                                            foreach (Member m in p.Members)
                                            {
                                                if (!listFoundMemberUniqueNames.Contains(m.UniqueName))
                                                {
                                                    //the quick full cube search above didn't find this member
                                                    AddMatch(new CubeSearchMatch(m));
                                                }
                                            }
                                        }
                                    }
                                }

                                _completedTaskCount++;
                                ProgressChanged.Invoke(this, new ProgressChangedEventArgs((int)(97 * _completedTaskCount / ((double)_totalTaskCount)) + 2, null));
                            }
                        }
                    }
                }
                //TODO: future: don't search ROLAP dimensions?
            }
            catch (Exception ex)
            {
                if (!_thread.CancellationPending)
                {
                    _error = ex;
                }
            }
            finally
            {
                _Complete = true;
                try
                {
                    ProgressChanged.Invoke(this, new ProgressChangedEventArgs(100, null));
                }
                catch { }
            }
        }

        private int GetSSASServerVersion()
        {
            int iPos = _cube.ParentConnection.ServerVersion.IndexOf('.');
            int iVersion;
            if (iPos > 0 && int.TryParse(_cube.ParentConnection.ServerVersion.Substring(0, iPos), out iVersion))
            {
                if (iVersion >= 10)
                    return 2008;
                else if (iVersion >= 9)
                    return 2005;
                else if (iVersion >= 8)
                    return 2000;
            }
            throw new Exception("Couldn't determine SSAS server version: " + _cube.ParentConnection.ServerVersion);
        }

        public enum CubeSearchScope
        {
            FieldList,
            DimensionData
        }

        private enum MEMBER_TYPE
        {
            MDMEMBER_TYPE_REGULAR = 1,
            MDMEMBER_TYPE_ALL = 2,
            MDMEMBER_TYPE_MEASURE = 3,
            MDMEMBER_TYPE_FORMULA = 4,
            MDMEMBER_TYPE_UNKNOWN = 0
        }

        public class CubeSearchMatch
        {
            private object _InnerObject;
            private MemberProperty _MemberProperty;
            private string _Name;
            private string _Type;
            private string _Folder;
            private string _Description;
            private bool _Checked = false;
            private bool _IsCalculated = false;

            public CubeSearchMatch(Measure m)
            {
                _InnerObject = m;
                _Name = m.Caption;
                if (!string.IsNullOrEmpty(m.Expression))
                {
                    _IsCalculated = true;
                    _Type = "Calculated Measure";
                }
                else
                {
                    _Type = "Measure";
                }
                string sMeasureGroup = Convert.ToString(m.Properties["MEASUREGROUP_NAME"].Value);
                if (!string.IsNullOrEmpty(sMeasureGroup))
                {
                    string sMeasureGroupCaption = CubeSearcher.GetMeasureGroupCaption(m.ParentCube.ParentConnection.Database, m.ParentCube.Name, sMeasureGroup);
                    _Folder = (sMeasureGroupCaption != null ? sMeasureGroupCaption : sMeasureGroup);
                }
                else
                    _Folder = "Values";
                if (!string.IsNullOrEmpty(m.DisplayFolder) && m.DisplayFolder != "\\" && m.DisplayFolder != "/")
                    _Folder += "\\" + m.DisplayFolder;
                _Description = m.Description;
            }

            public CubeSearchMatch(Member m)
            {
                _InnerObject = m;
                if (Convert.ToInt32(m.MemberProperties["MEMBER_TYPE"].Value) == (int)MEMBER_TYPE.MDMEMBER_TYPE_FORMULA)
                {
                    _IsCalculated = true;
                    _Type = "Calculated Member";
                }
                else
                {
                    _Type = "Member";
                }
                _Name = m.Caption;
                _Folder = m.ParentLevel.ParentHierarchy.ParentDimension.Caption;
                if (!string.IsNullOrEmpty(m.ParentLevel.ParentHierarchy.DisplayFolder) && m.ParentLevel.ParentHierarchy.DisplayFolder != "\\" && m.ParentLevel.ParentHierarchy.DisplayFolder != "/")
                    _Folder += "\\" + m.ParentLevel.ParentHierarchy.DisplayFolder;
                _Folder += "\\" + m.ParentLevel.ParentHierarchy.Caption;
            }

            public CubeSearchMatch(Member m, MemberProperty mp)
            {
                _InnerObject = m;
                if (Convert.ToInt32(m.MemberProperties["MEMBER_TYPE"].Value) == (int)MEMBER_TYPE.MDMEMBER_TYPE_FORMULA)
                {
                    _IsCalculated = true;
                    _Type = "Calculated Member";
                }
                else
                {
                    _Type = "Member";
                }
                _MemberProperty = mp;
                _Name = m.Caption;
                _Folder = m.ParentLevel.ParentHierarchy.ParentDimension.Caption;
                if (!string.IsNullOrEmpty(m.ParentLevel.ParentHierarchy.DisplayFolder) && m.ParentLevel.ParentHierarchy.DisplayFolder != "\\" && m.ParentLevel.ParentHierarchy.DisplayFolder != "/")
                    _Folder += "\\" + m.ParentLevel.ParentHierarchy.DisplayFolder;
                _Folder += "\\" + m.ParentLevel.ParentHierarchy.Caption + "\\" + m.ParentLevel.Caption;
                _Description = "Property [" + mp.Name + "] = " + mp.Value;
            }

            public CubeSearchMatch(Level l, MemberProperty mp, string description)
            {
                _InnerObject = l;
                _MemberProperty = mp;
                _Name = mp.Name;
                _Type = "Property";
                _Folder = l.ParentHierarchy.ParentDimension.Caption;
                if (!string.IsNullOrEmpty(l.ParentHierarchy.DisplayFolder) && l.ParentHierarchy.DisplayFolder != "\\" && l.ParentHierarchy.DisplayFolder != "/")
                    _Folder += "\\" + l.ParentHierarchy.DisplayFolder;
                _Folder += "\\" + l.ParentHierarchy.Caption + "\\" + l.Caption;
                _Description = description;
            }

            public CubeSearchMatch(Dimension d)
            {
                _InnerObject = d;
                _Name = d.Caption;
                _Type = "Dimension";
                _Description = d.Description;
            }

            public CubeSearchMatch(Hierarchy h)
            {
                _InnerObject = h;
                _Name = h.Caption;
                if (h.HierarchyOrigin == HierarchyOrigin.AttributeHierarchy)
                    _Type = "Attribute";
                else
                    _Type = "Hierarchy";
                _Folder = h.ParentDimension.Caption;
                if (!string.IsNullOrEmpty(h.DisplayFolder) && h.DisplayFolder != "\\" && h.DisplayFolder != "/")
                    _Folder += "\\" + h.DisplayFolder;
                _Description = h.Description;
            }

            public CubeSearchMatch(Level l)
            {
                _InnerObject = l;
                _Name = l.Caption;
                _Type = "Level";
                _Folder = l.ParentHierarchy.ParentDimension.Caption;
                if (!string.IsNullOrEmpty(l.ParentHierarchy.DisplayFolder) && l.ParentHierarchy.DisplayFolder != "\\" && l.ParentHierarchy.DisplayFolder != "/")
                    _Folder += "\\" + l.ParentHierarchy.DisplayFolder;
                _Folder += "\\" + l.ParentHierarchy.Caption;
                _Description = l.Description;
            }

            public CubeSearchMatch(Kpi k)
            {
                _InnerObject = k;
                _Name = k.Caption;
                _Type = "KPI";

                _Folder = string.Empty;
                if (k.ParentKpi != null)
                {
                    Kpi k2 = k;
                    while (k2.ParentKpi != null)
                    {
                        k2 = k2.ParentKpi;
                        if (_Folder.Length > 0) _Folder = _Folder + "\\";
                        _Folder = _Folder + k2.Caption;
                    }
                    if (_Folder.Length > 0)
                        _Folder = k2.DisplayFolder + "\\" + _Folder;
                    else
                        _Folder = k2.DisplayFolder;
                }
                else
                {
                    _Folder = k.DisplayFolder;
                }
                _Description = k.Description;
            }

            public CubeSearchMatch(NamedSet s)
            {
                _InnerObject = s;
                string sSetCaption = Convert.ToString(s.Properties["SET_CAPTION"].Value);
                string sDisplayFolder = Convert.ToString(s.Properties["SET_DISPLAY_FOLDER"].Value);
                string sDimensions = Convert.ToString(s.Properties["DIMENSIONS"].Value);
                string sDimension = sDimensions.Substring(1, sDimensions.IndexOf('.') - 2);
                Dimension d = s.ParentCube.Dimensions.Find(sDimension);
                if (d == null)
                    throw new Exception("Can't find dimension " + sDimension + " for named set " + sSetCaption);

                if (string.IsNullOrEmpty(sDisplayFolder)) sDisplayFolder = "Sets";

                _Name = sSetCaption;
                _Type = "Named Set";
                _Folder = d.Caption + " \\ " + sDisplayFolder;
                _Description = s.Description;
            }

            public bool Checked
            {
                get { return _Checked; }
                set { _Checked = value; }
            }

            public string Name
            {
                get { return _Name; }
            }

            public string Type
            {
                get { return _Type; }
            }

            public string Folder
            {
                get { return _Folder; }
            }

            public string Description
            {
                get { return _Description; }
            }

            public object InnerObject
            {
                get { return _InnerObject; }
            }

            public MemberProperty MemberProperty
            {
                get { return _MemberProperty; }
            }

            public bool IsCalculated
            {
                get { return _IsCalculated; }
            }

            public string UniqueName
            {
                get
                {
                    if (_InnerObject is Measure)
                        return ((Measure)_InnerObject).UniqueName;
                    else if (_InnerObject is Dimension)
                        return ((Dimension)_InnerObject).UniqueName;
                    else if (_InnerObject is Hierarchy)
                        return ((Hierarchy)_InnerObject).UniqueName;
                    else if (_InnerObject is Level)
                        return ((Level)_InnerObject).UniqueName;
                    else if (_InnerObject is Kpi)
                        return ((Kpi)_InnerObject).Name;
                    else if (_InnerObject is NamedSet)
                        return ((NamedSet)_InnerObject).Name;
                    else if (_InnerObject is Member)
                        return ((Member)_InnerObject).UniqueName;
                    throw new Exception("Unexpected InnerObject type");
                }
            }

            public bool IsFieldListField
            {
                get
                {
                    if (_InnerObject is Measure)
                        return true;
                    else if (_InnerObject is Dimension)
                        return true;
                    else if (_InnerObject is Hierarchy)
                        return true;
                    else if (_InnerObject is Level)
                        return true;
                    else if (_InnerObject is NamedSet)
                        return true;
                    else if (_InnerObject is Kpi)
                        return true;
                    return false;
                }
            }
        }
    }
}
