using System;
using System.Collections.Generic;
using System.Text;

using OlapPivotTableExtensions.AdomdClientWrappers;
//using Microsoft.AnalysisServices.AdomdClient; //if you're adding this code to your own app, then just uncomment this line and comment out the above line
using LevelTypeEnum = Microsoft.AnalysisServices.AdomdClient.LevelTypeEnum;
using HierarchyOrigin = Microsoft.AnalysisServices.AdomdClient.HierarchyOrigin;

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
        private string _searchStringOrStrings;
        private int _searchTermCount;
        private bool _exactMatch;
        private bool _searchMemberProperties = false;
        private string _searchOnly;
        private bool _searchOnlyIsLevel = false;
        private int _totalTaskCount;
        private int _completedTaskCount;
        private BackgroundWorker _thread;
        private Exception _error;
        private SortableList<CubeSearchMatch> _listMatches;
        private System.Collections.Hashtable _hashMatchedUniqueNames = new System.Collections.Hashtable();
        private AdomdCommand cmd;
        private bool _Complete = false;
        private System.Windows.Forms.Control _checkInvokeRequired;
        private static Dictionary<string, string> _measureGroupCaptions;
        private List<string> _listMatchedSearchTerms = new List<string>();
        
        private bool _impersonate;
        private string _username;
        private string _domain;
        private string _password;

        //private static int _processorCount = Environment.ProcessorCount;

        private static CubeSearcher _searchOptimizationsCubeSearcher;
        private static bool _caseInsensitive = true;
        private static string _lcaseMDX = "LCase";
        private static float _searchOnServerVsClientRatio = 0.8f;

        /// <summary>
        /// Call this static function as soon as you know what cube you're searching
        /// It will asynchronously run a few test queries to check for case insensitivity and compare searching on the server vs. searching on the client (i.e. copying the members to the client and searching in .NET)
        /// </summary>
        /// <param name="Cube"></param>
        public static void SetupSearchOptimizationsAsync(CubeDef Cube, bool bImpersonate, string sUsername, string sDomain, string sPassword)
        {
            _searchOptimizationsCubeSearcher = new CubeSearcher(Cube, bImpersonate, sUsername, sDomain, sPassword);

            _searchOptimizationsCubeSearcher._thread = new BackgroundWorker();
            _searchOptimizationsCubeSearcher._thread.WorkerSupportsCancellation = true;
            _searchOptimizationsCubeSearcher._thread.DoWork += new DoWorkEventHandler(_searchOptimizationsCubeSearcher._thread_DoSearchOptimizationsWork);
            _searchOptimizationsCubeSearcher._thread.RunWorkerAsync();
        }

        private CubeSearcher(CubeDef Cube, bool bImpersonate, string sUsername, string sDomain, string sPassword)
        {
            _cube = Cube;
            _impersonate = bImpersonate;
            _username = sUsername;
            _domain = sDomain;
            _password = sPassword;
        }

        private void _thread_DoSearchOptimizationsWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                cmd = new AdomdCommand();
                cmd.Connection = _cube.ParentConnection;

                string cubeName = _cube.Name;

                //apparently the collation on the cube itself controls whether all string comparisons done in MDX queries against that cube are case sensitive or not... apparently it doesn't matter about the individual collation of dimensions or column bindings... at least that's what my research observed
                cmd.CommandText = "with member [Measures].[IsCaseInsensitive] as \"X\"=\"x\" select [Measures].[IsCaseInsensitive] on 0 from [" + cubeName + "]";
                CellSet cs = cmd.ExecuteCellSet();
                _caseInsensitive = Convert.ToBoolean(cs.Cells[0].Value);

                if (_caseInsensitive)
                    _lcaseMDX = string.Empty;
                else
                    _lcaseMDX = "LCase";

                _searchOnServerVsClientRatio = 0.8f; //default

                if (_thread.CancellationPending) return;

                //sort dimensions based on cardinality
                List<Dimension> listDimensions = GetCardinalitySortedDimensions();

                //find the first dimension that's over 20000 members which should give us a good idea of search performance
                Dimension dimensionToTest = null;
                Level levelToTest = null;
                foreach (Dimension d in listDimensions)
                {
                    if (((uint)d.Properties["DIMENSION_CARDINALITY"].Value) > 20000)
                    {
                        dimensionToTest = d;
                        break;
                    }
                }
                if (dimensionToTest == null)
                {
                    //if there's not a large enough dimension then pick the largest
                    dimensionToTest = listDimensions[listDimensions.Count - 1];
                }
                foreach (Hierarchy h in dimensionToTest.Hierarchies)
                {
                    if (((uint)h.Properties["HIERARCHY_CARDINALITY"].Value) > 20000)
                    {
                        levelToTest = h.Levels[h.Levels.Count - 1];
                        break;
                    }
                }
                if (levelToTest == null)
                {
                    levelToTest = dimensionToTest.Hierarchies[0].Levels[dimensionToTest.Hierarchies[0].Levels.Count - 1];
                }

                System.Diagnostics.Stopwatch timer = new System.Diagnostics.Stopwatch();
                timer.Start();

                string sSearchStringTest = "OlapPivotTableExtensionsSearchOptimizationTest";
                cmd.CommandText = "select {} on 0, Filter(" + levelToTest.UniqueName + ".AllMembers, InStr(" + _lcaseMDX + "(" + levelToTest.ParentHierarchy.UniqueName + ".CurrentMember.Member_Caption), @SearchString) > 0) properties MEMBER_TYPE on 1 from [" + cubeName + "]";
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new AdomdParameter("SearchString", sSearchStringTest));

                if (_thread.CancellationPending) return;

                cs = cmd.ExecuteCellSet();

                if (_thread.CancellationPending) return;

                timer.Stop();
                long lngSearchOnServerTicks = timer.ElapsedTicks;
                timer.Reset();


                //test downloading all members to the client and doing the searching on the client
                timer.Start();
                cmd.CommandText = "with member [Measures].[OlapPivotTableExtensionsNull] as null select {[Measures].[OlapPivotTableExtensionsNull]} on 0, " + levelToTest.UniqueName + ".AllMembers dimension properties " + levelToTest.UniqueName + ".[MEMBER_UNIQUE_NAME], " + levelToTest.UniqueName + ".[MEMBER_CAPTION] on 1 from [" + _cube.Name + "]";
                cmd.Parameters.Clear();
                AdomdDataReader reader = cmd.ExecuteReader();
                int iCnt = 0;
                while (reader.Read())
                {
                    if (_thread.CancellationPending)
                    {
                        reader.Close();
                        return;
                    }
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        string sValue = Convert.ToString(reader[i]);
                        if (sValue != null && sValue.IndexOf(sSearchStringTest, 0, StringComparison.CurrentCultureIgnoreCase) > 0)
                        {
                            iCnt++;
                        }
                    }
                }
                reader.Close();

                timer.Stop();
                long lngSearchOnClientTicks = timer.ElapsedTicks;

                _searchOnServerVsClientRatio = ((float)lngSearchOnServerTicks) / ((float)lngSearchOnClientTicks);

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
            }
        }






        public CubeSearcher(CubeDef Cube, CubeSearchScope Scope, string SearchString, bool ExactMatch, bool SearchMemberProperties, string SearchOnly, System.Windows.Forms.Control ConsumingControl, bool bImpersonate, string sUsername, string sDomain, string sPassword)
        {
            _cube = Cube;
            _scope = Scope;
            _searchStringOrStrings = SearchString;
            _exactMatch = ExactMatch;
            _searchMemberProperties = SearchMemberProperties;
            _searchOnly = SearchOnly;
            _checkInvokeRequired = ConsumingControl;
            _impersonate = bImpersonate;
            _username = sUsername;
            _domain = sDomain;
            _password = sPassword;

            if (_searchOnly != null)
            {
                _searchOnly = _searchOnly.Trim();
                if (_searchOnly.Split(new string[] { "].[" }, StringSplitOptions.None).Length == 3)
                {
                    _searchOnlyIsLevel = true;
                }
            }
        }

        private List<Dimension> GetCardinalitySortedDimensions()
        {
            List<Dimension> listDimensions = new List<Dimension>(_cube.Dimensions.Count);
            foreach (Dimension d in _cube.Dimensions)
            {
                if (d.UniqueName.ToLower().StartsWith("[measures]")) continue;
                listDimensions.Add(d);
            }
            listDimensions.Sort(delegate(Dimension x, Dimension y) { return ((uint)x.Properties["DIMENSION_CARDINALITY"].Value).CompareTo((uint)y.Properties["DIMENSION_CARDINALITY"].Value); });
            return listDimensions;
        }

        public SortableList<CubeSearchMatch> Matches
        {
            get { return _listMatches; }
        }

        public void SearchAsync()
        {
            _thread = new BackgroundWorker();
            _thread.WorkerSupportsCancellation = true;

            try
            {
                _searchOptimizationsCubeSearcher.Cancel(); //just in case it's still running, cancel it
                //TODO: test this further
                while (!_searchOptimizationsCubeSearcher.Complete)
                {
                    System.Threading.Thread.Sleep(100);
                    if (_thread.CancellationPending) return;
                }
            }
            catch { }

            _listMatches = new SortableList<CubeSearchMatch>();

            _error = null;
            _completedTaskCount = 0;
            _totalTaskCount = 1;
            _measureGroupCaptions = new Dictionary<string, string>();

            _thread.DoWork += new DoWorkEventHandler(_thread_DoWork);
            _thread.RunWorkerAsync();
        }

        public void Cancel()
        {
            try
            {
                _thread.CancelAsync();
            }
            catch { }

            try
            {
                if (cmd != null)
                {
                    if (_impersonate)
                    {
                        using (new Impersonator(_username, _domain, _password))
                        {
                            cmd.Cancel(); //cancel opens a new connection so must be done under impersonation
                        }
                    }
                    else
                    {
                        cmd.Cancel();
                    }
                }
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

        public string[] MatchedSearchTerms
        {
            get { return _listMatchedSearchTerms.ToArray(); }
        }

        public int SearchTermCount
        {
            get { return _searchTermCount; }
        }

        private delegate void AddMatch_Delegate(CubeSearchMatch match);
        private void AddMatch(CubeSearchMatch match)
        {
            try
            {
                bool bExistsInHashtable = _hashMatchedUniqueNames.ContainsKey(match.UniqueName);
                if (bExistsInHashtable //searching hashtable for the unique name should quickly eliminate most new matches and short circuit
                && _listMatches.Contains(match))
                {
                    return; //don't add duplicates which can happen with multiple search terms that find the same match
                }
                if (_checkInvokeRequired != null && _checkInvokeRequired.InvokeRequired)
                {
                    //avoid the "cross-thread operation not valid" error message
                    //since a control is using this list as a BindingSource, we have to update the list this way
                    _checkInvokeRequired.Invoke(new AddMatch_Delegate(AddMatch), new object[] { match });
                }
                else
                {
                    _listMatches.Add(match);
                    if (!bExistsInHashtable)
                        _hashMatchedUniqueNames.Add(match.UniqueName, null);
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

        private bool IsExactMatch(string str, string[] searchTerms)
        {
            bool bMatch = false;
            foreach (string searchTerm in searchTerms)
            {
                if (string.Compare(str, searchTerm, true) == 0)
                {
                    AddSearchTermMatch(searchTerm);
                    bMatch = true;
                    //don't break as we want to see if other search terms also match this item so that we will report they match on the results screen
                }
            }
            return bMatch;
        }

        private bool IsPartialMatch(string str, string[] searchTerms)
        {
            bool bMatch = false;
            foreach (string searchTerm in searchTerms)
            {
                if (str.IndexOf(searchTerm, 0, StringComparison.CurrentCultureIgnoreCase) >= 0)
                {
                    AddSearchTermMatch(searchTerm);
                    bMatch = true;
                    //don't break as we want to see if other search terms also match this item so that we will report they match on the results screen
                }
            }
            return bMatch;
        }

        private void AddSearchTermMatch(string searchTerm)
        {
            if (!_listMatchedSearchTerms.Contains(searchTerm))
                _listMatchedSearchTerms.Add(searchTerm);
        }

        public event ProgressChangedEventHandler ProgressChanged;

        private void _thread_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                _Complete = false;
                _totalTaskCount = 0;

                if (_cube.ParentConnection.State != System.Data.ConnectionState.Open)
                {
                    if (_impersonate)
                    {
                        using (new Impersonator(_username, _domain, _password))
                        {
                            _cube.ParentConnection.Open();
                        }
                    }
                    else
                    {
                        _cube.ParentConnection.Open();
                    }
                }

                if (string.IsNullOrEmpty(_searchOnly))
                {
                    foreach (Dimension d in _cube.Dimensions)
                        _totalTaskCount += d.Hierarchies.Count;
                }
                else
                {
                    _totalTaskCount = 1;
                }

                string _searchStrings = this._searchStringOrStrings;
                _searchStrings = _searchStrings.Replace("\r\n", "\n").Replace('\r', '\n');
                string[] _searchStringArray = _searchStrings.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                _searchTermCount = _searchStringArray.Length;

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

                List<string> listFoundMemberUniqueNames = new List<string>();

                if (_scope == CubeSearchScope.FieldList || _scope == CubeSearchScope.MeasuresCaptionOnly)
                {
                    ////////////////////////////////////////////////////////////////////
                    // SEARCH FIELD LIST
                    ////////////////////////////////////////////////////////////////////

                    if (GetSSASServerVersion() >= 2005 && _scope != CubeSearchScope.MeasuresCaptionOnly)
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
                            if (IsExactMatch(m.Caption, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(m));
                            }
                        }
                        else
                        {
                            if (IsPartialMatch(m.Caption, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(m));
                            }
                            else if (_scope != CubeSearchScope.MeasuresCaptionOnly && IsPartialMatch(m.Description, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(m));
                            }
                            else if (_scope != CubeSearchScope.MeasuresCaptionOnly && IsPartialMatch(m.DisplayFolder, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(m));
                            }
                            else if (_scope != CubeSearchScope.MeasuresCaptionOnly)
                            {
                                //search the measure group caption
                                string sMeasureGroup = Convert.ToString(m.Properties["MEASUREGROUP_NAME"].Value);
                                if (!string.IsNullOrEmpty(sMeasureGroup))
                                {
                                    string sMeasureGroupCaption = CubeSearcher.GetMeasureGroupCaption(m.ParentCube.ParentConnection.Database, m.ParentCube.Name, sMeasureGroup);
                                    sMeasureGroup = (sMeasureGroupCaption != null ? sMeasureGroupCaption : sMeasureGroup);
                                    if (IsPartialMatch(sMeasureGroup, _searchStringArray))
                                    {
                                        AddMatch(new CubeSearchMatch(m));
                                    }
                                }
                            }
                        }
                    }
                    _completedTaskCount++;
                    ProgressChanged.Invoke(this, new ProgressChangedEventArgs(Math.Min((int)(100 * _completedTaskCount / ((double)_totalTaskCount)), 100), null));

                    if (_scope == CubeSearchScope.MeasuresCaptionOnly)
                    {
                        return;
                    }

                    foreach (Dimension d in _cube.Dimensions)
                    {
                        if (d.UniqueName.ToLower().StartsWith("[measures]")) continue; //work item 23021
                        if (_exactMatch)
                        {
                            if (IsExactMatch(d.Caption, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(d));
                            }
                        }
                        else
                        {
                            if (IsPartialMatch(d.Caption, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(d));
                            }
                            else if (IsPartialMatch(d.Description, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(d));
                            }
                        }
                        foreach (Hierarchy h in d.Hierarchies)
                        {
                            if (_thread.CancellationPending) return;
                            if (_exactMatch)
                            {
                                if (IsExactMatch(h.Caption, _searchStringArray))
                                {
                                    AddMatch(new CubeSearchMatch(h));
                                }
                            }
                            else
                            {
                                if (IsPartialMatch(h.Caption, _searchStringArray))
                                {
                                    AddMatch(new CubeSearchMatch(h));
                                }
                                else if (IsPartialMatch(h.Description, _searchStringArray))
                                {
                                    AddMatch(new CubeSearchMatch(h));
                                }
                                else if (IsPartialMatch(h.DisplayFolder, _searchStringArray))
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
                                        if (IsExactMatch(l.Caption, _searchStringArray))
                                        {
                                            AddMatch(new CubeSearchMatch(l));
                                        }
                                    }
                                    else
                                    {
                                        if (IsPartialMatch(l.Caption, _searchStringArray))
                                        {
                                            AddMatch(new CubeSearchMatch(l));
                                        }
                                        else if (IsPartialMatch(l.Description, _searchStringArray))
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
                                            if (IsExactMatch(sPropertyCaption, _searchStringArray))
                                            {
                                                bIsMatch = true;
                                            }
                                        }
                                        else
                                        {
                                            if (IsPartialMatch(sPropertyCaption, _searchStringArray))
                                            {
                                                bIsMatch = true;
                                            }
                                            else if (IsPartialMatch(sDescription, _searchStringArray))
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
                            ProgressChanged.Invoke(this, new ProgressChangedEventArgs(Math.Min((int)(100 * _completedTaskCount / ((double)_totalTaskCount)), 100), null));
                        }
                    }

                    //search KPIs
                    foreach (Kpi k in _cube.Kpis)
                    {
                        if (_exactMatch)
                        {
                            if (IsExactMatch(k.Caption, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(k));
                            }
                        }
                        else
                        {
                            if (IsPartialMatch(k.Caption, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(k));
                            }
                            else if (IsPartialMatch(k.Description, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(k));
                            }
                            else if (IsPartialMatch(k.DisplayFolder, _searchStringArray))
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
                            if (IsExactMatch(sSetCaption, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(s));
                            }
                        }
                        else
                        {
                            if (IsPartialMatch(sSetCaption, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(s));
                            }
                            else if (IsPartialMatch(s.Description, _searchStringArray))
                            {
                                AddMatch(new CubeSearchMatch(s));
                            }
                            else if (IsPartialMatch(sDisplayFolder, _searchStringArray))
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

                    bool bSearchOnServer = true;
                    if (_searchOnServerVsClientRatio * _searchStringArray.Length > 1.2)
                    {
                        bSearchOnServer = false;
                    }
                    else
                    {
                        _totalTaskCount = _totalTaskCount * _searchTermCount;
                    }

                    Hierarchy hierSearchOnly = null;

                    //now do traditional search by looping all hierarchies and executing one MDX query per
                    //put dimensions into structure that can be sorted by dimension size
                    //sort dimensions based on cardinality
                    List<Dimension> listDimensions = new List<Dimension>(_cube.Dimensions.Count);
                    if (string.IsNullOrEmpty(_searchOnly))
                    {
                        listDimensions = GetCardinalitySortedDimensions();
                    }
                    else
                    {
                        foreach (Dimension d in _cube.Dimensions)
                        {
                            foreach (Hierarchy h in d.Hierarchies)
                            {
                                if (string.Compare(h.UniqueName, _searchOnly, true) == 0 || (_searchOnlyIsLevel && _searchOnly.ToLower().StartsWith(h.UniqueName.ToLower())))
                                {
                                    hierSearchOnly = h;
                                    listDimensions.Add(d);
                                    break;
                                }
                            }
                        }
                    }

                    foreach (string _searchString in _searchStringArray)
                    {
                        if (_thread.CancellationPending) return;

                        if (!bSearchOnServer && _searchStringArray.Length > 1 && (!_exactMatch || _searchMemberProperties)) continue; //because we're searching multiple search terms, each MDSCHEMA_MEMBERS call takes some time so may not be worth it if we're searching on the client anyway... don't want to double the total search time by doing this quick exact match search

                        //do quick full cube search for exact match in any dimension... this code uses the name hash index and is very fast (except for ROLAP dimensions)
                        //even if not looking for an exact match, run this code every time because it is so much faster than Filter(AllMembers) function
                        AdomdRestrictionCollection restrictions = new AdomdRestrictionCollection();
                        restrictions.Add(new AdomdRestriction("CATALOG_NAME", _cube.ParentConnection.Database));
                        restrictions.Add(new AdomdRestriction("CUBE_NAME", _cube.Name));
                        if (!string.IsNullOrEmpty(_searchOnly))
                        {
                            if (_searchOnlyIsLevel)
                            {
                                restrictions.Add(new AdomdRestriction("LEVEL_UNIQUE_NAME", _searchOnly));
                            }
                            else
                            {
                                restrictions.Add(new AdomdRestriction("HIERARCHY_UNIQUE_NAME", _searchOnly));
                            }
                        }
                        restrictions.Add(new AdomdRestriction("MEMBER_CAPTION", _searchString));
                        System.Data.DataTable tblExactMatchMembers = _cube.ParentConnection.GetSchemaDataSet("MDSCHEMA_MEMBERS", restrictions).Tables[0];

                        Dictionary<Hierarchy, List<Member>> dictFoundHierarchyMembers = new Dictionary<Hierarchy, List<Member>>();
                        Dictionary<string, string> dictFoundHierarchyMembersString = new Dictionary<string, string>();
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
                                    AddSearchTermMatch(_searchString);
                                    AddMatch(new CubeSearchMatch(m));
                                }
                            }
                            ProgressChanged.Invoke(this, new ProgressChangedEventArgs(Math.Min((int)(100 * _completedTaskCount / ((double)_totalTaskCount)) + 1, 100), null));
                        }

                        if (_thread.CancellationPending) return;
                        ProgressChanged.Invoke(this, new ProgressChangedEventArgs(Math.Min((int)(100 * _completedTaskCount / ((double)_totalTaskCount)) + 1, 100), null));

                        if (!bSearchOnServer) continue; //will search on client below

                        ProgressChanged.Invoke(this, new ProgressChangedEventArgs(Math.Min((int)(100 * _completedTaskCount / ((double)_totalTaskCount)) + 2, 100), null));

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
                                            if (_searchOnlyIsLevel && string.Compare(l.UniqueName, _searchOnly, true) != 0) continue;

                                            string sLevelMembers = l.UniqueName + ".AllMembers";

                                            List<string> listProperties = new List<string>();
                                            string sFilterPropertiesMDX = string.Empty;
                                            string sDimensionPropertiesClause = l.UniqueName + ".[Member_Type]";
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
                                                        sFilterPropertiesMDX += " or " + _lcaseMDX + "(" + h.UniqueName + ".CurrentMember.Properties(\"" + sPropertyName + "\")) = @SearchString";
                                                    else
                                                        sFilterPropertiesMDX += " or InStr(" + _lcaseMDX + "(" + h.UniqueName + ".CurrentMember.Properties(\"" + sPropertyName + "\")), @SearchString) > 0";
                                                    listProperties.Add(sPropertyName);
                                                }
                                            }
                                            if (sDimensionPropertiesClause.Length > 0)
                                                sDimensionPropertiesClause = "dimension properties " + sDimensionPropertiesClause;

                                            if (_exactMatch)
                                                cmd.CommandText = "select {} on 0, Filter(" + sLevelMembers + ", " + _lcaseMDX + "(" + h.UniqueName + ".CurrentMember.Member_Caption) = @SearchString" + sFilterPropertiesMDX + ") " + sDimensionPropertiesClause + " on 1 from [" + _cube.Name + "]";
                                            else
                                                cmd.CommandText = "select {} on 0, Filter(" + sLevelMembers + ", InStr(" + _lcaseMDX + "(" + h.UniqueName + ".CurrentMember.Member_Caption), @SearchString) > 0" + sFilterPropertiesMDX + ") " + sDimensionPropertiesClause + " on 1 from [" + _cube.Name + "]";

                                            if (_thread.CancellationPending) return;

                                            cs = cmd.ExecuteCellSet();

                                            if (_thread.CancellationPending) return;

                                            if (cs != null && cs.Axes.Count > 1)
                                            {
                                                foreach (Position p in cs.Axes[1].Positions)
                                                {
                                                    foreach (Member m in p.Members)
                                                    {
                                                        if (_thread.CancellationPending) return;
                                                        AddSearchTermMatch(_searchString); //mark it as a match even if we've already found this member so that you'll get credit if multiple search terms find the same match
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
                                                            {
                                                                AddMatch(new CubeSearchMatch(m));
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        string sHierarchyMembers = h.UniqueName + ".AllMembers";
                                        if (_searchOnlyIsLevel)
                                        {
                                            sHierarchyMembers = _searchOnly + ".AllMembers";
                                        }

                                        if (_exactMatch)
                                            cmd.CommandText = "select {} on 0, Filter(" + sHierarchyMembers + ", " + _lcaseMDX + "(" + h.UniqueName + ".CurrentMember.Member_Caption) = @SearchString) dimension properties Member_Type on 1 from [" + _cube.Name + "]";
                                        else
                                            cmd.CommandText = "select {} on 0, Filter(" + sHierarchyMembers + ", InStr(" + _lcaseMDX + "(" + h.UniqueName + ".CurrentMember.Member_Caption), @SearchString) > 0) dimension properties Member_Type on 1 from [" + _cube.Name + "]";

                                        if (_thread.CancellationPending) return;

                                        cs = cmd.ExecuteCellSet();

                                        if (_thread.CancellationPending) return;

                                        if (cs != null && cs.Axes.Count > 1)
                                        {
                                            foreach (Position p in cs.Axes[1].Positions)
                                            {
                                                foreach (Member m in p.Members)
                                                {
                                                    if (_thread.CancellationPending) return;
                                                    AddSearchTermMatch(_searchString);
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
                                    ProgressChanged.Invoke(this, new ProgressChangedEventArgs(Math.Min((int)(97 * _completedTaskCount / ((double)_totalTaskCount)) + 2, 100), null));
                                }
                            }
                        }
                    }


                    //////////////////////////////
                    //SEARCH FOR MEMBERS ON CLIENT
                    //////////////////////////////
                    if (!bSearchOnServer && (!_exactMatch || _searchMemberProperties))
                    {

                        //search each hierarchy... start with the smallest dimensions
                        foreach (Dimension d in listDimensions)
                        {
                            foreach (Hierarchy h in d.Hierarchies)
                            {
                                if (_thread.CancellationPending) return;
                                if (hierSearchOnly != null && h != hierSearchOnly) continue;

                                uint iHierarchyMemberCount = ((uint)h.Properties["HIERARCHY_CARDINALITY"].Value);
                                uint iHierarchyMemberCountComplete = 0;
                                List<string> listFoundMembers = new List<string>();
                                Dictionary<string, string> dictFoundMembersViaMemberProperty = new Dictionary<string, string>();
                                string sAllDimensionPropertiesClause = string.Empty;
                                foreach (Level l in h.Levels)
                                {
                                    if (_searchOnlyIsLevel && string.Compare(l.UniqueName, _searchOnly, true) != 0) continue;
                                    if (l.LevelType == LevelTypeEnum.All) continue;

                                    string sLevelMembers = l.UniqueName + ".AllMembers";

                                    List<string> listProperties = new List<string>();
                                    string sFilterPropertiesMDX = string.Empty;
                                    string sDimensionPropertiesClause = l.UniqueName + ".[MEMBER_UNIQUE_NAME]";
                                    sDimensionPropertiesClause += ", " + l.UniqueName + ".[MEMBER_CAPTION]";
                                    foreach (System.Data.DataRow row in tblProperties.Rows)
                                    {
                                        string sLevelUniqueName = Convert.ToString(row["LEVEL_UNIQUE_NAME"]);
                                        if (sLevelUniqueName != l.UniqueName) continue;

                                        string sPropertyName = Convert.ToString(row["PROPERTY_NAME"]);
                                        if (sDimensionPropertiesClause.Length > 0) sDimensionPropertiesClause += ", ";
                                        sDimensionPropertiesClause += sLevelUniqueName + ".[" + sPropertyName + "]";
                                        if (!listProperties.Contains(sPropertyName))
                                        {
                                            listProperties.Add(sPropertyName);
                                        }
                                    }
                                    if (sDimensionPropertiesClause.Length > 0)
                                    {
                                        if (sAllDimensionPropertiesClause.Length > 0) sAllDimensionPropertiesClause += ", ";
                                        sAllDimensionPropertiesClause += sDimensionPropertiesClause;

                                        sDimensionPropertiesClause = "dimension properties " + sDimensionPropertiesClause;
                                    }

                                    cmd.CommandText = "with member [Measures].[OlapPivotTableExtensionsNull] as null select {[Measures].[OlapPivotTableExtensionsNull]} on 0, " + sLevelMembers + " " + sDimensionPropertiesClause + " on 1 from [" + _cube.Name + "]";

                                    if (_thread.CancellationPending) return;

                                    AdomdDataReader reader = cmd.ExecuteReader();

                                    if (_thread.CancellationPending)
                                    {
                                        reader.Close();
                                        return;
                                    }

                                    int iMemberCaptionColumn = 1;
                                    int iUniqueNameColumn = 0;
                                    List<int> listColumnIndexes = new List<int>();
                                    for (int i = 0; i < reader.FieldCount; i++)
                                    {
                                        if (string.Compare(reader.GetName(i), l.UniqueName + ".[MEMBER_UNIQUE_NAME]", true) == 0)
                                        {
                                            iUniqueNameColumn = i;
                                        }
                                        else if (reader.GetName(i).StartsWith(l.UniqueName))
                                        {
                                            listColumnIndexes.Add(i);
                                        }
                                        if (string.Compare(reader.GetName(i), l.UniqueName + ".[MEMBER_CAPTION]", true) == 0)
                                        {
                                            iMemberCaptionColumn = i;
                                        }
                                    }

                                    while (reader.Read())
                                    {
                                        if (_thread.CancellationPending)
                                        {
                                            reader.Close();
                                            return;
                                        }

                                        foreach (int i in listColumnIndexes)
                                        {
                                            if (_exactMatch)
                                            {
                                                if (IsExactMatch(Convert.ToString(reader[i]), _searchStringArray))
                                                {
                                                    listFoundMembers.Add(Convert.ToString(reader[iUniqueNameColumn]));
                                                    if (i != iMemberCaptionColumn)
                                                    {
                                                        string sProperty = reader.GetName(i).Substring(l.UniqueName.Length + 2).Replace("]", "");
                                                        dictFoundMembersViaMemberProperty.Add(Convert.ToString(reader[iUniqueNameColumn]), sProperty);
                                                    }
                                                    break;
                                                }
                                            }
                                            else
                                            {
                                                if (IsPartialMatch(Convert.ToString(reader[i]), _searchStringArray))
                                                {
                                                    listFoundMembers.Add(Convert.ToString(reader[iUniqueNameColumn]));
                                                    if (i != iMemberCaptionColumn)
                                                    {
                                                        string sProperty = reader.GetName(i).Substring(l.UniqueName.Length + 2).Replace("]", "");
                                                        dictFoundMembersViaMemberProperty.Add(Convert.ToString(reader[iUniqueNameColumn]), sProperty);
                                                    }
                                                    break;
                                                }
                                            }
                                        }

                                        iHierarchyMemberCountComplete++;
                                        if (iHierarchyMemberCountComplete % 5000 == 0) //record progress every 5,000 members
                                        {
                                            ProgressChanged.Invoke(this, new ProgressChangedEventArgs(Math.Min((int)(97 * (_completedTaskCount + Math.Min(1.0 * iHierarchyMemberCountComplete / iHierarchyMemberCount, 1)) / ((double)_totalTaskCount)) + 2, 100), null));
                                        }
                                    }
                                    reader.Close();
                                }

                                if (listFoundMembers.Count > 0)
                                {
                                    cmd.CommandText = "select {} on 0, {" + string.Join(", ", listFoundMembers.ToArray()) + "} dimension properties Member_Type, " + sAllDimensionPropertiesClause + " on 1 from [" + _cube.Name + "]";

                                    if (_thread.CancellationPending) return;

                                    CellSet cs = cmd.ExecuteCellSet();

                                    if (_thread.CancellationPending) return;

                                    foreach (Position p in cs.Axes[1].Positions)
                                    {
                                        foreach (Member m in p.Members)
                                        {
                                            if (dictFoundMembersViaMemberProperty.ContainsKey(m.UniqueName))
                                            {
                                                MemberProperty mp = m.MemberProperties[dictFoundMembersViaMemberProperty[m.UniqueName]];
                                                AddMatch(new CubeSearchMatch(m, mp));
                                            }
                                            else
                                            {
                                                AddMatch(new CubeSearchMatch(m));
                                            }
                                        }
                                    }
                                }

                                _completedTaskCount++;
                                ProgressChanged.Invoke(this, new ProgressChangedEventArgs(Math.Min((int)(97 * _completedTaskCount / ((double)_totalTaskCount)) + 2, 100), null));
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
                catch
                {
                }
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
            MeasuresCaptionOnly,
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
            : IEquatable<CubeSearchMatch>
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
                _Folder += "\\" + m.ParentLevel.Caption;
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

            //allows for List.Contains to function properly
            public bool Equals(CubeSearchMatch other)
            {
                if (other == null)
                    return false;

                if (this.Name == other.Name
                    && this.Type == other.Type
                    && this.Folder == other.Folder
                    && this.Description == other.Description
                    && this.UniqueName == this.UniqueName)
                    return true;
                else
                    return false;
            }

            public override bool Equals(Object obj)
            {
                if (obj == null)
                    return false;

                CubeSearchMatch other = obj as CubeSearchMatch;
                if (other == null)
                    return false;
                else
                    return Equals(other);
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
