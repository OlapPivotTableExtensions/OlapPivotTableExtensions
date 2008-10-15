using System;
using System.Collections.Generic;
using System.Text;

namespace OlapPivotTableExtensions
{
    /// <summary>
    /// Parses a connection string so you can inspect individual properties more easily.
    /// </summary>
    public class ConnectionStringParser
    {
        private Dictionary<string, string> _dictProperties = new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
        public ConnectionStringParser(string ConnectionString)
        {
            bool bInQuotes = false;
            bool bInKey = true;
            string sKey = string.Empty;
            string sValue = string.Empty;
            foreach (char c in ConnectionString.ToCharArray())
            {
                if (bInQuotes)
                {
                    if (c == '"')
                    {
                        bInQuotes = false;
                        if (string.Compare(sKey, "Extended Properties", true) == 0)
                        {
                            ConnectionStringParser extendedPropertiesParser = new ConnectionStringParser(sValue);
                            foreach (string k in extendedPropertiesParser._dictProperties.Keys)
                            {
                                _dictProperties[k] = extendedPropertiesParser._dictProperties[k];
                            }
                        }
                        else
                        {
                            _dictProperties[sKey] = sValue;
                        }
                        sKey = sValue = string.Empty;
                    }
                    else if (bInKey)
                    {
                        throw new Exception("Didn't expect quotes around property name in connection string: " + ConnectionString);
                    }
                    else
                    {
                        sValue += c;
                    }
                }
                else if (c == '"')
                {
                    bInQuotes = true;
                }
                else if (c == '=')
                {
                    if (!bInKey)
                        sValue += c;
                    else
                        bInKey = false;
                }
                else if (c == ';')
                {
                    _dictProperties[sKey] = sValue;
                    sKey = sValue = string.Empty;
                    bInKey = true;
                }
                else if (bInKey)
                {
                    sKey += c;
                }
                else
                {
                    sValue += c;
                }
            }
            if (!bInKey && !string.IsNullOrEmpty(sKey))
                _dictProperties[sKey] = sValue;
        }

        public Dictionary<string, string>.KeyCollection Keys
        {
            get
            {
                return _dictProperties.Keys;
            }
        }

        public bool ContainsKey(string Key)
        {
            return _dictProperties.ContainsKey(Key);
        }

        public string this[string key] {
            get
            {
                if (_dictProperties.ContainsKey(key))
                    return _dictProperties[key];
                return null;
            }
        }
    }
}
