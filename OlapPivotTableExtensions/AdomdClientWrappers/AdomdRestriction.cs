using System;
using System.Collections.Generic;
using System.Text;

namespace OlapPivotTableExtensions.AdomdClientWrappers
{
    public class AdomdRestriction
    {
        public AdomdRestriction(string name, object restrictionValue)
        {
            Name = name;
            Value = restrictionValue;
        }

        public string Name { get; set; }
        public object Value { get; set; }
    }

    public class AdomdRestrictionCollection : List<AdomdRestriction>
    {
    }

}
