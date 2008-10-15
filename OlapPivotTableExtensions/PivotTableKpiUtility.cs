using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.AnalysisServices.AdomdClient;

namespace OlapPivotTableExtensions
{
    public class PivotTableKpiUtility
    {
        private class IconSetDefinition
        {
            public IconSetDefinition() { }
            public IconSetDefinition(Excel.XlIconSet IconSet, bool Reverse, double[] ValueBoundaries)
            {
                this.IconSet = IconSet;
                this.Reverse = Reverse;
                this.ValueBoundaries = ValueBoundaries;
            }

            public Excel.XlIconSet IconSet = Excel.XlIconSet.xl5Arrows;
            public bool Reverse = false;
            public double[] ValueBoundaries = new double[] { };
        }

        private static Dictionary<string, IconSetDefinition> _dictIconSetLookup = new Dictionary<string, IconSetDefinition>(StringComparer.CurrentCultureIgnoreCase);
        static PivotTableKpiUtility()
        {
            //Excel appears to hardcode the mapping between the SSAS icons and the Excel icons
            _dictIconSetLookup.Add("Standard Arrow", new IconSetDefinition(Excel.XlIconSet.xl5ArrowsGray, false, new double[] { -0.5, -0.01, 0.01, 0.5 }));
            _dictIconSetLookup.Add("Status Arrow - Ascending", new IconSetDefinition(Excel.XlIconSet.xl5Arrows, false, new double[] { -0.5, -0.01, 0.01, 0.5 }));
            _dictIconSetLookup.Add("Status Arrow - Descending", new IconSetDefinition(Excel.XlIconSet.xl5Arrows, true, new double[] { -0.5, -0.01, 0.01, 0.5 }));
            _dictIconSetLookup.Add("Gauge - Ascending", new IconSetDefinition(Excel.XlIconSet.xl5Quarters, false, new double[] { -0.5, -0.01, 0.01, 0.5 }));
            _dictIconSetLookup.Add("Gauge - Descending", new IconSetDefinition(Excel.XlIconSet.xl5Quarters, true, new double[] { -0.5, -0.01, 0.01, 0.5 }));
            _dictIconSetLookup.Add("Shapes", new IconSetDefinition(Excel.XlIconSet.xl3Symbols, false, new double[] { -0.5, 0.5 }));
            _dictIconSetLookup.Add("Thermometer", new IconSetDefinition(Excel.XlIconSet.xl3Symbols, false, new double[] { -0.5, 0.5 }));
            _dictIconSetLookup.Add("Variance Arrow", new IconSetDefinition(Excel.XlIconSet.xl3Arrows, false, new double[] { -0.5, 0.5 }));
            _dictIconSetLookup.Add("Road Signs", new IconSetDefinition(Excel.XlIconSet.xl3Signs, false, new double[] { -0.5, 0.5 }));
            _dictIconSetLookup.Add("Cylinder", new IconSetDefinition(Excel.XlIconSet.xl3Signs, false, new double[] { -0.5, 0.5 }));
            _dictIconSetLookup.Add("Smiley Face", new IconSetDefinition(Excel.XlIconSet.xl3Signs, false, new double[] { -0.5, 0.5 }));
            _dictIconSetLookup.Add("Traffic Light", new IconSetDefinition(Excel.XlIconSet.xl3TrafficLights2, false, new double[] { -0.5, 0.5 }));
        }

        public static void AddKpiToPivotTable(Kpi k, Excel.PivotTable pvt)
        {
            foreach (string sKpiPart in new string[] { "KPI_VALUE", "KPI_GOAL", "KPI_STATUS", "KPI_TREND" })
            {
                string sKpiMeasure = Convert.ToString(k.Properties[sKpiPart].Value);
                if (string.IsNullOrEmpty(sKpiMeasure)) continue;

                Excel.CubeField field = pvt.CubeFields.get_Item(sKpiMeasure);
                if (field.Orientation == Excel.XlPivotFieldOrientation.xlDataField) continue;
                field.Orientation = Excel.XlPivotFieldOrientation.xlDataField;

                if (sKpiPart == "KPI_STATUS" || sKpiPart == "KPI_TREND")
                {
                    Excel.PivotItem pi = (Excel.PivotItem)pvt.DataPivotField.PivotItems(sKpiMeasure);
                    Excel.Range range = pi.DataRange;
                    Excel.IconSetCondition iconSet = (Excel.IconSetCondition)range.FormatConditions.AddIconSetCondition();
                    
                    string sStatusGraphic = (sKpiPart == "KPI_STATUS" ? k.StatusGraphic : k.TrendGraphic);
                    IconSetDefinition def = new IconSetDefinition();
                    if (_dictIconSetLookup.ContainsKey(sStatusGraphic))
                        def = _dictIconSetLookup[sStatusGraphic];
                    else
                        System.Windows.Forms.MessageBox.Show("Status graphic type " + sStatusGraphic + " not expected. Please contact the authors of OLAP PivotTable Extensions on the About tab.", "OLAP PivotTable Extensions");

                    iconSet.IconSet = pvt.Application.ActiveWorkbook.IconSets[def.IconSet];
                    try
                    {
                        iconSet.ScopeType = Microsoft.Office.Interop.Excel.XlPivotConditionScope.xlDataFieldScope;
                    }
                    catch { }
                    iconSet.ShowIconOnly = true;
                    iconSet.ReverseOrder = def.Reverse;

                    int i = 2;
                    foreach (double d in def.ValueBoundaries)
                    {
                        Excel.IconCriterion crit = iconSet.IconCriteria[i++];
                        crit.Type = Excel.XlConditionValueTypes.xlConditionValueNumber;
                        crit.Value = d;
                        crit.Operator = (int)(Excel.XlFormatConditionOperator.xlGreaterEqual);
                    }
                }
            }
        }
    }
}
