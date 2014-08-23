using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OlapPivotTableExtensions
{
    public partial class LevelChooserForm : Form
    {
        private Excel.PivotTable PivotTable;
        public LevelChooserForm()
        {
            InitializeComponent();
        }

        public LevelChooserForm(Excel.CubeField cubeField, Excel.PivotTable pt)
        {
            PivotTable = pt;
            InitializeComponent();
            this.chkLevels.Items.Clear();
            //cubeField.CreatePivotFields(); //shouldn't be necessary since it's already in the PivotTable
            foreach (Excel.PivotField pf in cubeField.PivotFields)
            {
                if (pf.IsMemberProperty) continue;
                this.chkLevels.Items.Add(new LevelContainer(pf), !pf.Hidden);
            }
        }

        private class LevelContainer
        {
            private string _Caption; //cache this
            public LevelContainer(Excel.PivotField pf)
            {
                PivotField = pf;
                _Caption = pf.Caption;
            }
            public Excel.PivotField PivotField;
            public override string ToString()
            {
                return _Caption;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            int iLevel = 0;
            try
            {
                bool bFoundCheckedLevel = false;
                PivotTable.ManualUpdate = true;

                //first make levels visible so that at least one level will be visible
                for (int i = 0; i < chkLevels.Items.Count; i++)
                {
                    iLevel = i + 1;
                    LevelContainer lc = (LevelContainer)chkLevels.Items[i];
                    bool bHidden = !chkLevels.GetItemChecked(i);
                    if (!bHidden)
                    {
                        bFoundCheckedLevel = true;
                        lc.PivotField.Hidden = bHidden;
                    }
                    if (!bFoundCheckedLevel
                        && !lc.PivotField.Hidden
                        && i + 1 < chkLevels.Items.Count) //don't drilldown the last level
                    {
                        lc.PivotField.DrilledDown = true; //drill down any hidden levels above the first visible level
                    }
                }

                //second make levels hidden
                for (int i = 0; i < chkLevels.Items.Count; i++)
                {
                    iLevel = i + 1;
                    LevelContainer lc = (LevelContainer)chkLevels.Items[i];
                    bool bHidden = !chkLevels.GetItemChecked(i);
                    if (bHidden)
                    {
                        lc.PivotField.Hidden = bHidden;
                    }
                }

                PivotTable.ManualUpdate = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error on level " + iLevel + " when clicking OK:\r\n" + ex.Message + "\r\n" + ex.StackTrace, "OLAP PivotTable Extensions");
                PivotTable.ManualUpdate = false;
            }
        }
    }
}
