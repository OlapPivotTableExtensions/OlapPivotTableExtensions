using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace OlapPivotTableExtensions
{
    public partial class MainForm : Form
    {
        private Excel.PivotTable pvt;
        private Excel.Application application;
        private CalculationsLibrary library;

        private int _LibraryComboDividerItemIndex = int.MaxValue;


        public MainForm(Excel.Application app)
        {
            InitializeComponent();

            try
            {
                System.Reflection.AssemblyFileVersionAttribute attrVersion = (System.Reflection.AssemblyFileVersionAttribute)typeof(MainForm).Assembly.GetCustomAttributes(typeof(System.Reflection.AssemblyFileVersionAttribute), true)[0];
                lblVersion.Text = "OLAP PivotTable Extensions v" + attrVersion.Version;

                application = app;
                pvt = app.ActiveCell.PivotTable;

                library = new CalculationsLibrary();
                library.Load();

                FillCalcsDropdown();

                chkShowCalcMembers.Checked = ThisAddIn.ShowCalcMembersByDefault;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                this.Visible = false;
                this.Close();
            }
        }

        private void SetMDX()
        {
            StringBuilder sMdxQuery = new StringBuilder(pvt.MDX);

            //add (session) calculated members to the query so that you can run it from SSMS
            if (pvt.CalculatedMembers.Count > 0)
            {
                StringBuilder sCalcs = new StringBuilder();
                foreach (Excel.CalculatedMember calc in pvt.CalculatedMembers)
                {
                    sCalcs.AppendFormat("MEMBER {0} as {1}\r\n", calc.Name, calc.Formula);
                }
                if (sMdxQuery.ToString().StartsWith("with", StringComparison.CurrentCultureIgnoreCase))
                {
                    sMdxQuery.Insert(5, sCalcs.ToString());
                }
                else
                {
                    sCalcs.Insert(0, "WITH\r\n");
                    sMdxQuery.Insert(0, sCalcs.ToString());
                }
            }

            txtMDX.Text = sMdxQuery.ToString();
            txtMDX.SelectionStart = 0;
            txtMDX.SelectionLength = sMdxQuery.Length;
            txtMDX.Focus();
        }

        private void btnDeleteCalc_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.CalculatedMember oCalcMember = GetCalculatedMember(comboCalcName.Text);
                if (oCalcMember != null)
                {
                    oCalcMember.Delete();
                    pvt.RefreshTable();
                }
                FillCalcsDropdown();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAddCalc_Click(object sender, EventArgs e)
        {
            try
            {
                bool bMeasure = true;
                string sName = comboCalcName.Text;
                string sFormula = txtCalcFormula.Text;
                if (!sName.StartsWith("[") && !sName.StartsWith("[Measures].", StringComparison.CurrentCultureIgnoreCase))
                {
                    sName = "[Measures].[" + sName.Replace("]", "]]") + "]";
                }
                else if (sName.StartsWith("[") && !sName.StartsWith("[Measures].", StringComparison.CurrentCultureIgnoreCase))
                {
                    bMeasure = false;
                }

                try
                {
                    library.AddCalculation(sName, sFormula);
                    library.Save();
                }
                catch (Exception ex)
                {
                    throw new Exception("There was a problem saving this calculation to the library at " + CalculationsLibrary.LibraryPath + ". " + ex.Message, ex);
                }

                Excel.CalculatedMember oCalcMember = GetCalculatedMember(sName);
                if (oCalcMember != null)
                    oCalcMember.Delete();

                try
                {
                    oCalcMember = pvt.CalculatedMembers.Add(sName, sFormula, System.Reflection.Missing.Value, Excel.XlCalculatedMemberType.xlCalculatedMember);
                    if (bMeasure)
                    {
                        pvt.RefreshTable();
                        pvt.CubeFields.get_Item(sName).Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                    }
                    else
                    {
                        pvt.ViewCalculatedMembers = true;
                        pvt.RefreshTable();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("There was a problem creating the calculation:\r\n" + ex.Message);
                }

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an unexpected error creating the calculation:\r\n" + ex.Message);
            }
        }

        private void FillCalcsDropdown()
        {
            comboCalcName.Items.Clear();
            List<string> listCalcs = new List<string>();
            foreach (Excel.CalculatedMember calc in pvt.CalculatedMembers)
            {
                listCalcs.Add(calc.Name);
            }
            listCalcs.Sort();

            foreach (string calc in listCalcs)
            {
                comboCalcName.Items.Add(calc);
            }

            comboCalcName.Items.Add(string.Empty);
            if (library.Calculations.Length > 0)
            {
                _LibraryComboDividerItemIndex = comboCalcName.Items.Add("---CALCULATIONS LIBRARY---");

                foreach (CalculationsLibrary.Calculation c in library.Calculations)
                {
                    comboCalcName.Items.Add(c.Name);
                }
            }

            comboCalcName.Text = string.Empty;
            comboCalcName.Focus();
            txtCalcFormula.Text = string.Empty;
        }

        //returns the calc member if it exists
        private Excel.CalculatedMember GetCalculatedMember(string sName)
        {
            try
            {
                return pvt.CalculatedMembers.get_Item(sName);
            }
            catch
            {
                return null;
            }
        }

        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl.SelectedTab == tabMDX)
            {
                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    SetMDX();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("There was a problem capturing the MDX query for this PivotTable.\r\n" + ex.Message);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void comboCalcName_TextChanged(object sender, EventArgs e)
        {
            if (comboCalcName.SelectedIndex == _LibraryComboDividerItemIndex)
            {
                comboCalcName.Text = string.Empty;
                btnDeleteCalc.Enabled = false;
            }
            else if (comboCalcName.SelectedIndex > _LibraryComboDividerItemIndex)
            {
                CalculationsLibrary.Calculation c = library.GetCalculation(comboCalcName.Text);
                if (c != null)
                {
                    txtCalcFormula.Text = c.Formula;
                }
                btnDeleteCalc.Enabled = false;
            }
            else
            {
                Excel.CalculatedMember oCalcMember = GetCalculatedMember(comboCalcName.Text);
                if (oCalcMember != null)
                {
                    txtCalcFormula.Text = oCalcMember.Formula;
                    btnDeleteCalc.Enabled = true;
                }
                else
                {
                    btnDeleteCalc.Enabled = false;
                }
            }
        }

        private void linkCodeplex_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.codeplex.com/OlapPivotTableExtend");
        }

        private void linkHelp_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.codeplex.com/OlapPivotTableExtend/Wiki/View.aspx?title=Calculations%20Help");
        }

        private void radioExport_CheckedChanged(object sender, EventArgs e)
        {
            if (radioExport.Checked)
            {
                listImportExportCalcs.Items.Clear();
                foreach (CalculationsLibrary.Calculation c in library.Calculations)
                {
                    listImportExportCalcs.Items.Add(c.Name, true);
                }

                listImportExportCalcs.Enabled = true;
                btnImportExportExecute.Enabled = true;
            }
        }

        private void radDelete_CheckedChanged(object sender, EventArgs e)
        {
            if (radDelete.Checked)
            {
                listImportExportCalcs.Items.Clear();
                foreach (CalculationsLibrary.Calculation c in library.Calculations)
                {
                    listImportExportCalcs.Items.Add(c.Name, false);
                }

                listImportExportCalcs.Enabled = true;
                btnImportExportExecute.Enabled = true;
            }
        }

        private void btnImportExportExecute_Click(object sender, EventArgs e)
        {
            try
            {
                if (radImport.Checked)
                {
                    CalculationsLibrary libraryImportExport = new CalculationsLibrary();
                    libraryImportExport.Load(txtImportFilePath.Text);
                    foreach (CalculationsLibrary.Calculation c in libraryImportExport.Calculations)
                    {
                        if (listImportExportCalcs.CheckedItems.Contains(c.Name))
                        {
                            library.AddCalculation(c.Name, c.Formula);
                        }
                    }
                    library.Save();
                }
                else if (radioExport.Checked)
                {
                    CalculationsLibrary libraryImportExport = new CalculationsLibrary();
                    List<CalculationsLibrary.Calculation> calcs = new List<CalculationsLibrary.Calculation>();
                    foreach (CalculationsLibrary.Calculation c in library.Calculations)
                    {
                        if (listImportExportCalcs.CheckedItems.Contains(c.Name))
                        {
                            calcs.Add(c);
                        }
                    }
                    libraryImportExport.Calculations = calcs.ToArray();
                    libraryImportExport.Save(txtExportFilePath.Text);
                    MessageBox.Show("Export completed successfully.");
                    return;
                }
                else if (radDelete.Checked)
                {
                    foreach (CalculationsLibrary.Calculation c in library.Calculations)
                    {
                        if (listImportExportCalcs.CheckedItems.Contains(c.Name))
                        {
                            library.DeleteCalculation(c.Name);
                        }
                    }
                    library.Save();
                }

                FillCalcsDropdown();
                tabControl.SelectedTab = tabCalcs;
                comboCalcName.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnImportFilePath_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Title = "Choose Calculation Library To Import...";
                dlg.Filter = "Calculation Library (*.xml)|*.xml";
                dlg.CheckFileExists = true;
                dlg.Multiselect = false;
                dlg.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    txtImportFilePath.Text = dlg.FileName;
                    radImport_CheckedChanged(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnExportFilePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Export Calculations To...";
            dlg.Filter = "Calculation Library (*.xml)|*.xml";
            dlg.CheckFileExists = false;
            dlg.Multiselect = false;
            dlg.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                this.txtExportFilePath.Text = dlg.FileName;
            }
        }

        private void radImport_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(txtImportFilePath.Text))
                {
                    CalculationsLibrary libraryImportExport = new CalculationsLibrary();
                    libraryImportExport.Load(txtImportFilePath.Text);
                    listImportExportCalcs.Items.Clear();
                    foreach (CalculationsLibrary.Calculation c in libraryImportExport.Calculations)
                    {
                        listImportExportCalcs.Items.Add(c.Name, true);
                    }

                    listImportExportCalcs.Enabled = true;
                    btnImportExportExecute.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was a problem loading that XML file: " + ex.Message);
            }
        }

        private void btnSaveDefaults_Click(object sender, EventArgs e)
        {
            try
            {
                ThisAddIn.ShowCalcMembersByDefault = chkShowCalcMembers.Checked;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


    }
}
