using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace OlapPivotTableExtensions
{
    //concept from http://www.codeproject.com/Articles/22667/Column-Header-CheckBox-in-DataGridView-Winfoms
    class DataGridViewCheckBoxColumnHeaderCell : DataGridViewColumnHeaderCell
    {
        private Rectangle CheckBoxRegion;
        private bool checkAll = false;

        protected override void Paint(Graphics graphics,
            Rectangle clipBounds, Rectangle cellBounds, int rowIndex,
            DataGridViewElementStates dataGridViewElementState,
            object value, object formattedValue, string errorText,
            DataGridViewCellStyle cellStyle, DataGridViewAdvancedBorderStyle advancedBorderStyle,
            DataGridViewPaintParts paintParts)
        {

            base.Paint(graphics, clipBounds, cellBounds, rowIndex, dataGridViewElementState, string.Empty,
                string.Empty, string.Empty, cellStyle, advancedBorderStyle, paintParts);

            //graphics.FillRectangle(new SolidBrush(cellStyle.BackColor), cellBounds); //this wipes out the border... quick workaround

            //left aligned
            //CheckBoxRegion = new Rectangle(
            //    cellBounds.Location.X + 1,
            //    cellBounds.Location.Y + 2,
            //    25, cellBounds.Size.Height - 4);

            //center aligned checkbox
            CheckBoxRegion = new Rectangle(
                Math.Max(cellBounds.Size.Width / 2 - 4, 0),
                cellBounds.Location.Y + 2,
                14, Math.Min(cellBounds.Size.Height - 4, 13));

            if (this.checkAll)
                ControlPaint.DrawCheckBox(graphics, CheckBoxRegion, ButtonState.Checked);
            else
                ControlPaint.DrawCheckBox(graphics, CheckBoxRegion, ButtonState.Normal);

            //Rectangle normalRegion =
            //    new Rectangle(
            //    cellBounds.Location.X + 1 + 25,
            //    cellBounds.Location.Y,
            //    cellBounds.Size.Width - 26,
            //    cellBounds.Size.Height);

            //graphics.DrawString(value.ToString(), cellStyle.Font, new SolidBrush(cellStyle.ForeColor), normalRegion);
        }

        protected override void OnMouseClick(DataGridViewCellMouseEventArgs e)
        {
            //Convert the CheckBoxRegion 
            //Rectangle rec = new Rectangle(new Point(0, 0), this.CheckBoxRegion.Size);
            this.checkAll = !this.checkAll;
            //if (rec.Contains(e.Location))
            //{
            this.DataGridView.Invalidate();
            //}
            base.OnMouseClick(e);
        }

        public bool CheckAll
        {
            get { return this.checkAll; }
            set { this.checkAll = value; }
        }
    }
}
