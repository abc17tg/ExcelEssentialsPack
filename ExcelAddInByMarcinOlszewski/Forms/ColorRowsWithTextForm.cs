using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelAddInByMarcinOlszewski.Scripts;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInByMarcinOlszewski.Forms
{
    public partial class ColorCellsWithTextForm : Form
    {
        public Color Color = Color.White;
        public Type RangeType = Type.Rows;
        private Excel.Application m_app;

        public ColorCellsWithTextForm(Excel.Application app, Type type)
        {
            InitializeComponent();
            m_app = app;
            RangeType = type;
            foreach (var cT in new List<TextBox>() { redTextBox, greenTextBox, blueTextBox })
                cT.KeyPress += colorTextBox_KeyPress;
            Color = Color.PaleGreen;
            redTextBox.Text = this.Color.R.ToString();
            greenTextBox.Text = this.Color.G.ToString();
            blueTextBox.Text = this.Color.B.ToString();
            searchWordTextBox.Focus();
        }

        private void colorPictureBox_Click(object sender, EventArgs e)
        {
            DialogResult result = colorDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                redTextBox.Text = colorDialog.Color.R.ToString();
                greenTextBox.Text = colorDialog.Color.G.ToString();
                blueTextBox.Text = colorDialog.Color.B.ToString();
            }
        }

        private void colorTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void redTextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox tb = sender as TextBox;
            int val;
            if (int.TryParse(tb.Text, out val))
            {
                tb.Text = Math.Max(0, Math.Min(255, val)).ToString();
                ColorChange();
            }
        }

        private void greenTextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox tb = sender as TextBox;
            int val;
            if (int.TryParse(tb.Text, out val))
            {
                tb.Text = Math.Max(0, Math.Min(255, val)).ToString();
                ColorChange();
            }
        }

        private void blueTextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox tb = sender as TextBox;
            int val;
            if (int.TryParse(tb.Text, out val))
            {
                tb.Text = Math.Max(0, Math.Min(255, val)).ToString();
                ColorChange();
            }
        }

        private void ColorChange()
        {
            byte r, g, b;
            if (byte.TryParse(redTextBox.Text, out r) && byte.TryParse(greenTextBox.Text, out g) && byte.TryParse(blueTextBox.Text, out b))
                Color = Color.FromArgb(r, g, b);

            if (Color.GetBrightness() > 0.5)
                invertFontColorCheckBox.Checked = false;
            else
                invertFontColorCheckBox.Checked = true;

            // Create a new Bitmap object
            Bitmap bmp = new Bitmap(colorPictureBox.Width, colorPictureBox.Height);

            // Create a Graphics object
            using (Graphics gfx = Graphics.FromImage(bmp))
            {
                // Use the FillRectangle method to fill the bitmap
                gfx.FillRectangle(new SolidBrush(Color), new Rectangle(0, 0, bmp.Width, bmp.Height));
            }

            // Set the PictureBox's Image property to the bitmap
            colorPictureBox.Image = bmp;
        }


        private void searchWordTextBox_TextChanged(object sender, EventArgs e)
        {
            if (searchWordTextBox.Text != "Search word" && searchWordTextBox.Text != string.Empty)
                okBtn.Enabled = true;
            else
                okBtn.Enabled = false;
        }

        private void okBtn_Click(object sender, EventArgs e)
        {
            Excel.Range rng = m_app.ActiveWindow.RangeSelection;
            ColorRange(rng, RangeType);
            this.Close();
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void searchWordTextBox_Enter(object sender, EventArgs e)
        {
            if (searchWordTextBox.Text == "Search word")
                searchWordTextBox.Text = string.Empty;
        }

        private void searchWordTextBox_Leave(object sender, EventArgs e)
        {
            if (searchWordTextBox.Text == string.Empty)
                searchWordTextBox.Text = "Search word";
        }

        private void searchWordTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return && searchWordTextBox.Text != "Search word" && searchWordTextBox.Text != string.Empty)
                okBtn.PerformClick();
        }

        private void ColorRange(Excel.Range rng, Type type)
        {
            if (!rng.Valid())
                return;
            using (new ExcelExecutionBlock(m_app))
            {
                switch (type)
                {
                    case Type.Rows:
                        foreach (var row in rng.Rows.Cast<Excel.Range>())
                        {
                            if (row.Cells.Cast<Excel.Range>().Any(p => p.Value2 != null && (p.Value2.ToString().Contains(searchWordTextBox.Text) || p.Text.ToString().Contains(searchWordTextBox.Text))))
                            {
                                row.Interior.Color = Color;
                                if (invertFontColorCheckBox.Checked)
                                    row.Font.Color = ColorTranslator.FromOle((int)(row.Cells[1, 1] as Excel.Range).Font.Color).Invert();
                            }
                        }
                        break;
                    case Type.Colums:
                        foreach (var col in rng.Columns.Cast<Excel.Range>())
                        {
                            if (col.Cells.Cast<Excel.Range>().Any(p => p.Value2 != null && (p.Value2.ToString().Contains(searchWordTextBox.Text) || p.Text.ToString().Contains(searchWordTextBox.Text))))
                            {
                                col.Interior.Color = Color;
                                if (invertFontColorCheckBox.Checked)
                                    col.Font.Color = ColorTranslator.FromOle((int)(col.Cells[1, 1] as Excel.Range).Font.Color).Invert();
                            }
                        }
                        break;
                    case Type.Cells:
                        foreach (var cell in rng.Cells.Cast<Excel.Range>())
                        {
                            if (cell.Value2 != null && (cell.Value2.ToString().Contains(searchWordTextBox.Text) || cell.Text.ToString().Contains(searchWordTextBox.Text)))
                            {
                                cell.Interior.Color = Color;
                                if (invertFontColorCheckBox.Checked)
                                    cell.Font.Color = ColorTranslator.FromOle((int)(cell.Cells[1, 1] as Excel.Range).Font.Color).Invert();
                            }
                        }
                        break;
                    default:
                        return;
                }
            }
        }

        public enum Type
        {
            Cells,
            Colums,
            Rows
        }
    }
}
