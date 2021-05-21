using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void bodyText_TextChanged(object sender, EventArgs e)
        {

        }

        private void buttonBold_Click_1(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = bodyText.SelectionFont;
            if (currentFont.Style != FontStyle.Bold)
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Bold);
            }
            else
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void buttonUnderline_Click(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = bodyText.SelectionFont;
            if (currentFont.Style != FontStyle.Underline)
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Underline);
            }
            else
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void buttonItalics_Click(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = bodyText.SelectionFont;
            if (currentFont.Style != FontStyle.Italic)
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Italic);
            }
            else
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void buttonBold1_Click(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = titleText.SelectionFont;
            if (currentFont.Style != FontStyle.Bold)
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Bold);
            }
            else
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void buttonUnderline1_Click(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = titleText.SelectionFont;
            if (currentFont.Style != FontStyle.Underline)
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Underline);
            }
            else
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void buttonItalics1_Click(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = titleText.SelectionFont;
            if (currentFont.Style != FontStyle.Italic)
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Italic);
            }
            else
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void buttonBGColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                bgColor.BackColor = colorDialog1.Color;
            }
        }

        private void buttonFontColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                fontColor.BackColor = colorDialog1.Color;
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonSearch_Click(object sender, EventArgs e)
        {
            string titlePlainText = titleText.Text;
            string tempBodyText = "";
            string bodyPlainText = bodyText.Rtf;
            while (bodyPlainText.IndexOf("\\b") != -1)
            {
                int start = bodyPlainText.IndexOf("\\b") + 2;
                int end = bodyPlainText.IndexOf("\\b0");
                tempBodyText += bodyPlainText.Substring(start, end - start) + '?';
                bodyPlainText = bodyPlainText.Substring(end + 3);
            }
            tempBodyText = tempBodyText.Substring(0, tempBodyText.Length - 1);
            string[] bodyTerms = tempBodyText.Split('?');
            bodyText.Text = bodyTerms.Length.ToString();
        }
    }
}
