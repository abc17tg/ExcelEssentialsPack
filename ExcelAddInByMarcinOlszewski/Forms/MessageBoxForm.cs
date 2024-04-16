using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddInByMarcinOlszewski.Forms
{
    public partial class MessageBoxForm : Form
    {
        public MessageBoxForm(string message,string title = "Message", bool topMost = false)
        {
            InitializeComponent();
            messageRichTextBox.Text = message;
            this.Text = title;
            this.TopMost = topMost;
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
