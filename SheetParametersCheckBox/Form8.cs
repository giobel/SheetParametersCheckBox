using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SheetParametersCheckBox
{
    public partial class Form8 : Form
    {
        public string NewSheetNoText { get; set; }

        public bool Sheetparams { get; set; }
        public Form8(string currentViewText)
        {
            InitializeComponent();
            labelCurrentView.Text = currentViewText;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            NewSheetNoText = tboxNewSheetNoText.Text;
            Sheetparams = checkBox1.Checked;
        }
    }
}
