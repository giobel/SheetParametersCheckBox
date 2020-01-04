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
    public partial class Form1 : Form
    {
        //we will set the content of this list from our macro
        public List<string> checkedListSource { get; set; }
        public CheckedListBox.CheckedItemCollection checkedItems { get; set; }

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
           checkedListBox1.DataSource = checkedListSource;
        }
        private void okButton_Click(object sender, EventArgs e)
        {
            checkedItems = checkedListBox1.CheckedItems;
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                CheckAll();
            }
            else
            {
                UncheckAll();
            }
        }
        public void CheckAll()
        {
            for (int i = 0; i <= checkedListBox1.Items.Count - 1; i++)
            {
                //check item
                checkedListBox1.SetItemChecked(i, true);

            }
        }
        public void UncheckAll()
        {
            for (int i = 0; i <= checkedListBox1.Items.Count - 1; i++)
            {
                //check item
                checkedListBox1.SetItemChecked(i, false);
            }
        }

    }
}
