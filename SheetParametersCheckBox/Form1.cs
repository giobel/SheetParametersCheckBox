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

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           checkedListBox1.DataSource = checkedListSource;
        }
    }
}
