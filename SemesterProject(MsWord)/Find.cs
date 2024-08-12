using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SemesterProject_MsWord_
{
    public partial class Find : Form
    {
        public Find()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(checkBox1.Checked==true)
            {
                Form1.matchcase=true;
            }
            else
            {
                Form1.matchcase=false;
            }
            Form1.FindText=txtSearch.Text;
            this.Close();
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            if(txtSearch.Text.Length > 0)
            {
                button1.Enabled= true;
            }
            else
            {
                button1.Enabled= false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form1.FindText = "";
            this.Close();
        }
    }
}
