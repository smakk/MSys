using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using System.Text.RegularExpressions;

namespace MSys
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void 口令_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void ShowForm2()
        {
            Form2 form2 = new Form2();
            this.Hide();
            form2.ShowDialog();
            this.Dispose();
        }

        private void DataShowForm2()
        {
            if (!File.Exists(@".\data.xls"))
            {
                DialogResult res = System.Windows.Forms.MessageBox.Show("data.xls不存在，是否按默认数据格式新建data.xls","提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (res == DialogResult.Yes)
                {
                    CommonData.ifFirstLogin = true;
                    ShowForm2();
                    return;
                }
            }
            CommonData.ifFirstLogin = false;
            ShowForm2();
            return;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //MatchCollection mc = Regex.Matches("编号0", @"编?");
            
            string s = textBox1.Text;
            if(s == "root")
            {
                CommonData.rootLogin = true;
                DataShowForm2();
            }
            else if(s == "user")
            {
                CommonData.rootLogin = false;
                DataShowForm2();
            }
            else
            {
                MessageBox.Show("口令错误");
            }
        }
    }
}
