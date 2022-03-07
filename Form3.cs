using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;

namespace MSys
{
    public partial class Form3 : Form
    {
        Person myP;
        int flag;
        //Person form2Peron;
        public Form3()
        {
            InitializeComponent();
        }

        public Form3(PersonInfo personInfo, ref Person p, int flag)
        {
            InitializeComponent();
            InitForm3(personInfo, ref p, flag);
        }

        public void InitForm3(PersonInfo personInfo, ref Person p, int flag)
        {
            int x = this.panel1.Width;
            int y = this.panel1.Height;
            Label lId = new Label();
            lId.Text = "编号";
            lId.Width = this.panel1.Width/2 -40;
            //lId.BackColor = Color.Black;
            //lId.Height = 20;
            lId.Location = new Point(10,10);
            TextBox t = new TextBox();
            t.Width = this.panel1.Width / 2 - 40;
            //t.Height = 20;
            t.Location = new Point(this.panel1.Width / 2 + 10, 10);
            t.Text = p.getPersonInfo(0);
            t.Enabled = false;
            
            this.panel1.Controls.Add(lId);
            this.panel1.Controls.Add(t);
            
            for (int i = 1; i < personInfo.GetInfoNum() - 3; i++)
            {
                Label l = new Label();
                l.Text = (string)personInfo.GetItem(i);
                l.Width = this.panel1.Width / 2 - 40;
                l.Location = new Point(10, 50 * i);
                this.panel1.Controls.Add(l);
                TextBox tTemp = new TextBox();
                tTemp.Enabled = true;
                //tTemp.Text = p.getPersonInfo(i);
                tTemp.Width = this.panel1.Width / 2 - 40;
                tTemp.Location = new Point(this.panel1.Width / 2 + 10, 50 * i);
                tTemp.Text = p.getPersonInfo(i);
                //tTemp.BackColor = Color.Black;
                this.panel1.Controls.Add(tTemp);
            }
            myP = p;
            //写回
            if (flag == 1)
            {
                this.button1.Text = "确认修改";
                //myP = p;
                //this.button1.Click += new System.EventHandler(this.button1_Click2);
            }
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            Form2.ifCreate = false;
        }

        private void Form3_FormClosed(object sender, FormClosedEventArgs e)
        {
            //Form2.ifCreate = false;
        }

        private void Form3_Load_1(object sender, EventArgs e)
        {
            Form2.ifCreate = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int pos = 0;
            foreach (Control ctrl in panel1.Controls)
            {
                if (ctrl is TextBox)
                {
                    myP.setPersonInfo(pos, ctrl.Text);
                    pos++;
                }
            }
            if (this.button1.Text == "确认修改")
            {
                Form2.ifChange = true;
            }
            Form2.ifCreate = true;
            this.Close();
        }
        private void button1_Click2(object sender, EventArgs e)
        {
            
        }
    }
}
