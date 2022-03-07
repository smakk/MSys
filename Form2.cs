using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Collections;

using System.Text.RegularExpressions;

using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace MSys
{

    public partial class Form2 : System.Windows.Forms.Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        public int currentComboBoxIndex;
        public const int WM_NCLBUTTONDBLCLK = 0xA3;
        const int WM_NCLBUTTONDOWN = 0x00A1;
        const int HTCAPTION = 2;
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_NCLBUTTONDOWN && m.WParam.ToInt32() == HTCAPTION)
                return;
            if (m.Msg == WM_NCLBUTTONDBLCLK)
                return;

            base.WndProc(ref m);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            InitSystem();
            if(CommonData.rootLogin == false)
            {
                this.button3.Enabled = false;
                this.button4.Enabled = false;
                this.button5.Enabled = false;
                this.button6.Enabled = false;
                this.button9.Enabled = false;
            }
            /*
            */
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int curItem = listBox1.SelectedIndex;
            listBox2.Items.Clear();
            if (curItem < 0 || curItem > CommonData.personData.Count)
                return ;
            //更改listbox2
            for (int i = 0; i < CommonData.personInfo.GetInfoNum(); i++)
            {
                listBox2.Items.Add(CommonData.personInfo.GetItem(i) + ": " + ((Person)CommonData.personData[curItem]).getPersonInfo(i));
            }

            //更改图片框
            if (CommonData.personData.Count != 0)
            {
                for (int i = 0; i < CommonData.personInfo.GetPictureNum(); i++)
                {
                    if (CommonFileIfPicture.FileIsPicture(((Person)CommonData.personData[0]).getPersonInfo(CommonData.personInfo.GetInfoNumNoPicture() + i)))
                    {
                        try
                        {
                            this.pictureBox1.Image = Image.FromFile(((Person)CommonData.personData[0]).getPersonInfo(CommonData.personInfo.GetInfoNumNoPicture() + i));
                            this.pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                        }
                        catch
                        {
                            this.pictureBox1.Image = Image.FromFile(@".\picture\nofile.jpeg");
                            MessageBox.Show("找不到图片" + ((Person)CommonData.personData[0]).getPersonInfo(CommonData.personInfo.GetInfoNumNoPicture() + i));
                            //this.Close();
                        }
                    }
                    //Console.WriteLine(((Person)CommonData.personData[0]).getPersonInfo(CommonData.personInfo.GetInfoNumNoPicture() + i));
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (currentComboBoxIndex == comboBox1.SelectedIndex)
                return;
            currentComboBoxIndex = comboBox1.SelectedIndex;
            listBox1.Items.Clear();
            for (int i = 0; i < CommonData.personData.Count; i++)
            {
                listBox1.Items.Add(((Person)CommonData.personData[i]).getPersonInfo(currentComboBoxIndex));
            }
            //MessageBox.Show("212");
        }

        public void FillFile(ArrayList personData, OleDbCommand cmd)
        {
            //OleDbCommand cmd = CommonData.oleConn.CreateCommand();
            foreach (Person p in personData)
            {
                string insertString = @"insert into [案件信息] values(";
                for(int i = 0; i < p.person.Count; i++)
                {
                    insertString += @""""+p.person[i] + @"""";
                    if (i != p.person.Count-1)
                        insertString += ",";
                }
                insertString += ")";
                cmd.CommandText = insertString;
                Console.WriteLine(insertString);
                cmd.ExecuteNonQuery();
                ///CommonData.oleConn.
                //cmd.Dispose();
            }
        }

        //填充主题数据内容，personInfo为有哪些字段，personData为数据内容
        private void FillForm2(PersonInfo personInfo, ArrayList personData)
        {
            //左侧列表栏
            for (int i = 0; i < personData.Count; i++)
            {
                listBox1.Items.Add(((Person)personData[i]).getPersonInfo(0));
            }

            //左侧下拉框
            for (int i = 0; i < personInfo.GetInfoNumNoPicture(); i++)
            {
                comboBox1.Items.Add("按照" + personInfo.GetItem(i) + "排序");
            }
            if(personInfo.GetInfoNum() != 0)
            {
                //SetSelected
                comboBox1.SelectedIndex = 0;
                currentComboBoxIndex = 0;
            }

            //右侧信息栏
            if(personData.Count != 0)
            {
                for (int i = 0; i < personInfo.GetInfoNum(); i++)
                    listBox2.Items.Add(personInfo.GetItem(i) + ": " + ((Person)personData[0]).getPersonInfo(i));
                listBox1.SetSelected(0,true);
                //listBox1.SelectIndex = 0;
            }

            //搜索框的选择栏
            comboBox2.Items.Add("");
            for (int i = 0; i < personInfo.GetInfoNumNoPicture(); i++)
            {
                comboBox2.Items.Add("按照" + personInfo.GetItem(i) + "搜索");
            }

            //右侧图片框
            if(personData.Count != 0)
            {
                for(int i = 0; i < personInfo.GetPictureNum(); i++)
                {
                    if (CommonFileIfPicture.FileIsPicture(((Person)personData[0]).getPersonInfo(personInfo.GetInfoNumNoPicture() + i)))
                    {
                        try
                        {
                            this.pictureBox1.Image = Image.FromFile(((Person)personData[0]).getPersonInfo(personInfo.GetInfoNumNoPicture() + i));
                            this.pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                        }
                        catch
                        {
                            this.pictureBox1.Image = Image.FromFile(@".\picture\nofile.jpeg");
                            MessageBox.Show("找不到图片" + ((Person)personData[0]).getPersonInfo(personInfo.GetInfoNumNoPicture() + i));
                            //this.Close();
                        }
                    }
                    //Console.WriteLine(((Person)personData[0]).getPersonInfo(personInfo.GetInfoNumNoPicture() + i));
                }
            }
            //this.pictureBox2.Paint();
        }

        //initsystemz中的两个部分拆分成这两个文件
        //从文件中读取所有的数据
        private void ReadFromFile()
        {
            CommonData.oleConn.Open();
            //OleDbCommand cmd = CommonData.oleConn.CreateCommand();
            OleDbDataAdapter OleDaExcel = new OleDbDataAdapter("select * from [案件信息]", CommonData.oleConn);
            DataSet OleDsExcle = new DataSet();
            OleDaExcel.Fill(OleDsExcle, "Person");
            foreach (DataRow mDr in OleDsExcle.Tables[0].Rows)
            {
                Person p = new Person();
                foreach (DataColumn mDc in OleDsExcle.Tables[0].Columns)
                {
                    p.Add(mDr[mDc].ToString());

                }
                CommonData.personData.Add(p);

            }
            //System.Data.DataTable table = CommonData.oleConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

            foreach (DataRow mDr in OleDsExcle.Tables[0].Rows)
            {
                foreach (DataColumn mDc in OleDsExcle.Tables[0].Columns)
                {
                    CommonData.personInfo.Add(mDc.ColumnName);
                }
                break;
            }
            //OleDaExcel.
            OleDaExcel.Dispose();
            CommonData.oleConn.Close();
        }

        private void WriteToFile()
        {
            CommonData.oleConn.Open();
            //File.Delete(@".\data.xls");
            FillFile(CommonData.personData, CommonData.oleConn.CreateCommand());
            CommonData.oleConn.Close();
        }

        private void CreateTableFillFile()
        {
            CommonData.oleConn.Open();
            OleDbCommand cmd = CommonData.oleConn.CreateCommand();
            string createTableString = "CREATE TABLE 案件信息 (";
            for (int i = 0; i < CommonData.personInfo.GetInfoNum(); i++)
            {
                createTableString += "[" + CommonData.personInfo.GetItem(i) + "] text";
                if (i != CommonData.personInfo.GetInfoNum() - 1)
                    createTableString += ",";
            }
            createTableString += ")";
            cmd.CommandText = createTableString;
            //Console.WriteLine(createTableString);
            cmd.ExecuteNonQuery();
            FillFile(CommonData.personData, CommonData.oleConn.CreateCommand());
            CommonData.oleConn.Close();
        }

        private void CreateDir()
        {
            Directory.CreateDirectory(@".\picture");
        }

        //初始化全局变量，填充数据
        private void InitSystem()
        {
            if(CommonData.ifFirstLogin == true)
            {
                CommonData.FillData();
                CreateTableFillFile();
                CreateDir();
                //this.pictureBox1.Refresh();

            }
            else
            {
                ReadFromFile();
            }
            FillForm2(CommonData.personInfo, CommonData.personData);
            /*
            CommonData.oleConn.Open();
            //第一次登录需要向其中写入数据，不是第一次则需要向其中读取数据
            if(CommonData.ifFirstLogin == true)
            {
                //需要向其中填入默认数据吗？
                CommonData.FillData();

                //初始化数据文件写入默认数据
                OleDbCommand cmd = CommonData.oleConn.CreateCommand();
                string createTableString = "CREATE TABLE 案件信息 (";
                for (int i = 0; i < CommonData.personInfo.GetInfoNum(); i++)
                {
                    createTableString += "[" + CommonData.personInfo.GetItem(i) + "] text";
                    if (i != CommonData.personInfo.GetInfoNum() - 1)
                        createTableString += ",";
                }
                createTableString += ")";
                cmd.CommandText = createTableString;
                //Console.WriteLine(createTableString);
                cmd.ExecuteNonQuery();
                FillFile(CommonData.personData);
            }
            else
            {
                //从文件中读取数据
                OleDbCommand cmd = CommonData.oleConn.CreateCommand();
                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter("select * from [案件信息]", CommonData.oleConn);
                DataSet OleDsExcle = new DataSet();
                OleDaExcel.Fill(OleDsExcle, "Person");
                foreach (DataRow mDr in OleDsExcle.Tables[0].Rows)
                {
                    Person p = new Person();
                    foreach (DataColumn mDc in OleDsExcle.Tables[0].Columns)
                    {
                        p.Add(mDr[mDc].ToString());
                            
                    }
                    CommonData.personData.Add(p);
                    
                }
                //System.Data.DataTable table = CommonData.oleConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                foreach (DataRow mDr in OleDsExcle.Tables[0].Rows)
                {
                    foreach (DataColumn mDc in OleDsExcle.Tables[0].Columns)
                    {
                        CommonData.personInfo.Add(mDc.ColumnName);
                    }
                    break;
                }
                    //OleDaExcel.
                OleDaExcel.Dispose();
            }
            FillForm2(CommonData.personInfo, CommonData.personData);
            CommonData.oleConn.Close();
            */
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Microsoft.Office.Interop.Excel.Application oApp;
            //Excel.Worksheet oSheet;
            //Excel.Workbook oBook;
            //oApp = new Excel.Application();
            ///oBook = oApp.Workbooks.Add();
            ///
            /*
            string outPutFilePath = @".\data.xls";
            if (File.Exists(outPutFilePath))
            {
                File.Delete(outPutFilePath);
            }
            OleDbConnection oleConn = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source =" + outPutFilePath + ";Extended Properties=Excel 8.0");
            oleConn.Open();
            OleDbCommand cmd = oleConn.CreateCommand();
            //string sheetName = "案件信息";
            string createTableString = "CREATE TABLE 案件信息 (";
            for(int i=0;i< CommonData.personInfo.GetInfoNum(); i++)
            {
                createTableString += "[" + CommonData.personInfo.GetItem(i) + "] text";
                if (i != CommonData.personInfo.GetInfoNum() - 1)
                    createTableString += ",";
            }
            createTableString += ")";
            //cmd.CommandText = "CREATE TABLE 案件信息 ([序号] INTEGER, [刀号] VarChar,[刀具规格] VarChar,[1月] VarChar,[2月] VarChar,[3月] VarChar,[4月] VarChar,[5月] VarChar,[6月] VarChar,[7月] VarChar,[8月] VarChar,[9月] VarChar,[10月] VarChar,[11月] VarChar,[12月] VarChar)";
            Console.WriteLine(createTableString);
            //Console.WriteLine(cmd.CommandText);
            cmd.CommandText = createTableString;
            cmd.ExecuteNonQuery();
             */
            string s = CommonFileDialog.ChoseDir();
            OleDbConnection o = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source =" + s+ @"\data.xls" + ";Extended Properties=Excel 8.0");
            o.Open();
            //Console.WriteLine(s);
            OleDbCommand cmd = o.CreateCommand();
            string createTableString = "CREATE TABLE 案件信息 (";
            for (int i = 0; i < CommonData.personInfo.GetInfoNum(); i++)
            {
                createTableString += "[" + CommonData.personInfo.GetItem(i) + "] text";
                if (i != CommonData.personInfo.GetInfoNum() - 1)
                    createTableString += ",";
            }
            createTableString += ")";
            cmd.CommandText = createTableString;
            //Console.WriteLine(createTableString);
            cmd.ExecuteNonQuery();
            FillFile(CommonData.personData, CommonData.oleConn.CreateCommand());
            //CommonFileDialog.GetFilePath("excel文件|*.xls");
            o.Close();
        }

        private void ClraerData()
        {
            this.listBox1.Items.Clear();
            //listBox1.Show();
            this.listBox2.Items.Clear();
            //listBox2.Show();
        }

        //搜索
        private void button1_Click(object sender, EventArgs e)
        {
            string target = comboBox2.Text;
            string searchString = this.textBox1.Text;
            target = target.Replace("按照", "");
            target = target.Replace("搜索", "");
            if(searchString == "")
            {
                ClraerData();
                FillForm2(CommonData.personInfo, CommonData.personData);
                CommonData.ifSearch = false;
                return;
            }
            int targetPos = -1;
            for(int i = 0; i < CommonData.personInfo.GetInfoNum(); i++)
            {
                //Console.WriteLine(CommonData.personInfo.GetItem(i));
                if (CommonData.personInfo.GetItem(i) == target)
                {
                    targetPos = i;
                    break;
                }
            }
            if(targetPos != -1)
            {
                for(int i = 0; i < CommonData.personData.Count; i++)
                {
                    if (((Person)CommonData.personData[i]).getPersonInfo(targetPos).Contains(searchString))
                    {
                        Person p = (Person)CommonData.personData[i];
                        p.searchPos = i;
                        CommonData.searchPersonData.Add(p);
                    }
                }
            }
            else
            {
                for (int i = 0; i < CommonData.personData.Count; i++)
                {
                    for(int j=0;j< CommonData.personInfo.GetInfoNum(); j++)
                    {
                        if (((Person)CommonData.personData[i]).getPersonInfo(j).Contains(searchString))
                        {
                            Person p = (Person)CommonData.personData[i];
                            p.searchPos = i;
                            CommonData.searchPersonData.Add(p);
                            break;
                        }
                    }
                }
            }
            ClraerData();
            //Console.WriteLine(CommonData.searchPersonData.Count);
            CommonData.ifSearch = true;
            FillForm2(CommonData.personInfo, CommonData.searchPersonData);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ClraerData();
            CommonData.searchPersonData.Clear();
            FillForm2(CommonData.personInfo, CommonData.personData);
            CommonData.ifSearch = false;
            //Console.WriteLine(this.listBox1.Items.Count);
        }

        public void DelateFileItem(string id)
        {
            //OleDb不支持删除一行
            OleDbCommand cmd = CommonData.oleConn.CreateCommand();
            cmd.CommandText = @"drop table [案件信息]";
            cmd.ExecuteNonQuery();

            //OleDbCommand cmd = CommonData.oleConn.CreateCommand();
            string createTableString = "CREATE TABLE 案件信息 (";
            for (int i = 0; i < CommonData.personInfo.GetInfoNum(); i++)
            {
                createTableString += "[" + CommonData.personInfo.GetItem(i) + "] text";
                if (i != CommonData.personInfo.GetInfoNum() - 1)
                    createTableString += ",";
            }
            createTableString += ")";
            cmd.CommandText = createTableString;
            //Console.WriteLine(createTableString);
            cmd.ExecuteNonQuery();
            FillFile(CommonData.personData, CommonData.oleConn.CreateCommand());

            //cmd.Dispose();
            //Microsoft.Office.Interop.Excel.Application.Workbooks.Open(".\data.xls").ActiveSheet.Rows(1).Delete True;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (CommonData.ifSearch == true)
            {
                MessageBox.Show("先清除搜索结果");
                return;
            }
            int pos = this.listBox1.SelectedIndex;
            if(CommonData.ifSearch == true)
            {
                listBox1.Items.RemoveAt(pos);
                CommonData.searchPersonData.RemoveAt(pos);
                pos = ((Person)CommonData.searchPersonData[pos]).searchPos;
                CommonData.personData.RemoveAt(((Person)CommonData.searchPersonData[pos]).searchPos);
            }
            else
            {
                if (pos == -1)
                    return;
                listBox1.Items.RemoveAt(pos);
                CommonData.personData.RemoveAt(pos);
            }
            //Console.WriteLine(CommonData.personData.Count);
            DelateFileItem(((Person)CommonData.personData[pos]).getPersonInfo(0));
            
            //DelateFileItem("编号3");
        }

        public void SimpleUpdateFile()
        {
            File.Delete(@".\data.xls");
            CreateTableFillFile();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            string picturePath = CommonFileIfPicture.GetPicturePath();
            if (picturePath == "") return;
            if (this.listBox1.SelectedIndex >= 0 && this.listBox1.SelectedIndex < CommonData.personData.Count)
            {
                string id = ((Person)CommonData.personData[listBox1.SelectedIndex]).getPersonInfo(0);
                File.Copy(picturePath, @".\picture\" + id + "1" + Path.GetFileName(picturePath), true);
                ((Person)CommonData.personData[listBox1.SelectedIndex]).setPersonPictureInfo(0, picturePath);
                //删除原文件？
                /*
                if (File.Exists(((Person)CommonData.personData[listBox1.SelectedIndex]).getPersonPicture(0)))
                {
                    File.Delete(((Person)CommonData.personData[listBox1.SelectedIndex]).getPersonPicture(0));
                }
                */
                //WriteToFile();
                SimpleUpdateFile();
            }
            try
            {
                this.pictureBox1.Image = Image.FromFile(picturePath);
                this.pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            }
            catch
            {
                MessageBox.Show("找不到图片" + picturePath);
            }
        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            //if(CommonData.ifFirstLogin == true)
            {
                e.Graphics.DrawString("点击选择", this.Font, Brushes.Black, 15, 15);
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            string picturePath = CommonFileIfPicture.GetPicturePath();
            if (picturePath == "") return;
            if (this.listBox1.SelectedIndex >= 0 && this.listBox1.SelectedIndex < CommonData.personData.Count)
            {
                string id = ((Person)CommonData.personData[listBox1.SelectedIndex]).getPersonInfo(0);
                File.Copy(picturePath, @".\picture\" + id + "1" + Path.GetFileName(picturePath), true);
                ((Person)CommonData.personData[listBox1.SelectedIndex]).setPersonPictureInfo(1, picturePath);
                //删除原文件？
                /*
                if (File.Exists(((Person)CommonData.personData[listBox1.SelectedIndex]).getPersonPicture(0)))
                {
                    File.Delete(((Person)CommonData.personData[listBox1.SelectedIndex]).getPersonPicture(0));
                }
                */
                //WriteToFile();
                SimpleUpdateFile();
            }
            try
            {
                this.pictureBox2.Image = Image.FromFile(picturePath);
                this.pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
            }
            catch
            {
                MessageBox.Show("找不到图片" + picturePath);
            }
        }

        private void pictureBox2_Paint(object sender, PaintEventArgs e)
        {
            //if (CommonData.ifFirstLogin == true)
            {
                e.Graphics.DrawString("点击选择", this.Font, Brushes.Black, 15, 15);
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string picturePath = CommonFileIfPicture.GetPicturePath();
            if (picturePath == "") return;
            if (this.listBox1.SelectedIndex >= 0 && this.listBox1.SelectedIndex < CommonData.personData.Count)
            {
                string id = ((Person)CommonData.personData[listBox1.SelectedIndex]).getPersonInfo(0);
                File.Copy(picturePath, @".\picture\" + id + "1" + Path.GetFileName(picturePath), true);
                ((Person)CommonData.personData[listBox1.SelectedIndex]).setPersonPictureInfo(2, picturePath);
                //删除原文件？
                /*
                if (File.Exists(((Person)CommonData.personData[listBox1.SelectedIndex]).getPersonPicture(0)))
                {
                    File.Delete(((Person)CommonData.personData[listBox1.SelectedIndex]).getPersonPicture(0));
                }
                */
                //WriteToFile();
                SimpleUpdateFile();
            }
            try
            {
                this.pictureBox3.Image = Image.FromFile(picturePath);
                this.pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            }
            catch
            {
                MessageBox.Show("找不到图片" + picturePath);
            }
        }

        private void pictureBox3_Paint(object sender, PaintEventArgs e)
        {
            //if (CommonData.ifFirstLogin == true)
            {
                e.Graphics.DrawString("点击选择", this.Font, Brushes.Black, 15, 15);
            }
        }

        public static ArrayList newPerson = new ArrayList();
        public static bool ifCreate = false;
        public static bool ifChange = false;

        private void button5_Click(object sender, EventArgs e)
        {
            if(CommonData.ifSearch == true)
            {
                MessageBox.Show("先清除搜索结果");
                return;
            }
            Person p = new Person();
            p.Add(CommonData.getNewID());
            for(int i = 1; i < CommonData.personInfo.GetInfoNum(); i++)
            {
                p.Add("");
            }
            Form3 f3 = new Form3(CommonData.personInfo, ref p, 0);
            f3.ShowDialog();
            if (ifCreate)
            {
                CommonData.personData.Add(p);
                //CommonData.personData.Sort();
                SimpleUpdateFile();
                ClraerData();
                FillForm2(CommonData.personInfo, CommonData.personData);
                ifCreate = false;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (CommonData.ifSearch == true)
            {
                MessageBox.Show("先清除搜索结果");
                return;
            }
            Person pNew = new Person(1);
            Person.PersonCopy((Person)CommonData.personData[listBox1.SelectedIndex], pNew);
            //Person p = (Person)CommonData.personData[listBox1.SelectedIndex];
            Form3 f3 = new Form3(CommonData.personInfo, ref pNew, 1);
            f3.ShowDialog();
            if(ifChange == true)
            {
                Person.PersonCopy(pNew, (Person)CommonData.personData[listBox1.SelectedIndex]);
                SimpleUpdateFile();
                ClraerData();
                FillForm2(CommonData.personInfo, CommonData.personData);
                ifChange = false;
            }
        }

        public static bool ifDelateItem = false;
        public static string delateName;

        //删除字段
        private void button4_Click(object sender, EventArgs e)
        {
            if (CommonData.ifSearch == true)
            {
                MessageBox.Show("先清除搜索结果");
                return;
            }
            Form5 f5 = new Form5();
            f5.ShowDialog();
            if (ifDelateItem)
            {
                ifDelateItem = false;
                int found = -1;
                for (int i = 1; i < CommonData.personInfo.GetInfoNum(); i++)
                {
                    if(CommonData.personInfo.GetItem(i) == delateName)
                    {
                        found = i;
                        break;
                    }
                }
                if(found == -1)
                {
                    MessageBox.Show("没有该字段");
                    return;
                }
                CommonData.personInfo.DeleteItem(found);
                for(int i = 0; i < CommonData.personData.Count; i++)
                {
                    ((Person)(CommonData.personData[i])).DeleteItem(found);
                }
                ClraerData();
                FillForm2(CommonData.personInfo, CommonData.personData);
                SimpleUpdateFile();
            }
        }

        public static bool ifNewItem = false;
        public static string sName;
        public static string sValue;
        // 增加字段
        private void button6_Click(object sender, EventArgs e)
        {
            if (CommonData.ifSearch == true)
            {
                MessageBox.Show("先清除搜索结果");
                return;
            }
            //Dictionary<string, string> myDictionary = new Dictionary<string, string>();
            Form4 fNew = new Form4();
            fNew.ShowDialog();
            if (ifNewItem)
            {
                ifNewItem = false;
                CommonData.personInfo.InsertItem(sName);
                for(int i = 0; i < CommonData.personData.Count; i++)
                {
                    ((Person)(CommonData.personData[i])).InsertItem(sValue);
                }
                ClraerData();
                FillForm2(CommonData.personInfo, CommonData.personData);
                SimpleUpdateFile();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string target = comboBox2.Text;
            string searchString = this.textBox1.Text;
            target = target.Replace("按照", "");
            target = target.Replace("搜索", "");
            if (searchString == "")
            {
                ClraerData();
                FillForm2(CommonData.personInfo, CommonData.personData);
                CommonData.ifSearch = false;
                return;
            }
            int targetPos = -1;
            for (int i = 0; i < CommonData.personInfo.GetInfoNum(); i++)
            {
                //Console.WriteLine(CommonData.personInfo.GetItem(i));
                if (CommonData.personInfo.GetItem(i) == target)
                {
                    targetPos = i;
                    break;
                }
            }
            try { 
            if (targetPos != -1)
            {
                //按字段搜索
                for (int i = 0; i < CommonData.personData.Count; i++)
                {
                    MatchCollection mc = Regex.Matches(((Person)CommonData.personData[i]).getPersonInfo(targetPos), searchString);
                    if (mc.Count != 0)
                    {
                        Person p = (Person)CommonData.personData[i];
                        p.searchPos = i;
                        CommonData.searchPersonData.Add(p);
                    }
                }
            }
            else
            {
                //全文搜索
                for (int i = 0; i < CommonData.personData.Count; i++)
                {
                    for (int j = 0; j < CommonData.personInfo.GetInfoNum(); j++)
                    {
                        MatchCollection mc = Regex.Matches(((Person)CommonData.personData[i]).getPersonInfo(j), searchString);
                        if (mc.Count != 0)
                        {
                            Person p = (Person)CommonData.personData[i];
                            p.searchPos = i;
                            CommonData.searchPersonData.Add(p);
                            break;
                        }
                    }
                }
            }
            }
            catch
            {
                MessageBox.Show("正则表达式错误");
            }
            ClraerData();
            //Console.WriteLine(CommonData.searchPersonData.Count);
            CommonData.ifSearch = true;
            FillForm2(CommonData.personInfo, CommonData.searchPersonData);
        }
    }
}
