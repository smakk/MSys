using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.Data.OleDb;


namespace MSys
{
    public class PersonInfo
    {
        //总共有多少个字段
        private ArrayList personInfoList = new ArrayList();

        public void Add(string info)
        {
            personInfoList.Add(info);
        }

        public int GetInfoNum()
        {
            return personInfoList.Count;
        }

        public int GetInfoNumNoPicture()
        {
            return GetInfoNum() - 3;
        }

        public int GetPictureNum()
        {
            return 3;
        }

        public string GetItem(int num)
        {
            return (string)personInfoList[num];
        }

        public void InsertItem(string s)
        {
            this.personInfoList.Insert(personInfoList.Count - 3, s);
        }

        public void DeleteItem(int index)
        {
            this.personInfoList.RemoveAt(index);
        }
    }

    public class Person : IComparable
    {
        //如果是搜索结果，该字段代表搜索结果在personData中的位置
        public int searchPos = -1;

        public ArrayList person = new ArrayList();
        //private ArrayList personPicture = new ArrayList();
        public Person(string s0, string s1, string s2, string s3, string s4, string s5, string s6, 
            string s7, string s8, string s9, string s10, string s11, string s12, string s13, string s14, string s15, 
            string s16, string s17, string s18, string s19, string s20, string s21, string s22, string s23, string s24, string s25)//, string s3,
            //string s1, string s2, string s3, string s1, string s2, string s3, string s1, string s2, string s3, )
        {
            person.Add(s0);
            person.Add("中文名"+s1);
            person.Add("英文名"+s2);
            person.Add("其他名字"+s3);
            person.Add("性别"+s4);
            person.Add("年龄"+s5);
            person.Add("出生日期"+s6);
            person.Add("身高"+s7);
            person.Add("血型"+s8);
            person.Add("民族"+s9);
            person.Add("宗教"+s10);
            person.Add("国籍"+s11);
            person.Add("婚姻"+s12);
            person.Add("党派"+s13);
            person.Add("证件类型"+s14);
            person.Add("证件号"+s15);
            person.Add("证件地址"+s16);
            person.Add("产权地址"+s17);
            person.Add("登记地址"+s18);
            person.Add("寄递地址"+s19);
            person.Add("租房地址"+s20);
            person.Add("工作单位"+s21);
            person.Add("工作地址"+s22);
            person.Add("图片1_" + s23);
            person.Add("图片2_" + s24);
            person.Add("图片3_" + s25);
        }

        public Person()
        {

        }

        public Person(int flag)
        {
            for(int i = 0; i < CommonData.personInfo.GetInfoNum(); i++)
            {
                this.person.Add("");
            }
        }

        public static void PersonCopy(Person p1, Person p2)
        {
            for(int i = 0; i < p1.person.Count; i++)
            {
                p2.person[i] = p1.person[i];
            }
        }
        /*
        public void PirtureAdd(string pictureAddr)
        {
            personPicture.Add(pictureAddr);
        }
        */
        public int CompareTo(object obj)
        {
            Person p = (Person)obj;

            return string.Compare(this.getPersonInfo(CommonData.keyInfo) , p.getPersonInfo(CommonData.keyInfo), true);
        }

        public string getPersonInfo(int num)
        {
            return person[num].ToString();
        }

        
        public string getPersonPicture(int num)
        {
            return this.person[person.Count - 3 + num].ToString();
        }
        

        public void Add(string s)
        {
            person.Add(s);
        }
        public void setPersonInfo(int num, string s)
        {
            person[num] = s;
        }

        public void setPersonPictureInfo(int num, string s)
        {
            person[person.Count - 3 + num] = s;
        }

        public void InsertItem(string s)
        {
            this.person.Insert(person.Count - 3, s);
        }

        public void DeleteItem(int index)
        {
            this.person.RemoveAt(index);
        }
    }

    public static class CommonFileDialog
    { 
        public static string GetFilePath(string filter)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = Application.StartupPath;
            ofd.Title = "请选择要打开的文件";
            ofd.Multiselect = true;
            ofd.Filter = filter;
            ofd.FilterIndex = 2;
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string filePath = ofd.FileName;
                string fileName = ofd.SafeFileName;
                Console.WriteLine(filePath + fileName);
                return filePath;
            }
            return "";
        }

        public static string ChoseDir()
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            return path.SelectedPath;
        }
    }

    public static class CommonFileIfPicture
    {
        public static bool FileIsPicture(string path)
        {
            return path.EndsWith("bmp") || path.EndsWith("jpg") || path.EndsWith("png") || path.EndsWith("jpeg");
        }

        public static string GetPicturePath()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = Application.StartupPath;
            ofd.Title = "请选择要打开的文件";
            ofd.Multiselect = true;
            //ofd.Filter = "bmp文件 | *.bmp | jpg文件 | *.jpg | png文件 | *.png | jpg文件 | *.jpg";
            ofd.Filter = @"图片文件(*.png;*.jpg;*.bmp;*.jpeg)|*.png;*.jpg;*.bmp;*.jpeg";
            ofd.FilterIndex = 2;
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string filePath = ofd.FileName;
                string fileName = ofd.SafeFileName;
                Console.WriteLine(filePath + fileName);
                return filePath;
            }
            return "";
        }
    }

    public static class CommonData
    {
        //Form2依照哪一个字段显示
        public static int keyInfo = 0;
        //是不是第一次登录，以目录是否存在data文件为标准
        public static bool ifFirstLogin = true;
        //是否是管理员登录
        public static bool rootLogin;
        //一个人有哪些信息
        public static PersonInfo personInfo = new PersonInfo();
        //所有人的信息，一个Person类型的数组
        public static ArrayList personData = new ArrayList();
        //搜索结果
        public static ArrayList searchPersonData = new ArrayList();
        //excel的连接
        public static OleDbConnection oleConn = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source =" + @".\data.xls" + ";Extended Properties=Excel 8.0");
        //是否显示的是搜索结果
        public static bool ifSearch = false;
        public static void InitStaticData()
        {
            CommonData.rootLogin = false;
        }

        public static string getNewID()
        {
            int res = -1;
            for(int i = 0; i < personData.Count; i++)
            {
                int temp = int.Parse(((Person)personData[i]).getPersonInfo(0));
                if (temp > res)
                    res = temp;
            }
            return (res+1).ToString();
        }

        //以第几个字段对内容进行排序
        public static void sortPersonData(int num)
        {
            keyInfo = num;
            personData.Sort();
        }

        //从数据库中填充一些默认数据
        public static int FillData()
        {
            //确定字段
            personInfo.Add("编号");
            personInfo.Add("中文名");
            personInfo.Add("英文名");
            personInfo.Add("其他名字");
            personInfo.Add("性别");
            personInfo.Add("年龄");
            personInfo.Add("出生日期");
            personInfo.Add("身高");
            personInfo.Add("血型");
            personInfo.Add("民族");
            personInfo.Add("宗教");
            personInfo.Add("国籍");
            personInfo.Add("婚姻");
            personInfo.Add("党派");
            personInfo.Add("证件类型");
            personInfo.Add("证件号");
            personInfo.Add("证件地址");
            personInfo.Add("产权地址");
            personInfo.Add("登记地址");
            personInfo.Add("寄递地址");
            personInfo.Add("租房地址");
            personInfo.Add("工作单位");
            personInfo.Add("工作地址");
            personInfo.Add("图片一");
            personInfo.Add("图片二");
            personInfo.Add("图片三");

            //填充数据
            for (int i = 0; i < 10; i++)
            {
                personData.Add(new Person(i+"", i+"",i + "", i + "", i + "", i + "", i + "", i + "", i + "", i + "", i + "",
                    i + "", i + "", i + "", i + "", i + "", i + "", i + "", i + "", i + "", i + "", i + "", i + "", i + "", i + "", i + ""));
            }
            
            return 0;
        }
    }

    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());

            CommonData.InitStaticData();
        }
    }
}
