using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.IO;
using WinForm = System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WpfApp4
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        Controller controller;
        public MainWindow()
        {
            InitializeComponent();

            if (Check())
            {
                controller = new Controller(Path());
                Console.WriteLine("1");
                controller.Load();
                List.Items.Clear();
                foreach (string n in controller.nameList)
                {
                    List.Items.Add(n);
                }
            }            
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            if (Id.Text.Length == 6 && !Id.Text.Equals("000000"))
            {
                MainData mainData = new MainData(Id.Text);
                string[] d = mainData.GetData();
                foreach (string s in d)
                {
                    Console.WriteLine(s);
                }
                int index = controller.GetIndex();
                Console.WriteLine(index);
                controller.Insert(mainData, index + 1);
                controller.Load();
                List.Items.Clear();
                foreach (string n in controller.nameList)
                {
                    List.Items.Add(n);
                }
            }
            else
            {
                MessageBox.Show("代码错误", "提示");
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            int count = List.SelectedItems.Count;
            List<string> itemValues = new List<string>();
            if (count != 0)
            {
                for (int i = 0; i < count; i++)
                {
                    controller.Delete(List.SelectedItems[i].ToString());
                }
                controller.Load();
                List.Items.Clear();
                foreach (string n in controller.nameList)
                {
                    List.Items.Add(n);
                }
            }
            else
            {
                MessageBox.Show("请选择！", "提示");
            }
        }

        private void Search_Click(object sender, RoutedEventArgs e)
        {
            int c = List.Items.Count;
            for (int i = 0; i < c; i++)
            {
                MainData mainData = new MainData(controller.idList[i]);
                controller.Insert(mainData, i + 2);
            }
        }

        private void SP_Click(object sender, RoutedEventArgs e)
        {
            WinForm.FolderBrowserDialog dialog = new WinForm.FolderBrowserDialog();
            dialog.ShowDialog();
            FileInfo f = new FileInfo(@"gshxsj/path.txt");
            StreamWriter w = f.CreateText();
            w.Write(dialog.SelectedPath);
            w.Close();
            controller = new Controller(dialog.SelectedPath);
            controller.Load();
            List.Items.Clear();
            foreach (string n in controller.nameList)
            {
                List.Items.Add(n);
            }
        }

        string Path()
        {
            StreamReader p = new StreamReader(@"Release/path.txt");
            string m_Data = p.ReadToEnd();
            p.Close();
            Console.WriteLine(m_Data);
            return m_Data;
        }

        Boolean Check()
        {
            string path = Path();
            if (!File.Exists(path + "\\公司核心数据.xls")) 
            {
                CK.Visibility = Visibility.Visible;
                return false;
            }
            return true;
        }

        private void Select_Click(object sender, RoutedEventArgs e)
        {
            WinForm.FolderBrowserDialog dialog = new WinForm.FolderBrowserDialog();
            dialog.ShowDialog();
            Path_Label.Content = dialog.SelectedPath;
        }

        private void Yes_Click(object sender, RoutedEventArgs e)
        {
            FileInfo f = new FileInfo(@"Release/path.txt");
            StreamWriter w = f.CreateText();
            w.Write(Path_Label.Content.ToString());
            w.Close();
            controller = new Controller(Path_Label.Content.ToString());
            controller.Load();
            List.Items.Clear();
            foreach (string n in controller.nameList)
            {
                List.Items.Add(n);
            }
            CK.Visibility = Visibility.Collapsed;
        }
    }
}
