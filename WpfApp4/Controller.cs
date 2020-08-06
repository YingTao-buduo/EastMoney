using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace WpfApp4
{
    class Controller
    {
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);


        string fileName = "";

        public List<string> idList = new List<string>();
        public List<string> nameList = new List<string>();

        public Controller(string fn)
        {
            fileName = fn + "\\公司核心数据.xls";
            try
            {
                Excel.Application EXC = new Excel.Application();
                EXC.Visible = false;
                Excel.Workbook workbook;
                Console.WriteLine(fileName);
                if (File.Exists(fileName))
                {
                    workbook = EXC.Workbooks.Open(fileName);
                    
                }
                else
                {
                    workbook = EXC.Workbooks.Add(true);
                    Excel._Worksheet sheet = EXC.Worksheets.Add();
                    sheet.Activate();
                    string[] head = { "公司名称", "股票代码", "收益", "PE（动）", "每股净资产", "市净率", "总收益", "总收益—同比", "净利润", "净利润-同比", "毛利率", "净利率", "ROE", "负债率", "总股本", "总值", "流通股", "流值", "每股未分配利润" };
                    for (int i = 0; i < head.Length; i++)
                    {
                        Excel.Range r = sheet.Cells[1, i + 1];
                        r.Columns.AutoFit();
                        r.Value2 = head[i];
                    }
                    sheet.SaveAs(fileName);
                }
                workbook.Close(true, Missing.Value, Missing.Value);
                EXC.Quit();
                System.GC.GetGeneration(EXC);
                IntPtr t = new IntPtr(EXC.Hwnd);
                int k = 0;
                GetWindowThreadProcessId(t, out k);
                Process p = Process.GetProcessById(k);
                p.Kill();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(EXC);

            }
            catch (Exception e)
            {
                Console.WriteLine("controller.Controller");
            }

        }

        public void Insert(MainData mainData, int index)
        {
            try
            {
                Excel.Application EXC = new Excel.Application();
                EXC.Visible = false;
                Excel.Workbook workbook = EXC.Workbooks.Open(fileName);
                Excel._Worksheet sheet = workbook.Sheets[1];
                sheet.Activate();
                string[] data = mainData.GetData();
                for (int i = 0; i < data.Length; i++)
                {
                    Excel.Range r = sheet.Cells[index, i + 1];
                    r.Value2 = data[i];
                }

                workbook.Close(true, Missing.Value, Missing.Value);
                EXC.Quit();
                System.GC.GetGeneration(EXC);
                IntPtr t = new IntPtr(EXC.Hwnd);
                int k = 0;
                GetWindowThreadProcessId(t, out k);
                Process p = Process.GetProcessById(k);
                p.Kill();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(EXC);

            }
            catch (Exception e)
            {
                Console.WriteLine("controller.Insert");
            }
        }

        public void Load()
        {
            try
            {
                idList.Clear();
                nameList.Clear();
                Excel.Application EXC = new Excel.Application();
                EXC.Visible = false;
                Excel.Workbook workbook = EXC.Workbooks.Open(fileName);
                Excel._Worksheet sheet = workbook.Sheets[1];
                sheet.Activate();
                int c = sheet.UsedRange.Rows.Count;
                for (int i = 0; i < c - 1; i++)
                {
                    Excel.Range r_n = sheet.Cells[i + 2, 1];
                    Excel.Range r_i = sheet.Cells[i + 2, 2];
                    nameList.Add(r_n.Text);
                    idList.Add(r_i.Text);

                }
                //sheet.Rows[2, Missing.Value].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                workbook.Close(true, Missing.Value, Missing.Value);
                EXC.Quit();
                System.GC.GetGeneration(EXC);
                IntPtr t = new IntPtr(EXC.Hwnd);
                int k = 0;
                GetWindowThreadProcessId(t, out k);
                Process p = Process.GetProcessById(k);
                p.Kill();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(EXC);
            }
            catch (Exception e)
            {
                Console.WriteLine("controller.Load");
            }
        }

        public void Delete(string name)
        {
            Excel.Application EXC = new Excel.Application();
            EXC.Visible = false;
            Excel.Workbook workbook = EXC.Workbooks.Open(fileName);
            Excel._Worksheet sheet = workbook.Sheets[1];
            sheet.Activate();
            int c = sheet.UsedRange.Rows.Count;
            for (int i = 0; i < c - 1; i++)
            {
                Excel.Range r_n = sheet.Cells[i + 2, 1];
                if (name.Equals(r_n.Text))
                {
                    sheet.Rows[i + 2, Missing.Value].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                }

            }
            //sheet.Rows[2, Missing.Value].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

            workbook.Close(true, Missing.Value, Missing.Value);
            EXC.Quit();
            System.GC.GetGeneration(EXC);
            IntPtr t = new IntPtr(EXC.Hwnd);
            int k = 0;
            GetWindowThreadProcessId(t, out k);
            Process p = Process.GetProcessById(k);
            p.Kill();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(EXC);
        }

        public int GetIndex()
        {
            int c = 0;
            try
            {
                Excel.Application EXC = new Excel.Application();
                EXC.Visible = false;
                Excel.Workbook workbook = EXC.Workbooks.Open(fileName);
                Excel._Worksheet sheet = workbook.Sheets[1];
                sheet.Activate();
                c = sheet.UsedRange.Rows.Count;
                workbook.Close(true, Missing.Value, Missing.Value);
                EXC.Quit();
                System.GC.GetGeneration(EXC);
                IntPtr t = new IntPtr(EXC.Hwnd);
                int k = 0;
                GetWindowThreadProcessId(t, out k);
                Process p = Process.GetProcessById(k);
                p.Kill();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(EXC);
            }
            catch (Exception e)
            {
                Console.WriteLine("controller.GetIndex");
            }
            

            return c;
        }

        public void SetFileName(string p)
        {
            fileName = p + "\\公司核心数据.xls"; 
        }
    }
}
