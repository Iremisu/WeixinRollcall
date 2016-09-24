using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using weixinRollCall.DAL.Model;
using weixinRollCall.DAL.DAO;

namespace weixinRollCall.Common
{
    public class Excel
    {
        private string path;
        private List<ClassStudent> cls;
        private HSSFWorkbook wb;
        private FileStream file;
        private HSSFSheet sheet1;

        public Excel(List<ClassStudent> CS,string path)//打开
        {
            cls = CS;
            this.path = path;
        }
        public Excel(List<ClassStudent> CS,string path,string classname)//创建模板
        {
            int l = 3;
            this.path = path;
            cls = CS;
            wb = new HSSFWorkbook();
            HSSFSheet sheet1 = (HSSFSheet)wb.CreateSheet(classname+"点名情况汇总");
            sheet1.CreateRow(0).CreateCell(0).SetCellValue(classname + "点名情况汇总");
            sheet1.AddMergedRegion(new NPOI.SS.Util.Region(0, 0, 0, 2));
            sheet1.CreateRow(3).CreateCell(0).SetCellValue("姓名");
            sheet1.GetRow(3).CreateCell(1).SetCellValue("学号");
            sheet1.GetRow(3).CreateCell(2).SetCellValue("班级");
            sheet1.SetColumnWidth(1, 15 * 256);
            sheet1.SetColumnWidth(2, 15 * 256);
            foreach (ClassStudent cs in CS)
            {
                l++;
                sheet1.CreateRow(l).CreateCell(0).SetCellValue(cs.StudentName);
                sheet1.GetRow(l).CreateCell(1).SetCellValue(cs.StudentID);
                sheet1.GetRow(l).CreateCell(2).SetCellValue(cs.StudentClass);
            }
            FileStream file1 = new FileStream(path, FileMode.Create);
            wb.Write(file1);
            file1.Close();
        }
        public void editexcel(string cid)
        {
            file = new FileStream(path, FileMode.Open);
            wb = new HSSFWorkbook(file);
            sheet1 = (HSSFSheet)wb.GetSheetAt(0);
            List<string> Date = new RollCallDAO().GetDate(cid);
            foreach(string date in Date)
            {
                List<string> s = new RollCallDAO().GetStatus(cls, date,cid);
                editcol(s, date);                
            }
            FileStream file1 = new FileStream(path, FileMode.Create);
            wb.Write(file1);
            file1.Close();
            file.Close();
        }
        public void editcol(List<string> S,string date)
        {
            int l = 0;
            int c = 3;
            l = sheet1.GetRow(c).LastCellNum;
            sheet1.SetColumnWidth(l, 17 * 256);
            sheet1.GetRow(c).CreateCell(l).SetCellValue(date);
            foreach (string s in S)
            {
                c++;
                sheet1.GetRow(c).CreateCell(l).SetCellValue(s);
            }
        }
    }
}