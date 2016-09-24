using System;
using System.Text;
using System.Data;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Configuration;
using System.Web;
//using COI.TLM.BLL;
//using COI.TLM.Model;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using weixinRollCall.DAL.Model;
public class ExcelHelper
{
    private Excel.Application excelApp;
    private Excel.Workbooks excelWorkBooks;
    private Excel.Workbook excelWorkBook;
    private Excel.Worksheet excelWorkSheet;
    private object oMissing = Missing.Value;
    private string templetFile = null;
    private string outputFile = null;
    private object missing = Missing.Value;
    private DateTime beforeTime=DateTime.Now;			//Excel启动之前时间
    private DateTime afterTime=DateTime.Now;	//Excel启动之后时间

    public Excel.Application ExcelApp
    {
        get
        {
            if (excelApp == null)
            {
                excelApp = new Excel.Application();
            }
            return excelApp;
        }
        set { excelApp = value; }
    }

    public Excel.Workbooks ExcelWorkBooks
    {
        get { return excelWorkBooks; }
        set { excelWorkBooks = value; }
    }
    

    public Excel.Workbook ExcelWorkBook
    {
        get { return excelWorkBook; }
        set { excelWorkBook = value; }
    }
    

    public Excel.Worksheet ExcelWorkSheet
    {
        get { return excelWorkSheet; }
        set { excelWorkSheet = value; }
    }


    /// <summary>
    /// 创建一个excel对象
    /// </summary>
    public ExcelHelper()
    {

        excelWorkBooks = ExcelApp.Workbooks;
        excelWorkBook = excelWorkBooks.Add(true);
        excelWorkSheet = excelWorkBook.Worksheets[1] as Excel.Worksheet;
    }
    /// <summary>
    /// 打开工作薄
    /// </summary>
    /// <param name="fileName">文件名</param>
    public ExcelHelper(string fileName)
    {
        excelWorkBooks = ExcelApp.Workbooks;
        excelWorkBook = excelWorkBooks.Add(fileName);
        excelWorkSheet = excelWorkBook.Worksheets[1] as Excel.Worksheet;
    }
    /// <summary>
    /// 构造函数，将一个已有Excel工作簿作为模板，并指定输出路径
    /// </summary>
    /// <param name="templetFilePath">Excel模板文件路径</param>
    /// <param name="outputFilePath">输出Excel文件路径</param>
    public ExcelHelper(string templetFilePath, string outputFilePath)
    {
        if (templetFilePath == null)
            throw new Exception("Excel模板文件路径不能为空！");

        if (outputFilePath == null)
            throw new Exception("输出Excel文件路径不能为空！");

        if (!File.Exists(templetFilePath))
            throw new Exception("指定路径的Excel模板文件不存在！");

        this.templetFile = templetFilePath;
        this.outputFile = outputFilePath;

        //创建一个Application对象并使其可见
        beforeTime = DateTime.Now;
        excelApp = new Excel.Application();
        excelApp.DisplayAlerts = false;
        excelApp.AlertBeforeOverwriting = false;//new Excel.Application();//
        //excelApp.Visible = true;
        afterTime = DateTime.Now;

        //打开模板文件，得到WorkBook对象
        excelWorkBook = excelApp.Workbooks.Open(templetFile, missing, missing, missing, missing, missing,
            missing, missing, missing, missing, missing, missing, missing, missing, missing);

        //得到WorkSheet对象
        excelWorkSheet = (Excel.Worksheet)excelWorkBook.Sheets.get_Item(1);

    }
    public int getcol()
    {
        for (int i = 1; ; i++)
        {
            if (string.IsNullOrEmpty(GetCell(4, i).Text))
            {
                return i;
            }
        }
    }
    public void RCDataToExcel(List<string> s)
    {
        int c = getcol();
        int l = 4;
        SetCellValue(l, c, DateTime.Now.ToString("yyyy/MM/dd"));
        foreach (string i in s)
        {
            l++;
            SetCellValue(l, c, i);
        }
    }
    public void newDataToExcel(List<ClassStudent> CS,string classname)
    {
        SetCellValue(1, 1, classname + "点名情况汇总");
        int l = 4;
        SetCellValue(l, 1, "学号");
        SetCellValue(l, 2, "姓名");
        SetCellValue(l, 3, "班级");
        foreach (ClassStudent cs in CS)
        {
            l++;
            SetCellValue(l, 1, cs.StudentID);
            SetCellValue(l, 2, cs.StudentName);
            SetCellValue(l, 3, cs.StudentClass);
        }
    }
    /* public void DataToExcel(List<淘宝抓取工具.BaseData.ReceiveData> activitys)
     {

         if (activitys != null && activitys.Count > 0)
         {
             for (int i = 0; i < activitys.Count; i++)
             {
                 SetCellValue(2 + i, 1, activitys[i].KeyWord);
                 SetCellValue(2 + i, 2, activitys[i].item_id);
                 SetCellValue(2 + i, 3, activitys[i].nick);
                 SetCellValue(2 + i, 4, activitys[i].name);
                 SetCellValue(2 + i, 5, activitys[i].price);
                 SetCellValue(2 + i, 6, activitys[i].priceWap);
                 SetCellValue(2 + i, 7, activitys[i].originalPrice);
                 SetCellValue(2 + i, 8, activitys[i].auctionURL);
                 SetCellValue(2 + i, 9, activitys[i].url);
                 SetCellValue(2 + i, 10, activitys[i].zkType);
                 SetCellValue(2 + i, 11, activitys[i].location);
                 SetCellValue(2 + i, 12, activitys[i].sold);
                 SetCellValue(2 + i, 13, activitys[i].commentCount);
                 SetCellValue(2 + i, 14, activitys[i].userType);
                 SetCellValue(2 + i, 15, activitys[i].area);
                 SetCellValue(2 + i, 16, activitys[i].freight);
                 SetCellValue(2 + i, 17, activitys[i].userId);
                 SetCellValue(2 + i, 18, activitys[i].isMobileEcard);
                 SetCellValue(2 + i, 19, activitys[i].shipping);
                 SetCellValue(2 + i, 20, activitys[i].fastPostFee);
                 SetCellValue(2 + i, 21, activitys[i].zkGroup);
                 SetCellValue(2 + i, 22, activitys[i].coinLimit);
                 SetCellValue(2 + i, 23, activitys[i].isB2c);
                 SetCellValue(2 + i, 24, activitys[i].iconList);
                 SetCellValue(2 + i, 25, activitys[i].category);
                 SetCellValue(2 + i, 26, activitys[i].SearthTime.ToString());
                 SetCellValue(2 + i, 27, activitys[i].UseTime);

             }
         }
     }
     */
    /*public void DataToExcel2(List<淘宝抓取工具.BaseData.ReceiveData> activitys)
    {

        if (activitys != null && activitys.Count > 0)
        {
            for (int i = 0; i < activitys.Count; i++)
            {
                SetCellValue(2 + i, 1, activitys[i].KeyWord);
                SetCellValue(2 + i, 2, activitys[i].item_id);
                SetCellValue(2 + i, 3, activitys[i].nick);
                SetCellValue(2 + i, 4, activitys[i].userType);
                SetCellValue(2 + i, 5, activitys[i].Credit);
                SetCellValue(2 + i, 6, activitys[i].location);
                
                SetCellValue(2 + i, 7, activitys[i].sold);
                SetCellValue(2 + i, 8, activitys[i].price);
                SetCellValue(2 + i, 9, activitys[i].commentCount);
                SetCellValue(2 + i, 10, activitys[i].name);
                SetCellValue(2 + i, 11, activitys[i].auctionURL);
            }
        }
    }

    */
    //public void DataToExcel(List<ActivityInfo> activitys)
    //{

    //    if (activitys != null && activitys.Count > 0)
    //    {
    //        for (int i = 0; i < activitys.Count; i++)
    //        {
    //            SetCellValue(2 + i, 1,activitys[i].ActivityType.Name );
    //            SetCellValue(2 + i, 2, activitys[i].Name);
    //            SetCellValue(2 + i, 3, WebUtility.FormatDate(activitys[i].SupportDate));
    //            SetCellValue(2 + i, 4, activitys[i].Address);
    //            SetCellValue(2 + i, 5, activitys[i].Company);
    //            SetCellValue(2 + i, 6, activitys[i].Alumni);
    //            SetCellValue(2 + i, 7, activitys[i].AlumniNums.ToString());
    //            SetCellValue(2 + i, 8, activitys[i].StudentNums.ToString());
    //            SetCellValue(2 + i, 9, activitys[i].LeaderAndTeacher);
    //            SetCellValue(2 + i, 10, activitys[i].Comment);
    //            SetCellValue(2 + i, 11, activitys[i].Unit.Name);
    //            SetCellValue(2 + i, 12, activitys[i].Poster.Name);
    //            SetCellValue(2 + i, 13, WebUtility.FormatDate(activitys[i].UpdateTime));


    //        }
    //    }


    //}
    //public void DataToExcel(List<BaseUseInfo> activitys)
    //{

    //    if (activitys != null && activitys.Count > 0)
    //    {
    //        for (int i = 0; i < activitys.Count; i++)
    //        {
    //            SetCellValue(2 + i, 1, activitys[i].BaseInfo.Name);
    //            SetCellValue(2 + i, 2, activitys[i].Name);
    //            SetCellValue(2 + i, 3, WebUtility.FormatDate(activitys[i].BeginDate));
    //            SetCellValue(2 + i, 4, WebUtility.FormatDate(activitys[i].EndDate));
    //            SetCellValue(2 + i, 5, activitys[i].Unit.Name);
    //            SetCellValue(2 + i, 6, activitys[i].StudentNums.ToString());
    //            SetCellValue(2 + i, 7, activitys[i].LeaderAndTeacher);
    //            SetCellValue(2 + i, 8, activitys[i].Comment);
    //            SetCellValue(2 + i, 9, activitys[i].Poster.Name);
    //            SetCellValue(2 + i, 10, WebUtility.FormatDate(activitys[i].UpdateTime));


    //        }
    //    }


    //}
    //public void DataToExcel(List<BaseInfo> activitys)
    //{

    //    if (activitys != null && activitys.Count > 0)
    //    {
    //        for (int i = 0; i < activitys.Count; i++)
    //        {
    //            SetCellValue(2 + i, 1, activitys[i].Name);
    //            SetCellValue(2 + i, 2, activitys[i].Company);
    //            SetCellValue(2 + i, 3, WebUtility.FormatDate(activitys[i].BuildDate));
    //            SetCellValue(2 + i, 4, activitys[i].Unit.Name);
    //            SetCellValue(2 + i, 5, activitys[i].Address);
    //            SetCellValue(2 + i, 6, activitys[i].PostNum);
    //            SetCellValue(2 + i, 7, activitys[i].Alumni);
    //            SetCellValue(2 + i, 8, activitys[i].College.Name);
    //            SetCellValue(2 + i, 9, activitys[i].Grade);
    //            SetCellValue(2 + i, 10, activitys[i].Duty);
    //            SetCellValue(2 + i, 11, activitys[i].Phone);
    //            SetCellValue(2 + i, 12, activitys[i].QQ);
    //            SetCellValue(2 + i, 13, activitys[i].Email);
    //            SetCellValue(2 + i, 14, activitys[i].ContactName);
    //            SetCellValue(2 + i, 16, activitys[i].ContactEmail);
    //            SetCellValue(2 + i, 15, activitys[i].ContactPhone);
    //            SetCellValue(2 + i, 17, activitys[i].Comment);
    //            SetCellValue(2 + i, 18, activitys[i].Poster.Name);
    //            SetCellValue(2 + i, 19, WebUtility.FormatDate(activitys[i].UpdateTime));


    //        }
    //    }


    //}
    //public void DataToExcel(List<CompanyInfo> activitys)
    //{

    //    if (activitys != null && activitys.Count > 0)
    //    {
    //        for (int i = 0; i < activitys.Count; i++)
    //        {
    //            SetCellValue(2 + i, 1, activitys[i].Name);
    //            SetCellValue(2 + i, 2, activitys[i].Keywords);
    //            SetCellValue(2 + i, 3, activitys[i].Address);
    //            SetCellValue(2 + i, 4, activitys[i].PostNum);
    //            SetCellValue(2 + i, 5, activitys[i].Website);
    //            SetCellValue(2 + i, 6, activitys[i].Alumni);
    //            SetCellValue(2 + i, 7, activitys[i].Duty);
    //            SetCellValue(2 + i, 8, activitys[i].College.Name);
    //            SetCellValue(2 + i, 9, WebUtility.FormatDate(activitys[i].GraduateDate));
    //            SetCellValue(2 + i, 10, activitys[i].Classs);
    //            SetCellValue(2 + i, 11, activitys[i].Phone);
    //            SetCellValue(2 + i, 12, activitys[i].QQ);
    //            SetCellValue(2 + i, 13, activitys[i].Email);
    //            SetCellValue(2 + i, 14, activitys[i].Comment);
    //            SetCellValue(2 + i, 15, activitys[i].Poster.Name);
    //            SetCellValue(2 + i, 16, WebUtility.FormatDate(activitys[i].UpdateTime));


    //        }
    //    }


    //}
    //public void DataToExcel(List<DirectorInfo> activitys)
    //{

    //    if (activitys != null && activitys.Count > 0)
    //    {
    //        for (int i = 0; i < activitys.Count; i++)
    //        {
    //            SetCellValue(2 + i, 1, activitys[i].Name);
    //            SetCellValue(2 + i, 2, activitys[i].Unit.Name);
    //            SetCellValue(2 + i, 3, WebUtility.FormatDate(activitys[i].EngageDate));
    //            SetCellValue(2 + i, 4, activitys[i].Title);
    //            SetCellValue(2 + i, 5, activitys[i].College.Name);
    //            SetCellValue(2 + i, 6, activitys[i].Grade);
    //            SetCellValue(2 + i, 7, activitys[i].Company);
    //            SetCellValue(2 + i, 8, activitys[i].Duty);
    //            SetCellValue(2 + i, 9, activitys[i].Phone);
    //            SetCellValue(2 + i, 10, activitys[i].QQ);
    //            SetCellValue(2 + i, 11, activitys[i].Email);
    //            SetCellValue(2 + i, 12, activitys[i].Address);
    //            SetCellValue(2 + i, 13, activitys[i].PostNum);
    //            SetCellValue(2 + i, 14, activitys[i].Classs);
    //            SetCellValue(2 + i, 19, activitys[i].Comment);
    //            SetCellValue(2 + i, 17, activitys[i].Poster.Name);
    //            SetCellValue(2 + i, 18, WebUtility.FormatDate(activitys[i].UpdateTime));


    //        }
    //    }


    //}
    //public void DataToExcel(List<DonateInfo> activitys)
    //{

    //    if (activitys != null && activitys.Count > 0)
    //    {
    //        for (int i = 0; i < activitys.Count; i++)
    //        {
    //            SetCellValue(2 + i, 1, activitys[i].Name);
    //            SetCellValue(2 + i, 2, activitys[i].Alumni);
    //            SetCellValue(2 + i, 3, activitys[i].College.Name);
    //            SetCellValue(2 + i, 4, activitys[i].Company);
    //            SetCellValue(2 + i, 5, activitys[i].DonateType.Name);
    //            SetCellValue(2 + i, 6, activitys[i].DonateName);
    //            SetCellValue(2 + i, 7, activitys[i].DonateAmount.ToString());
    //            SetCellValue(2 + i, 8, activitys[i].DonateSum.ToString());
    //            SetCellValue(2 + i, 9, WebUtility.FormatDate(activitys[i].DonateDate));
    //            SetCellValue(2 + i, 10, activitys[i].Comment);
    //            SetCellValue(2 + i, 11, activitys[i].Poster.Name);
    //            SetCellValue(2 + i, 12, WebUtility.FormatDate(activitys[i].UpdateTime));


    //        }
    //    }


    //}
    //public void DataToExcel(List<FundInfo> activitys)
    //{

    //    if (activitys != null && activitys.Count > 0)
    //    {
    //        for (int i = 0; i < activitys.Count; i++)
    //        {
    //            SetCellValue(2 + i, 1, activitys[i].Name);
    //            SetCellValue(2 + i, 2, activitys[i].Alumni);
    //            SetCellValue(2 + i, 3, activitys[i].Company);
    //            SetCellValue(2 + i, 4, activitys[i].FundSum.ToString());
    //            SetCellValue(2 + i, 5, WebUtility.FormatDate(activitys[i].BuildDate));
    //            SetCellValue(2 + i, 6, activitys[i].Unit.Name);
    //            SetCellValue(2 + i, 7, activitys[i].Comment);
    //            SetCellValue(2 + i, 8, activitys[i].Poster.Name);
    //            SetCellValue(2 + i, 9, WebUtility.FormatDate(activitys[i].UpdateTime));


    //        }
    //    }


    //}
    //public void DataToExcel(List<FundUseInfo> activitys)
    //{

    //    if (activitys != null && activitys.Count > 0)
    //    {
    //        for (int i = 0; i < activitys.Count; i++)
    //        {
    //            SetCellValue(2 + i, 1, activitys[i].Name);
    //            SetCellValue(2 + i, 2, activitys[i].Reward);
    //            SetCellValue(2 + i, 3, activitys[i].Fund.Name);
    //            SetCellValue(2 + i, 4, activitys[i].Award.ToString());
    //            SetCellValue(2 + i, 5, WebUtility.FormatDate(activitys[i].AwardDate));
    //            SetCellValue(2 + i, 6, activitys[i].Unit.Name);
    //            SetCellValue(2 + i, 7, activitys[i].Grade);
    //            SetCellValue(2 + i, 8, activitys[i].Company);
    //            SetCellValue(2 + i, 9, activitys[i].Phone);
    //            SetCellValue(2 + i, 10, activitys[i].Email);
    //            SetCellValue(2 + i, 11, activitys[i].QQ);
    //            SetCellValue(2 + i, 12, activitys[i].Poster.Name);
    //            SetCellValue(2 + i, 13, WebUtility.FormatDate(activitys[i].UpdateTime));


    //        }
    //    }


    //}

    //public void DataToExcel(List<GraduateInfo> activitys)
    //{

    //    if (activitys != null && activitys.Count > 0)
    //    {
    //        for (int i = 0; i < activitys.Count; i++)
    //        {
    //            SetCellValue(2 + i, 1, activitys[i].CollegeName);
    //            SetCellValue(2 + i, 2, activitys[i].Major);
    //            SetCellValue(2 + i, 3, activitys[i].Classs);
    //            SetCellValue(2 + i, 4, activitys[i].Id);
    //            SetCellValue(2 + i, 5, activitys[i].Name);
    //            SetCellValue(2 + i, 6, activitys[i].Sex);
    //            SetCellValue(2 + i, 7, activitys[i].Nation);
    //            SetCellValue(2 + i, 8, activitys[i].Place);
    //            SetCellValue(2 + i, 9, WebUtility.FormatDate(activitys[i].Birthday));
    //            SetCellValue(2 + i, 10, activitys[i].IdCard);
    //            SetCellValue(2 + i, 11, WebUtility.FormatDate(activitys[i].JoinDate));
    //            SetCellValue(2 + i, 12, WebUtility.FormatDate(activitys[i].GraduateDate));
    //            SetCellValue(2 + i, 13, activitys[i].ES);
    //            SetCellValue(2 + i, 14, activitys[i].EB);
    //            SetCellValue(2 + i, 15, activitys[i].Degree);
    //            SetCellValue(2 + i, 16, activitys[i].Phone);
    //            SetCellValue(2 + i, 17, activitys[i].QQ);
    //            SetCellValue(2 + i, 18, activitys[i].Email);
    //            SetCellValue(2 + i, 19, activitys[i].Address);
    //            SetCellValue(2 + i, 20, activitys[i].PostNum);
    //            SetCellValue(2 + i, 21, activitys[i].Unit);
    //            SetCellValue(2 + i, 22, activitys[i].Duty);
    //        }
    //    }


    //}
    /// <summary>
    /// 显示excel工作薄
    /// </summary>
    public void ShowExcelApp()
    {
        excelApp.Visible = true;
    }
    private void Dispose()
    {
        excelWorkBook.Close(null, null, null);
        excelApp.Workbooks.Close();
        excelApp.Quit();

        
        if (excelWorkSheet != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkSheet);
            excelWorkSheet = null;
        }
        if (excelWorkBook != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkBook);
            excelWorkBook = null;
        }
        if (excelApp != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            excelApp = null;
        }

        GC.Collect();

        

    }
    /// <summary>
    /// 将Excel文件另存为指定格式
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="format">HTML，CSV，TEXT，EXCEL，XML</param>
    public void SaveAsFile(string fileName, string format)
    {
        try
        {
            switch (format)
            {
                case "HTML":
                    {
                        excelWorkBook.SaveAs(fileName, Excel.XlFileFormat.xlHtml, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
                        break;
                    }
                case "CSV":
                    {
                        excelWorkBook.SaveAs(fileName, Excel.XlFileFormat.xlCSV, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
                        break;
                    }
                case "TEXT":
                    {
                        excelWorkBook.SaveAs(fileName, Excel.XlFileFormat.xlHtml, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
                        break;
                    }
                //					case "XML":
                //					{
                //						workBook.SaveAs(fileName,Excel.XlFileFormat.xlXMLSpreadsheet, Type.Missing, Type.Missing,
                //							Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                //							Type.Missing, Type.Missing, Type.Missing, Type.Missing,	Type.Missing);
                //						break;
                //					}
                default:
                    {
                        excelWorkBook.SaveAs(fileName, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
                        break;
                    }
            }
        }
        catch (Exception e)
        {
            throw e;
        }
        finally
        {
            this.Dispose();
        }
    }
    public void SaveAsFile()
    {
        if (this.outputFile == null)
            throw new Exception("没有指定输出文件路径！");

        try
        {
            excelWorkBook.SaveAs(outputFile, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
        }
        catch (Exception e)
        {
            //MyLog.Log("保存excel出错：" + e.Message);
        }
        finally
        {
            this.Dispose();
        }
    }
    public void SaveAsFile(string outputFile)
    {
        this.outputFile = outputFile;
        if (this.outputFile == null)
            throw new Exception("没有指定输出文件路径！");

        try
        {
            excelWorkBook.SaveAs(outputFile, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
        }
        catch (Exception e)
        {
            //MyLog.Log("保存excel出错：" + e.Message);
        }
        finally
        {
            this.Dispose();
        }
    }
    /// <summary>
    /// 根据sheet表的名字返回sheet表
    /// </summary>
    /// <param name="sheetName">sheet表名</param>
    /// <returns></returns>
    public Excel.Worksheet GetWorkSheet(string sheetName)
    {
        try
        {
            return excelWorkSheet = excelWorkBook.Worksheets[sheetName] as Excel.Worksheet;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 添加一个新工作表
    /// </summary>
    /// <param name="sheetName">新工作表名称</param>
    /// <returns></returns>
    public Excel.Worksheet AddWorkSheet(string sheetName)
    {
        excelWorkSheet = excelWorkBook.Worksheets.Add(oMissing, oMissing, oMissing, oMissing) as Excel.Worksheet ;
        excelWorkSheet.Name = sheetName;
        return excelWorkSheet;
    }

    /// <summary>
    /// 删除一个工作表
    /// </summary>
    /// <param name="sheetName"></param>
    public void DeleteSheet(string sheetName)
    {
        try
        {
            (excelWorkBook.Worksheets[sheetName] as Excel.Worksheet).Delete();
        }
        catch
        {
        }
    }

    /// <summary>
    /// 重命名工作表
    /// </summary>
    /// <param name="ws"></param>
    /// <param name="newName"></param>
    /// <returns></returns>
    public Excel.Worksheet ReNameSheet(Excel.Worksheet sheet, string newName)
    {
        sheet.Name = newName;
        return sheet;
    }

    /// <summ个ary>
    /// 读取某单元格的值
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="colunm">列号</param>
    /// <returns></returns>
    public string ReadCellValue(int row, int column)
    {
        Excel.Range range = excelWorkSheet.Cells[row, column] as Excel.Range;
        return range.Text.ToString();
    }

    /// <summary>
    /// 获取单元格
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="colunm">列号</param>
    /// <returns>单元格</returns>
    public Excel.Range GetCell(int row, int column)
    {
        return excelWorkSheet.Cells[row, column] as Excel.Range;
    }

    /// <summary>
    /// 获取区域
    /// </summary>
    /// <param name="row1">行号1</param>
    /// <param name="column1">列号1</param>
    /// <param name="row2">行号2</param>
    /// <param name="column2">列号2</param>
    /// <returns>区域</returns>
    public Excel.Range GetCells(int row1, int column1, int row2, int column2)
    {
        return excelWorkSheet.get_Range(excelWorkSheet.Cells[row1, column1], excelWorkSheet.Cells[row2, column2]) as Excel.Range;
    }


    /// <summary>
    /// 设置某一个单元格的值
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="column">列号</param>
    /// <param name="value">值</param>
    public void SetCellValue(int row, int column, string value)
    {
        Excel.Range range = excelWorkSheet.Cells[row, column] as Excel.Range;
        range.Value2 = value;
    }
    /// <summary>
    /// 设置某一区域单元格的值
    /// </summary>
    /// <param name="row1">行号1</param>
    /// <param name="column1">列号1</param>
    /// <param name="value">值</param>
    public void SetCellValue(int row1, int column1,int row2, int column2, string value)
    {
        Excel.Range range = excelWorkSheet.get_Range(excelWorkSheet.Cells[row1, column1], excelWorkSheet.Cells[row2, column2]);
        range.Value2 = value;
    }
    /// <summary>
    /// 设置单元格边框
    /// </summary>
    /// <param name="row"></param>
    /// <param name="column"></param>
    public void SetCellBorder(int row, int column)
    {
        Excel.Range range = excelWorkSheet.Cells[row, column] as Excel.Range;
        range.Borders.LineStyle = 1;
        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;//设置左边线加粗   
        range.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;//设置上边线加粗   
        range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;//设置右边线加粗   
        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;//设置下边线加粗
    }
    /// <summary>
    /// 设置范围边框
    /// </summary>
    /// <param name="row1"></param>
    /// <param name="column1"></param>
    /// <param name="row2"></param>
    /// <param name="column2"></param>
    public void SetCellBorder(int row1, int column1, int row2, int column2)
    {
        Excel.Range range = excelWorkSheet.get_Range(excelWorkSheet.Cells[row1, column1], excelWorkSheet.Cells[row2, column2]);
        range.Borders.LineStyle = 1;
        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;//设置左边线加粗   
        range.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;//设置上边线加粗   
        range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;//设置右边线加粗   
        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;//设置下边线加粗
    }
    /// <summary>
    /// 设置单元格自动换行
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="column">列号</param>
    /// <param name="isWrap">是否换行</param>
    public void SetCellWrapText(int row, int column, bool isWrap)
    {
        Excel.Range range = excelWorkSheet.Cells[row, column] as Excel.Range;
        range.WrapText = isWrap;
        if (isWrap)
            range.EntireRow.AutoFit();
    }
    /// <summary>
    /// 设置区域自动换行
    /// </summary>
    /// <param name="row1">行号1</param>
    /// <param name="column1">列号1</param>
    /// <param name="row2">行号2</param>
    /// <param name="column2">列号2</param>
    /// <param name="isWrap">是否换行</param>
    public void SetCellWrapText(int row1, int column1, int row2, int column2, bool isWrap)
    {
        Excel.Range range = excelWorkSheet.get_Range(excelWorkSheet.Cells[row1, column1], excelWorkSheet.Cells[row2, column2]);
        range.WrapText = isWrap;
    }
    public void SetCellBold(int row, int column)
    {
        Excel.Range range = excelWorkSheet.Cells[row, column] as Excel.Range;
        range.Font.Bold = true;
    }
    public void SetCellBold(int row1, int column1, int row2, int column2)
    {
        Excel.Range range = excelWorkSheet.get_Range(excelWorkSheet.Cells[row1, column1], excelWorkSheet.Cells[row2, column2]);
        range.Font.Bold = true;
    }
    /// <summary>
    /// 合并或分割区域
    /// </summary>
    /// <param name="row1">行号1</param>
    /// <param name="column1">列号1</param>
    /// <param name="row2">行号2</param>
    /// <param name="column2">列号2</param>
    /// <param name="merge">是否合并</param>
    public void SetCellMerge(int row1, int column1, int row2, int column2, bool merge)
    {
        Excel.Range range = excelWorkSheet.get_Range(excelWorkSheet.Cells[row1, column1], excelWorkSheet.Cells[row2, column2]);
        if (merge)
            range.Merge();
        else
            range.UnMerge();
    }
    /// <summary>
    /// 设置行高
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="height">高度</param>
    public void SetRowHeight(int row, decimal height)
    {
        Excel.Range range = excelWorkSheet.Cells[row, 1] as Excel.Range;
        range.RowHeight = height;
    }
    /// <summary>
    /// 设置行高
    /// </summary>
    /// <param name="row1">行号1</param>
    /// <param name="row2">行号2</param>
    /// <param name="height">高度</param>
    public void SetRowHeight(int row1, int row2, decimal height)
    {
        Excel.Range range = excelWorkSheet.get_Range(excelWorkSheet.Cells[row1, 1], excelWorkSheet.Cells[row2, 1]);
        range.RowHeight = height;
    }
    public void SetRangeHeightAutoFit(int row, int column, decimal minHeight)
    {
        Excel.Range range = excelWorkSheet.Cells[row, column] as Excel.Range;
        range.EntireRow.AutoFit();
        if (Convert.ToDecimal(range.RowHeight) < minHeight)
            range.RowHeight = minHeight;
    }
    public void SetRangeHeightAutoFit(int row1, int column1, int row2, int column2, decimal minHeight)
    {
        Excel.Range range = excelWorkSheet.get_Range(excelWorkSheet.Cells[row1, column1], excelWorkSheet.Cells[row2, column2]);
        decimal rangeWidth,firstCellWidth,autoFitHeight;
        Excel.Range firstCell = range.Cells[1, 1] as Excel.Range;
        int columns = column2- column1;
        rangeWidth = 0;
        for(int i = 1 ;i<= columns;i++)
            rangeWidth += Convert.ToDecimal((range.Cells[1,i] as Excel.Range).ColumnWidth);
        firstCellWidth = Convert.ToDecimal(firstCell.ColumnWidth);
        range.UnMerge();
        firstCell.ColumnWidth = rangeWidth;
        range.EntireRow.AutoFit();
        autoFitHeight = Convert.ToDecimal(range.RowHeight);
        if (minHeight > 0 && autoFitHeight < minHeight)
            autoFitHeight = minHeight;
        firstCell.ColumnWidth = firstCellWidth;
        range.Merge(true);
        range.RowHeight = autoFitHeight;
    }
    /// <summary>
    /// 获取行高
    /// </summary>
    /// <param name="row">行号</param>
    public decimal GetRowHeight(int row)
    {
        Excel.Range range = excelWorkSheet.Cells[row, 1] as Excel.Range;
        return Convert.ToDecimal(range.RowHeight);
    }

    public decimal GetRowHeight(int row1, int row2)
    {
        Excel.Range range = excelWorkSheet.get_Range(excelWorkSheet.Cells[row1, 1], excelWorkSheet.Cells[row2, 1]);
        return Convert.ToDecimal(range.RowHeight);
    }
    /// <summary>
    /// 设置列宽
    /// </summary>
    /// <param name="column">列号</param>
    /// <param name="width">宽度</param>
    public void SetColumnWidth(int column, decimal width)
    {
        Excel.Range range = excelWorkSheet.Cells[1, column] as Excel.Range;
        range.ColumnWidth = width;
    }
    /// <summary>
    /// 设置列宽
    /// </summary>
    /// <param name="column1">列号1</param>
    /// <param name="column2">列号2</param>
    /// <param name="width">宽度</param>
    public void SetColumnWidth(int column1, int column2, decimal width)
    {
        Excel.Range range = excelWorkSheet.get_Range(excelWorkSheet.Cells[1, column1], excelWorkSheet.Cells[1, column2]);
        range.ColumnWidth = width;
    }
    /// <summary>
    /// 设置单元格水平对齐
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="column">列号</param>
    /// <param name="align">对齐方式：0 右 1 左 2 两端 3 分散 4 居中 5 默认(数据类型) 6 填充 7 跨列</param>
    /// <param name="indent">缩进</param>
    public void SetHorizontalAlignment(int row, int column, int align, int indent)
    {
        Excel.Range range = excelWorkSheet.Cells[row, column] as Excel.Range;
        switch (align)
        {
            case 0:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                break;
            case 1:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                break;
            case 2:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify;
                break;
            case 3:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignDistributed;
                break;
            case 4:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                break;
            case 5:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral;
                break;
            case 6:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignFill;
                break;
            case 7:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenterAcrossSelection;
                break;
        }
        if (indent <= 15 && indent >= 0)
            range.IndentLevel = indent;
    }
    /// <summary>
    /// 设置单元格水平对齐
    /// </summary>
    /// <param name="row1">行号1</param>
    /// <param name="column1">列号1</param>
    /// <param name="row2">行号2</param>
    /// <param name="column2">列号2</param>
    /// <param name="align">对齐方式：0 右 1 左 2 两端 3 分散 4 居中 5 默认(数据类型) 6 填充 7 跨列</param>
    /// <param name="indent">缩进</param>
    public void SetHorizontalAlignment(int row1, int column1, int row2, int column2, int align, int indent)
    {
        Excel.Range range = excelWorkSheet.get_Range(excelWorkSheet.Cells[row1, column1], excelWorkSheet.Cells[row2, column2]);
        switch (align)
        {
            case 0:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                break;
            case 1:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                break;
            case 2:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify;
                break;
            case 3:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignDistributed;
                break;
            case 4:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                break;
            case 5:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral;
                break;
            case 6:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignFill;
                break;
            case 7:
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenterAcrossSelection;
                break;
        }
        if (indent <= 15 && indent >= 0)
            range.IndentLevel = indent;
    }
    public void KillExcelProcess()
    {
        Process[] myProcesses;
        DateTime startTime;
        myProcesses = Process.GetProcessesByName("Excel");

        //得不到Excel进程ID，暂时只能判断进程启动时间
        foreach (Process myProcess in myProcesses)
        {
            startTime = myProcess.StartTime;

            if (startTime > beforeTime && startTime < afterTime)
            {
                myProcess.Kill();
            }
        }
    }
    /// <summary>
    /// 关闭excel对象
    /// </summary>
    /// <param name="isSave"></param>
    /// <param name="fileName"></param>
    public void CloseExcel(bool isSave, string fileName)
    {
        
        if (excelWorkBook != null)
        {
            if (isSave)
            {
                if (excelWorkBook.Name == fileName)
                    excelWorkBook.Save();
                //else
                //    excelWorkBook.SaveAs(fileName,oMissing,oMissing,oMissing,oMissing,oMissing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,oMissing,oMissing,oMissing,oMissing,oMissing);
            }
            else
            {
                excelWorkBook.Close(false, oMissing, oMissing);
            }
        }
        if (excelWorkSheet != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkSheet);
        }
        if (excelWorkBook != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkBook);
        }
        if (excelWorkBooks != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkBooks);
        }
        if (excelApp != null)
        {
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
        GC.Collect();
        this.KillExcelProcess();
    }

        

}


