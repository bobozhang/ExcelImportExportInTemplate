using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;


public partial class Default2 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        
    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        string filePath = Server.MapPath("~/App_Data/zhangbo.xlsx");
        getSheetValue(filePath);  //读取zhangbo.xlsx工作簿的sheet中对应cell的值

    }

    private  void getSheetValue(string filePath)  //读取Sheet中字符串的数据
    {
        using( SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
        {
            string strSheet = "zhangbo";  //sheet名字
            object obj;
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == strSheet);
            WorkbookPart workBookPart = document.WorkbookPart;   //获取Workbookpart
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            SharedStringTablePart tablePart = document.WorkbookPart.SharedStringTablePart;
            Worksheet worksheet = worksheetPart.Worksheet;
            IEnumerable<Row> rows = worksheet.Descendants<Row>();  // 根据WorkbookPart和sheetName获取该Sheet下所有Row数据
           // IEnumerable<Row> rows = GetWorkBookPartRows(workBookPart, sheetnames.First());
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            string cellName = "";
            foreach (Row row in rows)
            {
                uint rowIndex = row.RowIndex;
                foreach (Cell cell in row)
                {
                    obj = GetValue(cell, tablePart);   //单元格的值
                    cellName = cell.CellReference.ToString();   //单元格序号
                    string indexCol = cell.CellReference.ToString().Substring(0,1);  //列号
                  //  if (indexCol.CompareTo("W")>0) break;
                }
            }


            List<string> sheetnames = GetSheetNames(workBookPart);
          
        }
    }

    /// <summary>
    /// 根据WorkbookPart获取所有SheetName
    /// </summary>
    /// <param name="workBookPart"></param>
    /// <returns>SheetName集合</returns>
    private  List<string> GetSheetNames(WorkbookPart workBookPart)
    {
        List<string> sheetNames = new List<string>();
        Sheets sheets = workBookPart.Workbook.Sheets;
        foreach (Sheet sheet in sheets)
        {
            string sheetName = sheet.Name;
            if (!string.IsNullOrEmpty(sheetName))
            {
                sheetNames.Add(sheetName);
            }
        }
        return sheetNames;
    }

    /// <summary>
    /// 根据WorkbookPart和sheetName获取该Sheet下所有Row数据
    /// </summary>
    /// <param name="workBookPart">WorkbookPart对象</param>
    /// <param name="sheetName">SheetName</param>
    /// <returns>该SheetName下的所有Row数据</returns>
    public IEnumerable<Row> GetWorkBookPartRows(WorkbookPart workBookPart, string sheetName)
    {
        IEnumerable<Row> sheetRows = null;
        //根据表名在WorkbookPart中获取Sheet集合
        IEnumerable<Sheet> sheets = workBookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);
        if (sheets.Count() == 0)
        {
            return null;//没有数据
        }

        WorksheetPart workSheetPart = workBookPart.GetPartById(sheets.First().Id) as WorksheetPart;
        //获取Excel中得到的行
        sheetRows = workSheetPart.Worksheet.Descendants<Row>();

        return sheetRows;
    }


    public static String GetValue(Cell cell, SharedStringTablePart stringTablePart)
    {
        if (cell.ChildElements.Count == 0)
            return null;
        //get cell value
        String value = cell.CellValue.InnerText;
        //Look up real value from shared string table
        if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            value = stringTablePart.SharedStringTable
            .ChildElements[Int32.Parse(value)]
            .InnerText;
        return value;
    }

}