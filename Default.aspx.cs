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

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        string filePath = Server.MapPath("~/App_Data/zhangbo-demo.xlsx");
        insertDataToExcel(filePath);

    }

    public void insertDataToExcel(string fileName)  //向 fileName工作簿中的Sheet1文件中插入数据
    {
        InsertText(fileName,"Sheet1", "520202337327965","D",5);
        InsertText(fileName, "Sheet1", "盘县火铺景美农家乐餐馆", "J", 5);
        InsertText(fileName, "Sheet1", "520202337327965", "Q", 5);
        InsertText(fileName, "Sheet1", "中国银行", "D", 6);
        InsertText(fileName, "Sheet1", "520202337327965", "J", 6);
        InsertText(fileName, "Sheet1", "520202337327965", "Q", 6);
        InsertText(fileName, "Sheet1", "520202337327965", "D", 7);

        char rowName = 'A';  //写入Sheet的行名
        int colNum = 1;    //写入Sheet的列号
        uint rowi = 0;     //记录行
        uint colj = 0;       //记录列
        rowi = (uint)getRowNumbers(fileName,"Sheet1");
        int tiaoj = getRowNumbers(fileName, "Sheet1") + 1;


        for (rowi = rowi + 1; rowi <= tiaoj; rowi++)  //新增行数自己设置
        {

           
            for (colj = 1; colj < 22; colj++)
            {
                char rowNames = (char)('A' + colj - 1);
                string inserString = "["+rowNames +"," + rowi.ToString() +"]";
                InsertText(fileName, "Sheet1", inserString, rowNames.ToString(), rowi);
            }
        }


    }


    public static int getRowNumbers(string filePath, string sheetName)
    {
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
        {
            string strSheet = sheetName;    //sheet的名字
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == strSheet);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return 0;
            }
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            SharedStringTablePart tablePart = document.WorkbookPart.SharedStringTablePart;
            Worksheet worksheet = worksheetPart.Worksheet;
            IEnumerable<Row> rows = worksheet.Descendants<Row>();
            return rows.Count();
        }
    }

    // Given a document name and text, 
    //  writes the text to cell [rowName,colNum] of the worksheet.
    //writes the text to cell
    public static void InsertText(string docName,string sheetName, string text,string rowName,uint colNum)
    {
        // Open the document for editing.从文档中创建类实例，请调用 Open 重载方法之一。
        //第一个参数采用表示要打开的文档的完整路径字符串。第二个参数是 true 或 false，表示是否要打开文件以进行编辑。如果此参数为 false，则不会保存对该文档所做的任何更改。

        using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))  
        {
          
            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem(text, shareStringPart);

            string strSheet = sheetName;    //strSheet为读写sheet的名字
            //sheets 接口用来存放需要读写的sheet
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == strSheet);

            //  获取工作表的worksheetPart
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet(rowName, colNum, worksheetPart);
            // Set the value of cell A1.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            // Save the new worksheet.
            worksheetPart.Worksheet.Save();
        }
    }
    // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create one.
        if (shareStringPart.SharedStringTable == null)
        {
            shareStringPart.SharedStringTable = new SharedStringTable();
        }
        int i = 0;
        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return i;
            }
            i++;
        }
        // The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        shareStringPart.SharedStringTable.Save();
        return i;
    }

    // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    // If the cell already exists, returns it. 
    private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
    {
        Worksheet worksheet = worksheetPart.Worksheet;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;
        if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }
        // If there is not a cell with the specified column name, insert one.  
        if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }
            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);
            worksheet.Save();
            return newCell;
        }
    }


}