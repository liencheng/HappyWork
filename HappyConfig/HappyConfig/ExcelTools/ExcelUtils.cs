using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using System.IO;

public class ExcelUtils
{
    public static _Worksheet OpenWorkSheet(string filePath, int idx)
    {
        if(!File.Exists(filePath))
        {
            LogUtil.LogDebug("File Not Exist,{0}", filePath);
            return null;
        }
        Application oApp = new Application();
        Workbooks oBooks= oApp.Workbooks;
        _Workbook book = oBooks.Add(filePath);
        _Worksheet oSheet = book.ActiveSheet;
        return oSheet;
        
    }

    public static bool WriteData2Exce(string filePath, List<TableHeader> heads)
    {
        for (int idx = 0; idx < heads.Count; ++idx)
        {

        }
        if (!File.Exists(filePath))
        {
            LogUtil.LogDebug("File Not Exist,{0}", filePath);
            return false;
        }
        Application oApp = new Application();
        Workbooks oBooks = oApp.Workbooks;
        _Workbook book = oBooks.Add(filePath);
        _Worksheet sheet = book.ActiveSheet;

        if (null == sheet)
        {
            LogUtil.LogDebug("Open Sheet Failed, Sheet:{0}", filePath);
            return false;
        }
        for (int idx = 0; idx < heads.Count; ++idx)
        {
            int lineNum = idx * TableHeader.C_LINE_PER_TABLE + 1;
            int C_PRE_CELL = 3;
            sheet.Cells[lineNum, 1] = idx;
            sheet.Cells[lineNum, 2] = heads[idx].TableName;
            sheet.Cells[lineNum, 3] = heads[idx].TableFullPath;
            List<HeaderCell> cols = heads[idx].GetCols();
            for (var col = 0; col < cols.Count; ++col)
            {
                sheet.Cells[lineNum, col + 1 + C_PRE_CELL] = cols[col].Name;
                sheet.Cells[lineNum+1, col + 1 + C_PRE_CELL] = cols[col].DataType;
                sheet.Cells[lineNum+2, col + 1 + C_PRE_CELL] = cols[col].Comment;
            }
        }

        book.SaveAs(filePath);

        CloseExcel(oApp, oBooks, book, sheet);
        return true;
    }
    public static bool WriteData2Excel(string filePath, int sheetIdx, List<string> rowData)
    {
        if(!File.Exists(filePath))
        {
            LogUtil.LogDebug("File Not Exist,{0}", filePath);
            return false;
        }
        if(sheetIdx<=0)
        {
            LogUtil.LogDebug("sheetIndex <=0, Please Fix It.");
        }
        Application oApp = new Application();
        {
            Workbooks oBooks = oApp.Workbooks;
            _Workbook book = oBooks.Add(filePath);
            _Worksheet sheet = book.ActiveSheet;

            if (null == sheet)
            {
                LogUtil.LogDebug("Open Sheet Failed, Sheet:{0}", filePath);
                return false;
            }
            if (null == rowData)
            {
                LogUtil.LogDebug("WriteData2Excel Failed. rowData is Null.");
                return false;
            }
            for (var idx = 0; idx < rowData.Count; ++idx)
            {
                sheet.Cells[sheetIdx, idx+1] = rowData[idx];
            }

            book.SaveAs(filePath);

            CloseExcel(oApp, oBooks, book, sheet);
        }
        return true;
    }
    public static  void CloseExcel(Application app, Workbooks books, _Workbook book, _Worksheet sheet)
    {
        ReleaseRes(sheet);
        book.Close(false);
        ReleaseRes(book);
        ReleaseRes(books);
        app.Quit();
        ReleaseRes(app);
    }
    public static void ReleaseRes(Object o)
    {
        try
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o);       //使用此方法，来释放引用某些资源的基础 COM 对象。
                                                                              //这里的o就是要释放的对象
        }
        catch { 
        }
        finally { o = null; }
    }
    public static bool CreateExcel(string filePath, string bForceCreate)
    {
        Application oApp = new Application();
        oApp.Visible = true;
        _Workbook oBook = oApp.Workbooks.Add();
        _Worksheet oSheet = oBook.ActiveSheet;
        oBook.SaveAs(filePath);
        LogUtil.LogDebug(filePath);
        return true;
    }
}
