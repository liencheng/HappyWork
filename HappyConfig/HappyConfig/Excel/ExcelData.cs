using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using System.IO;


public class ExcelRow
{
    public ExcelRow()
    {
        Clean();
    }
    public void Clean()
    {
        ColNum = 0;
        TableName = "";
        TableFullPath = "";
        ColName = new List<string>();
    }
    public void AddCol(string name)
    {
        ColName.Add(name);
    }
    public int ColNum { get; set; } = 0;
    public int TableName { get; set; } = "";
    public int TableFullPath { get; set; } = "";
    public List<string> ColName { get; set; } = new List<string>();
}


public class ExcelTableUnit
{
    public string TableName { get; set; } = "";
    public string TableFullPath { get; set; } = "";

    public List<ExcelRow> Rows { get; set; } = new List<ExcelRow>();
}


public class ExcelData
{
    private Application m_App;
    private Workbooks m_Books;
    private _Workbook m_Book;
    private _Worksheet m_Sheet;
    private DataTable m_Dt;
    private List<ExcelTableUnit> m_TableUnits { get; set; } = new List<ExcelTableUnit>();
    private ExcelTableUnit m_TmpUnit;

    private bool bExcelValid { get; set; } = false;
    private string FileName { get; set; } = "";
    private string FilePath { get; set; } = "";
    public ExcelData(string filepath)
    {
        if (!File.Exists(filepath))
        {
            bExcelValid = false;
            return;
        }
        m_App = new Application();
        m_Books= m_App.Workbooks;
        m_Book = m_Books.Add(filepath);
        m_Sheet = m_Book.ActiveSheet;
        bExcelValid = true;     
    }

    ~ExcelData()
    {
        CloseExcel();
    }
    public void ReadExcel()
    {
       if(!bExcelValid)
        {
            LogUtil.LogDebug("ReadExcel Failded.File:{0}", FilePath);
            return;
        }
        int nRowCt = m_Sheet.UsedRange.Rows.Count;
        int nColCt = m_Sheet.UsedRange.Columns.Count;

        LogUtil.LogDebug("ReadExcel, RowCount:{0}, ColCount{1}", nRowCt.ToString(), nColCt.ToString());
        
        //先遍历行
        for(int row= 0;row<nRowCt;++row)
        {
            //再遍历列
            ReadRow(row + 1, nColCt);
        }
    }

    string GetCellVal(int nRow, int nCol)
    {
        return ((Excel.Range)m_Sheet.Cells[nRow, nCol]).Text.ToString();
    }
    void ReadRow(int nRowIndex, int nColMax)
    {
        string idVal = GetCellVal(nRowIndex, Define.C_ExcelIdCol);
        string tableNameVal = GetCellVal(nRowIndex, Define.C_ExcelNameCol);
        string tableFullPathVal = GetCellVal(nRowIndex, Define.C_ExcelFullPathCol);

        bool bTableStartRow = false;
        if(!string.IsNullOrEmpty(idVal))
        {
            bTableStartRow = true;
        }
        ExcelRow row = new ExcelRow(); 
        for (int col = Define.C_ExcelBeginCol; col <= nColMax; ++col)
        {
            string cellval = GetCellVal(nRowIndex, col);
            row.AddCol(cellval);
        }
        ExcelTableUnit tableUnit;
        if(bTableStartRow || m_TmpUnit == null)
        {
            tableUnit = new ExcelTableUnit();
            m_TmpUnit = tableUnit;
            m_TmpUnit.TableName = tableNameVal;
            m_TmpUnit.TableFullPath = tableFullPathVal;
        }
        m_TmpUnit.Rows.Add(row);

        LogUtil.LogDebug("ReadRowl");
    }

 
    void CloseExcel()
    {
        ReleaseRes(m_Sheet);
        m_Book.Close(false);
        ReleaseRes(m_Book);
        ReleaseRes(m_Books);
        m_App.Quit();
        ReleaseRes(m_App);
    }
   void ReleaseRes(Object o)
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
    
}
