using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

/// <summary>
/// 需要严格和表格匹配上
/// </summary>
public enum CellType
{
    C_NAME,
    C_TYPE,
    C_LIMIT,
    C_COMMENT,
    C_MAX
}

public class HeaderCell
{
    public string Name { get; set; } = "NaN";
    public string Comment { get; set; } = "NaN";
    public string DataType { get; set; } = "NaN";

    public void SetVal(string val, CellType type)
    {
        switch(type)
        {
            case CellType.C_NAME:
                Name = val;
                break;
            case CellType.C_COMMENT:
                Comment = val;
                break;
            case CellType.C_TYPE:
                DataType = val;
                break;
            default:
                LogUtil.LogDebug("Header Cell, Type Error. val:{0}, type:{1}", val, type.ToString());
                break;
        }
    }
}
public class TableHeader
{
    public const int C_LINE_PER_TABLE = 4;
    public string TableName { get; set; }
    public string TableFullPath { get; set; }
    private List<HeaderCell> Cols = new List<HeaderCell>();

    public List<HeaderCell> GetCols()
    {
        return Cols;
    }
    public TableHeader(FileInfo fi)
    {
        TableName = fi.Name;
        TableFullPath = fi.FullName;
    }
    void MakeCols(int colnum)
    {
        for (int idx = Cols.Count; idx < colnum; ++idx)
        {
            Cols.Add(new HeaderCell());
        }
    }
    public void Init(string header, CellType cellType)
    {
        string []line_split = header.Split('\t', '\0');
        MakeCols(line_split.Length);
        
        for(int idx=0;idx<line_split.Length;++idx)
        {
            Cols[idx].SetVal(line_split[idx], cellType);
        }
    }

    public bool IsValid()
    {
        return Cols.Count > 0;
    }
}
