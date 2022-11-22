using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

public class TableCollector
{
    private List<TableHeader> m_TableHeaders = new List<TableHeader>();

    public TableCollector(){ CleanUp();}
    void CleanUp() { m_TableHeaders.Clear(); }
    public void CollectionHeaders(string tableFolder)
    {
        if (!Directory.Exists(tableFolder))
        {
            LogUtil.LogDebug("CollectiHeaders Failed. Folder:{0}", tableFolder);
            return;
        }
        RecursionReadTable(new DirectoryInfo(tableFolder));
    }
    public void UpdateRulerExcel()
    {
        string testExcel = "D:/HappyConfig.xlsx";
        //ExcelUtils.WriteData2Excel(testExcel, 1, new List<string> { "职业", "年龄", "性比", "身高" });
        ExcelUtils.WriteData2Exce(testExcel, m_TableHeaders);
        LogUtil.LogDebug("Update Rule SUCC.");
    }
    private void RecursionReadTable(DirectoryInfo dirInfo)
    {
        if(null == dirInfo)
        {
            return;
        }
        //File filelist = Directory.GetFiles(tableFolder);
        DirectoryInfo curDirInfo = dirInfo;
        FileInfo[] childFileList = curDirInfo.GetFiles();
        for(int idx=0;idx<childFileList.Length;++idx)
        {
            if(childFileList[idx].Extension.Contains("txt"))
                ReadTable(childFileList[idx]);
        }
        DirectoryInfo[] childDirInfo = curDirInfo.GetDirectories();
        for(int idx=0;idx<childDirInfo.Length;++idx)
        {
            RecursionReadTable(childDirInfo[idx]);
        }
    }

    private void ReadTable(FileInfo fInfo)
    {
        TableHeader header = new TableHeader(fInfo);
        using (StreamReader reader = fInfo.OpenText())
        {
            string line = "";
            CellType lineType = CellType.C_NAME;
            while ((line = reader.ReadLine()) != null)
            {
                header.Init(line, lineType);
                lineType = lineType + 1;

                if (lineType >= CellType.C_MAX)
                {
                    break;
                }
            }
            reader.Close();
        }
        if(header.IsValid())
        {
            m_TableHeaders.Add(header); 
        }
        else
        {
            LogUtil.LogDebug("ReadTable Failed, fInfo:{0}", fInfo.FullName);
        }
    }
}
