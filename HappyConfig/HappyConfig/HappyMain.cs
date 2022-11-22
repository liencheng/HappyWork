using System;
using System.Collections.Generic;

namespace HappyConfig
{
    class HappyMain
    {
        const string C_CLINET_TABLE = @"F:\Project\Main\Client\Assets\Game\Resources\Bundle\Table";
        const string C_PUBLIC_TABLE = "F:/Project/Main/Public/PublicTables";
        const string C_SERVER_TABLE = "F:/Project/Main/Public/ServerTables";
        TableCollector collector = new TableCollector();
        static void Main(string[] args)
        {
            //string testExcel = "D:/HappyConfig.xlsx";
            //ExcelUtils.CreateExcel(testExcel, "false");
            //ExcelUtils.WriteData2Excel(testExcel, 1, new List<string>{ "职业", "年龄", "性比", "身高"});
            HappyMain happyMain = new HappyMain();
            happyMain.ExtraTable();
            LogUtil.LogDebug("Hello World.");
        }

        public void ExtraTable()
        {
            collector.CollectionHeaders(C_CLINET_TABLE);
            collector.UpdateRulerExcel();
        }
    }
}
