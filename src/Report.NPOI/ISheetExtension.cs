using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace Report.NPOI
{
    public static class ISheetExtension
    {
        public static bool WriteReport(this ISheet sheet, ReportDataSource reportDataSource)
        {
            //获得表格中的标签信息
            LabelSet labelSet = new LabelSet(sheet);

            //进行标签替换
            labelSet.WriteReplace(reportDataSource.Data);

            //进行表格填充
            foreach (var table in reportDataSource.Tables)
            {
                labelSet.WriteTable(table.Key, table.Value);
            }

            labelSet.WriteFormula();

            return true;
        }

        /// <summary>
        /// 快速获得单元格
        /// </summary>
        public static ICell? GetCell(this ISheet sheet, int rownum, int cellnum)
        {
            return sheet.GetRow(rownum)?.GetCell(cellnum);
        }
    }

}


