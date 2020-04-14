using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace Report.NPOI
{
    public static class ISheetExtension
    {
        public static void WriteReport(this ISheet sheet, ReportDataSource reportDataSource)
        {
            //获得表格中的标签信息
            LabelSet labelSet = new LabelSet(sheet);
            labelSet.Fill(reportDataSource);
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


