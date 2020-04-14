using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Report.NPOI
{
    public static class ICellExtension
    {
        /// <summary>
        /// 根据数据类型决定保存类型
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void SetCellValueObject(this ICell cell, object? value)
        {
            if (value == null)
                cell.SetCellValue("");
            else if (value is string)
                cell.SetCellValue((string)value);
            else if (value is bool)
                cell.SetCellValue((bool)value);
            else if (value is int)
                cell.SetCellValue(Convert.ToDouble(value));
            else if (value is double)
                cell.SetCellValue((double)value);
            else if (value is decimal)
                cell.SetCellValue(Convert.ToDouble(value));
            else if (value is float)
                cell.SetCellValue(Convert.ToDouble(value));
            else if (value is DateTime)
                cell.SetCellValue((DateTime)value);
            else
                cell.SetCellValue(value.ToString());
        }
    }
}
