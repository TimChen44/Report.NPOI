using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace Report.NPOI
{

    /// <summary>
    /// 标签，用于标记单元格的用于
    /// 格式：$+类型
    /// </summary>
    public abstract class Label
    {
        public Label(ICell cell)
        {
            this.Cell = cell;
        }

        public ICell Cell { get; set; }

        public abstract string AddressString();
    }

    /// <summary>
    /// 替换标记
    /// 格式：$:[字段名称]
    /// 例如：$:Company
    /// </summary>
    public class ReplaceLabel : Label
    {
        public ReplaceLabel(ICell cell, string label) : base(cell)
        {
            this.Name = label.Substring(2);
        }

        /// <summary>
        /// 标签名称
        /// </summary>
        public string Name { get; set; }

        public override string ToString() => $"{Name}";

        public override string AddressString()
        {
            return Cell.Address.FormatAsString();
        }

    }

    /// <summary>
    /// 表格标签
    /// 格式：$.[表名].[字段名]
    /// 例如：$.Table.UserName
    /// </summary>
    public class TableLabel : Label, IFormulaLabel
    {
        public TableLabel(ICell cell, string label) : base(cell)
        {
            label = label.Substring(2);
            this.Table = label.Substring(0, label.IndexOf("."));

            label = label.Substring(label.IndexOf(".") + 1);

            if (label.Contains("=") == true)
            {
                this.Name = label.Substring(0, label.IndexOf("="));
                this.Formula = label.Substring(label.IndexOf("=") + 1);
            }
            else
            {
                this.Name = label;
            }
        }

        /// <summary>
        /// 主体，这个标签属于谁，比如属于那个表格，默认是空
        /// </summary>
        public string Table { get; set; }
        /// <summary>
        /// 标签名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 公式
        /// </summary>
        public string? Formula { get; set; } = null;

        /// <summary>
        /// 记录当前对象填充的数据的范围
        /// </summary>
        public CellRangeAddress? CellRange { get; set; }

        public override string ToString() => $"{Table}.{Name}";

        public override string AddressString()
        {
            return CellRange?.FormatAsString() ?? Cell.Address.FormatAsString();
        }
    }

    /// <summary>
    /// 公式标签
    /// 格式：$=[公式]
    /// 例如：$=SUM(
    /// </summary>
    public class FormulaLabel : Label, IFormulaLabel
    {
        public FormulaLabel(ICell cell, string label) : base(cell)
        {
            this.Formula = label.Substring(2);
        }

        /// <summary>
        /// 公式
        /// </summary>
        public string Formula { get; set; }

        public override string ToString() => Formula;

        public override string AddressString()
        {
            return Cell.Address.FormatAsString();
        }
    }

    public interface IFormulaLabel
    {
        string Formula { get; set; }
    }
}
