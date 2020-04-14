using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace Report.NPOI
{

    /// <summary>
    /// 标签集合，也是标签处工具
    /// </summary>
    public class LabelSet
    {
        //替换标签集合
        public List<ReplaceLabel> Replaces { get; set; } = new List<ReplaceLabel>();

        //表格标签集合
        public List<TableLabel> Tables { get; set; } = new List<TableLabel>();

        //公式标签集合
        public List<FormulaLabel> Formulas { get; set; } = new List<FormulaLabel>();

        ISheet Sheet;

        public LabelSet(ISheet sheet)
        {
            Sheet = sheet;
            //查找所有标签
            InitLabel();
        }

        /// <summary>
        ///  //查找所有标签
        /// </summary>
        private void InitLabel()
        {
            for (int rowIndex = 0; rowIndex <= Sheet.LastRowNum; rowIndex++)
            {
                var row = Sheet.GetRow(rowIndex);

                for (int cellIndex = 0; cellIndex < row?.LastCellNum; cellIndex++)
                {
                    var cell = row.GetCell(cellIndex);
                    if (cell == null) continue;
                    if (cell.CellType != CellType.String) continue;
                    if (cell.StringCellValue.StartsWith("$") == false) continue;

                    var label = cell.StringCellValue;
                    if (label.StartsWith("$:"))
                        Replaces.Add(new ReplaceLabel(cell, label));
                    else if (label.StartsWith("$$"))
                        Tables.Add(new TableLabel(cell, label));
                    if (label.StartsWith("$="))
                        Formulas.Add(new FormulaLabel(cell, label));
                }

            }
        }


        /// <summary>
        /// 构造公式，通过分析公式中的内容进行构造出Excel可以使用的公式
        /// </summary>
        /// <param name="formula"></param>
        /// <param name="getAddress"></param>
        private string ConstructFormula(string? formula, Func<Label, string?> getAddress)
        {
            if (string.IsNullOrWhiteSpace(formula)) return "";
            Regex regex = new Regex(@"(?<=[{]).*?(?=[}])");
            var variable = regex.Matches(formula);
            foreach (Match? item in variable)
            {
                Label varLabel = Tables.FirstOrDefault(x => x.ToString() == item?.Value);
                if (varLabel == null)
                    varLabel = Replaces.FirstOrDefault(x => x.ToString() == item?.Value);
                if (varLabel == null)
                    continue;

                formula = formula.Replace($"{{{item?.Value}}}", getAddress(varLabel));
                continue;
            }
            return formula;
        }

        /// <summary>
        /// 填写标签数据
        /// </summary>
        /// <param name="data"></param>
        public void WriteReplace(object data)
        {
            var properties = data.GetType().GetProperties();

            foreach (var label in Replaces)
            {
                var prop = properties.FirstOrDefault(x => x.Name == label.Name);
                if (prop == null)
                {
                    label.Cell.SetCellValue("");
                    continue;
                }
                object? value = prop.GetValue(data);
                label.Cell.SetCellValueObject(value);
            }
        }

        /// <summary>
        /// 填写表格数据
        /// </summary>
        /// <param name="tabelName"></param>
        /// <param name="data"></param>
        public void WriteTable(string tabelName, List<object> datas)
        {
            //找到所有表格标签
            var labels = Tables.Where(x => x.Table == tabelName).ToList();
            if (labels.Count == 0) return;

            //获得模板行
            var tempRow = Sheet.GetRow(labels.First().Cell.RowIndex);

            //根据表格行数量创建空表格
            for (int i = 0; i < datas.Count - 1; i++)
            {
                Sheet.CopyRow(tempRow.RowNum, tempRow.RowNum + 1);
            }

            //填充数据
            for (int dataRowIndex = 0; dataRowIndex < datas.Count; dataRowIndex++)
            {
                var properties = datas[dataRowIndex].GetType().GetProperties();

                foreach (var label in labels)
                {
                    if (string.IsNullOrWhiteSpace(label.Formula))
                    {//没有公式，那么只是普通的数据替换
                        var prop = properties.FirstOrDefault(x => x.Name == label.Name);
                        if (prop == null)
                        {
                            Sheet.GetCell(tempRow.RowNum + dataRowIndex, label.Cell.ColumnIndex)?.SetCellValue("");
                            continue;
                        }
                        object? value = prop.GetValue(datas[dataRowIndex]);
                        Sheet.GetCell(tempRow.RowNum + dataRowIndex, label.Cell.ColumnIndex)?.SetCellValueObject(value);
                    }
                    else
                    {//进入公式编辑方式
                        var formula = ConstructFormula(label.Formula, label => Sheet.GetCell(tempRow.RowNum + dataRowIndex, label.Cell.ColumnIndex)?.Address?.FormatAsString());
                        Sheet.GetCell(tempRow.RowNum + dataRowIndex, label.Cell.ColumnIndex)?.SetCellType(CellType.Formula);
                        Sheet.GetCell(tempRow.RowNum + dataRowIndex, label.Cell.ColumnIndex)?.SetCellFormula(formula);
                    }
                }
            }

            //计算每列的区域
            foreach (var label in labels)
            {
                label.CellRange = new CellRangeAddress(tempRow.RowNum, tempRow.RowNum + datas.Count - 1, label.Cell.ColumnIndex, label.Cell.ColumnIndex);
            }
        }

        /// <summary>
        /// 完成公式替换
        /// </summary>
        public void WriteFormula()
        {
            foreach (var label in Formulas)
            {
                var formula = ConstructFormula(label.Formula, label => label.AddressString());

                label.Cell.SetCellType(CellType.Formula);
                label.Cell.SetCellFormula(formula);
            }

        }
    }

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

        /// <summary>
        /// 公式
        /// </summary>
        public string? Formula { get; set; } = null;

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
    public class TableLabel : Label
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
    public class FormulaLabel : Label
    {
        public FormulaLabel(ICell cell, string label) : base(cell)
        {
            this.Formula = label.Substring(2);
        }

        public override string AddressString()
        {
            return Cell.Address.FormatAsString();
        }
    }
}
