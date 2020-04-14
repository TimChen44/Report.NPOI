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
        //替换标签和表格标签集合
        public List<Label> Labels { get; set; } = new List<Label>();

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
                        Labels.Add(new ReplaceLabel(cell, label));
                    else if (label.StartsWith("$$"))
                        Labels.Add(new TableLabel(cell, label));
                    if (label.StartsWith("$="))
                        Labels.Add(new FormulaLabel(cell, label));
                }

            }
        }

        /// <summary>
        /// 填充数据
        /// </summary>
        /// <param name="reportDataSource"></param>
        public void Fill(ReportDataSource reportDataSource)
        {
            //进行标签替换
            WriteReplace(reportDataSource.Data);

            //进行表格填充
            foreach (var table in reportDataSource.Tables)
            {
                WriteTable(table.Key, table.Value);
            }

            WriteFormula();
        }


        /// <summary>
        /// 构造公式，并且填充公式
        /// </summary>
        /// <param name="formula"></param>
        /// <param name="getAddress"></param>
        private void FillFormula(ICell currentCell, string formula, Func<Label, string?> getAddress)
        {
            Regex regex = new Regex(@"(?<=[{]).*?(?=[}])");
            var variable = regex.Matches(formula);
            foreach (Match? item in variable)
            {
                Label varLabel = Labels.FirstOrDefault(x => x.ToString() == item?.Value);

                formula = formula.Replace($"{{{item?.Value}}}", getAddress(varLabel));
                continue;
            }
            currentCell?.SetCellType(CellType.Formula);
            currentCell?.SetCellFormula(formula);
        }

        /// <summary>
        /// 填充值
        /// </summary>
        private void FillValue(ICell currentCell, string labelName, object data, PropertyInfo[] properties)
        {
            var prop = properties.FirstOrDefault(x => x.Name == labelName);
            if (prop == null)
            {
                currentCell.SetCellValue("");
                return;
            }
            object? value = prop.GetValue(data);
            currentCell.SetCellValueObject(value);
        }


        /// <summary>
        /// 填写标签数据
        /// </summary>
        /// <param name="data"></param>
        public void WriteReplace(object data)
        {
            var properties = data.GetType().GetProperties();

            foreach (ReplaceLabel label in Labels.Where(x => x is ReplaceLabel))
            {
                FillValue(label.Cell, label.Name, data, properties);
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
            var labels = Labels.Where(x => x is TableLabel label && label.Table == tabelName).Cast<TableLabel>().ToList();
            if (labels.Count == 0) return;

            //获得模板行
            //var tempRow = Sheet.GetRow(labels.First().Cell.RowIndex);
            var tempRow = labels.First().Cell.Row;

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
                    var currentCell = Sheet.GetCell(tempRow.RowNum + dataRowIndex, label.Cell.ColumnIndex);
                    if (currentCell == null) continue;

                    if (string.IsNullOrWhiteSpace(label.Formula))
                    {//没有公式，那么只是普通的数据替换
                        FillValue(currentCell, label.Name, datas[dataRowIndex], properties);
                    }
                    else
                    {//进入公式编辑方式
                        FillFormula(currentCell, label.Formula, label => Sheet.GetCell(tempRow.RowNum + dataRowIndex, label.Cell.ColumnIndex)?.Address?.FormatAsString());
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
            foreach (FormulaLabel label in Labels.Where(x => x is FormulaLabel))
            {
                FillFormula(label.Cell, label.Formula, label => label.AddressString());

            }

        }
    }

}
