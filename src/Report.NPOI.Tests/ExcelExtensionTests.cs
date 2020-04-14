using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace Report.NPOI.Extension.Tests
{
    [TestClass()]
    public class ExcelExtensionTests
    {
        [TestMethod()]
        public void WriteReportTest()
        {

            XSSFWorkbook workbook = new XSSFWorkbook("Template.xlsx");

            ReportDataSource rds = new ReportDataSource();
            rds.Data = new
            {
                Company = "Cmp",
                Phone = "137",
            };
            rds.Tables = new Dictionary<string, List<object>>
            {
                {
                    "Table",
                    new List<object>()
                    {
                        new {Title="A" , Price=8, Count=2 },
                        new {Title="V" ,Price=6.5, Count=5 }
                    }
                }
            };

            workbook.GetSheetAt(0).WriteReport(rds);

            var ms = new MemoryStream();
            workbook.Write(ms);


            workbook.Close();
        }

    }
}