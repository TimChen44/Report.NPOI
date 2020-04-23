# Report.NPOI
使用Excel模板，将文件中的标签做替换，实现各类报表输出

```
Install-Package Report.NPOI
```

# Demo

### 模板

![image](https://github.com/TimChen44/Report.NPOI/raw/master/doc/a.png)
 
### 输出

![image](https://github.com/TimChen44/Report.NPOI/raw/master/doc/b.png)

### 代码

``` csharp

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

```
