using System;
using System.Collections.Generic;
using System.Text;

namespace Report.NPOI
{
    public class ReportDataSource
    {
        public object Data { get; set; }

        public Dictionary<string, List<object>> Tables { get; set; }

    }


}
