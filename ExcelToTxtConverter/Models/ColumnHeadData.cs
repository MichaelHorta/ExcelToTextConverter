using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToTxtConverter
{
    public class ColumnHeadData
    {
        public string ExcelID { get; set; }
        public string TxtColumnText { get; set; }
        public int TxtTextPosition { get; set; }
        public int ColumnPosition { get; set; }
        public string CellFormat { get; set; }
        public OrderableAttribute Orderable { get; set; }
        public bool GroupKey { get; set; }
        public IDictionary<string, string> CustomAttributes { get; set; }
    }

    public class OrderableAttribute
    {
        public string Type { get; set; }
    }
}
