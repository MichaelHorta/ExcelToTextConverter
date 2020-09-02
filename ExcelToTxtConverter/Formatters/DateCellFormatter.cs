using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelToTxtConverter
{
    public class DateCellFormatter : ICellValueFormatter
    {
        public static string Identifier = "0";
        
        public string ApplyFormatToValue(string cellValue)
        {
            if (string.IsNullOrEmpty(cellValue))
                return cellValue;

            DateTime.TryParse(cellValue, new CultureInfo("en-US"), DateTimeStyles.None, out DateTime result);
            if (null != result)
                return result.ToString("yyyy-MM-dd");

            return cellValue;
        }
    }
}
