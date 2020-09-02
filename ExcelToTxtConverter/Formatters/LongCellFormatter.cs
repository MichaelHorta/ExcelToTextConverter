using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelToTxtConverter
{
    public class LongCellFormatter : ICellValueFormatter
    {
        public static string Identifier = "2";

        const NumberStyles numStyle = NumberStyles.AllowThousands;
        CultureInfo culture = new CultureInfo("en-US");
        public string ApplyFormatToValue(string cellValue)
        {
            long.TryParse(cellValue, numStyle, culture, out long retValue);
            return retValue.ToString();
        }
    }
}
