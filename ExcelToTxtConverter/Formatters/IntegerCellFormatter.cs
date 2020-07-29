using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelToTxtConverter
{
    public class IntegerCellFormatter : ICellValueFormatter
    {
        const NumberStyles numStyle = NumberStyles.AllowThousands;
        CultureInfo culture = new CultureInfo("en-US");
        public string ApplyFormatToValue(string cellValue)
        {
            Int32.TryParse(cellValue, numStyle, culture, out int retValue);
            return retValue.ToString();
        }
    }
}
