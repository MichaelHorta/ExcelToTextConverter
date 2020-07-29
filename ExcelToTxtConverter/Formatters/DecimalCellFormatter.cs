using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelToTxtConverter
{
    public class DecimalCellFormatter : ICellValueFormatter, ICellFormatter
    {
        const NumberStyles numStyle = NumberStyles.AllowThousands;
        CultureInfo culture = new CultureInfo("en-US");

        public void ApplyFormatToCell(ExcelRange excelRange)
        {
            excelRange.Style.Numberformat.Format = "0.0000000";
        }

        public string ApplyFormatToValue(string cellValue)
        {
            decimal.TryParse(cellValue, out decimal retValue);

            return string.Format("{0:G29}", retValue);
        }
    }
}
