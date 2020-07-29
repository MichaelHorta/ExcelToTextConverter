using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToTxtConverter
{
    public interface ICellValueFormatter
    {
        string ApplyFormatToValue(string cellValue);
    }
}
