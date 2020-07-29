using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToTxtConverter
{
    public interface ICellFormatter
    {
        void ApplyFormatToCell(ExcelRange excelRange);
    }
}
