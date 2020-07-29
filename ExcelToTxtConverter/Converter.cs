using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace ExcelToTxtConverter
{
    public class Converter
    {
        private XElement Definition { get; set; }
        public IList<ColumnHeadData> ColumnList { get; internal set; }

        public Converter(XElement definition)
        {
            Definition = definition;
        }

        public IDictionary<string, StringBuilder> Execute(byte[] excelData, Func<int, IList<ColumnHeadData>, ExcelWorksheet, string> grouperFunction = null)
        {
            try
            {
                var lceExcelDataStream = new MemoryStream(excelData);
                var package = new ExcelPackage();
                package.Load(lceExcelDataStream);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                if (null == grouperFunction)
                    grouperFunction = new Func<int, IList<ColumnHeadData>, ExcelWorksheet, string>(DefaultGrouperFunction);

                var lceTxtWriter = new TextWriterBase(worksheet, Definition);
                lceTxtWriter.Execute(grouperFunction);
                ColumnList = lceTxtWriter.ColumnList;

                return lceTxtWriter.GetBuilders();
            }
            catch (Exception ex)
            {
                throw new Exception("Error executing conversion", ex);
            }
        }

        static string defaultSeparatorGuid = null;
        private static string DefaultGrouperFunction(int indexRecord, IList<ColumnHeadData> lceColumnList, ExcelWorksheet excelWorksheet)
        {
            if (string.IsNullOrEmpty(defaultSeparatorGuid))
            {
                defaultSeparatorGuid = Guid.NewGuid().ToString();
            }
            return defaultSeparatorGuid;
        }
    }
}
