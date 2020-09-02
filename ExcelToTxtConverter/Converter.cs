using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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

        public IDictionary<string, StringBuilder> Execute(byte[] excelData, Func<int, IList<ColumnHeadData>, DataTable, string> grouperFunction = null, Func<int, IList<ColumnHeadData>, DataTable, bool> ignoreRowFunction = null)
        {
            try
            {   
                var tempFilenamePath = Path.GetTempFileName();
                File.WriteAllBytes(tempFilenamePath, excelData);
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = File.Open(tempFilenamePath, FileMode.Open, FileAccess.Read))
                {
                    // Auto-detect format, supports:
                    //  - Binary Excel files (2.0-2003 format; *.xls)
                    //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        DataTable dataTable = result.Tables[0];

                        var lceTxtWriter = new TextWriterBase(dataTable, Definition);
                        lceTxtWriter.Execute(grouperFunction, ignoreRowFunction);
                        ColumnList = lceTxtWriter.ColumnList;

                        reader.Close();

                        return lceTxtWriter.GetBuilders();
                    }
                }
                
            }
            catch (Exception ex)
            {
                throw new Exception("Error executing conversion", ex);
            }
        }
    }
}
