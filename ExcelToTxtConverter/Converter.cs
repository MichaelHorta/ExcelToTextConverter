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
                var dataReader = new DataReader();
                var dataTable = dataReader.Execute(excelData);

                var lceTxtWriter = new TextWriterBase(dataTable, Definition);
                lceTxtWriter.Execute(grouperFunction, ignoreRowFunction);
                ColumnList = lceTxtWriter.ColumnList;

                return lceTxtWriter.GetBuilders();
            }
            catch (Exception ex)
            {
                throw new Exception("Error executing conversion", ex);
            }
        }
    }
}
