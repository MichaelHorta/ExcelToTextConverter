using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace ExcelToTxtConverter
{
    public class TextWriterBase
    {
        public DataTable DataTable { get; set; }
        private XElement Definition { get; set; }
        public IList<ColumnHeadData> ColumnList { get; private set; }

        public TextWriterBase(DataTable dataTable, XElement definition)
        {
            DataTable = dataTable;
            Definition = definition;

            InitializeColumnList();
        }

        private void InitializeColumnList()
        {
            ColumnList = new List<ColumnHeadData>();

            var rootElement = Definition.Element("Table");
            if (null == rootElement)
                throw new Exception("Definition doesn't contains the Table element");

            var firstRow = DataTable.Rows[0];

            var columnDefinitionElements = rootElement.Elements();
            var enumerator = columnDefinitionElements.GetEnumerator();
            while(enumerator.MoveNext())
            {
                var element = enumerator.Current;
                var columnHeadData = new ColumnHeadData
                {
                    ExcelID = element.Attribute("ExcelID").Value,
                    TxtColumnText = element.Attribute("TxtColumnText").Value,
                    TxtTextPosition = int.Parse(element.Attribute("TxtTextPosition").Value),
                    ColumnPosition = Array.IndexOf(firstRow.ItemArray, element.Attribute("ExcelID").Value.ToString())
                };
                
                bool groupKey = false;
                if(null != element.Attribute("GroupKey"))
                {
                    bool.TryParse(element.Attribute("GroupKey").Value, out groupKey);
                }
                columnHeadData.GroupKey = groupKey;
                ColumnList.Add(columnHeadData);

                var cellFormatAttribute = element.Attribute("CellFormat");
                if (null != cellFormatAttribute && !string.IsNullOrEmpty(cellFormatAttribute.Value))
                {
                    columnHeadData.CellFormat = (CellFormat)Enum.Parse(typeof(CellFormat), cellFormatAttribute.Value);
                }

                IdentifyOrderableColumn(element, columnHeadData);
            }
            ColumnList = ColumnList.OrderBy(c => c.TxtTextPosition).ToList();
        }

        private ColumnHeadData orderableColumnHeadData;
        private void IdentifyOrderableColumn(XElement column, ColumnHeadData columnHeadData)
        {
            if (null != orderableColumnHeadData)
                return;

            var orderableAttribute = column.Attribute("Orderable");
            if (null != orderableAttribute && !string.IsNullOrEmpty(orderableAttribute.Value))
            {
                columnHeadData.Orderable = new OrderableAttribute
                {
                    Type = orderableAttribute.Value
                };
                orderableColumnHeadData = columnHeadData;
            }
        }

        TextOrderableBase textOrderableWriter;
        public void Execute(Func<int, IList<ColumnHeadData>, DataTable, string> grouperFunction, Func<int, IList<ColumnHeadData>, DataTable, bool> ignoreRowFunction)
        {

            if (null == grouperFunction)
                grouperFunction = new Func<int, IList<ColumnHeadData>, DataTable, string>(DefaultGrouperFunction);

            if (null == ignoreRowFunction)
            {
                ignoreRowFunction = new Func<int, IList<ColumnHeadData>, DataTable, bool>(DefaultIgnoreRowFunction);
            }

            textOrderableWriter = TextOrderableWriterFactory.Build(null != orderableColumnHeadData);
            textOrderableWriter.Execute(DataTable, ColumnList, Definition, grouperFunction, ignoreRowFunction);
        }

        static string defaultSeparatorGuid = null;
        private static string DefaultGrouperFunction(int indexRecord, IList<ColumnHeadData> columnList, DataTable dataTable)
        {
            if (string.IsNullOrEmpty(defaultSeparatorGuid))
            {
                defaultSeparatorGuid = Guid.NewGuid().ToString();
            }
            return defaultSeparatorGuid;
        }

        public static bool DefaultIgnoreRowFunction(int rowIndex, IList<ColumnHeadData> columnList, DataTable dataTable)
        {
            return false;
        }

        public IDictionary<string, StringBuilder> GetBuilders()
        {
            return textOrderableWriter.GetBuilders();
        }
    }
}
