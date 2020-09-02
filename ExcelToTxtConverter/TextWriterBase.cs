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
        protected XElement Definition { get; set; }
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
                    columnHeadData.CellFormat = cellFormatAttribute.Value;
                }

                columnHeadData.CustomAttributes = new Dictionary<string, string>();
                var customAttributes = element.Attributes().Where(attr => !attr.Name.LocalName.Equals("ExcelID") && !attr.Name.LocalName.Equals("TxtColumnText") && !attr.Name.LocalName.Equals("TxtTextPosition") && !attr.Name.LocalName.Equals("GroupKey") && !attr.Name.LocalName.Equals("CellFormat"));
                foreach (var customAttr in customAttributes)
                {
                    columnHeadData.CustomAttributes.Add(customAttr.Name.LocalName, customAttr.Value.ToString());
                }

                IdentifyOrderableColumn(element, columnHeadData);
            }
            ColumnList = ColumnList.OrderBy(c => c.TxtTextPosition).ToList();
        }

        protected ColumnHeadData orderableColumnHeadData;
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

        protected TextOrderableBase textOrderableWriter;
        public virtual void Execute(Func<int, IList<ColumnHeadData>, DataTable, string> grouperFunction, Func<int, IList<ColumnHeadData>, DataTable, bool> ignoreRowFunction)
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
        protected static string DefaultGrouperFunction(int indexRecord, IList<ColumnHeadData> columnList, DataTable dataTable)
        {
            if (string.IsNullOrEmpty(defaultSeparatorGuid))
            {
                defaultSeparatorGuid = Guid.NewGuid().ToString();
            }
            return defaultSeparatorGuid;
        }

        protected static bool DefaultIgnoreRowFunction(int rowIndex, IList<ColumnHeadData> columnList, DataTable dataTable)
        {
            return false;
        }

        public IDictionary<string, StringBuilder> GetBuilders()
        {
            return textOrderableWriter.GetBuilders();
        }
    }
}
