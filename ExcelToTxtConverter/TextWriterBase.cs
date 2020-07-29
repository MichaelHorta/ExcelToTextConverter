using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace ExcelToTxtConverter
{
    public class TextWriterBase
    {

        private IDictionary<CellFormat, ICellValueFormatter> cellFormattersToValueDictionary;
        private IDictionary<CellFormat, ICellFormatter> cellFormattersDictionary;

        public ExcelWorksheet Worksheet { get; set; }
        private XElement Definition { get; set; }
        public IList<ColumnHeadData> ColumnList { get; private set; }

        public TextWriterBase(ExcelWorksheet worksheet, XElement definition)
        {
            cellFormattersToValueDictionary = new Dictionary<CellFormat, ICellValueFormatter>
            {
                { CellFormat.Date, new DateCellFormatter() },
                { CellFormat.Integer, new IntegerCellFormatter() },
                { CellFormat.Long, new LongCellFormatter() },
                { CellFormat.Decimal, new DecimalCellFormatter() }
            };

            cellFormattersDictionary = new Dictionary<CellFormat, ICellFormatter>
            {
                { CellFormat.Decimal, new DecimalCellFormatter() }
            };

            Worksheet = worksheet;
            Definition = definition;

            InitializeColumnList();
        }

        private void InitializeColumnList()
        {
            int colndex = 1;
            ColumnList = new List<ColumnHeadData>();
            int totalCols = Worksheet.Dimension.Columns;

            var rootElement = Definition.Element("Table");
            if (null == rootElement)
                throw new Exception("Definition doesn't contains the Table element");

            var columnDefinitionElements = rootElement.Elements();

            #region Exploring head - Row 1
            while (colndex <= totalCols)
            {
                var cellValue = Worksheet.Cells[1, colndex].Value?.ToString().Trim().ToLower();
                if (string.IsNullOrEmpty(cellValue))
                {
                    colndex++;
                    continue;
                }

                var column = columnDefinitionElements.Where(lcd => lcd.Attribute("ExcelID").Value.ToLower().Equals(cellValue)).FirstOrDefault();
                if (null != column)
                {
                    var columnHeadData = new ColumnHeadData
                    {
                        ExcelID = column.Attribute("ExcelID").Value,
                        TxtColumnText = column.Attribute("TxtColumnText").Value,
                        TxtTextPosition = int.Parse(column.Attribute("TxtTextPosition").Value),
                        ColumnPosition = colndex
                    };
                    ColumnList.Add(columnHeadData);

                    var cellFormatAttribute = column.Attribute("CellFormat");
                    if (null != cellFormatAttribute && !string.IsNullOrEmpty(cellFormatAttribute.Value))
                    {
                        columnHeadData.CellFormat = (CellFormat)Enum.Parse(typeof(CellFormat), cellFormatAttribute.Value);
                    }

                    IdentifyOrderableColumn(column, columnHeadData);
                }
                colndex++;
            }
            ColumnList = ColumnList.OrderBy(c => c.TxtTextPosition).ToList();
            #endregion
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
        public void Execute(Func<int, IList<ColumnHeadData>, ExcelWorksheet, string> grouperFunction)
        {
            textOrderableWriter = TextOrderableWriterFactory.Build(null != orderableColumnHeadData);
            textOrderableWriter.Execute(Worksheet, ColumnList, Definition, grouperFunction);
        }

        public IDictionary<string, StringBuilder> GetBuilders()
        {
            return textOrderableWriter.GetBuilders();
        }
    }
}
