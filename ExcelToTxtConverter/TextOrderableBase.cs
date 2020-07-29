using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;

namespace ExcelToTxtConverter
{
    public abstract class TextOrderableBase
    {
        protected ExcelWorksheet excelWorksheet;
        protected IList<ColumnHeadData> columnList;
        protected XElement definition;
        private IDictionary<CellFormat, ICellValueFormatter> cellFormattersToValueDictionary;
        private IDictionary<CellFormat, ICellFormatter> cellFormattersDictionary;
        protected Func<int, IList<ColumnHeadData>, ExcelWorksheet, string> grouperFunction;
        protected string headLine = string.Empty;

        public TextOrderableBase()
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
        }

        public abstract void Execute(ExcelWorksheet excelWorksheet, IList<ColumnHeadData> lceColumnList, XElement definition, Func<int, IList<ColumnHeadData>, ExcelWorksheet, string> grouperFunction);

        protected void InitializeExecution(ExcelWorksheet excelWorksheet, IList<ColumnHeadData> columnList, XElement definition, Func<int, IList<ColumnHeadData>, ExcelWorksheet, string> grouperFunction)
        {
            this.excelWorksheet = excelWorksheet;
            this.columnList = columnList;
            this.definition = definition;
            this.grouperFunction = grouperFunction;
        }

        protected string RetrieveBuilderKey(int indexRecord)
        {
            string builderKey = grouperFunction(indexRecord, columnList, excelWorksheet);
            TryInitializeBuilder(builderKey);

            return builderKey;
        }

        protected abstract void TryInitializeBuilder(string builderKey);

        protected void ConcatenateHeadLine(string value, int witdthAtRight = 0)
        {
            headLine = headLine.PadRight(witdthAtRight);
            headLine += value;
        }

        protected void ApplyFormatToCell(ColumnHeadData columnHeadData, ExcelRange excelRange)
        {
            var cellFormat = columnHeadData.CellFormat;
            if (!cellFormat.HasValue)
                return;

            cellFormattersDictionary.TryGetValue(cellFormat.Value, out ICellFormatter cellFormatter);
            if (null == cellFormatter)
                return;

            cellFormatter.ApplyFormatToCell(excelRange);
        }

        protected string ApplyFormatToValue(ColumnHeadData columnHeadData, string cellValue)
        {
            var cellFormat = columnHeadData.CellFormat;
            if (!cellFormat.HasValue)
                return cellValue;

            cellFormattersToValueDictionary.TryGetValue(cellFormat.Value, out ICellValueFormatter cellFormatter);
            if (null == cellFormatter)
                return cellValue;

            return cellFormatter.ApplyFormatToValue(cellValue);
        }

        protected bool IsEmptyRow(int rowIndex)
        {
            for (int j = 0; j < columnList.Count; j++)
            {
                var col = columnList[j];
                var cellValue = excelWorksheet.Cells[rowIndex, col.ColumnPosition].Text?.Trim();
                if (!string.IsNullOrEmpty(cellValue))
                    return false;
            }
            return true;
        }

        public abstract IDictionary<string, StringBuilder> GetBuilders();
    }
}
