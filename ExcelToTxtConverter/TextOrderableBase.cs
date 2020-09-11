using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Xml.Linq;

namespace ExcelToTxtConverter
{
    public abstract class TextOrderableBase
    {
        protected DataTable dataTable;
        protected IList<ColumnHeadData> columnList;
        protected XElement definition;
        protected readonly IDictionary<string,ICellValueFormatter> cellFormattersToValueDictionary;
        protected Func<int, IList<ColumnHeadData>, DataTable, string> grouperFunction;
        protected string headLine = string.Empty;

        public TextOrderableBase()
        {
            cellFormattersToValueDictionary = new Dictionary<string, ICellValueFormatter>
            {
                { DateCellFormatter.Identifier, new DateCellFormatter() },
                { IntegerCellFormatter.Identifier, new IntegerCellFormatter() },
                { LongCellFormatter.Identifier, new LongCellFormatter() },
                { DecimalCellFormatter.Identifier, new DecimalCellFormatter() }
            };
        }

        public abstract void Execute(DataTable dataTable, IList<ColumnHeadData> lceColumnList, XElement definition, Func<int, IList<ColumnHeadData>, DataTable, string> grouperFunction, Func<int, IList<ColumnHeadData>, DataTable, bool> ignoreRowFunction);

        protected void InitializeExecution(DataTable dataTable, IList<ColumnHeadData> columnList, XElement definition, Func<int, IList<ColumnHeadData>, DataTable, string> grouperFunction)
        {
            this.dataTable = dataTable;
            this.columnList = columnList;
            this.definition = definition;
            this.grouperFunction = grouperFunction;
        }

        protected string RetrieveBuilderKey(int indexRecord)
        {
            string builderKey = grouperFunction(indexRecord, columnList, dataTable);
            TryInitializeBuilder(builderKey);

            return builderKey;
        }

        protected abstract void TryInitializeBuilder(string builderKey);

        protected void ConcatenateHeadLine(string value, int witdthAtRight = 0)
        {
            headLine = headLine.PadRight(witdthAtRight);
            headLine += value;
        }

        protected string ApplyFormatToValue(ColumnHeadData columnHeadData, string cellValue)
        {
            var cellFormat = columnHeadData.CellFormat;
            if (string.IsNullOrEmpty(cellFormat))
                return cellValue;

            cellFormattersToValueDictionary.TryGetValue(cellFormat, out ICellValueFormatter cellFormatter);
            if (null == cellFormatter)
                return cellValue;

            return cellFormatter.ApplyFormatToValue(cellValue);
        }

        protected bool IsEmptyRow(int rowIndex)
        {
            for (int j = 0; j < columnList.Count; j++)
            {
                var col = columnList[j];
                if (col.ColumnPosition == -1)
                {
                    continue;
                }
                var cellValue = dataTable.Rows[rowIndex][col.ColumnPosition]?.ToString().Trim();
                if (!string.IsNullOrEmpty(cellValue))
                    return false;
            }
            return true;
        }

        public abstract IDictionary<string, StringBuilder> GetBuilders();
    }
}
