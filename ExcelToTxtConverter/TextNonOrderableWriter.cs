using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;

namespace ExcelToTxtConverter
{
    public class TextNonOrderableWriter: TextOrderableBase
    {
        protected IDictionary<string, StringBuilder> buildersDictionary;

        public TextNonOrderableWriter() : base()
        {
            buildersDictionary = new Dictionary<string, StringBuilder>();
        }

        protected override void TryInitializeBuilder(string builderKey)
        {
            if (buildersDictionary.ContainsKey(builderKey))
                return;

            buildersDictionary.Add(builderKey, new StringBuilder());

            WriteHeadLine(builderKey);
        }

        private void WriteHeadLine(string builderKey)
        {
            buildersDictionary[builderKey].AppendLine(headLine);
        }

        public override void Execute(ExcelWorksheet excelWorksheet, IList<ColumnHeadData> columnList, XElement definition, Func<int, IList<ColumnHeadData>, ExcelWorksheet, string> grouperFunction)
        {
            InitializeExecution(excelWorksheet, columnList, definition, grouperFunction);

            for (int i = 0; i < columnList.Count; i++)
            {
                var col = columnList[i];
                ConcatenateHeadLine(col.TxtColumnText, columnList[i].TxtTextPosition);
            }

            int totalRows = excelWorksheet.Dimension.Rows;
            for (int i = 2; i <= totalRows; i++)
            {
                if (IsEmptyRow(i))
                    continue;
                string builderKey = RetrieveBuilderKey(i);

                for (int j = 0; j < columnList.Count - 1; j++)
                {
                    var col = columnList[j];

                    var cell = excelWorksheet.Cells[i, col.ColumnPosition];

                    ApplyFormatToCell(col, cell);

                    var cellValue = cell.Text?.ToString();

                    cellValue = ApplyFormatToValue(col, cellValue);

                    WriteRecord(new TextRecord()
                    {
                        Value = cellValue ?? string.Empty,
                        ColumnHeadData = col
                    }, builderKey, columnList[j + 1].TxtTextPosition - columnList[j].TxtTextPosition);
                }
                WriteRecord(new TextRecord()
                {
                    Value = excelWorksheet.Cells[i, columnList[columnList.Count - 1].ColumnPosition].Value?.ToString() ?? string.Empty,
                    ColumnHeadData = columnList[columnList.Count - 1]
                }, builderKey);

                if (i < totalRows)
                    AppendLine(builderKey);
            }
        }

        private void AppendLine(string builderKey)
        {
            buildersDictionary[builderKey].AppendLine();
        }

        private void WriteRecord(TextRecord record, string builderKey, int distance = -1)
        {
            if (distance != -1 && record.Value.Length > distance)
                record.Value = record.Value.Substring(0, distance);

            buildersDictionary[builderKey].Append(record.Value);
            if (distance != -1)
                buildersDictionary[builderKey].Append(' ', distance - record.Value.Length);
        }

        public override IDictionary<string, StringBuilder> GetBuilders()
        {
            return buildersDictionary;
        }
    }
}
