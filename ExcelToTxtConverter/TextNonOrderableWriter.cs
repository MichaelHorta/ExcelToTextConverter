using System;
using System.Collections.Generic;
using System.Data;
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

        public override void Execute(DataTable dataTable, IList<ColumnHeadData> columnList, XElement definition, Func<int, IList<ColumnHeadData>, DataTable, string> grouperFunction, Func<int, IList<ColumnHeadData>, DataTable, bool> ignoreRowFunction)
        {
            InitializeExecution(dataTable, columnList, definition, grouperFunction);

            for (int i = 0; i < columnList.Count; i++)
            {
                var col = columnList[i];
                ConcatenateHeadLine(col.TxtColumnText, columnList[i].TxtTextPosition);
            }

            int rowsCount = dataTable.Rows.Count;
            for (int i = 1; i < rowsCount; i++)
            {
                if (IsEmptyRow(i) || ignoreRowFunction(i, columnList, dataTable))
                      continue;

                string builderKey = RetrieveBuilderKey(i);

                for (int j = 0; j < columnList.Count; j++)
                {
                    var col = columnList[j];
                    var cellValue = string.Empty;
                    if (col.ColumnPosition != -1)
                    {
                        var cell = dataTable.Rows[i][col.ColumnPosition];
                        cellValue = cell?.ToString();
                        cellValue = ApplyFormatToValue(col, cellValue);
                    }

                    WriteRecord(new TextRecord()
                    {
                        Value = cellValue ?? string.Empty,
                        ColumnHeadData = col
                    }, builderKey, j < columnList.Count - 1 ? columnList[j + 1].TxtTextPosition - columnList[j].TxtTextPosition : -1);
                }

                if (i < rowsCount)
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
