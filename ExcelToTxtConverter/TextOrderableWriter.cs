using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;
using OfficeOpenXml;
using System.Linq;

namespace ExcelToTxtConverter
{
    public class TextOrderableWriter : TextOrderableBase
    {
        private IDictionary<string, OrderableGroup> buildersDictionary;
        private ColumnHeadData orderableColumnHeadData;

        public TextOrderableWriter() : base()
        {
            buildersDictionary = new Dictionary<string, OrderableGroup>();
        }

        protected override void TryInitializeBuilder(string builderKey)
        {
            if (buildersDictionary.ContainsKey(builderKey))
                return;

            buildersDictionary.Add(builderKey, new OrderableGroup());

            WriteHeadLine(builderKey);
        }

        private void WriteHeadLine(string builderKey)
        {
            buildersDictionary[builderKey].CommonBuilder.AppendLine(headLine);
        }

        public override void Execute(ExcelWorksheet excelWorksheet, IList<ColumnHeadData> columnList, XElement definition, Func<int, IList<ColumnHeadData>, ExcelWorksheet, string> grouperFunction)
        {
            InitializeExecution(excelWorksheet, columnList, definition, grouperFunction);

            IdentifyOrderableColumn();

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
                string line = string.Empty;
                string orderableKeyGroupValue = string.Empty;
                for (int j = 0; j < columnList.Count - 1; j++)
                {
                    var col = columnList[j];

                    var cell = excelWorksheet.Cells[i, col.ColumnPosition];

                    ApplyFormatToCell(col, cell);

                    var cellValue = cell.Text?.ToString();

                    cellValue = ApplyFormatToValue(col, cellValue);

                    if (col.Equals(orderableColumnHeadData))
                    {
                        orderableKeyGroupValue = cellValue;
                    }

                    line = WriteRecordInLine(line, new TextRecord()
                    {
                        Value = cellValue ?? string.Empty,
                        ColumnHeadData = col
                    }, columnList[j + 1].TxtTextPosition - columnList[j].TxtTextPosition);
                }
                line = WriteRecordInLine(line, new TextRecord()
                {
                    Value = excelWorksheet.Cells[i, columnList[columnList.Count - 1].ColumnPosition].Value?.ToString() ?? string.Empty,
                    ColumnHeadData = columnList[columnList.Count - 1]
                });

                WriteLine(line, builderKey, orderableKeyGroupValue);
            }
        }

        private void IdentifyOrderableColumn()
        {
            orderableColumnHeadData = columnList.FirstOrDefault(c => null != c.Orderable);
        }

        private string WriteRecordInLine(string line, TextRecord record, int distance = -1)
        {
            if (distance != -1 && record.Value.Length > distance)
                record.Value = record.Value.Substring(0, distance);

            line += record.Value;
            if (distance != -1)
                line += new String(' ', distance - record.Value.Length);

            return line;
        }

        private void WriteLine(string line, string builderKey, string orderableKeyGroupValue)
        {
            var orderableGroup = buildersDictionary[builderKey];
            orderableGroup.BuildersRecordsTable.TryGetValue(orderableKeyGroupValue, out StringBuilder stringBuilder);

            var newStringBuilder = new StringBuilder(line);

            if (null == stringBuilder)
                orderableGroup.BuildersRecordsTable.Add(orderableKeyGroupValue, newStringBuilder);
            else
                orderableGroup.BuildersRecordsTable[orderableKeyGroupValue] = stringBuilder.Append(newStringBuilder);

            orderableGroup.BuildersRecordsTable[orderableKeyGroupValue].AppendLine();
        }

        public override IDictionary<string, StringBuilder> GetBuilders()
        {
            var buildersResult = new Dictionary<string, StringBuilder>();

            var enumerator = buildersDictionary.GetEnumerator();
            while (enumerator.MoveNext())
            {
                var orderableGroup = enumerator.Current;
                var value = orderableGroup.Value;
                var builderRecords = value.BuildersRecordsTable;

                List<KeyValuePair<string, StringBuilder>> builderRecordsSorted = new List<KeyValuePair<string, StringBuilder>>();
                switch (orderableColumnHeadData.Orderable.Type)
                {
                    case "int":
                        builderRecordsSorted = builderRecords.OrderBy(o =>
                        {
                            int parsedValue;
                            bool success = int.TryParse(o.Key, out parsedValue);
                            if (success)
                                return parsedValue;
                            else
                                return int.MaxValue;
                        }).ToList();
                        break;
                    default:
                        builderRecordsSorted = builderRecords.OrderBy(o => o.Key).ToList();
                        break;
                }

                var finalStringBuilder = value.CommonBuilder;
                foreach (var builderRecord in builderRecordsSorted)
                {
                    finalStringBuilder = finalStringBuilder.Append(builderRecord.Value);
                }

                buildersResult.Add(orderableGroup.Key, finalStringBuilder);
            }

            return buildersResult;
        }
    }

    public class OrderableGroup
    {
        public StringBuilder CommonBuilder { get; set; }
        public IDictionary<string, StringBuilder> BuildersRecordsTable { get; set; }

        public OrderableGroup()
        {
            CommonBuilder = new StringBuilder();
            BuildersRecordsTable = new Dictionary<string, StringBuilder>();
        }
    }
}
