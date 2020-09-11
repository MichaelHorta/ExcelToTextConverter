ExcelToTxtConverter (.NET)
===============================

[![Nuget badge](https://buildstats.info/nuget/ExcelToTxtConverter)](https://www.nuget.org/packages/ExcelToTxtConverter)

Nuget
-----
You can find nuget package with name ```ExcelToTxtConverter```

Introduction
-------
.NET package for converting EXCEL to TXT

Benefits and Features
-------
- Optionally, sorts by column
- Optionally, group by column
- Supporting for ```Binary Excel files (2.0-2003 format; *.xls)``` and ```OpenXml Excel files (2007 format; *.xlsx, *.xlsb)```
- Supports ```.NET Standard 2.0```, ```.NET Framework 4.6```

Example
-------
```C#
Assembly assembly = typeof(MyClass).GetTypeInfo().Assembly;
Stream definitionStream = assembly.GetManifestResourceStream(definitionFile));
XElement definitionElement = XElement.Load(definitionStream);
Converter converter = new Converter(definitionElement);
IDictionary<string, StringBuilder> stringBuildersResult = converter.Execute(excelBytes);
```

Example of Definition
-------
```XML
<?xml version="1.0" encoding="utf-8" ?>
<Definition>
    <Table>
      <Column ExcelID="Codigo Cuenta" TxtColumnText="&lt;CodigoCuenta&gt;" TxtTextPosition="0"></Column>
      <Column ExcelID="Apertura Debe" TxtColumnText="&lt;AperturaDebe&gt;" TxtTextPosition="25" CellFormat="3"></Column>
      <Column ExcelID="Apertura Haber" TxtColumnText="&lt;AperturaHaber&gt;" TxtTextPosition="48" CellFormat="3"></Column>
      <Column ExcelID="tipo comprobante" TxtColumnText="&lt;TpoComp&gt;" TxtTextPosition="71"></Column>
      <Column ExcelID="Numero comprobante" TxtColumnText="&lt;NumComp&gt;" TxtTextPosition="84"></Column>
      <Column ExcelID="Fecha" TxtColumnText="&lt;FechaContable&gt;" TxtTextPosition="106" CellFormat="0"></Column>
      <Column ExcelID="Glosa Analisis" TxtColumnText="&lt;GlosaAnalisis&gt;" TxtTextPosition="124"></Column>
      <Column ExcelID="Rut" TxtColumnText="&lt;Rut&gt;" TxtTextPosition="248"></Column>
      <Column ExcelID="Nombre" TxtColumnText="&lt;Nombre&gt;" TxtTextPosition="262"></Column>
      <Column ExcelID="Tipo Documento" TxtColumnText="&lt;TpoDocum&gt;" TxtTextPosition="386"></Column>
      <Column ExcelID="Numero documento" TxtColumnText="&lt;Numero&gt;" TxtTextPosition="405"></Column>
      <Column ExcelID="Fecha Emision" TxtColumnText="&lt;FchEmision&gt;" TxtTextPosition="427"></Column>
      <Column ExcelID="Fecha Vcto" TxtColumnText="&lt;FchVencimiento&gt;" TxtTextPosition="441"></Column>
      <Column ExcelID="Glosa" TxtColumnText="&lt;Glosa&gt;" TxtTextPosition="461"></Column>
      <Column ExcelID="Ref" TxtColumnText="&lt;Ref&gt;" TxtTextPosition="495"></Column>
      <Column ExcelID="Debe" TxtColumnText="&lt;Debe&gt;" TxtTextPosition="510" CellFormat="3"></Column>
      <Column ExcelID="Haber" TxtColumnText="&lt;Haber&gt;" TxtTextPosition="542" CellFormat="3"></Column>
    </Table>
</Definition>
```

Sort by Column
-------
Indicate in the XML definition the orderable column
```XML
<Column ExcelID="Codigo Cuenta" TxtColumnText="&lt;CodigoCuenta&gt;" TxtTextPosition="0" Orderable="int|string"></Column>
```

Group by Column
-------
Indicate in the XML definition the grouper column
```XML
<Column ExcelID="Fecha" TxtColumnText="&lt;FechaContable&gt;" TxtTextPosition="44" CellFormat="0" GroupKey="true"></Column>
```
In this point its neccesary indicates a function that builds the group identifier
```C#
public class MyTxtWriter
{
    public static string RetrieveBuilderKey(int indexRecord, IList<ExcelToTxtConverter.ColumnHeadData> lceColumnList, System.Data.DataTable dataTable)
    {
        var col = lceColumnList.Where(o => o.GroupKey.Equals(true)).FirstOrDefault();
        if (null == col)
        {
            return "{guid}";
        }

        string cellValue = string.Empty;
        DateTime dateValue;
        try
        {

            var cell = dataTable.Rows[indexRecord][col.ColumnPosition];
            cellValue = cell?.ToString();

            if (col.CellFormat.Equals(ExcelToTxtConverter.DateCellFormatter.Identifier))
            {
                dateValue = DateTime.Parse(cellValue);
                return string.Format("{0}{1}" + dateValue.Year, dateValue.ToString("MM"));
            }
        }
        catch (Exception)
        {
            Console.WriteLine(string.Format("Error parsing date: {0}", cellValue));
            throw;
        }

        return "{guid}";
    }

    public static string MakeBuilderKey(string builderKey)
    {
        builderKey = builderKey.Replace("{guid}", Guid.NewGuid().ToString());
        return string.Format("{0}.txt", builderKey);
    }
}
```

```C#
Assembly assembly = typeof(MyClass).GetTypeInfo().Assembly;
Stream definitionStream = assembly.GetManifestResourceStream(definitionFile));
XElement definitionElement = XElement.Load(definitionStream);

Func<int, IList<ColumnHeadData>, System.Data.DataTable, string> retrieveGroupKeyFunction = new Func<int, IList<ColumnHeadData>, System.Data.DataTable, string>(MyTxtWriter.RetrieveBuilderKey);

Converter converter = new Converter(definitionElement, retrieveGroupKeyFunction);
IDictionary<string, StringBuilder> stringBuildersResult = converter.Execute(excelBytes);
var resultDictionary = new Dictionary<string, string>();
foreach (var generatedTxt in generatedTxts)
{
    resultDictionary.Add(MyTxtWriter.MakeBuilderKey(generatedTxt.Key), generatedTxt.Value.ToString());
}
```