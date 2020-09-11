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
In this point its neccesary indicates a function that builds the group identifier
```C#
public static string RetrieveGroupKey(int indexRecord, IList<ExcelToTxtConverter.ColumnHeadData> columnList, ExcelWorksheet excelWorksheet)
{
    string cellValue = string.Empty;
    string cellFormat = string.Empty;
    DateTime? dateValue = null;
    try
    {
        var col = columnList.Where(o => o.ExcelID.Equals("Fecha")).FirstOrDefault();
        var cell = excelWorksheet.Cells[indexRecord, col.ColumnPosition];
        cellValue = cell.Text?.ToString();
        cellFormat = cell.Style.Numberformat.Format;

        var filteredFormat = "dd\\-mm\\-yyyy";
        if (string.Equals(cellFormat, filteredFormat, StringComparison.InvariantCultureIgnoreCase))
        {
            cell.Style.Numberformat.Format = "mm/dd/yyyy";
            cellValue = cell.Text?.ToString();
        }

        dateValue = DateTime.Parse(cellValue);
    }
    catch (Exception)
    {
        Console.WriteLine(string.Format("Error parsing date: {0} with format: {1}", cellValue, cellFormat));
        throw;
    }

    return string.Format("{0}{1}" + dateValue.Value.Year, dateValue.Value.ToString("MM"));
}
```

```C#
Assembly assembly = typeof(MyClass).GetTypeInfo().Assembly;
Stream definitionStream = assembly.GetManifestResourceStream(definitionFile));
XElement definitionElement = XElement.Load(definitionStream);

Func<int, IList<ColumnHeadData>, ExcelWorksheet, string> retrieveGroupKeyFunction = new Func<int, IList<ColumnHeadData>, ExcelWorksheet, string>(RetrieveGroupKey);

Converter converter = new Converter(definitionElement, retrieveGroupKeyFunction);
IDictionary<string, StringBuilder> stringBuildersResult = converter.Execute(excelBytes);
```