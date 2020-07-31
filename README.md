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
- Only support for XLSX file
- Supports .NET Standard 1.3, .NET 4

Example
```C#
Assembly assembly = typeof(MyClass).GetTypeInfo().Assembly;
Stream definitionStream = assembly.GetManifestResourceStream(definitionFile));
XElement definitionElement = XElement.Load(definitionStream);
Converter converter = new Converter(definitionElement);
IDictionary<string, StringBuilder> stringBuildersResult = converter.Execute(excelBytes);
```