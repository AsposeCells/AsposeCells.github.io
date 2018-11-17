XLS is an old but widely used Microsoft Excel format. It is a binary file format and known as _Binary Interchange File Format (BIFF)_. Microsoft Excel still supports XLS format for backward compatibility but since it is an old format, many new features are not supported inside it. For that reason, XLS format is often converted to XLSX format which is a newer Open Office XML-based format and supports all new features of Microsoft Excel.

You can convert XLS to XLSX manually using Microsoft Excel 2007 or later versions e.g. Microsoft Excel 2016 etc. or programmatically using [Aspose.Cells for .NET](https://products.aspose.com/cells/net) API with few lines of code. In order to perform its operations, Aspose.Cells does not depend on Microsoft Excel or on any sort of Microsoft Office Automation, VBA (_Visual Basic for Applications_), VSTO (_Visual Studio Tools for Office_) etc.

**Article Description**

>The purpose of this article is to explain how developers can convert Microsoft Excel XLS to XLSX format in C# or in any other .NET Framework supported language e.g. VB.NET etc.

**Supported Platforms**

>[Aspose.Cells](https://products.aspose.com/cells/) API supports all .NET frameworks e.g. .NET 2.0, .NET 3.5, .NET 4.0, .NET 7.0, .NET Core, .NET Standard 2.0, Xamarin etc. It is also available in other platforms e.g. Java, C++, Android, JavaScript, PHP etc. Besides, [Aspose.Cells is available in Cloud as REST or RESTful APIs](https://products.aspose.cloud/cells).

# Maximum Number of Columns and Rows in XLS and XLSX

XLS format supports

* 65,536 Rows.
* 256 Columns.

XLSX format supports

* 1,048,576 Rows.
* 16,384 Columns.

# Sample Input Microsoft Excel XLS Document

You can convert any XLS document to XLSX using Aspose.Cells API. For illustration, we will use the following [sample input Microsoft Excel XLS document](https://github.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/blob/master/Convert-Microsoft-Excel-XLS-to-XLSX/SampleConvertMicrosoftExcelXLSToXLSX.xls) that contains textual and numerical formatted data about some companies. Whenever, you will open XLS document in Microsoft Excel, it will show _Compatibility Mode_ as indicated by _Red Arrow_ in the following snapshot.

![Sample Microsoft Excel XLS document to be converted to XLSX format using Aspose.Cells API.](https://raw.githubusercontent.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/master/Convert-Microsoft-Excel-XLS-to-XLSX/Input-Convert-XLS-to-XLSX-using-Aspose.Cells-API.png "Sample Microsoft Excel XLS document to be converted to XLSX format using Aspose.Cells API.")

# Sample Code

The following sample code converts XLS to XLSX by performing these steps.

* Load the input Microsoft Excel XLS document in _Aspose.Cells.Workbook_ object.
* Save the _Aspose.Cells.Workbook_ object in XLSX format.

>**Gist** - [Convert Microsoft Excel Xls To Xlsx In C#](https://gist.github.com/AsposeCells/46733b053f952d37b90d25498084c4d7)

```js
// Directory path of input and output files.
string dirPath = "D:/Download/";

// Specify load options Excel97To2003 i.e. XLS format. 
LoadOptions opts = new LoadOptions(LoadFormat.Excel97To2003);
            
// Load the input XLS file inside the Aspose.Cells workbook object.
Aspose.Cells.Workbook wb = new Workbook(dirPath + "SampleConvertMicrosoftExcelXLSToXLSX.xls", opts);

// Save the workbook as output XLSX file.
wb.Save(dirPath + "OutputConvertMicrosoftExcelXLSToXLSX.xlsx", SaveFormat.Xlsx);
```

# Output Microsoft Excel XLSX by Aspose.Cells

The following snapshot shows the [Converted or Output Microsoft Excel XLSX file by Aspose.Cells](https://github.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/blob/master/Convert-Microsoft-Excel-XLS-to-XLSX/OutputConvertMicrosoftExcelXLSToXLSX.xlsx) with the code given above. As you can see, the output XLSX file is exactly similar to XLS file. Similarly, you can convert any XLS file to XLSX with Aspose.Cells API easily.

![Output - Convert XLS to XLSX using Aspose.Cells API.](https://raw.githubusercontent.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/master/Convert-Microsoft-Excel-XLS-to-XLSX/Output-Convert-XLS-to-XLSX-using-Aspose.Cells-API.png "Output - Convert XLS to XLSX using Aspose.Cells API.")



