Microsoft Excel supports wide range of _Number Formats_. One of these number formats is _Currency_. If the cell value contains number, then you can apply currency format by specifying currency symbol and decimal places. Similarly, you can format Excel cell numbers to currency using [Aspose.Cells for Java](https://products.aspose.com/cells/java) with ease. In order to perform its operation, Aspose.Cells does not need Microsoft Excel or on any sort of Microsoft Office Automation, VBA (_Visual Basic for Applications_), VSTO (_Visual Studio Tools for Office_) etc.

**Article Description**
>The purpose of this article is to explain how developers can format Excel cell numbers to currency in Java.

**Supported Platforms**
>[Aspose.Cells](https://products.aspose.com/cells/) API supports number of platform e.g. Java, .NET, C++, Android, JavaScript, PHP etc. Besides, [Aspose.Cells is available in Cloud as REST or RESTful APIs](https://products.aspose.cloud/cells).

**Licensing**
>Aspose.Cells is paid or commercial api, so it is not free or open source. Without license, it will work in evaluation mode with some limitations. If you want to test Aspose.Cells without evaluation version limitations, you can also request a _30 Day Temporary License_. For more information, please go through [Licensing](https://docs.aspose.com/display/cellsjava/Licensing).

# Format Excel Cell Numbers to Currency using Microsoft Excel

You can format Excel cell numbers to currency using Microsoft Excel by performing these steps.

* _Right Click_ the cell that contains some numeric value.
* Click _Format Cells…_ from the context menu as shown in snapshot below.
* Select _Currency_ from _Number Category_ and press _OK_.

![Format Excel Cell Numbers to Currency using Microsoft Excel.](https://raw.githubusercontent.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/master/Format-Excel-Cell-Numbers-to-Currency/Format-Excel-Cell-Numbers-to-Currency-Microsoft-Excel.png "Format Excel Cell Numbers to Currency using Microsoft Excel.")

# Currency Custom Number Format Strings of Cell

You can use various types of currency custom number format strings to display currencies e.g. Dollar, Yuan, Pound, Euro etc. and many others.

## Dollar
```
"$"#,##0.00
```
## Yaun
```
[$¥-804]#,##0.00
```

## Pound
```
[$£-809]#,##0.00
```

## Euro
```
#,##0.00[$€-40B]
```

# Format Excel Cell Numbers to Currency using Aspose.Cells

In the next few sections, we will learn how to use Aspose.Cells API to format Excel cell numbers to currency.

# Sample Input Microsoft Excel Document

For demonstration, we will use the following [sample input Microsoft Excel document](https://github.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/blob/master/Format-Excel-Cell-Numbers-to-Currency/SampleFormatExcelCellNumbersToCurrency.xlsx) that contains some numbers in cells G3, G4, G5 and G6. We will apply currency format i.e. Dollar, Yuan, Pound, Euro on these cells respectively.

![Currency formats will be applied on Input Excel file Cells using Aspose.Cells API.](https://raw.githubusercontent.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/master/Format-Excel-Cell-Numbers-to-Currency/Input-Excel-File-Format-Excel-Cell-Numbers-To-Currency.png "Currency formats will be applied on Input Excel file Cells using Aspose.Cells API.")

# Sample Code

The following sample code formats Excel cells i.e. G3, G4, G5 and G6 with currency formats i.e. Dollar, Yuan, Pound, Euro respectively by performing these steps.

* Load the [input Excel file](https://github.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/blob/master/Format-Excel-Cell-Numbers-to-Currency/SampleFormatExcelCellNumbersToCurrency.xlsx) inside the _com.aspose.cells.Workbook_ object and access the first worksheet.
* Access first cell i.e. G3 and apply currency format using the _Style.setCustom()_ method.
* Repeat the second step for cell G4, G5 and G6 with further currency formats.
* Save the _com.aspose.cells.Workbook_ object in XLSX format. You can also save it to XLS or other Excel formats as per your needs.

>**Gist** - [Format Excel Cell Numbers to Currency - Java](https://gist.github.com/AsposeCells/0a3a94799f272c0e882ecec444e5988e)

```js
// Directory path for input and output Excel files.
String dirPath = "D:/Download/";

// Load the input Excel file inside workbook object.
com.aspose.cells.Workbook wb = new Workbook(dirPath + "SampleFormatExcelCellNumbersToCurrency.xlsx");
			
// Access first worksheet.
Worksheet ws = wb.getWorksheets().get(0);

// Format Cell G3 with Curreny > Dollar.
Cell cell = ws.getCells().get("G3");
Style st = cell.getStyle();
st.setCustom("\"$\"#,##0.00");
cell.setStyle(st);

// Format Cell G4 with Curreny > Yaun.
cell = ws.getCells().get("G4");
st = cell.getStyle();
st.setCustom("[$¥-804]#,##0.00");
cell.setStyle(st);

// Format Cell G5 with Curreny > Pound.
cell = ws.getCells().get("G5");
st = cell.getStyle();
st.setCustom("[$£-809]#,##0.00");
cell.setStyle(st);

// Format Cell G6 with Curreny > Euro.
cell = ws.getCells().get("G6");
st = cell.getStyle();
st.setCustom("#,##0.00[$€-40B]");
cell.setStyle(st);

// Save the workbook in XLSX format. 
// You can also save it to XLS or other formats.
wb.save(dirPath + "OutputFormatExcelCellNumbersToCurrency.xlsx");
```

# Output Microsoft Excel XLSX by Aspose.Cells after applying Currency formats

The following snapshot shows the [Output Microsoft Excel XLSX by Aspose.Cells](https://github.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/blob/master/Format-Excel-Cell-Numbers-to-Currency/OutputFormatExcelCellNumbersToCurrency.xlsx) after applying currency formats on the cells G3, G4, G5 and G6. Similarly, you can apply all sorts of currency formats on Excel cells using Aspose.Cells API.

![Currency formats applied on Excel cells using Aspose.Cells API.](https://raw.githubusercontent.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/master/Format-Excel-Cell-Numbers-to-Currency/Currency-Format-Applied-To-Excel-Cells-Aspose.Cells-API.png "Currency formats applied on Excel cells using Aspose.Cells API.")




