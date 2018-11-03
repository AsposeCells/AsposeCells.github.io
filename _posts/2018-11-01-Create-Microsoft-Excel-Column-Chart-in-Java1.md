Chart is a graphical representation of data that enables user to easily understand the large quantity of data and relationships between parts of data. There are many types of charts available in Microsoft Excel and almost all of them are supported 
by [Aspose.Cells for Java](https://products.aspose.com/cells/java). As a matter of fact, Aspose.Cells can be used to create, edit and manipulate Excel spreadsheets almost in any platform without any need to install Microsoft Excel or without using any sort of Microsoft Office automation.

**Article Description**
>The purpose of this article is to explain how developers can use AutoFilter to filter Excel data in Java.

**Supported Platforms**
>[Aspose.Cells](https://products.aspose.com/cells) API supports various platforms including Java, .NET, C++, Android, JavaScript, PHP etc. Besides, [Aspose.Cells is available in Cloud as RESTful APIs](https://products.aspose.cloud/cells).

# Types of Column Charts

There are various types of Column charts, some of them are listed below.

* Column
* Column Stacked
* Column 100% Stacked
* Column 3D
* Column 3D Clustered
* Column 3D Stacked
* Column 3D 100% Stacked

# Sample Input Microsoft Excel Document

For demonstration, we will use the following [sample input Microsoft Excel document](https://github.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/blob/master/Create-Microsoft-Excel-Column-Chart/sampleCreateMicrosoftExcelColumnChart.xlsx) that contains the chart data. Here, column A contains the category axis data and other columns B, C and D contain chart series data.

![Sample Input Microsoft Excel Document containing Chart Data.](https://raw.githubusercontent.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/master/Create-Microsoft-Excel-Column-Chart/Input-Microsoft-Excel-Document-containing-Chart-Data.png "Sample Input Microsoft Excel Document containing Chart Data.")

# Sample Code

The following sample code creates Microsoft Excel Column Chart by performing these steps.

* Load [sample input Microsoft Excel document](https://github.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/blob/master/Create-Microsoft-Excel-Column-Chart/sampleCreateMicrosoftExcelColumnChart.xlsx) containing the chart data.
* Create Column chart with specified dimensions.
* Set the chart title and format it.
* Add three vertical series, set their names and fill colors.
* Format various chart items e.g. plot area, value axis, category axis, major tick marks etc.
* Save the workbook in XLSX format. You can also save it in other formats e.g. XLS, XLSB, XLSM etc.

>**Gist** - [Create Microsoft Excel Column Chart in Java](https://gist.github.com/AsposeCells/e1029cb702a873dd0c6d434e32240e07)

```js
// Directory path of input and output files.
String dirPath = "D:/Download/";

// Load source Excel file containing the chart data.
Workbook wb = new Workbook(dirPath + "sampleCreateMicrosoftExcelColumnChart.xlsx");

// Access first worksheet.
Worksheet ws = wb.getWorksheets().get(0);

// Specify dimensions of the chart.
int upperLeftRow = 7;
int upperLeftColumn = 4;
int lowerRightRow = 24;
int lowerRightColumn = 13;

// Create Column chart with specified dimensions.
int idx = ws.getCharts().add(ChartType.COLUMN, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);

// Access the Column chart.
Chart ch = ws.getCharts().get(idx);

// Set the outline of chart area.
ch.getChartArea().getBorder().setColor(Color.getBlack());
ch.getChartArea().getBorder().setWeight(WeightType.SINGLE_LINE);

// Set the chart title, make it non-bold and set its font size.
ch.getTitle().setText("Classification of Languages");
ch.getTitle().getFont().setBold(false);
ch.getTitle().getFont().setSize(15);

// Add three vertical series in chart covering the range B2:D5.
ch.getNSeries().add("B2:D5", true);

// Set the category data covering the range A2:A5.
ch.getNSeries().setCategoryData("A2:A5");

// Set the names of the chart series taken from cells.
ch.getNSeries().get(0).setName("=B1");
ch.getNSeries().get(1).setName("=C1");
ch.getNSeries().get(2).setName("=D1");

// Set the 1st series fill color.
ch.getNSeries().get(0).getArea().setForegroundColor(Color.fromArgb(74, 127, 176));
ch.getNSeries().get(0).getArea().setFormatting(FormattingType.CUSTOM);

// Set the 2nd series fill color.
ch.getNSeries().get(1).getArea().setForegroundColor(Color.fromArgb(91, 155, 213));
ch.getNSeries().get(1).getArea().setFormatting(FormattingType.CUSTOM);

// Set the 3rd series fill color.
ch.getNSeries().get(2).getArea().setForegroundColor(Color.fromArgb(173, 198, 229));
ch.getNSeries().get(2).getArea().setFormatting(FormattingType.CUSTOM);

// Set plot area formatting as none and hide its border.
ch.getPlotArea().getArea().getFillFormat().setFillType(FillType.NONE);
ch.getPlotArea().getBorder().setVisible(false);

// Set value axis major tick mark as none and hide axis line. 
// Also set the color of value axis major grid lines.
ch.getValueAxis().setMajorTickMark(TickMarkType.NONE);
ch.getValueAxis().getAxisLine().setVisible(false);
ch.getValueAxis().getMajorGridLines().setColor(Color.fromArgb(217, 217, 217));

// Save the output Excel file in XLSX format.
wb.save(dirPath + "outputCreateMicrosoftExcelColumnChart.xlsx", SaveFormat.XLSX);
```

# Output Microsoft Excel Column Chart by Aspose.Cells

The following snapshot shows the [Output Microsoft Excel Column Chart generated by Aspose.Cells](https://github.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/blob/master/Create-Microsoft-Excel-Column-Chart/outputCreateMicrosoftExcelColumnChart.xlsx) with the code given above. Similarly, you can create all sorts of Column charts with Aspose.Cells API easily.

![Microsoft Excel Column Chart created by Aspose.Cells API.](https://raw.githubusercontent.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/master/Create-Microsoft-Excel-Column-Chart/Microsoft-Excel-Column-Chart-created-by-Aspose.Cells-API.png "Microsoft Excel Column Chart created by Aspose.Cells API.")
