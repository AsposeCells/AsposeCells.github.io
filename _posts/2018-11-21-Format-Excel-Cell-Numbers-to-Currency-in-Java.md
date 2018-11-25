Filtering data based on criteria is a very important feature. It helps the user to understand and analyze data easily. You can use the auto filter feature of Microsoft Excel to find, show or hide values in a single or multiple columns based on the choices that you select from list. When the filter is applied, then all the rows that do not meet your criteria are hidden completely.

You can use [Aspose.Cells for C++](https://products.aspose.com/cells/cpp) to apply filter on your Excel data programmatically in C# easily with few lines of code. It can also be used to perform wide range of functions on Excel documents e.g. you can create, edit and manipulate Excel spreadsheets in any platform without any need to install Microsoft Excel or without using any sort of Microsoft Office automation.

**Article Description**

>The purpose of this article is to explain how developers can use AutoFilter to filter Excel data in C++.

**Supported Platforms**

>[Aspose.Cells](https://products.aspose.com/cells/) API supports number of platform e.g. C++, .NET, Java, Android, JavaScript, PHP etc. Besides, [Aspose.Cells is available in Cloud as RESTful APIs](https://products.aspose.cloud/cells).

# Filter Data using AutoFilter in Microsoft Excel

Please do the following steps to filter data using AutoFilter in Microsoft Excel.

* Select the columns and click _Data > Filter_ button inside the _Sort & Filter_ section.
* Click the AutoFilter dropdown, select your choices from list and press OK.
* All the rows that do not match your criteria will be filtered out. Please see this snapshot for detail.

![Apply AutoFilter in Microsoft Excel which can also be done with Aspose.Cells API programmatically.](https://raw.githubusercontent.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/master/Use-AutoFilter-to-Filter-Excel-Data/Apply-AutoFilter-Microsoft-Excel-Aspose.Cells-API.png "Apply AutoFilter in Microsoft Excel which can also be done with Aspose.Cells API programmatically.")

# Filter Data using AutoFilter in Aspose.Cells

This section explains how to filter Excel data using AutoFilter with Aspose.Cells API.

# Sample Input Microsoft Excel Document

For demonstration, we will use the following [sample input Microsoft Excel document](https://github.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/blob/master/Use-AutoFilter-to-Filter-Excel-Data/sampleUseAutoFilterToFilterExcelData.xlsx) that contains some data in four columns. We will apply AutoFilter on _Vehicle_ and _Color_ columns. Once, rows are filtered out, some of them will become hidden and the **Grand Total** for _Qty1_ and _Qty2_ columns shown inside the red lines will be modified accordingly.

![Sample Input Microsoft Excel Document containing Data for applying AutoFilter using Aspose.Cells API.](https://raw.githubusercontent.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/master/Use%20AutoFilter%20to%20Filter%20Excel%20Data/Sample-Microsoft-Excel-Apply-AutoFilter-Aspose.Cells-API.png "Sample Input Microsoft Excel Document containing Data for applying AutoFilter using Aspose.Cells API.")

# Sample Code

The following sample code applies AutoFilter on Microsoft Excel data by performing these steps

* Load [sample input Microsoft Excel document](https://github.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/blob/master/Use-AutoFilter-to-Filter-Excel-Data/sampleUseAutoFilterToFilterExcelData.xlsx) containing the sample data for auto filter.
* Apply auto filter to range.
* Adds two filters to first column.
* Refresh the auto filter.
* Adds another two filters to second column.
* Refresh the auto filter.
* Save the workbook in XLSX format. You can also save it in other formats e.g. XLS, XLSB, XLSM etc.

> **Gist** - [Use AutoFilter to Filter Excel Data - C++](https://gist.github.com/AsposeCells/0919095bebf74907d0971077d454ae98)

```js
// Path of input Excel file.
intrusive_ptr<Aspose::Cells::System::String> inputExcelFile = new Aspose::Cells::System::String("D:/Download/sampleUseAutoFilterToFilterExcelData.xlsx");

// Path of output Excel file.
intrusive_ptr<Aspose::Cells::System::String> outputExcelFile = new Aspose::Cells::System::String("D:/Download/outputUseAutoFilterToFilterExcelData.xlsx");

// Declaration of some variables to be used later.
intrusive_ptr<Aspose::Cells::System::String> strRng;
intrusive_ptr<Aspose::Cells::System::String> strCriteria;

// Load the input Excel file containing the sample data.
intrusive_ptr<Aspose::Cells::IWorkbook> wb = Aspose::Cells::Factory::CreateIWorkbook(inputExcelFile);

// Access first worksheet.
intrusive_ptr<Aspose::Cells::IWorksheet> ws = wb->GetIWorksheets()->GetObjectByIndex(0);

// Apply auto filter to the range.
strRng = new Aspose::Cells::System::String("D3:G3");
ws->GetIAutoFilter()->SetRange(strRng);

// Add filter to first column (i.e. Vehicle) inside the range - Criteria --> Bike
strCriteria = new Aspose::Cells::System::String("Bike");
ws->GetIAutoFilter()->AddFilter(0, strCriteria);

// Add filter to first column (i.e. Vehicle) inside the range - Criteria --> Car
strCriteria = new Aspose::Cells::System::String("Car");
ws->GetIAutoFilter()->AddFilter(0, strCriteria);

// Refresh the auto filter.
ws->GetIAutoFilter()->Refresh();

// Add filter to second column (i.e. Color) inside the range - Criteria --> Green
strCriteria = new Aspose::Cells::System::String("Green");
ws->GetIAutoFilter()->AddFilter(1, strCriteria);

// Add filter to second column (i.e. Color) inside the range - Criteria --> Blue
strCriteria = new Aspose::Cells::System::String("Blue");
ws->GetIAutoFilter()->AddFilter(1, strCriteria);

// Refresh the auto filter.
ws->GetIAutoFilter()->Refresh();

// Save the workbook in XLSX format. 
// You can also save it to XLS or other formats.
wb->Save(outputExcelFile, Aspose::Cells::SaveFormat::SaveFormat_Xlsx);
```

# Output Microsoft Excel Document by Aspose.Cells after applying AutoFilter

The following snapshot shows the [Output Microsoft Excel Document generated by Aspose.Cells](https://github.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/blob/master/Use-AutoFilter-to-Filter-Excel-Data/outputUseAutoFilterToFilterExcelData.xlsx) after applying AutoFilter with the code given above. _As you can see, it now shows the filtered rows and new values of **Grand Total** displayed inside the red lines._

![Output Microsoft Excel Document by Aspose.Cells API after applying AutoFilter.](https://raw.githubusercontent.com/AsposeCells/AsposeCells-Screenshots-and-Sample-Files/master/Use-AutoFilter-to-Filter-Excel-Data/Output-Microsoft-Excel-Apply-AutoFilter-Aspose.Cells-API.png "Output Microsoft Excel Document by Aspose.Cells API after applying AutoFilter.")

