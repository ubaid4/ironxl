# What is IronXL?
[IronXL](https://ironsoftware.com/csharp/excel/) is an Excel library for .NET developers which provides the easiest way to communicate with Excel(XLS, XLSX, CSV and TSV ) files, without any dependencie, even without using `Microsoft.Office.Interop.Excel` library and installation of  `Microsoft Office` on target machine.You can get rid of a lot of complicated lines of code and can use `IronXL` to get the easiest way. It provides all type of functions which can be required by developer e.g:
* Create new Excel file, insert data programmatically using Excel functions and set the style (font,color,bold,italic and other cell properties) as well.
* Import Excel file in the project ,use its data effectively and manipulate it programmatically.
* Behave with Excel file as Dataset and Datatables.

**`IronXL` supports the following:**
* Net Framework 4.5+ (C#, VB.Net,ASP.Net WebForms and MVC)
* Net Core 2+
* Net Standard
* Xamarin
* Windows Application(Desktop applications)
* Windows Mobile
* Mono
* Azure Cloud hosting
  
**Supported Operating System(OS):**
* Windows
* MacOS
* Linux
* iOS
* Andriod

## `IronXL` Installation:
There are two following ways to install `IronXL` 

### 1. Using NuGet Package:
Using NuGet Package Manager in Visual Studio project, you can browse the `IronXL.Excel` and and install it.
> PM > Install-Package IronXL.Excel

`IronXL` classes can be access using `IronXL` namespace.
### 2. By Downloading IronXL.dll
[Download IronXL.dll](https://ironsoftware.com/csharp/excel/) and add its reference in your project. `IronXL` classes can be access using `IronXL` namespace.
 


# How to Write Excel in .Net?

It is extremely simple to write in Excel file using IronXL library . Just access the specific cell, where you want to write the value, assign your custom value without any bulk lines of code. 

## Write Value In Specific Cell:  
First of all add reference of `IronXL` in your project and import liberary by `Using IronXL`. In the followig case our Excel file name is `sample.xlsx` and it exists in `bin>Debug` folder of the project.
```CSharp 
using IronXL;

WorkBook workbook = WorkBook.Load("sample.xlsx"); //load excel file 
WorkSheet sheet = workbook.GetWorkSheet("Sheet1"); //Get sheet1 of sample.xlsx
sheet["A1"].Value = "new value"; //access A1 cell and assign value
sheet.SaveAs("sample.xlsx");   //save changes         
```
## Write Static Value In Many Cells:
It is very easy to write static value in many cells at a time by using (colon `:`).Its left side indicate starting cell and right side for last cell of specific column to be changed.
>`sheet[From:To]`

As show in the code below:
```CSharp
using IronXL;

WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
sheet["A1:A9"].Value = "new value";
sheet.SaveAs("sample.xlsx");            
```
It will write value `new value` from cell `A1` to `A9` of `column A`.

## Write Value In Many Cells By User Inputs
By the folloing way we can take value from user and write it in the excel file:
```c#
using IronXL;

string _from, _to,NewValue ;

Console.Write("Enter Starting Cell :");
_from = Console.ReadLine();

Console.Write("Enter Last Cell :");
_to = Console.ReadLine();

Console.Write("Enter value:");
NewValue = Console.ReadLine();

WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
sheet[_from + ":" + _to].Value = NewValue;
Console.WriteLine("Successfully Changed...!");
Console.ReadKey();
```
Above code diapay the following output and take inputs from user:

![output](https://github.com/ubaid4/ironxl/blob/master/user_input_write_excel.png)

**values changed from B4 to B9 in Excel sheet,We can see.**

![output](https://github.com/ubaid4/ironxl/blob/master/excl_result.png)







## Write Dynamic Values In Many Cells:
Dynamic Values can be set in many cells in excel file by the following way:
```CSharp
using IronXL;
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
for (int i = From; i < To; i++)
{
    sheet["A" + i].Value = "Value"+i;
}
sheet.SaveAs("sample.xlsx");  
```

