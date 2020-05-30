# How to Write Excel in .Net?

It is extremely simple to write in excel file using IronXL library .It provides the most simplest and easy way to communicate with excel files for .Net developers. Just access the specific cell, where you want to write the value, assign your custom value without any bulk lines of code. 

## Write Value In Specific Cell:  
First of all add reference of `IronXL` in your project and import liberary by `Using IronXL`. In the followig case our excel file name is `sample.xlsx` and it exists in `bin>Debug` folder of the project
```CSharp
using IronXL;

WorkBook workbook = WorkBook.Load("sample.xlsx"); //load excel file 
WorkSheet sheet = workbook.GetWorkSheet("Sheet1"); //Get sheet1 of sample.xlsx
sheet["A1"].Value = "new value"; //access A1 cell and assign value
sheet.SaveAs("sample.xlsx");   //save changes         
```
## Write Static Value In Many Cells:
it is very easy to write static value in many cells at a time by using (colon `:`).its left side indicate starting cell and right side for last cell of specific column to be changed.
>`sheet[From:To]`

As show in code below:
```CSharp
using IronXL;

WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
sheet["A1:A9"].Value = "new value";
sheet.SaveAs("sample.xlsx");            
```
it will write value `new value` from cell `A1` to `A9` of `column A`.

## Write Value In Many Cells By User Inputs
by the folloing way we can take value from user and write it in the excel file
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
above code diapay following output and take input from user:

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

