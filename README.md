# How to Write Excel in .Net?

It is extremely simple to write in excel file using IronXL library .It provides the most simplest and easy way to communicate with excel files for .Net developers. Just access the specific cell, where you want to write the value, assign your custom value without any bulk lines of code. 

## Write Value In Specific Cell:  
```CSharp
using IronXL;

WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
sheet["A1"].Value = "new value";
sheet.SaveAs("sample.xlsx");            
```
## Write Static Value In Many Cells:
it is very easy to write static value in many cells at a time by using (colon `:`).its left side indicate starting cell and right side for last cell of specific column 
>`From:To `

As show in code below:
```CSharp
using IronXL;

WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
sheet["A1:A9"].Value = "new value";
sheet.SaveAs("sample.xlsx");            
```
it will write value `new value` from cell `A1` to `A9` of `column A`.
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
![ookk](https://www.w3schools.com/images/picture.jpg)

