[![NuGet](https://img.shields.io/nuget/v/SheetHelper.svg)](https://www.nuget.org/packages/SheetHelper/)

<img src="https://raw.githubusercontent.com/stevenrskelton/flag-icon/master/png/16/country-4x3/br.png" width=2.0% height=2.0%> Veja a documentação em português [clicando aqui](SheetHelper/Globalization/Readme_pt-br.md).<br/>

# SheetHelper
Fast and lightweight library for easy read and conversion of large Excel files.<br/>

<img src="SheetHelper/Images/SheetHelper_publish.png" width=100% height=100%> 

AVAILABLE FEATURES:<br/>
✔ Compatible with reading .xlsx, .xlsm, .xls, .xlsb, .csv, .txt, .rpt files, among others;<br/>
✔ Get a datatable from a spreadsheet using the `GetDataTable` method;<br/>
✔ Use `SaveDataTable` to save dataTable in different formats and with restriction of columns and rows;<br/>
✔ Use the `CloseExcel` method to close all Excel processes, even those in the background;<br/>
✔ Use `GetIndexColumn` to get the column index by giving the name (e.g. "AB");<br/>
✔ The `GetNameColumn` method can be used to get the column name;<br/>
✔ Use `GetFirstRow` to get the first row of a DataTable;<br/>
✔ Convert a array to a DataRow using the `ConverToDataRow` method;<br/>
✔ Convert a spreadsheet to different formats using the 'Converter' method;;<br/>
✔ Allows to convert ranges of rows. Eg: "1:23, -34:56, 70, 75, -1";<br/>
✔ Possibility to convert ranges of columns. Eg: "A:H, 4:9, 4:-9, B, 75, -2";<br/>
✔ Replaces file if already converted;<br/>
✔ Option to choose the desired sheet for conversion using index or name (no sensitive case);<br/>
✔ Can choose the file format to be converted;<br/>
✔ Option to choose the file name, destination location and format to be saved;<br/>
✔ Allowed to change the delimiter;<br/>
✔ Supports conversion of hidden columns, rows, and sheets;<br/>
✔ Possibility to choose specific columns and rows to be converted;<br/>
✔ Allows you to track the percentage of completion through the `"Progress"` property;<br/>
✔ Has handling for end user when file not found with MessageBox for NetFramework;<br/>
✔ Supports unzipping of .GZ (using `UnGZ`) and .ZIP (using `UnZIP`) files. Or use `UnzipAuto` to unzip automatically.<br/>

<br/>Uses the library [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) to perform the reading.<br/>

<br/><br/>

## MAIN POSSIBLE CONVERSIONS:<br/>
<img src="SheetHelper/Images/Conversions.png" width=80% height=80%> 

### INSTALLATION:
```
 dotnet add package SheetHelper
```


## EXAMPLE OF USE:
```c#
using SH;

namespace App
{
    static class Program
    {
        static void Main()
        {
            string source  = "C:\\Users\\Diego\\Files\\Report.xlsx.gz";
            string destination = "C:\\Users\\Diego\\Files\\Report.csv";

            string sheet = "1"; // Use "1" for the first sheet (index or name)
            string delimiter  = ";";
            string columns  = "A, 3, b, 12:-1"; // or null to convert all columns or "A:BC" for a column range
            string rows = ":4, -2"; // Eg: Extracts from the 1nd to the 4nd row and also the penultimate row      

            bool success = SheetHelper.Converter(source, destination, sheet, delimiter, columns, rows);

            if (success ) Console.WriteLine("The file was converted successfully!");
            else Console.WriteLine("Failed to convert the file!");
        }
    }
}

```