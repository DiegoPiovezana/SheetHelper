## [Version 1.3.0] - 2023-12-18

- Possible to obtain the DataSet from a file
- Cells that have line breaks will now be correctly converted
- The `NormalizeText` method allows removing all accents and spaces from a text
- Possible to convert all tabs of the spreadsheet considering all rows and columns
- Use `GetRowArray` to obtain a row from a DataTable
- The `GetSheets` method allows obtaining all sheets from the spreadsheet in DataTable
- Fixed selection of first line that is not the header
- Now it's possible to perform the conversion of tabs that have only one row

## [Version 1.2.0] - 2023-07-10

- Allowed to specify the name of the sheet disregarding sensitive case
- Allowed to save dataTable in different formats and with restriction of columns and rows
- Added possibility to get the first row of a DataTable
- Added handling for end user when file not found or in use, with MessageBox for NetFramework
- Fixed conversion considering header format


## [Version 1.1.1] - 2023-06-01

- Dealing with unnecessary conversion between .CSV and .TXT


## [Version 1.1.0] - 2023-05-24

- Added possibility to convert ranges of rows. Eg: "1:23, -34:56, 70, 75, -1"
- Possible to convert column ranges. Eg: "A:H, 4:9, 4:-9, B, 75, -2"
- Integrated ExcelDataReader library
- Compatible with net462, netstandard 2.0 and netstandard 2.1
- Several bugs fixed
- Performance improvement


## [Version 1.0.0.6] - 2022-06-13

- Added option to perform continuous column (A:AB) conversion and choose sheet by name or index


## [Version 1.0.0.5] - 2022-03-29

- Added option to convert RPT files


## [Version 1.0.0.4] - 2022-03-28

- Added option to choose rows, progressBar and unzip files


## [Version 1.0.0.3] - 2022-02-01

- Adding XSLB, XSL, HTML and CSV file conversion possibility and header size choice


## [Version 1.0.0.2] - 2022-01-19

- Allowing visualization of the summary of the methods


## [Version 1.0.0.1] - 2022-01-18

- Icon, readme and naming update
- Bugs fixed


## [Version 1.0.0] - 2022-01-17

- Release