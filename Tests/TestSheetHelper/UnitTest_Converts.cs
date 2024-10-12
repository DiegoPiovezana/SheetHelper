using SH;

namespace TestSheetHelper
{
    // https://github.com/DiegoPiovezana/SheetHelper

    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void TestPass()
        {
            Assert.Pass();
        }

        // --------------------------------------------------------------------------------
        [Test, Repeat(1)]
        public void TestManipulacaoDtGZ()
        {
            const string origin = @"C:\Users\diego\Desktop\Tests\Converter\ColunasExcel.xlsx.gz";
            const string destination = @"C:\Users\diego\Desktop\Tests\Convertidos\Especial_xlsx.csv";
            const string sheet = "1";
            const string delimiter = ";";
            const string columns = "A, 3, b, 12:-1";
            const string rows = ":4, -2";

            var sh = new SheetHelper();
            var dt = sh.GetDataTable(origin, sheet);

            // Validating that the table was loaded correctly
            Assert.IsNotNull(dt, "The DataTable was not loaded correctly.");
            Assert.IsTrue(dt.Rows.Count > 0, "No rows were loaded from the spreadsheet.");

            // Extracting the first row and validating values
            var first = sh.GetRowArray(dt);
            Assert.That(first, Is.Not.Null, "Extracting the first line failed.");
            Assert.AreEqual("100", first[99], "The value at position [99] is not '100' as expected.");

            bool success = sh.SaveDataTable(dt, destination, delimiter, columns, rows);
            Assert.IsTrue(success, "Failed to save DataTable in CSV format.");
        }

        [Test]
        public void TestDefaultReadme()
        {
            //string source = "C:\\Users\\Diego\\Files\\Report.xlsx.gz";
            //string destination = "C:\\Users\\Diego\\Files\\Report.csv";

            string source = @"C:\Users\diego\Desktop\Tests\Converter\ColunasExcel.zip";
            string destination = "C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\Report.csv";

            string sheet = "1"; // Use "1" for the first sheet (index or name)
            string delimiter = ";";
            string columns = "A, 3, b, -5:-1"; // or null to convert all columns or "A:BC" for a column range
            string rows = "1:2,:4, -2"; // Eg: Extracts from the 1nd to the 4nd row and also the penultimate row      

            var sheetHelper = new SheetHelper();
            bool success = sheetHelper.Converter(source, destination, sheet, delimiter, columns, rows);

            Assert.That(success, Is.True);
        }

        [Test]
        public void TestDefaultReadmeFull()
        {
            // Define paths for source and destination files
            string source = @"C:\Users\diego\Desktop\Tests\Converter\ColunasExcel.zip";
            string destination = @"C:\Users\diego\Desktop\Tests\Convertidos\Report.csv";

            // Specify parameters for the conversion
            string sheet = "1"; // Use "1" for the first sheet (index or name)
            string delimiter = ";";
            string columns = "A, 3, b, -5:-1"; // or null to convert all columns or "A:BC" for a column range
            string rows = "1:2,:4, -1"; // Extracts from the 1st to the 4th row and also the penultimate row      

            var sheetHelper = new SheetHelper();

            // Perform the conversion
            bool success = sheetHelper.Converter(source, destination, sheet, delimiter, columns, rows);

            // Validate that the conversion was successful
            Assert.That(success, Is.True, "Conversion process failed.");

            // Check if the output file exists
            Assert.That(File.Exists(destination), Is.True, $"The file '{destination}' was not created.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destination);
            Assert.That(csvLines.Length, Is.GreaterThan(0), "The converted CSV file is empty.");

            // Validate the number of rows matches the expected rows
            int expectedRows = 7; // Based on rows "1:2, :4, -2" = 1, 2, 1, 2, 3, 4, -1
            Assert.That(csvLines.Length, Is.EqualTo(expectedRows), $"Expected {expectedRows} rows but got {csvLines.Length}.");

            // Check if the delimiter is correct in the first line of the CSV
            Assert.That(csvLines[0], Does.Contain(delimiter), $"Expected delimiter '{delimiter}' not found in the output.");

            // 7 rows = 1, 2, 1, 2, 3, 4, -1
            // 8 columns = A, 3, b, -5, -4, -3, -2, -1  
            Assert.Multiple(() =>
            {
                // Check content of the first row -- 1
                Assert.That(csvLines[0].Split(delimiter)[0], Is.EqualTo("1"), $"Expected '1' not found in the output.");

                // Check content of the penultimate row -- 4
                Assert.That(csvLines[5].Split(delimiter)[2], Is.EqualTo("B4"), $"Expected 'B4' not found in the output.");

                // Check content of the last row and last column -- CV101254               
                Assert.That(csvLines[6].Split(delimiter)[7], Is.EqualTo("CV101254"), $"Expected 'CV101254' not found in the output.");
            });
        }

        [Test, Repeat(1)]
        public void TestUTF8()
        {
            string source = @"C:\Users\diego\Desktop\Tests\Converter\ExcelUTF8.xlsx";
            string destination = "C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ExcelUTF8.csv";

            string sheet = "2"; // Use "2" for the second sheet (index or name)
            string delimiter = ";";
            string columns = "a:F"; // Column range to convert
            string rows = "1:8"; // Row range to extract

            var sheetHelper = new SheetHelper();
            bool success = sheetHelper.Converter(source, destination, sheet, delimiter, columns, rows);

            // Check if the conversion was successful
            Assert.That(success, Is.True, "Conversion process failed.");

            // Validate that the file exists after conversion
            Assert.That(File.Exists(destination), Is.True, $"The file '{destination}' was not created.");

            // Load the CSV file and validate its content
            var csvLines = File.ReadAllLines(destination);
            Assert.That(csvLines.Length, Is.GreaterThan(0), "The converted CSV file is empty.");

            // Check if the number of rows matches the expected row range
            int expectedRows = 8; // Based on rows "1:8"
            Assert.That(csvLines.Length, Is.EqualTo(expectedRows), $"Expected {expectedRows} rows but got {csvLines.Length}.");

            // Verify if the delimiter is correct
            Assert.That(csvLines[0].Contains(delimiter), Is.True, $"Expected delimiter '{delimiter}' not found in the output.");

            Assert.Multiple(() =>
            {
                // Check content G2 = "!@#$%&*() T_est"
                Assert.That(csvLines[1].Split(delimiter)[6], Is.EqualTo("!@#$%&*() T_est"), $"Expected '!@#$%&*() T_est' not found in the output.");

                // Check content D3 - "30/09/2024  21:40:13"
                Assert.That(csvLines[2].Split(delimiter)[2], Is.EqualTo("30/09/2024  21:40:13"), $"Expected '30/09/2024  21:40:13' not found in the output.");
            });
        }


        [Test, Repeat(1)]
        public void TestFormatColumns()
        {
            // Define paths for source and destination files
            string origem = @"C:\Users\diego\Desktop\Tests\Converter\TesteFormatoColunas.xlsx";
            string destino = @"C:\Users\diego\Desktop\Tests\Convertidos\TesteFormatoColunas.csv";

            // Specify parameters for the conversion
            string aba = "1"; // Use "1" for the first sheet (index or name)
            string separador = ";";
            string? colunas = "A:"; // Include all columns from A to the end
            string? linhas = ":"; // Include all rows

            var sh = new SheetHelper();

            // Perform the conversion
            bool retorno = sh.Converter(origem, destino, aba, separador, colunas, linhas);

            // Validate that the conversion was successful
            Assert.That(retorno, Is.True, "Conversion process failed.");

            // Check if the output file exists
            Assert.That(File.Exists(destino), Is.True, $"The file '{destino}' was not created.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destino);
            Assert.That(csvLines.Length, Is.GreaterThan(0), "The converted CSV file is empty.");

            // Optionally, check if the delimiter is correct in the first line of the CSV
            Assert.That(csvLines[0].Contains(separador), Is.True, $"Expected delimiter '{separador}' not found in the output.");

            // Validate if the number of columns is as expected
            int expectedColumns = 10; // Since we are selecting from column A to the end
            var firstLineColumns = csvLines[0].Split(separador);
            Assert.That(firstLineColumns.Length, Is.EqualTo(expectedColumns), $"Expected {expectedColumns} columns but got {firstLineColumns.Length}.");

            // Validate content of the first row
            Assert.That(firstLineColumns[0], Is.Not.Empty, "The first cell in the first row is empty.");

            // Validate content of the last row
            var lastLineColumns = csvLines[^1].Split(separador);
            Assert.That(lastLineColumns[1], Is.EqualTo("2,885649922405185"), "The second cell in the last row isn't 2,88564992240519 => 2,885649922405185.");

            // Validate the last column of the penultimate row
            var penultimateLineColumns = csvLines[^2].Split(separador);
            Assert.That(penultimateLineColumns[^1], Is.EqualTo("1,291666666666667"), "The last cell in the penultimate row isn't 1,29166666666667 => 1,291666666666667.");
        }



        [Test, Repeat(1)]
        public void TestTxtCsv()
        {
            // Define paths for source and destination files
            string origem = @"C:\Users\diego\Desktop\Tests\Converter\ColunasExcel.txt";
            string destino = @"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_txtCsv.csv";

            // Specify parameters for the conversion
            string aba = "1"; // Use "1" for the first sheet (index or name)
            string delimitador = ";";
            string? colunas = "A:"; // Include all columns from A to the end
            string? linhas = ":"; // Include all rows

            // Create an instance of SheetHelper and perform the conversion
            var sheetHelper = new SheetHelper();
            bool retorno = sheetHelper.Converter(origem, destino, aba, delimitador, colunas, linhas);

            // Validate that the conversion was successful
            Assert.That(retorno, Is.True, "Conversion process failed.");

            // Check if the output file exists
            Assert.That(File.Exists(destino), Is.True, $"The file '{destino}' was not created.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destino);
            Assert.That(csvLines.Length, Is.GreaterThan(0), "The converted CSV file is empty.");

            // Check if the first line contains the expected delimiter
            Assert.That(csvLines[0].Contains(delimitador), Is.True, $"Expected delimiter '{delimitador}' not found in the output.");

            // Validate the number of columns based on expected input from the TXT file
            int expectedColumns = 1; // Adjust this value based on the expected number of columns
            var firstLineColumns = csvLines[0].Split(delimitador);
            Assert.That(firstLineColumns.Length, Is.EqualTo(expectedColumns), $"Expected {expectedColumns} columns but got {firstLineColumns.Length}.");

            // Validate specific data if known          
            Assert.That(firstLineColumns[0], Is.EqualTo("ExpectedValue"), "The first column value is not as expected.");

            // Verify if the delimiter is correct
            Assert.That(csvLines[0].Contains(delimitador), Is.True, $"Expected delimiter '{delimitador}' not found in the output.");

            Assert.Multiple(() =>
            {
                // Check content of the first row -- 1
                Assert.That(csvLines[0].Split(delimitador)[0], Is.EqualTo("1"), $"Expected '1' not found in the output.");

                // Check content of the last row -- A101254
                Assert.That(csvLines[5].Split(delimitador)[0], Is.EqualTo("A101254"), $"Expected 'A101254' not found in the output.");

                // Check content of the penultimate row and last column -- AB101253               
                Assert.That(csvLines[4].Split(delimitador)[^1], Is.EqualTo("AB101253"), $"Expected 'AB101253' not found in the output.");
            });
        }


        [Test, Repeat(1)]
        public void TestUnzipOnly()
        {
            string origem = @"C:\Users\diego\Desktop\Tests\Converter\ColunasExcel.zip";
            string destino = @"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_unzipOnly.xlsx";

            string aba = "1"; // Use "1" for the first sheet (index or name)
            string separador = ";";

            string? colunas = null; // Include all columns
            string? linhas = null; // Include all rows

            // Perform the conversion
            var sh = new SheetHelper();
            bool retorno = sh.Converter(origem, destino, aba, separador, colunas, linhas);

            // Validate that the conversion was successful
            Assert.That(retorno, Is.True, "Conversion process failed.");

            // Check if the output file exists
            Assert.That(File.Exists(destino), Is.True, $"The file '{destino}' was not created.");

            var dataSet = sh.GetDataSet(origem);

            // Ensure the dataset has at least one table
            Assert.That(dataSet.Tables.Count, Is.GreaterThan(0), "The unzipped file does not contain any worksheets.");

            // Get the first table (worksheet)
            System.Data.DataTable worksheet = dataSet.Tables[aba]; // Use index 0 for the first table
            Assert.That(worksheet.Rows.Count, Is.GreaterThan(0), "Worksheet has no data.");

            // Validate the number of rows and columns in the worksheet
            int numberOfRows = worksheet.Rows.Count;
            int numberOfColumns = worksheet.Columns.Count;

            Assert.That(numberOfRows, Is.EqualTo(101254), "The worksheet has not 101254 rows.");
            Assert.That(numberOfColumns, Is.EqualTo(28), "The worksheet has not 28 columns.");

            // Validate the content of the first cell
            var firstCellValue = worksheet.Rows[0][0].ToString();
            Assert.That(firstCellValue, Is.EqualTo("1"), "The first cell in the worksheet isn't 1.");

            // Validate the last cell's content in the last row and last column
            var lastCellValue = worksheet.Rows[numberOfRows - 1][numberOfColumns - 1].ToString();
            Assert.That(lastCellValue, Is.EqualTo("AB101254"), "The last cell in the worksheet isn't AB101254.");
        }


        [Test, Repeat(1)]
        public void TestUnzip()
        {
            // Define paths for source and destination files
            string origem = @"C:\Users\diego\Desktop\Tests\Converter\ColunasExcel.zip";
            string destino = @"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_unzipXlsx.csv";

            string aba = "1"; // Specify the worksheet (sheet)
            string separador = ";"; // Define the delimiter

            // Specify column and row parameters as null to include all
            string? colunas = null;
            string? linhas = null;

            // Perform the conversion
            bool retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);

            // Validate that the conversion was successful
            Assert.That(retorno, Is.True, "Conversion process failed.");

            // Check if the output file exists
            Assert.That(File.Exists(destino), Is.True, $"The file '{destino}' was not created.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destino);
            Assert.That(csvLines.Length, Is.GreaterThan(0), "The converted CSV file is empty.");

            // Validate the content of the first row (header) and last row
            var firstRow = csvLines[0].Split(separador);
            Assert.That(firstRow.Length, Is.GreaterThan(0), "The first row in the CSV file has no columns.");
            Assert.That(firstRow[0], Is.Not.Empty, "The first cell in the first row is empty.");

            var lastRow = csvLines[^1].Split(separador);
            Assert.That(lastRow.Length, Is.GreaterThan(0), "The last row in the CSV file has no columns.");
            Assert.That(lastRow[0], Is.Not.Empty, "The first cell in the last row is empty.");

            // Optionally, validate the total number of rows (expected value may need to be adjusted based on your data)
            int expectedRowCount = 10; // Set the expected row count based on your input file
            Assert.That(csvLines.Length, Is.EqualTo(expectedRowCount), $"Expected {expectedRowCount} rows but got {csvLines.Length}.");
        }


        // --------------------------------------------------------------------------------

        [Test, Repeat(1)]
        public void TestConvertParticularXLSX()
        {
            // Define paths for source and destination files
            string origem = @"C:\Users\diego\Desktop\Tests\Converter\ColunasExcel.xlsx";
            string destino = @"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_xlsx.csv";

            // Specify parameters for the conversion
            string aba = "sheet1"; // Use "sheet1" for the first sheet by name
            string separador = ";";
            string? colunas = "A,2,c"; // Specify the columns to convert
            string? linhas = "1:3,4,-1"; // Specify the rows to extract (5)

            // Perform the conversion
            bool retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);

            // Validate that the conversion was successful
            Assert.That(retorno, Is.True, "Conversion process failed.");

            // Check if the output file exists
            Assert.That(File.Exists(destino), Is.True, $"The file '{destino}' was not created.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destino);
            Assert.That(csvLines.Length, Is.GreaterThan(0), "The converted CSV file is empty.");

            // Validate the number of rows matches the expected rows
            int expectedRows = 5; // Based on rows "1:3,4,-1" = 1, 2, 3, 4, -1
            Assert.That(csvLines.Length, Is.EqualTo(expectedRows), $"Expected {expectedRows} rows but got {csvLines.Length}.");

            // Validate the content of the first row
            var firstRow = csvLines[0].Split(separador);
            Assert.That(firstRow.Length, Is.EqualTo(4), "The first row in the CSV file has not 4 columns.");
            Assert.That(firstRow[0], Is.EqualTo("1"), "The first cell in the first row isn't 1.");

            // Validate the content of the last row
            var lastRow = csvLines[^1].Split(separador);
            Assert.That(lastRow.Length, Is.EqualTo(4), "The last row in the CSV file has not 4 columns.");
            Assert.That(lastRow[0], Is.EqualTo("A101254"), "The first cell in the last row isnt A101254.");

            // Validate specific cell values           
            var row = csvLines[4].Split(separador);
            Assert.That(row[1], Is.EqualTo("B101254"), "The value in the row is not as expected (B101254).");
        }

        [Test, Repeat(1)]
        public void TestConvertParticularXLSX_FullColumnsRows()
        {
            // Define paths for source and destination files
            string origem = @"C:\Users\diego\Desktop\Tests\Converter\ColunasExcel.xlsx";
            string destino = @"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_xlsx.csv";

            // Specify parameters for the conversion
            string aba = "sheet1"; // Use "sheet1" for the first sheet by name
            string separador = ";";
            string? colunas = null; // A:CV (1:100)
            string? linhas = null; // 1:101254

            // Perform the conversion
            bool retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);

            // Validate that the conversion was successful
            Assert.That(retorno, Is.True, "Conversion process failed.");

            // Check if the output file exists
            Assert.That(File.Exists(destino), Is.True, $"The file '{destino}' was not created.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destino);
            Assert.That(csvLines.Length, Is.GreaterThan(0), "The converted CSV file is empty.");

            // Validate the number of rows matches the expected rows
            int expectedRows = 101254; 
            Assert.That(csvLines.Length, Is.EqualTo(expectedRows), $"Expected {expectedRows} rows but got {csvLines.Length}.");

            // Validate the content of the first row
            var firstRow = csvLines[0].Split(separador);
            Assert.That(firstRow.Length, Is.GreaterThan(0), "The first row in the CSV file has no columns.");
            Assert.That(firstRow[0], Is.Not.Empty, "The first cell in the first row is empty.");

            // Validate the content of the last row
            var lastRow = csvLines[^1].Split(separador);
            Assert.That(lastRow.Length, Is.GreaterThan(0), "The last row in the CSV file has no columns.");
            Assert.That(lastRow[0], Is.Not.Empty, "The first cell in the last row is empty.");

            // Validate specific cell values           
            var rowLast = csvLines[101253].Split(separador);
            Assert.That(rowLast[10], Is.EqualTo("K101254"), "The value in the row is not as expected (K101254).");
        }


        [Test, Repeat(1)]
        public void TestConvertParticularCSV()
        {
            // Define paths for source and destination files
            string origem = @"C:\Users\diego\Desktop\Tests\Converter\ColunasExcel.csv";
            string destino = @"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_csv.csv";

            // Specify parameters for the conversion
            string aba = "1"; // Not applicable for CSV, but included for consistency
            string separador = ";";
            string? colunas = null; // Convert all columns
            string linhas = "1:3"; // Specify the rows to extract

            // Perform the conversion
            bool retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);

            // Validate that the conversion was successful
            Assert.That(retorno, Is.True, "Conversion process failed.");

            // Check if the output file exists
            Assert.That(File.Exists(destino), Is.True, $"The file '{destino}' was not created.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destino);
            Assert.That(csvLines.Length, Is.GreaterThan(0), "The converted CSV file is empty.");

            // Validate the number of rows matches the expected rows
            int expectedRows = 3; // Based on rows "1:3"
            Assert.That(csvLines.Length, Is.EqualTo(expectedRows), $"Expected {expectedRows} rows but got {csvLines.Length}.");

            // Validate the content of the first row
            var firstRow = csvLines[0].Split(separador);
            Assert.That(firstRow.Length, Is.GreaterThan(0), "The first row in the CSV file has no columns.");
            Assert.That(firstRow[0], Is.EqualTo("1"), "The first cell in the first row isn't 1.");

            // Validate the content of the last row
            var lastRow = csvLines[^1].Split(separador);
            Assert.That(lastRow.Length, Is.GreaterThan(0), "The last row in the CSV file has no columns.");
            Assert.That(lastRow[27], Is.EqualTo("AB3"), "The 27th cell in the last row isn't AB3.");       
        }

        [Test, Repeat(1)]
        public void TestConvertParticularCSV_FullColumnsRows()
        {
            // Define paths for source and destination files
            string origem = @"C:\Users\diego\Desktop\Tests\Converter\ColunasExcel.csv";
            string destino = @"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_csv.csv";

            // Specify parameters for the conversion
            string aba = "1"; // Not applicable for CSV, but included for consistency
            string separador = ";";
            string? colunas = null; // A:CV (1:100)
            string? linhas = null; // 1:101254

            // Perform the conversion
            bool retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);

            // Validate that the conversion was successful
            Assert.That(retorno, Is.True, "Conversion process failed.");

            // Check if the output file exists
            Assert.That(File.Exists(destino), Is.True, $"The file '{destino}' was not created.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destino);
            Assert.That(csvLines.Length, Is.GreaterThan(0), "The converted CSV file is empty.");

            // Validate the number of rows matches the expected rows
            int expectedRows = 101254; // Based on rows "null" => full => 1:101254
            Assert.That(csvLines.Length, Is.EqualTo(expectedRows), $"Expected {expectedRows} rows but got {csvLines.Length}.");

            // Validate the content of the first row
            var firstRow = csvLines[0].Split(separador);
            Assert.That(firstRow.Length, Is.GreaterThan(0), "The first row in the CSV file has no columns.");
            Assert.That(firstRow[0], Is.EqualTo("1"), "The first cell in the first row isn't 1.");

            // Validate the content of the last row
            var lastRow = csvLines[^1].Split(separador);
            Assert.That(lastRow.Length, Is.EqualTo(100), "The last row in the CSV file has not 100 columns.");
            Assert.That(lastRow[27], Is.EqualTo("AB101254"), "The 101254th cell in the last row isn't AB101254."); // Last fill

            // Validate specific cell values in the last row
            var rowLast = csvLines[101253].Split(separador);
            Assert.That(rowLast[15], Is.EqualTo("P101254"), "The value in the row is not as expected (P101254).");
        }


        [Test]
        public void TestConvertFromCSV()
        {
            // Define paths for source and destination files
            string source = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\SemCabecalho.csv";
            string destination = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\SemCabecalho_csv.xls";

            // Specify parameters for the conversion
            string sheet = "1"; // Use "1" for the first sheet (index or name)
            string delimiter = ";"; // Define the delimiter for the output file
            string? columns = null; // Convert all columns (A:AB) = 28
            string rows = "1:5"; // Extract rows 1 to 5

            var sh = new SheetHelper();

            // Get the data table from the source file
            var dataTable = sh.GetDataTable(source, sheet);

            // Ensure the data table is not null and has rows
            Assert.That(dataTable, Is.Not.Null, "The data table is null.");
            Assert.That(dataTable.Rows.Count, Is.GreaterThan(0), "The data table has no rows.");

            // Get the first row as an array
            var firstRow = sh.GetRowArray(dataTable);

            // Check if the first row has expected number of columns
            Assert.That(firstRow.Length, Is.GreaterThan(0), "The first row is empty.");

            // Save the data table to the destination file
            var success = sh.SaveDataTable(dataTable, destination, delimiter, columns, rows);

            // Validate that the conversion was successful
            Assert.That(success, Is.EqualTo(true), "Conversion process failed.");

            // Check if the output file exists
            Assert.That(File.Exists(destination), Is.True, $"The file '{destination}' was not created.");

            // Load the converted file and validate its content
            var savedData = File.ReadAllLines(destination);
            Assert.That(savedData.Length, Is.GreaterThan(0), "The saved file is empty.");

            // Validate the number of lines saved matches the expected rows
            int expectedRows = 5; // Based on rows "1:5"
            Assert.That(savedData.Length, Is.EqualTo(expectedRows), $"Expected {expectedRows} rows but got {savedData.Length}.");

            // Validate the content of the first cell in the first row
            int expectedColumns = 28; // Based on columns "A:AB"
            var firstRowSaved = savedData[0].Split(delimiter);
            Assert.That(firstRowSaved.Length, Is.EqualTo(expectedColumns), "The first cell in the saved file has not 28 columns.");
            Assert.That(firstRowSaved[9], Is.EqualTo("J2"), "The 10th cell in the saved file is not J2.");
        }


        [Test]
        public void TestReadCSV()
        {
            // Define the path for the source CSV file
            string source = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.csv";
            string sheet = "1"; // Specify the sheet to read (for CSV, it's typically not used)

            var sh = new SheetHelper();

            // Get the data table from the source CSV file
            var dataTable = sh.GetDataTable(source, sheet);

            // Ensure the data table is not null and has rows
            Assert.That(dataTable, Is.Not.Null, "The data table is null.");
            Assert.That(dataTable.Rows.Count, Is.GreaterThan(0), "The data table has no rows.");

            // Get the first row as an array
            var firstRow = sh.GetRowArray(dataTable);

            // Ensure the first row has enough columns
            Assert.That(firstRow.Length, Is.GreaterThan(2), "The first row does not have enough columns.");

            // Validate the value in the third column (index 2)
            Assert.That(firstRow[2], Is.EqualTo("3"), "The value in the third column is not as expected (3).");
        }


        [Test]
        public void TestReadCSVIrregularHeader()
        {
            // Define the path for the source CSV file with irregular headers
            string source = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\CabecalhoIrregular.csv";

            var sh = new SheetHelper();

            // Add any specific exceptions to ignore during the conversion process
            sh.TryIgnoreExceptions.AddRange(new List<string> { "E-4041-SH" });

            // Get the data table from the source CSV file
            var dataTable = sh.GetDataTable(source);

            // Ensure the data table is not null
            Assert.That(dataTable, Is.Not.Null, "The data table is null.");

            // Ensure the data table has rows
            Assert.That(dataTable.Rows.Count, Is.GreaterThan(0), "The data table has no rows.");

            // Get the first row as an array
            var firstRow = sh.GetRowArray(dataTable);

            // Ensure the first row has enough columns
            Assert.That(firstRow.Length, Is.EqualTo(6), "The first row does not have enough columns.");

            // Validate specific cell values for irregular headers
            Assert.That(firstRow[5], Is.EqualTo("6"), "The value in the sixth column is not as expected (6).");
            Assert.That(firstRow[6], Is.EqualTo("ColumnEmpty6"), "The value in the seventh column is not as expected (Column6).");
        }


        [Test]
        public void TestReadCSVWithoutHeader() // Read
        {
            // Define the path for the source CSV file without a header
            string source = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\SemCabecalho.csv";
            string sheet = "1"; // Specify the sheet (not typically used for CSV)

            var sh = new SheetHelper();

            // Get the data table from the source CSV file
            var dataTable = sh.GetDataTable(source, sheet);

            // Ensure the data table is not null
            Assert.That(dataTable, Is.Not.Null, "The data table is null.");

            // Ensure the data table has rows
            Assert.That(dataTable.Rows.Count, Is.GreaterThan(0), "The data table has no rows.");

            // Get the first row as an array
            var firstRow = sh.GetRowArray(dataTable);

            // Ensure the first row has enough columns
            Assert.That(firstRow.Length, Is.GreaterThan(2), "The first row does not have enough columns.");

            // Validate specific cell values
            Assert.That(firstRow[2], Is.EqualTo("C2"), "The value in the third column is not as expected (C2).");
        }


        [Test]
        public void TestReadCSVCabecalhoVazio() // Read
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\CabecalhoVazio.csv";
            string aba = "1";

            var sh = new SheetHelper();
            sh.TryIgnoreExceptions.AddRange(new List<string>() { "E-4041-SH" });
            var dt = sh.GetDataTable(origem, aba);
            var first = sh.GetRowArray(dt);

            Assert.That(first[2].Equals("3"), Is.EqualTo(true));
        }

        [Test]
        public void TestReadCSVEmptyHeader() // Read
        {
            // Define the path for the source CSV file with an empty header
            string source = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\CabecalhoVazio.csv";
            string sheet = "1"; // Specify the sheet (not typically used for CSV)

            var sh = new SheetHelper();

            // Add exceptions to ignore during processing
            sh.TryIgnoreExceptions.AddRange(new List<string>() { "E-4041-SH" });

            // Get the data table from the source CSV file
            var dataTable = sh.GetDataTable(source, sheet);

            // Ensure the data table is not null
            Assert.That(dataTable, Is.Not.Null, "The data table is null.");

            // Ensure the data table has rows
            Assert.That(dataTable.Rows.Count, Is.GreaterThan(0), "The data table has no rows.");

            // Get the first row as an array
            var firstRow = sh.GetRowArray(dataTable);

            // Ensure the first row has enough columns
            Assert.That(firstRow.Length, Is.GreaterThan(2), "The first row does not have enough columns.");

            // Validate the specific value in the third column
            Assert.That(firstRow[2], Is.EqualTo("3"), "The value in the third column is not as expected (3).");
        }


        [Test, Repeat(1)]
        public void TestConvertParticularXLSB()
        {
            // Define the path for the source XLSB file
            string source = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xlsb";
            string destination = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_xlsb.csv";

            string sheet = "1"; // Specify the sheet to convert
            string delimiter = ";"; // Set the delimiter for the CSV output
            string? columns = null; // Specify columns to convert (null for all)
            string? rows = null; // Specify rows to convert (null for all)

            var sheetHelper = new SheetHelper();

            // Perform the conversion
            bool success = sheetHelper.Converter(source, destination, sheet, delimiter, columns, rows);

            // Validate that the conversion was successful
            Assert.That(success, Is.EqualTo(true), "Conversion from XLSB to CSV failed.");

            // Validate that the destination file was created
            Assert.That(File.Exists(destination), Is.True, $"The file '{destination}' was not created.");

            // Load the CSV file and validate its content (optional)
            var csvLines = File.ReadAllLines(destination);
            Assert.That(csvLines.Length, Is.GreaterThan(0), "The converted CSV file is empty.");

            // Optionally, validate the first line for expected delimiter presence
            Assert.That(csvLines[0].Contains(delimiter), Is.True, $"Expected delimiter '{delimiter}' not found in the output.");
        }


        [Test]
        public void TestConvertTryIgnoreExpect_ColumnOutOfRange()
        {
            // Define the paths for the source and destination files
            string source = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\Small.xlsx";
            string destination = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\Small_xlsx.csv";

            string sheet = "1"; // Specify the sheet to convert
            string delimiter = ";"; // Set the delimiter for the CSV output
            string? columns = "1:12"; // Attempting to select columns 1 to 12 (max column is 10)
            string? rows = "1:10"; // Specify the rows to convert

            var sheetHelper = new SheetHelper();

            // Add exceptions to ignore during conversion
            sheetHelper.TryIgnoreExceptions.AddRange(new List<string>() { "E-4042-SH" });

            // Perform the conversion while ignoring exceptions
            bool success = sheetHelper.Converter(source, destination, sheet, delimiter, columns, rows);
            Assert.That(success, Is.EqualTo(true), "Conversion process failed while ignoring exceptions.");

            // Clear the exceptions list for the next test
            sheetHelper.TryIgnoreExceptions.Clear();

            // Assert that an exception is thrown when trying to convert without ignoring exceptions
            Assert.Throws<SH.Exceptions.ColumnOutRangeSHException>(
                () => sheetHelper.Converter(source, destination, sheet, delimiter, columns, rows),
                "Expected ColumnOutRangeSHException was not thrown.");
        }


        [Test]
        public void TestConvertTryIgnoreExpect_FileOriginOpened()
        {
            // Define the paths for the source and destination files
            string source = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\Small.csv";
            string destination = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\Small_csv.txt";

            string sheet = "1"; // Specify the sheet to convert (not applicable for CSV but kept for consistency)
            string delimiter = ";"; // Set the delimiter for the output file
            string? columns = null; // No specific columns to convert
            string? rows = ""; // No specific rows to convert

            var sheetHelper = new SheetHelper();

            // First conversion attempt (should succeed)
            bool initialConversionSuccess = sheetHelper.Converter(source, destination, sheet, delimiter, columns, rows);
            Assert.That(initialConversionSuccess, Is.EqualTo(true), "Initial conversion failed unexpectedly.");

            // Open the source file in write mode to simulate it being in use
            File.OpenWrite(source); // This should throw an exception

            // Clear any ignored exceptions for the next test
            sheetHelper.TryIgnoreExceptions.Clear();

            // Assert that an exception is thrown when trying to convert with the source file in use
            Assert.Throws<SH.Exceptions.FileOriginInUseSHException>(
                () => sheetHelper.Converter(source, destination, sheet, delimiter, columns, rows),
                "Expected FileOriginInUseSHException was not thrown when the source file is in use."
            );

            // Add exception to ignore for the next assertion
            sheetHelper.TryIgnoreExceptions.AddRange(new List<string>() { "E-0541-SH" });

            // Assert that the exception is still thrown even when ignoring the specified exception
            Assert.Throws<SH.Exceptions.FileOriginInUseSHException>(
                () => sheetHelper.Converter(source, destination, sheet, delimiter, columns, rows),
                "Expected FileOriginInUseSHException was not thrown even when ignoring specific exceptions."
            );
        }


        [Test]
        public void TestConvert_FileNotExists()
        {
            // Define the paths for the non-existent source file and destination file
            string sourceFilePath = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\FileNotExists.xlsx";
            string destinationFilePath = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\Small_csv.txt";

            string sheet = "1"; // Specify the sheet to convert (not applicable for non-existent file)
            string delimiter = ";"; // Set the delimiter for the output file
            string? columns = null; // No specific columns to convert
            string? rows = ""; // No specific rows to convert

            var sheetHelper = new SheetHelper();

            // Assert that an exception is thrown when trying to convert a non-existent file
            var exception = Assert.Throws<SH.Exceptions.FileNotFoundSHException>(
                () => sheetHelper.Converter(sourceFilePath, destinationFilePath, sheet, delimiter, columns, rows),
                "Expected FileNotFoundSHException was not thrown when converting a non-existent file."
            );

            // Optionally, you can verify the exception message if your implementation provides it
            Assert.That(exception.Message, Does.Contain("E-4048-SH"), "The exception message does not indicate a file not found error.");
        }


        [Test]
        public void TestConvert_FileNotExistsFolder()
        {
            // Define the path for the non-existent source file in a non-existent directory
            string sourceFilePath = "C:\\Users\\diego\\Desktop\\Tests\\DirectoryNotExists\\FileNotExists.xlsx";
            string destinationFilePath = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\Small_csv.txt";

            string sheet = "1"; // Specify the sheet to convert (not applicable for non-existent file)
            string delimiter = ";"; // Set the delimiter for the output file
            string? columns = null; // No specific columns to convert
            string? rows = ""; // No specific rows to convert

            var sheetHelper = new SheetHelper();

            // Assert that an exception is thrown when trying to convert a file from a non-existent directory
            var exception = Assert.Throws<SH.Exceptions.FileNotFoundSHException>(
                () => sheetHelper.Converter(sourceFilePath, destinationFilePath, sheet, delimiter, columns, rows),
                "Expected FileNotFoundSHException was not thrown when converting a file from a non-existent directory."
            );

            // Optionally, you can verify the exception message if your implementation provides it
            Assert.That(exception.Message, Does.Contain("E-4048-SH"), "The exception message does not indicate a file not found error.");
        }


        [Test]
        public void TestConvert_FileNotSupport()
        {
            // Define the path for a non-supported file type
            string sourceFilePath = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\FileNotSupport.xps";
            string destinationFilePath = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\Small_csv.txt";

            string sheet = "1"; // Specify the sheet to convert (not applicable for non-supported files)
            string delimiter = ";"; // Set the delimiter for the output file
            string? columns = null; // No specific columns to convert
            string? rows = ""; // No specific rows to convert

            var sheetHelper = new SheetHelper();

            // Assert that an exception is thrown when trying to convert a file type that is not supported
            var exception = Assert.Throws<SH.Exceptions.FileOriginNotReadSupportSHException>(
                () => sheetHelper.Converter(sourceFilePath, destinationFilePath, sheet, delimiter, columns, rows),
                "Expected FileOriginNotReadSupportSHException was not thrown when converting a non-supported file type."
            );

            // Optionally, verify the exception message to ensure it indicates the file type issue
            Assert.That(exception.Message, Does.Contain("E-0541-SH"), "The exception message does not indicate that the file type is not supported.");
        }


        [Test]
        public void TestConvert_FolderDestinationNotExists()
        {
            // Define the source file path and destination file path
            string sourceFilePath = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\Small.xlsx";
            string destinationFilePath = $"C:\\Users\\diego\\Desktop\\Tests\\FolderNotExists\\Small_csv.txt";

            string sheet = "1"; // Specify the sheet to convert
            string delimiter = ";"; // Set the delimiter for the output file
            string? columns = null; // No specific columns to convert
            string? rows = ""; // No specific rows to convert

            var sheetHelper = new SheetHelper();

            // Perform the conversion to ensure the initial conversion works
            bool initialConversionResult = sheetHelper.Converter(sourceFilePath, destinationFilePath, sheet, delimiter, columns, rows);
            Assert.That(initialConversionResult, Is.EqualTo(true), "Initial conversion should succeed.");

            // Delete the destination folder to simulate the non-existence of the folder
            Directory.Delete(Path.GetDirectoryName(destinationFilePath), true);

            // Clear any previously ignored exceptions
            sheetHelper.TryIgnoreExceptions.Clear();

            // Assert that an exception is thrown when attempting to convert to a non-existing destination folder
            var exception = Assert.Throws<SH.Exceptions.DirectoryDestinationNotFoundSHException>(
                () => sheetHelper.Converter(sourceFilePath, destinationFilePath, sheet, delimiter, columns, rows),
                "Expected DirectoryDestinationNotFoundSHException was not thrown when the destination folder does not exist."
            );

            // Optionally, verify the exception message to ensure it indicates the folder not found issue
            Assert.That(exception.Message, Does.Contain("E-4049-SH"), "The exception message does not indicate that the destination directory is not found.");
        }

        [Test]
        public void TestConvertOneSheetBig()
        {
            // Define the source and destination file paths
            string sourceFilePath = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ExcelBig_OneSheetBig_ABN204960.xlsx";
            string destinationFilePath = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ExcelBig_OneSheetBig_ABN204960_xlsx.csv";

            string sheet = "1"; // Specify the sheet index to convert
            string delimiter = ";"; // Set the delimiter for the output CSV
            string? columns = ""; // Specify columns to convert ("" means all columns)
            string? rows = "1:"; // Specify rows to convert (starting from the first row to the last)

            var sheetHelper = new SheetHelper();

            // Perform the conversion and check if it was successful
            bool conversionResult = sheetHelper.Converter(sourceFilePath, destinationFilePath, sheet, delimiter, columns, rows);

            // Assert that the conversion was successful
            Assert.That(conversionResult, Is.EqualTo(true), "The conversion from Excel to CSV for the large sheet should succeed.");
        }

        [Test]
        public void TestConvertManySheetBig()
        {
            Assert.Ignore("Skipping long test ConvertManySheetBig!");

            // Define the source and destination file paths
            string sourceFilePath = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ExcelBig_ManySheetBig_AB1048576.xlsx";
            string destinationFilePath = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ExcelBig_ManySheetBig_AB1048576_xlsx.csv";

            string sheet = "1"; // Specify the sheet index to convert
            string delimiter = ";"; // Set the delimiter for the output CSV
            string? columns = ""; // Specify columns to convert ("" means all columns)
            string? rows = "1:"; // Specify rows to convert (starting from the first row to the last)

            var sheetHelper = new SheetHelper();

            // Set prohibited items to be replaced during conversion
            sheetHelper.ProhibitedItems = new Dictionary<string, string>
            {
                { "\n", " " }, // Replace newline characters with space
                { "\r", " " }, // Replace carriage return characters with space
                { ";", "," },   // Replace semicolons with commas
            };

            // Perform the conversion and check if it was successful
            bool conversionResult = sheetHelper.Converter(sourceFilePath, destinationFilePath, sheet, delimiter, columns, rows);

            // Assert that the conversion was successful
            Assert.That(conversionResult, Is.EqualTo(true), "The conversion from Excel to CSV for multiple large sheets should succeed.");
        }


        [Test]
        public void TestConvertMultiDestinations()
        {
            // Define the source and destination file paths
            string sourceFilePath = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\AbasExcel.xlsx";
            string destinationFilePath = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\AbasExcel_xlsx.csv";

            // Specify the sheets to convert
            var sheets = new List<string>() { "2", "1", "sheet2" };
            string delimiter = ";"; // Set the delimiter for the output CSV

            // Define the columns to convert for each sheet
            string[]? columns = new string[] { "A, B:C", "1:10", "B,A" };

            // Define the rows to convert for each sheet
            List<string>? rows = new() { "1:3", "1:10", "1" };

            int minimumRows = 1; // Specify the minimum number of rows to be converted

            // Perform the conversion and check if the number of successful conversions matches the number of sheets
            var conversionResult = new SheetHelper().Converter(sourceFilePath, destinationFilePath, sheets, delimiter, columns, rows, minimumRows);

            // Assert that the conversion result matches the expected number of sheets
            Assert.That(conversionResult, Is.EqualTo(sheets.Count), "The conversion should succeed for the specified number of sheets.");
        }


        [Test]
        public void TestConvertMultiSheets1()
        {
            // Define the source and destination file paths
            string sourceFilePath = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\AbasExcel.xlsx";
            string destinationFilePath = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\AbasExcel_xlsx.csv";

            // Specify the sheets to convert
            var sheetsToConvert = new List<string>() { "aba 3", "1", "aba 3" }; // Repeating "aba 3" to test multiple conversions

            string delimiter = ";"; // Set the delimiter for the output CSV
            string[]? columns = null; // No specific columns to convert
            List<string>? rows = new() { "1:3", "1:10", "1" }; // Specify the row ranges to convert for each sheet
            int minimumRows = 1; // Set the minimum number of rows required for conversion

            // Perform the conversion
            var conversionResult = new SheetHelper().Converter(sourceFilePath, destinationFilePath, sheetsToConvert, delimiter, columns, rows, minimumRows);

            // Assert that the conversion result matches the expected number of sheets processed
            Assert.That(conversionResult, Is.EqualTo(sheetsToConvert.Count), "The conversion should succeed for the specified number of sheets.");
        }


        [Test]
        public void TestConvertMultiSheets2()
        {
            // Initialize the SheetHelper
            var sheetHelper = new SheetHelper();

            // Define the source and destination file paths
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\AbasExcel.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\AbasExcel_xlsx.csv";

            // Set the sheet names (null indicates a missing parameter)
            string? abas = null;
            string separador = ";"; // Define the separator
            string[]? colunas = null; // Column selection is also null
            List<string>? linhas = null; // Row selection is null
            int minRows = 1; // Minimum required rows

            // Assert that a specific exception is thrown for missing parameters
            Assert.Throws<SH.Exceptions.ParamMissingConverterSHException>(() =>
                sheetHelper.Converter(origem, destino, abas, separador, colunas, linhas, minRows)
            );
        }


        [Test]
        public void TestConvertMultiSheets3()
        {
            // Define the source and destination file paths
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\AbasExcel.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\AbasExcel_xlsx.csv";

            // Specify the sheet names to convert
            var abas = new List<string>() { "Sheet2", "1", "aba 3" }; // "Sheet1" | "Sheet2" | "aba 3"
            string delimitador = ";"; // Define the separator for the CSV

            // Column selection is not specified; null indicates all columns will be included
            string[]? colunas = null;
            // Row selections for each sheet; the rows to be included in the conversion
            List<string>? linhas = new() { "1:3, -1", "1:10", "1" };
            int minRows = 1; // Minimum required rows for conversion

            // Instantiate SheetHelper and perform the conversion
            var sheetHelper = new SheetHelper();
            int retorno = sheetHelper.Converter(origem, destino, abas, delimitador, colunas, linhas, minRows);

            // Assert that the conversion was successful for the expected number of sheets
            Assert.That(retorno, Is.EqualTo(abas.Count), "The number of successfully converted sheets does not match the expected count.");

            // Verify if the output file (sheet 2 in file - name = "Sheet2")
            string destino2 = Path.Combine(Path.GetDirectoryName(destino), Path.GetFileNameWithoutExtension(destino) + "__Sheet2.csv");

            // Verify the output file existence and content
            Assert.IsTrue(File.Exists(destino2), "The output file does not exist after conversion.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destino2);

            // Verify if the delimiter is correct
            Assert.That(csvLines[0].Contains(delimitador), Is.True, $"Expected delimiter '{delimitador}' not found in the output.");

            Assert.Multiple(() =>
            {
                // Check content of the first row -- 1
                Assert.That(csvLines[0].Split(delimitador)[0], Is.EqualTo("1"), $"Expected '1' not found in the output.");

                // Check content of the last row -- 1:3, -1 => -1 => 10
                Assert.That(csvLines[3].Split(delimitador)[0], Is.EqualTo("A10"), $"Expected 'A10' not found in the output.");

                // Check content of the penultimate row and last column -- 1:3, -1 => 3 => J3          
                Assert.That(csvLines[2].Split(delimitador)[^1], Is.EqualTo("J3"), $"Expected 'J3' not found in the output.");
            });
        }

        [Test]
        public void TestProhibitedItems()
        {
            // Define the source and destination file paths
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ProhibitedItems.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ProhibitedItems_xlsx.csv";

            string aba = "3"; // Specify the sheet to convert
            string separador = ";"; // Define the separator for the CSV
            string? colunas = null; // Column selection is not specified; null indicates all columns will be included
            string? linhas = null; // Row selection is not specified; null indicates all rows will be included

            // Define prohibited items in a dictionary
            var dic = new Dictionary<string, string>
            {
                { "\n", " " },
                { "\r", " " },
                { ";", "," },
            };

            // Instantiate SheetHelper
            var sh = new SheetHelper();
            sh.ProhibitedItems = dic; // Set prohibited items in SheetHelper

            // Convert the dictionary to JSON format for verification
            string jsonDictionary = System.Text.Json.JsonSerializer.Serialize(dic);
            string jsonDictionaryFromHelper = sh.GetJsonDictionary(dic);

            // Check if the dictionary can be retrieved correctly from JSON
            var retrievedDict = sh.GetDictionaryJson(jsonDictionaryFromHelper);

            // Perform the conversion
            var retorno = sh.Converter(origem, destino, aba, separador, colunas, linhas);

            // Assert that the conversion was successful
            Assert.That(retorno, Is.EqualTo(true), "Conversion did not succeed as expected.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destino);

            // Check if the prohibited items have been replaced as intended
            foreach (var kvp in dic)
            {
                Assert.That(string.Join(" ", csvLines).Contains(kvp.Value),
                            Is.True,
                            $"Expected '{kvp.Value}' not found in the output for prohibited item '{kvp.Key}'.");
            }
        }

        [Test]
        public void TestConvertRowsToBack()
        {
            // Define the source and destination file paths
            string origin = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xlsx";
            string destination = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_BackRows_xlsx.csv";

            // Specify the sheet name to convert
            string sheet = "1";
            string separator = ";"; // Define the separator for the CSV
            string? columns = ""; // All columns will be included

            // Specify the rows to convert; negative indices indicate counting from the end
            string? rows = "3:1, -1:-2; -2:-1";

            // Instantiate SheetHelper
            var sh = new SheetHelper();

            // Perform the conversion
            bool result = sh.Converter(origin, destination, sheet, separator, columns, rows);

            // Assert that the conversion was successful
            Assert.That(result, Is.EqualTo(true), "The conversion did not succeed as expected.");

            // Verify if the output file was created
            Assert.IsTrue(File.Exists(destination), "The output file was not created.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destination);

            // Checks can be performed here to validate the content of the csvLines       
            // "3:1, -1:-2; -2:-1" => 3,2,1,-1,-2,-2,-1 => 3+2+2 = 7 rows
            Assert.That(csvLines.Length, Is.EqualTo(7), "The output CSV file is not 7 rows.");
        }


        [Test]
        public void TestConvertColumnsToBack()
        {
            // Define the source and destination file paths
            string origin = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xlsx";
            string destination = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_BackColumns_xlsx.csv";

            // Specify the sheet name to convert
            string sheet = "1";
            string separator = ";"; // Define the separator for the CSV

            // Specify the columns to convert; negative indices indicate counting from the end
            string? columns = "C:A, A; -1, -2:-3; -3:-1";

            // Specify the rows to convert
            string? rows = "1:3, 1, 3:1";

            // Instantiate SheetHelper
            var sh = new SheetHelper();

            // Perform the conversion
            bool result = sh.Converter(origin, destination, sheet, separator, columns, rows);

            // Assert that the conversion was successful
            Assert.That(result, Is.EqualTo(true), "The conversion did not succeed as expected.");

            // Verify if the output file was created
            Assert.IsTrue(File.Exists(destination), "The output file was not created.");

            // Load the converted CSV file and validate its content
            var csvLines = File.ReadAllLines(destination);

            // Checks count rows
            // "1:3, 1, 3:1" => 1, 2, 3, 1, 3, 2, 1=> 7 rows
            Assert.That(csvLines.Length, Is.EqualTo(7), "The output CSV file haven't 7 rows.");

            // Checks can be performed here to validate the content of the csvLines
            // "C:A, A; -1, -2:-3; -3:-1" => C, B, A, A, -1, -2, -3, -3, -2, -1 => 10 columns
            Assert.That(csvLines[0].Split(separator).Length-1, Is.EqualTo(10), "The output CSV file haven't 10 coluns.");

            // Validate specific content in the CSV lines - penultimate column of the first row
            Assert.That(csvLines[0].Split(separator)[^3], Is.EqualTo("99"));
        }

        [Test]
        public void TestAllFeaturesAccess()
        {
            var sh = new SheetHelper();

            var progress = sh.Progress;
            sh.ProhibitedItems = null;
            sh.TryIgnoreExceptions = null;

            Assert.Pass();

            sh.CloseExcel(null);
            sh.GetIndexColumn(null);
            sh.GetNameColumn(1);
            sh.UnGZ(null, null);
            sh.UnZIP(null, null);
            sh.UnzipAuto(null, null, false);
            sh.ConvertToDataRow(null, null);
            sh.GetRowArray(null);
            sh.GetAllSheets(null);
            sh.NormalizeText(" Hot Café");
            sh.FixItems(null);
            sh.GetDictionaryJson(null);
            sh.GetJsonDictionary(null);
            sh.GetDataSet(null);
            sh.GetDataTable(null, null);
            sh.SaveDataTable(null, null, null, null, null);
            sh.Converter(null, null, null, null, null, null);
            sh.ConvertAllSheets(null, null, 0, null);  
        }


        // --------------------------------------------------------------------------------

        [TestCase("2:", 1, ExpectedResult = true, TestName = "2:")]
        [TestCase(":10", 2, ExpectedResult = true, TestName = ":10")]
        [TestCase("1:20", 3, ExpectedResult = true, TestName = "1:20")]
        [TestCase("", 4, ExpectedResult = true, TestName = "Linhas String vazia")]
        [TestCase(null, 5, ExpectedResult = true, TestName = "Linhas Nulo")]
        [TestCase("1, 2, 4", 6, ExpectedResult = true, TestName = "1, 2, 4")]
        [TestCase("7", 7, ExpectedResult = true, TestName = "7")]
        [TestCase("-1", 8, ExpectedResult = true, TestName = "-1")]
        [TestCase("1:-2", 9, ExpectedResult = true, TestName = "1:-2")]
        [TestCase("1,2,3,-3:", 10, ExpectedResult = true, TestName = "1,2,3,-3:-1")]
        public bool TestRows(string linhas, int id)
        {
            const string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.csv";
            string destino = @$"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_row{id}.csv";

            const string aba = "1";
            const string separador = ";";
            const string colunas = "A, 2,c";

            var sh = new SheetHelper();
            bool result = sh.Converter(origem, destino, aba, separador, colunas, linhas);

            // Verificação básica de que o arquivo de destino foi criado
            Assert.That(File.Exists(destino), Is.True, $"O arquivo de destino {destino} não foi criado.");

            // Carrega o conteúdo gerado e validar as linhas
            var linhasGeradas = File.ReadAllLines(destino);
            Assert.That(linhasGeradas.Length, Is.GreaterThan(0), "O arquivo CSV gerado está vazio.");

            return result;
        }

        [TestCase("A, C, b", 1, ExpectedResult = true, TestName = "Uppercase and lowercase columns out of order...")]
        [TestCase("A:D", 2, ExpectedResult = true, TestName = "Continuous range of columns...")]
        [TestCase("", 3, ExpectedResult = true, TestName = "Columns with empty string...")]
        [TestCase(" ", 4, ExpectedResult = true, TestName = "Empty columns...")]
        [TestCase(null, 5, ExpectedResult = true, TestName = "Null columns...")]
        public bool TestColumns(string colunas, int id)
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xls";
            string destino = @$"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_column{id}.csv";

            string aba = "1";
            string separador = ";";
            string linhas = "2:";

            var sh = new SheetHelper();
            bool result = sh.Converter(origem, destino, aba, separador, colunas, linhas);

            // Basic check to ensure the destination file was created
            Assert.IsTrue(File.Exists(destino), $"The destination file {destino} was not created.");

            // Optional: Load the generated content and validate columns, if necessary
            var linhasGeradas = File.ReadAllLines(destino);
            Assert.That(linhasGeradas.Length, Is.GreaterThan(0), "The generated CSV file is empty.");

            // Here you can add more specific validations about the content if needed

            return result;
        }


        // --------------------------------------------------------------------------------



        [TestFixture]
        public class TestsFormats
        {
            private const string SheetName = "1";
            private const string Separator = ";";
            private const string Columns = "A, b, 4, -1";
            private const string Rows = ":10, -1";

            [Test, TestCaseSource(typeof(TestFormatCases), nameof(TestFormatCases.FormatConversions))]
            public void ValidateValidFormats(string origin, string destination)
            {
                bool conversionResult = TestFormats(origin, destination);
                Assert.IsTrue(conversionResult, $"Conversion from '{origin}' to '{destination}' failed.");

                // Optionally, you can also assert if the destination file exists.
                Assert.IsTrue(File.Exists(destination), $"Destination file '{destination}' was not created.");
            }

            private static bool TestFormats(string origin, string destination)
            {
                return new SheetHelper().Converter(origin, destination, SheetName, Separator, Columns, Rows);
            }
        }

        public class TestFormatCases
        {
            public static IEnumerable<TestCaseData> FormatConversions
            {
                get
                {
                    string path = "C:\\Users\\diego\\Desktop\\Tests";

                    return new List<TestCaseData>
                    {
                        // TXT
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.txt", $"{path}\\Converted\\ColunasExcel_TXT.csv").Returns(true).SetName("TXT to CSV"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.txt", $"{path}\\Converted\\ColunasExcel_TXT.txt").Returns(true).SetName("TXT to TXT"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.txt", $"{path}\\Converted\\ColunasExcel_TXT.xls").Returns(true).SetName("TXT to XLS"),

                        // CSV
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.csv", $"{path}\\Converted\\ColunasExcel_CSV.csv").Returns(true).SetName("CSV to CSV"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.csv", $"{path}\\Converted\\ColunasExcel_CSV.txt").Returns(true).SetName("CSV to TXT"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.csv", $"{path}\\Converted\\ColunasExcel_CSV.xls").Returns(true).SetName("CSV to XLS"),

                        // XLS
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xls", $"{path}\\Converted\\ColunasExcel_XLS.csv").Returns(true).SetName("XLS to CSV"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xls", $"{path}\\Converted\\ColunasExcel_XLS.txt").Returns(true).SetName("XLS to TXT"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xls", $"{path}\\Converted\\ColunasExcel_XLS.xls").Returns(true).SetName("XLS to XLS"),

                        // XLSB
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsb", $"{path}\\Converted\\ColunasExcel_XLSB.csv").Returns(true).SetName("XLSB to CSV"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsb", $"{path}\\Converted\\ColunasExcel_XLSB.txt").Returns(true).SetName("XLSB to TXT"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsb", $"{path}\\Converted\\ColunasExcel_XLSB.xls").Returns(true).SetName("XLSB to XLS"),

                        // XLSX
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx", $"{path}\\Converted\\ColunasExcel_XSLX.csv").Returns(true).SetName("XLSX to CSV"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx", $"{path}\\Converted\\ColunasExcel_XSLX.txt").Returns(true).SetName("XLSX to TXT"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx", $"{path}\\Converted\\ColunasExcel_XSLX.html").Returns(true).SetName("XLSX to HTML"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx", $"{path}\\Converted\\ColunasExcel_XSLX.xls").Returns(true).SetName("XLSX to XLS"),

                        // XLSM
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsm", $"{path}\\Converted\\ColunasExcel_XLSM.csv").Returns(true).SetName("XLSM to CSV"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsm", $"{path}\\Converted\\ColunasExcel_XLSM.txt").Returns(true).SetName("XLSM to TXT"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsm", $"{path}\\Converted\\ColunasExcel_XLSM.html").Returns(true).SetName("XLSM to HTML"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsm", $"{path}\\Converted\\ColunasExcel_XLSM.xls").Returns(true).SetName("XLSM to XLS"),

                        // Zip formats
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx.zip", $"{path}\\Converted\\ColunasExcel_XLSX_ZIP.csv").Returns(true).SetName("ZIP to CSV"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx.gz", $"{path}\\Converted\\ColunasExcel_XLSX_GZ.csv").Returns(true).SetName("GZ to CSV"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.gz", $"{path}\\Converted\\ColunasExcel_GZ.csv").Returns(true).SetName("GZ to CSV"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.tar.gz", $"{path}\\Converted\\ColunasExcel_TAR_GZ.csv").Returns(true).SetName("TAR GZ to CSV"),
                        new TestCaseData($"{path}\\Converter\\ColunasExcel.zip", $"{path}\\Converted\\ColunasExcel_ZIP.csv").Returns(true).SetName("ZIP to CSV"),
                    };
                }
            }
        }




    }
}