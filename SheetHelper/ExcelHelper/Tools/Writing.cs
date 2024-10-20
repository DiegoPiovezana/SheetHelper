using SH.ExcelHelper.Treatments;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace SH.ExcelHelper.Tools
{
    /// <summary>
    /// Fast and lightweight library for easy conversion of large Excel files
    /// </summary>
    internal class Writing
    {
        private readonly SheetHelper _sheetHelper;
        private readonly Definitions _definitions;
        private readonly Adjustments _adjustments;
        private readonly Features _features;

        public Writing(SheetHelper sheetHelper, Validations validations, Features features)
        {
            _sheetHelper = sheetHelper;
            _definitions = new Definitions(sheetHelper, validations);
            _adjustments = new Adjustments(sheetHelper);
            _features = features;
        }




        //internal static bool Converter(string origin, string destination, string sheet, string delimiter, string? columns, string? rows)
        //{
        //    SheetHelper.Progress = 0;

        //    Validations.Validate(origin, destination, sheet, delimiter, columns, rows);
        //    SheetHelper.Progress += 5; // 5 

        //    origin = SheetHelper.UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
        //    if (origin == null) return false;

        //    if (!Validations.CheckConvertNecessary(origin, destination, sheet, delimiter, columns, rows))
        //    {
        //        // If no conversion is needed
        //        SheetHelper.Progress = 100;
        //        File.Copy(origin, destination, true);
        //        return true;
        //    }

        //    DataTable table = SheetHelper.GetDataTable(origin, sheet);

        //    return ConverterDataTable(table, destination, delimiter, columns, rows);
        //}

        internal bool SaveDataTable(DataTable table, string destination, string delimiter, string? columns, string? rows)
        {

            StringBuilder output = new();
            string[] rowFull;

            // Defines the number of all rows to be considered
            int[] rowsNumber = _definitions.DefineRows(rows ?? "", table);
            _sheetHelper.Progress += 5; // 45                

            // Define in ASCII, which will be all the columns to be converted
            int[] columnsASCII = _definitions.DefineColumnsASCII(columns ?? "", table);
            _sheetHelper.Progress += 5; // 50 (tratativas ok)

            double countPercPrg = 40.0 / rowsNumber.Count(); // Percentage to be progressed for each row of the worksheet
            double percPrg = countPercPrg;

            //table.Rows.Add(); // To avoid IndexOutOfRangeException (last rows will be ignored)

            //using (StreamWriter writer = new (destination))
            //{

            // If you want to include header
            if (rowsNumber[0].Equals(1))
            {
                // Get the header (coluns name)
                //rowFull = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
                rowFull = table.Columns.Cast<DataColumn>().Select(column =>
                {
                    //string cellValue = column.ColumnName;
                    //if (cellValue.Contains("\n") || cellValue.Contains("\r")) // Check if the cell contains a line break
                    //{
                    //    // Apply double quotes to surround the value and escape the inner double quotes
                    //    cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
                    //}
                    //return cellValue;
                    return _adjustments.TreatCell(column.ColumnName, delimiter);
                }).ToArray();
            }
            else
            {
                // Get the first row selected (after header - index-2)              
                //rowFull = table.Rows[rowsNumber[0]].ItemArray.Select(cell => cell.ToString()).ToArray();
                rowFull = table.Rows[rowsNumber[0] - 2].ItemArray.Select(cell =>
                {
                    //string cellValue = cell.ToString();
                    //if (cellValue.Contains("\n") || cellValue.Contains("\r") || cellValue.Contains(delimiter)) // Check if the cell contains a line break
                    //{
                    //    // Apply double quotes to surround the value and escape the inner double quotes
                    //    cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
                    //}
                    //return cellValue;

                    return _adjustments.TreatCell(cell.ToString(), delimiter);
                }).ToArray();
            }

            // Save all rows by start and end  
            foreach (int rowIndex in rowsNumber.Skip(1).Concat(new[] { rowsNumber.Last() })) // For each row in the worksheet
            {
                if (columnsASCII[0].Equals(0)) // If columns not specified - All
                {
                    output.AppendLine(string.Join(delimiter, rowFull)); // Add all row columns
                                                                        //writer.Write(String.Join(delimiter, rowFull));                       
                }
                else // If specified columns - Selected
                {
                    StringBuilder rowSelected = new(); // Store the selected columns of the row                           

                    foreach (int column in columnsASCII) // For each column of rows
                    {
                        // Select column considering ASCII table and add separately                            
                        rowSelected.Append(rowFull[column - 1]).Append(delimiter);
                    }
                    output.AppendLine(string.Join(delimiter, rowSelected)); // Add the row with the selected columns                           
                                                                            //writer.Write(String.Join(delimiter, rowSelected));                    
                }

                if (countPercPrg >= 1) // If applicable, load the progress
                {
                    _sheetHelper.Progress += (int)countPercPrg; // 90                                                               
                    countPercPrg -= (int)countPercPrg;
                }

                countPercPrg += percPrg; // Increment progress counter                      

                // Get the next row
                if (rowIndex - 1 >= 0 && rowIndex - 2 < table.Rows.Count)
                {
                    //rowFull = table.Rows[rowIndex - 2].ItemArray.Select(cell => cell.ToString()).ToArray();
                    //rowFull = table.Rows[rowIndex - 2].ItemArray.Select(cell =>
                    //{
                    //    string cellValue = cell.ToString();
                    //    if (cellValue.Contains("\n") || cellValue.Contains("\r")) // Check if the cell contains a line break
                    //    {
                    //        // Apply double quotes to surround the value and escape the inner double quotes
                    //        cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
                    //    }
                    //    return cellValue;
                    //}).ToArray();
                    //rowFull = TreatCell(table.Rows[rowIndex - 2].ItemArray.Select(cell => cell.ToString()).ToArray());
                    //rowFull = table.Rows[rowIndex - 2].ItemArray.Select(cell =>
                    //{
                    //    return TreatCell(cell.ToString(), delimiter);

                    //}).ToArray();

                    if (rowIndex.Equals(1))  // If header
                    {
                        // Get the header (coluns name)                       
                        rowFull = table.Columns.Cast<DataColumn>().Select(column =>
                        {
                            return _adjustments.TreatCell(column.ColumnName, delimiter);
                        }).ToArray();
                    }
                    else
                    {
                        // Get the first row selected (after header - index-2) 
                        rowFull = table.Rows[rowIndex - 2].ItemArray.Select(cell =>
                        {
                            return _adjustments.TreatCell(cell.ToString(), delimiter);
                        }).ToArray();
                    }
                }

                //writer.WriteLine();
            }

            _sheetHelper.Progress += 90 - _sheetHelper.Progress; // If necessary, complete up to 90%

            // Write new converted file (overwrite if existing)
            //File.WriteAllText(destination, output.ToString(), Encoding.UTF8);
            using (StreamWriter writer = new(destination, false, Encoding.UTF8)) { writer.Write(output.ToString()); }

            //if (Directory.Exists(@".\SheetHelper\")) Directory.Delete(@".\SheetHelper\", true);

            _sheetHelper.Progress += 10; // 100
            return File.Exists(destination);
        }

        //internal static T[] TreatCell<T>(T[] cells, string delimiter = ";")
        internal string GenerateCsv(string fileName, int numRows, int numColumns, string delimiter)
        {
            try
            {
                // Header with column names (for row 1)
                var columnNames = Enumerable.Range(1, numColumns)
                                            .Select(i => _features.GetNameColumn(i))
                                            .ToArray();

                // Buffer size for efficient writing
                const int bufferSize = 65536; // 64KB buffer

                using (var fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize))
                using (var streamWriter = new StreamWriter(fileStream, Encoding.UTF8, bufferSize))
                {
                    // Write the header to A1, B1, C1, etc.
                    streamWriter.WriteLine(string.Join(delimiter, columnNames.Select((col, index) => $"{col}1")));

                    // Process in chunks to manage memory
                    const int chunkSize = 100_000;
                    int numChunks = (numRows + chunkSize - 2) / chunkSize;

                    for (int chunk = 0; chunk < numChunks; chunk++)
                    {
                        // Start data at row 2 (header is row 1)
                        int startRow = chunk * chunkSize + 2;
                        int rowsInChunk = Math.Min(chunkSize, (numRows - 1) - chunk * chunkSize);

                        // Generate lines for the current chunk
                        for (int row = startRow; row < startRow + rowsInChunk; row++)
                        {
                            var cellReferences = columnNames.Select(col => $"{col}{row}");
                            streamWriter.WriteLine(string.Join(delimiter, cellReferences));
                        }

                        Debug.WriteLine($"Processed chunk {chunk + 1}/{numChunks} (rows {startRow} to {startRow + rowsInChunk - 1})");
                    }
                }

                Debug.WriteLine("CSV file generated successfully!");
                Debug.WriteLine($"File path: {Path.GetFullPath(fileName)}");

                return Path.GetFullPath(fileName);
            }
            catch (IOException ioEx)
            {
                Debug.WriteLine($"File I/O error: {ioEx.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error generating CSV file: {ex.Message}");
                throw;
            }
        }



    }
}