using SH.ExcelHelper.Treatments;
using System.Data;
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

        public Writing(SheetHelper sheetHelper, Validations validations)
        {
            _sheetHelper = sheetHelper;
            _definitions = new Definitions(sheetHelper, validations);
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
                    return TreatCell(column.ColumnName, delimiter);
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

                    return TreatCell(cell.ToString(), delimiter);
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
                            return TreatCell(column.ColumnName, delimiter);
                        }).ToArray();
                    }
                    else
                    {
                        // Get the first row selected (after header - index-2) 
                        rowFull = table.Rows[rowIndex - 2].ItemArray.Select(cell =>
                        {
                            return TreatCell(cell.ToString(), delimiter);
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
        internal string TreatCell(string cellValue, string delimiter)
        {
            //// Header
            //rowFull = table.Columns.Cast<DataColumn>().Select(column =>
            //{
            //    string cellValue = column.ColumnName;
            //    if (cellValue.Contains("\n") || cellValue.Contains("\r")) // Check if the cell contains a line break
            //    {
            //        // Apply double quotes to surround the value and escape the inner double quotes
            //        cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
            //    }
            //    return cellValue;
            //}).ToArray();

            //// Row 1
            //rowFull = table.Rows[rowsNumber[0] - 2].ItemArray.Select(cell =>
            //{
            //    string cellValue = cell.ToString();
            //    if (cellValue.Contains("\n") || cellValue.Contains("\r") || cellValue.Contains(delimiter)) // Check if the cell contains a line break
            //    {
            //        // Apply double quotes to surround the value and escape the inner double quotes
            //        cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
            //    }
            //    return cellValue;
            //}).ToArray();

            //// Other rows
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

            // Generic
            //return cells.Select(cell =>
            //{
            //    string cellValue = cell.ToString();
            //    if (cellValue.Contains("\n") || cellValue.Contains("\r") || cellValue.Contains(delimiter))
            //    {
            //        cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
            //    }
            //    return (T)Convert.ChangeType(cellValue, typeof(T));
            //}).ToArray();

            if (cellValue.Contains("\n") 
                || cellValue.Contains("\r")                
                || cellValue.Contains(delimiter) 
                || cellValue.Contains("\""))
            {
                cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\""; // Apply ""
            }

            if (_sheetHelper.ProhibitedItems != null && _sheetHelper.ProhibitedItems.Count > 0)
            {
                foreach (var item in _sheetHelper.ProhibitedItems)
                {
                    cellValue = cellValue.Replace(item.Key, item.Value);
                }
            }

            return cellValue;
        }

    }

}

