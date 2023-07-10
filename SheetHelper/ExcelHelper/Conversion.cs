using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace SH
{
    /// <summary>
    /// Fast and lightweight library for easy conversion of large Excel files
    /// </summary>
    internal static class Conversion
    {
        internal static bool Converter(string origin, string destiny, string sheet, string separator, string? columns, string? rows)
        {
            SheetHelper.Progress = 0;

            Treatment.Validate(origin, destiny, sheet, separator, columns, rows);
            SheetHelper.Progress += 5; // 5 

            origin = SheetHelper.UnzipAuto(origin, @".\SheetHelper\Extractions\", false);

            if (!Treatment.CheckConvert(origin, destiny, sheet, separator, columns, rows))
            {
                // If no conversion is needed
                SheetHelper.Progress = 100;
                File.Copy(origin, destiny, true);
                return true;
            }

            DataTable table = SheetHelper.GetDataTable(origin, sheet);

            return ConverterDataTable(table, destiny, separator, columns, rows);
        }

        internal static bool ConverterDataTable(DataTable table, string destiny, string separator, string? columns, string? rows)
        {
            StringBuilder output = new();
            string[] rowFull;

            // Defines the number of all rows to be considered
            int[] rowsNumber = Treatment.DefineRows(rows ?? "", table);
            SheetHelper.Progress += 5; // 45                

            // Define in ASCII, which will be all the columns to be converted
            int[] columnsASCII = Treatment.DefineColumnsASCII(columns ?? "", table);
            SheetHelper.Progress += 5; // 50 (tratativas ok)

            double countPercPrg = 40.0 / rowsNumber.Count(); // Percentage to be progressed for each row of the worksheet
            double percPrg = countPercPrg;

            //table.Rows.Add(); // To avoid IndexOutOfRangeException (last rows will be ignored)

            //using (StreamWriter writer = new (destiny))
            //{

            // If you want to include header
            if (rowsNumber[0].Equals(1))
            {
                // Get the header (coluns name)
                rowFull = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
            }
            else
            {
                // Get the first row selected (after header)              
                rowFull = table.Rows[rowsNumber[0]].ItemArray.Select(cell => cell.ToString()).ToArray();
            }

            // Save all other rows by start and end  
            foreach (int rowIndex in rowsNumber.Skip(1).Concat(new[] { rowsNumber.Last() })) // For each row in the worksheet
            {         
                if (columnsASCII[0].Equals(0)) // If columns not specified - All
                {
                    output.AppendLine(string.Join(separator, rowFull)); // Add all row columns
                    //writer.Write(String.Join(separator, rowFull));                       
                }
                else // If specified columns - Selected
                {
                    StringBuilder rowSelected = new(); // Store the selected columns of the row                           

                    foreach (int column in columnsASCII) // For each column of rows
                    {
                        // Select column considering ASCII table and add separately                            
                        rowSelected.Append(rowFull[column - 1]).Append(separator);
                    }
                    output.AppendLine(string.Join(separator, rowSelected)); // Add the row with the selected columns                           
                    //writer.Write(String.Join(separator, rowSelected));                    
                }

                if (countPercPrg >= 1) // If applicable, load the progress
                {
                    SheetHelper.Progress += (int)countPercPrg; // 90                                                               
                    countPercPrg -= (int)countPercPrg;
                }

                countPercPrg += percPrg; // Increment progress counter                      

                // Get the next row               
                rowFull = table.Rows[rowIndex - 2].ItemArray.Select(cell => cell.ToString()).ToArray();

                //writer.WriteLine();
            }

            SheetHelper.Progress += (90 - SheetHelper.Progress); // If necessary, complete up to 90%

            // Write new converted file (overwrite if existing)
            File.WriteAllText(destiny, output.ToString());
            //}

            if (Directory.Exists(@".\SheetHelper\Extractions\")) Directory.Delete(@".\SheetHelper\Extractions\", true);

            SheetHelper.Progress += 10; // 100
            return true;
        }



    }
}

