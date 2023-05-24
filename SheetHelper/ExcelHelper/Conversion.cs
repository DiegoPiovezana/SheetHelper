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

            origin = SheetHelper.UnzipAuto(origin, @".\ExcelHelper\Extractions\",false);

            if (!Treatment.CheckConvert(origin, destiny, sheet, separator, columns, rows))
            {
                // If no conversion is needed
                SheetHelper.Progress = 100;
                File.Copy(origin, destiny, true);
                return true;
            }

            DataTable table = SheetHelper.GetDataTable(origin, sheet);

            StringBuilder output = new();

            // Defines the number of all rows to be considered
            int[] rowsNumber = Treatment.DefineRows(rows, table);
            SheetHelper.Progress += 5; // 45                

            // Define in ASCII, which will be all the columns to be converted
            int[] columnsASCII = Treatment.DefineColumnsASCII(columns, table);
            SheetHelper.Progress += 5; // 50 (tratativas ok)

            double countPercPrg = 40.0 / (rowsNumber[1] - rowsNumber[0] + 1); // Percentage to be progressed for each row of the worksheet
            double percPrg = countPercPrg;

            table.Rows.Add(); // To avoid IndexOutOfRangeException (last rows will be ignored)

            //using (StreamWriter writer = new (destiny))
            //{
            // Save all other lines by start and end  
            foreach (int rowIndex in rowsNumber) // For each row in the worksheet
            {
                // Get the row               
                string[] rowFull = table.Rows[rowIndex - 1].ItemArray.Select(cell => cell.ToString()).ToArray();

                if (columnsASCII[0].Equals(0)) // If columns not specified - All
                {
                    output.AppendLine(String.Join(separator, rowFull)); // Add all row columns
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
                    output.AppendLine(String.Join(separator, rowSelected)); // Add the row with the selected columns                           
                    //writer.Write(String.Join(separator, rowSelected));                    
                }

                if (countPercPrg >= 1) // If applicable, load the progress
                {
                    SheetHelper.Progress += (int)countPercPrg; // 90                                                               
                    countPercPrg -= (int)countPercPrg;
                }

                countPercPrg += percPrg; // Increment progress counter                        

                //writer.WriteLine();
            }

            SheetHelper.Progress += (90 - SheetHelper.Progress); // If necessary, complete up to 90%

            // Write new converted file (overwrite if existing)
            File.WriteAllText(destiny, output.ToString());
            //}

            if(Directory.Exists(@".\ExcelHelper\Extractions\")) Directory.Delete(@".\ExcelHelper\Extractions\",true);

            SheetHelper.Progress += 10; // 100
            return true;
        }



    }
}

