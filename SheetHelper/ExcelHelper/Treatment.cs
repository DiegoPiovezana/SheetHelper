using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SH
{
    /// <summary>
    /// Used to process the information provided by the user
    /// </summary>
    internal static class Treatment
    {
        #region Validates

        /// <summary>
        /// Checks if it is necessary to convert the file.
        /// </summary>
        /// <returns>True if conversion is required.</returns>
        internal static bool CheckConvert(string origin, string destiny, string sheet, string separator, string? columns, string? rows)
        {
            bool checkFormat = Path.GetExtension(origin).Equals(Path.GetExtension(destiny), StringComparison.OrdinalIgnoreCase); // The formats is the same?
            bool isOriginTextFormat = Path.GetExtension(origin).Equals(".csv", StringComparison.OrdinalIgnoreCase) || Path.GetExtension(origin).Equals(".txt", StringComparison.OrdinalIgnoreCase);                                                                             // 
            bool isDestinyTextFormat = Path.GetExtension(destiny).Equals(".csv", StringComparison.OrdinalIgnoreCase) || Path.GetExtension(destiny).Equals(".txt", StringComparison.OrdinalIgnoreCase);
            bool checkFormatText = isOriginTextFormat == isDestinyTextFormat;
            bool checkColumns = string.IsNullOrEmpty(columns) || columns.Trim().Equals(":") || columns.Trim().Equals("1:") || columns.Trim().Equals("A:"); // All columns?
            bool checkRows = string.IsNullOrEmpty(rows) || rows.Trim().Equals(":") || rows.Trim().Equals("1:"); // All rows?

            return !(checkFormat && checkFormatText && checkColumns && checkRows);
        }

        private static void ValidateOrigin(string origin)
        {
            if (!File.Exists(origin))
            {
                throw new FileNotFoundException("Origin file not found.", origin);
            }
        }

        private static void ValidateDestiny(string destiny)
        {
            try
            {
                File.WriteAllText(destiny, ""); // To check if the destination file is accessible
                File.Delete(destiny); // Delete to prevent the file from being opened during conversion
            }
            catch (Exception ex)
            {
                throw new Exception("Error validating destiny file.", ex);
            }
        }

        private static void ValidateSheet(string sheet)
        {
            // "1" or "Fist_Sheet_Name"

            if (String.IsNullOrEmpty(sheet))
            {
                throw new ArgumentException("Invalid sheet name.", nameof(sheet));
            }
        }

        private static void ValidateSeparator(string separator)
        {
            // ";"

            if (String.IsNullOrEmpty(separator))
            {
                throw new ArgumentException("Invalid separator.", nameof(separator));
            }
        }

        private static void ValidateColumns(string? columns)
        {
            // "A:H, 4:9, B, 75, -2"

            // TODO: Add specific validation logic for columns
        }

        private static void ValidateRows(string? rows)
        {
            // "1:23, 34:56, 70, 75, -1"


            // throw new ArgumentException("Invalid rows.", nameof(rows));


            // TODO: Add specific validation logic for rows
        }

        internal static void Validate(string origin, string destiny, string sheet, string separator, string? columns, string? rows)
        {
            List<Task> validates = new()
            {
                Task.Run(() => ValidateOrigin(origin)),
                Task.Run(() => ValidateDestiny(destiny)),
                Task.Run(() => ValidateSheet(sheet)),
                Task.Run(() => ValidateSeparator(separator)),
                Task.Run(() => ValidateColumns(columns)),
                Task.Run(() => ValidateRows(rows))
            };

            Task.WhenAll(validates).Wait();
        }

        #endregion

        #region Defines

        /// <summary>
        /// Fixes a string containing items by replacing line breaks and semicolons with commas,
        /// removing spaces, single quotes, and double quotes, and ensuring proper comma separation.
        /// </summary>
        /// <param name="items">The string containing items to be fixed.</param>
        /// <returns>The fixed string with proper item separation.</returns>
        public static string FixItems(string items)
        {
            if (!String.IsNullOrEmpty(items))
            {
                items = items.Replace("\n", ",").Replace(";", ","); // Replace line breaks and semicolons with commas
                items = Regex.Replace(items, @"\s+|['""]+", ""); // Remove spaces, single quotes, and double quotes
                items = Regex.Replace(items, ",{2,}", ",").Trim(','); // Remove repeated commas and excess spaces
            }
            return items; // "123123,13514,31234"
        }

        /// <summary>
        /// Receives rows as a string and returns an array of integers with the first and last row.
        /// </summary>
        internal static int[] DefineRows(string rows, DataTable table)
        {
            int limitIndexRows = table.Rows.Count;
            List<int> indexRows = new();

            if (string.IsNullOrEmpty(rows)) // If rows not specified           
                return new[] { 1, limitIndexRows }; // Convert all rows

            rows = FixItems(rows);

            //"1:23,34:-56,23:1,70,75,-1"

            foreach (string row in rows.Split(','))
            {
                if (row.Contains(":")) // "1:23", "34:-56" or "23:1"
                {
                    string[] rowsArray = row.Split(':'); // e.g.: {"A","Z"}

                    if (rowsArray.Length != 2)
                        throw new Exception($"Row '{row}' is not a valid pattern!");

                    if (string.IsNullOrEmpty(rowsArray[0])) // If first row not defined
                        rowsArray[0] = "1"; // Then, convert from the first row

                    if (string.IsNullOrEmpty(rowsArray[1])) // If last row not defined
                        rowsArray[1] = limitIndexRows.ToString(); // Then, convert until the last row

                    int firstRowIndex = ConvertIndexRow(rowsArray[0]);
                    int lastRowIndex = ConvertIndexRow(rowsArray[1]);

                    if (firstRowIndex > lastRowIndex)
                        indexRows.AddRange(Enumerable.Range(firstRowIndex, lastRowIndex - firstRowIndex + 1).Reverse());
                    else
                        indexRows.AddRange(Enumerable.Range(firstRowIndex, lastRowIndex - firstRowIndex + 1));
                }
                else // "70", "75" or "-1"
                {
                    indexRows.Add(ConvertIndexRow(row));
                }
            }
            return indexRows.ToArray();


            int ConvertIndexRow(string row)
            {
                if (row.All(c => char.IsLetter(c))) throw new Exception($"The row '{row}' is not a valid!");

                int indexRow = Convert.ToInt32(row);

                if (indexRow >= 0)  // "75"
                {
                    if (indexRow == 0 || indexRow > limitIndexRows)
                        throw new Exception($"The row '{row}' is out of range (min 1, max {limitIndexRows})!");

                    return indexRow;
                }
                else // "-2"
                {
                    if (limitIndexRows + indexRow + 1 > limitIndexRows)
                        throw new Exception($"The row '{row}' is out of range, because it refers to row '{limitIndexRows + indexRow + 1}' (min 1, max {limitIndexRows})!");

                    return limitIndexRows + indexRow + 1;
                }
            }
        }

        /// <summary>
        /// Defines the index of all columns to be converted.
        /// </summary>
        /// <param name="columns">Columns to be converted. e.g.: "B:H"</param>
        /// <param name="table">DataTable a ser analisado.</param>        
        /// <returns>Array with all indexes of the columns to be converted.</returns>
        /// <exception cref="Exception"></exception>
        internal static int[] DefineColumnsASCII(string columns, DataTable table)
        {
            int limitIndexColumn = table.Columns.Count;
            List<int> indexColumns = new();

            if (string.IsNullOrEmpty(columns)) // If columns not specified                
                return new int[] { 0 }; // Behavior to convert all columns

            columns = FixItems(columns);

            // "A:H,4:9,4:-9,B,75,-2"

            foreach (string column in columns.Split(','))
            {
                if (column.Contains(":")) // "A:H" or "4:9" or 4:-9
                {
                    string[] columnsArray = column.Split(':'); // e.g.: {"A","Z"}

                    if (columnsArray.Length != 2)
                        throw new Exception($"Column '{column}' is not a valid pattern!");

                    if (string.IsNullOrEmpty(columnsArray[0])) // If first row not defined
                        columnsArray[0] = "1"; // Then, convert from the first row

                    if (string.IsNullOrEmpty(columnsArray[1])) // If last row not defined
                        columnsArray[1] = limitIndexColumn.ToString(); // Then, convert until the last row

                    int firstColumnIndex = ConvertIndexColumn(columnsArray[0]);
                    int lastColumnIndex = ConvertIndexColumn(columnsArray[1]);

                    if (lastColumnIndex.Equals(limitIndexColumn) && firstColumnIndex.Equals(1))
                        return new int[] { 0 }; // Behavior to convert all columns

                    if (firstColumnIndex > lastColumnIndex)
                        indexColumns.AddRange(Enumerable.Range(firstColumnIndex, lastColumnIndex - firstColumnIndex + 1).Reverse());
                    else
                        indexColumns.AddRange(Enumerable.Range(firstColumnIndex, lastColumnIndex - firstColumnIndex + 1));
                }
                else // "B", "75" or "-2"
                {
                    indexColumns.Add(ConvertIndexColumn(column));
                }
            }
            return indexColumns.ToArray();


            int ConvertIndexColumn(string column)
            {
                int indexColumn;

                if (column.All(c => char.IsLetter(c))) indexColumn = SheetHelper.GetIndexColumn(column);
                else indexColumn = Convert.ToInt32(column);

                if (indexColumn >= 0)  // "75"
                {
                    if (indexColumn == 0 || indexColumn > limitIndexColumn)
                        throw new Exception($"The column '{column}' is out of range (min 1, max {limitIndexColumn})!");

                    return indexColumn;
                }
                else // "-2"
                {
                    if (limitIndexColumn + indexColumn + 1 > limitIndexColumn)
                        throw new Exception($"The column '{column}' is out of range, because it refers to column '{limitIndexColumn + indexColumn + 1}' (min 1, max {limitIndexColumn})!");

                    return limitIndexColumn + indexColumn + 1;
                }
            }
        }

        #endregion

       


    }
}
