using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace SH.ExcelHelper
{

    /// <summary>
    /// Used to process the information provided by the user
    /// </summary>
    internal static class Validations
    {
        /// <summary>
        /// Checks if it is necessary to convert the file.
        /// </summary>
        /// <returns>True if conversion is required.</returns>
        internal static bool CheckConvertNecessary(string origin, string destiny, string sheet, string separator, string? columns, string? rows)
        {
            bool checkFormat = Path.GetExtension(origin).Equals(Path.GetExtension(destiny), StringComparison.OrdinalIgnoreCase); // The formats is the same?
            bool isOriginTextFormat = Path.GetExtension(origin).Equals(".csv", StringComparison.OrdinalIgnoreCase) || Path.GetExtension(origin).Equals(".txt", StringComparison.OrdinalIgnoreCase);                                                                             // 
            bool isDestinyTextFormat = Path.GetExtension(destiny).Equals(".csv", StringComparison.OrdinalIgnoreCase) || Path.GetExtension(destiny).Equals(".txt", StringComparison.OrdinalIgnoreCase);
            bool checkFormatText = isOriginTextFormat == isDestinyTextFormat;
            bool checkColumns = string.IsNullOrEmpty(columns) || columns.Trim().Equals(":") || columns.Trim().Equals("1:") || columns.Trim().Equals("A:"); // All columns?
            bool checkRows = string.IsNullOrEmpty(rows) || rows.Trim().Equals(":") || rows.Trim().Equals("1:"); // All rows?

            return !((checkFormat || checkFormatText) && checkColumns && checkRows);
        }

        internal static void ValidateOrigin(string? origin)
        {
            try
            {
                if (string.IsNullOrEmpty(origin)) throw new Exception($"E-0000-SH: The '{nameof(origin)}' is null or empty.");

                if (!File.Exists(origin))
                {
                    throw new FileNotFoundException("E-0000-SH: Origin file not found.", origin);
                }
                else
                {
                    File.OpenRead(origin).Dispose(); // To check if the file is accessible
                }
            }
            catch (UnauthorizedAccessException)
            {
                throw new InvalidOperationException("E-0000-SH: The origin file is in use by another process.");
            }
            catch (Exception ex)
            {
                throw new Exception("E-0000-SH: An error occurred while validating the origin file.", ex);
            }
        }

        internal static void ValidateDestiny(string destiny)
        {
            try
            {
                File.WriteAllText(destiny, ""); // To check if the destination file is accessible
                File.Delete(destiny); // Delete to prevent the file from being opened during conversion
            }
            catch (UnauthorizedAccessException)
            {
                throw new InvalidOperationException("E-0000-SH: The destination file is in use by another process.");
            }
            catch (Exception ex)
            {
                throw new Exception("E-0000-SH: An error occurred while validating the destination file.", ex);
            }
        }

        internal static void ValidateSheet(string sheet)
        {
            // "1" or "Fist_Sheet_Name"

            if (string.IsNullOrEmpty(sheet))
            {
                throw new ArgumentException("E-0000-SH: Invalid sheet name.", nameof(sheet));
            }

            if (int.TryParse(sheet, out int sheetNumber) && sheetNumber <= 0)
            {
                throw new ArgumentException("E-0000-SH: The first sheet is '1'!");
            }
        }

        internal static void ValidateSeparator(string separator)
        {
            // ";"

            if (string.IsNullOrEmpty(separator))
            {
                throw new ArgumentException("E-0000-SH: Invalid separator.", nameof(separator));
            }
        }

        internal static void ValidateColumns(string? columns)
        {
            // "A:H, 4:9, B, 75, -2"

            // TODO: Add specific validation logic for columns
        }

        internal static void ValidateRows(string? rows)
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

        internal static void Validate(string destiny, string separator, string? columns, string? rows)
        {
            List<Task> validates = new()
            {
                Task.Run(() => ValidateDestiny(destiny)),
                Task.Run(() => ValidateSeparator(separator)),
                Task.Run(() => ValidateColumns(columns)),
                Task.Run(() => ValidateRows(rows))
            };

            Task.WhenAll(validates).Wait();
        }

    }
}
