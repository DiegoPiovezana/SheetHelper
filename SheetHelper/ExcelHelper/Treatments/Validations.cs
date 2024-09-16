using SH.Exceptions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SH.ExcelHelper.Treatments
{

    /// <summary>
    /// Used to process the information provided by the user
    /// </summary>
    internal class Validations
    {
        readonly TryHandlerExceptions _tryHandlerExceptions;
        readonly Definitions _definitions;

        public Validations(SheetHelper sheetHelper)
        {
            _tryHandlerExceptions = new TryHandlerExceptions(sheetHelper);
            _definitions = new Definitions(sheetHelper, this);
        }





        /// <summary>
        /// Checks if it is necessary to convert the file.
        /// </summary>
        /// <returns>True if conversion is required.</returns>
        internal bool CheckConvertNecessary(string origin, string destination, string sheet, string separator, string? columns, string? rows)
        {
            bool checkFormat = Path.GetExtension(origin).Equals(Path.GetExtension(destination), StringComparison.OrdinalIgnoreCase); // The formats is the same?
            bool isOriginTextFormat = Path.GetExtension(origin).Equals(".csv", StringComparison.OrdinalIgnoreCase) || Path.GetExtension(origin).Equals(".txt", StringComparison.OrdinalIgnoreCase);                                                                             // 
            bool isDestinationTextFormat = Path.GetExtension(destination).Equals(".csv", StringComparison.OrdinalIgnoreCase) || Path.GetExtension(destination).Equals(".txt", StringComparison.OrdinalIgnoreCase);
            bool checkFormatText = isOriginTextFormat == isDestinationTextFormat;
            bool checkColumns = string.IsNullOrEmpty(columns) || columns.Trim().Equals(":") || columns.Trim().Equals("1:") || columns.Trim().Equals("A:"); // All columns?
            bool checkRows = string.IsNullOrEmpty(rows) || rows.Trim().Equals(":") || rows.Trim().Equals("1:"); // All rows?

            return !((checkFormat || checkFormatText) && checkColumns && checkRows);
        }

        internal string GetCallingMethodName(int indStack)
        {
            //StackFrame frame = new(indStack);
            //StringBuilder nameMethod = new();

            //for (int i = 1; frame.GetILOffset() != -1 && i <= levelPath; i++)
            //{
            //    nameMethod.Insert(0, "/" + frame.GetMethod().Name);
            //    frame = new StackFrame(indStack + i);
            //}

            //return nameMethod.ToString();           

            StackFrame callerFrame = new StackTrace().GetFrame(indStack); // 1 to get the current method caller
            return callerFrame.GetMethod().Name;

            // TODO: Get the method name after SheetHelper main
        }

        internal void ValidateIntMin(int number, string paramName, string methodName, int min = 0)
        {
            if (number < min)
            {
                throw new ArgumentMinSHException(paramName, methodName, number, min);
            }
        }

        internal void ValidateArgumentNull(object? param, string paramName, string methodName)
        {
            if (param is null)
            {
                throw new ArgumentNullOrEmptySHException(paramName, methodName);
            }
        }

        internal void ValidateStringNullOrEmpty(string? param, string paramName, string methodName)
        {
            if (string.IsNullOrEmpty(param?.Trim()))
            {
                throw new ArgumentNullOrEmptySHException(paramName, methodName);
            }
        }

        internal void ValidateFile(string? pathFile, string paramName, string methodName)
        {
            ValidateStringNullOrEmpty(pathFile, paramName, methodName);

            if (!File.Exists(pathFile))
            {
                throw new FileNotFoundSHException(pathFile);
            }
        }

        internal void ValidateOriginFile(string? pathOrigin, string paramName, string methodName)
        {

            int countOpen = 0;

        again:
            try
            {
                ValidateFile(pathOrigin, paramName, methodName);
                File.OpenRead(pathOrigin).Dispose(); // To check if the file is accessible
            }
            catch (UnauthorizedAccessException ex)
            {
                switch (_tryHandlerExceptions.FileExcelInUse(ex, pathOrigin, countOpen, true))
                {
                    case 0: return;
                    case 1: goto again;
                    default: throw new FileDestinationInUseSHException(pathOrigin);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("E-0000-SH: An error occurred while validating the origin file.", ex);
            }
        }

        internal void ValidateDestinationFile(string? destination, string methodName)
        {
            int countOpen = 0;

        again:
            try
            {
                File.WriteAllText(destination, ""); // To check if the destination file is accessible
                File.Delete(destination); // Delete to prevent the file from being opened during conversion
            }
            catch (UnauthorizedAccessException ex)
            {
                switch (_tryHandlerExceptions.FileExcelInUse(ex, destination, countOpen, false))
                {
                    case 0: return;
                    case 1: goto again;
                    default: throw new FileDestinationInUseSHException(destination);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("E-0000-SH: An error occurred while validating the destination file.", ex);
            }
        }

        internal void ValidateDestinationFolder(string destination, bool createIfNotExist, string paramName, string methodName)
        {
            try
            {
                ValidateStringNullOrEmpty(destination, paramName, methodName);

                if (createIfNotExist) { Directory.CreateDirectory(destination); }
                if (!Directory.Exists(destination)) throw new DirectoryNotFoundSHException("E-0000-SH: Destination folder not found.");
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

        internal void ValidateSheetIdInput(string? sheet)
        {
            // "1" or "Fist_Sheet_Name"

            //if (string.IsNullOrEmpty(sheet))
            //{
            //    throw new ArgumentNullOrEmptySHException(nameof(sheet), GetCallingMethodName(2));
            //}

            if (int.TryParse(sheet, out int sheetNumber) && sheetNumber <= 0)
            {
                throw new ArgumentException("E-0000-SH: The first sheet is '1'!");
            }
        }

        internal void ValidateSheetsDictionaryInput(Dictionary<string, DataTable>? sheetsDictionary)
        {
            if (sheetsDictionary is null || sheetsDictionary.Count == 0)
            {
                throw new ArgumentNullOrEmptySHException(nameof(sheetsDictionary), GetCallingMethodName(2));
            }
        }

        internal void ValidateSeparatorInput(string? separatorInput)
        {
            // ";"

            if (string.IsNullOrEmpty(separatorInput))
            {
                throw new ArgumentException("E-0000-SH: Invalid separator.", nameof(separatorInput));
            }
        }

        internal void ValidateColumnsInput(string? columnsInput)
        {
            // "A:H, 4:9, B, 75, -2"
            // Null Ok

            // TODO: Add specific validation logic for columns
        }

        internal void ValidateRowsInput(string? rowsInput)
        {
            // "1:23, 34:56, 70, 75, -1"


            // throw new ArgumentException("Invalid rows.", nameof(rows));


            // TODO: Add specific validation logic for rows
        }

        internal void ValidateConverter(string? originInput, object? destinationsInput, object? sheetsInput, object? separatorsInput, object? columnsInput, object? rowsInput, string methodName)
        {
            try
            {
                ICollection<string?> destinationsCollection = (ICollection<string?>)destinationsInput;
                ICollection<string?> sheetsCollection = (ICollection<string?>)sheetsInput;
                ICollection<string?> separatorsCollection = (ICollection<string?>)separatorsInput;
                ICollection<string?> columnsCollection = (ICollection<string?>)columnsInput;
                ICollection<string?> rowsCollection = (ICollection<string?>)rowsInput;

                // Remove duplicates  
                HashSet<string?> destinations = new(destinationsCollection);
                HashSet<string?> sheets = new(sheetsCollection);
                HashSet<string?> separators = new(separatorsCollection);
                HashSet<string?> columns = new(columnsCollection);
                HashSet<string?> rows = new(rowsCollection);

                List<Task> validates = new()
                {
                    Task.Run(() => ValidateOriginFile(originInput, nameof(originInput), methodName))
                };

                validates.AddRange(destinations.Select(destination => Task.Run(() => ValidateDestinationFile(destination, methodName))));
                validates.AddRange(sheets.Select(sheet => Task.Run(() => ValidateSheetIdInput(sheet))));
                validates.AddRange(separators.Select(separator => Task.Run(() => ValidateSeparatorInput(separator))));
                validates.AddRange(columns.Select(column => Task.Run(() => ValidateColumnsInput(column))));
                validates.AddRange(rows.Select(row => Task.Run(() => ValidateRowsInput(row))));

                Task.WaitAll(validates.ToArray()); // TODO: Task.WhenAll

                int countConversions = sheetsCollection.Count;
                if (destinationsCollection.Count != countConversions || separatorsCollection.Count != countConversions || columnsCollection.Count != countConversions || rowsCollection.Count != countConversions)
                {
                    //throw new ArgumentException("All parameters must have the same number of items or be single values.");
                    throw new ParamMissingConverterSHException();
                }
            }
            catch (AggregateException ex)
            {
                throw new AggregateException("One or more validations failed.", ex.InnerExceptions);
            }
        }


        internal async Task ValidateOneConverterAsync(string? origin, string? destination, string? sheet, string? separator, string? columns, string? rows, string methodName)
        {
            try
            {
                List<Task> validates = new()
                {
                    Task.Run(() => ValidateOriginFile(origin,nameof(origin), methodName)),
                    Task.Run(() => ValidateDestinationFile(destination, methodName)),
                    Task.Run(() => ValidateSheetIdInput(sheet)),
                    Task.Run(() => ValidateSeparatorInput(separator)),
                    Task.Run(() => ValidateColumnsInput(columns)),
                    Task.Run(() => ValidateRowsInput(rows))
                };

                //Task.WhenAll(validates).Wait();
                await Task.WhenAll(validates.ToArray());
            }
            catch (AggregateException ex)
            {
                throw new AggregateException("One or more validations failed.", ex.InnerExceptions);
            }
        }

        internal void ValidateSaveDataTable(string destination, string separator, string? columns, string? rows, string methodName)
        {
            List<Task> validates = new()
            {
                Task.Run(() => ValidateDestinationFile(destination, methodName)),
                Task.Run(() => ValidateSeparatorInput(separator)),
                Task.Run(() => ValidateColumnsInput(columns)),
                Task.Run(() => ValidateRowsInput(rows))
            };
            Task.WhenAll(validates).Wait();
        }

        //internal static void ValidateFormatFileOrigin(string pathFile, string desiredFormat)
        //{
        //    if (!string.IsNullOrEmpty(pathFile?.Trim()))
        //    {
        //        if (!File.Exists(pathFile)) throw new ParamException(nameof(zipFile), nameof(UnZIP));
        //    }
        //    else
        //    {
        //        throw new ParamException(nameof(zipFile), nameof(UnZIP));
        //    }
        //}               

        internal void ValidateRowsMinDt(DataTable? table, int minRows, string methodName)
        {
            if (table == null)
                throw new ArgumentNullOrEmptySHException(nameof(table), methodName);

            if (table.Rows.Count < minRows - 1)
                throw new RowsMinDtSHException(table.TableName);
        }

        internal void ValidateHeader(DataTable dataTable, bool headerMandatory, int numberColumnsExpected = -1)
        {
            DataRow firstRow = dataTable.Rows[0];
            if (firstRow == null || (numberColumnsExpected > 0 && firstRow.Table.Columns.Count != numberColumnsExpected))
                _tryHandlerExceptions.HeaderInvalid(dataTable);
        }

        internal void ValidateColumnOutOfRange(string idColumn, int limitIndexColumn, int indexColumn, DataTable table)
        {
            if (indexColumn == 0 || indexColumn > limitIndexColumn)
            {
                //switch (_tryHandlerExceptions.ColumnNotExist(idColumn, limitIndexColumn, table))
                //{
                //    case 0: return;
                //    case 1: goto again;
                //    default: throw new ColumnOutRangeSHException(indexColumn, limitIndexColumn);
                //}
            }
        }

        internal void ValidateColumnRefOutOfRange(string idColumn, int limitIndexColumn, int indexColumn)
        {
            //if (limitIndexColumn + indexColumn + 1 > limitIndexColumn)
            //{

            //    fthrow new Exception($"E-4042-SH: The column '{idColumn}' is out of range, because it refers to column '{limitIndexColumn + indexColumn + 1}' (min 1, max {limitIndexColumn})!");
            //}
        }

    }
}
