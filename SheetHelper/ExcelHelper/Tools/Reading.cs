using ExcelDataReader;
using SH.Exceptions;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace SH.ExcelHelper.Tools
{
    public class SheetReader : IDisposable
    {
        public Stream Stream { get; }
        public IExcelDataReader Reader { get; }

        public SheetReader(Stream stream, IExcelDataReader reader)
        {
            Stream = stream;
            Reader = reader;
        }

        public void Dispose()
        {
            Reader?.Dispose();
            Stream?.Dispose();
        }
    }

    internal class Reading
    {
        internal SheetReader ReadSheet(string filePath, string sheet)
        {
            FileStream? stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            IExcelDataReader? reader = Path.GetExtension(filePath).ToLower() switch
            {
                ".rpt" or ".txt" or ".csv" => ExcelReaderFactory.CreateCsvReader(stream),
                _ => ExcelReaderFactory.CreateReader(stream) // .xlsx, .xls, .xlsb, .xlsm
            };

            int sheetIndex = -1; // To track the index of the current sheet
            do
            {
                sheetIndex++;

                // Check if the sheet is specified by index or name
                if (string.Equals(sheet, (sheetIndex + 1).ToString(), StringComparison.OrdinalIgnoreCase) || // Compare with index as string
                    string.Equals(reader.Name, sheet, StringComparison.OrdinalIgnoreCase)) // Compare with sheet name
                {
                    Debug.WriteLine($"Reading sheet: {reader.Name}");
                    //while (reader.Read()) // Cells will be read as string
                    //{
                    //    //for (int i = 0; i < reader.FieldCount; i++) // FieldCount is the number of columns
                    //    //{
                    //    //    Debug.Write($"{reader.GetValue(i)} ");
                    //    //}
                    //    //Debug.WriteLine("");
                    //}
                    //break;

                    return new SheetReader(stream, reader);
                }
            } while (reader.NextResult());

            throw new ArgumentException($"Sheet '{sheet}' not found.");
        }


        /// <summary>
        /// Reads .xls, .xlsx and .xlsb files
        /// </summary>
        internal DataSet ReadXLS(FileStream stream)
        {
            // Autodetect format, supports:
            // - Binary Excel files (2.0-2003 format; *.xls)
            // - Excel OpenXml files (2007 format; *.xlsx, *.xlsb)
            using var reader = ExcelReaderFactory.CreateReader(stream);
            IDataReader dataReader = reader;

            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) =>
                new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true,
                    EmptyColumnNamePrefix = "EmptyColumn"
                }

            });

            return result;
        }

        /// <summary>
        /// Reads .csv files.
        /// </summary>
        internal DataSet ReadCSV(FileStream stream)
        {
            using var reader = ExcelReaderFactory.CreateCsvReader(stream);

            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) =>
                new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true,
                    EmptyColumnNamePrefix = "EmptyColumn"
                }

            });

            return result;
        }

        /// <summary>
        /// Get the desired sheet from the dataset.
        /// </summary>
        /// <param name="sheet">Name or index of the desired sheet.</param>
        /// <param name="result">Dataset of the spreadsheet.</param>
        /// <returns>The DataTable representing the desired sheet.</returns>
        /// <exception cref="Exception">Throws an exception if there is an error locating the sheet.</exception>
        internal DataTable GetTableByDataSet(string sheet, DataSet result)
        {
            if (int.TryParse(sheet, out int sheetIndex)) // If the index of the desired sheet is provided
            {
                // If there are no sheets in the dataset or the provided index is incorrect
                if (result.Tables.Count <= 0 || sheetIndex <= -1 || sheetIndex > result.Tables.Count)
                {
                    throw new Exception($"E-0000-SH: Error selecting the desired sheet! Please check if the sheet index '{sheetIndex}' is correct.");
                }

                return result.Tables[sheetIndex - 1];

            }
            else // If sheet name is provided
            {
                if (!result.Tables.Contains(sheet)) // If sheet name is not found
                {
                    throw new Exception($"E-0000-SH: Unable to find the desired sheet '{sheet}'! Please check if the sheet name is correct.");
                }

                //return result.Tables[sheet];
                // TODO: ?SheetHelper.NormalizeText(table.TableName)
                return result.Tables
                    .Cast<DataTable>()
                    .FirstOrDefault(table => table.TableName.Trim().ToLower() == sheet.Trim().ToLower());
            }
        }


        /// <summary>
        /// Open the file and perform the reading
        /// </summary>       
        internal DataSet GetDataSet(string origin)
        {
            try
            {
                using var stream = File.Open(origin, FileMode.Open, FileAccess.Read);
                return Path.GetExtension(origin).ToLower() switch
                {
                    ".rpt" or ".txt" or ".csv" => ReadCSV(stream),
                    _ => ReadXLS(stream), // .xlsx, .xls, .xlsb, .xlsm
                };
            }
            catch (ExcelDataReader.Exceptions.HeaderException heEx) when (heEx.HResult.Equals(-2147024894) || heEx.HResult.Equals(-2146233088))
            {
                //throw new Exception($"Erro E-99101-SH: Sem suporte para converter o arquivo '{Path.GetExtension(origin)}'.");
                throw new FileOriginNotReadSupportSHException(origin, heEx);
            }
        }







    }
}
