using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace SH
{
    internal class Reading
    {
        /// <summary>
        /// Reads .xls, .xlsx and .xlsb files
        /// </summary>
        internal static DataSet ReadXLS(FileStream stream)
        {
            // Autodetect format, supports:
            // - Binary Excel files (2.0-2003 format; *.xls)
            // - Excel OpenXml files (2007 format; *.xlsx, *.xlsb)
            using var reader = ExcelReaderFactory.CreateReader(stream);
            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }

            });

            return result;
        }

        /// <summary>
        /// Reads .csv files
        /// </summary>
        internal static DataSet ReadCSV(FileStream stream)
        {           
            using var reader = ExcelReaderFactory.CreateCsvReader(stream);
            return reader.AsDataSet();            
        }

        /// <summary>
        /// Get the desired sheet
        /// </summary>
        /// <param name="sheet">Nome ou índice da aba desejada</param>
        /// <param name="result">Dataset da planilha</param>    
        /// <exception cref="Exception">Erro ao localizar aba</exception>
        internal static DataTable GetTableByDataSet(string sheet, DataSet result)
        {
            if (int.TryParse(sheet, out int sh)) // Se informado índice da aba desejada
            {
                // Se existir abas na planilha e a desejada estiver correta
                if (result.Tables.Count <= 0 || sh <= -1 || sh > result.Tables.Count)
                {
                    throw new Exception("Erro ao selecionar a aba desejada! Verifique se o índice da aba está correto.");
                }

                return result.Tables[sh - 1]; // Obtem a aba desejada

            } // Se nome da aba for informado
            else
            {
                if (!result.Tables.Contains(sheet)) // Se nome da aba não localizado
                {
                    throw new Exception($"Não foi possível encontrar a aba '{sheet}' desejada! Verifique se o nome da aba está correto.");
                }

                return result.Tables[sheet]; // Obtem a aba desejada
            }
        }

        /// <summary>
        /// Retrieves the first row of a DataTable based on the file extension and header settings.
        /// </summary>
        /// <param name="extension">The file extension to determine the reading logic.</param>
        /// <param name="table">The DataTable containing the data.</param>
        /// <param name="ignoreCSV">If true and if the extension is a CSV or similar, it will return null.</param>
        /// <returns>An array of strings representing the first row of the DataTable.</returns>
        /// <exception cref="Exception">Thrown when header is required for CSV, TXT, or RPT files.</exception>
        internal static string[] GetFirstRow(string extension, DataTable table, bool ignoreCSV = false)
        {
            if (!IsCsvTxtRptExtension(extension))
            {
                return table.Columns.Cast<DataColumn>()
                    .Select(column => column.ColumnName)
                    .ToArray();
            }
            else // If CSV the first row of the table is indeed the first row
            {
                if(ignoreCSV) { return null; }

                return table.Rows[0].ItemArray
                        .Select(item => item.ToString())
                        .ToArray();
            }

            static bool IsCsvTxtRptExtension(string extension)
            {
                string[] allowedExtensions = { ".csv", ".txt", ".rpt" };
                return allowedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// Open the file and perform the reading
        /// </summary>       
        internal static DataSet GetDataSet(string origin)
        {
            using var stream = File.Open(origin, FileMode.Open, FileAccess.Read);
            return Path.GetExtension(origin).ToLower() switch
            {
                ".rpt" or ".txt" or ".csv" => ReadCSV(stream),
                _ => ReadXLS(stream), // .xlsx, .xls, .xlsb, .xlsm
            };
        }




    }
}
