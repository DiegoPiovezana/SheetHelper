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

            //do
            //{

            //    reader.NextResult();


            //    while (reader.Read())
            //    {
            //        // reader.GetDouble(0);
            //    }
            //} while (reader.NextResult());

            //return null;

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
        /// <param name="sheet">Name or index of the desired sheet</param>
        /// <param name="result">Dataset of the spreadsheet</param>
        /// <exception cref="Exception">Error locating sheet</exception>
        internal static DataTable GetTableByDataSet(string sheet, DataSet result)
        {
            if (int.TryParse(sheet, out int sh)) // If the index of the desired tab is informed
            {
                // If there are sheets in the worksheet and the desired one is correct
                if (result.Tables.Count <= 0 || sh <= -1 || sh > result.Tables.Count)
                {
                    throw new Exception("Erro ao selecionar a aba desejada! Verifique se o índice da aba está correto.");
                }

                return result.Tables[sh - 1]; // Get the desired tab

            } // If sheet name is given
            else
            {
                if (!result.Tables.Contains(sheet)) // If sheet name not found
                {
                    throw new Exception($"Não foi possível encontrar a aba '{sheet}' desejada! Verifique se o nome da aba está correto.");
                }

                //return result.Tables[sheet];
                // TODO: ?SheetHelper.NormalizeText(table.TableName)
                return result.Tables.Cast<DataTable>().FirstOrDefault(table => table.TableName.Trim().ToLower() == sheet.Trim().ToLower()); // Obtem a aba desejada
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

        private static bool IsCsvTxtRptExtension(string extension)
        {
            string[] allowedExtensions = { ".csv", ".txt", ".rpt" };
            return allowedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase);
        }

        internal static DataTable TreatHeader(DataTable dataTable, string extension)
        {
            if (IsCsvTxtRptExtension(extension))
            {
                DataRow firstRow = dataTable.Rows[0];

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    dataTable.Columns[i].ColumnName = firstRow[i].ToString();
                }

                dataTable.Rows.RemoveAt(0);
            }

            return dataTable;
        }




    }
}
