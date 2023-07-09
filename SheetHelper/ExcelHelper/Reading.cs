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

               //return result.Tables[sheet];
                return result.Tables.Cast<DataTable>().FirstOrDefault(table => table.TableName.ToLower() == sheet.ToLower()); // Obtem a aba desejada
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
            //string[]? header = GetFirstRow(Path.GetExtension(origin), dataTable, true);
            //if (header != null) { dataTable.Rows.InsertAt(SheetHelper.ConvertToDataRow(header, dataTable), 0); }

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
