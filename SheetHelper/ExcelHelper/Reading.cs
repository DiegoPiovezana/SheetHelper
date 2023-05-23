using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace SH
{
    internal class Reading
    {
        /// <summary>
        /// Realiza a leitura de arquivos .xls, .xlsx e .xlsb
        /// </summary>
        internal static DataSet ReadXLS(FileStream stream)
        {
            // Formato de detecção automática, suporta: 
            //   - Arquivos Excel binários (formato 2.0-2003; *.xls) 
            //   - Arquivos Excel OpenXml (formato 2007; *.xlsx, *.xlsb)
            using var reader = ExcelReaderFactory.CreateReader(stream);
            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }

            });

            return result;
        }

        /// <summary>
        /// Realiza a leitura de arquivos .csv
        /// </summary>
        internal static DataSet ReadCSV(FileStream stream)
        {
            // Realiza a leitura do arquivo Excel CSV
            using var reader = ExcelReaderFactory.CreateCsvReader(stream);
            DataSet result = reader.AsDataSet();
            return result;
        }


        /// <summary>
        /// Obtem a aba desejada
        /// </summary>
        /// <param name="sheet">Nome ou índice da aba desejada</param>
        /// <param name="result">Dataset da planilha</param>    
        /// <exception cref="Exception">Erro ao localizar aba</exception>
        internal static DataTable GetTable(string sheet, DataSet result)
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

        //public static string[] GetFirstRow(string extension, DataTable table, bool header)
        //{
        //    List<string> row;

        //    if (!extension.Equals(".csv") && !extension.Equals(".rpt") && !extension.Equals(".txt")) // A tratativa para o cabeçalho csv é diferente
        //    { // Se não for CSV

        //        // Se deseja incluir cabeçalho
        //        if (header)
        //        {
        //            var colunsData = table.Columns.Cast<DataColumn>().ToList(); // Salva cabeçalho
        //            row = new List<string>(colunsData.Count);

        //            foreach (var item in colunsData) // Realiza a conversão das Listas
        //                row.Add(item.ToString());
        //        }
        //        else // Se não deseja incluir cabeçalho
        //        {
        //            row = table.Rows[1].ItemArray.Select(f => f.ToString()).ToList(); // linha 2 primeira => index é 1 (-1) e cabeçalho ja retirado (-1)
        //        }

        //    }
        //    else // Se leitura CSV, elimina cabeçalho 'Column' e considera index 0
        //    {
        //        ////if (!extension.Equals(".csv"))                       
        //        //if (_i == table.Rows.Count + 1) // Se automaticamente alterado para última linha
        //        //    throw new Exception("Para tratar arquivos CSV, TXT ou RPT é necessário informar qual será a última linha!");

        //        // Realiza a leitura da primeira linha (cabeçalho)
        //        row = table.Rows[0].ItemArray.Select(f => f.ToString()).ToList(); // linha 2 primeira => index é 1 (-1) e cabeçalho ja retirado (-1)            
        //    }

        //    return row.ToArray();
        //}

        /// <summary>
        /// Retrieves the first row of a DataTable based on the file extension and header settings.
        /// </summary>
        /// <param name="extension">The file extension to determine the reading logic.</param>
        /// <param name="table">The DataTable containing the data.</param>
        /// <param name="header">Specifies whether to include the header row.</param>
        /// <returns>An array of strings representing the first row of the DataTable.</returns>
        /// <exception cref="Exception">Thrown when header is required for CSV, TXT, or RPT files.</exception>
        public static string[] GetFirstRow(string extension, DataTable table, bool header = true)
        {
            List<string> row;

            if (!IsCsvTxtRptExtension(extension))
            {
                if (header)
                {
                    row = table.Columns.Cast<DataColumn>()
                        .Select(column => column.ColumnName)
                        .ToList();
                }
                else
                {
                    row = table.Rows[0].ItemArray
                        .Select(item => item.ToString())
                        .ToList();
                }
            }
            else
            {
                row = table.Rows[1].ItemArray
                        .Select(item => item.ToString())
                        .ToList();


                //if (header)
                //{
                //    row = table.Rows[1].ItemArray
                //        .Select(item => item.ToString())
                //        .ToList();
                //}
                //else
                //{
                //    throw new Exception("For CSV, TXT, or RPT files, a header row is required.");
                //}
            }

            return row.ToArray();

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
            using (var stream = File.Open(origin, FileMode.Open, FileAccess.Read))
            {
                switch (Path.GetExtension(origin).ToLower())
                {
                    case ".rpt":
                    case ".txt":
                    case ".csv":
                        return ReadCSV(stream);

                    default: // .xlsx, .xls, .xlsb, .xlsm
                        return ReadXLS(stream);
                }
            }
        }




    }
}
