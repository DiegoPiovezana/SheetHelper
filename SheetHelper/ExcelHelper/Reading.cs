using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace SheetHelper
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

        public static string[] GetFirstRow(string extension, bool header, DataTable table)
        {
            List<string> row;

            if (!extension.Equals(".csv") && !extension.Equals(".rpt") && !extension.Equals(".txt")) // A tratativa para o cabeçalho csv é diferente
            { // Se não for CSV

                // Se deseja incluir cabeçalho
                if (header)
                {
                    var colunsData = table.Columns.Cast<DataColumn>().ToList(); // Salva cabeçalho
                    row = new List<string>(colunsData.Count);

                    foreach (var item in colunsData) // Realiza a conversão das Listas
                        row.Add(item.ToString());
                }
                else // Se não deseja incluir cabeçalho
                {
                    row = table.Rows[1].ItemArray.Select(f => f.ToString()).ToList(); // linha 2 primeira => index é 1 (-1) e cabeçalho ja retirado (-1)
                }

            }
            else // Se leitura CSV, elimina cabeçalho 'Column' e considera index 0
            {
                ////if (!extension.Equals(".csv"))                       
                //if (_i == table.Rows.Count + 1) // Se automaticamente alterado para última linha
                //    throw new Exception("Para tratar arquivos CSV, TXT ou RPT é necessário informar qual será a última linha!");

                // Realiza a leitura da primeira linha (cabeçalho)
                row = table.Rows[0].ItemArray.Select(f => f.ToString()).ToList(); // linha 2 primeira => index é 1 (-1) e cabeçalho ja retirado (-1)            
            }

            return row.ToArray();
        }

        /// <summary>
        /// Abre o arquivo e realiza a leitura
        /// </summary>       
        internal static DataSet GetDataSet(string origin, string destiny)
        {

        restart:

            // Abre o arquivo
            using (var stream = File.Open(origin, FileMode.Open, FileAccess.Read))
            {
                DataSet result;

                // Realiza a leitura do arquivo
                switch (Path.GetExtension(origin).ToLower())
                {
                    case ".gz":
                        origin = SH.UnGZ(stream, Path.GetDirectoryName(destiny) + "\\");
                        goto restart;

                    case ".zip":
                        stream.Close();
                        origin = SH.UnZIP(origin, Path.GetDirectoryName(destiny));
                        goto restart;

                    case ".rpt":
                    case ".txt":
                    case ".csv":
                        result = ReadCSV(stream);
                        break;

                    default: // .xlsx, .xls, .xlsb, .xlsm
                        result = ReadXLS(stream);
                        break;
                }

                return result;
            }
        }




    }
}
