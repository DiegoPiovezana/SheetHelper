using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Compression;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;
using ExcelDataReader.Core;
using System.Text;

namespace SH
{
    /// <summary>
    /// Fast and lightweight library for easy conversion of large Excel files
    /// </summary>
    public class SheetHelper
    {
        /// <summary>
        /// Represents the conversion progress. E.g.: If 100%, the conversion is fully completed.
        /// </summary>
        public static int Progress { get; set; }


        /// <summary>
        /// Terminates all Excel processes
        /// </summary>
        public static void CloseExcel()
        {
            try
            {
                var processes = from p in Process.GetProcessesByName("EXCEL") select p;
                foreach (var process in processes) process.Kill();
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Receives the column name and returns the index in the worksheet
        /// </summary>
        /// <param name="columnName">Column name. E.g.: "A"</param>
        /// <returns>Index. E.g.: "A" = 1</returns>
        public static int GetIndexColumn(string columnName)
        {
            try
            {
                int sum = 0;

                foreach (var character in columnName)
                {
                    sum *= 26;
                    sum += (character - 'A' + 1);
                }

                return sum; // E.g.: A = 1, Z = 26, AA = 27
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Gets the column name by index
        /// </summary>
        /// <param name="columnIndex"> Column index</param>
        /// <returns>Column name (e.g.: "AB")</returns>
        public static string GetNameColumn(int columnIndex)
        {
            try
            {
                string columnName = String.Empty;

                while (columnIndex > 0)
                {
                    int remainder = (columnIndex - 1) % 26;
                    columnName = Convert.ToChar('A' + remainder) + columnName;
                    columnIndex = (columnIndex - remainder) / 26;
                }

                return columnName;
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Descompacta arquivo .GZ
        /// </summary>
        /// <param name="compressedFileStream">Arquivo  a ser convertido obtido através do método File.Open. </param> 
        /// <param name="pathDestiny">Diretório onde será salvo o arquivo descompactado (contendo OU NAO o nome do arquivo destino). E.g.: 'C:\\Arquivos\\ ou 'C:\\Arquivos\\Convertido.xlsx'</param>
        public static string? UnGZ(FileStream compressedFileStream, string pathDestiny)
        {
            try
            {
                string fileConverted;

                if (Path.GetExtension(pathDestiny) == "") // Se formato a ser convertido não especificado, tenta obter do nome
                {
                    string originalFileName = Path.GetFileName(compressedFileStream.Name).Replace(".gz", "").Replace(".GZ", "");
                    string formatOriginal = Regex.Match(Path.GetExtension(originalFileName), @"\.[A-Za-z]*").Value;
                    fileConverted = $"{pathDestiny}{Path.GetFileNameWithoutExtension(originalFileName)}{formatOriginal}";
                }
                else
                {
                    fileConverted = pathDestiny;
                }

                //FileStream compressedFileStream = File.Open(compressedFileName, FileMode.Open); // "compressed.xlsx.gz"
                FileStream outputFileStream = File.Create(fileConverted); // "decompressed.xlsx"
                var decompressor = new GZipStream(compressedFileStream, CompressionMode.Decompress);
                decompressor.CopyTo(outputFileStream);

                // Encerra uso dos arquivos
                compressedFileStream.Close();
                outputFileStream.Close();

                return File.Exists(fileConverted) ? fileConverted : null;
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Extracts a .ZIP file.
        /// </summary>
        /// <param name="zipFile">The location and name of the compressed file. E.g.: 'C:\\Files\\Report.zip'</param>
        /// <param name="pathDestiny">The directory where the extracted file will be saved. E.g.: 'C:\\Files\\'</param>
        /// <returns>The path of the extracted file.</returns>
        public static string UnZIP(string zipFile, string pathDestiny)
        {
            try
            {
                string directoryZIP = $"{pathDestiny}\\CnvrtdZIP\\";

                // Extract to a new directory
                ZipFile.ExtractToDirectory(zipFile, directoryZIP);

                IEnumerable<string> files = Directory.EnumerateFiles(directoryZIP);
                string fileLocation = files.First(); // Get the location of the file
                string fileDestiny = $"{pathDestiny}\\{Path.GetFileName(fileLocation)}"; // Destination location of the file

                if (File.Exists(fileDestiny)) // If the file already exists, delete it
                    File.Delete(fileDestiny);

                File.Move(fileLocation, fileDestiny); // Move it to the target location            
                Directory.Delete(directoryZIP); // Delete the previously created directory

                return $"{pathDestiny}\\{Path.GetFileName(fileLocation)}";
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Converts a string array to a DataRow and returns the resulting DataRow.
        /// </summary>
        /// <param name="row">The string array to be converted.</param>
        /// <param name="table">The target DataTable to which the new DataRow will be added.</param>
        /// <returns>The newly created DataRow populated with values from the string array.</returns>
        public static DataRow ConverToDataRow(string[] row, DataTable table)
        {
            DataRow newRow = table.NewRow();

            for (int i = 0; i < row.Length; i++)
            {
                newRow[i] = row[i];
            }

            return newRow;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="origin">Diretorio + nome do arquivo de origem + formato. E.g.: "C:\\Users\\ArquivoExcel.xlsx"</param>
        /// <param name="destiny">Diretorio + nome do arquivo de destino + formato. E.g.: "C:\\Users\\ArquivoExcel.csv"</param>
        /// <param name="sheet">Aba da planilha a ser convertida. E.g.: "1" (primeira aba) ou "NomeAba"</param>
        /// <param name="separator">Separador a ser utilizado para realizar a conversão. E.g.: ";"</param>
        /// <param name="columns">"Vetor de caracteres (maiúsculo ou minúsculo) contendo todas as colunas desejadas. E.g.: { "A", "b", "E", "C" } ou "{ "A:BC" } </param>
        /// <param name="rows">"Informe a primeira e última linha (ou deixe em branco). E.g.: "1:50 (linha 1 até linha 50)"</param>  
        /// <returns>DataTable</returns>
        public static DataTable GetDataTable(string origin, string destiny, string sheet, string separator, string? columns, string? rows)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                Progress += 5; // 5 

                DataSet result = Reading.GetDataSet(origin);
                Progress += 30; // 35 (pós leitura do arquivo)

                // Obtem a aba a ser convertida
                DataTable table = Reading.GetTable(sheet, result);
                var header = ConverToDataRow(Reading.GetFirstRow(Path.GetExtension(origin), table, true), table);
                table.Rows.InsertAt(header, 0); // Treatment to allow considering header
                Progress += 5; // 40

                return table;
            }
            catch (Exception)
            {

                throw;
            }
        }


        /// <summary>
        /// Realiza a conversão do arquivo Excel localizado em <paramref name="origin"/>, salva em <paramref name="destiny"/>
        /// com tratativa de exceçoes para o usuário final (arquivo inexistente no diretorio ou aberto durante a conversão)
        /// e retorna 'true' caso a conversão tenha ocorrido com sucesso
        /// </summary>
        /// <param name="origin">Diretorio + nome do arquivo de origem + formato. E.g.: "C:\\Users\\ArquivoExcel.xlsx"</param>
        /// <param name="destiny">Diretorio + nome do arquivo de destino + formato. E.g.: "C:\\Users\\ArquivoExcel.csv"</param>
        /// <param name="sheet">Aba da planilha a ser convertida. E.g.: "1" (primeira aba) ou "NomeAba"</param>
        /// <param name="separator">Separador a ser utilizado para realizar a conversão. E.g.: ";"</param>
        /// <param name="columns">"Vetor de caracteres (maiúsculo ou minúsculo) contendo todas as colunas desejadas. E.g.: { "A", "b", "E", "C" } ou "{ "A:BC" } </param>
        /// <param name="rows">"Informe a primeira e última linha (ou deixe em branco). E.g.: "1:50 (linha 1 até linha 50)"</param>        
        /// <returns>"true" se convertido com sucesso. "false" se não convertido.</returns>
        public static bool Converter(string origin, string destiny, string sheet, string separator, string? columns, string? rows)
        {
            try
            {
                return Conversion.Converter(origin, destiny, sheet, separator, columns, rows);
            }

            // TODO: considerar retirar todos os catchs

            #region Se arquivo nao localizado        
            catch (FileNotFoundException nffEx) when (nffEx.HResult.Equals(-2147024894))
            {

                //var result3 = MessageBox.Show(
                //                   "O arquivo '" + Path.GetFileName(origin) + "' não foi localizado. Por favor, verifique se o arquivo está presente no repositório de origem e confirme para continuar: "
                //                   + "\n\n" + origin,
                //                   "Aviso",
                //                   MessageBoxButtons.OKCancel,
                //                   MessageBoxIcon.Exclamation);




                return false;
            }
            #endregion

            #region Se arquivo esta em uso
            catch (IOException eiEx) when (eiEx.HResult.Equals(-2147024864))
            {

                //countOpen++; // Contador de tentativas com falha de arquivo aberto

                //if (countOpen >= 2) // Se necessario forçar o fechamento do Excel (a partir do 2 caso)
                //{

                //    var result1 = MessageBox.Show(
                //       "Parece que o arquivo ainda continua em uso. Deseja forçar o encerramento do Excel e tentar novamente? \n\nAs alterações não serão salvas.",
                //       "Aviso",
                //       MessageBoxButtons.YesNo,
                //       MessageBoxIcon.Exclamation);

                //    if (result1 == DialogResult.Yes)
                //    {
                //        CloseExcel(); // Encerra todos os processos do Excel
                //        Thread.Sleep(1500); // Aguarda o excel fechar por completo durante 1,5 segundo
                //        goto again; // Tenta realizar a conversão novamente

                //    } // Se No, continua execucao abaixo

                //}

                //var result2 = MessageBox.Show(
                //    "O arquivo '" + Path.GetFileName(origin) + "' ou '" + Path.GetFileName(destiny) + "' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida confirme para continuar.",
                //    "Aviso",
                //    MessageBoxButtons.OKCancel,
                //    MessageBoxIcon.Error);


                //if (result2 == DialogResult.OK)
                //{
                //    goto again; // Tenta realizar a conversão novamente
                //}
                //else // Se cancelar
                //{
                //    return false;
                //}

                return false;

            }

            #endregion

            #region Se arquivo em formato não suportado
            catch (ExcelDataReader.Exceptions.HeaderException heEx) when (heEx.HResult.Equals(-2147024894))
            {

                throw new Exception($"Erro! Sem suporte para converter o arquivo '{Path.GetExtension(origin)}'.");

            }
            #endregion


        }



    }
}
