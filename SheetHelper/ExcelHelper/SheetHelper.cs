using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

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

                if (File.Exists(fileDestiny)) File.Delete(fileDestiny); // If the file already exists, delete it                   

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
        /// Unzip a .zip or .gz file.    
        /// </summary>
        /// <param name="zipFile">The location and name of the compressed file. E.g.: 'C:\\Files\\Report.zip'</param>
        /// <param name="pathDestiny">The directory where the extracted file will be saved. E.g.: 'C:\\Files\\'</param>
        /// <param name="mandatory">If true, it indicates that the extraction must occur, otherwise, it will show an error. If false, if the conversion does not happen, nothing happens.</param>
        /// <returns>The path of the extracted file.</returns>
        public static string UnzipAuto(string zipFile, string pathDestiny, bool mandatory = true)
        {
            try
            {

            restart:

                using (var stream = File.Open(zipFile, FileMode.Open, FileAccess.Read))
                {
                    switch (Path.GetExtension(zipFile).ToLower())
                    {
                        case ".gz":
                            zipFile = UnGZ(stream, pathDestiny);
                            goto restart;

                        case ".zip":
                            stream.Close();
                            zipFile = UnZIP(zipFile, pathDestiny);
                            goto restart;

                        default:
                            if (mandatory) throw new Exception("Unable to extract this file!");
                            else return zipFile;
                    }
                }
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
        /// Reads the file and gets the datatable of the specified sheet
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx"</param>
        /// <param name="sheet">Tab of the worksheet to be converted. E.g.: "1" (first sheet) or "TabName"</param>
        /// <returns>DataTable</returns>
        public static DataTable GetDataTable(string origin, string sheet)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                Progress += 5; // 5 

                DataSet result = Reading.GetDataSet(origin);
                Progress += 30; // 35 (after reading the file)

                // Get the sheet to be converted
                DataTable table = Reading.GetTableByDataSet(sheet, result);

                // Handling to allow header consideration (XLS case)
                string[]? header = Reading.GetFirstRow(Path.GetExtension(origin), table, true);
                if (header != null) { table.Rows.InsertAt(ConverToDataRow(header, table), 0); }
                Progress += 5; // 40

                return table;
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Performs the conversion of the Excel file located in <paramref name="origin"/>, saves in <paramref name="destiny"/>
        /// Handling exceptions for the end user (file does not exist in directory or opened during conversion)
        /// and returns 'true' if the conversion was successful
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx"</param>
        /// <param name="destiny">Directory + destination file name + format. E.g.: "C:\\Users\\FileExcel.csv"</param>
        /// <param name="sheet">Tab of the worksheet to be converted. E.g.: "1" (first sheet) or "TabName"</param>
        /// <param name="separator">Separator to be used to perform the conversion. E.g.: ";"</param>
        /// <param name="columns">"Enter the columns or their range. E.g.: "A:H, 4:9, 4:-9, B, 75, -2" </param>
        /// <param name="rows">"Enter the rows or their range. E.g.: "1:23, -34:56, 70, 75, -1"</param>
        /// <returns>"true" if converted successfully. "false" if not converted.</returns>
        public static bool Converter(string origin, string destiny, string sheet, string separator, string? columns, string? rows)
        {
            try
            {
                return Conversion.Converter(origin, destiny, sheet, separator, columns, rows);
            }
            catch (Exception)
            {
                throw;
            }
        }



    }
}
