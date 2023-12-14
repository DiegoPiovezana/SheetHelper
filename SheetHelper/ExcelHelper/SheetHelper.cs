using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

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
                string columnName = string.Empty;

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
        /// Unpacks a .GZ file.
        /// </summary>
        /// <param name="zipFile">The location and name of the compressed file. E.g.: 'C:\\Files\\Report.zip'</param>
        /// <param name="pathDestiny">The directory where the uncompressed file will be saved (with or without the destination file name). E.g.: 'C:\\Files\\' or 'C:\\Files\\Converted.xlsx'</param>
        /// <returns>The path of the uncompressed file if successful, otherwise null.</returns>
        public static string? UnGZ(string zipFile, string pathDestiny)
        {
            try
            {
                using var compressedFileStream = File.Open(zipFile, FileMode.Open, FileAccess.Read);
                string fileConverted;

                if (string.IsNullOrEmpty(Path.GetExtension(pathDestiny))) // If the format to be converted is not specified, try to get it from the file name
                {
                    string originalFileName = Path.GetFileName(compressedFileStream.Name).Replace(".gz", "").Replace(".GZ", "");
                    string formatOriginal = Regex.Match(Path.GetExtension(originalFileName), @"\.[A-Za-z]*").Value;
                    fileConverted = $"{pathDestiny}{Path.GetFileNameWithoutExtension(originalFileName)}{formatOriginal}";
                }
                else
                {
                    fileConverted = pathDestiny;
                }

                using FileStream outputFileStream = File.Create(fileConverted);
                using var decompressor = new GZipStream(compressedFileStream, CompressionMode.Decompress);
                decompressor.CopyTo(outputFileStream);

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
        public static string? UnZIP(string? zipFile, string pathDestiny)
        {
            try
            {
                string directoryZIP = Path.Combine(pathDestiny, "CnvrtdZIP");

                // Extract to a new directory
                ZipFile.ExtractToDirectory(zipFile, directoryZIP);

                string fileLocation = Directory.EnumerateFiles(directoryZIP).First(); // Get the location of the file
                string fileDestiny = Path.Combine(pathDestiny, Path.GetFileName(fileLocation)); // Destination location of the file

                if (File.Exists(fileDestiny))
                    File.Delete(fileDestiny); // If the file already exists, delete it

                File.Move(fileLocation, fileDestiny); // Move it to the target location
                Directory.Delete(directoryZIP, true); // Delete the previously created directory

                return fileDestiny;
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
        public static string? UnzipAuto(string? zipFile, string pathDestiny, bool mandatory = true)
        {
            try
            {
                if (string.IsNullOrEmpty(zipFile)) return null;

                Directory.CreateDirectory(pathDestiny);

            restart:

                //using (var stream = File.Open(zipFile, FileMode.Open, FileAccess.Read))
                //{
                switch (Path.GetExtension(zipFile).ToLower())
                {
                    case ".gz":
                        zipFile = UnGZ(zipFile, pathDestiny);
                        goto restart;

                    case ".zip":
                        //stream.Close();
                        zipFile = UnZIP(zipFile, pathDestiny);
                        goto restart;

                    default:
                        if (mandatory) throw new Exception("Unable to extract this file!");
                        else return zipFile;
                }
                //}
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
        public static DataRow ConvertToDataRow(string[] row, DataTable table)
        {
            DataRow newRow = table.NewRow();

            if (row.Length <= table.Columns.Count)
            {
                for (int i = 0; i < row.Length; i++)
                {
                    newRow[i] = row[i];
                }
            }
            else
            {
                throw new ArgumentException("The length of the row array exceeds the number of columns in the table.");
            }

            return newRow;
        }


        /// <summary>
        /// Retrieves the first row of a DataTable.
        /// </summary>       
        /// <param name="table">The DataTable containing the data.</param>
        /// <param name="header">If true, it will get the header (columns name).</param>
        /// <returns>An array of strings representing the first row of the DataTable.</returns>
        public static string[] GetFirstRow(DataTable table, bool header = true)
        {
            if (header)
            {
                return table.Columns.Cast<DataColumn>()
                    .Select(column => column.ColumnName)
                    .ToArray();
            }
            else
            {
                if (table.Rows.Count > 0)
                {
                    return table.Rows[0].ItemArray
                        .Select(item => item.ToString())
                        .ToArray();
                }
                else
                {
                    return Array.Empty<string>();
                }
            }
        }

        /// <summary>
        /// Gets the name of the sheets in the workbook
        /// </summary>
        /// <param name="filePath">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx"</param>   
        /// <returns></returns>
        public static List<DataTable> GetSheets(string filePath)
        {
            try
            {
                var dataSet = GetDataSet(filePath);
                return dataSet.Tables.Cast<DataTable>().ToList(); // DataTableCollection              
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Gets the name of the sheets in the workbook with the respective dataTable.
        /// </summary>
        /// <param name="filePath">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx"</param>
        /// <param name="minQtdRows">The minimum number of lines a tab needs to have, otherwise it will be ignored.</param>
        /// <param name="formatName"></param>
        /// <returns>Dictionary containing the name of the tabs and the DataTable. If desired, consider using 'sheetDictionary.Values.ToList()' to obtain a list of DataTables.</returns>
        public static Dictionary<string, DataTable> GetNameSheets(string filePath, int minQtdRows = 0, bool formatName = false)
        {
            try
            {
                var dataSet = GetDataSet(filePath);

                if (minQtdRows == 0 && formatName == false)
                {
                    return dataSet.Tables.Cast<DataTable>().ToDictionary(table => table.TableName);
                }

                Dictionary<string, DataTable> sheetDictionary = new();

                foreach (var sheet in dataSet.Tables.Cast<DataTable>())
                {
                    if (sheet.Rows.Count >= minQtdRows)
                    {
                        if (formatName)
                        {
                            sheetDictionary.Add(sheet.TableName, sheet);
                        }
                        else
                        {
                            sheetDictionary.Add(NormalizeText(sheet.TableName), sheet);
                        }
                    }
                }

                return sheetDictionary;

            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Normalizes the text by removing accents and spaces.
        /// </summary>
        /// <param name="text"></param>
        /// <returns>Text normalized.</returns>
        public static string NormalizeText(string text)
        {
            string normalizedString = text.Normalize(NormalizationForm.FormD);
            StringBuilder stringBuilder = new();

            foreach (char c in normalizedString)
            {
                UnicodeCategory unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark) { stringBuilder.Append(c); }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC).Replace(" ", "").ToLower();
        }

        /// <summary>
        /// Reads the file and gets the dataset of worksheet.
        /// <br>Note.: The header is the name of the columns.</br>
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx"</param>       
        /// <returns>DataSet</returns>
        public static DataSet GetDataSet(string origin)
        {
            try
            {
                Treatment.ValidateOrigin(origin);
                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                Progress += 5; // 5 

                DataSet result = Reading.GetDataSet(origin);
                Progress += 25; // 35 (after reading the file)

                return result;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Reads the file and gets the datatable of the specified sheet.
        /// <br>Note.: The header is the name of the columns.</br>
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx"</param>
        /// <param name="sheet">Tab of the worksheet to be converted. E.g.: "1" (first sheet) or "TabName"</param>
        /// <returns>DataTable</returns>
        public static DataTable GetDataTable(string origin, string sheet)
        {
            try
            {
                Treatment.ValidateSheet(sheet);

                var result = GetDataSet(origin); // 35 (after reading the file)

                // Get the sheet to be converted
                DataTable table = Reading.GetTableByDataSet(sheet, result);

                // Handling to allow header consideration (XLS case)
                table = Reading.TreatHeader(table, Path.GetExtension(origin));
                Progress += 5; // 40

                return table;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Performs the conversion of the <paramref name="dataTable"/>, saves in <paramref name="destiny"/>. 
        /// </summary>
        /// <param name="dataTable">DataTable to be converted.</param>
        /// <param name="destiny">Directory + destination file name + format. E.g.: "C:\\Users\\FileExcel.csv".</param>
        /// <param name="separator">Separator to be used to perform the conversion. E.g.: ";".</param>
        /// <param name="columns">"Enter the columns or their range. E.g.: "A:H, 4:9, 4:-9, B, 75, -2".</param>
        /// <param name="rows">"Enter the rows or their range. E.g.: "1:23, -34:56, 70, 75, -1".</param>
        /// <returns>"true" if converted successfully. "false" if not converted.</returns>
        public static bool SaveDataTable(DataTable dataTable, string destiny, string separator, string? columns, string? rows)
        {
            try
            {
                Treatment.Validate(destiny, separator, columns, rows);
                return Conversion.ConverterDataTable(dataTable, destiny, separator, columns, rows);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Performs the conversion of the Excel file located in <paramref name="origin"/>, saves in <paramref name="destiny"/>.      
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx."</param>
        /// <param name="destiny">Directory + destination file name + format. E.g.: "C:\\Users\\FileExcel.csv."</param>
        /// <param name="sheet">Tab of the worksheet to be converted. E.g.: "1" (first sheet) or "TabName".</param>
        /// <param name="separator">Separator to be used to perform the conversion. E.g.: ";".</param>
        /// <param name="columns">"Enter the columns or their range. E.g.: "A:H, 4:9, 4:-9, B, 75, -2".</param>
        /// <param name="rows">"Enter the rows or their range. E.g.: "1:23, -34:56, 70, 75, -1".</param>
        /// <returns>"true" if converted successfully. "false" if not converted.</returns>
        public static bool Converter(string origin, string destiny, string sheet, string separator, string? columns, string? rows)
        {
            int countOpen = 0; // Count of times Excel was open

        again:

            try
            {
                Progress = 5;

                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
                if (origin == null) return false;

                if (!Treatment.CheckConvertNecessary(origin, destiny, sheet, separator, columns, rows))
                {
                    // If no conversion is needed
                    Progress = 100;
                    File.Copy(origin, destiny, true);
                    return true;
                }

                DataTable table = GetDataTable(origin, sheet); // Progress 40        

                return SaveDataTable(table, destiny, separator, columns, rows);
            }


#if NETFRAMEWORK

            #region If file not found       
            catch (Exception ex) when (ex.InnerException.Message.Contains("file not found"))
            {
                var result1 = MessageBox.Show(
                                   "O arquivo '" + Path.GetFileName(origin) + "' não foi localizado. Por favor, verifique se o arquivo está presente no repositório de origem e confirme para continuar: "
                                   + "\n\n" + origin,
                                   "Aviso",
                                   MessageBoxButtons.OKCancel,
                                   MessageBoxIcon.Exclamation);


                if (result1 == DialogResult.OK)
                {
                    goto again; // Try conversion again
                }

                return false;
            }
            #endregion

            #region If file is in use
            catch (Exception eiEx) when (eiEx.Message.Contains("file being used by another process") || eiEx.Message.Contains("sendo usado por outro processo"))
            {
                countOpen++; // Counter for failed attempts with open file

                if (countOpen >= 2) // If it is necessary to force Excel closure (from the 2nd attempt onwards)                
                {

                    var result2 = MessageBox.Show(
                       "Parece que o arquivo ainda continua em uso. Deseja forçar o encerramento do Excel e tentar novamente? \n\nTodos os Excel abertos serão fechados e as alterações não serão salvas!",
                       "Aviso",
                       MessageBoxButtons.YesNo,
                       MessageBoxIcon.Exclamation);

                    if (result2 == DialogResult.Yes)
                    {
                        CloseExcel(); // Close all Excel processes
                        System.Threading.Thread.Sleep(1500); // Wait for Excel to close completely for 1.5 seconds
                        goto again; // Try conversion again

                    } // If No, continue execution below
                }

                var result3 = MessageBox.Show(
                    "O arquivo '" + Path.GetFileName(origin) + "' ou '" + Path.GetFileName(destiny) + "' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida confirme para continuar.",
                    "Aviso",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Error);

                if (result3 == DialogResult.OK)
                {
                    goto again; // Try conversion again
                }
                else // If canceled
                {
                    return false;
                }
            }

            #endregion
#endif


            #region If file in unsupported format
            catch (ExcelDataReader.Exceptions.HeaderException heEx) when (heEx.HResult.Equals(-2147024894))
            {

                throw new Exception($"Erro! Sem suporte para converter o arquivo '{Path.GetExtension(origin)}'.");

            }
            #endregion


            catch (Exception ex)
            {
                throw ex;
            }
        }


    }
}
