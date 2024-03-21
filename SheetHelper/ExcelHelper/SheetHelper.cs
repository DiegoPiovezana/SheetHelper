using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SH
{
    /// <summary>
    /// Fast and lightweight library for easy read and conversion of large Excel files
    /// </summary>
    public class SheetHelper
    {
        /// <summary>
        /// Represents the conversion progress. E.g.: If 100%, the conversion is fully completed.
        /// </summary>    
        public static int Progress { get; internal set; }

        /// <summary>
        /// (Optional) The dictionary can specify characters that should not be maintained after conversion (line breaks, for example) and which replacements should be performed in each case.
        /// </summary>
        public static Dictionary<string, string>? ProhibitedItems { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// (Optional) Ignored exceptions will attempt to be handled internally. If it is not possible, it will just return false and the exception will not be thrown.
        /// <para>By default, it will ignore the exception when the source or destination file is in use. If .NET Framework will display a warning to close the file, otherwise it will return false.</para>
        /// </summary>
        public static List<string>? TryIgnoreExceptions { get; set; } = new List<string>() { "E-0001-SH" };


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

                ZipFile.ExtractToDirectory(zipFile, directoryZIP); // Extract to a new directory

                string fileLocation = Directory.EnumerateFiles(directoryZIP).First(); // Get the location of the file
                string fileDestiny = Path.Combine(pathDestiny, Path.GetFileName(fileLocation)); // Destination location of the file

                if (File.Exists(fileDestiny)) File.Delete(fileDestiny); // If the file already exists, delete it

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
        /// <para>Please visit https://bit.ly/SheetHelper_Features to learn more.</para>
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
                        if (mandatory) throw new Exception("E-0000-SH: Unable to extract this file!");
                        else return zipFile;
                }
                //}
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (Directory.Exists(@".\SheetHelper")) Directory.Delete(@".\SheetHelper", true);
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
                for (int i = 0; i < row.Length; i++) { newRow[i] = row[i]; }
            }
            else
            {
                throw new ArgumentException("E-0000-SH: The length of the row array exceeds the number of columns in the table.");
            }

            return newRow;
        }


        /// <summary>
        /// Retrieves a row of a DataTable.
        /// </summary>       
        /// <param name="table">The DataTable containing the data.</param>
        /// <param name="header">If true, it will get the header (columns name).</param>
        /// <param name="indexRow">index of the line to be obtained (in addition to the header). Enter negative to get just the header.</param>
        /// <returns>An array of strings representing a row of the DataTable.</returns>
        public static string[] GetRowArray(DataTable table, bool header = true, int indexRow = 0)
        {
            if (header)
            {
                return table.Columns.Cast<DataColumn>()
                    .Select(column => column.ColumnName)
                    .ToArray();
            }
            else
            {
                if (table.Rows.Count > 0 && indexRow >= 0)
                {
                    return table.Rows[indexRow].ItemArray
                        .Select(cell => cell.ToString())
                        .ToArray();
                }
                else
                {
                    return Array.Empty<string>();
                }
            }
        }

        /// <summary>
        /// Gets the sheets in the workbook with the respective dataTable.
        /// </summary>
        /// <param name="filePath">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx"</param>
        /// <param name="minQtdRows">The minimum number of lines a tab needs to have, otherwise it will be ignored.</param>
        /// <param name="formatName">If true, all spaces and special characters from tab names will be removed.</param>
        /// <returns>Dictionary containing the name of the tabs and the DataTable. If desired, consider using 'sheetDictionary.Values.ToList()' to obtain a list of DataTables.</returns>
        public static Dictionary<string, DataTable> GetAllSheets(string filePath, int minQtdRows = 0, bool formatName = false)
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
                    if ((sheet.Rows.Count + (sheet.Columns.Count > 0 ? 1 : 0)) >= minQtdRows)
                    {
                        if (!formatName) { sheetDictionary.Add(sheet.TableName, sheet); }
                        else { sheetDictionary.Add(NormalizeText(sheet.TableName), sheet); }
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
        /// <para>Example: " Hot Café" => "hot_cafe" </para>
        /// </summary>
        /// <param name="text">Text to be normalized.</param>
        /// <param name="replaceSpace">Character to replace spaces. E.g.: "_"</param>
        /// <param name="toLower">If true, the text will be converted to lowercase.</param>
        /// <returns>Text normalized.</returns>
        public static string NormalizeText(string? text, char replaceSpace = '_', bool toLower = true)
        {
            try
            {
                if (string.IsNullOrEmpty(text)) throw new ArgumentException("E-0000-SH: The text is null or empty.");
                string normalizedString = text.Trim().Normalize(NormalizationForm.FormD);
                StringBuilder stringBuilder = new();

                foreach (char c in normalizedString)
                {
                    UnicodeCategory unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                    if (unicodeCategory != UnicodeCategory.NonSpacingMark) { stringBuilder.Append(c); }
                }

                if (toLower) return stringBuilder.ToString().Normalize(NormalizationForm.FormC).Replace(' ', replaceSpace).ToLower();
                return stringBuilder.ToString().Normalize(NormalizationForm.FormC).Replace(' ', replaceSpace);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Fixes a string containing items by replacing line breaks and semicolons with commas,
        /// removing spaces, single quotes, and double quotes, and ensuring proper comma separation.
        /// </summary>
        /// <param name="items">The string containing items to be fixed.</param>
        /// <returns>The fixed string with proper item separation.</returns>
        public static string FixItems(string items)
        {
            if (!string.IsNullOrEmpty(items))
            {
                items = items.Replace("\n", ",").Replace(";", ","); // Replace line breaks and semicolons with commas
                items = Regex.Replace(items, @"\s+|['""]+", ""); // Remove spaces, single quotes, and double quotes
                items = Regex.Replace(items, ",{2,}", ",").Trim(','); // Remove repeated commas and excess spaces
            }
            return items; // "123123,13514,31234"
        }

        /// <summary>
        /// Converts a string in JSON format to a dictionary.
        /// <para>Example 1: "{ \"key1\" : \"value1\", \"key2\" : \"value2\", \"key3\" : \"value3\" }"</para> 
        /// <para>Example 2: "{\"\\n\": \" \", \"\\r\": \"\", \";\": \",\"}"</para>        
        /// </summary>
        /// <param name="jsonTextItems">String JSON containing the key-value pairs to be converted.</param>        
        /// <returns>A dictionary containing the extracted key-value pairs from the string.</returns>
        public static Dictionary<string, string>? GetDictionaryJson(string jsonTextItems)
        {
            try
            {
                if (string.IsNullOrEmpty(jsonTextItems))
                    throw new ArgumentException($"E-0000-SH: The '{nameof(jsonTextItems)}' parameter is null or empty.");

                return JsonSerializer.Deserialize<Dictionary<string, string>>(jsonTextItems);
            }
            catch (Exception ex)
            {
                throw new Exception("E-0000-SH: An error occurred while processing the text items in JSON format.", ex);
            }
        }

        /// <summary>
        /// Serializes a dictionary of strings into a JSON representation.
        /// </summary>
        /// <param name="dictionary">The dictionary to be serialized into JSON.</param>
        /// <returns>A string containing the JSON representation of the provided dictionary.</returns>
        /// <exception cref="ArgumentException">Thrown if the dictionary is null or empty.</exception>
        /// <exception cref="Exception">Thrown if an error occurs while serializing the dictionary to JSON.</exception>
        public static string GetJsonDictionary(Dictionary<string, string> dictionary)
        {
            try
            {
                if (dictionary == null || dictionary.Count == 0)
                    throw new ArgumentException("E-0000-SH: The dictionary is null or empty.");

                return JsonSerializer.Serialize(dictionary);
            }
            catch (Exception ex)
            {
                throw new Exception("E-0000-SH: An error occurred while serializing the dictionary to JSON.", ex);
            }
        }

        /// <summary>
        /// Reads the file and gets the dataset of worksheet.
        /// <br>Note.: The header is the name of the columns.</br>
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx"</param>       
        /// <returns>DataSet</returns>
        public static DataSet GetDataSet(string? origin)
        {
            try
            {
                Treatment.ValidateOrigin(origin);
                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
                if (string.IsNullOrEmpty(origin)) throw new Exception("E-0000-SH: The 'origin' is null or empty.");

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
        public static DataTable? GetDataTable(string origin, string sheet = "1")
        {
            int countOpen = 0; // Count of times Excel was open

        again:

            try
            {
                Treatment.ValidateSheet(sheet);

                var result = GetDataSet(origin); // 35 (after reading the file)

                // Get the sheet to be converted
                DataTable table = Reading.GetTableByDataSet(sheet, result);

                // Handling to allow header consideration (XLS case)
                table = Reading.FirstRowToHeader(table, Path.GetExtension(origin));
                Progress += 5; // 40

                return table;
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

                return null;
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
                    $"O arquivo '{Path.GetFileName(origin)}' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida confirme para continuar.",
                    "Aviso",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Error);

                if (result3 == DialogResult.OK)
                {
                    goto again; // Try conversion again
                }
                else // If canceled
                {
                    return null;
                }
            }

            #endregion
#endif


            #region If file in unsupported format
            catch (ExcelDataReader.Exceptions.HeaderException heEx) when (heEx.HResult.Equals(-2147024894))
            {
                throw new Exception($"Erro E-99101-SH: Sem suporte para converter o arquivo '{Path.GetExtension(origin)}'.");
            }
            #endregion

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
        public static bool SaveDataTable(DataTable dataTable, string destiny, string separator = ";", string? columns = null, string? rows = null)
        {
            int countOpen = 0; // Count of times Excel was open

        again:

            try
            {
                Treatment.Validate(destiny, separator, columns, rows);
                return Conversion.SaveDataTable(dataTable, destiny, separator, columns, rows);
            }
#if NETFRAMEWORK                        

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
                        CloseExcel(); 
                        System.Threading.Thread.Sleep(1500);
                        goto again;

                    } // If No, continue execution below
                }

                var result3 = MessageBox.Show(
                    $"O arquivo '{Path.GetFileName(destiny)}' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida confirme para continuar.",
                    "Aviso",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Error);

                if (result3 == DialogResult.OK)
                {
                    goto again;
                }
                else // If canceled
                {
                    return false;
                }
            }

            #endregion
#endif
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
        /// <param name="minRows">(Optional) The minimum number of lines a tab needs to have, otherwise it will be ignored.</param>
        /// <returns>"true" if converted successfully. "false" if not converted.</returns>
        public static bool Converter(string? origin, string? destiny, string sheet, string separator, string? columns, string? rows, int minRows = 1)
        {
            try
            {
                Progress = 5;

                if (string.IsNullOrEmpty(destiny)) throw new ArgumentException($"E-0000-SH: The '{nameof(destiny)}' is null or empty.", destiny);
                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
                if (string.IsNullOrEmpty(origin)) throw new Exception("E-0000-SH: The 'origin' is null or empty.");

                if (!Treatment.CheckConvertNecessary(origin, destiny, sheet, separator, columns, rows))
                {
                    // If no conversion is needed
                    Progress = 100;
                    File.Copy(origin, destiny, true);
                    if (Directory.Exists(@".\SheetHelper\")) Directory.Delete(@".\SheetHelper\", true);
                    return true;
                }

                DataTable? table = GetDataTable(origin, sheet); // Progress 40        
                if (table == null || table.Rows.Count < minRows - 1) throw new Exception("E-0000-SH: The sheet does not have the minimum number of rows.");

                return SaveDataTable(table, destiny, separator, columns, rows);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Converts all spreadsheet tabs considering all rows and columns.
        /// <para>NOTE.: Use the Convert or SaveDataTable method for further customizations.</para>
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx."</param>
        /// <param name="destiny">Directory + destination file name + format. E.g.: "C:\\Users\\FileExcel.csv."</param>
        /// <param name="sheets">Enter the names or indexes of the sheets to be converted. Enter null to convert all sheets.</param>
        /// <param name="separator">Separator to be used to perform the conversion. E.g.: ";".</param>
        /// <param name="columns">"Enter the the group of columns or their range for each sheet.</param>
        /// <param name="rows">"Enter the group of rows or their range for each sheet.</param>
        /// <param name="minRows">(Optional) The minimum number of lines a tab needs to have, otherwise it will be ignored.</param>
        /// <returns>Number of tabs successfully saved.</returns>
        public static int Converter(string? origin, string? destiny, ICollection<string>? sheets, string separator = ";", ICollection<string>? columns = default, ICollection<string>? rows = default, int minRows = 1)
        {
            try
            {
                if (string.IsNullOrEmpty(destiny)) throw new Exception("E-0000-SH: The 'destiny' is null or empty.");
                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
                if (string.IsNullOrEmpty(origin)) throw new Exception("E-0000-SH: The 'origin' is null or empty.");

                var sheetsDictionary = GetAllSheets(origin, minRows, true);

                if (sheets == null || sheets.Count == 0) sheets = sheetsDictionary.Keys;
                if (columns == null || columns.Count == 0) columns = Enumerable.Repeat("", sheets.Count).ToList();
                if (rows == null || rows.Count == 0) rows = Enumerable.Repeat("", sheets.Count).ToList();

                if (sheets.Count != columns.Count || sheets.Count != rows.Count)
                {
                    throw new Exception("E-0000-SH: The number of sheets, columns and rows must be the same.");
                }

                int saveSuccess = default;

                for (int i = 0; i < sheets.Count; i++) // Name or index of the sheet              
                {
                    var sheetId = NormalizeText(sheets.Skip(i).FirstOrDefault());

                    DataTable? dtSheet = null;

                    if (int.TryParse(sheetId, out int indexSheet)) // Index of the sheet
                    {
                        dtSheet = sheetsDictionary.ElementAtOrDefault(indexSheet - 1).Value;
                    }
                    else if (sheetsDictionary.ContainsKey(sheetId))// Name of the sheet
                    {
                        dtSheet = sheetsDictionary[sheetId];

                        //indexSheet =
                        //    sheetsDictionary.FirstOrDefault(x => x.Value == sheetsDictionary[sheetId]).Key != null ?
                        //    Array.IndexOf(sheetsDictionary.Keys.ToArray(), sheetsDictionary.FirstOrDefault(x => x.Value == sheetsDictionary[sheetId]).Key) :
                        //    -1;
                    }

                    if (dtSheet == null) throw new Exception("E-0000-SH: Failed to locate sheet to be converted.");

                    //var columnSheet = columns.Skip(indexSheet).FirstOrDefault();
                    //var rowSheet = rows.Skip(indexSheet).FirstOrDefault();

                    var columnSheet = columns.Skip(i).FirstOrDefault();
                    var rowSheet = rows.Skip(i).FirstOrDefault();

                    string dest = Path.Combine(Path.GetDirectoryName(destiny), $"{Path.GetFileNameWithoutExtension(destiny)}__{sheetId}{Path.GetExtension(destiny)}");
                    saveSuccess += SaveDataTable(dtSheet, dest, separator, columnSheet, rowSheet) ? 1 : 0;
                }

                return saveSuccess;
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        /// Converts all spreadsheet tabs considering all rows and columns.
        /// <para>NOTE.: Use the Convert or SaveDataTable method for further customizations.</para>
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx."</param>
        /// <param name="destiny">Directory + destination file name + format. E.g.: "C:\\Users\\FileExcel.csv."</param>
        /// <param name="minRows">(Optional) The minimum number of lines a tab needs to have, otherwise it will be ignored.</param>
        /// <param name="separator">Separator to be used to perform the conversion. E.g.: ";".</param>
        /// <returns>True, if success.</returns>
        public static bool ConvertAllSheets(string? origin, string? destiny, int minRows = 1, string separator = ";")
        {
            try
            {
                if (string.IsNullOrEmpty(destiny)) throw new Exception("E-0000-SH: The 'destiny' is null or empty.");
                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
                if (string.IsNullOrEmpty(origin)) throw new Exception("E-0000-SH: The 'origin' is null or empty.");

                foreach (var sheet in GetAllSheets(origin, minRows, true))
                {
                    string dest = Path.Combine(Path.GetDirectoryName(destiny), $"{Path.GetFileNameWithoutExtension(destiny)}__{sheet.Key}{Path.GetExtension(destiny)}");
                    SaveDataTable(sheet.Value, dest, separator, "", "");
                }

                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }


    }
}
