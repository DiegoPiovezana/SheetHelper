using SH.ExcelHelper.Treatments;
using SH.Exceptions;
using SH.Globalization;
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

namespace SH.ExcelHelper.Tools
{
    internal class Features
    {
        private readonly SheetHelper _sheetHelper;
        private readonly Validations _validations;
        private readonly Writing _writing;

        public Features(SheetHelper sheetHelper)
        {
            _sheetHelper = sheetHelper;
            _validations = new Validations(sheetHelper);
            _writing = new Writing(sheetHelper);

        }



        internal void CloseExcel()
        {
            try
            {
                var processes = from p in Process.GetProcessesByName("EXCEL") select p;
                foreach (var process in processes) process.Kill();
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(CloseExcel), ex), ex);
            }
        }

        internal int GetIndexColumn(string? columnName)
        {
            try
            {
                if (string.IsNullOrEmpty(columnName?.Trim())) throw new ArgumentSHException(nameof(columnName), nameof(GetIndexColumn));

                int sum = 0;
                foreach (var character in columnName)
                {
                    sum *= 26;
                    sum += character - 'A' + 1;
                }

                return sum; // E.g.: A = 1, Z = 26, AA = 27
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(GetIndexColumn), ex), ex);
            }
        }

        internal string GetNameColumn(int columnIndex)
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
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(GetNameColumn), ex), ex);
            }
        }

        internal string? UnGZ(string gzFile, string pathDestiny)
        {
            try
            {
                if (string.IsNullOrEmpty(gzFile?.Trim()) || !File.Exists(gzFile)) throw new ArgumentSHException(nameof(gzFile), nameof(UnGZ));

                if (!string.IsNullOrEmpty(pathDestiny))
                {
                    if (!Directory.Exists(pathDestiny)) Directory.CreateDirectory(pathDestiny);
                }
                else
                {
                    throw new ArgumentSHException(nameof(pathDestiny), nameof(UnGZ));
                }

                using var compressedFileStream = File.Open(gzFile, FileMode.Open, FileAccess.Read);
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
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(UnGZ), ex), ex);
            }
        }

        internal string? UnZIP(string? zipFile, string pathDestiny)
        {
            try
            {
                if (!string.IsNullOrEmpty(zipFile?.Trim()))
                {
                    if (!File.Exists(zipFile)) throw new FileNotFoundSHException(nameof(zipFile));
                }
                else
                {
                    throw new PathFileNullSHException(nameof(zipFile));
                }

                if (!string.IsNullOrEmpty(pathDestiny))
                {
                    if (!Directory.Exists(pathDestiny)) Directory.CreateDirectory(pathDestiny);
                }
                else
                {
                    throw new ArgumentSHException(nameof(pathDestiny), nameof(UnZIP));
                }

                string directoryZIP = Path.Combine(pathDestiny, "CnvrtdZIP");

                ZipFile.ExtractToDirectory(zipFile, directoryZIP); // Extract to a new directory

                string fileLocation = Directory.EnumerateFiles(directoryZIP).First(); // Get the location of the file
                string fileDestiny = Path.Combine(pathDestiny, Path.GetFileName(fileLocation)); // Destination location of the file

                if (File.Exists(fileDestiny)) File.Delete(fileDestiny); // If the file already exists, delete it

                File.Move(fileLocation, fileDestiny); // Move it to the target location
                Directory.Delete(directoryZIP, true); // Delete the previously created directory

                return fileDestiny;
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(UnZIP), ex), ex);
            }
        }

        internal string? UnzipAuto(string? zipFile, string pathDestiny, bool mandatory = true)
        {
            try
            {
                Directory.CreateDirectory(pathDestiny);

            restart:

                if (string.IsNullOrEmpty(zipFile)) return null;
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
                        if (mandatory) throw new UnableUnzipSHException(zipFile);
                        else return zipFile;
                }
                //}
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(UnzipAuto), ex), ex);
            }
            finally
            {
                if (Directory.Exists(@".\SheetHelper")) Directory.Delete(@".\SheetHelper", true);
            }
        }

        internal DataRow ConvertToDataRow(string[] row, DataTable table)
        {
            try
            {
                DataRow newRow = table.NewRow();

                if (row.Length <= table.Columns.Count)
                {
                    for (int i = 0; i < row.Length; i++) { newRow[i] = row[i]; }
                }
                else
                {
                    throw new RowArrayOverflowDteSHException();
                }

                return newRow;
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(ConvertToDataRow), ex), ex);
            }
        }

        internal string[] GetRowArray(DataTable table, bool header = true, int indexRow = 0)
        {
            try
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
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(GetRowArray), ex), ex);
            }
        }

        internal Dictionary<string, DataTable> GetAllSheets(string filePath, int minQtdRows = 0, bool formatName = false)
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
                    if (sheet.Rows.Count + (sheet.Columns.Count > 0 ? 1 : 0) >= minQtdRows)
                    {
                        if (!formatName) { sheetDictionary.Add(sheet.TableName, sheet); }
                        else { sheetDictionary.Add(NormalizeText(sheet.TableName), sheet); }
                    }
                }

                return sheetDictionary;
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(GetAllSheets), ex), ex);
            }
        }

        internal string NormalizeText(string? text, char replaceSpace = '_', bool toLower = true)
        {
            try
            {
                if (string.IsNullOrEmpty(text)) throw new ArgumentSHException(nameof(text), nameof(NormalizeText));
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
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(NormalizeText), ex), ex);
            }
        }

        internal string FixItems(string items)
        {
            try
            {
                if (!string.IsNullOrEmpty(items))
                {
                    items = items.Replace("\n", ",").Replace(";", ","); // Replace line breaks and semicolons with commas
                    items = Regex.Replace(items, @"\s+|['""]+", ""); // Remove spaces, single quotes, and double quotes
                    items = Regex.Replace(items, ",{2,}", ",").Trim(','); // Remove repeated commas and excess spaces
                }
                return items; // "123123,13514,31234"
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(FixItems), ex), ex);
            }
        }

        internal Dictionary<string, string>? GetDictionaryJson(string jsonTextItems)
        {
            try
            {
                _validations.ValidateString(jsonTextItems,nameof(jsonTextItems), _validations.GetCallingMethodName(1));

                //if (string.IsNullOrEmpty(jsonTextItems))
                //    throw new ParamExceptionSHException(nameof(jsonTextItems), nameof(GetDictionaryJson));

                return JsonSerializer.Deserialize<Dictionary<string, string>>(jsonTextItems);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(GetDictionaryJson), ex), ex);
            }
        }

        internal string GetJsonDictionary(Dictionary<string, string> dictionary)
        {
            try
            {
                if (dictionary == null || dictionary.Count == 0)
                    throw new ArgumentSHException(nameof(dictionary), nameof(GetJsonDictionary));

                return JsonSerializer.Serialize(dictionary);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(GetJsonDictionary), ex), ex);
            }
        }

        internal DataSet GetDataSet(string? origin)
        {
            try
            {
                _validations.ValidateOriginFile(origin, nameof(origin), nameof(GetDataTable));
                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
                _validations.ValidateOriginFile(origin, nameof(origin), nameof(GetDataTable));               

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                _sheetHelper.Progress += 5; // 5 

                DataSet result = Reading.GetDataSet(origin);
                _sheetHelper.Progress += 25; // 35 (after reading the file)

                return result;
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(GetDataSet), ex), ex);
            }
        }

        internal DataTable? GetDataTable(string origin, string sheet = "1")
        {
            int countOpen = 0; // Count of times Excel was open

        again:

            try
            {
                _validations.ValidateOriginFile(origin, nameof(origin), nameof(GetDataTable));
                _validations.ValidateSheet(sheet);

                var result = GetDataSet(origin); // 35 (after reading the file)

                // Get the sheet to be converted
                DataTable table = Reading.GetTableByDataSet(sheet, result);

                // Handling to allow header consideration (XLS case)
                table = Reading.FirstRowToHeader(table, Path.GetExtension(origin));
                _sheetHelper.Progress += 5; // 40

                return table;
            }

//#if NETFRAMEWORK

//            //#region If file not found       
//            //catch (Exception ex) when (ex.InnerException.Message.Contains("file not found"))
//            //{
//            //    var result1 = MessageBox.Show(
//            //                       "O arquivo '" + Path.GetFileName(origin) + "' não foi localizado. Por favor, verifique se o arquivo está presente no repositório de origem e confirme para continuar: "
//            //                       + "\n\n" + origin,
//            //                       "Aviso",
//            //                       MessageBoxButtons.OKCancel,
//            //                       MessageBoxIcon.Exclamation);


//            //    if (result1 == DialogResult.OK)
//            //    {
//            //        goto again; // Try conversion again
//            //    }

//            //    return null;
//            //}
//            //#endregion

//            //#region If file is in use
//            //catch (Exception eiEx) when (eiEx.Message.Contains("file being used by another process") || eiEx.Message.Contains("sendo usado por outro processo"))
//            //{
//            //    countOpen++; // Counter for failed attempts with open file

//            //    if (countOpen >= 2) // If it is necessary to force Excel closure (from the 2nd attempt onwards)                
//            //    {

//            //        var result2 = MessageBox.Show(
//            //           "Parece que o arquivo ainda continua em uso. Deseja forçar o encerramento do Excel e tentar novamente? \n\nTodos os Excel abertos serão fechados e as alterações não serão salvas!",
//            //           "Aviso",
//            //           MessageBoxButtons.YesNo,
//            //           MessageBoxIcon.Exclamation);

//            //        if (result2 == DialogResult.Yes)
//            //        {
//            //            CloseExcel(); // Close all Excel processes
//            //            System.Threading.Thread.Sleep(1500); // Wait for Excel to close completely for 1.5 seconds
//            //            goto again; // Try conversion again

//            //        } // If No, continue execution below
//            //    }

//            //    var result3 = MessageBox.Show(
//            //        $"O arquivo '{Path.GetFileName(origin)}' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida confirme para continuar.",
//            //        "Aviso",
//            //        MessageBoxButtons.OKCancel,
//            //        MessageBoxIcon.Error);

//            //    if (result3 == DialogResult.OK)
//            //    {
//            //        goto again; // Try conversion again
//            //    }
//            //    else // If canceled
//            //    {
//            //        return null;
//            //    }
//            //}

//            //#endregion
//#endif


            #region If file in unsupported format
            catch (ExcelDataReader.Exceptions.HeaderException heEx) when (heEx.HResult.Equals(-2147024894))
            {
                throw new Exception($"Erro E-99101-SH: Sem suporte para converter o arquivo '{Path.GetExtension(origin)}'.");
            }
            #endregion

            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(GetDataTable), ex), ex);
            }
        }

        internal bool SaveDataTable(DataTable dataTable, string destiny, string separator = ";", string? columns = null, string? rows = null)
        {

            try
            {
                _validations.Validate(destiny, separator, columns, rows);
                return _writing.SaveDataTable(dataTable, destiny, separator, columns, rows);
            }

//#if NETFRAMEWORK                        

//            #region If file is in use
//            catch (Exception eiEx) when (eiEx.Message.Contains("file being used by another process") || eiEx.Message.Contains("sendo usado por outro processo"))
//            {
//                countOpen++; // Counter for failed attempts with open file

//                if (countOpen >= 2) // If it is necessary to force Excel closure (from the 2nd attempt onwards)                
//                {
//                    var result2 = MessageBox.Show(
//                       "Parece que o arquivo ainda continua em uso. Deseja forçar o encerramento do Excel e tentar novamente? \n\nTodos os Excel abertos serão fechados e as alterações não serão salvas!",
//                       "Aviso",
//                       MessageBoxButtons.YesNo,
//                       MessageBoxIcon.Exclamation);

//                    if (result2 == DialogResult.Yes)
//                    {
//                        CloseExcel();
//                        System.Threading.Thread.Sleep(1500);
//                        goto again;

//                    } // If No, continue execution below
//                }

//                var result3 = MessageBox.Show(
//                    $"O arquivo '{Path.GetFileName(destiny)}' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida confirme para continuar.",
//                    "Aviso",
//                    MessageBoxButtons.OKCancel,
//                    MessageBoxIcon.Error);

//                if (result3 == DialogResult.OK)
//                {
//                    goto again;
//                }
//                else // If canceled
//                {
//                    return false;
//                }
//            }

//            #endregion
//#endif
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(SaveDataTable), ex), ex);
            }
        }

        internal bool Converter(string? origin, string? destiny, string sheet, string separator, string? columns, string? rows, int minRows = 1)
        {
            try
            {
                _sheetHelper.Progress = 5;

                if (string.IsNullOrEmpty(destiny)) throw new ArgumentException($"E-0000-SH: The '{nameof(destiny)}' is null or empty.", destiny);
                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
                if (string.IsNullOrEmpty(origin)) throw new Exception("E-0000-SH: The 'origin' is null or empty.");

                if (!_validations.CheckConvertNecessary(origin, destiny, sheet, separator, columns, rows))
                {
                    // If no conversion is needed
                    _sheetHelper.Progress = 100;
                    File.Copy(origin, destiny, true);
                    if (Directory.Exists(@".\SheetHelper\")) Directory.Delete(@".\SheetHelper\", true);
                    return true;
                }

                DataTable? table = GetDataTable(origin, sheet); // Progress 40        
                if (table == null || table.Rows.Count < minRows - 1) throw new Exception("E-0000-SH: The sheet does not have the minimum number of rows.");

                return SaveDataTable(table, destiny, separator, columns, rows);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(Converter), ex), ex);
            }
        }

        internal int Converter(string? origin, string? destiny, ICollection<string>? sheets, string separator = ";", ICollection<string>? columns = default, ICollection<string>? rows = default, int minRows = 1)
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
                        if (dtSheet == null) throw new Exception("E-0000-SH: Failed to locate sheet to be converted.");
                    }
                    else if (sheetsDictionary.ContainsKey(sheetId))// Name of the sheet
                    {
                        dtSheet = sheetsDictionary[sheetId];
                        if (dtSheet == null) throw new Exception("E-0000-SH: Failed to locate sheet to be converted.");

                        //indexSheet =
                        //    sheetsDictionary.FirstOrDefault(x => x.Value == sheetsDictionary[sheetId]).Key != null ?
                        //    Array.IndexOf(sheetsDictionary.Keys.ToArray(), sheetsDictionary.FirstOrDefault(x => x.Value == sheetsDictionary[sheetId]).Key) :
                        //    -1;
                    }

                    //var columnSheet = columns.Skip(indexSheet).FirstOrDefault();
                    //var rowSheet = rows.Skip(indexSheet).FirstOrDefault();

                    var columnSheet = columns.Skip(i).FirstOrDefault();
                    var rowSheet = rows.Skip(i).FirstOrDefault();

                    string dest = Path.Combine(Path.GetDirectoryName(destiny), $"{Path.GetFileNameWithoutExtension(destiny)}__{sheetId}{Path.GetExtension(destiny)}");
                    saveSuccess += SaveDataTable(dtSheet, dest, separator, columnSheet, rowSheet) ? 1 : 0;
                }

                return saveSuccess;
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(Converter), ex), ex);
            }
        }

        internal bool ConvertAllSheets(string? origin, string? destiny, int minRows = 1, string separator = ";")
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
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(ConvertAllSheets), ex), ex);
            }
        }


    }
}
