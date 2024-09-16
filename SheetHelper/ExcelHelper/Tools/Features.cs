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

namespace SH.ExcelHelper.Tools
{
    internal class Features : ISheetHelper
    {
        private readonly SheetHelper _sheetHelper;
        private readonly Validations _validations;
        private readonly Reading _reading;
        private readonly Writing _writing;
        private readonly Definitions _definitions;

        public Features(SheetHelper sheetHelper)
        {
            _sheetHelper = sheetHelper;
            _validations = new Validations(sheetHelper);
            _reading = new Reading();
            _writing = new Writing(sheetHelper);
            _definitions = new Definitions(sheetHelper, _validations);
        }



        public void CloseExcel()
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

        public int GetIndexColumn(string? columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName)) throw new ArgumentException("Column name cannot be null or empty.");

            int sum = 0;
            foreach (var character in columnName)
            {
                if (character < 'A' || character > 'Z') throw new ArgumentException("Invalid character in column name.");
                sum = sum * 26 + (character - 'A' + 1);
            }

            return sum;
        }


        public string GetNameColumn(int columnIndex)
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

        public string? UnGZ(string gzFile, string pathDestination)
        {
            try
            {
                using var compressedFileStream = File.Open(gzFile, FileMode.Open, FileAccess.Read);
                string fileConverted;

                // If the format to be converted is not specified, try to get it from the file name
                if (string.IsNullOrEmpty(Path.GetExtension(pathDestination)))
                {
                    string originalFileName = Path.GetFileName(compressedFileStream.Name).Replace(".gz", "").Replace(".GZ", "");
                    string formatOriginal = Regex.Match(Path.GetExtension(originalFileName), @"\.[A-Za-z]*").Value;
                    fileConverted = $"{pathDestination}{Path.GetFileNameWithoutExtension(originalFileName)}{formatOriginal}";
                }
                else
                {
                    fileConverted = pathDestination;
                }

                if(!Directory.Exists(pathDestination)) Directory.CreateDirectory(pathDestination);
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

        public string? UnZIP(string? zipFile, string pathDestination)
        {
            try
            {
                string directoryZIP = Path.Combine(pathDestination, "CnvrtdZIP");
                if (!Directory.Exists(directoryZIP)) Directory.CreateDirectory(directoryZIP);

                ZipFile.ExtractToDirectory(zipFile, directoryZIP);

                string fileLocation = Directory.EnumerateFiles(directoryZIP).First();
                string fileDestination = Path.Combine(pathDestination, Path.GetFileName(fileLocation));

                if (File.Exists(fileDestination)) File.Delete(fileDestination);
                File.Move(fileLocation, fileDestination);
                Directory.Delete(directoryZIP, true);

                return fileDestination;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public string? UnzipAuto(string? zipFile, string pathDestination, bool mandatory = true)
        {
            try
            {

            restart:

                _validations.ValidateFile(zipFile, nameof(zipFile), _validations.GetCallingMethodName(1));

                switch (Path.GetExtension(zipFile).ToLower())
                {
                    case ".gz":
                        zipFile = UnGZ(zipFile, pathDestination);
                        goto restart;

                    case ".zip":
                        //stream.Close();
                        zipFile = UnZIP(zipFile, pathDestination);
                        goto restart;

                    default:
                        if (mandatory) throw new UnableUnzipSHException(zipFile);
                        else return zipFile;
                }
            }
            catch (Exception)
            {
                throw;
            }
            //finally
            //{
            //    if (Directory.Exists(@".\SheetHelper")) Directory.Delete(@".\SheetHelper", true);
            //}
        }

        public DataRow ConvertToDataRow(string[] row, DataTable table)
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
            catch (Exception)
            {
                throw;
            }
        }

        public string[] GetRowArray(DataTable table, bool header = true, int indexRow = 0)
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
            catch (Exception)
            {
                throw;
            }
        }

        public Dictionary<string, DataTable> GetAllSheets(string filePath, int minQtdRows = 0, bool formatName = false)
        {
            try
            {
                var dataSet = GetDataSet(filePath);

                if (dataSet.Tables.Count == 0)
                {
                    throw new Exception("E-0000-SH: No sheets found in the file.");
                }

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
            catch (Exception)
            {
                throw;
            }
        }

        public string NormalizeText(string? text, char replaceSpace = '_', bool toLower = true)
        {
            try
            {
                if (string.IsNullOrEmpty(text?.Trim())) return "";

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

        public string FixItems(string items)
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
            catch (Exception)
            {
                throw;
            }
        }

        public Dictionary<string, string>? GetDictionaryJson(string jsonTextItems)
        {
            try
            {
                //if (string.IsNullOrEmpty(jsonTextItems))
                //    throw new ParamExceptionSHException(nameof(jsonTextItems), nameof(GetDictionaryJson));

                return JsonSerializer.Deserialize<Dictionary<string, string>>(jsonTextItems);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public string GetJsonDictionary(Dictionary<string, string> dictionary)
        {
            try
            {
                if (dictionary == null || dictionary.Count == 0)
                    throw new ArgumentNullOrEmptySHException(nameof(dictionary), nameof(GetJsonDictionary));

                return JsonSerializer.Serialize(dictionary);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public DataSet GetDataSet(string? origin)
        {
            try
            {
                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
                _validations.ValidateOriginFile(origin, nameof(origin), nameof(GetDataTable));

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                _sheetHelper.Progress += 5; // 5 

                DataSet result = _reading.GetDataSet(origin);
                _sheetHelper.Progress += 25; // 35 (after reading the file)

                return result;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public DataTable? GetDataTable(string origin, string sheet = "1")
        {
            try
            {
                var result = GetDataSet(origin); // 35 (after reading the file)

                // Get the sheet to be converted
                DataTable table = _reading.GetTableByDataSet(sheet, result);

                // Handling to allow header consideration (XLS case)
                // TODO: Refactor              
                table = _definitions.DefineFirstRowToHeader(table, Path.GetExtension(origin), _sheetHelper.TryIgnoreExceptions.Contains("E-4041-SH"));
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

            catch (Exception)
            {
                throw;
            }
        }

        public bool SaveDataTable(DataTable dataTable, string destination, string separator = ";", string? columns = null, string? rows = null)
        {

            try
            {
                return _writing.SaveDataTable(dataTable, destination, separator, columns, rows);
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
            //                    $"O arquivo '{Path.GetFileName(destination)}' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida confirme para continuar.",
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

        public bool Converter(string origin, string destination, string sheet, string separator, string columns, string rows, int minRows = 1)
        {
            try
            {
                _sheetHelper.Progress = 5;

                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
                _validations.ValidateFile(origin, nameof(origin), nameof(Converter));

                if (!_validations.CheckConvertNecessary(origin, destination, sheet, separator, columns, rows))
                {
                    _sheetHelper.Progress = 100;
                    File.Copy(origin, destination, true);
                    //if (Directory.Exists(@".\SheetHelper\")) Directory.Delete(@".\SheetHelper\", true);
                    return true;
                }

                DataTable? table = GetDataTable(origin, sheet);
                _validations.ValidateRowsMinDt(table, minRows, nameof(Converter));

                return SaveDataTable(table, destination, separator, columns, rows);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public int Converter(string origin, object destinations, object sheets, object separators, object columns, object rows, int minRows)
        {
            try
            {
                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
                _validations.ValidateFile(origin, nameof(origin), nameof(Converter));

                var destinationsCollection = destinations as ICollection<string?>;
                var sheetsCollection = sheets as ICollection<string?>;
                var separatorsCollection = separators as ICollection<string?>;
                var columnsCollection = columns as ICollection<string?>;
                var rowsCollection = rows as ICollection<string?>;

                Dictionary<string, DataTable>? sheetsDictionary = GetAllSheets(origin, minRows, true);
                _validations.ValidateSheetsDictionary(sheetsDictionary);

                //if (sheets == null || sheets.Count == 0) sheets = sheetsDictionary.Keys;
                //if (columns == null || columns.Count == 0) columns = Enumerable.Repeat("", sheets.Count).ToList();
                //if (rows == null || rows.Count == 0) rows = Enumerable.Repeat("", sheets.Count).ToList();


                _definitions.DefineSheets(ref sheetsCollection, sheetsDictionary);
                _definitions.DefineDestinations(ref destinationsCollection, sheetsCollection);

                //int count = sheetsCollection.Count;
                //if (count == destinationsCollection.Count && count == separatorsCollection.Count && count == columnsCollection.Count && count == rowsCollection.Count)
                //{
                //    throw new Exception("E-0000-SH: The number of sheets, columns and rows must be the same.");
                //}

                _validations.ValidateConverter(origin, destinationsCollection, sheetsCollection, separatorsCollection, columnsCollection, rowsCollection, nameof(Converter));

                int saveSuccess = default;

                for (int i = 0; i < sheetsCollection.Count; i++) // Name or index of the sheet              
                {
                    var sheetId = NormalizeText(sheetsCollection.Skip(i).FirstOrDefault());

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

                    var columnSheet = columnsCollection.Skip(i).FirstOrDefault();
                    var rowSheet = rowsCollection.Skip(i).FirstOrDefault();
                    var destination = destinationsCollection.Skip(i).FirstOrDefault();
                    var separator = separatorsCollection.Skip(i).FirstOrDefault();

                    //string dest = Path.Combine(Path.GetDirectoryName(destination), $"{Path.GetFileNameWithoutExtension(destination)}__{sheetId}{Path.GetExtension(destination)}");
                    saveSuccess += SaveDataTable(dtSheet, destination, separator, columnSheet, rowSheet) ? 1 : 0;
                }

                return saveSuccess;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public bool ConvertAllSheets(string? origin, string? destination, int minRows = 1, string separator = ";")
        {
            try
            {
                origin = UnzipAuto(origin, @".\SheetHelper\Extractions\", false);
                _validations.ValidateFile(origin, nameof(origin), nameof(Converter));

                foreach (var sheet in GetAllSheets(origin, minRows, true))
                {
                    string dest = Path.Combine(Path.GetDirectoryName(destination), $"{Path.GetFileNameWithoutExtension(destination)}__{sheet.Key}{Path.GetExtension(destination)}");
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
