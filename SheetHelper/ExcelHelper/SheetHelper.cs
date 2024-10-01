using SH.ExcelHelper.Tools;
using SH.ExcelHelper.Treatments;
using SH.Exceptions;
using SH.Globalization;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace SH
{
    /// <summary>
    /// Fast and lightweight library for easy read and conversion of large Excel files
    /// </summary>
    public class SheetHelper : ISheetHelper
    {
        /// <summary>
        /// Represents the conversion progress. If 100%, the conversion is fully completed.
        /// </summary>    
        public int Progress { get; internal set; }

        /// <summary>
        /// (Optional) The dictionary can specify characters that should not be maintained after conversion (line breaks, for example) and which replacements should be performed in each case.
        /// </summary>
        public Dictionary<string, string>? ProhibitedItems { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// (Optional) Ignored exceptions will attempt to be handled internally. If it is not possible, it will just return false and the exception will not be thrown.
        /// <para>By default, it will ignore the exception when the source or destination file is in use. If .NET Framework will display a warning to close the file, otherwise it will return false.</para>
        /// </summary>
        public List<string> TryIgnoreExceptions { get; set; } = new List<string>() { "E-0001-SH", "E-4041-SH", "E-4049-SH" };


        private readonly Features _features;
        private readonly Validations _validations;
        private readonly Definitions _definitions;

        /// <summary>
        /// Fast and lightweight library for easy read and conversion of large Excel files
        /// </summary>
        public SheetHelper()
        {
            _features = new(this);
            _validations = new(this);
            _definitions = new(this, _validations);
        }



        /// <inheritdoc/>     
        public void CloseExcel(string? filterTitle = null)
        {
            try
            {
                _features.CloseExcel(filterTitle);
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

        /// <inheritdoc/>  
        public int GetIndexColumn(string columnName)
        {
            try
            {
                _validations.ValidateStringNullOrEmpty(columnName, nameof(columnName), nameof(GetIndexColumn));
                return _features.GetIndexColumn(columnName);
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

        /// <inheritdoc/> 
        public string GetNameColumn(int columnIndex)
        {
            try
            {
                _validations.ValidateIntMin(columnIndex, nameof(columnIndex), nameof(GetNameColumn), 1);
                return _features.GetNameColumn(columnIndex);
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

        /// <inheritdoc/> 
        public string? UnGZ(string gzFile, string pathDestination)
        {
            try
            {
                _validations.ValidateFileExists(gzFile, nameof(gzFile), nameof(UnGZ));
                _validations.ValidateDestinationFolder(pathDestination, nameof(pathDestination), nameof(UnGZ));
                return _features.UnGZ(gzFile, pathDestination);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(UnGZ), ex), ex);
            }
            finally
            {
                if (Directory.Exists(@".\SheetHelper")) Directory.Delete(@".\SheetHelper", true);
            }
        }

        /// <inheritdoc/> 
        public string? UnZIP(string? zipFile, string pathDestination)
        {
            try
            {
                _validations.ValidateFileExists(zipFile, nameof(zipFile), nameof(UnZIP));
                _validations.ValidateDestinationFolder(pathDestination, nameof(pathDestination), nameof(UnZIP));
                return _features.UnZIP(zipFile, pathDestination);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(UnZIP), ex), ex);
            }
            finally
            {
                if (Directory.Exists(@".\SheetHelper")) Directory.Delete(@".\SheetHelper", true);
            }
        }

        /// <inheritdoc/> 
        public string? UnzipAuto(string? zipFile, string pathDestination, bool mandatory = true)
        {
            try
            {
                _validations.ValidateFileExists(zipFile, nameof(zipFile), _validations.GetCallingMethodName(1));
                _validations.ValidateDestinationFolder(pathDestination, nameof(pathDestination), nameof(UnzipAuto));
                return _features.UnzipAuto(zipFile, pathDestination, mandatory);
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

        /// <inheritdoc/> 
        public DataRow ConvertToDataRow(string[] row, DataTable table)
        {
            try
            {
                _validations.ValidateArgumentNull(row, nameof(row), nameof(ConvertToDataRow));
                _validations.ValidateArgumentNull(table, nameof(table), nameof(ConvertToDataRow));
                return _features.ConvertToDataRow(row, table);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(ConvertToDataRow), ex), ex);
            }
            finally
            {
                if (Directory.Exists(@".\SheetHelper")) Directory.Delete(@".\SheetHelper", true);
            }
        }

        /// <inheritdoc/> 
        public string[] GetRowArray(DataTable table, bool header = true, int indexRow = 0)
        {
            try
            {
                _validations.ValidateArgumentNull(table, nameof(table), nameof(GetRowArray));
                _validations.ValidateIntMin(indexRow, nameof(indexRow), nameof(GetRowArray));
                return _features.GetRowArray(table, header, indexRow);
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

        /// <inheritdoc/> 
        public Dictionary<string, DataTable> GetAllSheets(string filePath, int minQtdRows = 0, bool formatName = false)
        {
            try
            {
                _validations.ValidateFileExists(filePath, nameof(filePath), nameof(GetAllSheets));
                _validations.ValidateIntMin(minQtdRows, nameof(minQtdRows), nameof(GetAllSheets));
                return _features.GetAllSheets(filePath, minQtdRows, formatName);
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

        /// <inheritdoc/> 
        public string NormalizeText(string? text, char replaceSpace = '_', bool toLower = true)
        {
            try
            {
                return _features.NormalizeText(text, replaceSpace, toLower);
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

        /// <inheritdoc/> 
        public string FixItems(string items)
        {
            try
            {
                return _features.FixItems(items);
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

        /// <inheritdoc/> 
        public Dictionary<string, string>? GetDictionaryJson(string jsonTextItems)
        {
            try
            {
                _validations.ValidateStringNullOrEmpty(jsonTextItems, nameof(jsonTextItems), _validations.GetCallingMethodName(1));
                return _features.GetDictionaryJson(jsonTextItems);
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

        /// <inheritdoc/>
        public string GetJsonDictionary(Dictionary<string, string> dictionary)
        {
            try
            {
                return GetJsonDictionary(dictionary);
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

        /// <inheritdoc/>
        public DataSet GetDataSet(string? origin)
        {
            try
            {
                _validations.ValidateOriginFile(origin, nameof(origin), nameof(GetDataTable));
                return _features.GetDataSet(origin);
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

        /// <inheritdoc/>
        public DataTable? GetDataTable(string origin, string sheet = "1")
        {
            try
            {
                _validations.ValidateOriginFile(origin, nameof(origin), nameof(GetDataTable));
                _validations.ValidateSheetIdInput(sheet);
                return _features.GetDataTable(origin, sheet);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(GetDataTable), ex), ex);
            }
        }

        /// <inheritdoc/>
        public bool SaveDataTable(DataTable dataTable, string destination, string delimiter = ";", string? columns = null, string? rows = null)
        {
            try
            {
                _validations.ValidateSaveDataTable(destination, delimiter, columns, rows, nameof(SaveDataTable));
                return _features.SaveDataTable(dataTable, destination, delimiter, columns, rows);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(SaveDataTable), ex), ex);
            }
            finally
            {
                if (Directory.Exists(@".\SheetHelper")) Directory.Delete(@".\SheetHelper", true);
            }
        }

        /// <inheritdoc/>
        public bool Converter(string? origin, string? destination, string? sheet, string? delimiter, string? columns, string? rows, int minRows = 1)
        {
            try
            {
                _validations.ValidateOneConverterAsync(origin, destination, sheet, delimiter, columns, rows, nameof(Converter)).GetAwaiter().GetResult();
                return _features.Converter(origin, destination, sheet, delimiter, columns, rows, minRows);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(Converter), ex), ex);
            }
            finally
            {
                if (Directory.Exists(@".\SheetHelper")) Directory.Delete(@".\SheetHelper", true);
            }
        }

        /// <inheritdoc/> 
        public int Converter(string? origin, object? destinations, object? sheets, object? delimiters, object? columns, object? rows, int minRows = 1)
        {
            try
            {
                _definitions.DefineMultiplesInputsConverter(ref destinations, ref sheets, ref delimiters, ref columns, ref rows);
                _validations.ValidateConverter(origin, destinations, sheets, delimiters, columns, rows, nameof(Converter));
                return _features.Converter(origin, destinations, sheets, delimiters, columns, rows, minRows);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(Converter), ex), ex);
            }
            finally
            {
                if (Directory.Exists(@".\SheetHelper")) Directory.Delete(@".\SheetHelper", true);
            }
        }

        /// <inheritdoc/> 
        public bool ConvertAllSheets(string? origin, string? destination, int minRows = 1, string delimiter = ";")
        {
            try
            {
                _validations.ValidateFileExists(origin, nameof(origin), nameof(Converter));
                _validations.ValidateDestinationFile(destination, nameof(Converter));
                _validations.ValidateStringNullOrEmpty(delimiter, nameof(delimiter), nameof(Converter));
                return _features.ConvertAllSheets(origin, destination, minRows, delimiter);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(ConvertAllSheets), ex), ex);
            }
            finally
            {
                if (Directory.Exists(@".\SheetHelper")) Directory.Delete(@".\SheetHelper", true);
            }
        }


    }
}
