﻿using SH.ExcelHelper.Tools;
using SH.ExcelHelper.Treatments;
using SH.Exceptions;
using SH.Globalization;
using System;
using System.Collections.Generic;
using System.Data;

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
        public List<string>? TryIgnoreExceptions { get; set; } = new List<string>() { "E-0001-SH" };


        private readonly Features _features;
        private readonly Validations _validations;

        /// <summary>
        /// Fast and lightweight library for easy read and conversion of large Excel files
        /// </summary>
        public SheetHelper()
        {
            _features = new(this);
            _validations = new(this);
        }



        /// <inheritdoc/>     
        public void CloseExcel()
        {
            try
            {
                _features.CloseExcel();
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
        public string? UnGZ(string gzFile, string pathDestiny)
        {
            try
            {
                _validations.ValidateFile(gzFile, nameof(gzFile), nameof(UnGZ));
                _validations.ValidateDestinyFolder(pathDestiny, true, nameof(pathDestiny), nameof(UnGZ));
                return _features.UnGZ(gzFile, pathDestiny);
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

        /// <inheritdoc/> 
        public string? UnZIP(string? zipFile, string pathDestiny)
        {
            try
            {
                _validations.ValidateFile(zipFile, nameof(zipFile), nameof(UnZIP));
                _validations.ValidateDestinyFolder(pathDestiny, true, nameof(pathDestiny), nameof(UnZIP));
                return _features.UnZIP(zipFile, pathDestiny);
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

        /// <inheritdoc/> 
        public string? UnzipAuto(string? zipFile, string pathDestiny, bool mandatory = true)
        {
            try
            {
                _validations.ValidateFile(zipFile, nameof(zipFile), _validations.GetCallingMethodName(1));
                _validations.ValidateDestinyFolder(pathDestiny, true, nameof(pathDestiny), nameof(UnzipAuto));
                return _features.UnzipAuto(zipFile, pathDestiny, mandatory);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(UnzipAuto), ex), ex);
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
                _validations.ValidateFile(filePath, nameof(filePath), nameof(GetAllSheets));
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
                _validations.ValidateSheetId(sheet);
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
        public bool SaveDataTable(DataTable dataTable, string destiny, string separator = ";", string? columns = null, string? rows = null)
        {
            try
            {
                _validations.Validate(destiny, separator, columns, rows, nameof(SaveDataTable));
                return _features.SaveDataTable(dataTable, destiny, separator, columns, rows);
            }
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(SaveDataTable), ex), ex);
            }
        }

        /// <inheritdoc/> 
        public bool Converter(string? origin, string? destiny, string sheet, string separator, string? columns, string? rows, int minRows = 1)
        {
            try
            {
                _validations.Validate(origin, destiny, sheet, separator, columns, rows, nameof(Converter));
                return _features.Converter(origin, destiny, sheet, separator, columns, rows, minRows);
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

        /// <inheritdoc/> 
        public int Converter(string? origin, string? destiny, ICollection<string>? sheets, string separator = ";", ICollection<string>? columns = default, ICollection<string>? rows = default, int minRows = 1)
        {
            try
            {
                _validations.Validate(origin, destiny, sheets, separator, columns, rows, nameof(Converter));               
                return _features.Converter(origin, destiny, sheets, separator, columns, rows, minRows);
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

        /// <inheritdoc/> 
        public bool ConvertAllSheets(string? origin, string? destiny, int minRows = 1, string separator = ";")
        {
            try
            {
                _validations.ValidateFile(origin, nameof(origin), nameof(Converter));
                _validations.ValidateDestinyFile(destiny, nameof(Converter));
                _validations.ValidateStringNullOrEmpty(separator, nameof(separator), nameof(Converter));
                return _features.ConvertAllSheets(origin, destiny, minRows, separator);
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
