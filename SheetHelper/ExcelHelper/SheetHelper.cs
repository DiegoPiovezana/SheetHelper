using SH.ExcelHelper.Tools;
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


        /// <summary>
        /// Fast and lightweight library for easy read and conversion of large Excel files
        /// </summary>
        public SheetHelper()
        {
            _features = new(this);
        }



        /// <inheritdoc/>     
        public void CloseExcel()
        {
            _features.CloseExcel();
        }

        /// <inheritdoc/>  
        public int GetIndexColumn(string columnName)
        {
            return _features.GetIndexColumn(columnName);
        }

        /// <inheritdoc/> 
        public string GetNameColumn(int columnIndex)
        {
            return _features.GetNameColumn(columnIndex);
        }

        /// <inheritdoc/> 
        public string? UnGZ(string zipFile, string pathDestiny)
        {
            return _features.UnGZ(zipFile, pathDestiny);
        }

        /// <inheritdoc/> 
        public string? UnZIP(string? zipFile, string pathDestiny)
        {
            return _features.UnZIP(zipFile, pathDestiny);
        }

        /// <inheritdoc/> 
        public string? UnzipAuto(string? zipFile, string pathDestiny, bool mandatory = true)
        {
            return _features.UnzipAuto(zipFile, pathDestiny, mandatory);
        }

        /// <inheritdoc/> 
        public DataRow ConvertToDataRow(string[] row, DataTable table)
        {
            return _features.ConvertToDataRow(row, table);
        }

        /// <inheritdoc/> 
        public string[] GetRowArray(DataTable table, bool header = true, int indexRow = 0)
        {
            return _features.GetRowArray(table, header, indexRow);
        }

        /// <inheritdoc/> 
        public Dictionary<string, DataTable> GetAllSheets(string filePath, int minQtdRows = 0, bool formatName = false)
        {
            return _features.GetAllSheets(filePath, minQtdRows, formatName);
        }

        /// <inheritdoc/> 
        public string NormalizeText(string? text, char replaceSpace = '_', bool toLower = true)
        {
            return _features.NormalizeText(text, replaceSpace, toLower);
        }

        /// <inheritdoc/> 
        public string FixItems(string items)
        {
            return _features.FixItems(items);
        }

        /// <inheritdoc/> 
        public Dictionary<string, string>? GetDictionaryJson(string jsonTextItems)
        {
            return _features.GetDictionaryJson(jsonTextItems);
        }

        /// <inheritdoc/> 
        public string GetJsonDictionary(Dictionary<string, string> dictionary)
        {
            return GetJsonDictionary(dictionary);
        }

        /// <inheritdoc/> 
        public DataSet GetDataSet(string? origin)
        {
            return _features.GetDataSet(origin);
        }

        /// <inheritdoc/> 
        public DataTable? GetDataTable(string origin, string sheet = "1")
        {
            return _features.GetDataTable(origin, sheet);
        }

        /// <inheritdoc/> 
        public bool SaveDataTable(DataTable dataTable, string destiny, string separator = ";", string? columns = null, string? rows = null)
        {
            return _features.SaveDataTable(dataTable, destiny, separator, columns, rows);
        }

        /// <inheritdoc/> 
        public bool Converter(string? origin, string? destiny, string sheet, string separator, string? columns, string? rows, int minRows = 1)
        {
            return _features.Converter(origin, destiny, sheet, separator, columns, rows, minRows);
        }

        /// <inheritdoc/> 
        public int Converter(string? origin, string? destiny, ICollection<string>? sheets, string separator = ";", ICollection<string>? columns = default, ICollection<string>? rows = default, int minRows = 1)
        {
            return _features.Converter(origin, destiny, sheets, separator, columns, rows, minRows);
        }

        /// <inheritdoc/> 
        public bool ConvertAllSheets(string? origin, string? destiny, int minRows = 1, string separator = ";")
        {
            return _features.ConvertAllSheets(origin, destiny, minRows, separator);
        }

    }
}
