﻿using System;
using System.Collections.Generic;
using System.Data;

namespace SH.ExcelHelper.Tools
{
    interface ISheetHelper
    {
        /// <summary>
        /// Terminates Excel processes. If a title keyword is provided, only Excel processes with window titles containing the specified keyword will be terminated.
        /// </summary>
        /// <param name="filterTitle">Optional keyword to filter Excel windows by title. If null or empty, all Excel processes will be terminated.</param>
        void CloseExcel(string? filterTitle = null);


        /// <summary>
        /// Receives the column name and returns the index in the worksheet
        /// </summary>
        /// <param name="columnName">Column name. E.g.: "A"</param>
        /// <returns>Index. E.g.: "A" = 1</returns>
        int GetIndexColumn(string columnName);


        /// <summary>
        /// Gets the column name by index
        /// </summary>
        /// <param name="columnIndex"> Column index</param>
        /// <returns>Column name (e.g.: "AB")</returns>
        string GetNameColumn(int columnIndex);


        /// <summary>
        /// Unpacks a .GZ file.
        /// </summary>
        /// <param name="zipFile">The location and name of the compressed file. E.g.: 'C:\\Files\\Report.zip'</param>
        /// <param name="pathDestination">The directory where the uncompressed file will be saved (with or without the destination file name). E.g.: 'C:\\Files\\' or 'C:\\Files\\Converted.xlsx'</param>
        /// <returns>The path of the uncompressed file if successful, otherwise null.</returns>
        public string? UnGZ(string zipFile, string pathDestination);


        /// <summary>
        /// Extracts a .ZIP file.
        /// </summary>
        /// <param name="zipFile">The location and name of the compressed file. E.g.: 'C:\\Files\\Report.zip'</param>
        /// <param name="pathDestination">The directory where the extracted file will be saved. E.g.: 'C:\\Files\\'</param>
        /// <returns>The path of the extracted file.</returns>
        public string? UnZIP(string? zipFile, string pathDestination);


        /// <summary>
        /// Unzip a .zip or .gz file.
        /// <para>Please visit https://bit.ly/SheetHelper_Features to learn more.</para>
        /// </summary>
        /// <param name="zipFile">The location and name of the compressed file. E.g.: 'C:\\Files\\Report.zip'</param>
        /// <param name="pathDestination">The directory where the extracted file will be saved. E.g.: 'C:\\Files\\'</param>
        /// <param name="mandatory">If true, it indicates that the extraction must occur, otherwise, it will show an error. If false, if the conversion does not happen, nothing happens.</param>
        /// <returns>The path of the extracted file.</returns>
        public string? UnzipAuto(string? zipFile, string pathDestination, bool mandatory = true);


        /// <summary>
        /// Converts a string array to a DataRow and returns the resulting DataRow.
        /// </summary>
        /// <param name="row">The string array to be converted.</param>
        /// <param name="table">The target DataTable to which the new DataRow will be added.</param>
        /// <returns>The newly created DataRow populated with values from the string array.</returns>
        public DataRow ConvertToDataRow(string[] row, DataTable table);


        /// <summary>
        /// Retrieves a row of a DataTable.
        /// </summary>       
        /// <param name="table">The DataTable containing the data.</param>
        /// <param name="header">If true, it will get the header (columns name).</param>
        /// <param name="indexRow">index of the line to be obtained (in addition to the header). Enter negative to get just the header.</param>
        /// <returns>An array of strings representing a row of the DataTable.</returns>
        public string[] GetRowArray(DataTable table, bool header = true, int indexRow = 0);


        /// <summary>
        /// Gets the sheets in the workbook with the respective dataTable.
        /// </summary>
        /// <param name="filePath">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx"</param>
        /// <param name="minQtdRows">The minimum number of lines a tab needs to have, otherwise it will be ignored.</param>
        /// <param name="formatName">If true, all spaces and special characters from tab names will be removed.</param>
        /// <returns>Dictionary containing the name of the tabs and the DataTable. If desired, consider using 'sheetDictionary.Values.ToList()' to obtain a list of DataTables.</returns>
        public Dictionary<string, DataTable> GetAllSheets(string filePath, int minQtdRows = 0, bool formatName = false);


        /// <summary>
        /// Normalizes the text by removing accents and spaces.
        /// <para>Example: " Hot Café" => "hot_cafe" </para>
        /// </summary>
        /// <param name="text">Text to be normalized.</param>
        /// <param name="replaceSpace">Character to replace spaces. E.g.: "_"</param>
        /// <param name="toLower">If true, the text will be converted to lowercase.</param>
        /// <returns>Text normalized.</returns>
        public string NormalizeText(string? text, char replaceSpace = '_', bool toLower = true);


        /// <summary>
        /// Fixes a string containing items by replacing line breaks and semicolons with commas,
        /// removing spaces, single quotes, and double quotes, and ensuring proper comma separation.
        /// </summary>
        /// <param name="items">The string containing items to be fixed.</param>
        /// <returns>The fixed string with proper item separation.</returns>
        public string FixItems(string items);


        /// <summary>
        /// Converts a string in JSON format to a dictionary.
        /// <para>Example 1: "{ \"key1\" : \"value1\", \"key2\" : \"value2\", \"key3\" : \"value3\" }"</para> 
        /// <para>Example 2: "{\"\\n\": \" \", \"\\r\": \"\", \";\": \",\"}"</para>        
        /// </summary>
        /// <param name="jsonTextItems">String JSON containing the key-value pairs to be converted.</param>        
        /// <returns>A dictionary containing the extracted key-value pairs from the string.</returns>
        public Dictionary<string, string>? GetDictionaryJson(string jsonTextItems);


        /// <summary>
        /// Serializes a dictionary of strings into a JSON representation.
        /// </summary>
        /// <param name="dictionary">The dictionary to be serialized into JSON.</param>
        /// <returns>A string containing the JSON representation of the provided dictionary.</returns>
        /// <exception cref="ArgumentException">Thrown if the dictionary is null or empty.</exception>
        /// <exception cref="Exception">Thrown if an error occurs while serializing the dictionary to JSON.</exception>
        public string GetJsonDictionary(Dictionary<string, string> dictionary);


        /// <summary>
        /// Reads the file and gets the dataset of worksheet.
        /// <br>Note.: The header is the name of the columns.</br>
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx"</param>       
        /// <returns>DataSet</returns>
        public DataSet GetDataSet(string? origin);


        /// <summary>
        /// Reads the file and gets the datatable of the specified sheet.
        /// <br>Note.: The header is the name of the columns.</br>
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx"</param>
        /// <param name="sheet">Tab of the worksheet to be converted. E.g.: "1" (first sheet) or "TabName"</param>
        /// <returns>DataTable</returns>
        public DataTable? GetDataTable(string origin, string sheet = "1");


        /// <summary>
        /// Performs the conversion of the <paramref name="dataTable"/>, saves in <paramref name="destination"/>. 
        /// </summary>
        /// <param name="dataTable">DataTable to be converted.</param>
        /// <param name="destination">Directory + destination file name + format. E.g.: "C:\\Users\\FileExcel.csv".</param>
        /// <param name="separator">Separator to be used to perform the conversion. E.g.: ";".</param>
        /// <param name="columns">"Enter the columns or their range. E.g.: "A:H, 4:9, 4:-9, B, 75, -2".</param>
        /// <param name="rows">"Enter the rows or their range. E.g.: "1:23, -34:56, 70, 75, -1".</param>
        /// <returns>"true" if converted successfully. "false" if not converted.</returns>
        public bool SaveDataTable(DataTable dataTable, string destination, string separator = ";", string? columns = null, string? rows = null);


        /// <summary>
        /// Performs the conversion of the Excel file located in <paramref name="origin"/>, saves in <paramref name="destination"/>.      
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx."</param>
        /// <param name="destination">Directory + destination file name + format. E.g.: "C:\\Users\\FileExcel.csv."</param>
        /// <param name="sheet">Tab of the worksheet to be converted. E.g.: "1" (first sheet) or "TabName".</param>
        /// <param name="separator">Separator to be used to perform the conversion. E.g.: ";".</param>
        /// <param name="columns">"Enter the columns or their range. E.g.: "A:H, 4:9, 4:-9, B, 75, -2".</param>
        /// <param name="rows">"Enter the rows or their range. E.g.: "1:23, -34:56, 70, 75, -1".</param>
        /// <param name="minRows">(Optional) The minimum number of lines a tab needs to have, otherwise it will be ignored.</param>
        /// <returns>"true" if converted successfully. "false" if not converted.</returns>
        public bool Converter(string? origin, string? destination, string? sheet, string? separator, string? columns, string? rows, int minRows = 1);


        /// <summary>
        /// Converts the desired spreadsheet tabs, considering the destination in the desired format, using the desired separator, columns and rows for each conversion.
        /// <para>NOTE.: Use the Convert or SaveDataTable method for further customizations.</para>
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx."</param>
        /// <param name="destinations">Enter the group of directory + destinations files names + format."</param>
        /// <param name="sheets">Enter the names or indexes of the sheets to be converted.</param>
        /// <param name="separators">The group of separators to be used to perform the conversion. E.g.: ";".</param>
        /// <param name="columns">"Enter the group of columns or their range for each sheet.</param>
        /// <param name="rows">"Enter the group of rows or their range for each sheet.</param>
        /// <param name="minRows">(Optional) The minimum number of lines a tab needs to have, otherwise it will be ignored.</param>
        /// <returns>Number of successful conversions.</returns>
        public int Converter(string? origin, object? destinations, object? sheets, object? separators, object? columns, object? rows, int minRows = 1);


        /// <summary>
        /// Converts all spreadsheet tabs considering all rows and columns.
        /// <para>NOTE.: Use the Convert or SaveDataTable method for further customizations.</para>
        /// </summary>
        /// <param name="origin">Directory + source file name + format. E.g.: "C:\\Users\\FileExcel.xlsx."</param>
        /// <param name="destination">Directory + destination file name + format. E.g.: "C:\\Users\\FileExcel.csv."</param>
        /// <param name="minRows">(Optional) The minimum number of lines a tab needs to have, otherwise it will be ignored.</param>
        /// <param name="separator">Separator to be used to perform the conversion. E.g.: ";".</param>
        /// <returns>True, if success.</returns>
        public bool ConvertAllSheets(string? origin, string? destination, int minRows = 1, string separator = ";");

    }
}
