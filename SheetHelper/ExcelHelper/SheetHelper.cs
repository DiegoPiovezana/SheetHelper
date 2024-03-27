﻿using SH.Exceptions;
using SH.Globalization;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.Remoting.Messaging;
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
        /// Represents the conversion progress. If 100%, the conversion is fully completed.
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
                Features.CloseExcel();
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
                return Features.GetIndexColumn(columnName);
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
                return Features.GetNameColumn(columnIndex);
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
               return Features.UnGZ(zipFile, pathDestiny);
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
               return Features.UnZIP(zipFile, pathDestiny);
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
                return Features.UnzipAuto(zipFile, pathDestiny, mandatory);
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
            try
            {
                return Features.ConvertToDataRow(row, table);
            }
            catch (Exception)
            {
                throw;
            }
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
            try
            {
                return Features.GetRowArray(table, header, indexRow);
            }
            catch (Exception)
            {
                throw;
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
                return Features.GetAllSheets(filePath, minQtdRows, formatName);
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
                return Features.NormalizeText(text, replaceSpace, toLower);
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
            try
            {
                return Features.FixItems(items);
            }
            catch (Exception)
            {
                throw;
            }
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
                return Features.GetDictionaryJson(jsonTextItems);
            }
            catch (FileOriginInUse)
            {
                TryHandlerExceptions.TryTreatEx_FileOriginInUse(origin, countOpen);
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
               return GetJsonDictionary(dictionary);
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
                return Features.GetDataSet(origin);              
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
               return Features.GetDataTable(origin, sheet);
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
            try
            {
                Features.SaveDataTable(dataTable, destiny, separator, columns, rows);
            }            
            catch (SHException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception(Messages.UnmappedException(nameof(SaveDataTable), ex), ex);
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
                return Features.Converter(origin, destiny, sheet, separator, columns, rows, minRows);
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
                return Features.Converter(origin, destiny, sheets, separator, columns, rows, minRows);
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
                return Features.ConvertAllSheets(origin, destiny, minRows, separator);
            }
            catch (Exception)
            {
                throw;
            }
        }


    }
}
