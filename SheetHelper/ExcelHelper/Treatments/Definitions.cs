using SH.Exceptions;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace SH.ExcelHelper.Treatments
{

    /// <summary>
    /// Used to process the information provided by the user
    /// </summary>
    internal class Definitions
    {
        readonly SheetHelper _sheetHelper;
        readonly Validations _validations;

        public Definitions(SheetHelper sheetHelper, Validations validations)
        {
            _sheetHelper = sheetHelper;            
            _validations = validations;
        }


        /// <summary>
        /// Receives rows as a string and returns an array of integers with the first and last row.
        /// </summary>
        internal int[] DefineRows(string rows, DataTable table)
        {
            int limitIndexRows = table.Rows.Count + 1; // Add 1 to consider header
            List<int> indexRows = new();

            if (string.IsNullOrEmpty(rows) || string.IsNullOrEmpty(rows.Trim())) // If rows not specified
            {
                indexRows.AddRange(Enumerable.Range(1, limitIndexRows)); // Convert all rows                
                return indexRows.ToArray();
            }

            rows = _sheetHelper.FixItems(rows); //"1:23,34:-56,23:1,70,75,-1"


            foreach (string row in rows.Split(','))
            {
                if (row.Contains(":")) // "1:23", "34:-56" or "23:1"
                {
                    string[] rowsArray = row.Split(':'); // e.g.: {"A","Z"}

                    if (rowsArray.Length != 2)
                        throw new Exception($"E-0000-SH: Row '{row}' is not a valid pattern!");

                    if (string.IsNullOrEmpty(rowsArray[0])) // If first row not defined
                        rowsArray[0] = "1"; // Then, convert from the first row

                    if (string.IsNullOrEmpty(rowsArray[1])) // If last row not defined
                        rowsArray[1] = limitIndexRows.ToString(); // Then, convert until the last row

                    int firstRowIndex = ConvertIndexRow(rowsArray[0]);
                    int lastRowIndex = ConvertIndexRow(rowsArray[1]);

                    if (firstRowIndex > lastRowIndex)
                        indexRows.AddRange(Enumerable.Range(lastRowIndex, firstRowIndex - lastRowIndex + 1).Reverse());
                    else
                        indexRows.AddRange(Enumerable.Range(firstRowIndex, lastRowIndex - firstRowIndex + 1));
                }
                else // "70", "75" or "-1"
                {
                    indexRows.Add(ConvertIndexRow(row));
                }
            }
            return indexRows.ToArray();


            int ConvertIndexRow(string row)
            {
                if (row.All(c => char.IsLetter(c))) throw new Exception($"E-0000-SH: The row '{row}' is not a valid!");

                int indexRow = Convert.ToInt32(row);

                if (indexRow >= 0)  // "75"
                {
                    if (indexRow == 0 || indexRow > limitIndexRows)
                        throw new Exception($"E-0000-SH: The row '{row}' is out of range (min 1, max {limitIndexRows})!");

                    return indexRow;
                }
                else // "-2"
                {
                    if (limitIndexRows + indexRow + 1 > limitIndexRows)
                        throw new Exception($"E-0000-SH: The row '{row}' is out of range, because it refers to row '{limitIndexRows + indexRow + 1}' (min 1, max {limitIndexRows})!");

                    return limitIndexRows + indexRow + 1;
                }
            }
        }

        /// <summary>
        /// Defines the index of all columns to be converted.
        /// </summary>
        /// <param name="columns">Columns to be converted. e.g.: "B:H"</param>
        /// <param name="table">DataTable a ser analisado.</param>        
        /// <returns>Array with all indexes of the columns to be converted.</returns>
        /// <exception cref="Exception"></exception>
        internal int[] DefineColumnsASCII(string columns, DataTable table)
        {
            int limitIndexColumn = table.Columns.Count;
            List<int> indexColumns = new();

            if (string.IsNullOrEmpty(columns) || string.IsNullOrEmpty(columns.Trim())) // If columns not specified                
                return new int[] { 0 }; // Behavior to convert all columns

            columns = _sheetHelper.FixItems(columns);
            columns = columns.ToUpper();

            // "A:H,4:9,4:-9,B,75,-2"

            foreach (string column in columns.Split(','))
            {
                if (column.Contains(":")) // "A:H" or "4:9" or 4:-9
                {
                    string[] columnsArray = column.Split(':'); // e.g.: {"A","Z"}

                    if (columnsArray.Length != 2)
                        throw new Exception($"E-0000-SH: Column '{column}' is not a valid pattern!");

                    if (string.IsNullOrEmpty(columnsArray[0])) // If first row not defined
                        columnsArray[0] = "1"; // Then, convert from the first row

                    if (string.IsNullOrEmpty(columnsArray[1])) // If last row not defined
                        columnsArray[1] = limitIndexColumn.ToString(); // Then, convert until the last row

                    int firstColumnIndex = ConvertIndexColumn(columnsArray[0]);
                    int lastColumnIndex = ConvertIndexColumn(columnsArray[1]);

                    if (lastColumnIndex.Equals(limitIndexColumn) && firstColumnIndex.Equals(1))
                        return new int[] { 0 }; // Behavior to convert all columns

                    if (firstColumnIndex > lastColumnIndex)
                        indexColumns.AddRange(Enumerable.Range(lastColumnIndex, firstColumnIndex - lastColumnIndex + 1).Reverse());
                    else
                        indexColumns.AddRange(Enumerable.Range(firstColumnIndex, lastColumnIndex - firstColumnIndex + 1));
                }
                else // "B", "75" or "-2"
                {
                    indexColumns.Add(ConvertIndexColumn(column));
                }
            }
            return indexColumns.ToArray();


            int ConvertIndexColumn(string column)
            {
                int indexColumn;

                if (column.All(c => char.IsLetter(c))) indexColumn = _sheetHelper.GetIndexColumn(column);
                else indexColumn = Convert.ToInt32(column);

                if (indexColumn >= 0)  // "75"
                {
                    //if (indexColumn == 0 || indexColumn > limitIndexColumn)
                    //{
                    //    switch (_tryHandlerExceptions.ColumnNotExist(ex, pathOrigin, countOpen, true))
                    //    {
                    //        case 0: return;
                    //        case 1: goto again;
                    //        default: throw new RowOutRangeSHException(column, limitIndexColumn); 
                    //    }                        
                    //}

                    _validations.ValidateColumnOutffRange(indexColumn, limitIndexColumn);
                    return indexColumn;
                }
                else // "-2"
                {
                    //if (limitIndexColumn + indexColumn + 1 > limitIndexColumn)
                    //{

                    //    throw new Exception($"E-4042-SH: The column '{column}' is out of range, because it refers to column '{limitIndexColumn + indexColumn + 1}' (min 1, max {limitIndexColumn})!");
                    //}                        

                    _validations.ValidateColumnRefOutffRange(indexColumn, limitIndexColumn, column);
                    return limitIndexColumn + indexColumn + 1;
                }
            }
        }

        internal void DefineSheets(ref ICollection<string?> sheets, Dictionary<string, DataTable> sheetsDictionary)
        {
            if (string.IsNullOrEmpty(sheets.FirstOrDefault()) && sheets.Count == 1)
            {
                sheets = sheetsDictionary.Keys;
            }
            //else if (sheets is ICollection<string> sheetsCollection)
            //{
            //    if (sheetsCollection.Count == 0)
            //    {
            //        sheets = Enumerable.Range(1, countSheets).Select(i => i.ToString()).ToList();
            //    }
            //}
            //else
            //{
            //    throw new ArgumentException("The sheets parameter must be a string or a collection of strings.");
            //}
        }

        internal void DefineDestinations(ref ICollection<string?> destinationsCollection, ICollection<string?> sheetsCollection)
        {
            if (destinationsCollection.Count > 1)
            {
                Dictionary<string, int> counts = new();
                List<string?> result = new();
                var sheetsCollect = new List<string?>(sheetsCollection);

                foreach (var dest in destinationsCollection)
                {
                    string directory = Path.GetDirectoryName(dest);
                    string fileName = Path.GetFileNameWithoutExtension(dest);
                    string extension = Path.GetExtension(dest);

                    string item = $"{Path.Combine(directory, fileName)}__{sheetsCollect[counts.Count]}";

                    if (counts.ContainsKey(item))
                    {
                        counts[item]++;
                        result.Add($"{item}_{counts[item]}{extension}");
                    }
                    else
                    {
                        counts[item] = 1;
                        result.Add($"{item}{extension}");
                    }
                }

                destinationsCollection = result;
            }
        }

        internal int DefineMultiplesInputsConverter(ref object destinations, ref object sheets, ref object separators, ref object columns, ref object rows)
        {
            /* POSSIBILITIES:
                 * origin = 1;
                 * destinations = 1 or X;
                 * sheets = 1 or X;
                 * separators = 1 or X;
                 * columns = 1/null or X;
                 * rows = 1/null or X;
                 * minRows = 1 or X.                
                 */

            ICollection<string?> destinationsCollection = TryConvertToCollection(destinations);
            ICollection<string?> sheetsCollection = TryConvertToCollection(sheets);
            ICollection<string?> separatorsCollection = TryConvertToCollection(separators);
            ICollection<string?> columnsCollection = TryConvertToCollection(columns);
            ICollection<string?> rowsCollection = TryConvertToCollection(rows);

            var collections = new List<ICollection<string?>> { destinationsCollection, sheetsCollection, separatorsCollection, columnsCollection, rowsCollection };
            int countConversions = collections.Max(collection => collection.Count);

            destinationsCollection = ExpandCollection(destinationsCollection, countConversions);
            sheetsCollection = ExpandCollection(sheetsCollection, countConversions);
            separatorsCollection = ExpandCollection(separatorsCollection, countConversions);
            columnsCollection = ExpandCollection(columnsCollection, countConversions);
            rowsCollection = ExpandCollection(rowsCollection, countConversions);

            destinations = destinationsCollection;
            sheets = sheetsCollection;
            separators = separatorsCollection;
            columns = columnsCollection;
            rows = rowsCollection;

            return countConversions;

            static ICollection<string?> TryConvertToCollection(object? obj)
            {
                if (obj is null) return new List<string?>() { "" };

                return obj switch
                {
                    ICollection<string?> collection => collection,
                    //string str => new ReadOnlyCollection<string?>(new List<string?> { str }),
                    string str => new List<string?> { str },
                    _ => throw new ArgumentException("Unsupported input type."),
                };
            }

            static ICollection<string?> ExpandCollection(ICollection<string?> collection, int targetCount)
            {
                if (collection.Count == 1 && targetCount > 1)
                {
                    string? item = collection.FirstOrDefault();
                    return Enumerable.Repeat(item, targetCount).ToList();
                }
                else if (collection.Count == 0)
                {
                    return Enumerable.Repeat<string?>(null, targetCount).ToList();
                }
                else
                {
                    return collection;
                }
            }

        }

        internal DataTable DefineFirstRowToHeader(DataTable dataTable, string extension, bool ignoreEmptyColumns)
        {
            static bool IsCsvTxtRptExtension(string extension)
            {
                string[] allowedExtensions = { ".csv", ".txt", ".rpt" };
                return allowedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase);
            }

            if (IsCsvTxtRptExtension(extension))
            {
                DataRow firstRow = dataTable.Rows[0];

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    if (string.IsNullOrEmpty(firstRow[i]?.ToString()))
                    {
                        if (ignoreEmptyColumns) { dataTable.Columns[i].ColumnName = $"EmptyColumn{i+1}"; }
                        else throw new Exception($"E-4041-SH: Column header '{i+1}' (column name) is not valid because it is blank!");
                    }
                    else
                    {
                        dataTable.Columns[i].ColumnName = firstRow[i].ToString();
                    }                    
                }

                dataTable.Rows.RemoveAt(0);
            }

            return dataTable;
        }
    }
}
