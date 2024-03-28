﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace SH.ExcelHelper.Treatments
{

    /// <summary>
    /// Used to process the information provided by the user
    /// </summary>
    internal static class Definitions
    {

        /// <summary>
        /// Receives rows as a string and returns an array of integers with the first and last row.
        /// </summary>
        internal static int[] DefineRows(string rows, DataTable table)
        {
            int limitIndexRows = table.Rows.Count + 1; // Add 1 to consider header
            List<int> indexRows = new();

            if (string.IsNullOrEmpty(rows) || string.IsNullOrEmpty(rows.Trim())) // If rows not specified
            {
                indexRows.AddRange(Enumerable.Range(1, limitIndexRows)); // Convert all rows                
                return indexRows.ToArray();
            }

            rows = SheetHelper.FixItems(rows); //"1:23,34:-56,23:1,70,75,-1"


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
        internal static int[] DefineColumnsASCII(string columns, DataTable table)
        {
            int limitIndexColumn = table.Columns.Count;
            List<int> indexColumns = new();

            if (string.IsNullOrEmpty(columns) || string.IsNullOrEmpty(columns.Trim())) // If columns not specified                
                return new int[] { 0 }; // Behavior to convert all columns

            columns = SheetHelper.FixItems(columns);
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

                if (column.All(c => char.IsLetter(c))) indexColumn = SheetHelper.GetIndexColumn(column);
                else indexColumn = Convert.ToInt32(column);

                if (indexColumn >= 0)  // "75"
                {
                    if (indexColumn == 0 || indexColumn > limitIndexColumn)
                        throw new Exception($"E-0000-SH: The column '{column}' is out of range (min 1, max {limitIndexColumn})!");

                    return indexColumn;
                }
                else // "-2"
                {
                    if (limitIndexColumn + indexColumn + 1 > limitIndexColumn)
                        throw new Exception($"E-0000-SH: The column '{column}' is out of range, because it refers to column '{limitIndexColumn + indexColumn + 1}' (min 1, max {limitIndexColumn})!");

                    return limitIndexColumn + indexColumn + 1;
                }
            }
        }

    }
}