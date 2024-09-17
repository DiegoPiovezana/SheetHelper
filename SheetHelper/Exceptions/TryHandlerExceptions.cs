using ExcelDataReader.Core;
using ExcelDataReader.Log.Logger;
using SH.Globalization;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace SH.Exceptions
{
    /// <summary>
    /// <para>0: Ignore the exception and no warning will be thrown;</para>
    /// <para>1: Try again.</para>
    /// <para>Or throw exception</para>
    /// </summary>
    internal class TryHandlerExceptions
    {
        private readonly SheetHelper _sheethelper;

        public TryHandlerExceptions(SheetHelper sheetHelper)
        {
            _sheethelper = sheetHelper;
        }





        //internal static int ExceptionManager(SHException exception, List<string>? tryIgnoreExceptions)
        //{
        //    switch (exception.Code)
        //    {




        //        default:
        //            break;
        //    }
        //}



        /// <summary>
        /// <para>0: Ignore the exception and no warning will be thrown;</para>
        /// <para>1: Try again.</para>
        /// <para>Or throw exception</para>
        /// </summary>
        internal int FileExcelInUse(Exception exception, string pathFile, int countOpen, bool fileOrigin)
        {
            //if (_sheethelper.TryIgnoreExceptions == null || !_sheethelper.TryIgnoreExceptions.Contains(exception.Code)) throw exception;

#if !NETFRAMEWORK
            throw exception; // Handle exception only to NETFRAMEWORK
#endif

            countOpen++;

            if (countOpen >= 2) // If it is necessary to force Excel closure (from the 2nd attempt onwards)                
            {
                var result2 = MessageBox.Show(
                   fileOrigin ? Messages.FileOriginInUseAndCloseExcel(pathFile) : Messages.FileDestinationInUseAndCloseExcel(pathFile),
                   Messages.Warning(),
                   MessageBoxButtons.YesNo,
                   MessageBoxIcon.Exclamation);

                if (result2 == DialogResult.Yes)
                {
                    _sheethelper.CloseExcel();
                    System.Threading.Thread.Sleep(1500);
                    return 1;

                } // If No, continue execution below
            }

            var result1 = MessageBox.Show(
                fileOrigin ? Messages.FileOriginInUse(pathFile) : Messages.FileDestinationInUse(pathFile),
                Messages.Warning(),
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Error);

            if (result1 == DialogResult.OK) { return 1; }
            else { return 0; } // If canceled
        }

        /// <summary>
        /// Incomplete headers will be filled.
        /// </summary>
        /// <param name="dataTable"></param>      
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        internal bool HeaderValid(DataTable dataTable)
        {
            DataRow firstRow = dataTable.Rows[0];

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                if (string.IsNullOrEmpty(firstRow[i]?.ToString()))
                {
                    var except = new ColumnNameHeaderInvalidSHException(i);
                    bool ignoreEmptyColumns = _sheethelper.TryIgnoreExceptions != null && _sheethelper.TryIgnoreExceptions.Contains(except.Code);
                    if (!ignoreEmptyColumns)
                    {
                        //dataTable.Columns[i].ColumnName = $"EmptyColumn{i + 1}";
                        dataTable.Rows[0][i]= $"EmptyColumn{i + 1}";
                    }
                    else throw except;
                }
                //else
                //{
                //    dataTable.Columns[i].ColumnName = firstRow[i].ToString();
                //}
            }
            return true;
        }

        //internal int ColumnNotExist(string idColumn, int limitIndexColumn, DataTable dataTable)
        //{
        //    var except = new ColumnOutRangeSHException(idColumn);
        //    if (_sheethelper.TryIgnoreExceptions == null || !_sheethelper.TryIgnoreExceptions.Contains(except.Code)) throw except;

        //    foreach (DataColumn column in dataTable.Columns)
        //    {
        //        if (column)


        //    }



        //    for (int i = 0; i < dataTable.Columns.Count; i++)
        //    {
        //        if (string.IsNullOrEmpty(firstRow[i]?.ToString()))
        //        {
        //            if (ignoreEmptyColumns) { dataTable.Columns[i].ColumnName = $"EmptyColumn{i + 1}"; }
        //            else throw new Exception($"E-4041-SH: Column header '{i + 1}' (column name) is not valid because it is blank!");
        //        }
        //        else
        //        {
        //            dataTable.Columns[i].ColumnName = firstRow[i].ToString();
        //        }
        //    }
        //}



    }
}
