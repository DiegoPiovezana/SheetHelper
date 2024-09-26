using ExcelDataReader.Core;
using ExcelDataReader.Log.Logger;
using SH.ExcelHelper.Tools;
using SH.Globalization;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
        //private readonly SheetHelper _sheetHelper;
        private readonly List<string> _ignoreExceptions;

        public TryHandlerExceptions(SheetHelper sheetHelper)
        {
            //_sheetHelper = sheetHelper;
            _ignoreExceptions = sheetHelper.TryIgnoreExceptions;
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
        internal int FileExcelInUse(SHException exception, string pathFile, int countOpen, bool fileOrigin)
        {
            if (_ignoreExceptions == null || !_ignoreExceptions.Contains(exception.Code)) throw exception;

#if NETFRAMEWORK       

            //countOpen++;

            if (countOpen >= 2) // If it is necessary to force Excel closure (from the 2nd attempt onwards)                
            {
                var result2 = MessageBox.Show(
                   fileOrigin ? Messages.FileOriginInUseAndCloseExcel(pathFile) : Messages.FileDestinationInUseAndCloseExcel(pathFile),
                   Messages.Warning(),
                   MessageBoxButtons.YesNo,
                   MessageBoxIcon.Exclamation);

                if (result2 == DialogResult.Yes)
                {
                    new Features().CloseExcel(Path.GetFileName(pathFile));
                    System.Threading.Thread.Sleep(1500);
                    return 1;

                } // If No, continue execution below
            }

            var result1 = MessageBox.Show(
                fileOrigin ? Messages.FileOriginInUse(pathFile) : Messages.FileDestinationInUse(pathFile),
                Messages.Warning(),
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Exclamation);

            if (result1 == DialogResult.OK) { return 1; }
            else { return 0; } // If canceled
#else
            throw exception;
#endif
        }

        /// <summary>
        /// Incomplete headers will be filled.
        /// </summary>
        internal void HeaderIncomplete(DataTable dataTable, int i, ColumnNameHeaderInvalidSHException except)
        {
            bool ignoreEmptyColumns = _ignoreExceptions != null && _ignoreExceptions.Contains(except.Code);
            if (ignoreEmptyColumns)
            {
                dataTable.Columns[i].ColumnName = $"EmptyColumn{i + 1}";
                //dataTable.Rows[0][i] = $"EmptyColumn{i + 1}";
            }
            else throw except;
        }

        internal int ColumnNotExist(int indexColumn, DataTable dataTable, SHException except)
        {
            if (_ignoreExceptions == null || !_ignoreExceptions.Contains(except.Code)) throw except;

            while (dataTable.Columns.Count < indexColumn)
            {               
                dataTable.Columns.Add($"NewColumn{dataTable.Columns.Count + 1}");
            }

            return 1;           
        }    
        
        internal int DirectoryNotExists(string pathFolder, SHException except)
        {
            if (_ignoreExceptions == null || !_ignoreExceptions.Contains(except.Code)) throw except;
            Directory.CreateDirectory(pathFolder);
            return 1;
        }


    }
}
