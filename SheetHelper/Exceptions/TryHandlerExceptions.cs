using ExcelDataReader.Log.Logger;
using SH.Globalization;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SH.Exceptions
{
    /// <summary>
    /// <para>0: Ignore the exception and no warning will be thrown;</para>
    /// <para>1: Try again.</para>
    /// <para>Or throw exception</para>
    /// </summary>
    internal static class TryHandlerExceptions
    {
        //internal static int ExceptionManager(SHException exception, List<string>? tryIgnoreExceptions)
        //{
        //    switch (exception.Code)
        //    {




        //        default:
        //            break;
        //    }
        //}




        internal static int FileExcelInUse(SHException exception, string pathFile, int countOpen, bool fileOrigin)
        {
            //if (tryIgnoreExceptions == null || !tryIgnoreExceptions.Contains(exception.Code)) throw exception;

            //#if !NETFRAMEWORK            
            //            throw exception; // Handle exception only to NETFRAMEWORK
            //#endif

            countOpen++;

            if (countOpen >= 2) // If it is necessary to force Excel closure (from the 2nd attempt onwards)                
            {
                var result2 = MessageBox.Show(
                   fileOrigin ? Messages.FileOriginInUseAndCloseExcel(pathFile) : Messages.FileDestinyInUseAndCloseExcel(pathFile),
                   Messages.Warning(),
                   MessageBoxButtons.YesNo,
                   MessageBoxIcon.Exclamation);

                if (result2 == DialogResult.Yes)
                {
                    SheetHelper.CloseExcel();
                    System.Threading.Thread.Sleep(1500);
                    return 1;

                } // If No, continue execution below
            }

            var result1 = MessageBox.Show(
                fileOrigin ? Messages.FileOriginInUse(pathFile) : Messages.FileDestinyInUse(pathFile),
                Messages.Warning(),
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Error);

            if (result1 == DialogResult.OK) { return 1; }
            else { return 0; } // If canceled
        }





    }
}
