﻿using SH.Exceptions;
using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace SH.ExcelHelper.Treatments
{

    /// <summary>
    /// Used to process the information provided by the user
    /// </summary>
    internal class Validations
    {
        readonly TryHandlerExceptions _tryHandlerExceptions;

        public Validations(SheetHelper sheetHelper)
        {
            _tryHandlerExceptions = new TryHandlerExceptions(sheetHelper);
        }





        /// <summary>
        /// Checks if it is necessary to convert the file.
        /// </summary>
        /// <returns>True if conversion is required.</returns>
        internal bool CheckConvertNecessary(string origin, string destiny, string sheet, string separator, string? columns, string? rows)
        {
            bool checkFormat = Path.GetExtension(origin).Equals(Path.GetExtension(destiny), StringComparison.OrdinalIgnoreCase); // The formats is the same?
            bool isOriginTextFormat = Path.GetExtension(origin).Equals(".csv", StringComparison.OrdinalIgnoreCase) || Path.GetExtension(origin).Equals(".txt", StringComparison.OrdinalIgnoreCase);                                                                             // 
            bool isDestinyTextFormat = Path.GetExtension(destiny).Equals(".csv", StringComparison.OrdinalIgnoreCase) || Path.GetExtension(destiny).Equals(".txt", StringComparison.OrdinalIgnoreCase);
            bool checkFormatText = isOriginTextFormat == isDestinyTextFormat;
            bool checkColumns = string.IsNullOrEmpty(columns) || columns.Trim().Equals(":") || columns.Trim().Equals("1:") || columns.Trim().Equals("A:"); // All columns?
            bool checkRows = string.IsNullOrEmpty(rows) || rows.Trim().Equals(":") || rows.Trim().Equals("1:"); // All rows?

            return !((checkFormat || checkFormatText) && checkColumns && checkRows);
        }

        internal void ValidateString(string param, string paramName, string methodName)
        {
            if (string.IsNullOrEmpty(param))
            {
                throw new ArgumentSHException(paramName, methodName); // GetCallingMethodName(indStack)
            }
        }

        internal void ValidateFile(string pathFile, string paramName, string methodName)
        {
            ValidateString(pathFile, paramName, methodName);

            if (!File.Exists(pathFile))
            {
                throw new FileNotFoundSHException(pathFile);
            }
        }

        internal void ValidateOriginFile(string? pathOrigin, string paramName, string methodName)
        {

            int countOpen = 0;

        again:
            try
            {
                ValidateFile(pathOrigin, paramName, methodName);
                File.OpenRead(pathOrigin).Dispose(); // To check if the file is accessible
            }
            catch (UnauthorizedAccessException ex)
            {
                switch (_tryHandlerExceptions.FileExcelInUse(ex, pathOrigin, countOpen, true))
                {
                    case 0: return;
                    case 1: goto again;
                    default: throw new FileDestinyInUseSHException(pathOrigin);
                }                
            }
            catch (Exception ex)
            {
                throw new Exception("E-0000-SH: An error occurred while validating the origin file.", ex);
            }
        }

        internal void ValidateDestinyFile(string destiny)
        {
            int countOpen = 0;

        again:
            try
            {

                File.WriteAllText(destiny, ""); // To check if the destination file is accessible
                File.Delete(destiny); // Delete to prevent the file from being opened during conversion
            }
            catch (UnauthorizedAccessException ex)
            {
                switch (_tryHandlerExceptions.FileExcelInUse(ex, destiny, countOpen, false))
                {
                    case 0: return;
                    case 1: goto again;
                    default: throw new FileDestinyInUseSHException(destiny);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("E-0000-SH: An error occurred while validating the destination file.", ex);
            }
        }

        internal void ValidateDestinyFolder(string destiny, bool createIfNotExist)
        {
            try
            {
                if (createIfNotExist) { Directory.CreateDirectory(destiny); }
                if (!Directory.Exists(destiny)) throw new Exceptions.DirectoryNotFoundSHException("E-0000-SH: Destiny folder not found.");
            }
            catch (UnauthorizedAccessException)
            {
                throw new InvalidOperationException("E-0000-SH: The destination file is in use by another process.");
            }
            catch (Exception ex)
            {
                throw new Exception("E-0000-SH: An error occurred while validating the destination file.", ex);
            }
        }

        internal void ValidateSheet(string sheet)
        {
            // "1" or "Fist_Sheet_Name"

            if (string.IsNullOrEmpty(sheet))
            {
                throw new ArgumentException("E-0000-SH: Invalid sheet name.", nameof(sheet));
            }

            if (int.TryParse(sheet, out int sheetNumber) && sheetNumber <= 0)
            {
                throw new ArgumentException("E-0000-SH: The first sheet is '1'!");
            }
        }

        internal void ValidateSeparator(string separator)
        {
            // ";"

            if (string.IsNullOrEmpty(separator))
            {
                throw new ArgumentException("E-0000-SH: Invalid separator.", nameof(separator));
            }
        }

        internal void ValidateColumns(string? columns)
        {
            // "A:H, 4:9, B, 75, -2"

            // TODO: Add specific validation logic for columns
        }

        internal void ValidateRows(string? rows)
        {
            // "1:23, 34:56, 70, 75, -1"


            // throw new ArgumentException("Invalid rows.", nameof(rows));


            // TODO: Add specific validation logic for rows
        }

        internal void Validate(string origin, string destiny, string sheet, string separator, string? columns, string? rows, string methodName)
        {
            List<Task> validates = new()
            {
                Task.Run(() => ValidateOriginFile(origin,nameof(origin), methodName)),
                Task.Run(() => ValidateDestinyFile(destiny)),
                Task.Run(() => ValidateSheet(sheet)),
                Task.Run(() => ValidateSeparator(separator)),
                Task.Run(() => ValidateColumns(columns)),
                Task.Run(() => ValidateRows(rows))
            };

            try
            {
                //Task.WhenAll(validates).Wait();
                Task.WaitAll(validates.ToArray());
            }
            catch (AggregateException ex)
            {                
                throw new AggregateException("One or more validations failed.", ex.InnerExceptions);
            }
        }

        internal void Validate(string destiny, string separator, string? columns, string? rows)
        {
            List<Task> validates = new()
            {
                Task.Run(() => ValidateDestinyFile(destiny)),
                Task.Run(() => ValidateSeparator(separator)),
                Task.Run(() => ValidateColumns(columns)),
                Task.Run(() => ValidateRows(rows))
            };
            Task.WhenAll(validates).Wait();
        }

        //internal static void ValidateFormatFileOrigin(string pathFile, string desiredFormat)
        //{
        //    if (!string.IsNullOrEmpty(pathFile?.Trim()))
        //    {
        //        if (!File.Exists(pathFile)) throw new ParamException(nameof(zipFile), nameof(UnZIP));
        //    }
        //    else
        //    {
        //        throw new ParamException(nameof(zipFile), nameof(UnZIP));
        //    }
        //}

        internal string GetCallingMethodName(int indStack)
        {
            //StackFrame frame = new(indStack);
            //StringBuilder nameMethod = new();

            //for (int i = 1; frame.GetILOffset() != -1 && i <= levelPath; i++)
            //{
            //    nameMethod.Insert(0, "/" + frame.GetMethod().Name);
            //    frame = new StackFrame(indStack + i);
            //}

            //return nameMethod.ToString();           

            StackFrame callerFrame = new StackTrace().GetFrame(indStack); // 1 para obter o chamador do método atual
            return callerFrame.GetMethod().Name;
        }

    }
}
