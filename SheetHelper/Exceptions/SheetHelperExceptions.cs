using SH.Globalization;
using System;

namespace SH.Exceptions
{
    // https://bit.ly/SheetHelper_Exceptions

    #region Principal

    [Serializable]
    public class SHException : Exception
    {
        protected virtual int Number { get; } = 0;
        public string Code => $"E-{Number:D4}-SH"; // E-0000-SH
        public string Message { get; } = string.Empty;
        //public string LinkDoc { get; } = string.Empty; -- Not

        internal SHException() { }

        internal SHException(string message) : base(message)
        {
            Message = $"{Code}: {message} \n\n {Messages.VisitDocExceptions()}";
            HResult = Number;
        }

        internal SHException(string message, Exception innerException) : base(message, innerException)
        {
            Message = $"{Code}: {message} \n\n {Messages.VisitDocExceptions()}";
            HResult = Number;
        }
    }

    #endregion

    #region Generic Params

    [Serializable]
    public class ArgumentNullOrEmptySHException : SHException
    {
        protected override int Number => 0;

        internal ArgumentNullOrEmptySHException(string argumentName, string methodName) : base(Messages.ArgumentNullOrEmptyException(argumentName, methodName)) { }

        internal ArgumentNullOrEmptySHException(string argumentName, string methodName, Exception innerException) : base(Messages.ArgumentNullOrEmptyException(argumentName, methodName), innerException) { }
    }

    [Serializable]
    public class ArgumentMinSHException : SHException
    {
        protected override int Number => 0;

        internal ArgumentMinSHException(string argumentName, string methodName, int value, int min) : base(Messages.ArgumentMinException(argumentName, methodName, value, min)) { }

        public ArgumentMinSHException(string argumentName, string methodName, int value, int min, Exception innerException) : base(Messages.ArgumentMinException(argumentName, methodName, value, min), innerException) { }
    }

    #endregion

    #region File Generic

    [Serializable]
    public class FileNotFoundSHException : SHException
    {
        protected override int Number => 0;

        internal FileNotFoundSHException(string pathFile) : base(Messages.FileNotFound(pathFile)) { }

        public FileNotFoundSHException(string pathFile, Exception innerException) : base(Messages.FileNotFound(pathFile), innerException) { }
    }

    [Serializable]
    public class PathFileNullSHException : SHException
    {
        protected override int Number => 0;

        internal PathFileNullSHException(string pathFile) : base(Messages.PathFileNull(pathFile)) { }

        public PathFileNullSHException(string pathFile, Exception innerException) : base(Messages.PathFileNull(pathFile), innerException) { }
    }


    #endregion

    #region FileOrigin

    [Serializable]
    public class FileOriginInUseSHException : SHException
    {
        protected override int Number => 541;

        internal FileOriginInUseSHException(string pathFile) : base(Messages.FileOriginInUse(pathFile)) { }

        public FileOriginInUseSHException(string pathFile, Exception innerException) : base(Messages.FileOriginInUse(pathFile), innerException) { }
    }

    [Serializable]
    public class FileOriginNotReadSupportSHException : SHException
    {
        protected override int Number => 541;

        public FileOriginNotReadSupportSHException(string pathFile, ExcelDataReader.Exceptions.HeaderException innerException) : base(Messages.FileOriginNotReadSupport(pathFile, innerException), innerException) { }
    }



    #endregion

    #region FileDestination

    [Serializable]
    public class FileDestinationInUseSHException : SHException
    {
        protected override int Number => 541;

        internal FileDestinationInUseSHException(string pathFile) : base(Messages.FileDestinationInUse(pathFile)) { }

        public FileDestinationInUseSHException(string pathFile, Exception innerException) : base(Messages.FileDestinationInUse(pathFile), innerException) { }
    }

    #endregion

    #region Directory

    [Serializable]
    public class DirectoryNotFoundSHException : SHException
    {
        protected override int Number => 0;

        internal DirectoryNotFoundSHException(string pathFile) : base(Messages.FileDestinationInUse(pathFile)) { }

        public DirectoryNotFoundSHException(string pathFile, Exception innerException) : base(Messages.FileDestinationInUse(pathFile), innerException) { }
    }

    #endregion

    #region Unzip

    [Serializable]
    public class UnableUnzipSHException : SHException
    {
        protected override int Number => 0;

        internal UnableUnzipSHException(string pathFileZip) : base(Messages.UnableUnzip(pathFileZip)) { }

        public UnableUnzipSHException(string pathFileZip, Exception innerException) : base(Messages.UnableUnzip(pathFileZip), innerException) { }
    }

    #endregion

    #region Rows

    [Serializable]
    public class RowArrayOverflowDteSHException : SHException
    {
        protected override int Number => 0;

        internal RowArrayOverflowDteSHException() : base(Messages.RowArrayOverflowDt()) { }

        public RowArrayOverflowDteSHException(Exception innerException) : base(Messages.RowArrayOverflowDt(), innerException) { }
    }

    [Serializable]
    public class RowsMinDtSHException : SHException
    {
        protected override int Number => 0;

        internal RowsMinDtSHException(string nameDt) : base(Messages.RowsMinDt(nameDt)) { }

        public RowsMinDtSHException(string nameDt, Exception innerException) : base(Messages.RowsMinDt(nameDt), innerException) { }
    }

    [Serializable]
    public class RowOutRangeSHException : SHException
    {
        protected override int Number => 4042;

        internal RowOutRangeSHException(string row, int limitIndexRows) : base(Messages.RowOutRange(row, limitIndexRows)) { }

        public RowOutRangeSHException(string row, int limitIndexRows, Exception innerException) : base(Messages.RowOutRange(row, limitIndexRows), innerException) { }
    }

    [Serializable]
    public class RowRefOutRangeSHException : SHException
    {
        protected override int Number => 4042;

        internal RowRefOutRangeSHException(string row, int limitIndexRows, int indexRow) : base(Messages.RowRefOutRange(row, limitIndexRows, indexRow)) { }

        public RowRefOutRangeSHException(string row, int limitIndexRows, int indexRow, Exception innerException) : base(Messages.RowRefOutRange(row, limitIndexRows, indexRow), innerException) { }
    }

    #endregion

    #region Columns

    [Serializable]
    public class ColumnNameHeaderInvalidSHException : SHException
    {
        protected override int Number => 4041;

        internal ColumnNameHeaderInvalidSHException(int indexColumn) : base(Messages.ColumnNameHeaderInvalidRange(indexColumn)) { }

        public ColumnNameHeaderInvalidSHException(int indexColumn, Exception innerException) : base(Messages.ColumnNameHeaderInvalidRange(indexColumn), innerException) { }
    }

    [Serializable]
    public class ColumnOutRangeSHException : SHException
    {
        protected override int Number => 4042;

        internal ColumnOutRangeSHException(int indexColumn, int limitIndexColumn) : base(Messages.ColumnOutRange(indexColumn, limitIndexColumn)) { }

        public ColumnOutRangeSHException(int indexColumn, int limitIndexColumn, Exception innerException) : base(Messages.ColumnOutRange(indexColumn, limitIndexColumn), innerException) { }
    }

    [Serializable]
    public class ColumnRefOutRangeSHException : SHException
    {
        protected override int Number => 4042;

        internal ColumnRefOutRangeSHException(int indexColumn, int limitIndexColumn) : base(Messages.ColumnRefOutRange(indexColumn, limitIndexColumn)) { }

        public ColumnRefOutRangeSHException(int indexColumn, int limitIndexColumn, Exception innerException) : base(Messages.ColumnRefOutRange(indexColumn, limitIndexColumn), innerException) { }
    }

    #endregion

    #region Sheets

    //[Serializable]
    //public class SheetNullSHException : SHException
    //{
    //    protected override int Number => 0;

    //    internal RowArrayOverflowDteSHException() : base(Messages.RowArrayOverflowDt()) { }

    //    public RowArrayOverflowDteSHException(Exception innerException) : base(Messages.RowArrayOverflowDt(), innerException) { }
    //}

    //[Serializable]
    //public class SheetNullSHException : SHException
    //{
    //    protected override int Number => 0;

    //    internal RowArrayOverflowDteSHException() : base(Messages.RowArrayOverflowDt()) { }

    //    public RowArrayOverflowDteSHException(Exception innerException) : base(Messages.RowArrayOverflowDt(), innerException) { }
    //}

    #endregion

    #region Converter

    [Serializable]
    public class ParamMissingConverterSHException : SHException
    {
        protected override int Number => 0;

        internal ParamMissingConverterSHException() : base(Messages.ParamMissingConverter()) { }

        public ParamMissingConverterSHException(Exception innerException) : base(Messages.ParamMissingConverter(), innerException) { }
    }

    #endregion


}
