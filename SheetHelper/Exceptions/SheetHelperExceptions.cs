using SH.Globalization;
using System;

namespace SH.Exceptions
{
    #region Principal

    [Serializable]
    internal class SHException : Exception
    {
        protected int Number { get; } = 0;
        public string Code => $"E-{Number:D4}-SH"; // E-0000-SH
        public new string Message { get; } = string.Empty;

        internal SHException() { }

        internal SHException(string message) : base(message)
        {
            Message = $"{Code}: {message} \n\n {Messages.VisitDocExceptions()}";
        }

        internal SHException(string message, Exception innerException) : base(message, innerException)
        {
            Message = $"{Code}: {message} \n\n {Messages.VisitDocExceptions()}";
        }
    }

    #endregion

    #region Generic Params

    [Serializable]
    internal class ArgumentNullOrEmptySHException : SHException
    {
        protected new int Number { get; } = 0;

        internal ArgumentNullOrEmptySHException(string argumentName, string methodName) : base(Messages.ArgumentNullOrEmptyException(argumentName, methodName)) { }

        public ArgumentNullOrEmptySHException(string argumentName, string methodName, Exception innerException) : base(Messages.ArgumentNullOrEmptyException(argumentName, methodName), innerException) { }
    }

    [Serializable]
    internal class ArgumentMinSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal ArgumentMinSHException(string argumentName, string methodName, int value, int min) : base(Messages.ArgumentMinException(argumentName, methodName, value, min)) { }

        public ArgumentMinSHException(string argumentName, string methodName, int value, int min, Exception innerException) : base(Messages.ArgumentMinException(argumentName, methodName, value, min), innerException) { }
    }

    #endregion

    #region File Generic

    [Serializable]
    internal class FileNotFoundSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal FileNotFoundSHException(string pathFile) : base(Messages.FileNotFound(pathFile)) { }

        public FileNotFoundSHException(string pathFile, Exception innerException) : base(Messages.FileNotFound(pathFile), innerException) { }
    }

    [Serializable]
    internal class PathFileNullSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal PathFileNullSHException(string pathFile) : base(Messages.PathFileNull(pathFile)) { }

        public PathFileNullSHException(string pathFile, Exception innerException) : base(Messages.PathFileNull(pathFile), innerException) { }
    }


    #endregion

    #region FileOrigin

    [Serializable]
    internal class FileOriginInUseSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal FileOriginInUseSHException(string pathFile) : base(Messages.FileOriginInUse(pathFile)) { }

        public FileOriginInUseSHException(string pathFile, Exception innerException) : base(Messages.FileOriginInUse(pathFile), innerException) { }
    }

    #endregion

    #region FileDestination

    [Serializable]
    internal class FileDestinationInUseSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal FileDestinationInUseSHException(string pathFile) : base(Messages.FileDestinationInUse(pathFile)) { }

        public FileDestinationInUseSHException(string pathFile, Exception innerException) : base(Messages.FileDestinationInUse(pathFile), innerException) { }
    }

    #endregion

    #region Directory

    [Serializable]
    internal class DirectoryNotFoundSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal DirectoryNotFoundSHException(string pathFile) : base(Messages.FileDestinationInUse(pathFile)) { }

        public DirectoryNotFoundSHException(string pathFile, Exception innerException) : base(Messages.FileDestinationInUse(pathFile), innerException) { }
    }

    #endregion

    #region Unzip

    [Serializable]
    internal class UnableUnzipSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal UnableUnzipSHException(string pathFileZip) : base(Messages.UnableUnzip(pathFileZip)) { }

        public UnableUnzipSHException(string pathFileZip, Exception innerException) : base(Messages.UnableUnzip(pathFileZip), innerException) { }
    }

    #endregion

    #region Rows

    [Serializable]
    internal class RowArrayOverflowDteSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal RowArrayOverflowDteSHException() : base(Messages.RowArrayOverflowDt()) { }

        public RowArrayOverflowDteSHException(Exception innerException) : base(Messages.RowArrayOverflowDt(), innerException) { }
    }

    [Serializable]
    internal class RowsMinDtSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal RowsMinDtSHException(string nameDt) : base(Messages.RowsMinDt(nameDt)) { }

        public RowsMinDtSHException(string nameDt, Exception innerException) : base(Messages.RowsMinDt(nameDt), innerException) { }
    }

    #endregion




}
