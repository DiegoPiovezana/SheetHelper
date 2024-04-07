using SH.Globalization;
using System;
using System.Diagnostics;
using System.Text;

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
    internal class ArgumentSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal ArgumentSHException(string argumentName, string methodName) : base(Messages.ArgumentException(argumentName, methodName)) { }

        public ArgumentSHException(string argumentName, string methodName, Exception innerException) : base(Messages.ArgumentException(argumentName, methodName), innerException) { }
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

    #region FileDestiny

    [Serializable]
    internal class FileDestinyInUseSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal FileDestinyInUseSHException(string pathFile) : base(Messages.FileDestinyInUse(pathFile)) { }

        public FileDestinyInUseSHException(string pathFile, Exception innerException) : base(Messages.FileDestinyInUse(pathFile), innerException) { }
    }

    #endregion

    #region Directory

    [Serializable]
    internal class DirectoryNotFoundSHException : SHException
    {
        protected new int Number { get; } = 0;

        internal DirectoryNotFoundSHException(string pathFile) : base(Messages.FileDestinyInUse(pathFile)) { }

        public DirectoryNotFoundSHException(string pathFile, Exception innerException) : base(Messages.FileDestinyInUse(pathFile), innerException) { }
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

    #endregion




}
