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
    internal class ParamException : SHException
    {
        protected new int Number { get; } = 0;

        internal ParamException(string argumentName, string methodName) : base(Messages.ArgumentException(argumentName, methodName)) { }

        public ParamException(string argumentName, string methodName, Exception innerException) : base(Messages.ArgumentException(argumentName, methodName), innerException) { }
    }

    #endregion

    #region File Generic

    [Serializable]
    internal class FileNotFound : SHException
    {
        protected new int Number { get; } = 0;

        internal FileNotFound(string pathFile) : base(Messages.FileNotFound(pathFile)) { }

        public FileNotFound(string pathFile, Exception innerException) : base(Messages.FileNotFound(pathFile), innerException) { }
    }

    [Serializable]
    internal class PathFileNull : SHException
    {
        protected new int Number { get; } = 0;

        internal PathFileNull(string pathFile) : base(Messages.PathFileNull(pathFile)) { }

        public PathFileNull(string pathFile, Exception innerException) : base(Messages.PathFileNull(pathFile), innerException) { }
    }


    #endregion

    #region FileOrigin

    [Serializable]
    internal class FileOriginInUse : SHException
    {
        protected new int Number { get; } = 0;

        internal FileOriginInUse(string pathFile) : base(Messages.FileOriginInUse(pathFile)) { }

        public FileOriginInUse(string pathFile, Exception innerException) : base(Messages.FileOriginInUse(pathFile), innerException) { }
    }    

    #endregion

    #region FileDestiny

    [Serializable]
    internal class FileDestinyInUse : SHException
    {
        protected new int Number { get; } = 0;

        internal FileDestinyInUse(string pathFile) : base(Messages.FileDestinyInUse(pathFile)) { }

        public FileDestinyInUse(string pathFile, Exception innerException) : base(Messages.FileDestinyInUse(pathFile), innerException) { }
    }

    #endregion

    #region Directory

    [Serializable]
    internal class DirectoryNotFoundException : SHException
    {
        protected new int Number { get; } = 0;

        internal DirectoryNotFoundException(string pathFile) : base(Messages.FileDestinyInUse(pathFile)) { }

        public DirectoryNotFoundException(string pathFile, Exception innerException) : base(Messages.FileDestinyInUse(pathFile), innerException) { }
    }

    #endregion

    #region Unzip

    [Serializable]
    internal class UnableUnzip : SHException
    {
        protected new int Number { get; } = 0;

        internal UnableUnzip(string pathFileZip) : base(Messages.UnableUnzip(pathFileZip)) { }

        public UnableUnzip(string pathFileZip, Exception innerException) : base(Messages.UnableUnzip(pathFileZip), innerException) { }
    }

    #endregion

    #region Rows

    [Serializable]
    internal class RowArrayOverflowDteException : SHException
    {
        protected new int Number { get; } = 0;

        internal RowArrayOverflowDteException() : base(Messages.RowArrayOverflowDt()) { }

        public RowArrayOverflowDteException(Exception innerException) : base(Messages.RowArrayOverflowDt(), innerException) { }
    }

    #endregion




}
