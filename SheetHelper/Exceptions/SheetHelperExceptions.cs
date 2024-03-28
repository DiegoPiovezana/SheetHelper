using SH.Globalization;
using System;

namespace SH.Exceptions
{
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



    [Serializable]
    internal class ParamException : SHException
    {
        protected new int Number { get; } = 0;

        internal ParamException(string argumentName, string methodName) : base(Messages.ArgumentException(argumentName, methodName)) { }

        public ParamException(string argumentName, string methodName, Exception innerException) : base(Messages.ArgumentException(argumentName, methodName), innerException) { }
    }


    #region FileOrigin

    [Serializable]
    internal class FileOriginInUse : SHException
    {
        protected new int Number { get; } = 0;

        internal FileOriginInUse(string pathFile) : base(Messages.FileOriginInUse(pathFile)) { }

        public FileOriginInUse(string pathFile, Exception innerException) : base(Messages.FileOriginInUse(pathFile), innerException) { }
    }

    [Serializable]
    internal class FileOriginNotFound : SHException
    {
        protected new int Number { get; } = 0;

        internal FileOriginNotFound(string pathFile) : base(Messages.FileOriginNotFound(pathFile)) { }

        public FileOriginNotFound(string pathFile, Exception innerException) : base(Messages.FileOriginNotFound(pathFile), innerException) { }
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



}
