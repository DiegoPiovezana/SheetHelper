using SH.Globalization;
using System;

namespace SH
{
    [Serializable]
    internal class CustomExceptionBase : Exception
    {
        protected int Number { get; } = 0;
        protected string Code => $"E-{Number:D4}-SH"; // E-0000-SH
        protected new string Message { get; } = string.Empty;

        internal CustomExceptionBase() { }

        internal CustomExceptionBase(string message) : base(message)
        {
            Message = $"{Code}: {message} \n\n Please visit https://bit.ly/SheetHelper_Exceptions to learn more.";
        }

        internal CustomExceptionBase(string message, Exception innerException) : base(message, innerException)
        {
            //Code = $"E-{Number:D4}-SH";
            Message = $"{Code}: {message} \n\n Please visit https://bit.ly/SheetHelper_Exceptions to learn more.";
        }
    }

    [Serializable]
    internal class FileOriginInUse : CustomExceptionBase
    {
        protected new int Number { get; } = 0;

        internal FileOriginInUse(string pathFile) : base(Messages.FileOriginInUse(pathFile)) { }

        public FileOriginInUse(string pathFile, Exception innerException) : base(Messages.FileOriginInUse(pathFile), innerException) { }
    }

    [Serializable]
    internal class FileDestinyInUse : CustomExceptionBase
    {
        protected new int Number { get; } = 0;

        internal FileDestinyInUse(string pathFile) : base(Messages.FileDestinyInUse(pathFile)) { }

        public FileDestinyInUse(string pathFile, Exception innerException) : base(Messages.FileDestinyInUse(pathFile), innerException) { }
    }

   


}
