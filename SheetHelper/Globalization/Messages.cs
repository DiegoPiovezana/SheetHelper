using System.Globalization;
using System.IO;

namespace SH.Globalization
{
    internal static class Messages
    {
        internal static string FileInUse(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O arquivo '{Path.GetFileName(pathFile)}' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida tente novamente.",
                _ => $"The file '{Path.GetFileName(pathFile)}' is being used by another process. Please finish its use and then try again.",
            };
        }





    }
}
