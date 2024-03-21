using System.Globalization;
using System.IO;

namespace SH.Globalization
{
    internal static class Messages
    {
        #region Warning

        internal static string Warning()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Aviso",
                _ => $"Warning",
            };
        }

        internal static string VisitDocExceptions()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Por favor, visite https://bit.ly/SheetHelper_Exceptions para saber mais.",
                _ => $"Please visit https://bit.ly/SheetHelper_Exceptions to learn more.",
            };
        }


        #endregion

        #region Sheet

        internal static string SheetErrorIndex(int sheetIndex)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Erro ao selecionar a aba desejada! Verifique se o índice '{sheetIndex}' da aba está correto.",
                _ => $"Error selecting the desired sheet! Please check if the sheet index '{sheetIndex}' is correct.",
            };
        }

        internal static string SheetNameNotFind(string sheetName)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Não foi possível encontrar a planilha desejada '{sheetName}'! Verifique se o nome da planilha está correto.",
                _ => $"Unable to find the desired sheet '{sheetName}'! Please check if the sheet name is correct.",
            };
        }

        internal static string SheetFirstIndex()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A primeira aba é 1!",
                _ => $"The first sheet is 1!",
            };
        }

        internal static string SheetNameInvalid(string sheetName)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Nome de aba '{sheetName}' é inválido.",
                _ => $"Invalid sheet name '{sheetName}'.",
            };
        }

        #endregion

        #region Column

        internal static string ColumnOutRange(string column, int limitIndexColumn, int indexColumn)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A coluna '{column}' está fora do intervalo porque se refere à coluna '{limitIndexColumn + indexColumn + 1}' (min 1, max {limitIndexColumn})!",
                _ => $"The column '{column}' is out of range, because it refers to column '{limitIndexColumn + indexColumn + 1}' (min 1, max {limitIndexColumn})!",
            };
        }

        internal static string ColumnOutRange(string column, int limitIndexColumn)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A coluna '{column}' está fora do intervalo (min 1, max {limitIndexColumn})!",
                _ => $"The column '{column}' is out of range (min 1, max {limitIndexColumn})!",
            };
        }

        internal static string ColumnNotPattern(string column)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A coluna '{column}' informada não possui um padrão válido!",
                _ => $"The column '{column}' entered does not have a valid pattern",
            };
        }

        #endregion

        #region Row

        internal static string RowOutRange(string row, int limitIndexRows, int indexRow)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A linha '{row}' está fora do intervalo porque faz referência à linha '{limitIndexRows + indexRow + 1}' (min 1, max {limitIndexRows})!",
                _ => $"The row '{row}' is out of range, because it refers to row '{limitIndexRows + indexRow + 1}' (min 1, max {limitIndexRows})!",
            };
        }

        internal static string RowOutRange(string row, int limitIndexRows)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A linha '{row}' está fora do intervalo (min 1, max {limitIndexRows})!",
                _ => $"The row '{row}' is out of range (min 1, max {limitIndexRows})!",
            };
        }

        internal static string RowInvalid(string row)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A linha '{row}' não é válida!",
                _ => $"The row '{row}' is not a valid!",
            };
        }

        internal static string RowNotPattern(string row)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A linha '{row}' não possui um padrão válido!",
                _ => $"Row '{row}' is not a valid pattern!",
            };
        }

        #endregion

        #region Separator

        internal static string SeparatorInvalid(string separator)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Invalid separator '{separator}'!",
                _ => $"Separador '{separator}' é inválido!",
            };
        }

        #endregion

        #region File

        internal static string FileOriginInUse(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O arquivo de origem '{Path.GetFileName(pathFile)}' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida tente novamente.",
                _ => $"The origin file '{Path.GetFileName(pathFile)}' is being used by another process. Please finish its use and then try again.",
            };
        }

        internal static string FileOriginErrorValid()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Ocorreu um erro ao validar o arquivo de origem!",
                _ => $"An error occurred while validating the origin file!",
            };
        }

        internal static string FileOriginNotFound(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Arquivo de origem não encontrado em '{pathFile}'.",
                _ => $"Origin file not found in '{pathFile}'.",
            };
        }

        internal static string FileOriginNameNull()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O arquivo 'origin' é nulo ou vazio.",
                _ => $"The 'origin' file is null or empty.",
            };
        }

        internal static string FileDestinyInUse(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O arquivo de destino '{Path.GetFileName(pathFile)}' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida tente novamente",
                _ => $"The destiny file '{Path.GetFileName(pathFile)}' is being used by another process. Please finish its use and then try again",
            };
        }

        internal static string FileDestinyErrorValid()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Ocorreu um erro ao validar o caminho do arquivo de destino!",
                _ => $"An error occurred while validating the destination file path!",
            };
        }

        internal static string FileDestinyNameNull()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O caminho de destino para o arquivo ('destiny') é nulo ou vazio.",
                _ => $"The destination path for the file ('destiny') is null or empty.",
            };
        }

        internal static string FileZipUnable(string extension)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Não é possível extrair este arquivo, pois '{extension}' ainda não é suportada.",
                _ => $"Can't extract this file, because '{extension}' is not supported yet.",
            };
        }


        #endregion

    }
}
