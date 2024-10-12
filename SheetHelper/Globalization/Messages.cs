using SH.ExcelHelper.Tools;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace SH.Globalization
{
    //TODO: use Resouces

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
                "pt-BR" => $"Por favor, visite https://bit.ly/SheetHelper_Exceptions para saber mais. Tente pesquisar pelo código do erro.",
                _ => $"Please visit https://bit.ly/SheetHelper_Exceptions to learn more. Try searching for the exception code.",
            };
        }

        internal static string ArgumentNullOrEmptyException(string argumentName, string methodName)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O parâmetro '{argumentName}' do método '{methodName}' não é válido pois está vazio ou inexistente! Por favor, verifique se está preenchido adequadamente. Considere consultar a documentação disponível em https://bit.ly/SheetHelper.",
                _ => $"The '{argumentName}' parameter of the '{methodName}' method is not valid because it is empty or non-existent! Please check that it is filled out properly. Consider consulting the documentation available at https://bit.ly/SheetHelper.",
            };
        }

        internal static string ArgumentMinException(string argumentName, string methodName, int value, int min)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O parâmetro '{argumentName}' do método '{methodName}' não é válido pois é '{value}' e deve ser no mínimo '{min}'! Por favor, verifique se está preenchido adequadamente. Considere consultar a documentação disponível em https://bit.ly/SheetHelper.",
                _ => $"The '{argumentName}' parameter of the '{methodName}' method is not valid because it is '{value}' and must be at least '{min}'! Please check that it is filled out properly. Consider consulting the documentation available at https://bit.ly/SheetHelper.",
            };
        }

        internal static string UnmappedException(string methodName, Exception ex)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Um erro ocorreu ao utilizar o método '{methodName}'. Por favor, verifique se os parâmetros utilizados estão corretos. Se desejar, reporte seu erro no card 'Unmapped Error' disponível em  https://bit.ly/SheetHelper_Exceptions ou consulte a documentação disponível em https://bit.ly/SheetHelper. \n\n{ex.Message} \n\n{ex.InnerException?.Message}",
                _ => $"An error occurred when using the '{methodName}' method. Please check that the parameters entered are correct. If you wish, report your error in the 'Unmapped Error' card available at https://bit.ly/SheetHelper_Exceptions or consult the documentation available at https://bit.ly/SheetHelper. \n\n{ex.Message} \n\n{ex.InnerException?.Message}",
            };
        }


        #endregion

        #region Converter

        internal static string ParamMissingConverter()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Nem todas as conversões possuem todos os parametros definidos. A quantidade total de parâmetros não está correta. Por favor verifique se há alguma informação ausente.",
                _ => $"Not all conversions have all parameters defined. The total number of parameters is not correct. Please check if there is any missing information.",
            };
        }

        #endregion

        #region Sheet

        internal static string SheetIndexNotFound(int sheetIndex)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Erro ao selecionar a aba desejada! Verifique se '{sheetIndex}' é o índice da aba correta.",
                _ => $"Error selecting the desired sheet! Please check if the sheet index '{sheetIndex}' is correct.",
            };
        }

        internal static string SheetNameNotFound(string sheetName)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Não foi possível encontrar a aba '{sheetName}' desejada! Verifique se o nome da aba está correto.",
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

        internal static string ColumnNameHeaderInvalidRange(int indexColumn)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O cabeçalho da coluna '{indexColumn}' do arquivo (CSV ou TXT por exemplo) que está sendo lido é inválida pois não pode estar em branco.",
                _ => $"The column header '{indexColumn}' of the file (CSV or TXT for example) being read is invalid as it cannot be blank.",
            };
        }

        internal static string ColumnOutRange(int indexColumn, int limitIndexColumn)
        {
            string nameLastColumn = new Features().GetNameColumn(limitIndexColumn);

            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A coluna '{indexColumn}' está fora do intervalo (min 1 ou A, max {limitIndexColumn} ou {nameLastColumn})!",
                _ => $"The column '{indexColumn}' is out of range (min 1 or A, max {limitIndexColumn} or {nameLastColumn})!",
            };
        }

        internal static string ColumnRefOutRange(int indexColumn, int limitIndexColumn)
        {
            string nameLastColumn = new Features().GetNameColumn(limitIndexColumn);
            string letterColumn = new Features().GetNameColumn(indexColumn);

            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A coluna {indexColumn} '{letterColumn}' está fora do intervalo porque se refere à coluna '{limitIndexColumn + indexColumn + 1}' (min 1 ou A, max {limitIndexColumn} ou {nameLastColumn})!",
                _ => $"The column {indexColumn} '{letterColumn}' is out of range, because it refers to column '{limitIndexColumn + indexColumn + 1}' (min 1 or A, max {limitIndexColumn} or {nameLastColumn})!",
            };
        }

        internal static string ColumnNotPattern(string column)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A coluna '{column}' informada não possui um padrão válido!",
                _ => $"The column '{column}' entered does not have a valid pattern!",
            };
        }

        #endregion

        #region Rows

        internal static string RowRefOutRange(string row, int limitIndexRows, int indexRow)
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

        internal static string RowCountNotMinimum(string sheet)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A aba '{sheet}' da planilha não possui o número mínimo de linhas desejado.",
                _ => $"The '{sheet}' tab of the spreadsheet does not have the minimum number of rows desired.",
            };
        }

        internal static string RowArrayOverflowDt()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O tamanho do vetor excede o número de linhas da dataTable!",
                _ => $"The length of the array exceeds the number of rows in the dataTable!",
            };
        }

        internal static string RowsMinDt(string nameDt)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O dataTable '{nameDt}' não possui a quantidade mínima de linhas desejada!",
                _ => $"The dataTable '{nameDt}' does not have the minimum number of rows desired!",
            };
        }


        #endregion

        #region Delimiter

        internal static string DelimiterInvalid(string delimiter)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Invalid delimiter '{delimiter}'!",
                _ => $"Separador '{delimiter}' é inválido!",
            };
        }

        #endregion

        #region File Generic

        internal static string FileNotFound(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O arquivo '{Path.GetFileName(pathFile)}' não foi localizado em '{pathFile}'.",
                _ => $"The file '{Path.GetFileName(pathFile)}' was not found in '{pathFile}'..",
            };
        }

        internal static string PathFileNull(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O caminho informado não é válido!",
                _ => $"The path is not valid!",
            };
        }

        #endregion

        #region FileOrigin

        internal static string FileOriginInUse(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O arquivo de origem '{Path.GetFileName(pathFile)}' não pode ser lido pois está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida tente novamente clicando em OK.",
                _ => $"The origin file '{Path.GetFileName(pathFile)}' cannot be read because it is being used by another process. Please finish its use and then try again by clicking OK.",
            };
        }

        internal static string FileOriginInUseAndCloseExcel(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Parece que o arquivo '{Path.GetFileName(pathFile)}' ainda continua em uso. Deseja forçar o encerramento de todos os Excel e tentar novamente? \n\nATENÇÃO: Todos os Excel abertos serão fechados e as alterações da(s) planilha(s) não serão salvas!",
                _ => $"It appears that the file '{Path.GetFileName(pathFile)}' is still in use. Do you want to force quit all Excel and try again? \n\nATTENTION: All open Excel will be closed and changes to the worksheet(s) will not be saved!",
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

        internal static string FileOriginNameNull()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O arquivo de origem 'origin' não pode ser lido pois o parâmetro está nulo ou vazio.",
                _ => $"The 'origin' file cannot be read as the parameter is null or empty.",
            };
        }

        internal static string FileOriginNotReadSupport(string pathOrigin, ExcelDataReader.Exceptions.HeaderException innerEx)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Sem suporte para realizar a leitura do arquivo '{Path.GetFileName(pathOrigin)}' com formato '{Path.GetExtension(pathOrigin)}'. \n\n{innerEx.Message}",
                _ => $"No support for reading the file '{Path.GetFileName(pathOrigin)}' in '{Path.GetExtension(pathOrigin)}' format. \n\n{innerEx.Message}",
            };
        }

        #endregion

        #region FileDestination

        internal static string FileDestinationInUse(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O arquivo de destino '{Path.GetFileName(pathFile)}' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida tente novamente.",
                _ => $"The destination file '{Path.GetFileName(pathFile)}' is being used by another process. Please finish its use and then try again.",
            };
        }

        internal static string FileDestinationInUseAndCloseExcel(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Parece que o arquivo '{Path.GetFileName(pathFile)}' ainda continua em uso. Deseja forçar o encerramento desse Excel e tentar novamente? \n\nATENÇÃO: O Excel aberto será fechado e as possiveis alterações não salvas podem ser perdidas.",
                _ => $"It appears that the file '{Path.GetFileName(pathFile)}' is still in use. Do you want to force close this Excel and try again? \n\nATTENTION: The opened Excel will be closed and any unsaved changes may be lost.",
            };
        }

        internal static string FileDestinationErrorValid()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Ocorreu um erro ao validar o caminho do arquivo de destino!",
                _ => $"An error occurred while validating the destination file path!",
            };
        }

        internal static string FileDestinationNameNull()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O caminho de destino para o arquivo 'destination' é nulo ou vazio.",
                _ => $"The destination path for the file 'destination' is null or empty.",
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

        #region FolderDestination

        internal static string FolderDestinationNotExists(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O diretório de destino '{Path.GetDirectoryName(pathFile)}' para o arquivo '{Path.GetFileName(pathFile)}' não existe!",
                _ => $"The destination directory '{Path.GetDirectoryName(pathFile)}' for the file '{Path.GetFileName(pathFile)}' does not exist!",
            };
        }

        #endregion

        #region JsonDictionary

        internal static string JsonTextItemsNull()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A string JSON 'jsonTextItems' que contém os pares de chave-valor a serem convertidos como parâmetro do método 'GetDictionaryJson' é nula ou vazia.",
                _ => $"The JSON string 'jsonTextItems' containing the key-value pairs to be converted as a parameter of the 'GetDictionaryJson' method is null or empty.",
            };
        }

        internal static string JsonError()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Ocorreu um erro ao processar os itens no formato JSON do método 'GetDictionaryJson'.",
                _ => $"An error occurred while processing items in JSON format from the 'GetDictionaryJson' method.",
            };
        }

        internal static string JsonDicNull()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O dicionário a ser serializado em JSON a partir do método 'GetDictionaryJson' é nulo ou vazio.",
                _ => $"The dictionary to be serialized into JSON from the 'GetDictionaryJson' method is null or empty.",
            };
        }

        internal static string JsonDicError()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Ocorreu um erro ao serializar o dicionário para JSON em 'GetJsonDictionary'.",
                _ => $"An error occurred while serializing the dictionary to JSON in 'GetJsonDictionary'.",
            };
        }

        #endregion

        #region Unzip

        internal static string UnableUnzip(string pathFileZip)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Não é possível extrair o arquivo '{Path.GetFileNameWithoutExtension(pathFileZip)}' pois '{Path.GetExtension(pathFileZip).ToUpper()}' ainda não é suportado!",
                _ => $"Unable to extract this file '{Path.GetFileNameWithoutExtension(pathFileZip)}' because '{Path.GetExtension(pathFileZip).ToUpper()}' isn't yet supported!",
            };
        }



        #endregion






    }
}
