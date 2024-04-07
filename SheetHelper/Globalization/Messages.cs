﻿using System.Diagnostics;
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
                "pt-BR" => $"Por favor, visite https://bit.ly/SheetHelper_Exceptions para saber mais.",
                _ => $"Please visit https://bit.ly/SheetHelper_Exceptions to learn more.",
            };
        }

        internal static string ArgumentException(string argument, string method)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O parâmetro '{argument}' do método '{method}' não é válido! Por favor, verifique se está preenchido adequadamente. Considere consultar a documentação disponível em https://bit.ly/SheetHelper.",
                _ => $"The '{argument}' parameter of the '{method}' method is not valid! Please check that it is filled out properly. Consider consulting the documentation available at https://bit.ly/SheetHelper.",
            };
        }

        internal static string UnmappedException(string nameMethod, System.Exception ex)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Um erro ocorreu ao utilizar o método '{nameMethod}'! Por favor, verifique se os parâmetros utilizados estão corretos. Se desejar, reporte seu erro no card 'Unmapped Error' disponível em  https://bit.ly/SheetHelper_Exceptions. \n\n{ex.Message}",
                _ => $"An error occurred when using the '{nameMethod}' method! Please check that the parameters entered are correct. If you wish, report your error in the 'Unmapped Error' card available at https://bit.ly/SheetHelper_Exceptions. \n\n{ex.Message}",
            };
        }


        #endregion

        #region Sheet

        internal static string SheetIndexNotFound(int sheetIndex)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Erro ao selecionar a aba desejada! Verifique se o índice '{sheetIndex}' da aba está correto.",
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

        internal static string SheetRowColumnNumberSame()
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"A quantidade de abas, colunas e linhas definidas deve ser igual.",
                _ => $"The number of tabs, columns and lines defined must be the same.",
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

        #region Rows

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
                "pt-BR" => $"O tamanho do vetor excede o número de colunas da dataTable!",
                _ => $"The length of the row array exceeds the number of columns in the dataTable!",
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
                "pt-BR" => $"O caminho '{pathFile}' não é válido!",
                _ => $"The path '{pathFile}' is not valid!",
            };
        }

        #endregion

        #region FileOrigin

        internal static string FileOriginInUse(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O arquivo de origem '{Path.GetFileName(pathFile)}' não pode ser lido pois está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida tente novamente.",
                _ => $"The origin file '{Path.GetFileName(pathFile)}' cannot be read because it is being used by another process. Please finish its use and then try again.",
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

        internal static string FileOriginNotReadSupport(string pathOrigin)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Sem suporte para realizar a leitura do arquivo com formato '{Path.GetExtension(pathOrigin)}'.",
                _ => $"No support for reading the file in '{Path.GetExtension(pathOrigin)}' format.",
            };
        }

        #endregion

        #region FileDestiny

        internal static string FileDestinyInUse(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"O arquivo de destino '{Path.GetFileName(pathFile)}' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida tente novamente",
                _ => $"The destiny file '{Path.GetFileName(pathFile)}' is being used by another process. Please finish its use and then try again",
            };
        }

        internal static string FileDestinyInUseAndCloseExcel(string pathFile)
        {
            return CultureInfo.CurrentCulture.Name switch
            {
                "pt-BR" => $"Parece que o arquivo '{Path.GetFileName(pathFile)}' ainda continua em uso. Deseja forçar o encerramento de todos os Excel e tentar novamente? \n\nATENÇÃO: Todos os Excel abertos serão fechados e as alterações da(s) planilha(s) não serão salvas!",
                _ => $"It appears that the file '{Path.GetFileName(pathFile)}' is still in use. Do you want to force quit all Excel and try again? \n\nATTENTION: All open Excel will be closed and changes to the worksheet(s) will not be saved!",
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
