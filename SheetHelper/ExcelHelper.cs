﻿using ExcelDataReader;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace SheetHelper
{
    /// <summary>
    /// Biblioteca rápida e leve, para fácil conversão de grandes arquivos Excel
    /// </summary>
    public static class ExcelHelper
    {
        /// <summary>
        /// Encerra todos os processos do Excel
        /// </summary>
        public static void CloseExcel()
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes) process.Kill();

        }

        /// <summary>
        /// Recebe as colunas em string, converte para caracter e retorna em ASCII correspondente
        /// </summary>
        private static int[] DefineColunms(string[] columns)
        {
            int[] columnsASCII = new int[columns.Length];

            for (int i = 0; i < columns.Length; i++) // Para cada coluna
            {
                string column = columns[i].ToUpper();
                int sum = 0;

                foreach (var character in column)
                {
                    sum *= 26;
                    sum += (character - 'A' + 1);
                }
                columnsASCII[i] = sum - 1;

            }

            return columnsASCII;
        }

        /// <summary>
        /// Realiza a leitura de arquivos .xls, .xlsx e .xlsb
        /// </summary>
        private static DataSet ReadXLS(FileStream stream)
        {
            // Formato de detecção automática, suporta: 
            //   - Arquivos Excel binários (formato 2.0-2003; *.xls) 
            //   - Arquivos Excel OpenXml (formato 2007; *.xlsx, *.xlsb)
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }

                });

                return result;
            }
        }

        /// <summary>
        /// Realiza a leitura de arquivos .csv
        /// </summary>
        private static DataSet ReadCSV(FileStream stream)
        {
            // Realiza a leitura do arquivo Excel CSV
            using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
            {
                DataSet result = reader.AsDataSet();

                return result;
            }
        }



        /// <summary>
        /// Realiza a conversão do arquivo Excel localizado em <paramref name="origin"/>, salva em <paramref name="destiny"/>
        /// e retorna 'true' caso a conversão tenha ocorrido com sucesso.
        /// Utilize o método "ConverterExcept" para realizar a conversão e tratar algumas exceções!
        /// </summary>
        /// <param name="origin">Diretorio + nome do arquivo de origem + formato. Ex.: "C:\\Users\\ArquivoExcel.xlsx"</param>
        /// <param name="destiny">Diretorio + nome do arquivo de destino + formato. Ex.: "C:\\Users\\ArquivoExcel.csv"</param>
        /// <param name="sheet">Aba da planilha a ser convertida. Ex.: 1 (segunda aba)</param>
        /// <param name="separator">Separador a ser utilizado para realizar a conversão. Ex.: ";"</param>
        /// <param name="header">Tamanho do cabeçalho a ser retirado (0 para manter a primeira linha). Ex.: 1</param>
        /// <param name="columns">"Vetor de caracteres (maiúsculo ou minúsculo) contendo todas as colunas desejadas. Ex.: "{ 'A', 'b', 'C', 'E' }"</param>
        /// <returns>"true" se convertido com sucesso. "false" se não convertido.</returns>
        public static bool Converter(string origin, string destiny, int sheet, string separator, int header, string[] columns)
        {

            // Abre o arquivo
            using (var stream = File.Open(origin, FileMode.Open, FileAccess.Read))
            {

                File.WriteAllText(destiny, ""); // Para verificar se arquivo de destino esta acessivel
                File.Delete(destiny); // Deleta para evitar que usuario abra o arquivo durante a conversao


                DataSet result = null;
                string format = Path.GetExtension(origin);

                // .xlsx, .xls, .xlsb, .csv ou .txt;
                switch (Path.GetExtension(origin))
                {
                    case ".csv":
                        result = ReadCSV(stream);
                        break;

                    default: // .xlsx, .xls, .xlsb
                        result = ReadXLS(stream);
                        break;
                }


                // Se existir abas na planilha e a desejada estiver correta
                if (result.Tables.Count > 0 && sheet > -1 && sheet < result.Tables.Count)
                {
                    StringBuilder output = new StringBuilder();

                    // Obtem a aba desejada
                    DataTable table = result.Tables[sheet];

                    // Se deseja incluir cabeçalho (primeira linha)
                    if (header == 0)
                        // Salva os nomes das colunas
                        output.AppendLine(String.Join(separator, table.Columns.Cast<DataColumn>().ToList()));
                    else
                        // Remove as demais linhas do cabeçalho de acordo com a quantidade desejada
                        for (int i = 1; i < header; i++)
                        { // A partir da 2ª linha
                            DataRow dr = table.Rows[0]; // Remove a linha do topo
                            table.Rows.Remove(dr);
                        }

                    // Salva todas as linhas não removidas
                    foreach (DataRow dr in table.Rows)
                    {
                        var row = dr.ItemArray.Select(f => f.ToString()).ToList();

                        if (columns == null || columns.Length == 0) // Se colunas nao especificadas
                        {
                            output.AppendLine(String.Join(separator, row)); // Adiciona toda as colunas
                        }

                        else
                        {
                            StringBuilder rowSelected = new StringBuilder(); // Armazena as colunas selecionadas da linha
                            int[] columnsASCII = DefineColunms(columns);

                            foreach (int column in columnsASCII) // Para cada coluna
                            {

                                // Seleciona a coluna considerando tabela ASCII e adiciona separadamente                                
                                rowSelected.Append(row[column]).Append(separator); //rowSelected.Append(row[Convert.ToInt32(Char.ToUpper(column)) - 65]).Append(separator);

                            }

                            output.AppendLine(String.Join(separator, rowSelected)); // Adiciona a linha com as colunas selecionadas

                        }

                    }

                    // Escreve o novo arquivo convertido (substitui se ja existente)
                    File.WriteAllText(destiny, output.ToString());
                    return true;
                }
                else
                {
                    throw new Exception("Erro ao selecionar a aba desejada");
                }
            }
        }


        /// <summary>
        /// Realiza a conversão do arquivo Excel localizado em <paramref name="origin"/>, salva em <paramref name="destiny"/>
        /// com tratativa de exceçoes para o usuário final (arquivo inexistente no diretorio ou aberto durante a conversão)
        /// e retorna 'true' caso a conversão tenha ocorrido com sucesso
        /// </summary>
        /// <param name="origin">Diretorio + nome do arquivo de origem + formato. Ex.: "C:\\Users\\ArquivoExcel.xlsx"</param>
        /// <param name="destiny">Diretorio + nome do arquivo de destino + formato. Ex.: "C:\\Users\\ArquivoExcel.csv"</param>
        /// <param name="sheet">Aba da planilha a ser convertida. Ex.: 1 (segunda aba)</param>
        /// <param name="separator">Separador a ser utilizado para realizar a conversão. Ex.: ";"</param>
        /// <param name="header">Informe o tamanho do cabeçalho a ser retirado (0 para manter a primeira linha). Ex.: 1</param>
        /// <param name="columns">"Vetor de caracteres (maiúsculo ou minúsculo) contendo todas as colunas desejadas. Ex.: "{ 'A', 'b', 'C', 'E' }. Passe null ou um vetor vazio caso precise de todas as colunas convertidas"</param>
        /// <returns>"true" se convertido com sucesso. "false" se não convertido.</returns>
        public static bool ConverterExcept(string origin, string destiny, int sheet, string separator, int header, string[] columns)
        {

            int countOpen = 0; // Contagem de vezes que o Excel estava aberto

        again:

            try
            {
                return Converter(origin, destiny, sheet, separator, header, columns);
            }


            #region Se arquivo nao localizado            
            catch (FileNotFoundException fnfEx) when (fnfEx.ToString().Contains("not find file") || fnfEx.ToString().Contains("arquivo não encontrado"))
            {
                var result3 = MessageBox.Show(
                                   "O arquivo '" + Path.GetFileName(origin) + "' não foi localizado. Por favor, verifique se o arquivo está presente no repositório de origem e confirme para continuar: "
                                   + "\n\n" + origin,
                                   "Aviso",
                                   MessageBoxButtons.OKCancel,
                                   MessageBoxIcon.Exclamation);


                if (result3 == DialogResult.OK)
                {
                    goto again; // Tenta realizar a conversão novamente
                }

                return false;
            }
            #endregion

            #region Se arquivo esta em uso
            catch (IOException eiEx) when (eiEx.ToString().Contains("being used by another process") || eiEx.ToString().Contains("sendo usado por outro processo"))
            {
                countOpen++; // Contador de tentativas com falha de arquivo aberto

                if (countOpen >= 2) // Se necessario forçar o fechamento do Excel (a partir do 2 caso)
                {

                    var result1 = MessageBox.Show(
                       "Parece que o arquivo ainda continua em uso. Deseja forçar o encerramento do Excel e tentar novamente? \n\nAs alterações não serão salvas.",
                       "Aviso",
                       MessageBoxButtons.YesNo,
                       MessageBoxIcon.Exclamation);

                    if (result1 == DialogResult.Yes)
                    {
                        CloseExcel(); // Encerra todos os processos do Excel
                        Thread.Sleep(1500); // Aguarda o excel fechar por completo durante 1,5 segundo
                        goto again; // Tenta realizar a conversão novamente

                    } // Se No, continua execucao abaixo

                }

                var result2 = MessageBox.Show(
                    "O arquivo '" + Path.GetFileName(origin) + "' ou '" + Path.GetFileName(destiny) + "' está sendo utilizado em outro processo. Por favor, finalize seu uso e em seguida confirme para continuar.",
                    "Aviso",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Error);


                if (result2 == DialogResult.OK)
                {
                    goto again; // Tenta realizar a conversão novamente
                }
                else // Se cancelar
                {
                    return false;
                }


            }

            #endregion

            #region Se arquivo em formato não suportado
            catch (ExcelDataReader.Exceptions.HeaderException heEx) when (heEx.ToString().Contains("Invalid file signature") || heEx.ToString().Contains("Assinatura de arquivo inválida"))
            {
                throw new Exception($"Erro! Sem suporte para converter arquivos de origem '{Path.GetExtension(origin)}'.");

            }
            #endregion


        }
    }
}
