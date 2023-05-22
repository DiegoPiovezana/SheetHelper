using ExcelDataReader.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace SheetHelper
{
    /// <summary>
    /// Fast and lightweight library for easy conversion of large Excel files
    /// </summary>
    internal static class Conversion
    {
        /// <summary>
        /// Represents the conversion progress. If 100%, the conversion is fully completed.
        /// </summary>
        public static int Progress { get; set; }

        //private static int _i;
        //private static int _j;

        /// <summary>
        /// Realiza a conversão do arquivo Excel localizado em <paramref name="origin"/>, salva em <paramref name="destiny"/>
        /// e retorna 'true' caso a conversão tenha ocorrido com sucesso.
        /// Utilize o método "ConverterExcept" para realizar a conversão e tratar algumas exceções!
        /// </summary>
        /// <param name="origin">Diretorio + nome do arquivo de origem + formato. E.g.: "C:\\Users\\ArquivoExcel.xlsx"</param>
        /// <param name="destiny">Diretorio + nome do arquivo de destino + formato. E.g.: "C:\\Users\\ArquivoExcel.csv"</param>
        /// <param name="sheet">Aba da planilha a ser convertida. E.g.: "1" (primeira aba) ou "NomeAba"</param>
        /// <param name="separator">Separador a ser utilizado para realizar a conversão. E.g.: ";"</param>
        /// <param name="columns">"Vetor de caracteres (maiúsculo ou minúsculo) contendo todas as colunas desejadas. E.g.: { "A", "b", "E", "C" } ou "{ "A:BC" } </param>
        /// <param name="rows">"Informe a primeira e última linha (ou deixe em branco). E.g.: "1:50 (linha 1 até linha 50)"</param>        
        /// <returns>"true" se convertido com sucesso. "false" se não convertido.</returns>
        public static bool Converter(string origin, string destiny, string sheet, string separator, string columns, string rows)
        {
            Progress = 0;

            Treatment.Validate(origin, destiny, sheet, separator, columns, rows);
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Progress += 5; // 5 

            DataSet result = Reading.GetDataSet(origin, destiny);
            Progress += 30; // 35 (pós leitura do arquivo)

            // Obtem a aba a ser convertida
            DataTable table = Reading.GetTable(sheet, result);
            StringBuilder output = new();
            Progress += 5; // 40

            // Define o número de todas as linhas a serem consideradas
            int[] rowsNumber = Treatment.DefineRows(rows, table);
            Progress += 5; // 45                

            // Define em ASCII, quais serão todas as colunas a serem convertidas
            int[] columnsASCII = Treatment.DefineColumnsASCII(columns, table);
            Progress += 5; // 50 (tratativas ok)

            //List<string> rowFull = Reading.GetFirstRow(Path.GetExtension(origin), table);

            double countPercPrg = 40.0 / (rowsNumber[1] - rowsNumber[0] + 1); // Percentual a ser progredido a cada linha da planilha
            double percPrg = countPercPrg;

            table.Rows.Add(); // Para evitar IndexOutOfRangeException (última linha será ignorada)

            //_i = rowsNumber[0]; // Primeira linha a ser convertida
            //_j = 0; // Deslocamento

            //using (StreamWriter writer = new (destiny))
            //{
            // Salva todas as demais linhas mediante início e fim   
            foreach (int rowIndex in rowsNumber) // Para cada linha da planilha  
            {
                // Obter a linha               
                string[] rowFull = table.Rows[rowIndex - 1].ItemArray.Select(cell => cell.ToString()).ToArray();
               
                if (columnsASCII.Equals("full")) // Se colunas não especificadas - Todas 
                {
                    output.AppendLine(String.Join(separator, rowFull)); // Adiciona toda as colunas da linha
                                                                        //writer.Write(String.Join(separator, rowFull));                       
                }
                else // Se colunas especificadas - Selecionadas
                {
                    StringBuilder rowSelected = new(); // Armazena as colunas selecionadas da linha                            

                    foreach (int column in columnsASCII) // Para cada coluna das linhas
                    {
                        // Seleciona a coluna considerando tabela ASCII e adiciona separadamente                               
                        rowSelected.Append(rowFull[column - 1]).Append(separator);
                    }
                    output.AppendLine(String.Join(separator, rowSelected)); // Adiciona a linha com as colunas selecionadas                            
                    //writer.Write(String.Join(separator, rowSelected));                    
                }

                if (countPercPrg >= 1) // Se aplicável, carrega o progresso
                {
                    Progress += (int)countPercPrg; // 90                                                               
                    countPercPrg -= (int)countPercPrg;
                }

                countPercPrg += percPrg; // Incrementa contador de progresso                        

                // Obtem a próxima linha
                //List<string> rowFull = table.Rows[_i - 1].ItemArray.Select(f => f.ToString()).ToList();
                //writer.WriteLine();
            }

            Progress += (90 - Progress); // Se necessário, completa até 90%

            // Escreve o novo arquivo convertido (substitui se ja existente)
            File.WriteAllText(destiny, output.ToString());
            //}

            Progress += 10; // 100
            return true;
        }



    }
}

