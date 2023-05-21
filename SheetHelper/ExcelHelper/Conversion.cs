using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
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

        private static int _i;
        private static int _j;

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
        public static bool Converter(string origin, string destiny, string sheet, string separator, string[] columns, string rows)
        {
            Progress = 0;
            Treatment.ValidateString(new string[] { origin, destiny, sheet, separator});

            File.WriteAllText(destiny, "");  // Para verificar se arquivo de destino esta acessivel
            File.Delete(destiny); // Deleta para evitar que usuario abra o arquivo durante a conversao

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Progress += 5; // 5 

            DataSet result = Reading.GetDataSet(origin, destiny);

            Progress += 30; // 35 (pós leitura do arquivo)

            // Obtem a aba a ser convertida
            DataTable table = Reading.GetTable(sheet, result);

            StringBuilder output = new StringBuilder();

            Progress += 5; // 40

            // Define qual será a primeira e última linha a ser convertida
            int[] rowsNumber = Treatment.DefineRows(rows, table.Rows.Count + 1);
            Progress += 5; // 45                


            int[] columnsASCII = null;
            _i = rowsNumber[0]; // Primeira linha a ser convertida
            _j = 0; // Deslocamento

            List<string> row = Reading.GetFirstRow(Path.GetExtension(origin), table);


            // Se deseja selecionar colunas específicas
            if (columns != null && columns.Length != 0) // null OR {}
            {
                if (columns[0].Contains(":"))
                { // Se primeira celula do array. E.g.: {"A:G"}
                    columnsASCII = Treatment.DefineColunms(columns[0], SH.GetNameColumn(row.Count()));
                }
                else
                { // Se colunas definidas individualmente. E.g.: {"A", "B"}
                    columnsASCII = Treatment.DefineColunms(columns);
                }

            }

            Progress += 5; // 50 (tratativas)

            double countPercPrg = 40.0 / (rowsNumber[1] - rowsNumber[0] + 1); // Percentual a ser progredido a cada linha da planilha
            double percPrg = countPercPrg;

            table.Rows.Add(); // Para evitar IndexOutOfRangeException (última linha será ignorada)


            // Salva todas as demais linhas mediante início e fim         
            for (; _i <= rowsNumber[1] + _j; _i++) // Para cada linha da planilha
            {
                if (columnsASCII == null) // Se colunas não especificadas
                {
                    output.AppendLine(String.Join(separator, row)); // Adiciona toda as colunas da linha                            
                }

                else // Se colunas especificadas
                {
                    StringBuilder rowSelected = new StringBuilder(); // Armazena as colunas selecionadas da linha                            

                    foreach (int column in columnsASCII) // Para cada coluna das linhas
                    {
                        // Seleciona a coluna considerando tabela ASCII e adiciona separadamente                               
                        rowSelected.Append(row[column - 1]).Append(separator); //rowSelected.Append(row[Convert.ToInt32(Char.ToUpper(column)) - 65]).Append(separator);
                    }

                    output.AppendLine(String.Join(separator, rowSelected)); // Adiciona a linha com as colunas selecionadas                            
                }

                if (countPercPrg >= 1) // Se aplicável, carrega a ProgressBar
                {
                    Progress += (int)countPercPrg; // 90                                                               
                    countPercPrg -= (int)countPercPrg;
                }

                countPercPrg += percPrg; // Incrementa contador da ProgressBar                        

                // Obtem a próxima linha
                row = table.Rows[_i - 1].ItemArray.Select(f => f.ToString()).ToList();
            }



            Progress += (90 - Progress); // Se necessário, completa até 90%

            // Escreve o novo arquivo convertido (substitui se ja existente)
            File.WriteAllText(destiny, output.ToString());
            Progress += 10; // 100
            return true;
        }



    }
}

