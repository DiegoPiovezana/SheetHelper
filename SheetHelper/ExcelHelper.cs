using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

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
        /// Recebe as linhas em string, e retorna um vetor de inteiros com a primeira e última linha
        /// </summary>
        public static int[] DefineRows(string rows, int limitRows)
        {
            if (rows == null || rows.Equals("")) // Se linhas não especificadas           
                return new[] { 1, limitRows }; // Converte todas as linhas
                                               
            rows = rows.Trim();

            if (!rows.Contains(":"))
                throw new Exception("Use a ':' to define the first and last row!");

            string[] rowsArray = rows.Split(':');

            if (rowsArray[0].Equals("")) // Se primeira linha não definida
                rowsArray[0] = "1"; // Então, deseja-se converter desde a primeira linha

            if (rowsArray[1].Equals("")) // Se última linha não definida
                rowsArray[1] = limitRows.ToString(); // Então, deseja-se converter até a última linha


            if (Int32.Parse(rowsArray[0]) < 1 || Int32.Parse(rowsArray[0]) > limitRows)
                throw new Exception($"A primeira linha está fora do limite (min 1, max {limitRows})!");

            if (Int32.Parse(rowsArray[1]) < 1 || Int32.Parse(rowsArray[1]) > limitRows)
                throw new Exception($"Última linha fora dos limites (min 1, max {limitRows})!");

            if (Int32.Parse(rowsArray[0]) > Int32.Parse(rowsArray[1]))
                throw new Exception($"A linha inicial deve ser menor que a última linha!");


            return new[] {Int32.Parse(rowsArray[0]), Int32.Parse(rowsArray[1])};

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
        /// <param name="columns">"Vetor de caracteres (maiúsculo ou minúsculo) contendo todas as colunas desejadas. Ex.: "{ 'A', 'b', 'E', 'C' }"</param>
        /// <param name="rows">"Informe a primeira e última linha (ou deixe em branco). Ex.: "1:50 (linha 1 até linha 50)"</param>
        /// <param name="pgbar">"Caso desejado, passe uma ProgressBar para ser carregada em 100 (ou null). Ex.: "ProgressBar pgbar = new ProgressBar()"</param>
        /// <returns>"true" se convertido com sucesso. "false" se não convertido.</returns>
        public static bool Converter(string origin, string destiny, int sheet, string separator, string[] columns, string rows, ProgressBar pgbar)
        {
            // TODO: Corrigir inclusão de cabeçalho ao converter de CSV
            // TODO: Última linha não convertida 
            // TODO: Verificar incrementador




            if (separator != null)
                separator = separator.Trim();
            else
                throw new Exception("Separador inválido!");

            if (pgbar == null)
                pgbar = new ProgressBar();

            // Abre o arquivo
            using (var stream = File.Open(origin, FileMode.Open, FileAccess.Read))
            {

                File.WriteAllText(destiny, ""); // Para verificar se arquivo de destino esta acessivel
                File.Delete(destiny); // Deleta para evitar que usuario abra o arquivo durante a conversao
                pgbar.Value += 5; // 5       

                DataSet result = null;

                // Realiza a leitura do arquivo
                switch (Path.GetExtension(origin)) // TODO: txt?
                {
                    case ".csv": // .csv 
                        result = ReadCSV(stream);
                        break;

                    default: // .xlsx, .xls, .xlsb, .xlsm
                        result = ReadXLS(stream);
                        break;
                }

                pgbar.Value += 30; // 35 (pós leitura do arquivo)

                // Se existir abas na planilha e a desejada estiver correta
                if (result.Tables.Count > 0 && sheet > -1 && sheet < result.Tables.Count)
                {
                    StringBuilder output = new StringBuilder();

                    // Obtem a aba desejada
                    DataTable table = result.Tables[sheet];
                    pgbar.Value += 5; // 40

                    // Define qual será a primeira e última linha a ser convertida
                    int[] rowsNumber = ExcelHelper.DefineRows(rows, table.Rows.Count + 1);
                    pgbar.Value += 5; // 45                

                    List<string> row = null;
                    int[] columnsASCII = null;

                    // Se deseja incluir cabeçalho
                    if (rowsNumber[0] == 1)
                    {
                        var colunsData = table.Columns.Cast<DataColumn>().ToList(); // Salva cabeçalho
                        row = new List<string>(colunsData.Count);

                        foreach (var item in colunsData) // Realiza a conversão das Listas
                            row.Add(item.ToString());
                    }
                    else
                        row = table.Rows[rowsNumber[0] - 2].ItemArray.Select(f => f.ToString()).ToList(); // linha 2 primeira => index é 1 (-1) e cabeçalho ja retirado (-1)

                    // Se deseja selecionar colunas específicas
                    if (columns != null && columns.Length != 0)
                    {
                        columnsASCII = DefineColunms(columns);
                    }

                    pgbar.Value += 5; // 50 (tratativas)

                    double percPrg = 40.0 / (rowsNumber[1] - rowsNumber[0] + 1); // Percentual a ser progredido a cada linha da planilha

                    // Salva todas as demais linhas mediante início e fim         
                    for (int i = rowsNumber[0]; i < rowsNumber[1]; i++) // Para cada linha da planilha
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
                                rowSelected.Append(row[column]).Append(separator); //rowSelected.Append(row[Convert.ToInt32(Char.ToUpper(column)) - 65]).Append(separator);
                            }

                            output.AppendLine(String.Join(separator, rowSelected)); // Adiciona a linha com as colunas selecionadas                            
                        }

                        if (percPrg >= 1) // Se aplicável, carrega
                        {
                            pgbar.Value += (int)percPrg; // 90
                            percPrg -= (int)percPrg;
                        }
                        else
                        {
                            percPrg += percPrg;
                        }

                        // Obtem a próxima linha
                        row = table.Rows[i - 1].ItemArray.Select(f => f.ToString()).ToList();
                    }

                    pgbar.Value += (90 - pgbar.Value);

                    // Escreve o novo arquivo convertido (substitui se ja existente)
                    File.WriteAllText(destiny, output.ToString());
                    pgbar.Value += 10; // 100
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
        /// <param name="columns">"Vetor de caracteres (maiúsculo ou minúsculo) contendo todas as colunas desejadas. Ex.: "{ 'A', 'b', 'E', 'C' }. Passe null ou um vetor vazio caso precise de todas as colunas convertidas"</param>
        /// <param name="rows">"Informe a primeira e última linha (ou deixe em branco). Ex.: "1:50 (linha 1 até linha 50)"</param>
        /// <param name="pgbar">"Caso desejado, passe uma ProgressBar para ser carregada em 100 (ou null). Ex.: "ProgressBar pgbar = new ProgressBar()"</param>
        /// <returns>"true" se convertido com sucesso. "false" se não convertido.</returns>
        public static bool ConverterExcept(string origin, string destiny, int sheet, string separator, string[] columns, string rows, ProgressBar pgbar)
        {

            int countOpen = 0; // Contagem de vezes que o Excel estava aberto

        again:

            try
            {
                return Converter(origin, destiny, sheet, separator, columns, rows, pgbar);
            }


            #region Se arquivo nao localizado        
            catch (FileNotFoundException nffEx) when (nffEx.HResult.Equals(-2147024894))
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
            catch (IOException eiEx) when (eiEx.HResult.Equals(-2147024864))
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
            catch (ExcelDataReader.Exceptions.HeaderException heEx) when (heEx.HResult.Equals(-2147024894))
            {

                throw new Exception($"Erro! Sem suporte para converter o arquivo '{Path.GetExtension(origin)}'.");

            }
            #endregion


        }
    }
}

