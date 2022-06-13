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
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace SheetHelper
{
    /// <summary>
    /// Biblioteca rápida e leve, para fácil conversão de grandes arquivos Excel
    /// </summary>
    public static class ExcelHelper
    {
        private static int _i;
        private static int _j;




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
        /// Recebe o nome da coluna e retorna o índice na planilha
        /// </summary>
        /// <param name="columnName">Nome da coluna. Ex.: "A"</param>
        /// <returns>Índice</returns>
        public static int GetIndexColumn(string columnName)
        {
            int sum = 0;

            foreach (var character in columnName)
            {
                sum *= 26;
                sum += (character - 'A' + 1);
            }

            return sum; // Ex.: A = 1, Z = 26, AA = 27
        }

        /// <summary>
        /// Obtem o nome (ex.: "AB") da coluna mediante índice
        /// </summary>
        /// <param name="columnNumber"> Índice da coluna</param>
        /// <returns></returns>
        public static string GetNameColumn(int columnNumber)
        {
            string columnName = String.Empty;

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }


        /// <summary>
        /// Recebe as colunas em string, converte para caracter e retorna em ASCII correspondente
        /// </summary>
        private static int[] DefineColunms(string[] columns)
        {
            int[] columnsASCII = new int[columns.Length];

            for (int i = 0; i < columns.Length; i++) // Para cada coluna
            {
                columnsASCII[i] = GetIndexColumn(columns[i].ToUpper());
            }

            return columnsASCII;
        }

        /// <summary>
        /// Define o índice de todas as colunas que serão convertidas
        /// </summary>
        /// <param name="columns">Colunas a serem convertidas. Ex.: "B:H"</param>
        /// <param name="lastColumnName">Última coluna da planilha (limite máximo de colunas). Ex.: "AZ"</param>
        /// <returns>Vetor com todos os índices das colunas a serem convertidas</returns>
        /// <exception cref="Exception"></exception>
        private static int[] DefineColunms(string columns, string lastColumnName)
        {
            int indexLastColumn = GetIndexColumn(lastColumnName);

            if (columns == null || columns.Equals("")) // Se coluna não especificadas           
                return new[] { 0, indexLastColumn }; // Converte todas as colunas

            columns = columns.Trim();

            if (!columns.Contains(":"))
                throw new Exception("Use um ':' para definir a primeira e última coluna!");

            string[] rowsArray = columns.Split(':'); // Ex.: {"A","Z"}

            if (rowsArray[0].Equals("")) // Se primeira coluna não definida
                rowsArray[0] = "A"; // Então, deseja-se converter desde a primeira coluna

            if (rowsArray[1].Equals("")) // Se última coluna não definida
                rowsArray[1] = lastColumnName; // Então, deseja-se converter até a última coluna


            if (GetIndexColumn(rowsArray[0]) <= 0 || GetIndexColumn(rowsArray[0]) > indexLastColumn)
                throw new Exception($"A primeira coluna está fora do limite (min A, max {lastColumnName})!");

            if (GetIndexColumn(rowsArray[1]) <= 0 || GetIndexColumn(rowsArray[1]) > indexLastColumn)
                throw new Exception($"Última coluna fora dos limites (min A, max {lastColumnName})!");

            if (GetIndexColumn(rowsArray[0]) > GetIndexColumn(rowsArray[1])) // Para este caso (incrementador)
                throw new Exception($"A coluna inicial deve vir antes da última coluna!");

            int firstColumn = GetIndexColumn(rowsArray[0]);
            int lastColumn = GetIndexColumn(rowsArray[1]);
            int[] columnsIndex = new int[lastColumn - firstColumn + 1];
            int count = 0;

            // Preenche o vetor com os índices de todas as colunas que serão convertidas
            for (int index = firstColumn; index <= lastColumn; index++)
            {
                columnsIndex[count] = index;
                count++;
            }

            return columnsIndex;
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


            return new[] { Int32.Parse(rowsArray[0]), Int32.Parse(rowsArray[1]) };

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
        /// Descompacta arquivo .GZ
        /// </summary>
        /// <param name="compressedFileStream">Arquivo  a ser convertido obtido através do método File.Open. </param> 
        /// <param name="pathDestiny">Diretório onde será salvo o arquivo descompactado (contendo OU NAO o nome do arquivo destino). Ex.: 'C:\\Arquivos\\ ou 'C:\\Arquivos\\Convertido.xlsx'</param>
        public static string UnGZ(FileStream compressedFileStream, string pathDestiny)
        {
            string fileConverted;

            if (Path.GetExtension(pathDestiny) == "") // Se formato a ser convertido não especificado, tenta obter do nome
            {
                string originalFileName = Path.GetFileName(compressedFileStream.Name).Replace(".gz", "").Replace(".GZ", "");
                string formatOriginal = Regex.Match(Path.GetExtension(originalFileName), @"\.[A-Za-z]*").Value;
                fileConverted = $"{pathDestiny}{Path.GetFileNameWithoutExtension(originalFileName)}{formatOriginal}";
            }
            else
            {
                fileConverted = pathDestiny;
            }

            //FileStream compressedFileStream = File.Open(compressedFileName, FileMode.Open); // "compressed.xlsx.gz"
            FileStream outputFileStream = File.Create(fileConverted); // "decompressed.xlsx"
            var decompressor = new GZipStream(compressedFileStream, CompressionMode.Decompress);
            decompressor.CopyTo(outputFileStream);

            // Encerra uso dos arquivos
            compressedFileStream.Close();
            outputFileStream.Close();

            return File.Exists(fileConverted) ? fileConverted : null;
        }

        /// <summary>
        /// Descompacta um arquivo .ZIP
        /// </summary>
        /// <param name="zipFile">Local e nome do arquivo compactado. Ex.: 'C:\\Arquivos\\Relatorio.zip</param>
        /// <param name="pathDestiny">Diretório onde será salvo o arquivo descompactado. Ex.: 'C:\\Arquivos\\'</param>
        /// <returns></returns>
        public static string UnZIP(string zipFile, string pathDestiny)
        {
            string directoryZIP = $"{pathDestiny}\\CnvrtdZIP\\";
            string directoryDestiny = pathDestiny;

            // Realiza a extração para um novo diretório
            ZipFile.ExtractToDirectory(zipFile, directoryZIP);

            IEnumerable<string> files = Directory.EnumerateFiles(directoryZIP);
            string fileLocation = files.First(); // Obtem o local do arquivo 
            string fileDestiny = $"{directoryDestiny}\\{Path.GetFileName(fileLocation)}"; // Local destinatário do arquivo

            if (File.Exists(fileDestiny)) // Se arquivo existente, apaga
                File.Delete(fileDestiny);

            File.Move(fileLocation, fileDestiny); // Move-o para o local de destino            
            Directory.Delete(directoryZIP); // Deleta o diretorio criado anteriormente

            return $"{directoryDestiny}\\{Path.GetFileName(fileLocation)}";
        }


        //public static string GetColumns(int row, bool fill)
        //{

        //}

        private static bool ValidateString(string[] strings)
        {
            foreach (string str in strings)
            {
                if (string.IsNullOrEmpty(str))
                    throw new Exception($"'{str}' inválido!");
            }
            return true;
        }

        /// <summary>
        /// Obtem a aba desejada
        /// </summary>
        /// <param name="sheet">Nome ou índice da aba desejada</param>
        /// <param name="result">Dataset da planilha</param>    
        /// <exception cref="Exception">Erro ao localizar aba</exception>
        private static DataTable GetTable(string sheet, DataSet result)
        {
            if (int.TryParse(sheet, out int sh)) // Se informado índice da aba desejada
            {
                // Se existir abas na planilha e a desejada estiver correta
                if (result.Tables.Count <= 0 || sh <= -1 || sh > result.Tables.Count)
                {
                    throw new Exception("Erro ao selecionar a aba desejada! Verifique se o índice da aba está correto.");
                }

                return result.Tables[sh - 1]; // Obtem a aba desejada

            } // Se nome da aba for informado
            else
            {
                if (!result.Tables.Contains(sheet)) // Se nome da aba não localizado
                {
                    throw new Exception($"Não foi possível encontrar a aba '{sheet}' desejada! Verifique se o nome da aba está correto.");
                }

                return result.Tables[sheet]; // Obtem a aba desejada
            }
        }

        private static List<string> GetFirstRow(string extension, DataTable table)
        {
            List<string> row;

            if (!extension.Equals(".csv") && !extension.Equals(".rpt") && !extension.Equals(".txt")) // A tratativa para o cabeçalho csv é diferente
            { // Se não for CSV

                // Se deseja incluir cabeçalho
                if (_i == 1)
                {
                    var colunsData = table.Columns.Cast<DataColumn>().ToList(); // Salva cabeçalho
                    row = new List<string>(colunsData.Count);

                    foreach (var item in colunsData) // Realiza a conversão das Listas
                        row.Add(item.ToString());
                }
                else // Se não deseja incluir cabeçalho
                {
                    row = table.Rows[_i - 2].ItemArray.Select(f => f.ToString()).ToList(); // linha 2 primeira => index é 1 (-1) e cabeçalho ja retirado (-1)
                }

            }
            else // Se leitura CSV, elimina cabeçalho 'Column' e considera index 0
            {
                //if (!extension.Equals(".csv"))                       
                if (_i == table.Rows.Count + 1) // Se automaticamente alterado para última linha
                    throw new Exception("Para tratar arquivos CSV, TXT ou RPT é necessário informar qual será a última linha!");

                // Realiza a leitura da primeira linha (cabeçalho)
                row = table.Rows[_i - 1].ItemArray.Select(f => f.ToString()).ToList(); // linha 2 primeira => index é 1 (-1) e cabeçalho ja retirado (-1)
                _i += 1; // Próxima leitura será a 2ª linha
                _j += 1;
            }

            return row;
        }

        /// <summary>
        /// Abre o arquivo e realiza a leitura
        /// </summary>       
        private static DataSet GetDataSet(string origin, string destiny)
        {

        restart:

            // Abre o arquivo
            using (var stream = File.Open(origin, FileMode.Open, FileAccess.Read))
            {
                DataSet result;

                // Realiza a leitura do arquivo
                switch (Path.GetExtension(origin).ToLower())
                {
                    case ".gz":
                        origin = UnGZ(stream, Path.GetDirectoryName(destiny) + "\\");
                        goto restart;

                    case ".zip":
                        stream.Close();
                        origin = UnZIP(origin, Path.GetDirectoryName(destiny));
                        goto restart;

                    case ".rpt":
                    case ".txt":
                    case ".csv":
                        result = ReadCSV(stream);
                        break;

                    default: // .xlsx, .xls, .xlsb, .xlsm
                        result = ReadXLS(stream);
                        break;
                }

                return result;
            }




        }



        /// <summary>
        /// Realiza a conversão do arquivo Excel localizado em <paramref name="origin"/>, salva em <paramref name="destiny"/>
        /// com tratativa de exceçoes para o usuário final (arquivo inexistente no diretorio ou aberto durante a conversão)
        /// e retorna 'true' caso a conversão tenha ocorrido com sucesso
        /// </summary>
        /// <param name="origin">Diretorio + nome do arquivo de origem + formato. Ex.: "C:\\Users\\ArquivoExcel.xlsx"</param>
        /// <param name="destiny">Diretorio + nome do arquivo de destino + formato. Ex.: "C:\\Users\\ArquivoExcel.csv"</param>
        /// <param name="sheet">Aba da planilha a ser convertida. Ex.: "1" (primeira aba) ou "NomeAba"</param>
        /// <param name="separator">Separador a ser utilizado para realizar a conversão. Ex.: ";"</param>
        /// <param name="columns">"Vetor de caracteres (maiúsculo ou minúsculo) contendo todas as colunas desejadas. Ex.: { "A", "b", "E", "C" } ou "{ "A:BC" } </param>
        /// <param name="rows">"Informe a primeira e última linha (ou deixe em branco). Ex.: "1:50 (linha 1 até linha 50)"</param>
        /// <param name="pgbar">"Caso desejado, passe uma ProgressBar para ser carregada em 100 (ou null). Ex.: "ProgressBar pgbar = new ProgressBar()"</param>
        /// <returns>"true" se convertido com sucesso. "false" se não convertido.</returns>
        public static bool ConverterExcept(string origin, string destiny, string sheet, string separator, string[] columns, string rows, ProgressBar pgbar)
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



        /// <summary>
        /// Realiza a conversão do arquivo Excel localizado em <paramref name="origin"/>, salva em <paramref name="destiny"/>
        /// e retorna 'true' caso a conversão tenha ocorrido com sucesso.
        /// Utilize o método "ConverterExcept" para realizar a conversão e tratar algumas exceções!
        /// </summary>
        /// <param name="origin">Diretorio + nome do arquivo de origem + formato. Ex.: "C:\\Users\\ArquivoExcel.xlsx"</param>
        /// <param name="destiny">Diretorio + nome do arquivo de destino + formato. Ex.: "C:\\Users\\ArquivoExcel.csv"</param>
        /// <param name="sheet">Aba da planilha a ser convertida. Ex.: "1" (primeira aba) ou "NomeAba"</param>
        /// <param name="separator">Separador a ser utilizado para realizar a conversão. Ex.: ";"</param>
        /// <param name="columns">"Vetor de caracteres (maiúsculo ou minúsculo) contendo todas as colunas desejadas. Ex.: { "A", "b", "E", "C" } ou "{ "A:BC" } </param>
        /// <param name="rows">"Informe a primeira e última linha (ou deixe em branco). Ex.: "1:50 (linha 1 até linha 50)"</param>
        /// <param name="pgbar">"Caso desejado, passe uma ProgressBar para ser carregada em 100 (ou null). Ex.: "ProgressBar pgbar = new ProgressBar()"</param>
        /// <returns>"true" se convertido com sucesso. "false" se não convertido.</returns>
        public static bool Converter(string origin, string destiny, string sheet, string separator, string[] columns, string rows, ProgressBar pgbar)
        {

            ValidateString(new string[] { origin, destiny, sheet, separator, rows, columns[0] });


            if (pgbar == null)
                pgbar = new ProgressBar();


            File.WriteAllText(destiny, ""); // Para verificar se arquivo de destino esta acessivel
            File.Delete(destiny); // Deleta para evitar que usuario abra o arquivo durante a conversao
            pgbar.Value += 5; // 5 

            DataSet result = GetDataSet(origin, destiny);

            pgbar.Value += 30; // 35 (pós leitura do arquivo)

            // Obtem a aba a ser convertida
            DataTable table = GetTable(sheet, result);

            StringBuilder output = new StringBuilder();

            pgbar.Value += 5; // 40

            // Define qual será a primeira e última linha a ser convertida
            int[] rowsNumber = ExcelHelper.DefineRows(rows, table.Rows.Count + 1);
            pgbar.Value += 5; // 45                


            int[] columnsASCII = null;
            _i = rowsNumber[0]; // Primeira linha a ser convertida
            _j = 0; // Deslocamento

            List<string> row = GetFirstRow(Path.GetExtension(origin), table);


            // Se deseja selecionar colunas específicas
            if (columns != null && columns.Length != 0) // null OR {}
            {
                if (columns[0].Contains(":"))
                { // Se primeira celula do array. Ex.: {"A:G"}
                    columnsASCII = DefineColunms(columns[0], GetNameColumn(row.Count()));
                }
                else
                { // Se colunas definidas individualmente. Ex.: {"A", "B"}
                    columnsASCII = DefineColunms(columns);
                }

            }

            pgbar.Value += 5; // 50 (tratativas)

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
                    pgbar.Value += (int)countPercPrg; // 90                                                               
                    countPercPrg -= (int)countPercPrg;
                }

                countPercPrg += percPrg; // Incrementa contador da ProgressBar                        

                // Obtem a próxima linha
                row = table.Rows[_i - 1].ItemArray.Select(f => f.ToString()).ToList();
            }



            pgbar.Value += (90 - pgbar.Value); // Se necessário, completa até 90%

            // Escreve o novo arquivo convertido (substitui se ja existente)
            File.WriteAllText(destiny, output.ToString());
            pgbar.Value += 10; // 100
            return true;



        }



    }
}

