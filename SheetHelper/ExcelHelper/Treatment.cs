using System;

namespace SheetHelper
{
    /// <summary>
    /// Used to process the information provided by the user
    /// </summary>
    internal static class Treatment
    {
        /// <summary>
        /// Define o índice de todas as colunas que serão convertidas
        /// </summary>
        /// <param name="columns">Colunas a serem convertidas. E.g.: "B:H"</param>
        /// <param name="lastColumnName">Última coluna da planilha (limite máximo de colunas). E.g.: "AZ"</param>
        /// <returns>Vetor com todos os índices das colunas a serem convertidas</returns>
        /// <exception cref="Exception"></exception>
        internal static int[] DefineColunms(string columns, string lastColumnName)
        {
            int indexLastColumn = SH.GetIndexColumn(lastColumnName);

            if (columns == null || columns.Equals("")) // Se coluna não especificadas           
                return new[] { 0, indexLastColumn }; // Converte todas as colunas

            columns = columns.Trim();

            if (!columns.Contains(":"))
                throw new Exception("Use um ':' para definir a primeira e última coluna!");

            string[] rowsArray = columns.Split(':'); // E.g.: {"A","Z"}

            if (rowsArray[0].Equals("")) // Se primeira coluna não definida
                rowsArray[0] = "A"; // Então, deseja-se converter desde a primeira coluna

            if (rowsArray[1].Equals("")) // Se última coluna não definida
                rowsArray[1] = lastColumnName; // Então, deseja-se converter até a última coluna


            if (SH.GetIndexColumn(rowsArray[0]) <= 0 || SH.GetIndexColumn(rowsArray[0]) > indexLastColumn)
                throw new Exception($"A primeira coluna está fora do limite (min A, max {lastColumnName})!");

            if (SH.GetIndexColumn(rowsArray[1]) <= 0 || SH.GetIndexColumn(rowsArray[1]) > indexLastColumn)
                throw new Exception($"Última coluna fora dos limites (min A, max {lastColumnName})!");

            if (SH.GetIndexColumn(rowsArray[0]) > SH.GetIndexColumn(rowsArray[1])) // Para este caso (incrementador)
                throw new Exception($"A coluna inicial deve vir antes da última coluna!");

            int firstColumn = SH.GetIndexColumn(rowsArray[0]);
            int lastColumn = SH.GetIndexColumn(rowsArray[1]);
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
        internal static int[] DefineRows(string rows, int limitRows)
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
        /// Recebe as colunas em string, converte para caracter e retorna em ASCII correspondente
        /// </summary>
        internal static int[] DefineColunms(string[] columns)
        {
            int[] columnsASCII = new int[columns.Length];

            for (int i = 0; i < columns.Length; i++) // Para cada coluna
            {
                columnsASCII[i] = SH.GetIndexColumn(columns[i].ToUpper());
            }

            return columnsASCII;
        }

        internal static bool ValidateString(string[] strings)
        {
            foreach (string str in strings)
            {
                if (string.IsNullOrEmpty(str))
                    throw new Exception($"'{str}' inválido!");
            }
            return true;
        }


    }
}
