[![NuGet](https://img.shields.io/nuget/v/SheetHelper.svg)](https://www.nuget.org/packages/SheetHelper/)

Fast and lightweight library for easy conversion of large Excel files. <br/>

<img src="SheetHelper\Images\SheetHelper_publish.png" width=100% height=100%> 
AVAILABLE FEATURES: <br/>
✔ Compatible with reading .xlsx, .xlsm, .xls, .xlsb, .csv, .txt, .rpt files, among others; <br/>
✔ Contains exception handling for the end user (use the ConverterExcept method)¹; <br/>
✔ Allows defining the first and last rows to be converted; <br/>
✔ Replaces file if already converted; <br/>
✔ Option to choose the desired sheet for conversion using index or name; <br/>
✔ If the file is being used by another process, offers an option to close Excel; <br/>
✔ Can choose the file format to be converted; <br/>
✔ Option to choose the file name, destination location, and format to be saved; <br/>
✔ Allowed to change the delimiter; <br/>
✔ Supports conversion of hidden columns, rows, and sheets; <br/>
✔ Possibility to choose specific columns to be converted; <br/>
✔ Ability to load a Progress Bar;<br/>
✔ Supports decompression of .GZ and .ZIP files.<br/>

Uses the library ExcelDataReader version 3.7.0 <br/>

¹ File does not exist in the source directory or is open during conversion.
<br/><br/>

## PRINCIPAIS CONVERSÕES POSSÍVEIS: <br/>
<img src="SheetHelper\Images\Conversions.png" width=80% height=80%> 

## EXEMPLO DE USO:
```c#
using SheetHelper;
using System.Windows.Forms;

namespace WindowsFormsAppNetFrameworkMain
{
    static class Program
    {
        static void Main()
        {
            string origem = "C:\\Users\\diego\\Arquivos\\Relatorio.xlsx.gz";
            string destino = "C:\\Users\\diego\\Arquivos\\Relatorio.csv";

            string aba = "1"; // Utilize "1" para a primeira aba (índice ou nome)
            string separador = ";";
            string[] colunas = {"A", "C", "b"}; // ou null, para converter todas as colunas ou {"A:BC"} para intervalo de colunas
            string linhas = "2:"; // Ex.: extrai a partir da 2ª linha da planilha até a última (retira a 1ª linha)          

            bool retorno = ExcelHelper.ConverterExcept(origem, destino, aba, separador, colunas, linhas, null);

            if (retorno) MessageBox.Show("O arquivo foi convertido com sucesso!");
            else MessageBox.Show("Não foi possível converter o arquivo!");
        }
    }
}

```