[![NuGet](https://img.shields.io/nuget/v/SheetHelper.svg)](https://www.nuget.org/packages/SheetHelper/)


Biblioteca rápida e leve, para fácil conversão de grandes arquivos Excel <br/>

<img src="SheetHelper\Images\SheetHelper_publish.png" width=100% height=100%> 

## RECURSOS DISPONÍVEIS: <br/>
✔ Compatível com leitura de arquivos .xlsx, .xlsm, .xlsb, .xls, .csv, .txt, .rpt, entre outros; <br/>
✔ Possibilita definir a primeira e última linha que serão convertidas; <br/>
✔ Substitui arquivo se já convertido; <br/>
✔ Possibilidade de escolher a aba desejada para conversão utilizando índice ou nome; <br/>
✔ Se arquivo está sendo usado por outro processo, oferece opção para encerrar o Excel; <br/>
✔ É possível escolher o formato do arquivo a ser convertido; <br/>
✔ Opção para escolha do nome do arquivo, local de destino e formato a ser salvo; <br/>
✔ Permitido alterar o delimitador; <br/>
✔ Suporta conversão de colunas, linhas e abas ocultas; <br/>
✔ Possibilidade de escolha de colunas a serem convertidas; <br/>
✔ Suporte a descompactação de arquivos .GZ e .ZIP.<br/>

Faz uso da biblioteca [ExcelDataReader versão 3.7.0](https://github.com/ExcelDataReader/ExcelDataReader) <br/>


<br/><br/>

## PRINCIPAIS CONVERSÕES POSSÍVEIS: <br/>
<img src="SheetHelper\Images\Conversions_pt-br.png" width=80% height=80%> 

## EXEMPLO DE USO:
```c#
using SheetHelper;

namespace App
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