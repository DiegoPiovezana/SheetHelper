[![NuGet](https://img.shields.io/nuget/v/SheetHelper.svg)](https://www.nuget.org/packages/SheetHelper/)


Biblioteca rápida e leve, para fácil conversão de grandes arquivos Excel <br/>

<img src="https://github.com/SANTODEVS/SheetHelper/blob/master/SheetHelper/Images/SheetHelper.png?raw=true" width=50% height=50%>

## RECURSOS DISPONÍVEIS: <br/>
✔ Compatível com leitura de arquivos .xlsx, .xlsm, .xls, .xlsb, .csv, entre outros; <br/>
✔ Contém tratativas de exceções para o usuário final (utilize o método ConverterExcept)¹; <br/>
✔ Possibilita definir a primeira e última linha que serão convertidas; <br/>
✔ Substitui arquivo se já convertido; <br/>
✔ Possibilidade de escolher a aba desejada para conversão; <br/>
✔ Se arquivo está sendo usado por outro processo, oferece opção para encerrar o Excel; <br/>
✔ Opção para escolha do formato do arquivo a ser convertido; <br/>
✔ Opção para escolha do nome do arquivo, local de destino e formato a ser salvo; <br/>
✔ Opção para alterar delimitador; <br/>
✔ Possibilidade de escolha de colunas a serem convertidas; <br/>
✔ Capacidade de realizar carregamento de Progress Bar;<br/>
✔ Suporte a descompactação de arquivos .GZ e .ZIP.<br/>

Faz uso da biblioteca [ExcelDataReader versão 3.7.0](https://github.com/ExcelDataReader/ExcelDataReader) <br/>

¹ Arquivo inexistente no diretorio de origem ou aberto durante a conversão. 
<br/><br/>

## PRINCIPAIS CONVERSÕES POSSÍVEIS: <br/>
<img src="https://github.com/SANTODEVS/SheetHelper/blob/master/SheetHelper/Images/Conversions.png?raw=true" width=65% height=65%>
<br/>

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

            int aba = 0; // Utilize 0 para a primeira aba
            string separador = ";";
            string[] colunas = {"A", "C", "b"}; // ou null, para converter todas as colunas
            string linhas = "2:"; // Ex.: extrai a partir da 2ª linha da planilha até a última (retira a 1ª linha)          

            bool retorno = ExcelHelper.ConverterExcept(origem, destino, aba, separador, colunas, linhas, null);

            if (retorno) MessageBox.Show("O arquivo foi convertido com sucesso!");
            else MessageBox.Show("Não foi possível converter o arquivo!");

        }
    }
}

```