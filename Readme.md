[![NuGet](https://img.shields.io/nuget/v/SheetHelper.svg)](https://www.nuget.org/packages/SheetHelper/)


Biblioteca rápida e leve, para fácil conversão de grandes arquivos Excel <br/>

<img src="https://github.com/SANTODEVS/SheetHelper/blob/master/SheetHelper/Images/SheetHelper.png?raw=true" width=50% height=50%>

## RECURSOS DISPONÍVEIS: <br/>
✔ Compatível com leitura de arquivos .xlsx, .xls, .xlsb e .csv; <br/>
✔ Contém tratativas de exceções para o usuário final (utilize o método ConverterExcept)¹; <br/>
✔ Possibilidade de exclusão de cabeçalho; <br/>
✔ Substitui arquivo se já convertido; <br/>
✔ Possibilidade de escolher a aba desejada para conversão; <br/>
✔ Se arquivo está sendo usado por outro processo, oferece opção para encerrar o Excel; <br/>
✔ Opção para escolha do formato do arquivo a ser convertido; <br/>
✔ Opção para escolha do nome do arquivo, local de destino e formato a ser salvo; <br/>
✔ Opção para alterar delimitador; <br/>
✔ Possibilidade de escolha de colunas a serem convertidas; <br/>
✔ Capacidade de definir um tamanho personalizado para o cabeçalho (escolha da linha de início).<br/>

Faz uso da biblioteca [ExcelDataReader v.](https://github.com/ExcelDataReader/ExcelDataReader) <br/>

¹ Arquivo inexistente no diretorio de origem ou aberto durante a conversão. 
<br/>

## CONVERSÕES POSSÍVEIS: <br/>
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
            string origem = "C:\\Users\\Usuario\\Arquivos\\Excel_original.xlsx";
            string destino = "C:\\Users\\Usuario\\Arquivos\\Excel_convertido.csv";

            int aba = 1; // Utilize 0 para a primeira aba
            string separador = ";";
            int cabecalho = 2; // Remove a primeira e segunda linha da planilha
            string[] colunas = { }; // ou null, para converter todas as colunas
             
            bool retorno = ExcelHelper.ConverterExcept(origem, destino, aba, separador, cabecalho, colunas);

            if (retorno) MessageBox.Show("O arquivo foi convertido com sucesso!");
            else MessageBox.Show("Não foi possível converter o arquivo!");

        }
    }
}

```