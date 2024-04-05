[![NuGet](https://img.shields.io/nuget/v/SheetHelper.svg)](https://www.nuget.org/packages/SheetHelper/)

<img src="https://raw.githubusercontent.com/DiegoPiovezana/DiegoPiovezana/main/Images/us.png" width=2.0% height=2.0%> See the documentation in English by [clicking here](../../Readme.md).<br/>

# SheetHelper
Biblioteca rápida e leve para fácil leitura e conversão de grandes arquivos Excel.<br/>

<img src="../Images/SheetHelper_publish.png" width=100% height=100%> 

RECURSOS DISPONÍVEIS:<br/>
✔ Compatível com leitura de arquivos .xlsx, .xlsm, .xls, .xlsb, .csv, .txt, .rpt e outros;<br/>
✔ Obtenha um DataTable de uma planilha usando o método `GetDataTable`;<br/>
✔ Use `SaveDataTable` para salvar um DataTable em diferentes formatos e com restrição de colunas e linhas;<br/>
✔ Use o método `CloseExcel` para fechar todos os processos do Excel, inclusive os em segundo plano;<br/>
✔ Use `GetIndexColumn` para obter o índice da coluna fornecendo o nome (por exemplo, "AB");<br/>
✔ O método `GetNameColumn` pode ser usado para obter o nome da coluna;<br/>
✔ Use `GetRowArray` para obter uma linha de um DataTable;<br/>
✔ Converta um array em uma DataRow usando o método `ConvertToDataRow`;<br/>
✔ Converta uma planilha para diferentes formatos usando o método `Converter`;<br/>
✔ Permite converter intervalos de linhas. Ex: "1:23, -34:56, 70:40, 75, -1";<br/>
✔ Possibilidade de converter intervalos de colunas. Ex: "A:H, 4:9, 4:-9, B, 75, -2";<br/>
✔ Substitui o arquivo se já convertido;<br/>
✔ Opção de escolher a planilha desejada para conversão usando índice ou nome (sem diferenciação entre maiúsculas e minúsculas);<br/>
✔ Pode escolher o formato de arquivo a ser convertido;<br/>
✔ Opção de escolher o nome do arquivo, local de destino e formato a ser salvo;<br/>
✔ Permite alterar o delimitador;<br/>
✔ Suporta a conversão de colunas, linhas e planilhas ocultas;<br/>
✔ Possibilidade de escolher colunas e linhas específicas para conversão;<br/>
✔ Permite acompanhar a porcentagem de conclusão por meio da propriedade `Progress`;<br/>
✔ Lida com o usuário final quando o arquivo não é encontrado com MessageBox para NetFramework;<br/>
✔ Suporta descompactação de arquivos .GZ (usando `UnGZ`) e .ZIP (usando `UnZIP`). Ou use `UnzipAuto` para descompactar automaticamente.<br/>

<br/>Utiliza a biblioteca [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) para realizar a leitura.<br/>

<br/>

### CONTATO:
https://bit.ly/FeedbackHappyHelper


<br/><br/>

## PRINCIPAIS CONVERSÕES POSSÍVEIS:<br/>
<img src="../Images/Conversions_pt-br.png" width=80% height=80%> 


### INSTALAÇÃO:
```
 dotnet add package SheetHelper
```

<br/>

## EXEMPLO DE USO:
```c#
using SH;

namespace App
{
    static class Program
    {
        static void Main()
        {
            string origem  = "C:\\Users\\Diego\\Files\\Report.xlsx.gz";
            string destino = "C:\\Users\\Diego\\Files\\Report.csv";

            string aba = "1"; // Use "1" para a primeira aba (possível informar índice ou nome)
            string delimitador  = ";";
            string colunas  = "A, 3, b, 12:-1"; // ou null para converter todas as colunas ou "A:BC" para um intervalo de colunas
            string linhas = ":4, -2"; // Ex: Extrai da 1ª à 4ª linha e também a penúltima linha            

            bool resultado = SheetHelper.Converter(origem, destino, aba, delimitador, colunas, linhas);
            
            Console.WriteLine(resultado ? "O arquivo foi convertido com sucesso!" : "Falha ao converter o arquivo!")
        }
    }
}

```