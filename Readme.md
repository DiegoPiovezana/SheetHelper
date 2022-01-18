[![NuGet](https://img.shields.io/nuget/v/SheetHelper.svg)](https://www.nuget.org/packages/SheetHelper/)


Biblioteca rápida e leve, para fácil conversão de grandes arquivos Excel <br/>

![Screenshot](images/icon.png)

RECURSOS DISPONÍVEIS: <br/>
✔ Compatível com arquivos .xlsx, .xls, .xlsb, .csv e .txt; <br/>
✔ Contém algumas tratativas de exceções para o usuário final (ConverterExcept)¹; <br/>
✔ Possibilidade de exclusão de cabeçalho; <br/>
✔ Substitui arquivo se já convertido; <br/>
✔ Possibilidade de escolher a aba desejada para conversão; <br/>
✔ Se arquivo está sendo usado por outro processo, oferece opção para encerrar o Excel; <br/>
✔ Opção para escolha do formato do arquivo a ser convertido; <br/>
✔ Opção para escolha do nome do arquivo e local de destino; <br/>
✔ Opção para alterar delimitador; <br/>
✔ Possibilidade de escolha de colunas a serem convertidas. <br/>

Faz uso da biblioteca [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) <br/>

¹ Arquivo inexistente no diretorio de origem ou aberto durante a conversão. 