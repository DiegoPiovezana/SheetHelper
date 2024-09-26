<img src="https://raw.githubusercontent.com/DiegoPiovezana/DiegoPiovezana/main/Images/us.png" width=2.0% height=2.0%> See in English by [clicking here](../../Changelog.md).<br/>

## [Versão 1.4.0] - 2023-12-01

- Capacidade de manipular e salvar planilhas contendo caracteres especiais (UTF-8)
- O diretório temporário "SheetHelper" é automaticamente excluído após a conclusão
- Capacidade de selecionar linhas em ordem reversa (ex.: "3:1")
- Colunas podem ser selecionadas em ordem reversa (ex.: "-1:-5")
- Dicionário de termos proibidos em células que serão substituídos
- Pontos e vírgulas (";") nas células são mantidos dentro da célula, em vez de criar uma nova coluna

## [Versão 1.3.0] - 2023-12-18

- Possibilidade de obter o `DataSet` de um arquivo
- Células que possuem quebras de linha agora serão convertidas corretamente
- O método `NormalizeText` permite remover todos os acentos e espaços de um texto
- Possível converter todas as abas da planilha, considerando todas as linhas e colunas
- Use `GetRowArray` para obter uma linha de um `DataTable`
- O método `GetSheets` permite obter todas as abas da planilha em `DataTable`
- Correção na seleção da primeira linha que não é o cabeçalho
- Agora é possível realizar a conversão de abas que possuem apenas uma linha

## [Versão 1.2.0] - 2023-07-10

- Permitido especificar o nome da aba, desconsiderando a diferenciação entre maiúsculas e minúsculas
- Permitido salvar `DataTable` em diferentes formatos e com restrição de colunas e linhas
- Adicionada a possibilidade de obter a primeira linha de um `DataTable`
- Adicionado tratamento para o usuário final quando o arquivo não é encontrado ou está em uso, com `MessageBox` para NetFramework
- Correção na conversão considerando o formato do cabeçalho

## [Versão 1.1.1] - 2023-06-01

- Tratamento para conversão desnecessária entre `.CSV` e `.TXT`

## [Versão 1.1.0] - 2023-05-24

- Adicionada a possibilidade de converter intervalos de linhas. Ex.: "1:23, -34:56, 70, 75, -1"
- Possível converter intervalos de colunas. Ex.: "A:H, 4:9, 4:-9, B, 75, -2"
- Biblioteca ExcelDataReader integrada
- Compatível com net462, netstandard 2.0 e netstandard 2.1
- Vários bugs corrigidos
- Melhoria de performance

## [Versão 1.0.0.6] - 2022-06-13

- Adicionada opção para realizar a conversão contínua de colunas (A:AB) e escolher a aba pelo nome ou índice

## [Versão 1.0.0.5] - 2022-03-29

- Adicionada opção para converter arquivos RPT

## [Versão 1.0.0.4] - 2022-03-28

- Adicionada opção para escolher linhas, barra de progresso e descompactar arquivos

## [Versão 1.0.0.3] - 2022-02-01

- Adicionada a possibilidade de conversão de arquivos XSLB, XSL, HTML e CSV e escolha do tamanho do cabeçalho

## [Versão 1.0.0.2] - 2022-01-19

- Permitida a visualização do resumo dos métodos

## [Versão 1.0.0.1] - 2022-01-18

- Atualização de ícone, readme e nomenclatura
- Correções de bugs

## [Versão 1.0.0] - 2022-01-17

- Lançamento
