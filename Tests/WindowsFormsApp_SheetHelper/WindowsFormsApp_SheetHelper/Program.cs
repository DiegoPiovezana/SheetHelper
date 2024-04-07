using SH;
using System;
using System.Reflection.Emit;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp_SheetHelper
{
    internal static class Program
    {
        public static bool Converter(ProgressBar carregamento)
        {

            // TODO: Tratativa converter para XLSX
            // TODO: Converter de .TXT para .CSV


            string origem;
            string destino;
            //bool retorno;

            string aba = "1"; // Utilize "1" para a primeira aba
            string separador = ";";
            string colunas = "A:F"; // ou null, para converter todas as colunas
            string linhas; // Ex.: extrai a partir da 2ª linha da planilha até a última (retira a 1ª linha)          

            //var abas = SheetHelper.GetSheets(
            //    "C:\\Users\\diego\\Desktop\\Lixo\\Converter\\ColunasExcel.xlsx",
            //    1,
            //    true
            //    );

            //retorno = abas.Count > 0;

            //var retorno = SheetHelper.ConvertAllSheets(
            //    "C:\\Users\\diego\\Desktop\\Lixo\\Converter\\ColunasExcel.xlsx",
            //    "C:\\Users\\diego\\Desktop\\Lixo\\Convertidos\\Conversao_XLSX_abas.csv"                
            //    );

            //MessageBox.Show(retorno ? $"O arquivo foi convertido com sucesso!" : "Não foi possível converter o arquivo!");

            //var text1 = SheetHelper.NormalizeText("Teste de texto com acentuação e ç");
            //var text2 = SheetHelper.NormalizeText("Teste de teXto com  acentuação e ç");
            //var text3 = SheetHelper.NormalizeText(" Teste de teXto com acentuação e ç");
            //var text4 = SheetHelper.NormalizeText(" Teste de teXto com acentuação e ç ");

            //retorno = true;
            //carregamento.Value = SheetHelper.Progress;
            //aba = "3";
            //linhas = "";
            //colunas = "";
            //origem = "C:\\Users\\diego\\Desktop\\Lixo\\Converter\\ColunasExcel.xlsx";
            //destino = "C:\\Users\\diego\\Desktop\\Lixo\\Convertidos\\Conversao_XLSX_abas.csv";
            //retorno = SheetHelper.Converter(origem, destino, aba, separador, colunas, linhas);

            //carregamento.Value = ref SheetHelper.Progress;
            //SheetHelper.OnProgressChanged += newValue => { carregamento.Value = newValue; };            
            //SheetHelper.Progress.ValueChanged += (sender, e) => { value = e.NewValue; };            
            //aba = "Sheet7";
            //origem = "C:\\Users\\diego\\Desktop\\Lixo\\Converter\\ColunasExcel.xlsx";
            //destino = "C:\\Users\\diego\\Desktop\\Lixo\\Convertidos\\Conversao_QuebraLinha.csv";
            //retorno = SheetHelper.Converter(origem, destino, aba, separador, null, null);


            //linhas = "1:"; // 502.383
            //origem = "C:\\Users\\diego\\Desktop\\Lixo\\Converter\\ColunasExcelBig.xlsb";
            //destino = "C:\\Users\\diego\\Desktop\\Lixo\\Convertidos\\ConversaoBig_XLSB.csv";
            //ref int progress = ref SheetHelper.Progress;
            //retorno = SheetHelper.ConverterAllSheet(origem, destino);
            //var retorno = SheetHelper.Converter(origem, destino, aba, separador, colunas, linhas);



            //var retorno = false;




            //Task<bool> taskConvert = Task.Run(() => SheetHelper.Converter(origem, destino, aba, separador, colunas, linhas));
            //taskConvert.Wait();
            //retorno = taskConvert.Result;

            //linhas = "1:10";
            ////carregamento.Value = 0;
            //origem = "C:\\Users\\diego\\Desktop\\Lixo\\Converter\\ColunasExcel.csv";
            //destino = "C:\\Users\\diego\\Desktop\\Lixo\\Convertidos\\ColunasExcel.txt";
            //retorno = SheetHelper.Converter(origem, destino, aba, separador, colunas, linhas);

            //origem = "C:\\Users\\diego\\Desktop\\Lixo\\Converter\\Teste.rpt";
            //destino = "C:\\Users\\diego\\Desktop\\Lixo\\Convertidos\\TesteRPT.csv";
            //separador = ";";
            //linhas = "1:40";
            //colunas = "A:F";
            //carregamento.Value = 0;
            //var retorno = SheetHelper.Converter(origem, destino, aba, separador, colunas, linhas);

            origem = "C:\\Users\\diego\\Desktop\\Lixo\\Converter\\ColunasExcel.xlsx";
            destino = "C:\\Users\\diego\\Desktop\\Lixo\\Convertidos\\TesteXLSX.csv";
            separador = ";";
            linhas = "3:5,2,3";
            colunas = "A,C,B";            
            var retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);

            //linhas = "1:10";
            //carregamento.Value = 0;
            //origem = "C:\\Users\\diego\\Desktop\\Lixo\\Converter\\Teste.txt";
            //destino = "C:\\Users\\diego\\Desktop\\Lixo\\Convertidos\\Teste4.csv";
            //retorno = SheetHelper.Converter(origem, destino, aba, separador, colunas, linhas);

            //linhas = "1:10";
            //carregamento.Value = 0;
            //origem = "C:\\Users\\diego\\Desktop\\Lixo\\Converter\\Teste.txt";
            //destino = "C:\\Users\\diego\\Desktop\\Lixo\\Convertidos\\Teste5.txt";
            //retorno = SheetHelper.Converter(origem, destino, aba, separador, colunas, linhas);

            // Inicie um loop para atualizar a barra de progresso
            //while (SheetHelper.Progress < 100)
            //{
            //    Console.WriteLine(SheetHelper.Progress);
            //    carregamento.Value = SheetHelper.Progress;

            //    // Aguarde um intervalo de tempo antes de verificar novamente
            //    // Isso evita que o loop fique consumindo muitos recursos
            //    Thread.Sleep(100); // Aguarda 100 milissegundos (0,1 segundo) antes de verificar novamente
            //}

            //if (retorno) MessageBox.Show("O arquivo foi convertido com sucesso!");
            //else MessageBox.Show("Não foi possível converter o arquivo!");

            return retorno;
        }




        static void Main()
        {

            SheetHelper_Menu form1 = new SheetHelper_Menu();
            form1.ShowDialog();



        }
    }
}
