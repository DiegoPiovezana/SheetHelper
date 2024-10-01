using Microsoft.VisualStudio.TestTools.UnitTesting;
using SH;

namespace UnitTestSheetHelper
{
    [TestClass]
    public class UnitTest_viaNuget
    {
        [TestMethod]
        public void TestConvertParticular()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_0.csv";

            string aba = "1";
            string separador = ";";           
            string colunas = null;
            string linhas = "1:";          

            bool retorno = SheetHelper.Converter(origem, destino, aba, separador, colunas, linhas);
            //Assert.That(retorno, Is.EqualTo(true));
            Assert.AreEqual(retorno, true);
        }

        //[DataTestMethod]
        //[TestCase(new string[] { "A", "C", "b" }, 1, ExpectedResult = true, TestName = "Colunas maiuscula e minuscula e fora de ordem...")]
        //[TestCase(new string[] { "A:D" }, 2, ExpectedResult = true, TestName = "Colunas em intervalo contínuo...")]
        //[TestCase(new string[] { "" }, 3, ExpectedResult = true, TestName = "Array com string vazia...")]
        //[TestCase(new string[] { }, 4, ExpectedResult = true, TestName = "Arary vazio...")]
        //[TestCase(null, 5, ExpectedResult = true, TestName = "Array nulo...")]
        //[TestCase(new string[] { null }, 6, ExpectedResult = true, TestName = "Array com string nula...")]
        //public bool TestColuns(string[] colunas, int id)
        //{
        //    string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xls";
        //    string destino = @$"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_column{id}.csv";

        //    string aba = "1";
        //    string separador = ";";
        //    string linhas = "2:";          

        //    return ExcelHelper.ConverterExcept(origem, destino, aba, separador, colunas, linhas, null);

        //}

        //// --------------------------------------------------------------------------------

        //[TestCase("2:", 1, ExpectedResult = true, TestName = "2:")]
        //[TestCase(":10", 2, ExpectedResult = true, TestName = ":10")]
        //[TestCase("1:20", 3, ExpectedResult = true, TestName = "1:20")]
        //[TestCase("", 4, ExpectedResult = true, TestName = "String vazia")]
        //[TestCase(null, 5, ExpectedResult = true, TestName = "Nulo")]
        //[TestCase("1, 2, 4", 6, ExpectedResult = false, TestName = "1, 2, 4")]
        //[TestCase("7", 7, ExpectedResult = true, TestName = "7")]
        //public bool TestRows(string linhas, int id)
        //{
        //    string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xlsx";
        //    string destino = @$"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_row{id}.csv";

        //    string aba = "1"; 
        //    string separador = ";";
        //    string[] colunas = { "A", "C", "b" }; 
        //    //string linhas = "2:"; 

        //    return ExcelHelper.ConverterExcept(origem, destino, aba, separador, colunas, linhas, null);

        //}

        //// --------------------------------------------------------------------------------


        //[TestFixture]
        //public class TestsFormats
        //{
        //    [Test, TestCaseSource(typeof(CasosDeTesteDeFormatos), nameof(CasosDeTesteDeFormatos.FormatsConverts))]
        //    public bool ValidarFormatosValidos(string origin, string destination) => TestFormats(origin, destination);
        //}

        //public class CasosDeTesteDeFormatos
        //{
        //    public static List<TestCaseData> FormatsConverts
        //    {
        //        get
        //        {
        //            string path = "C:\\Users\\diego\\Desktop\\Tests";

        //            return new List<TestCaseData>()
        //             {
        //                 new TestCaseData("", "").Returns(new Exception().Message == "'' é inválido!").SetName("Origem e destino vazio"),
        //                 new TestCaseData($"{path}\\ColunasExcel.xlsm",$"{path}\\Convertidos\\ColunasExcel_XLSM.csv").Returns(true).SetName("XLSM__CSV"),
        //                 new TestCaseData($"{path}\\ColunasExcel.xlsx.gz",$"{path}\\Convertidos\\ColunasExcel_XLSX_GZ.csv").Returns(true).SetName("XLSX_GZ__CSV"),
        //                 new TestCaseData($"{path}\\ColunasExcel.gz", $"{path}\\Convertidos\\ColunasExcel_GZ.csv").Returns(true).SetName("GZ__CSV"),
        //                 new TestCaseData($"{path}\\ColunasExcel.tar.gz", $"{path}\\Convertidos\\ColunasExcel_TAR_GZ.csv").Returns(true).SetName("TAR_GZ__CSV"),
        //                 new TestCaseData($"{path}\\ColunasExcel.xlsx.zip", $"{path}\\Convertidos\\ColunasExcel_XLSX_ZIP.txt").Returns(true).SetName("XLSX_ZIP__TXT"),
        //                 new TestCaseData($"{path}\\ColunasExcel.zip", $"{path}\\Convertidos\\ColunasExcel_ZIP.csv").Returns(true).SetName("ZIP__CSV"),
        //                 new TestCaseData($"{path}\\ColunasExcel.xls", $"{path}\\Convertidos\\ColunasExcel_XLS.csv").Returns(true).SetName("XLS__CSV"),
        //                 new TestCaseData($"{path}\\ColunasExcel.csv", $"{path}\\Convertidos\\ColunasExcel_CSV.csv").Returns(true).SetName("CSV__CSV"),
        //                 new TestCaseData($"{path}\\ColunasExcel.csv", $"{path}\\Convertidos\\ColunasExcel_CSV.txt").Returns(true).SetName("CSV__TXT"),
        //                 new TestCaseData($"{path}\\ColunasExcel.xlsb", $"{path}\\Convertidos\\ColunasExcel_XLSB.csv").Returns(true).SetName("XLSB__CSV"),
        //                 new TestCaseData($"{path}\\ColunasExcel.xlsx", $"{path}\\Convertidos\\ColunasExcel_XSLX.csv").Returns(true).SetName("XSLX__CSV"),
        //                 new TestCaseData($"{path}\\ColunasExcel.txt", $"{path}\\Convertidos\\ColunasExcel_TXT.txt").Returns(true).SetName("TXT__TXT")
        //             };
        //        }
        //    }
        //}

        //private static bool TestFormats(string origem, string destino)
        //{
        //    //string origem = "C:\\Users\\diego\\Arquivos\\Relatorio.xlsx.gz";
        //    //string destino = "C:\\Users\\diego\\Arquivos\\Relatorio.csv";

        //    string aba = "1";
        //    string separador = ";";
        //    string[]? colunas = { "A", "b", "D" };
        //    string? linhas = ":";

        //    return ExcelHelper.ConverterExcept(origem, destino, aba, separador, colunas, linhas, null);

        //}

    }
}
