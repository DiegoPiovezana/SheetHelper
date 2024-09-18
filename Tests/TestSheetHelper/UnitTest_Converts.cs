using SH;

namespace TestSheetHelper
{
    // https://github.com/DiegoPiovezana/SheetHelper

    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void TestPass()
        {
            Assert.Pass();
        }

        // --------------------------------------------------------------------------------
        [Test, Repeat(1)]
        public void TestManipulacaoDtGZ()
        {
            string origin = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xlsx.gz";
            string destination = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\Especial_xlsx.csv";

            string sheet = "1"; // Use "1" for the first sheet (index or name)
            string delimiter = ";";
            string columns = "A, 3, b, 12:-1"; // or null to convert all columns or "A:BC" for a column range
            string rows = ":4, -2"; // Eg: Extracts from the 1nd to the 4nd row and also the penultimate row      

            var sh = new SheetHelper();
            var dt = sh.GetDataTable(origin, "1");
            var first = sh.GetRowArray(dt);
            bool success = sh.SaveDataTable(dt, destination, delimiter, columns, rows);

            Assert.That(success && first[99] == "100", Is.EqualTo(true));
        }


        [Test, Repeat(1)]
        public void TestDefaultReadme()
        {
            //string source = "C:\\Users\\Diego\\Files\\Report.xlsx.gz";
            //string destination = "C:\\Users\\Diego\\Files\\Report.csv";

            string source = @"C:\Users\diego\Desktop\Tests\Converter\ColunasExcel.zip";
            string destination = "C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\Report.csv";

            string sheet = "1"; // Use "1" for the first sheet (index or name)
            string delimiter = ";";
            string columns = "A, 3, b, -5:-1"; // or null to convert all columns or "A:BC" for a column range
            string rows = "1:2,:4, -2"; // Eg: Extracts from the 1nd to the 4nd row and also the penultimate row      

            var sh = new SheetHelper();
            bool success = sh.Converter(source, destination, sheet, delimiter, columns, rows);
            Assert.That(success, Is.EqualTo(true));
        }

        [Test, Repeat(1)]
        public void TestUTF8()
        {
            string source = @"C:\Users\diego\Desktop\Tests\Converter\ExcelUTF8.xlsx";
            string destination = "C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ExcelUTF8.csv";

            string sheet = "2"; // Use "1" for the first sheet (index or name)
            string delimiter = ";";
            string columns = "a:F"; // or null to convert all columns or "A:BC" for a column range
            string rows = "1:8"; // Eg: Extracts from the 1nd to the 4nd row and also the penultimate row      

            var sh = new SheetHelper();
            bool success = sh.Converter(source, destination, sheet, delimiter, columns, rows);
            Assert.That(success, Is.EqualTo(true));
        }

        [Test, Repeat(1)]
        public void TesteFormatoColunas()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\TesteFormatoColunas.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\TesteFormatoColunas.csv";

            string aba = "1";
            string separador = ";";

            string? colunas = "A:";
            string? linhas = ":";

            var sh = new SheetHelper();
            bool retorno = sh.Converter(origem, destino, aba, separador, colunas, linhas);
            Assert.That(retorno, Is.EqualTo(true));
        }

        [Test, Repeat(1)]
        public void TestTxtCsv()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.txt";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_txtCsv.csv";

            string aba = "1";
            string separador = ";";

            string? colunas = "A:";
            string? linhas = ":";

            bool retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);
            Assert.That(retorno, Is.EqualTo(true));
        }

        [Test, Repeat(1)]
        public void TestUnzipOnly()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.zip";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_unzipOnly.xlsx";

            string aba = "1";
            string separador = ";";

            string? colunas = null;
            string? linhas = null;

            bool retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);
            Assert.That(retorno, Is.EqualTo(true));
        }

        [Test, Repeat(1)]
        public void TestUnzip()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.zip";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_unzipXlsx.csv";

            string aba = "1";
            string separador = ";";

            string? colunas = null;
            string? linhas = null;

            bool retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);
            Assert.That(retorno, Is.EqualTo(true));
        }

        // --------------------------------------------------------------------------------

        [Test, Repeat(1)]
        public void TestConvertParticularXLSX()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_xlsx.csv";

            string aba = "sheet1"; // Utilize "1" para a primeira aba (�ndice ou nome)
            string separador = ";";
            //string[]? colunas = { "" }; // { "A", "C", "b" } ou null, para converter todas as colunas ou {"A:BC"} para intervalo de colunas
            string? colunas = "A,2,c";
            string? linhas = "1:3,4,-1"; // Ex.: se "2:" extrai a partir da 2� linha da planilha at� a �ltima (retira a 1� linha)          

            bool retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);
            Assert.That(retorno, Is.EqualTo(true));
        }

        [Test, Repeat(1)]
        public void TestConvertParticularCSV()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.csv";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_csv.csv";

            string aba = "1";
            string separador = ";";
            string? colunas = null;
            string linhas = "1:3";

            bool retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);
            Assert.That(retorno, Is.EqualTo(true));
        }

        [Test]
        public void TestConvertFromCSV()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\SemCabecalho.csv";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\SemCabecalho_csv.xls";

            string aba = "1";
            string separador = ";";
            string? colunas = null;
            string linhas = "1:5";

            var sh = new SheetHelper();
            var dt = sh.GetDataTable(origem, aba);
            var first = sh.GetRowArray(dt);
            var success = sh.SaveDataTable(dt, destino, separador, colunas, linhas);

            Assert.That(success, Is.EqualTo(true));
        }

        [Test]
        public void TestReadCSV()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.csv";
            string aba = "1";       

            var sh = new SheetHelper();
            var dt = sh.GetDataTable(origem, aba);
            var first = sh.GetRowArray(dt);

            Assert.That(first[2].Equals("3"), Is.EqualTo(true));
        }     

        [Test]
        public void TestReadCSVCabecalhoIrregular() // Read
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\CabecalhoIrregular.csv";
            //string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\SeparadorPipeline_CabecalhoIrregular.txt";
            string aba = "1";        

            var sh = new SheetHelper();
            sh.TryIgnoreExceptions.AddRange(new List<string>() { "E-4041-SH" });
            var dt = sh.GetDataTable(origem, aba);
            var first = sh.GetRowArray(dt);

            Assert.That(first[5].Equals("6") && first[6].Equals("Column6"), Is.EqualTo(true));
        }

        [Test]
        public void TestReadCSVSemCabecalho() // Read
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\SemCabecalho.csv";
            string aba = "1";

            var sh = new SheetHelper();
            var dt = sh.GetDataTable(origem, aba);
            var first = sh.GetRowArray(dt);

            Assert.That(first[2].Equals("C2"), Is.EqualTo(true));
        }

        [Test]
        public void TestReadCSVCabecalhoVazio() // Read
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\CabecalhoVazio.csv";
            string aba = "1";       

            var sh = new SheetHelper();
            sh.TryIgnoreExceptions.AddRange(new List<string>() { "E-4041-SH" });
            var dt = sh.GetDataTable(origem, aba);
            var first = sh.GetRowArray(dt);

            Assert.That(first[2].Equals("3"), Is.EqualTo(true));
        }

        [Test]
        public void TestReadCSVSeparatorPipeline()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\SeparadorPipeline.csv";
            string aba = "1";

            var sh = new SheetHelper();
            var dt = sh.GetDataTable(origem, aba);
            var first = sh.GetRowArray(dt);

            Assert.That(first[2].Equals("Last Name"), Is.EqualTo(true));
        }

        [Test, Repeat(1)]
        public void TestConvertParticularXLSB()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xlsb";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_xlsb.csv";

            string aba = "1";
            string separador = ";";
            string? colunas = null;
            string? linhas = null;

            bool retorno = new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);
            Assert.That(retorno, Is.EqualTo(true));
        }

        [Test]
        public void TestConvertTryIgnoreExpect()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\Small.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\Small_xlsb.csv";

            string aba = "1";
            string separador = ";";
            string? colunas = "1:12"; // Max column 10
            string? linhas = "1:10";

            SheetHelper sheethelper = new ();
            sheethelper.TryIgnoreExceptions.AddRange(new List<string>(){ "E-4042-SH" });

            bool retorno = sheethelper.Converter(origem, destino, aba, separador, colunas, linhas);
            Assert.That(retorno, Is.EqualTo(true));
        }

        [Test]
        public void TestConvertOneSheetBig()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ExcelBig_OneSheetBig_ABN204960.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ExcelBig_OneSheetBig_ABN204960_xlsx.csv";

            string aba = "1";
            string separador = ";";
            string? colunas = "";
            string? linhas = "1:";

            var sh = new SheetHelper();

            //sh.ProhibitedItems = new Dictionary<string, string>
            //{
            //    { "\n", " " },
            //    { "\r", " " },
            //    { ";", "," },
            //};

            bool retorno = sh.Converter(origem, destino, aba, separador, colunas, linhas);
            Assert.That(retorno, Is.EqualTo(true));
        }

        [Test]
        public void TestConvertManySheetBig()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ExcelBig_ManySheetBig_AB1048576.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ExcelBig_ManySheetBig_AB1048576_xlsx.csv";

            string aba = "1";
            string separador = ";";
            string? colunas = "";
            string? linhas = "1:";

            var sh = new SheetHelper();

            sh.ProhibitedItems = new Dictionary<string, string>
            {
                { "\n", " " },
                { "\r", " " },
                { ";", "," },
            };

            bool retorno = sh.Converter(origem, destino, aba, separador, colunas, linhas);
            Assert.That(retorno, Is.EqualTo(true));
        }

        [Test]
        public void TestConvertMultiDestinations()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\AbasExcel.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\AbasExcel_xlsx.csv";
                        
            var abas = new List<string>() { "2", "1", "sheet2" };
            string separador = ";";  
            string[]? colunas = new string[] { "A, B:C", "1:10", "B,A" };            
            List<string>? linhas = new() { "1:3", "1:10", "1" };
            int minRows = 1;

            var retorno = new SheetHelper().Converter(origem, destino, abas, separador, colunas, linhas, minRows);

            Assert.That(retorno == abas.Count, Is.EqualTo(true));
        }

        [Test]
        public void TestConvertMultiSheets1()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\AbasExcel.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\AbasExcel_xlsx.csv";

            //string aba = "2";
            //var abas = new List<string>() { "aba 3", "1", "sheet3" }; // Sheet1, Sheet2, aba 3
            var abas = new List<string>() { "aba 3", "1", "aba 3" };
            string separador = ";";
            //string? colunas = null;
            string[]? colunas = null;
            //string[]? colunas = new string[] { "A, B:C", "1:10", "B,A" };
            //string? linhas = null;
            //List<string>? linhas = new ();
            List<string>? linhas = new() { "1:3", "1:10", "1" };
            int minRows = 1;


            var retorno = new SheetHelper().Converter(origem, destino, abas, separador, colunas, linhas, minRows);

            Assert.That(retorno == abas.Count, Is.EqualTo(true));
        }

        [Test]
        public void TestConvertMultiSheets2()
        {
            var sheetHelper = new SheetHelper();

            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\AbasExcel.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\AbasExcel_xlsx.csv";

            //string aba = "2";
            string abas = null;
            string separador = ";";
            //string? colunas = null;
            string[]? colunas = null;
            //string[]? colunas = new string[] { "A, B:C", "1:10", "B,A" };
            //string? linhas = null;
            //List<string>? linhas = new ();
            List<string>? linhas = null;
            int minRows = 1;

            Assert.Throws<SH.Exceptions.ParamMissingConverterSHException>(() =>
                       sheetHelper.Converter(origem, destino, abas, separador, colunas, linhas, minRows)
                   );
        }

        [Test]
        public void TestConvertMultiSheets3()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\AbasExcel.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\AbasExcel_xlsx.csv";

            //string aba = "2";
            var abas = new List<string>() { "Sheet2", "1", "aba 3" };
            string separador = ";";
            //string? colunas = null;
            string[]? colunas = null;
            //string[]? colunas = new string[] { "A, B:C", "1:10", "B,A" };
            //string? linhas = null;
            //List<string>? linhas = new ();
            List<string>? linhas = new() { "1:3", "1:10", "1" };
            int minRows = 1;


            var retorno = new SheetHelper().Converter(origem, destino, abas, separador, colunas, linhas, minRows);

            Assert.That(retorno == abas.Count, Is.EqualTo(true));
        }

        [Test]
        public void TestProhibitedItems()
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ProhibitedItems.xlsx";
            string destino = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ProhibitedItems_xlsx.csv";

            string aba = "3";
            string separador = ";";
            string? colunas = null;
            string? linhas = null;

            var dic = new Dictionary<string, string>
            {
                { "\n", " " },
                { "\r", " " },
                { ";", "," },
            };

            var sh = new SheetHelper();

            //sh.ProhibitedItems = dic;

            // "{"key1": "value1", "key2": "value2", "key3": "value3"}";
            string test1 = "{ \"key1\" : \"value1\", \"key2\" : \"value2\", \"key3\" : \"value3\" }";
            string test2 = "{\"\\n\": \" \", \"\\r\": \"\", \";\": \",\"}";
            string jsonDictionary1 = System.Text.Json.JsonSerializer.Serialize(dic);
            string jsonDictionary2 = sh.GetJsonDictionary(dic);

            sh.ProhibitedItems = sh.GetDictionaryJson(jsonDictionary2);

            var retorno = sh.Converter(origem, destino, aba, separador, colunas, linhas);

            Assert.That(retorno, Is.EqualTo(true));
        }

        [Test]
        public void TestConvertRowsToBack()
        {
            string origin = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xlsx";
            string destination = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_BackRows_xlsx.csv";

            string sheet = "1";
            string separator = ";";
            string? columns = "";
            //string? rows = "1:3, 5, 1; 2";
            string? rows = "3:1, -1:-2; -2:-1";

            var sh = new SheetHelper();

            bool result = sh.Converter(origin, destination, sheet, separator, columns, rows);
            Assert.That(result, Is.EqualTo(true));
        }

        [Test]
        public void TestConvertColumnsToBack()
        {
            string origin = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xlsx";
            string destination = $"C:\\Users\\diego\\Desktop\\Tests\\Convertidos\\ColunasExcel_BackColumns_xlsx.csv";

            string sheet = "1";
            string separator = ";";
            string? columns = "C:A, A; -1, -2:-3; -3:-1";
            string? rows = "1:3, 1, 3:1";

            var sh = new SheetHelper();
            bool result = sh.Converter(origin, destination, sheet, separator, columns, rows);
            Assert.That(result, Is.EqualTo(true));
        }

        // --------------------------------------------------------------------------------

        [TestCase("2:", 1, ExpectedResult = true, TestName = "2:")]
        [TestCase(":10", 2, ExpectedResult = true, TestName = ":10")]
        [TestCase("1:20", 3, ExpectedResult = true, TestName = "1:20")]
        [TestCase("", 4, ExpectedResult = true, TestName = "Linhas String vazia")]
        [TestCase(null, 5, ExpectedResult = true, TestName = "Linhas Nulo")]
        [TestCase("1, 2, 4", 6, ExpectedResult = true, TestName = "1, 2, 4")]
        [TestCase("7", 7, ExpectedResult = true, TestName = "7")]
        [TestCase("-1", 8, ExpectedResult = true, TestName = "-1")]
        [TestCase("1:-2", 9, ExpectedResult = true, TestName = "1:-2")]
        [TestCase("1,2,3,-3:", 10, ExpectedResult = true, TestName = "1,2,3,-3:-1")]
        public bool TestRows(string linhas, int id)
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.csv";
            string destino = @$"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_row{id}.csv";

            string aba = "1";
            string separador = ";";
            string colunas = "A, 2,c";

            var sh = new SheetHelper();
            return sh.Converter(origem, destino, aba, separador, colunas, linhas);
        }

        [TestCase("A, C, b", 1, ExpectedResult = true, TestName = "Colunas maiuscula e minuscula e fora de ordem...")]
        [TestCase("A:D", 2, ExpectedResult = true, TestName = "Colunas em intervalo cont�nuo...")]
        [TestCase("", 3, ExpectedResult = true, TestName = "Colunas com string vazia...")]
        [TestCase(" ", 4, ExpectedResult = true, TestName = "Colunas vazio...")]
        [TestCase(null, 5, ExpectedResult = true, TestName = "Colunas nulo...")]
        public bool TestColuns(string colunas, int id)
        {
            string origem = "C:\\Users\\diego\\Desktop\\Tests\\Converter\\ColunasExcel.xls";
            string destino = @$"C:\Users\diego\Desktop\Tests\Convertidos\ColunasExcel_column{id}.csv";

            string aba = "1";
            string separador = ";";
            string linhas = "2:";

            var sh = new SheetHelper();
            return sh.Converter(origem, destino, aba, separador, colunas, linhas);
        }

        // --------------------------------------------------------------------------------


        [TestFixture]
        public class TestsFormats
        {
            [Test, TestCaseSource(typeof(CasosDeTesteDeFormatos), nameof(CasosDeTesteDeFormatos.FormatsConverts))]
            public bool ValidarFormatosValidos(string origin, string destination) => TestFormats(origin, destination);
        }

        public class CasosDeTesteDeFormatos
        {
            public static List<TestCaseData> FormatsConverts
            {
                get
                {
                    string path = "C:\\Users\\diego\\Desktop\\Tests";

                    return new List<TestCaseData>()
                     {
                         //new TestCaseData("", "").Returns(new Exception().Message == "'' � inv�lido!").SetName("Origem e destino vazio"),
                         
                         // Txt
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.txt", $"{path}\\Convertidos\\ColunasExcel_TXT.csv").Returns(true).SetName("TXT__CSV"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.txt", $"{path}\\Convertidos\\ColunasExcel_TXT.txt").Returns(true).SetName("TXT__TXT"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.txt", $"{path}\\Convertidos\\ColunasExcel_TXT.xls").Returns(true).SetName("TXT__XLS"),

                         // Csv
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.csv", $"{path}\\Convertidos\\ColunasExcel_CSV.csv").Returns(true).SetName("CSV__CSV"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.csv", $"{path}\\Convertidos\\ColunasExcel_CSV.txt").Returns(true).SetName("CSV__TXT"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.csv", $"{path}\\Convertidos\\ColunasExcel_CSV.xls").Returns(true).SetName("CSV__XLS"),
                             
                         // Xls
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xls", $"{path}\\Convertidos\\ColunasExcel_XLS.csv").Returns(true).SetName("XLS__CSV"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xls", $"{path}\\Convertidos\\ColunasExcel_XLS.txt").Returns(true).SetName("XLS__TXT"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xls", $"{path}\\Convertidos\\ColunasExcel_XLS.xls").Returns(true).SetName("XLS__XLS"),

                         // Xlsb
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsb", $"{path}\\Convertidos\\ColunasExcel_XLSB.csv").Returns(true).SetName("XLSB__CSV"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsb", $"{path}\\Convertidos\\ColunasExcel_XLSB.txt").Returns(true).SetName("XLSB__TXT"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsb", $"{path}\\Convertidos\\ColunasExcel_XLSB.xls").Returns(true).SetName("XLSB__XLS"),

                         // Xlsx
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx", $"{path}\\Convertidos\\ColunasExcel_XSLX.csv").Returns(true).SetName("XSLX__CSV"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx", $"{path}\\Convertidos\\ColunasExcel_XSLX.txt").Returns(true).SetName("XSLX__TXT"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx", $"{path}\\Convertidos\\ColunasExcel_XSLX.html").Returns(true).SetName("XSLX__HTML"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx", $"{path}\\Convertidos\\ColunasExcel_XSLX.xls").Returns(true).SetName("XSLX__XLS"),

                         // Xlsm
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsm",$"{path}\\Convertidos\\ColunasExcel_XLSM.csv").Returns(true).SetName("XLSM__CSV"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsm",$"{path}\\Convertidos\\ColunasExcel_XLSM.txt").Returns(true).SetName("XLSM__TXT"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsm",$"{path}\\Convertidos\\ColunasExcel_XLSM.html").Returns(true).SetName("XLSM__HTML"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsm",$"{path}\\Convertidos\\ColunasExcel_XLSM.xls").Returns(true).SetName("XLSM__XLS"),
                             
                         // Zip
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx.zip", $"{path}\\Convertidos\\ColunasExcel_XLSX_ZIP.csv").Returns(true).SetName("Z_XLSX_ZIP__csv"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.xlsx.gz",$"{path}\\Convertidos\\ColunasExcel_XLSX_GZ.csv").Returns(true).SetName("Z_XLSX_GZ__CSV"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.gz", $"{path}\\Convertidos\\ColunasExcel_GZ.csv").Returns(true).SetName("Z_GZ__CSV"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.tar.gz", $"{path}\\Convertidos\\ColunasExcel_TAR_GZ.csv").Returns(true).SetName("Z_TAR_GZ__CSV"),
                         new TestCaseData($"{path}\\Converter\\ColunasExcel.zip", $"{path}\\Convertidos\\ColunasExcel_ZIP.csv").Returns(true).SetName("Z_ZIP__CSV"),



                    };
                }
            }
        }

        private static bool TestFormats(string origem, string destino)
        {
            //string origem = "C:\\Users\\diego\\Arquivos\\Relatorio.xlsx.gz";
            //string destino = "C:\\Users\\diego\\Arquivos\\Relatorio.csv";

            string aba = "1";
            string separador = ";";
            string? colunas = "A, b, 4,-1";
            string? linhas = ":10,-1";

            return new SheetHelper().Converter(origem, destino, aba, separador, colunas, linhas);

        }




    }
}