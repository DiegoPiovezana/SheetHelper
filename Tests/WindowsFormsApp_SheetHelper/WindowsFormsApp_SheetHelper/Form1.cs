using SH;
using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace WindowsFormsApp_SheetHelper
{
    public partial class SheetHelper_Menu : Form
    {
        // Crie um Timer
        private readonly System.Windows.Forms.Timer updateTimer;


        public SheetHelper_Menu()
        {
            InitializeComponent();

            // Configurar o Timer
            updateTimer = new System.Windows.Forms.Timer { Interval = 100 };
            updateTimer.Tick += UpdateProgressBar; // Define o método a ser chamado pelo Timer           
        }

        // Método para atualizar a barra de progresso
        private void UpdateProgressBar(object sender, EventArgs e)
        {
            // Atualiza a barra de progresso com o valor atual de SheetHelper.Progress
            pgBarConvert.Value = SheetHelper.Progress;
            this.lblConvertendo.Text = $"Convertendo... {pgBarConvert.Value}%";
        }

        private void SheetHelper_Menu_Load(object sender, EventArgs e)
        {
            lblConvertendo.Visible = false;
            pgBarConvert.Style = ProgressBarStyle.Continuous;
        }

        // Código para iniciar a conversão
        private void Button1_Click(object sender, EventArgs e)
        {
            this.lblConvertendo.Visible = true;
            BtnConverter.Enabled = false;

            updateTimer.Start();

            bool retorno = false;
            // Inicia a conversão em uma nova thread
            Thread converterThread = new Thread(() =>
            {
                // Inicia o cronômetro
                Stopwatch stopwatch = Stopwatch.StartNew();

                // Realiza a conversão
                //retorno = Converter();
                retorno = Program.Converter(pgBarConvert);

                updateTimer.Stop();

                // Após a conclusão da conversão, exibe o tempo decorrido
                stopwatch.Stop();
                TimeSpan tempoDecorrido = stopwatch.Elapsed;                            

                Debug.WriteLine("____________________________________________________\n\n\n\n\n");
                Debug.WriteLine($"Tempo necessário para conversão: {tempoDecorrido:mm\\:ss\\.fff}\n\n");
                Debug.WriteLine("____________________________________________________");

                // Atualiza a interface do usuário após a conclusão da conversão
                this.Invoke((MethodInvoker)delegate
                {
                    this.lblConvertendo.Text = $"Conversão finalizada!";
                    pgBarConvert.Value = 100;
                    //Debug.WriteLine(pgBarConvert.Value);
                    //Debug.WriteLine(SheetHelper.Progress);

                    MessageBox.Show(retorno ? $"O arquivo foi convertido com sucesso em {tempoDecorrido.TotalMilliseconds} ms!" : "Não foi possível converter o arquivo!");

                    this.Enabled = true;
                    BtnConverter.Enabled = true;
                });
            });

            converterThread.Start();   

            // Código a ser executado ao mesmo tempo que a conversão
            // ...
        }

        private bool Converter()
        {

            string origem = "C:\\Users\\diego\\Desktop\\Lixo\\Converter\\ColunasExcelBig.xlsb";
            string destino = "C:\\Users\\diego\\Desktop\\Lixo\\Convertidos\\ConversaoBig_XLSB.csv";
            string aba = "2";
            string linhas = "1:"; // 502.383
            string colunas = "";

            // TESTES DE TEMPO - ColunasExcelBig.xlsb:
            // 00:42.421


            //retorno = SheetHelper.ConvertAllSheet(origem, destino);
            return SheetHelper.Converter(origem, destino, aba, ";", colunas, linhas);

        }
    }
}
