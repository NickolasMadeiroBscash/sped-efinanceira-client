using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExemploAssinadorXML.Models;
using ExemploAssinadorXML.Services;

namespace ExemploAssinadorXML.Forms
{
    public partial class ProcessamentoForm : Form
    {
        private GroupBox grpControles;
        private Button btnProcessarAbertura;
        private Button btnProcessarMovimentacao;
        private Button btnProcessarFechamento;
        private Button btnCancelar;
        private CheckBox chkApenasProcessar;

        private GroupBox grpProgresso;
        private Label lblEtapaAtual;
        private Label lblMensagemAtual;
        private ProgressBar progressBarGeral;
        private Label lblProgressoGeral;
        private ListBox lstLog;

        private GroupBox grpEstatisticas;
        private Label lblTotalLotes;
        private Label lblLotesProcessados;
        private Label lblLotesAssinados;
        private Label lblLotesCriptografados;
        private Label lblLotesEnviados;
        private Label lblLotesComErro;
        private Label lblTempoDecorrido;
        private Label lblTempoEstimado;

        private StatusProcessamento status;
        private bool cancelarProcessamento = false;

        public ConfiguracaoForm ConfigForm { get; set; }
        public ConsultaForm ConsultaForm { get; set; }

        public ProcessamentoForm()
        {
            InitializeComponent();
            status = new StatusProcessamento();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Controles
            grpControles = new GroupBox();
            grpControles.Text = "Controles";
            grpControles.Location = new Point(10, 10);
            grpControles.Size = new Size(750, 80);
            grpControles.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            btnProcessarAbertura = new Button();
            btnProcessarAbertura.Text = "Processar Abertura";
            btnProcessarAbertura.Location = new Point(10, 25);
            btnProcessarAbertura.Size = new Size(150, 35);
            btnProcessarAbertura.Click += BtnProcessarAbertura_Click;

            btnProcessarMovimentacao = new Button();
            btnProcessarMovimentacao.Text = "Processar Movimentação";
            btnProcessarMovimentacao.Location = new Point(170, 25);
            btnProcessarMovimentacao.Size = new Size(150, 35);
            btnProcessarMovimentacao.Click += BtnProcessarMovimentacao_Click;

            btnProcessarFechamento = new Button();
            btnProcessarFechamento.Text = "Processar Fechamento";
            btnProcessarFechamento.Location = new Point(330, 25);
            btnProcessarFechamento.Size = new Size(150, 35);
            btnProcessarFechamento.Click += BtnProcessarFechamento_Click;

            btnCancelar = new Button();
            btnCancelar.Text = "Cancelar";
            btnCancelar.Location = new Point(490, 25);
            btnCancelar.Size = new Size(100, 35);
            btnCancelar.Enabled = false;
            btnCancelar.Click += BtnCancelar_Click;

            chkApenasProcessar = new CheckBox();
            chkApenasProcessar.Text = "Apenas Processar (não enviar)";
            chkApenasProcessar.Location = new Point(600, 32);
            chkApenasProcessar.Size = new Size(200, 20);
            chkApenasProcessar.Checked = false;

            grpControles.Controls.AddRange(new Control[] {
                btnProcessarAbertura, btnProcessarMovimentacao, btnProcessarFechamento,
                btnCancelar, chkApenasProcessar
            });

            // Progresso
            grpProgresso = new GroupBox();
            grpProgresso.Text = "Progresso";
            grpProgresso.Location = new Point(10, 100);
            grpProgresso.Size = new Size(750, 200);
            grpProgresso.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            lblEtapaAtual = new Label();
            lblEtapaAtual.Text = "Etapa: Aguardando...";
            lblEtapaAtual.Location = new Point(10, 25);
            lblEtapaAtual.Size = new Size(730, 20);
            lblEtapaAtual.Font = new Font(lblEtapaAtual.Font, FontStyle.Bold);

            lblMensagemAtual = new Label();
            lblMensagemAtual.Text = "Mensagem: -";
            lblMensagemAtual.Location = new Point(10, 50);
            lblMensagemAtual.Size = new Size(730, 20);

            progressBarGeral = new ProgressBar();
            progressBarGeral.Location = new Point(10, 80);
            progressBarGeral.Size = new Size(730, 25);
            progressBarGeral.Style = ProgressBarStyle.Continuous;

            lblProgressoGeral = new Label();
            lblProgressoGeral.Text = "0%";
            lblProgressoGeral.Location = new Point(10, 110);
            lblProgressoGeral.Size = new Size(730, 20);

            lstLog = new ListBox();
            lstLog.Location = new Point(10, 135);
            lstLog.Size = new Size(730, 60);
            lstLog.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            grpProgresso.Controls.AddRange(new Control[] {
                lblEtapaAtual, lblMensagemAtual, progressBarGeral,
                lblProgressoGeral, lstLog
            });

            // Estatísticas
            grpEstatisticas = new GroupBox();
            grpEstatisticas.Text = "Estatísticas";
            grpEstatisticas.Location = new Point(10, 310);
            grpEstatisticas.Size = new Size(750, 150);
            grpEstatisticas.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            lblTotalLotes = new Label();
            lblTotalLotes.Text = "Total de Lotes: 0";
            lblTotalLotes.Location = new Point(10, 25);
            lblTotalLotes.Size = new Size(200, 20);

            lblLotesProcessados = new Label();
            lblLotesProcessados.Text = "Processados: 0";
            lblLotesProcessados.Location = new Point(220, 25);
            lblLotesProcessados.Size = new Size(200, 20);

            lblLotesAssinados = new Label();
            lblLotesAssinados.Text = "Assinados: 0";
            lblLotesAssinados.Location = new Point(430, 25);
            lblLotesAssinados.Size = new Size(200, 20);

            lblLotesCriptografados = new Label();
            lblLotesCriptografados.Text = "Criptografados: 0";
            lblLotesCriptografados.Location = new Point(10, 55);
            lblLotesCriptografados.Size = new Size(200, 20);

            lblLotesEnviados = new Label();
            lblLotesEnviados.Text = "Enviados: 0";
            lblLotesEnviados.Location = new Point(220, 55);
            lblLotesEnviados.Size = new Size(520, 20);

            lblLotesComErro = new Label();
            lblLotesComErro.Text = "Com Erro: 0";
            lblLotesComErro.Location = new Point(10, 85);
            lblLotesComErro.Size = new Size(200, 20);
            lblLotesComErro.ForeColor = Color.Red;

            lblTempoDecorrido = new Label();
            lblTempoDecorrido.Text = "Tempo Decorrido: 00:00:00";
            lblTempoDecorrido.Location = new Point(220, 85);
            lblTempoDecorrido.Size = new Size(300, 20);

            lblTempoEstimado = new Label();
            lblTempoEstimado.Text = "Tempo Estimado Restante: -";
            lblTempoEstimado.Location = new Point(320, 85);
            lblTempoEstimado.Size = new Size(300, 20);

            grpEstatisticas.Controls.AddRange(new Control[] {
                lblTotalLotes, lblLotesProcessados, lblLotesAssinados,
                lblLotesCriptografados, lblLotesEnviados, lblLotesComErro,
                lblTempoDecorrido, lblTempoEstimado
            });

            this.Controls.AddRange(new Control[] {
                grpControles, grpProgresso, grpEstatisticas
            });

            this.ResumeLayout(false);
        }

        private void BtnProcessarAbertura_Click(object sender, EventArgs e)
        {
            string erro = ValidarDadosAbertura();
            if (!string.IsNullOrEmpty(erro))
            {
                string mensagemFinal = erro.Contains("Campo:") || erro.Contains("Aba") 
                    ? $"Não é possível processar o evento de Abertura.\n\nOs seguintes campos obrigatórios estão faltando:\n\n{erro}\n\nPor favor, preencha os campos indicados acima na aba Configuração."
                    : $"Não é possível processar o evento de Abertura.\n\n{erro}";
                MessageBox.Show(mensagemFinal, 
                    "Campos Obrigatórios Faltando", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnProcessarAbertura.Enabled = false;
            btnProcessarMovimentacao.Enabled = false;
            btnProcessarFechamento.Enabled = false;
            btnCancelar.Enabled = true;
            cancelarProcessamento = false;

            Task.Run(() => ProcessarAbertura());
        }

        private void BtnProcessarMovimentacao_Click(object sender, EventArgs e)
        {
            string erro = ValidarDadosMovimentacao();
            if (!string.IsNullOrEmpty(erro))
            {
                string mensagemFinal = erro.Contains("Campo:") || erro.Contains("Aba") 
                    ? $"Não é possível processar o evento de Movimentação.\n\nOs seguintes campos obrigatórios estão faltando:\n\n{erro}\n\nPor favor, preencha os campos indicados acima na aba Configuração."
                    : $"Não é possível processar o evento de Movimentação.\n\n{erro}";
                MessageBox.Show(mensagemFinal, 
                    "Campos Obrigatórios Faltando", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnProcessarAbertura.Enabled = false;
            btnProcessarMovimentacao.Enabled = false;
            btnProcessarFechamento.Enabled = false;
            btnCancelar.Enabled = true;
            cancelarProcessamento = false;

            Task.Run(() => ProcessarMovimentacao());
        }

        private void BtnProcessarFechamento_Click(object sender, EventArgs e)
        {
            string erro = ValidarDadosFechamento();
            if (!string.IsNullOrEmpty(erro))
            {
                string mensagemFinal = erro.Contains("Campo:") || erro.Contains("Aba") 
                    ? $"Não é possível processar o evento de Fechamento.\n\nOs seguintes campos obrigatórios estão faltando:\n\n{erro}\n\nPor favor, preencha os campos indicados acima na aba Configuração."
                    : $"Não é possível processar o evento de Fechamento.\n\n{erro}";
                MessageBox.Show(mensagemFinal, 
                    "Campos Obrigatórios Faltando", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnProcessarAbertura.Enabled = false;
            btnProcessarMovimentacao.Enabled = false;
            btnProcessarFechamento.Enabled = false;
            btnCancelar.Enabled = true;
            cancelarProcessamento = false;

            Task.Run(() => ProcessarFechamento());
        }

        private void BtnCancelar_Click(object sender, EventArgs e)
        {
            cancelarProcessamento = true;
            AdicionarLog("Processamento cancelado pelo usuário.");
        }

        private string ValidarDadosAbertura()
        {
            List<string> camposFaltando = new List<string>();

            if (ConfigForm == null)
            {
                camposFaltando.Add("Acesse a aba 'Configuração' e preencha todos os campos obrigatórios");
                camposFaltando.Add("Campo: 'CNPJ do Declarante' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado para Assinatura' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado do Servidor' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Período' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Diretório de Lotes' na seção 'Configuração Geral'");
                camposFaltando.Add("Aba 'Dados de Abertura': Preencha todos os campos obrigatórios");
                return string.Join("\n", camposFaltando);
            }

            if (ConfigForm.Config == null)
            {
                camposFaltando.Add("Configuração geral não foi inicializada. Preencha os seguintes campos:");
                camposFaltando.Add("Campo: 'CNPJ do Declarante' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado para Assinatura' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado do Servidor' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Período' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Diretório de Lotes' na seção 'Configuração Geral'");
                return string.Join("\n", camposFaltando);
            }

            var config = ConfigForm.Config;
            var dadosAbertura = ConfigForm.DadosAbertura;

            if (dadosAbertura == null)
            {
                camposFaltando.Add("Dados de abertura não foram configurados. Preencha os seguintes campos:");
                camposFaltando.Add("Aba 'Dados de Abertura': Preencha 'Data de Início'");
                camposFaltando.Add("Aba 'Dados de Abertura': Preencha 'Data de Fim'");
                camposFaltando.Add("Aba 'Dados de Abertura': Configure 'Responsável RMF'");
                camposFaltando.Add("Aba 'Dados de Abertura': Configure 'Responsável e-Financeira'");
                camposFaltando.Add("Aba 'Dados de Abertura': Configure 'Representante Legal'");
                return string.Join("\n", camposFaltando);
            }

            // Validações gerais
            if (string.IsNullOrWhiteSpace(config.CnpjDeclarante))
                camposFaltando.Add("Campo: 'CNPJ do Declarante' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.CertThumbprint))
                camposFaltando.Add("Campo: 'Certificado para Assinatura' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.CertServidorThumbprint))
                camposFaltando.Add("Campo: 'Certificado do Servidor' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.Periodo))
                camposFaltando.Add("Campo: 'Período' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.DiretorioLotes))
                camposFaltando.Add("Campo: 'Diretório de Lotes' na seção 'Configuração Geral'");

            // Validações específicas de abertura
            if (string.IsNullOrWhiteSpace(dadosAbertura.DtInicio))
                camposFaltando.Add("Campo: 'Data de Início' na aba 'Dados de Abertura'");

            if (string.IsNullOrWhiteSpace(dadosAbertura.DtFim))
                camposFaltando.Add("Campo: 'Data de Fim' na aba 'Dados de Abertura'");

            // Validações de Responsável RMF
            if (dadosAbertura.IndicarMovOpFin)
            {
                if (dadosAbertura.ResponsavelRMF == null)
                {
                    camposFaltando.Add("Seção completa: 'Responsável RMF' na aba 'Dados de Abertura' não está configurada");
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.Cnpj))
                        camposFaltando.Add("Campo: 'CNPJ' na seção 'Responsável RMF' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.Cpf))
                        camposFaltando.Add("Campo: 'CPF' na seção 'Responsável RMF' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.Nome))
                        camposFaltando.Add("Campo: 'Nome' na seção 'Responsável RMF' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.Setor))
                        camposFaltando.Add("Campo: 'Setor' na seção 'Responsável RMF' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.TelefoneDDD))
                        camposFaltando.Add("Campo: 'DDD' do Telefone na seção 'Responsável RMF' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.TelefoneNumero))
                        camposFaltando.Add("Campo: 'Número' do Telefone na seção 'Responsável RMF' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.EnderecoLogradouro))
                        camposFaltando.Add("Campo: 'Logradouro' na seção 'Responsável RMF' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.EnderecoNumero))
                        camposFaltando.Add("Campo: 'Número' do Endereço na seção 'Responsável RMF' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.EnderecoBairro))
                        camposFaltando.Add("Campo: 'Bairro' na seção 'Responsável RMF' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.EnderecoCEP))
                        camposFaltando.Add("Campo: 'CEP' na seção 'Responsável RMF' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.EnderecoMunicipio))
                        camposFaltando.Add("Campo: 'Município' na seção 'Responsável RMF' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.ResponsavelRMF.EnderecoUF))
                        camposFaltando.Add("Campo: 'UF' na seção 'Responsável RMF' (aba Dados de Abertura)");
                }
            }

            // Validações de Responsável e-Financeira
            if (dadosAbertura.IndicarMovOpFin)
            {
                if (dadosAbertura.RespeFin == null)
                {
                    camposFaltando.Add("Seção completa: 'Responsável e-Financeira' na aba 'Dados de Abertura' não está configurada");
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.Cpf))
                        camposFaltando.Add("Campo: 'CPF' na seção 'Responsável e-Financeira' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.Nome))
                        camposFaltando.Add("Campo: 'Nome' na seção 'Responsável e-Financeira' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.Setor))
                        camposFaltando.Add("Campo: 'Setor' na seção 'Responsável e-Financeira' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.TelefoneDDD))
                        camposFaltando.Add("Campo: 'DDD' do Telefone na seção 'Responsável e-Financeira' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.TelefoneNumero))
                        camposFaltando.Add("Campo: 'Número' do Telefone na seção 'Responsável e-Financeira' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.EnderecoLogradouro))
                        camposFaltando.Add("Campo: 'Logradouro' na seção 'Responsável e-Financeira' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.EnderecoNumero))
                        camposFaltando.Add("Campo: 'Número' do Endereço na seção 'Responsável e-Financeira' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.EnderecoBairro))
                        camposFaltando.Add("Campo: 'Bairro' na seção 'Responsável e-Financeira' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.EnderecoCEP))
                        camposFaltando.Add("Campo: 'CEP' na seção 'Responsável e-Financeira' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.EnderecoMunicipio))
                        camposFaltando.Add("Campo: 'Município' na seção 'Responsável e-Financeira' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.EnderecoUF))
                        camposFaltando.Add("Campo: 'UF' na seção 'Responsável e-Financeira' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RespeFin.Email))
                        camposFaltando.Add("Campo: 'E-mail' na seção 'Responsável e-Financeira' (aba Dados de Abertura)");
                }
            }

            // Validações de Representante Legal
            if (dadosAbertura.IndicarMovOpFin)
            {
                if (dadosAbertura.RepresLegal == null)
                {
                    camposFaltando.Add("Seção completa: 'Representante Legal' na aba 'Dados de Abertura' não está configurada");
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(dadosAbertura.RepresLegal.Cpf))
                        camposFaltando.Add("Campo: 'CPF' na seção 'Representante Legal' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RepresLegal.Setor))
                        camposFaltando.Add("Campo: 'Setor' na seção 'Representante Legal' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RepresLegal.TelefoneDDD))
                        camposFaltando.Add("Campo: 'DDD' do Telefone na seção 'Representante Legal' (aba Dados de Abertura)");

                    if (string.IsNullOrWhiteSpace(dadosAbertura.RepresLegal.TelefoneNumero))
                        camposFaltando.Add("Campo: 'Número' do Telefone na seção 'Representante Legal' (aba Dados de Abertura)");
                }
            }

            if (camposFaltando.Count > 0)
            {
                return "Os seguintes campos obrigatórios estão faltando:\n\n" + string.Join("\n", camposFaltando);
            }

            return null;
        }

        private string ValidarDadosMovimentacao()
        {
            List<string> camposFaltando = new List<string>();

            if (ConfigForm == null)
            {
                camposFaltando.Add("Acesse a aba 'Configuração' e preencha todos os campos obrigatórios");
                camposFaltando.Add("Campo: 'CNPJ do Declarante' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado para Assinatura' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado do Servidor' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Período' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Diretório de Lotes' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Page Size' na seção 'Configuração de Processamento'");
                return string.Join("\n", camposFaltando);
            }

            if (ConfigForm.Config == null)
            {
                camposFaltando.Add("Configuração geral não foi inicializada. Preencha os seguintes campos:");
                camposFaltando.Add("Campo: 'CNPJ do Declarante' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado para Assinatura' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado do Servidor' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Período' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Diretório de Lotes' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Page Size' na seção 'Configuração de Processamento'");
                return string.Join("\n", camposFaltando);
            }

            var config = ConfigForm.Config;

            // Validações gerais obrigatórias para movimentação
            if (string.IsNullOrWhiteSpace(config.CnpjDeclarante))
                camposFaltando.Add("Campo: 'CNPJ do Declarante' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.CertThumbprint))
                camposFaltando.Add("Campo: 'Certificado para Assinatura' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.CertServidorThumbprint))
                camposFaltando.Add("Campo: 'Certificado do Servidor' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.Periodo))
                camposFaltando.Add("Campo: 'Período' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.DiretorioLotes))
                camposFaltando.Add("Campo: 'Diretório de Lotes' na seção 'Configuração Geral'");

            if (config.PageSize <= 0)
                camposFaltando.Add("Campo: 'Page Size' na seção 'Configuração de Processamento' (deve ser maior que zero)");

            if (camposFaltando.Count > 0)
            {
                return "Os seguintes campos obrigatórios estão faltando:\n\n" + string.Join("\n", camposFaltando);
            }

            return null;
        }

        private string ValidarDadosFechamento()
        {
            List<string> camposFaltando = new List<string>();

            if (ConfigForm == null)
            {
                camposFaltando.Add("Acesse a aba 'Configuração' e preencha todos os campos obrigatórios");
                camposFaltando.Add("Campo: 'CNPJ do Declarante' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado para Assinatura' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado do Servidor' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Período' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Diretório de Lotes' na seção 'Configuração Geral'");
                camposFaltando.Add("Aba 'Dados de Fechamento': Preencha 'Data de Início'");
                camposFaltando.Add("Aba 'Dados de Fechamento': Preencha 'Data de Fim'");
                return string.Join("\n", camposFaltando);
            }

            if (ConfigForm.Config == null)
            {
                camposFaltando.Add("Configuração geral não foi inicializada. Preencha os seguintes campos:");
                camposFaltando.Add("Campo: 'CNPJ do Declarante' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado para Assinatura' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Certificado do Servidor' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Período' na seção 'Configuração Geral'");
                camposFaltando.Add("Campo: 'Diretório de Lotes' na seção 'Configuração Geral'");
                return string.Join("\n", camposFaltando);
            }

            var config = ConfigForm.Config;
            var dadosFechamento = ConfigForm.DadosFechamento;

            if (dadosFechamento == null)
            {
                camposFaltando.Add("Dados de fechamento não foram configurados. Preencha os seguintes campos:");
                camposFaltando.Add("Aba 'Dados de Fechamento': Preencha 'Data de Início'");
                camposFaltando.Add("Aba 'Dados de Fechamento': Preencha 'Data de Fim'");
                return string.Join("\n", camposFaltando);
            }

            // Validações gerais
            if (string.IsNullOrWhiteSpace(config.CnpjDeclarante))
                camposFaltando.Add("Campo: 'CNPJ do Declarante' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.CertThumbprint))
                camposFaltando.Add("Campo: 'Certificado para Assinatura' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.CertServidorThumbprint))
                camposFaltando.Add("Campo: 'Certificado do Servidor' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.Periodo))
                camposFaltando.Add("Campo: 'Período' na seção 'Configuração Geral'");

            if (string.IsNullOrWhiteSpace(config.DiretorioLotes))
                camposFaltando.Add("Campo: 'Diretório de Lotes' na seção 'Configuração Geral'");

            // Validações específicas de fechamento
            if (string.IsNullOrWhiteSpace(dadosFechamento.DtInicio))
                camposFaltando.Add("Campo: 'Data de Início' na aba 'Dados de Fechamento'");

            if (string.IsNullOrWhiteSpace(dadosFechamento.DtFim))
                camposFaltando.Add("Campo: 'Data de Fim' na aba 'Dados de Fechamento'");

            if (camposFaltando.Count > 0)
            {
                return "Os seguintes campos obrigatórios estão faltando:\n\n" + string.Join("\n", camposFaltando);
            }

            return null;
        }

        private async Task ProcessarAbertura()
        {
            try
            {
                AtualizarEtapa("Iniciando processamento de abertura...");
                status.InicioProcessamento = DateTime.Now;
                status.TotalLotes = 1;
                status.ProtocolosEnviados.Clear();

                var config = ConfigForm.Config;
                var dadosAbertura = ConfigForm.DadosAbertura;

                // 1. Gerar XML
                AtualizarEtapa("Gerando XML de abertura...");
                var geradorService = new EfinanceiraGeradorXmlService();
                string arquivoXml = geradorService.GerarXmlAbertura(dadosAbertura, config.DiretorioLotes);
                AdicionarLog($"XML gerado: {arquivoXml}");

                if (cancelarProcessamento) return;

                // Registrar lote no banco após gerar XML
                int quantidadeEventos = ProtocoloPersistenciaService.ContarEventosNoXml(arquivoXml);
                long idLoteBanco = 0;
                try
                {
                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                    string ambienteStr = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : "HOMOLOG";
                    idLoteBanco = persistenceService.RegistrarLote(
                        TipoLote.Abertura,
                        config.Periodo,
                        quantidadeEventos,
                        config.CnpjDeclarante,
                        arquivoXml,
                        null, // Ainda não assinado
                        null, // Ainda não criptografado
                        ambienteStr
                    );
                    persistenceService.RegistrarLogLote(idLoteBanco, "GERACAO", $"XML gerado: {Path.GetFileName(arquivoXml)}");
                    AdicionarLog($"Lote registrado no banco (ID: {idLoteBanco}).");
                }
                catch (Exception exDb)
                {
                    string erroCompleto = $"⚠ ERRO ao registrar lote no banco: {exDb.Message}";
                    if (exDb.InnerException != null)
                    {
                        erroCompleto += $"\nDetalhes: {exDb.InnerException.Message}";
                    }
                    AdicionarLog(erroCompleto);
                    MessageBox.Show($"Erro ao registrar lote no banco de dados:\n\n{erroCompleto}", 
                        "Erro de Banco de Dados", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                // 2. Assinar
                AtualizarEtapa("Assinando XML...");
                var assinaturaService = new EfinanceiraAssinaturaService();
                X509Certificate2 cert = BuscarCertificado(config.CertThumbprint);
                var xmlAssinado = assinaturaService.AssinarEventosDoArquivo(arquivoXml, cert);
                string arquivoAssinado = arquivoXml.Replace(".xml", "-ASSINADO.xml");
                xmlAssinado.Save(arquivoAssinado);
                status.LotesAssinados = 1;
                AtualizarEstatisticas();
                AdicionarLog($"XML assinado: {arquivoAssinado}");

                // Atualizar lote no banco após assinar
                if (idLoteBanco > 0)
                {
                    try
                    {
                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                        persistenceService.AtualizarLote(idLoteBanco, "ASSINADO");
                        persistenceService.RegistrarLogLote(idLoteBanco, "ASSINATURA", $"XML assinado: {Path.GetFileName(arquivoAssinado)}");
                    }
                    catch (Exception exDb)
                    {
                        AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após assinatura: {exDb.Message}");
                    }
                }

                if (cancelarProcessamento) return;

                // 3. Criptografar
                AtualizarEtapa("Criptografando XML...");
                var criptografiaService = new EfinanceiraCriptografiaService();
                string arquivoCriptografado = criptografiaService.CriptografarLote(arquivoAssinado, config.CertServidorThumbprint);
                status.LotesCriptografados = 1;
                AtualizarEstatisticas();
                AdicionarLog($"XML criptografado: {arquivoCriptografado}");

                // Atualizar lote no banco após criptografar
                if (idLoteBanco > 0)
                {
                    try
                    {
                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                        persistenceService.AtualizarLote(idLoteBanco, "CRIPTOGRAFADO");
                        persistenceService.RegistrarLogLote(idLoteBanco, "CRIPTOGRAFIA", $"XML criptografado: {Path.GetFileName(arquivoCriptografado)}");
                    }
                    catch (Exception exDb)
                    {
                        AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após criptografia: {exDb.Message}");
                    }
                }

                if (cancelarProcessamento) return;

                // Registrar lote processado (mesmo sem envio) - manter compatibilidade com sistema antigo
                ProtocoloPersistenciaService.RegistrarProtocolo(
                    TipoLote.Abertura,
                    arquivoCriptografado,
                    null, // Protocolo será preenchido após envio
                    config.Periodo,
                    quantidadeEventos
                );
                AdicionarLog($"Lote registrado na lista de processados ({quantidadeEventos} evento(s)).");
                
                // Atualizar lista na aba Consulta
                if (ConsultaForm != null)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        ConsultaForm.AtualizarListaLotes();
                    });
                }

                // 4. Enviar (se não estiver marcado "Apenas Processar")
                if (!chkApenasProcessar.Checked)
                {
                    try
                    {
                        AtualizarEtapa("Enviando para e-Financeira...");
                        var envioService = new EfinanceiraEnvioService();
                        var resposta = envioService.EnviarLote(arquivoCriptografado, config, cert);
                        status.LotesEnviados = 1;
                        AtualizarEstatisticas();
                        
                        if (resposta.CodigoResposta == 1)
                        {
                            // Garantir que o protocolo foi extraído (pode não ter vindo no objeto resposta)
                            string protocoloFinal = resposta.Protocolo;
                            
                            // Se não veio no objeto resposta, tentar extrair do XML
                            if (string.IsNullOrEmpty(protocoloFinal) && !string.IsNullOrEmpty(resposta.XmlCompleto))
                            {
                                try
                                {
                                    System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                    doc.LoadXml(resposta.XmlCompleto);
                                    
                                    // Primeiro tentar com GetElementsByTagName (sem namespace, como no Java)
                                    System.Xml.XmlNodeList protocoloList = doc.GetElementsByTagName("protocoloEnvio");
                                    if (protocoloList != null && protocoloList.Count > 0)
                                    {
                                        protocoloFinal = protocoloList[0].InnerText.Trim();
                                    }
                                    else
                                    {
                                        // Tentar com XPath (com namespace)
                                        System.Xml.XmlNamespaceManager nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                                        nsmgr.AddNamespace("ns", "http://www.eFinanceira.gov.br/schemas/envioLoteCriptografado/v1_2_0");
                                        
                                        System.Xml.XmlNode protocoloNode = doc.SelectSingleNode("//protocoloEnvio")
                                            ?? doc.SelectSingleNode("//ns:protocoloEnvio", nsmgr)
                                            ?? doc.SelectSingleNode("//protocolo")
                                            ?? doc.SelectSingleNode("//ns:protocolo", nsmgr);
                                        
                                        if (protocoloNode != null)
                                        {
                                            protocoloFinal = protocoloNode.InnerText.Trim();
                                        }
                                        else
                                        {
                                            // Tentar buscar por numeroProtocolo (como no método extrairProtocolo do Java)
                                            System.Xml.XmlNodeList numeroProtocoloList = doc.GetElementsByTagName("numeroProtocolo");
                                            if (numeroProtocoloList != null && numeroProtocoloList.Count > 0)
                                            {
                                                protocoloFinal = numeroProtocoloList[0].InnerText.Trim();
                                            }
                                        }
                                    }
                                    
                                    if (!string.IsNullOrEmpty(protocoloFinal))
                                    {
                                        resposta.Protocolo = protocoloFinal; // Atualizar no objeto resposta também
                                        AdicionarLog($"✓ Protocolo extraído do XML: {protocoloFinal}");
                                    }
                                }
                                catch (Exception exXml)
                                {
                                    AdicionarLog($"⚠ Erro ao extrair protocolo do XML: {exXml.Message}");
                                }
                            }
                            
                            // AGUARDAR PROTOCOLO - Não finalizar sem protocolo
                            if (!string.IsNullOrEmpty(protocoloFinal))
                            {
                                // Adicionar protocolo à lista
                                if (!status.ProtocolosEnviados.Contains(protocoloFinal))
                                {
                                    status.ProtocolosEnviados.Add(protocoloFinal);
                                }
                                
                                AdicionarLog($"✓ Lote de abertura enviado com sucesso! Protocolo: {protocoloFinal}");
                                AdicionarLog($"════════════════════════════════════════");
                                AdicionarLog($"PROTOCOLO DO LOTE DE ABERTURA:");
                                AdicionarLog($"{protocoloFinal}");
                                AdicionarLog($"════════════════════════════════════════");
                                
                                // Atualizar lote no banco após envio
                                if (idLoteBanco > 0)
                                {
                                    try
                                    {
                                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                                        string xmlResposta = resposta.XmlCompleto ?? "";
                                        persistenceService.AtualizarLote(
                                            idLoteBanco,
                                            "ENVIADO",
                                            protocoloFinal,
                                            resposta.CodigoResposta,
                                            resposta.Descricao,
                                            xmlResposta,
                                            null,
                                            null,
                                            null,
                                            DateTime.Now,
                                            null,
                                            null
                                        );
                                        persistenceService.RegistrarLogLote(idLoteBanco, "ENVIO", $"Lote enviado com sucesso. Protocolo: {protocoloFinal}");
                                    }
                                    catch (Exception exDb)
                                    {
                                        AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após envio: {exDb.Message}");
                                    }
                                }
                                
                                // Atualizar protocolo no lote já registrado (sistema antigo)
                                ProtocoloPersistenciaService.RegistrarProtocolo(
                                    TipoLote.Abertura, 
                                    arquivoCriptografado, 
                                    protocoloFinal,
                                    config.Periodo,
                                    quantidadeEventos
                                );
                                
                                // Aguardar um momento para garantir que o protocolo foi salvo
                                System.Threading.Thread.Sleep(500);
                                
                                // Atualizar lista na aba Consulta com o protocolo já preenchido
                                if (ConsultaForm != null)
                                {
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        ConsultaForm.AtualizarListaLotes();
                                    });
                                }
                                
                                // Exibir MessageBox destacando o protocolo
                                MessageBox.Show(
                                    $"Lote de abertura enviado com sucesso!\n\n" +
                                    $"PROTOCOLO: {protocoloFinal}\n\n" +
                                    $"Este protocolo foi salvo e pode ser consultado na aba 'Consulta'.",
                                    "Envio Concluído",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );
                                
                                AtualizarEstatisticas();
                            }
                            else
                            {
                                // Se não recebeu protocolo, aguardar e tentar novamente ou informar erro
                                AdicionarLog($"⚠ ATENÇÃO: Lote enviado mas protocolo não foi retornado!");
                                AdicionarLog($"  Aguardando resposta do servidor...");
                                
                                // Tentar aguardar um pouco mais e verificar se há protocolo na resposta XML
                                System.Threading.Thread.Sleep(2000);
                                
                                // Se ainda não tiver protocolo, registrar como erro
                                if (idLoteBanco > 0)
                                {
                                    try
                                    {
                                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                                        persistenceService.AtualizarLote(
                                            idLoteBanco,
                                            "ENVIADO_SEM_PROTOCOLO",
                                            null,
                                            resposta.CodigoResposta,
                                            resposta.Descricao + " (Protocolo não retornado)",
                                            resposta.XmlCompleto ?? "",
                                            null,
                                            null,
                                            null,
                                            DateTime.Now,
                                            null,
                                            "Protocolo não retornado pelo servidor"
                                        );
                                    }
                                    catch (Exception exDb)
                                    {
                                        AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                                    }
                                }
                                
                                MessageBox.Show(
                                    $"Lote enviado, mas o protocolo não foi retornado pelo servidor.\n\n" +
                                    $"Código de Resposta: {resposta.CodigoResposta}\n" +
                                    $"Descrição: {resposta.Descricao}\n\n" +
                                    $"Verifique o XML de resposta ou consulte o lote mais tarde.",
                                    "Aviso - Protocolo Não Recebido",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning
                                );
                            }
                        }
                        else if (resposta.CodigoResposta == 7)
                        {
                            AdicionarLog($"✗ Lote de abertura REJEITADO - Código: {resposta.CodigoResposta}");
                            AdicionarLog($"  Descrição: {resposta.Descricao}");
                            
                            // Atualizar lote no banco com erro
                            if (idLoteBanco > 0)
                            {
                                try
                                {
                                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                                    string erroMsg = $"REJEITADO - Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}";
                                    persistenceService.AtualizarLote(
                                        idLoteBanco,
                                        "REJEITADO",
                                        null,
                                        resposta.CodigoResposta,
                                        resposta.Descricao,
                                        resposta.XmlCompleto ?? "",
                                        null,
                                        null,
                                        null,
                                        DateTime.Now,
                                        null,
                                        erroMsg
                                    );
                                    persistenceService.RegistrarLogLote(idLoteBanco, "ENVIO_ERRO", erroMsg);
                                }
                                catch (Exception exDb)
                                {
                                    AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após rejeição: {exDb.Message}");
                                }
                            }
                            
                            if (resposta.Ocorrencias != null && resposta.Ocorrencias.Count > 0)
                            {
                                foreach (var ocorr in resposta.Ocorrencias)
                                {
                                    AdicionarLog($"  Ocorrência: {ocorr.Codigo} - {ocorr.Descricao} ({ocorr.Tipo})");
                                }
                            }
                        }
                        else
                        {
                            // SEMPRE tentar extrair e salvar o protocolo, mesmo se código não for 1
                            string protocoloExtraido = resposta.Protocolo;
                            
                            // Se não veio no objeto resposta, tentar extrair do XML (seguindo lógica do Java)
                            if (string.IsNullOrEmpty(protocoloExtraido) && !string.IsNullOrEmpty(resposta.XmlCompleto))
                            {
                                try
                                {
                                    System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                    doc.LoadXml(resposta.XmlCompleto);
                                    
                                    // Primeiro tentar com GetElementsByTagName (sem namespace, como no Java)
                                    System.Xml.XmlNodeList protocoloList = doc.GetElementsByTagName("protocoloEnvio");
                                    if (protocoloList != null && protocoloList.Count > 0)
                                    {
                                        protocoloExtraido = protocoloList[0].InnerText.Trim();
                                    }
                                    else
                                    {
                                        // Tentar com XPath (com namespace)
                                        System.Xml.XmlNamespaceManager nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                                        nsmgr.AddNamespace("ns", "http://www.eFinanceira.gov.br/schemas/envioLoteCriptografado/v1_2_0");
                                        
                                        System.Xml.XmlNode protocoloNode = doc.SelectSingleNode("//protocoloEnvio")
                                            ?? doc.SelectSingleNode("//ns:protocoloEnvio", nsmgr)
                                            ?? doc.SelectSingleNode("//protocolo")
                                            ?? doc.SelectSingleNode("//ns:protocolo", nsmgr);
                                        
                                        if (protocoloNode != null)
                                        {
                                            protocoloExtraido = protocoloNode.InnerText.Trim();
                                        }
                                        else
                                        {
                                            // Tentar buscar por numeroProtocolo (como no método extrairProtocolo do Java)
                                            System.Xml.XmlNodeList numeroProtocoloList = doc.GetElementsByTagName("numeroProtocolo");
                                            if (numeroProtocoloList != null && numeroProtocoloList.Count > 0)
                                            {
                                                protocoloExtraido = numeroProtocoloList[0].InnerText.Trim();
                                            }
                                        }
                                    }
                                    
                                    if (!string.IsNullOrEmpty(protocoloExtraido))
                                    {
                                        AdicionarLog($"✓ Protocolo extraído do XML: {protocoloExtraido}");
                                    }
                                }
                                catch (Exception exXml)
                                {
                                    AdicionarLog($"⚠ Erro ao extrair protocolo do XML: {exXml.Message}");
                                    System.Diagnostics.Debug.WriteLine($"Erro ao extrair protocolo do XML: {exXml.Message}");
                                }
                            }
                            
                            if (!string.IsNullOrEmpty(protocoloExtraido))
                            {
                                AdicionarLog($"✓ Protocolo extraído: {protocoloExtraido}");
                                
                                // Adicionar protocolo à lista
                                if (!status.ProtocolosEnviados.Contains(protocoloExtraido))
                                {
                                    status.ProtocolosEnviados.Add(protocoloExtraido);
                                }
                                
                                // Atualizar protocolo no lote já registrado (sistema antigo)
                                ProtocoloPersistenciaService.RegistrarProtocolo(
                                    TipoLote.Abertura, 
                                    arquivoCriptografado, 
                                    protocoloExtraido,
                                    config.Periodo,
                                    quantidadeEventos
                                );
                                
                                // Aguardar um momento para garantir que o protocolo foi salvo
                                System.Threading.Thread.Sleep(500);
                                
                                // Atualizar lista na aba Consulta
                                if (ConsultaForm != null)
                                {
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        ConsultaForm.AtualizarListaLotes();
                                    });
                                }
                            }
                            
                            AdicionarLog($"⚠ Lote de abertura - Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}");
                            
                            // Atualizar lote no banco com resposta e protocolo (se encontrado)
                            if (idLoteBanco > 0)
                            {
                                try
                                {
                                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                                    persistenceService.AtualizarLote(
                                        idLoteBanco,
                                        "ENVIADO_COM_RESPOSTA",
                                        protocoloExtraido, // Salvar protocolo mesmo se código não for 1
                                        resposta.CodigoResposta,
                                        resposta.Descricao,
                                        resposta.XmlCompleto ?? "",
                                        null,
                                        null,
                                        null,
                                        DateTime.Now,
                                        null,
                                        null
                                    );
                                    
                                    if (!string.IsNullOrEmpty(protocoloExtraido))
                                    {
                                        persistenceService.RegistrarLogLote(idLoteBanco, "ENVIO", 
                                            $"Lote enviado. Protocolo: {protocoloExtraido}, Código: {resposta.CodigoResposta}");
                                    }
                                }
                                catch (Exception exDb)
                                {
                                    AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                                }
                            }
                        }
                    }
                    catch (Exception exEnv)
                    {
                        AdicionarLog($"✗ Erro ao enviar lote de abertura: {exEnv.Message}");
                    }
                }
                else
                {
                    AdicionarLog("Envio não realizado (modo 'Apenas Processar' ativo).");
                    
                    // Atualizar lista na aba Consulta mesmo sem protocolo
                    if (ConsultaForm != null)
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            ConsultaForm.AtualizarListaLotes();
                        });
                    }
                }

                status.LotesProcessados = 1;
                AtualizarEstatisticas();
                AtualizarEtapa("Processamento concluído com sucesso!");
                AdicionarLog("Processamento de abertura concluído.");

                this.Invoke((MethodInvoker)delegate
                {
                    btnProcessarAbertura.Enabled = true;
                    btnProcessarMovimentacao.Enabled = true;
                    btnProcessarFechamento.Enabled = true;
                    btnCancelar.Enabled = false;
                });
            }
            catch (Exception ex)
            {
                AdicionarLog($"ERRO: {ex.Message}");
                status.LotesComErro = 1;
                this.Invoke((MethodInvoker)delegate
                {
                    btnProcessarAbertura.Enabled = true;
                    btnProcessarMovimentacao.Enabled = true;
                    btnProcessarFechamento.Enabled = true;
                    btnCancelar.Enabled = false;
                    MessageBox.Show($"Erro ao processar: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }

        private async Task ProcessarMovimentacao()
        {
            try
            {
                AtualizarEtapa("Iniciando processamento de movimentação...");
                status.InicioProcessamento = DateTime.Now;
                status.TotalLotes = 0;
                status.LotesProcessados = 0;
                status.LotesAssinados = 0;
                status.LotesCriptografados = 0;
                status.LotesEnviados = 0;
                status.LotesComErro = 0;
                status.ProtocolosEnviados.Clear();

                var config = ConfigForm.Config;
                var dadosAbertura = ConfigForm.DadosAbertura;

                // Validar e calcular período
                string periodoStr = config.Periodo;
                if (string.IsNullOrWhiteSpace(periodoStr))
                {
                    throw new Exception("Período não configurado. Configure o campo 'Período' na seção Configuração Geral.");
                }

                // Validar formato do período
                if (!EfinanceiraPeriodoUtil.ValidarPeriodo(periodoStr))
                {
                    throw new Exception($"Período inválido: {periodoStr}. Deve estar no formato YYYYMM onde MM deve ser:\n" +
                        $"  • 01 ou 06 = Primeiro semestre (Janeiro a Junho)\n" +
                        $"  • 02 ou 12 = Segundo semestre (Julho a Dezembro)\n" +
                        $"Exemplos: 202301 (Jan-Jun/2023) ou 202302 (Jul-Dez/2023)");
                }

                // Calcular datas do período semestral
                var (dataInicio, dataFim) = EfinanceiraPeriodoUtil.CalcularPeriodoSemestral(periodoStr);
                
                AdicionarLog($"Período configurado: {periodoStr}");
                AdicionarLog($"Datas calculadas pelo EfinanceiraPeriodoUtil: {dataInicio} a {dataFim}");
                
                DateTime dtInicio = DateTime.Parse(dataInicio);
                DateTime dtFim = DateTime.Parse(dataFim);
                
                AdicionarLog($"Datas parseadas: dtInicio={dtInicio:yyyy-MM-dd}, dtFim={dtFim:yyyy-MM-dd}");
                
                int ano = dtInicio.Year;
                int mesInicial = dtInicio.Month;
                int mesFinal = dtFim.Month;

                AdicionarLog($"Ano extraído: {ano}");
                AdicionarLog($"Mês Inicial extraído de dtInicio.Month: {mesInicial}");
                AdicionarLog($"Mês Final extraído de dtFim.Month: {mesFinal}");

                // Validação: garantir que os meses estão corretos
                if (mesInicial < 1 || mesInicial > 12 || mesFinal < 1 || mesFinal > 12)
                {
                    throw new Exception($"Meses inválidos calculados: Mês Inicial={mesInicial}, Mês Final={mesFinal}. Verifique o período configurado.");
                }

                // Para períodos semestrais, validar:
                // - Se mesInicial = 1, mesFinal deve ser 6 (Jan-Jun)
                // - Se mesInicial = 7, mesFinal deve ser 12 (Jul-Dez)
                if ((mesInicial == 1 && mesFinal != 6) || (mesInicial == 7 && mesFinal != 12))
                {
                    throw new Exception($"Período semestral inválido: Mês Inicial={mesInicial}, Mês Final={mesFinal}. " +
                        $"Esperado: 1-6 (Jan-Jun) ou 7-12 (Jul-Dez). Verifique o período '{periodoStr}' configurado.");
                }

                // Garantir que mesInicial <= mesFinal (para a query SQL funcionar corretamente)
                if (mesInicial > mesFinal)
                {
                    throw new Exception($"Erro: Mês Inicial ({mesInicial}) é maior que Mês Final ({mesFinal}). " +
                        $"Isso não é válido para um período semestral. Período configurado: {periodoStr}");
                }

                // Testar conexão com banco
                AtualizarEtapa("Testando conexão com banco de dados...");
                var dbService = new EfinanceiraDatabaseService();
                if (!dbService.TestarConexao())
                {
                    throw new Exception("Não foi possível conectar ao banco de dados. Verifique as credenciais.");
                }
                AdicionarLog("Conexão com banco de dados estabelecida.");

                // Buscar pessoas com contas (paginado)
                AtualizarEtapa("Buscando dados do banco de dados...");
                int pageSize = config.PageSize;
                int offset = config.OffsetRegistros;
                int maxLotes = config.MaxLotes ?? int.MaxValue;
                int lotesGerados = 0;
                int eventosOffset = config.EventoOffset - 1; // Ajustar para índice base 0

                while (lotesGerados < maxLotes && !cancelarProcessamento)
                {
                    AdicionarLog($"Buscando página: offset={offset}, limit={pageSize}...");
                    var pessoas = dbService.BuscarPessoasComContas(ano, mesInicial, mesFinal, pageSize, offset);
                    
                    if (pessoas.Count == 0)
                    {
                        AdicionarLog("Não há mais registros para processar.");
                        break;
                    }

                    AdicionarLog($"Encontradas {pessoas.Count} pessoas nesta página.");

                    // Validar eventosOffset antes de processar
                    if (eventosOffset < 0)
                    {
                        AdicionarLog($"Aviso: EventoOffset ajustado de {eventosOffset} para 0.");
                        eventosOffset = 0;
                    }
                    if (eventosOffset >= pessoas.Count)
                    {
                        AdicionarLog($"Aviso: EventoOffset ({eventosOffset}) é maior ou igual ao número de pessoas ({pessoas.Count}). " +
                            $"Pulando esta página e incrementando offset.");
                        offset += pageSize;
                        continue;
                    }

                    // Gerar lotes desta página (usar valor configurável de eventos por lote)
                    int eventosPorLote = config.EventosPorLote > 0 && config.EventosPorLote <= 50 
                        ? config.EventosPorLote : 50; // Garantir que está entre 1 e 50
                    
                    AdicionarLog($"Gerando lotes com {eventosPorLote} evento(s) por lote (configurado: {config.EventosPorLote})");
                    
                    for (int i = eventosOffset; i < pessoas.Count && lotesGerados < maxLotes; i += eventosPorLote)
                    {
                        if (cancelarProcessamento) break;

                        // Validar índice antes de usar Skip
                        if (i >= pessoas.Count)
                        {
                            AdicionarLog($"Aviso: Índice {i} está fora do intervalo (0-{pessoas.Count - 1}). Pulando.");
                            break;
                        }

                        var pessoasLote = pessoas.Skip(i).Take(eventosPorLote).ToList();
                        if (pessoasLote.Count == 0) break;

                        lotesGerados++;
                        status.TotalLotes = lotesGerados;

                        AtualizarEtapa($"Gerando lote {lotesGerados} ({pessoasLote.Count} eventos)...");
                        
                        // Gerar XML
                        var geradorService = new EfinanceiraGeradorXmlService();
                        string arquivoXml = geradorService.GerarXmlMovimentacao(
                            pessoasLote, 
                            config.CnpjDeclarante, 
                            periodoStr, 
                            config.Ambiente == EfinanceiraAmbiente.PROD ? 1 : 2,
                            eventosOffset,
                            config.DiretorioLotes
                        );
                        AdicionarLog($"Lote {lotesGerados} gerado: {Path.GetFileName(arquivoXml)}");

                        if (cancelarProcessamento) break;

                        // Registrar lote no banco após gerar XML
                        int quantidadeEventos = pessoasLote.Count;
                        long idLoteBanco = 0;
                        try
                        {
                            var persistenceService = new EfinanceiraDatabasePersistenceService();
                            string ambienteStr = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : "HOMOLOG";
                            idLoteBanco = persistenceService.RegistrarLote(
                                TipoLote.Movimentacao,
                                periodoStr,
                                quantidadeEventos,
                                config.CnpjDeclarante,
                                arquivoXml,
                                null, // Ainda não assinado
                                null, // Ainda não criptografado
                                ambienteStr
                            );
                            persistenceService.RegistrarLogLote(idLoteBanco, "GERACAO", $"XML gerado: {Path.GetFileName(arquivoXml)}");
                            
                            // Registrar eventos do lote
                            persistenceService.RegistrarEventosDoLote(idLoteBanco, pessoasLote);
                            AdicionarLog($"Lote {lotesGerados} registrado no banco (ID: {idLoteBanco}).");
                        }
                        catch (Exception exDb)
                        {
                            AdicionarLog($"⚠ Aviso: Erro ao registrar lote no banco: {exDb.Message}");
                        }

                        // Assinar
                        AtualizarEtapa($"Assinando lote {lotesGerados}...");
                        var assinaturaService = new EfinanceiraAssinaturaService();
                        X509Certificate2 cert = BuscarCertificado(config.CertThumbprint);
                        var xmlAssinado = assinaturaService.AssinarEventosDoArquivo(arquivoXml, cert);
                        string arquivoAssinado = arquivoXml.Replace(".xml", "-ASSINADO.xml");
                        xmlAssinado.Save(arquivoAssinado);
                        status.LotesAssinados++;
                        AtualizarEstatisticas();
                        AdicionarLog($"Lote {lotesGerados} assinado.");

                        // Atualizar lote no banco após assinar
                        if (idLoteBanco > 0)
                        {
                            try
                            {
                                var persistenceService = new EfinanceiraDatabasePersistenceService();
                                persistenceService.AtualizarLote(idLoteBanco, "ASSINADO");
                                persistenceService.RegistrarLogLote(idLoteBanco, "ASSINATURA", $"XML assinado: {Path.GetFileName(arquivoAssinado)}");
                            }
                            catch (Exception exDb)
                            {
                                AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após assinatura: {exDb.Message}");
                            }
                        }

                        if (cancelarProcessamento) break;

                        // Criptografar
                        AtualizarEtapa($"Criptografando lote {lotesGerados}...");
                        var criptografiaService = new EfinanceiraCriptografiaService();
                        string arquivoCriptografado = criptografiaService.CriptografarLote(arquivoAssinado, config.CertServidorThumbprint);
                        status.LotesCriptografados++;
                        AtualizarEstatisticas();
                        AdicionarLog($"Lote {lotesGerados} criptografado.");

                        // Atualizar lote no banco após criptografar
                        if (idLoteBanco > 0)
                        {
                            try
                            {
                                var persistenceService = new EfinanceiraDatabasePersistenceService();
                                persistenceService.AtualizarLote(idLoteBanco, "CRIPTOGRAFADO");
                                persistenceService.RegistrarLogLote(idLoteBanco, "CRIPTOGRAFIA", $"XML criptografado: {Path.GetFileName(arquivoCriptografado)}");
                            }
                            catch (Exception exDb)
                            {
                                AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após criptografia: {exDb.Message}");
                            }
                        }

                        if (cancelarProcessamento) break;

                        // Registrar lote processado (mesmo sem envio) - manter compatibilidade com sistema antigo
                        ProtocoloPersistenciaService.RegistrarProtocolo(
                            TipoLote.Movimentacao,
                            arquivoCriptografado,
                            null, // Protocolo será preenchido após envio
                            periodoStr,
                            quantidadeEventos
                        );
                        AdicionarLog($"Lote {lotesGerados} registrado na lista de processados ({quantidadeEventos} evento(s)).");

                        // Enviar (se não estiver marcado "Apenas Processar")
                        if (!chkApenasProcessar.Checked)
                        {
                            try
                            {
                                AtualizarEtapa($"Enviando lote {lotesGerados}...");
                                var envioService = new EfinanceiraEnvioService();
                                var resposta = envioService.EnviarLote(arquivoCriptografado, config, cert);
                                status.LotesEnviados++;
                                AtualizarEstatisticas();
                                
                                if (resposta.CodigoResposta == 1)
                                {
                                    // Garantir que o protocolo foi extraído (pode não ter vindo no objeto resposta)
                                    string protocoloFinal = resposta.Protocolo;
                                    
                                    // Se não veio no objeto resposta, tentar extrair do XML
                                    if (string.IsNullOrEmpty(protocoloFinal) && !string.IsNullOrEmpty(resposta.XmlCompleto))
                                    {
                                        try
                                        {
                                            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                            doc.LoadXml(resposta.XmlCompleto);
                                            
                                            // Primeiro tentar com GetElementsByTagName (sem namespace, como no Java)
                                            System.Xml.XmlNodeList protocoloList = doc.GetElementsByTagName("protocoloEnvio");
                                            if (protocoloList != null && protocoloList.Count > 0)
                                            {
                                                protocoloFinal = protocoloList[0].InnerText.Trim();
                                            }
                                            else
                                            {
                                                // Tentar com XPath (com namespace)
                                                System.Xml.XmlNamespaceManager nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                                                nsmgr.AddNamespace("ns", "http://www.eFinanceira.gov.br/schemas/envioLoteCriptografado/v1_2_0");
                                                
                                                System.Xml.XmlNode protocoloNode = doc.SelectSingleNode("//protocoloEnvio")
                                                    ?? doc.SelectSingleNode("//ns:protocoloEnvio", nsmgr)
                                                    ?? doc.SelectSingleNode("//protocolo")
                                                    ?? doc.SelectSingleNode("//ns:protocolo", nsmgr);
                                                
                                                if (protocoloNode != null)
                                                {
                                                    protocoloFinal = protocoloNode.InnerText.Trim();
                                                }
                                                else
                                                {
                                                    // Tentar buscar por numeroProtocolo (como no método extrairProtocolo do Java)
                                                    System.Xml.XmlNodeList numeroProtocoloList = doc.GetElementsByTagName("numeroProtocolo");
                                                    if (numeroProtocoloList != null && numeroProtocoloList.Count > 0)
                                                    {
                                                        protocoloFinal = numeroProtocoloList[0].InnerText.Trim();
                                                    }
                                                }
                                            }
                                            
                                            if (!string.IsNullOrEmpty(protocoloFinal))
                                            {
                                                resposta.Protocolo = protocoloFinal; // Atualizar no objeto resposta também
                                                AdicionarLog($"✓ Protocolo extraído do XML: {protocoloFinal}");
                                            }
                                        }
                                        catch (Exception exXml)
                                        {
                                            AdicionarLog($"⚠ Erro ao extrair protocolo do XML: {exXml.Message}");
                                        }
                                    }
                                    
                                    // AGUARDAR PROTOCOLO - Não finalizar sem protocolo
                                    if (!string.IsNullOrEmpty(protocoloFinal))
                                    {
                                        // Adicionar protocolo à lista
                                        if (!status.ProtocolosEnviados.Contains(protocoloFinal))
                                        {
                                            status.ProtocolosEnviados.Add(protocoloFinal);
                                        }
                                        
                                        AdicionarLog($"✓ Lote {lotesGerados} enviado com sucesso! Protocolo: {protocoloFinal}");
                                        AdicionarLog($"════════════════════════════════════════");
                                        AdicionarLog($"PROTOCOLO DO LOTE {lotesGerados} DE MOVIMENTAÇÃO:");
                                        AdicionarLog($"{protocoloFinal}");
                                        AdicionarLog($"════════════════════════════════════════");
                                        
                                        // Atualizar lote no banco após envio
                                        if (idLoteBanco > 0)
                                        {
                                            try
                                            {
                                                var persistenceService = new EfinanceiraDatabasePersistenceService();
                                                string xmlResposta = resposta.XmlCompleto ?? "";
                                                persistenceService.AtualizarLote(
                                                    idLoteBanco,
                                                    "ENVIADO",
                                                    protocoloFinal,
                                                    resposta.CodigoResposta,
                                                    resposta.Descricao,
                                                    xmlResposta,
                                                    null,
                                                    null,
                                                    null,
                                                    DateTime.Now,
                                                    null,
                                                    null
                                                );
                                                persistenceService.RegistrarLogLote(idLoteBanco, "ENVIO", $"Lote enviado com sucesso. Protocolo: {protocoloFinal}");
                                            }
                                            catch (Exception exDb)
                                            {
                                                AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após envio: {exDb.Message}");
                                            }
                                        }
                                        
                                        // Atualizar protocolo no lote já registrado (sistema antigo)
                                        ProtocoloPersistenciaService.RegistrarProtocolo(
                                            TipoLote.Movimentacao, 
                                            arquivoCriptografado, 
                                            protocoloFinal,
                                            periodoStr,
                                            quantidadeEventos
                                        );
                                        
                                        // Aguardar um momento para garantir que o protocolo foi salvo
                                        System.Threading.Thread.Sleep(500);
                                        
                                        // Atualizar lista na aba Consulta com o protocolo já preenchido
                                        if (ConsultaForm != null)
                                        {
                                            this.Invoke((MethodInvoker)delegate
                                            {
                                                ConsultaForm.AtualizarListaLotes();
                                            });
                                        }
                                        
                                        // Exibir MessageBox destacando o protocolo
                                        MessageBox.Show(
                                            $"Lote {lotesGerados} de movimentação enviado com sucesso!\n\n" +
                                            $"PROTOCOLO: {protocoloFinal}\n\n" +
                                            $"Este protocolo foi salvo e pode ser consultado na aba 'Consulta'.",
                                            "Envio Concluído",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Information
                                        );
                                        
                                        AtualizarEstatisticas();
                                    }
                                    else
                                    {
                                        // Se não recebeu protocolo, aguardar e tentar novamente ou informar erro
                                        AdicionarLog($"⚠ ATENÇÃO: Lote {lotesGerados} enviado mas protocolo não foi retornado!");
                                        AdicionarLog($"  Aguardando resposta do servidor...");
                                        
                                        // Tentar aguardar um pouco mais e verificar se há protocolo na resposta XML
                                        System.Threading.Thread.Sleep(2000);
                                        
                                        // Se ainda não tiver protocolo, registrar como erro
                                        if (idLoteBanco > 0)
                                        {
                                            try
                                            {
                                                var persistenceService = new EfinanceiraDatabasePersistenceService();
                                                persistenceService.AtualizarLote(
                                                    idLoteBanco,
                                                    "ENVIADO_SEM_PROTOCOLO",
                                                    null,
                                                    resposta.CodigoResposta,
                                                    resposta.Descricao + " (Protocolo não retornado)",
                                                    resposta.XmlCompleto ?? "",
                                                    null,
                                                    null,
                                                    null,
                                                    DateTime.Now,
                                                    null,
                                                    "Protocolo não retornado pelo servidor"
                                                );
                                            }
                                            catch (Exception exDb)
                                            {
                                                AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                                            }
                                        }
                                        
                                        MessageBox.Show(
                                            $"Lote {lotesGerados} enviado, mas o protocolo não foi retornado pelo servidor.\n\n" +
                                            $"Código de Resposta: {resposta.CodigoResposta}\n" +
                                            $"Descrição: {resposta.Descricao}\n\n" +
                                            $"Verifique o XML de resposta ou consulte o lote mais tarde.",
                                            "Aviso - Protocolo Não Recebido",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Warning
                                        );
                                    }
                                }
                                else if (resposta.CodigoResposta == 7)
                                {
                                    AdicionarLog($"✗ Lote {lotesGerados} REJEITADO - Código: {resposta.CodigoResposta}");
                                    AdicionarLog($"  Descrição: {resposta.Descricao}");
                                    
                                    // Atualizar lote no banco com erro
                                    if (idLoteBanco > 0)
                                    {
                                        try
                                        {
                                            var persistenceService = new EfinanceiraDatabasePersistenceService();
                                            string erroMsg = $"REJEITADO - Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}";
                                            persistenceService.AtualizarLote(
                                                idLoteBanco,
                                                "REJEITADO",
                                                null,
                                                resposta.CodigoResposta,
                                                resposta.Descricao,
                                                resposta.XmlCompleto ?? "",
                                                null,
                                                null,
                                                null,
                                                DateTime.Now,
                                                null,
                                                erroMsg
                                            );
                                            persistenceService.RegistrarLogLote(idLoteBanco, "ENVIO_ERRO", erroMsg);
                                        }
                                        catch (Exception exDb)
                                        {
                                            AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após rejeição: {exDb.Message}");
                                        }
                                    }
                                    
                                    if (resposta.Ocorrencias != null && resposta.Ocorrencias.Count > 0)
                                    {
                                        foreach (var ocorr in resposta.Ocorrencias)
                                        {
                                            AdicionarLog($"  Ocorrência: {ocorr.Codigo} - {ocorr.Descricao} ({ocorr.Tipo})");
                                        }
                                    }
                                }
                                else
                                {
                                    // SEMPRE tentar extrair e salvar o protocolo, mesmo se código não for 1
                                    string protocoloExtraido = resposta.Protocolo;
                                    
                                    // Se não veio no objeto resposta, tentar extrair do XML
                                    if (string.IsNullOrEmpty(protocoloExtraido) && !string.IsNullOrEmpty(resposta.XmlCompleto))
                                    {
                                        try
                                        {
                                            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                            doc.LoadXml(resposta.XmlCompleto);
                                            
                                            System.Xml.XmlNamespaceManager nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                                            nsmgr.AddNamespace("ns", "http://www.eFinanceira.gov.br/schemas/envioLoteCriptografado/v1_2_0");
                                            
                                            System.Xml.XmlNode protocoloNode = doc.SelectSingleNode("//protocoloEnvio")
                                                ?? doc.SelectSingleNode("//ns:protocoloEnvio", nsmgr)
                                                ?? doc.SelectSingleNode("//protocolo")
                                                ?? doc.SelectSingleNode("//ns:protocolo", nsmgr);
                                            
                                            if (protocoloNode != null)
                                            {
                                                protocoloExtraido = protocoloNode.InnerText.Trim();
                                            }
                                        }
                                        catch (Exception exXml)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"Erro ao extrair protocolo do XML: {exXml.Message}");
                                        }
                                    }
                                    
                                    if (!string.IsNullOrEmpty(protocoloExtraido))
                                    {
                                        AdicionarLog($"✓ Protocolo extraído: {protocoloExtraido}");
                                        
                                        // Adicionar protocolo à lista
                                        if (!status.ProtocolosEnviados.Contains(protocoloExtraido))
                                        {
                                            status.ProtocolosEnviados.Add(protocoloExtraido);
                                        }
                                        
                                        // Atualizar protocolo no lote já registrado (sistema antigo)
                                        ProtocoloPersistenciaService.RegistrarProtocolo(
                                            TipoLote.Movimentacao, 
                                            arquivoCriptografado, 
                                            protocoloExtraido,
                                            periodoStr,
                                            quantidadeEventos
                                        );
                                        
                                        // Aguardar um momento para garantir que o protocolo foi salvo
                                        System.Threading.Thread.Sleep(500);
                                        
                                        // Atualizar lista na aba Consulta
                                        if (ConsultaForm != null)
                                        {
                                            this.Invoke((MethodInvoker)delegate
                                            {
                                                ConsultaForm.AtualizarListaLotes();
                                            });
                                        }
                                    }
                                    
                                    AdicionarLog($"⚠ Lote {lotesGerados} - Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}");
                                    
                                    // Atualizar lote no banco com resposta e protocolo (se encontrado)
                                    if (idLoteBanco > 0)
                                    {
                                        try
                                        {
                                            var persistenceService = new EfinanceiraDatabasePersistenceService();
                                            persistenceService.AtualizarLote(
                                                idLoteBanco,
                                                "ENVIADO_COM_RESPOSTA",
                                                protocoloExtraido, // Salvar protocolo mesmo se código não for 1
                                                resposta.CodigoResposta,
                                                resposta.Descricao,
                                                resposta.XmlCompleto ?? "",
                                                null,
                                                null,
                                                null,
                                                DateTime.Now,
                                                null,
                                                null
                                            );
                                            
                                            if (!string.IsNullOrEmpty(protocoloExtraido))
                                            {
                                                persistenceService.RegistrarLogLote(idLoteBanco, "ENVIO", 
                                                    $"Lote enviado. Protocolo: {protocoloExtraido}, Código: {resposta.CodigoResposta}");
                                            }
                                        }
                                        catch (Exception exDb)
                                        {
                                            AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                                        }
                                    }
                                }
                            }
                            catch (Exception exEnv)
                            {
                                AdicionarLog($"✗ Erro ao enviar lote {lotesGerados}: {exEnv.Message}");
                            }
                        }
                        else
                        {
                            AdicionarLog($"Lote {lotesGerados} não enviado (modo 'Apenas Processar' ativo).");
                            
                            // Atualizar lista na aba Consulta mesmo sem protocolo
                            if (ConsultaForm != null)
                            {
                                this.Invoke((MethodInvoker)delegate
                                {
                                    ConsultaForm.AtualizarListaLotes();
                                });
                            }
                        }

                        status.LotesProcessados++;
                        AtualizarEstatisticas();
                        
                        // Resetar eventosOffset após o primeiro lote (só aplica no primeiro evento do primeiro lote)
                        if (lotesGerados == 1)
                        {
                            eventosOffset = 0;
                            AdicionarLog("EventoOffset resetado para 0 após o primeiro lote.");
                        }
                    }
                    
                    // Resetar eventosOffset para a próxima página (já foi aplicado no primeiro lote)
                    eventosOffset = 0;

                    offset += pageSize;
                    
                    // Se não encontrou mais registros, sair do loop
                    if (pessoas.Count < pageSize)
                    {
                        break;
                    }
                }

                AtualizarEtapa("Processamento de movimentação concluído!");
                AdicionarLog($"Processamento concluído. Total de lotes processados: {status.LotesProcessados}");

                this.Invoke((MethodInvoker)delegate
                {
                    btnProcessarAbertura.Enabled = true;
                    btnProcessarMovimentacao.Enabled = true;
                    btnProcessarFechamento.Enabled = true;
                    btnCancelar.Enabled = false;
                });
            }
            catch (Exception ex)
            {
                AdicionarLog($"ERRO: {ex.Message}");
                if (ex.InnerException != null)
                {
                    AdicionarLog($"Detalhes: {ex.InnerException.Message}");
                }
                status.LotesComErro++;
                this.Invoke((MethodInvoker)delegate
                {
                    btnProcessarAbertura.Enabled = true;
                    btnProcessarMovimentacao.Enabled = true;
                    btnProcessarFechamento.Enabled = true;
                    btnCancelar.Enabled = false;
                    MessageBox.Show($"Erro ao processar movimentação: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }

        private async Task ProcessarFechamento()
        {
            try
            {
                AtualizarEtapa("Iniciando processamento de fechamento...");
                status.InicioProcessamento = DateTime.Now;
                status.TotalLotes = 1;
                status.ProtocolosEnviados.Clear();

                var config = ConfigForm.Config;
                var dadosFechamento = ConfigForm.DadosFechamento;

                // 1. Gerar XML
                AtualizarEtapa("Gerando XML de fechamento...");
                var geradorService = new EfinanceiraGeradorXmlService();
                string arquivoXml = geradorService.GerarXmlFechamento(dadosFechamento, config.DiretorioLotes);
                AdicionarLog($"XML gerado: {arquivoXml}");

                if (cancelarProcessamento) return;

                // Registrar lote no banco após gerar XML
                int quantidadeEventos = ProtocoloPersistenciaService.ContarEventosNoXml(arquivoXml);
                long idLoteBanco = 0;
                try
                {
                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                    string ambienteStr = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : "HOMOLOG";
                    idLoteBanco = persistenceService.RegistrarLote(
                        TipoLote.Fechamento,
                        config.Periodo,
                        quantidadeEventos,
                        config.CnpjDeclarante,
                        arquivoXml,
                        null, // Ainda não assinado
                        null, // Ainda não criptografado
                        ambienteStr
                    );
                    persistenceService.RegistrarLogLote(idLoteBanco, "GERACAO", $"XML gerado: {Path.GetFileName(arquivoXml)}");
                    AdicionarLog($"Lote registrado no banco (ID: {idLoteBanco}).");
                }
                catch (Exception exDb)
                {
                    string erroCompleto = $"⚠ ERRO ao registrar lote no banco: {exDb.Message}";
                    if (exDb.InnerException != null)
                    {
                        erroCompleto += $"\nDetalhes: {exDb.InnerException.Message}";
                    }
                    AdicionarLog(erroCompleto);
                    MessageBox.Show($"Erro ao registrar lote no banco de dados:\n\n{erroCompleto}", 
                        "Erro de Banco de Dados", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                // 2. Assinar
                AtualizarEtapa("Assinando XML...");
                var assinaturaService = new EfinanceiraAssinaturaService();
                X509Certificate2 cert = BuscarCertificado(config.CertThumbprint);
                var xmlAssinado = assinaturaService.AssinarEventosDoArquivo(arquivoXml, cert);
                string arquivoAssinado = arquivoXml.Replace(".xml", "-ASSINADO.xml");
                xmlAssinado.Save(arquivoAssinado);
                status.LotesAssinados = 1;
                AtualizarEstatisticas();
                AdicionarLog($"XML assinado: {arquivoAssinado}");

                // Atualizar lote no banco após assinar
                if (idLoteBanco > 0)
                {
                    try
                    {
                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                        persistenceService.AtualizarLote(idLoteBanco, "ASSINADO");
                        persistenceService.RegistrarLogLote(idLoteBanco, "ASSINATURA", $"XML assinado: {Path.GetFileName(arquivoAssinado)}");
                    }
                    catch (Exception exDb)
                    {
                        AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após assinatura: {exDb.Message}");
                    }
                }

                if (cancelarProcessamento) return;

                // 3. Criptografar
                AtualizarEtapa("Criptografando XML...");
                var criptografiaService = new EfinanceiraCriptografiaService();
                string arquivoCriptografado = criptografiaService.CriptografarLote(arquivoAssinado, config.CertServidorThumbprint);
                status.LotesCriptografados = 1;
                AtualizarEstatisticas();
                AdicionarLog($"XML criptografado: {arquivoCriptografado}");

                // Atualizar lote no banco após criptografar
                if (idLoteBanco > 0)
                {
                    try
                    {
                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                        persistenceService.AtualizarLote(idLoteBanco, "CRIPTOGRAFADO");
                        persistenceService.RegistrarLogLote(idLoteBanco, "CRIPTOGRAFIA", $"XML criptografado: {Path.GetFileName(arquivoCriptografado)}");
                    }
                    catch (Exception exDb)
                    {
                        AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após criptografia: {exDb.Message}");
                    }
                }

                if (cancelarProcessamento) return;

                // Registrar lote processado (mesmo sem envio) - manter compatibilidade com sistema antigo
                ProtocoloPersistenciaService.RegistrarProtocolo(
                    TipoLote.Fechamento,
                    arquivoCriptografado,
                    null, // Protocolo será preenchido após envio
                    config.Periodo,
                    quantidadeEventos
                );
                AdicionarLog($"Lote registrado na lista de processados ({quantidadeEventos} evento(s)).");

                // 4. Enviar (se não estiver marcado "Apenas Processar")
                if (!chkApenasProcessar.Checked)
                {
                    try
                    {
                        AtualizarEtapa("Enviando para e-Financeira...");
                        var envioService = new EfinanceiraEnvioService();
                        var resposta = envioService.EnviarLote(arquivoCriptografado, config, cert);
                        status.LotesEnviados = 1;
                        AtualizarEstatisticas();
                        
                        if (resposta.CodigoResposta == 1)
                        {
                            // Garantir que o protocolo foi extraído (pode não ter vindo no objeto resposta)
                            string protocoloFinal = resposta.Protocolo;
                            
                            // Se não veio no objeto resposta, tentar extrair do XML
                            if (string.IsNullOrEmpty(protocoloFinal) && !string.IsNullOrEmpty(resposta.XmlCompleto))
                            {
                                try
                                {
                                    System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                    doc.LoadXml(resposta.XmlCompleto);
                                    
                                    // Primeiro tentar com GetElementsByTagName (sem namespace, como no Java)
                                    System.Xml.XmlNodeList protocoloList = doc.GetElementsByTagName("protocoloEnvio");
                                    if (protocoloList != null && protocoloList.Count > 0)
                                    {
                                        protocoloFinal = protocoloList[0].InnerText.Trim();
                                    }
                                    else
                                    {
                                        // Tentar com XPath (com namespace)
                                        System.Xml.XmlNamespaceManager nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                                        nsmgr.AddNamespace("ns", "http://www.eFinanceira.gov.br/schemas/envioLoteCriptografado/v1_2_0");
                                        
                                        System.Xml.XmlNode protocoloNode = doc.SelectSingleNode("//protocoloEnvio")
                                            ?? doc.SelectSingleNode("//ns:protocoloEnvio", nsmgr)
                                            ?? doc.SelectSingleNode("//protocolo")
                                            ?? doc.SelectSingleNode("//ns:protocolo", nsmgr);
                                        
                                        if (protocoloNode != null)
                                        {
                                            protocoloFinal = protocoloNode.InnerText.Trim();
                                        }
                                        else
                                        {
                                            // Tentar buscar por numeroProtocolo (como no método extrairProtocolo do Java)
                                            System.Xml.XmlNodeList numeroProtocoloList = doc.GetElementsByTagName("numeroProtocolo");
                                            if (numeroProtocoloList != null && numeroProtocoloList.Count > 0)
                                            {
                                                protocoloFinal = numeroProtocoloList[0].InnerText.Trim();
                                            }
                                        }
                                    }
                                    
                                    if (!string.IsNullOrEmpty(protocoloFinal))
                                    {
                                        resposta.Protocolo = protocoloFinal; // Atualizar no objeto resposta também
                                        AdicionarLog($"✓ Protocolo extraído do XML: {protocoloFinal}");
                                    }
                                }
                                catch (Exception exXml)
                                {
                                    AdicionarLog($"⚠ Erro ao extrair protocolo do XML: {exXml.Message}");
                                }
                            }
                            
                            // AGUARDAR PROTOCOLO - Não finalizar sem protocolo
                            if (!string.IsNullOrEmpty(protocoloFinal))
                            {
                                // Adicionar protocolo à lista
                                if (!status.ProtocolosEnviados.Contains(protocoloFinal))
                                {
                                    status.ProtocolosEnviados.Add(protocoloFinal);
                                }
                                
                                AdicionarLog($"✓ Lote de fechamento enviado com sucesso! Protocolo: {protocoloFinal}");
                                AdicionarLog($"════════════════════════════════════════");
                                AdicionarLog($"PROTOCOLO DO LOTE DE FECHAMENTO:");
                                AdicionarLog($"{protocoloFinal}");
                                AdicionarLog($"════════════════════════════════════════");
                                
                                // Atualizar lote no banco após envio
                                if (idLoteBanco > 0)
                                {
                                    try
                                    {
                                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                                        string xmlResposta = resposta.XmlCompleto ?? "";
                                        persistenceService.AtualizarLote(
                                            idLoteBanco,
                                            "ENVIADO",
                                            protocoloFinal,
                                            resposta.CodigoResposta,
                                            resposta.Descricao,
                                            xmlResposta,
                                            null,
                                            null,
                                            null,
                                            DateTime.Now,
                                            null,
                                            null
                                        );
                                        persistenceService.RegistrarLogLote(idLoteBanco, "ENVIO", $"Lote enviado com sucesso. Protocolo: {protocoloFinal}");
                                    }
                                    catch (Exception exDb)
                                    {
                                        AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após envio: {exDb.Message}");
                                    }
                                }
                                
                                // Atualizar protocolo no lote já registrado (sistema antigo)
                                ProtocoloPersistenciaService.RegistrarProtocolo(
                                    TipoLote.Fechamento, 
                                    arquivoCriptografado, 
                                    protocoloFinal,
                                    config.Periodo,
                                    quantidadeEventos
                                );
                                
                                // Aguardar um momento para garantir que o protocolo foi salvo
                                System.Threading.Thread.Sleep(500);
                                
                                // Atualizar lista na aba Consulta com o protocolo já preenchido
                                if (ConsultaForm != null)
                                {
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        ConsultaForm.AtualizarListaLotes();
                                    });
                                }
                                
                                // Exibir MessageBox destacando o protocolo
                                MessageBox.Show(
                                    $"Lote de fechamento enviado com sucesso!\n\n" +
                                    $"PROTOCOLO: {protocoloFinal}\n\n" +
                                    $"Este protocolo foi salvo e pode ser consultado na aba 'Consulta'.",
                                    "Envio Concluído",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );
                                
                                AtualizarEstatisticas();
                            }
                            else
                            {
                                // Se não recebeu protocolo, aguardar e tentar novamente ou informar erro
                                AdicionarLog($"⚠ ATENÇÃO: Lote de fechamento enviado mas protocolo não foi retornado!");
                                AdicionarLog($"  Aguardando resposta do servidor...");
                                
                                // Tentar aguardar um pouco mais e verificar se há protocolo na resposta XML
                                System.Threading.Thread.Sleep(2000);
                                
                                // Se ainda não tiver protocolo, registrar como erro
                                if (idLoteBanco > 0)
                                {
                                    try
                                    {
                                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                                        persistenceService.AtualizarLote(
                                            idLoteBanco,
                                            "ENVIADO_SEM_PROTOCOLO",
                                            null,
                                            resposta.CodigoResposta,
                                            resposta.Descricao + " (Protocolo não retornado)",
                                            resposta.XmlCompleto ?? "",
                                            null,
                                            null,
                                            null,
                                            DateTime.Now,
                                            null,
                                            "Protocolo não retornado pelo servidor"
                                        );
                                    }
                                    catch (Exception exDb)
                                    {
                                        AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                                    }
                                }
                                
                                MessageBox.Show(
                                    $"Lote de fechamento enviado, mas o protocolo não foi retornado pelo servidor.\n\n" +
                                    $"Código de Resposta: {resposta.CodigoResposta}\n" +
                                    $"Descrição: {resposta.Descricao}\n\n" +
                                    $"Verifique o XML de resposta ou consulte o lote mais tarde.",
                                    "Aviso - Protocolo Não Recebido",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning
                                );
                            }
                        }
                        else if (resposta.CodigoResposta == 7)
                        {
                            AdicionarLog($"✗ Lote de fechamento REJEITADO - Código: {resposta.CodigoResposta}");
                            AdicionarLog($"  Descrição: {resposta.Descricao}");
                            
                            // Atualizar lote no banco com erro
                            if (idLoteBanco > 0)
                            {
                                try
                                {
                                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                                    string erroMsg = $"REJEITADO - Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}";
                                    persistenceService.AtualizarLote(
                                        idLoteBanco,
                                        "REJEITADO",
                                        null,
                                        resposta.CodigoResposta,
                                        resposta.Descricao,
                                        resposta.XmlCompleto ?? "",
                                        null,
                                        null,
                                        null,
                                        DateTime.Now,
                                        null,
                                        erroMsg
                                    );
                                    persistenceService.RegistrarLogLote(idLoteBanco, "ENVIO_ERRO", erroMsg);
                                }
                                catch (Exception exDb)
                                {
                                    AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após rejeição: {exDb.Message}");
                                }
                            }
                            
                            if (resposta.Ocorrencias != null && resposta.Ocorrencias.Count > 0)
                            {
                                foreach (var ocorr in resposta.Ocorrencias)
                                {
                                    AdicionarLog($"  Ocorrência: {ocorr.Codigo} - {ocorr.Descricao} ({ocorr.Tipo})");
                                }
                            }
                        }
                        else
                        {
                            // SEMPRE tentar extrair e salvar o protocolo, mesmo se código não for 1
                            string protocoloExtraido = resposta.Protocolo;
                            
                            // Se não veio no objeto resposta, tentar extrair do XML (seguindo lógica do Java)
                            if (string.IsNullOrEmpty(protocoloExtraido) && !string.IsNullOrEmpty(resposta.XmlCompleto))
                            {
                                try
                                {
                                    System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                    doc.LoadXml(resposta.XmlCompleto);
                                    
                                    // Primeiro tentar com GetElementsByTagName (sem namespace, como no Java)
                                    System.Xml.XmlNodeList protocoloList = doc.GetElementsByTagName("protocoloEnvio");
                                    if (protocoloList != null && protocoloList.Count > 0)
                                    {
                                        protocoloExtraido = protocoloList[0].InnerText.Trim();
                                    }
                                    else
                                    {
                                        // Tentar com XPath (com namespace)
                                        System.Xml.XmlNamespaceManager nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                                        nsmgr.AddNamespace("ns", "http://www.eFinanceira.gov.br/schemas/envioLoteCriptografado/v1_2_0");
                                        
                                        System.Xml.XmlNode protocoloNode = doc.SelectSingleNode("//protocoloEnvio")
                                            ?? doc.SelectSingleNode("//ns:protocoloEnvio", nsmgr)
                                            ?? doc.SelectSingleNode("//protocolo")
                                            ?? doc.SelectSingleNode("//ns:protocolo", nsmgr);
                                        
                                        if (protocoloNode != null)
                                        {
                                            protocoloExtraido = protocoloNode.InnerText.Trim();
                                        }
                                        else
                                        {
                                            // Tentar buscar por numeroProtocolo (como no método extrairProtocolo do Java)
                                            System.Xml.XmlNodeList numeroProtocoloList = doc.GetElementsByTagName("numeroProtocolo");
                                            if (numeroProtocoloList != null && numeroProtocoloList.Count > 0)
                                            {
                                                protocoloExtraido = numeroProtocoloList[0].InnerText.Trim();
                                            }
                                        }
                                    }
                                    
                                    if (!string.IsNullOrEmpty(protocoloExtraido))
                                    {
                                        AdicionarLog($"✓ Protocolo extraído do XML: {protocoloExtraido}");
                                    }
                                }
                                catch (Exception exXml)
                                {
                                    AdicionarLog($"⚠ Erro ao extrair protocolo do XML: {exXml.Message}");
                                    System.Diagnostics.Debug.WriteLine($"Erro ao extrair protocolo do XML: {exXml.Message}");
                                }
                            }
                            
                            if (!string.IsNullOrEmpty(protocoloExtraido))
                            {
                                AdicionarLog($"✓ Protocolo extraído: {protocoloExtraido}");
                                
                                // Adicionar protocolo à lista
                                if (!status.ProtocolosEnviados.Contains(protocoloExtraido))
                                {
                                    status.ProtocolosEnviados.Add(protocoloExtraido);
                                }
                                
                                // Atualizar protocolo no lote já registrado (sistema antigo)
                                ProtocoloPersistenciaService.RegistrarProtocolo(
                                    TipoLote.Fechamento, 
                                    arquivoCriptografado, 
                                    protocoloExtraido,
                                    config.Periodo,
                                    quantidadeEventos
                                );
                                
                                // Aguardar um momento para garantir que o protocolo foi salvo
                                System.Threading.Thread.Sleep(500);
                                
                                // Atualizar lista na aba Consulta
                                if (ConsultaForm != null)
                                {
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        ConsultaForm.AtualizarListaLotes();
                                    });
                                }
                            }
                            
                            AdicionarLog($"⚠ Lote de fechamento - Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}");
                            
                            // Atualizar lote no banco com resposta e protocolo (se encontrado)
                            if (idLoteBanco > 0)
                            {
                                try
                                {
                                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                                    persistenceService.AtualizarLote(
                                        idLoteBanco,
                                        "ENVIADO_COM_RESPOSTA",
                                        protocoloExtraido, // Salvar protocolo mesmo se código não for 1
                                        resposta.CodigoResposta,
                                        resposta.Descricao,
                                        resposta.XmlCompleto ?? "",
                                        null,
                                        null,
                                        null,
                                        DateTime.Now,
                                        null,
                                        null
                                    );
                                    
                                    if (!string.IsNullOrEmpty(protocoloExtraido))
                                    {
                                        persistenceService.RegistrarLogLote(idLoteBanco, "ENVIO", 
                                            $"Lote enviado. Protocolo: {protocoloExtraido}, Código: {resposta.CodigoResposta}");
                                    }
                                }
                                catch (Exception exDb)
                                {
                                    AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                                }
                            }
                        }
                    }
                    catch (Exception exEnv)
                    {
                        AdicionarLog($"✗ Erro ao enviar lote de fechamento: {exEnv.Message}");
                    }
                }
                else
                {
                    AdicionarLog("Envio não realizado (modo 'Apenas Processar' ativo).");
                    
                    // Atualizar lista na aba Consulta mesmo sem protocolo
                    if (ConsultaForm != null)
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            ConsultaForm.AtualizarListaLotes();
                        });
                    }
                }

                status.LotesProcessados = 1;
                AtualizarEstatisticas();
                AtualizarEtapa("Processamento concluído com sucesso!");
                AdicionarLog("Processamento de fechamento concluído.");

                this.Invoke((MethodInvoker)delegate
                {
                    btnProcessarAbertura.Enabled = true;
                    btnProcessarMovimentacao.Enabled = true;
                    btnProcessarFechamento.Enabled = true;
                    btnCancelar.Enabled = false;
                });
            }
            catch (Exception ex)
            {
                AdicionarLog($"ERRO: {ex.Message}");
                status.LotesComErro = 1;
                this.Invoke((MethodInvoker)delegate
                {
                    btnProcessarAbertura.Enabled = true;
                    btnProcessarMovimentacao.Enabled = true;
                    btnProcessarFechamento.Enabled = true;
                    btnCancelar.Enabled = false;
                    MessageBox.Show($"Erro ao processar: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }

        private X509Certificate2 BuscarCertificado(string thumbprint)
        {
            if (string.IsNullOrEmpty(thumbprint))
            {
                throw new Exception("Thumbprint do certificado não configurado.");
            }

            string thumbprintNormalizado = thumbprint.Replace(" ", "").Replace("-", "").ToUpper();

            X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            try
            {
                store.Open(OpenFlags.ReadOnly);
                foreach (X509Certificate2 cert in store.Certificates)
                {
                    if (cert.Thumbprint.Replace(" ", "").Replace("-", "").ToUpper() == thumbprintNormalizado)
                    {
                        return cert;
                    }
                }
            }
            finally
            {
                store.Close();
            }

            throw new Exception($"Certificado com thumbprint '{thumbprint}' não encontrado.");
        }

        private void AtualizarEtapa(string etapa)
        {
            this.Invoke((MethodInvoker)delegate
            {
                status.EtapaAtual = etapa;
                lblEtapaAtual.Text = $"Etapa: {etapa}";
                status.TempoDecorrido = DateTime.Now - status.InicioProcessamento;
                lblTempoDecorrido.Text = $"Tempo Decorrido: {status.TempoDecorrido:hh\\:mm\\:ss}";
            });
        }

        private void AdicionarLog(string mensagem)
        {
            this.Invoke((MethodInvoker)delegate
            {
                try
                {
                    string logEntry = $"[{DateTime.Now:HH:mm:ss}] {mensagem}";
                    lstLog.Items.Add(logEntry);
                    
                    // Verificar se há itens antes de definir SelectedIndex
                    if (lstLog.Items.Count > 0)
                    {
                        lstLog.SelectedIndex = lstLog.Items.Count - 1;
                        lstLog.TopIndex = lstLog.Items.Count - 1; // Scroll para o final
                    }
                    
                    status.MensagemAtual = mensagem;
                    lblMensagemAtual.Text = $"Mensagem: {mensagem}";
                }
                catch (Exception ex)
                {
                    // Log do erro mas não interrompe o processamento
                    System.Diagnostics.Debug.WriteLine($"Erro ao adicionar log: {ex.Message}");
                }
            });
        }

        private void AtualizarEstatisticas()
        {
            this.Invoke((MethodInvoker)delegate
            {
                lblTotalLotes.Text = $"Total de Lotes: {status.TotalLotes}";
                lblLotesProcessados.Text = $"Processados: {status.LotesProcessados}";
                lblLotesAssinados.Text = $"Assinados: {status.LotesAssinados}";
                lblLotesCriptografados.Text = $"Criptografados: {status.LotesCriptografados}";
                
                // Atualizar "Enviados" com protocolos se houver
                if (status.ProtocolosEnviados != null && status.ProtocolosEnviados.Count > 0)
                {
                    string protocolosStr = string.Join(", ", status.ProtocolosEnviados);
                    lblLotesEnviados.Text = $"Enviados: {status.LotesEnviados} ({protocolosStr})";
                }
                else
                {
                    lblLotesEnviados.Text = $"Enviados: {status.LotesEnviados}";
                }
                
                lblLotesComErro.Text = $"Com Erro: {status.LotesComErro}";

                if (status.TotalLotes > 0)
                {
                    int progresso = (int)((double)status.LotesProcessados / status.TotalLotes * 100);
                    progressBarGeral.Value = Math.Min(progresso, 100);
                    lblProgressoGeral.Text = $"{progresso}%";
                }
            });
        }
    }
}
