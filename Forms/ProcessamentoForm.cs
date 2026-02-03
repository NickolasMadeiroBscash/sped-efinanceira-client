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
using ExemploAssinadorXML;

namespace ExemploAssinadorXML.Forms
{
    public partial class ProcessamentoForm : Form
    {
        // Constantes para strings repetidas
        private const string PROTOCOLO_ENVIO = "protocoloEnvio";
        private const string NUMERO_PROTOCOLO = "numeroProtocolo";
        private const string PROTOCOLO_XPATH = "//protocoloEnvio";
        private const string PROTOCOLO_NS_XPATH = "//ns:protocoloEnvio";
        private const string PROTOCOLO_SIMPLE_XPATH = "//protocolo";
        private const string PROTOCOLO_NS_SIMPLE_XPATH = "//ns:protocolo";
        private const string NAMESPACE_ENVIO_LOTE = "http://www.eFinanceira.gov.br/schemas/envioLoteCriptografado/v1_2_0";
        private const string SUFIXO_ASSINADO = "-ASSINADO.xml";
        private const string STATUS_ASSINADO = "ASSINADO";
        private const string STATUS_ASSINATURA = "ASSINATURA";
        private const string STATUS_CRIPTOGRAFADO = "CRIPTOGRAFADO";
        private const string STATUS_CRIPTOGRAFIA = "CRIPTOGRAFIA";
        private const string STATUS_ENVIADO = "ENVIADO";
        private const string STATUS_REJEITADO = "REJEITADO";
        private const string STATUS_ENVIO = "ENVIO";
        private const string AMBIENTE_HOMOLOG = "HOMOLOG";
        private const string LOG_GERACAO = "GERACAO";
        private const string CAMPO_PREFIXO = "Campo:";
        private const string TITULO_CAMPOS_FALTANDO = "Campos Obrigatórios Faltando";
        private const string MSG_ACESSE_CONFIG = "Acesse a aba 'Configuração' e preencha todos os campos obrigatórios";
        private const string MSG_CAMPO_CNPJ = "Campo: 'CNPJ do Declarante' na seção 'Configuração Geral'";
        private const string MSG_CAMPO_CERT_ASSINATURA = "Campo: 'Certificado para Assinatura' na seção 'Configuração Geral'";
        private const string MSG_CAMPO_CERT_SERVIDOR = "Campo: 'Certificado do Servidor' na seção 'Configuração Geral'";
        private const string MSG_CAMPO_PERIODO = "Campo: 'Período' na seção 'Configuração Geral'";
        private const string MSG_CAMPO_DIRETORIO = "Campo: 'Diretório de Lotes' na seção 'Configuração Geral'";
        private const string MSG_CONFIG_NAO_INICIALIZADA = "Configuração geral não foi inicializada. Preencha os seguintes campos:";
        private const string MSG_CAMPOS_FALTANDO = "Os seguintes campos obrigatórios estão faltando:\n\n";
        
        // Constante para formato de data
        private static readonly System.Globalization.CultureInfo CULTURE_INFO_PT_BR = System.Globalization.CultureInfo.GetCultureInfo("pt-BR");
        private GroupBox grpControles;
        private Button btnProcessarAbertura;
        private Button btnProcessarMovimentacao;
        private Button btnProcessarFechamento;
        private Button btnProcessarCadastroDeclarante;
        private Button btnProcessarCompleto;
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
        private Label lblStatusEtapa;
        private Label lblAberturaFinalizada;
        private Label lblTempoMedioPorLote;

        private readonly StatusProcessamento status;
        private bool cancelarProcessamento = false;
        private System.Windows.Forms.Timer timerAtualizacao;
        private bool processamentoAtivo = false;

        public ConfiguracaoForm ConfigForm { get; set; }
        public ConsultaForm ConsultaForm { get; set; }
        public MainForm MainFormParent { get; set; }
        
        /// <summary>
        /// Desabilita todos os controles durante o processamento, exceto o botão Cancelar
        /// </summary>
        private void DesabilitarControlesDuranteProcessamento()
        {
            if (this.InvokeRequired)
            {
                this.Invoke((MethodInvoker)delegate { DesabilitarControlesDuranteProcessamento(); });
                return;
            }
            
            // Desabilitar botões de processamento
            btnProcessarAbertura.Enabled = false;
            btnProcessarMovimentacao.Enabled = false;
            btnProcessarFechamento.Enabled = false;
            btnProcessarCadastroDeclarante.Enabled = false;
            btnProcessarCompleto.Enabled = false;
            
            // Habilitar apenas o botão Cancelar
            btnCancelar.Enabled = true;
            
            // Desabilitar checkbox
            chkApenasProcessar.Enabled = false;
            
            // Desabilitar abas do TabControl principal
            if (MainFormParent != null && MainFormParent.TabControlPrincipal != null)
            {
                foreach (TabPage tab in MainFormParent.TabControlPrincipal.TabPages)
                {
                    tab.Enabled = false;
                }
                // Manter a aba de Processamento habilitada para ver o progresso
                if (MainFormParent.TabControlPrincipal.TabPages.Count > 2)
                {
                    MainFormParent.TabControlPrincipal.TabPages[2].Enabled = true; // Aba Processamento (índice 2)
                }
            }
        }
        
        /// <summary>
        /// Reabilita todos os controles após o processamento ou cancelamento
        /// </summary>
        private void ReabilitarControlesAposProcessamento()
        {
            if (this.InvokeRequired)
            {
                this.Invoke((MethodInvoker)delegate { ReabilitarControlesAposProcessamento(); });
                return;
            }
            
            // Reabilitar botões de processamento
            btnProcessarAbertura.Enabled = true;
            btnProcessarMovimentacao.Enabled = true;
            btnProcessarFechamento.Enabled = true;
            btnProcessarCadastroDeclarante.Enabled = true;
            btnProcessarCompleto.Enabled = true;
            
            // Desabilitar botão Cancelar
            btnCancelar.Enabled = false;
            
            // Reabilitar checkbox
            chkApenasProcessar.Enabled = true;
            
            // Reabilitar abas do TabControl principal
            if (MainFormParent != null && MainFormParent.TabControlPrincipal != null)
            {
                foreach (TabPage tab in MainFormParent.TabControlPrincipal.TabPages)
                {
                    tab.Enabled = true;
                }
            }
        }

        public ProcessamentoForm()
        {
            InitializeComponent();
            status = new StatusProcessamento();
            status.StatusEtapa = 0; // Inicializar como 0 (Aguardando)
            
            // Criar timer para atualização em tempo real
            timerAtualizacao = new System.Windows.Forms.Timer();
            timerAtualizacao.Interval = 500; // Atualizar a cada 500ms
            timerAtualizacao.Tick += TimerAtualizacao_Tick;
        }
        
        private void TimerAtualizacao_Tick(object sender, EventArgs e)
        {
            if (processamentoAtivo && status.InicioProcessamento != default(DateTime))
            {
                // Atualizar tempo decorrido em tempo real
                status.TempoDecorrido = DateTime.Now - status.InicioProcessamento;
                this.Invoke((MethodInvoker)delegate
                {
                    lblTempoDecorrido.Text = $"Tempo Decorrido: {status.TempoDecorrido:hh\\:mm\\:ss}";
                    
                    // Atualizar tempo médio por lote em tempo real durante processamento
                    if ((status.StatusEtapa == 3 || status.StatusEtapa == 4) && status.TotalLotesMovimentacao > 0)
                    {
                        if (status.TemposLotesMovimentacao.Count > 0 && status.LotesMovimentacaoProcessados > 0)
                        {
                            TimeSpan tempoTotal = DateTime.Now - status.TemposLotesMovimentacao[0];
                            TimeSpan tempoMedio = TimeSpan.FromMilliseconds(tempoTotal.TotalMilliseconds / status.LotesMovimentacaoProcessados);
                            lblTempoMedioPorLote.Text = $"Tempo Médio/Lote: {tempoMedio:mm\\:ss} (calculando...)";
                        }
                    }
                    
                    // Atualizar barra de progresso com animação marquee quando estiver consultando
                    if (status.StatusEtapa == 0 || (status.TotalLotes == 0 && processamentoAtivo))
                    {
                        if (progressBarGeral.Style != ProgressBarStyle.Marquee)
                        {
                            progressBarGeral.Style = ProgressBarStyle.Marquee;
                            progressBarGeral.MarqueeAnimationSpeed = 30;
                        }
                        lblProgressoGeral.Text = "Consultando banco de dados...";
                    }
                    else if (status.TotalLotes > 0)
                    {
                        if (progressBarGeral.Style != ProgressBarStyle.Continuous)
                        {
                            progressBarGeral.Style = ProgressBarStyle.Continuous;
                        }
                        int progresso = (int)((double)status.LotesProcessados / status.TotalLotes * 100);
                        progressBarGeral.Value = Math.Min(Math.Max(progresso, 0), 100);
                        lblProgressoGeral.Text = $"{progresso}%";
                    }
                });
            }
        }
        
        private void LimparEstatisticas()
        {
            status.TotalLotes = 0;
            status.LotesProcessados = 0;
            status.LotesAssinados = 0;
            status.LotesCriptografados = 0;
            status.LotesEnviados = 0;
            status.LotesComErro = 0;
            status.ProtocolosEnviados.Clear();
            status.StatusEtapa = 0;
            status.AberturaFinalizada = false;
            status.LotesMovimentacaoProcessados = 0;
            status.TotalLotesMovimentacao = 0;
            status.TempoMedioPorLote = null;
            status.TemposLotesMovimentacao.Clear();
            status.EtapaAtual = "Aguardando...";
            status.MensagemAtual = "-";
            status.InicioProcessamento = DateTime.Now;
            status.TempoDecorrido = TimeSpan.Zero;
            
            this.Invoke((MethodInvoker)delegate
            {
                progressBarGeral.Value = 0;
                progressBarGeral.Style = ProgressBarStyle.Marquee;
                progressBarGeral.MarqueeAnimationSpeed = 30;
                lblProgressoGeral.Text = "Iniciando...";
                lstLog.Items.Clear();
                AtualizarEstatisticas();
            });
        }
        
        private void IniciarProcessamento()
        {
            processamentoAtivo = true;
            timerAtualizacao.Start();
            LimparEstatisticas();
        }
        
        private void FinalizarProcessamento(bool cancelado = false, string mensagemErro = null)
        {
            processamentoAtivo = false;
            timerAtualizacao.Stop();
            
            this.Invoke((MethodInvoker)delegate
            {
                // Parar animação marquee
                if (progressBarGeral.Style == ProgressBarStyle.Marquee)
                {
                    progressBarGeral.Style = ProgressBarStyle.Continuous;
                    progressBarGeral.MarqueeAnimationSpeed = 0;
                }
                
                if (!string.IsNullOrEmpty(mensagemErro))
                {
                    // Se houve erro, mostrar mensagem de erro
                    if (status.TotalLotes > 0)
                    {
                        int progresso = (int)((double)status.LotesProcessados / status.TotalLotes * 100);
                        progressBarGeral.Value = Math.Min(Math.Max(progresso, 0), 100);
                        lblProgressoGeral.Text = $"Erro ({progresso}%)";
                    }
                    else
                    {
                        progressBarGeral.Value = 0;
                        lblProgressoGeral.Text = "Erro";
                    }
                    AtualizarEtapa($"Erro: {mensagemErro}");
                }
                else if (cancelado)
                {
                    // Se foi cancelado, manter o progresso atual e mostrar mensagem
                    if (status.TotalLotes > 0)
                    {
                        int progresso = (int)((double)status.LotesProcessados / status.TotalLotes * 100);
                        progressBarGeral.Value = Math.Min(Math.Max(progresso, 0), 100);
                        lblProgressoGeral.Text = $"Cancelado ({progresso}%)";
                    }
                    else
                    {
                        progressBarGeral.Value = 0;
                        lblProgressoGeral.Text = "Cancelado";
                    }
                    AtualizarEtapa("Processamento cancelado pelo usuário.");
                }
                else
                {
                    // Se foi finalizado normalmente
                    if (status.TotalLotes > 0)
                    {
                        int progresso = (int)((double)status.LotesProcessados / status.TotalLotes * 100);
                        progressBarGeral.Value = Math.Min(Math.Max(progresso, 0), 100);
                        lblProgressoGeral.Text = $"Concluído ({progresso}%)";
                    }
                    else
                    {
                        progressBarGeral.Value = 100;
                        lblProgressoGeral.Text = "Concluído (100%)";
                    }
                }
            });
        }

        /// <summary>
        /// Extrai o protocolo da resposta do envio, tentando múltiplas estratégias
        /// </summary>
        private string ExtrairProtocolo(RespostaEnvioEfinanceira resposta)
        {
            string protocolo = resposta.Protocolo;

            // Se já veio no objeto resposta, retornar
            if (!string.IsNullOrEmpty(protocolo))
            {
                return protocolo;
            }

            // Se não veio no objeto resposta, tentar extrair do XML
            if (string.IsNullOrEmpty(resposta.XmlCompleto))
            {
                return null;
            }

            try
            {
                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                doc.LoadXml(resposta.XmlCompleto);

                // Primeiro tentar com GetElementsByTagName (sem namespace, como no Java)
                System.Xml.XmlNodeList protocoloList = doc.GetElementsByTagName(PROTOCOLO_ENVIO);
                if (protocoloList != null && protocoloList.Count > 0)
                {
                    protocolo = protocoloList[0].InnerText.Trim();
                    if (!string.IsNullOrEmpty(protocolo))
                    {
                        System.Diagnostics.Debug.WriteLine($"Protocolo encontrado via GetElementsByTagName: {protocolo}");
                        return protocolo;
                    }
                }

                // Tentar com XPath (com namespace)
                System.Xml.XmlNamespaceManager nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("ns", NAMESPACE_ENVIO_LOTE);

                System.Xml.XmlNode protocoloNode = doc.SelectSingleNode(PROTOCOLO_XPATH)
                    ?? doc.SelectSingleNode(PROTOCOLO_NS_XPATH, nsmgr)
                    ?? doc.SelectSingleNode(PROTOCOLO_SIMPLE_XPATH)
                    ?? doc.SelectSingleNode(PROTOCOLO_NS_SIMPLE_XPATH, nsmgr);

                if (protocoloNode != null)
                {
                    protocolo = protocoloNode.InnerText.Trim();
                    if (!string.IsNullOrEmpty(protocolo))
                    {
                        System.Diagnostics.Debug.WriteLine($"Protocolo encontrado via XPath: {protocolo}");
                        return protocolo;
                    }
                }

                // Tentar buscar por numeroProtocolo (como no método extrairProtocolo do Java)
                System.Xml.XmlNodeList numeroProtocoloList = doc.GetElementsByTagName(NUMERO_PROTOCOLO);
                if (numeroProtocoloList != null && numeroProtocoloList.Count > 0)
                {
                    protocolo = numeroProtocoloList[0].InnerText.Trim();
                    if (!string.IsNullOrEmpty(protocolo))
                    {
                        System.Diagnostics.Debug.WriteLine($"Protocolo encontrado via numeroProtocolo: {protocolo}");
                        return protocolo;
                    }
                }

                System.Diagnostics.Debug.WriteLine("AVISO: Protocolo não encontrado no XML de resposta.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao extrair protocolo do XML: {ex.Message}");
                AdicionarLog($"⚠ Erro ao extrair protocolo do XML: {ex.Message}");
            }

            return null;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Controles
            grpControles = new GroupBox();
            grpControles.Text = "Controles";
            grpControles.Location = new Point(10, 10);
            grpControles.Size = new Size(900, 80);
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

            btnProcessarCadastroDeclarante = new Button();
            btnProcessarCadastroDeclarante.Text = "Processar Cadastro";
            btnProcessarCadastroDeclarante.Location = new Point(490, 25);
            btnProcessarCadastroDeclarante.Size = new Size(150, 35);
            btnProcessarCadastroDeclarante.Click += BtnProcessarCadastroDeclarante_Click;

            btnProcessarCompleto = new Button();
            btnProcessarCompleto.Text = "Processar Completo";
            btnProcessarCompleto.Location = new Point(650, 25);
            btnProcessarCompleto.Size = new Size(150, 35);
            btnProcessarCompleto.BackColor = Color.LightGreen;
            btnProcessarCompleto.Click += BtnProcessarCompleto_Click;

            btnCancelar = new Button();
            btnCancelar.Text = "Cancelar";
            btnCancelar.Location = new Point(810, 25);
            btnCancelar.Size = new Size(100, 35);
            btnCancelar.Enabled = false;
            btnCancelar.Click += BtnCancelar_Click;

            chkApenasProcessar = new CheckBox();
            chkApenasProcessar.Text = "Apenas Processar (não enviar)";
            chkApenasProcessar.Location = new Point(920, 32);
            chkApenasProcessar.Size = new Size(200, 20);

            grpControles.Size = new Size(1130, 80);

            grpControles.Controls.AddRange(new Control[] {
                btnProcessarAbertura, btnProcessarMovimentacao, btnProcessarFechamento,
                btnProcessarCadastroDeclarante, btnProcessarCompleto, btnCancelar, chkApenasProcessar
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

            lblStatusEtapa = new Label();
            lblStatusEtapa.Text = "Status Etapa: Aguardando";
            lblStatusEtapa.Location = new Point(10, 115);
            lblStatusEtapa.Size = new Size(250, 20);
            lblStatusEtapa.Font = new Font(lblStatusEtapa.Font, FontStyle.Bold);
            lblStatusEtapa.ForeColor = Color.Blue;

            lblAberturaFinalizada = new Label();
            lblAberturaFinalizada.Text = "Abertura: Não iniciada";
            lblAberturaFinalizada.Location = new Point(270, 115);
            lblAberturaFinalizada.Size = new Size(200, 20);

            lblTempoMedioPorLote = new Label();
            lblTempoMedioPorLote.Text = "Tempo Médio/Lote: -";
            lblTempoMedioPorLote.Location = new Point(480, 115);
            lblTempoMedioPorLote.Size = new Size(250, 20);

            grpEstatisticas.Size = new Size(750, 150);

            grpEstatisticas.Controls.AddRange(new Control[] {
                lblTotalLotes, lblLotesProcessados, lblLotesAssinados,
                lblLotesCriptografados, lblLotesEnviados, lblLotesComErro,
                lblTempoDecorrido, lblTempoEstimado, lblStatusEtapa,
                lblAberturaFinalizada, lblTempoMedioPorLote
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
                string mensagemFinal = erro.Contains(CAMPO_PREFIXO) || erro.Contains("Aba") 
                    ? $"Não é possível processar o evento de Abertura.\n\n{MSG_CAMPOS_FALTANDO}{erro}\n\nPor favor, preencha os campos indicados acima na aba Configuração."
                    : $"Não é possível processar o evento de Abertura.\n\n{erro}";
                MessageBox.Show(mensagemFinal, 
                    TITULO_CAMPOS_FALTANDO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DesabilitarControlesDuranteProcessamento();
            cancelarProcessamento = false;

            Task.Run(() => ProcessarAbertura());
        }

        private void BtnProcessarMovimentacao_Click(object sender, EventArgs e)
        {
            string erro = ValidarDadosMovimentacao();
            if (!string.IsNullOrEmpty(erro))
            {
                string mensagemFinal = erro.Contains(CAMPO_PREFIXO) || erro.Contains("Aba") 
                    ? $"Não é possível processar o evento de Movimentação.\n\n{MSG_CAMPOS_FALTANDO}{erro}\n\nPor favor, preencha os campos indicados acima na aba Configuração."
                    : $"Não é possível processar o evento de Movimentação.\n\n{erro}";
                MessageBox.Show(mensagemFinal, 
                    TITULO_CAMPOS_FALTANDO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DesabilitarControlesDuranteProcessamento();
            cancelarProcessamento = false;

            Task.Run(() => ProcessarMovimentacao());
        }

        private void BtnProcessarCompleto_Click(object sender, EventArgs e)
        {
            string erro = ValidarPeriodosCompletos();
            if (!string.IsNullOrEmpty(erro))
            {
                MessageBox.Show(erro, "Validação de Períodos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            erro = ValidarDadosAbertura();
            if (!string.IsNullOrEmpty(erro))
            {
                string mensagemFinal = erro.Contains(CAMPO_PREFIXO) || erro.Contains("Aba") 
                    ? $"{TITULO_CAMPOS_FALTANDO}\n\n{erro}\n\n{MSG_ACESSE_CONFIG}"
                    : erro;
                MessageBox.Show(mensagemFinal, TITULO_CAMPOS_FALTANDO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            erro = ValidarDadosFechamento();
            if (!string.IsNullOrEmpty(erro))
            {
                string mensagemFinal = erro.Contains(CAMPO_PREFIXO) || erro.Contains("Aba") 
                    ? $"{TITULO_CAMPOS_FALTANDO}\n\n{erro}\n\n{MSG_ACESSE_CONFIG}"
                    : erro;
                MessageBox.Show(mensagemFinal, TITULO_CAMPOS_FALTANDO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var resultado = MessageBox.Show(
                "Deseja processar tudo automaticamente?\n\n" +
                "O sistema irá:\n" +
                "1. Processar abertura\n" +
                "2. Processar todas as movimentações\n" +
                "3. Processar fechamento\n\n" +
                "Tudo será feito sequencialmente de forma automatizada.",
                "Processamento Completo",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (resultado != DialogResult.Yes)
                return;

            DesabilitarControlesDuranteProcessamento();
            cancelarProcessamento = false;

            // Resetar status
            status.StatusEtapa = 0; // Aguardando
            status.AberturaFinalizada = false;
            status.LotesMovimentacaoProcessados = 0;
            status.TotalLotesMovimentacao = 0;
            status.TempoMedioPorLote = null;
            status.TemposLotesMovimentacao.Clear();

            Task.Run(() => ProcessarCompleto());
        }

        private void BtnProcessarFechamento_Click(object sender, EventArgs e)
        {
            string erro = ValidarDadosFechamento();
            if (!string.IsNullOrEmpty(erro))
            {
                string mensagemFinal = erro.Contains(CAMPO_PREFIXO) || erro.Contains("Aba") 
                    ? $"Não é possível processar o evento de Fechamento.\n\n{MSG_CAMPOS_FALTANDO}{erro}\n\nPor favor, preencha os campos indicados acima na aba Configuração."
                    : $"Não é possível processar o evento de Fechamento.\n\n{erro}";
                MessageBox.Show(mensagemFinal, 
                    TITULO_CAMPOS_FALTANDO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DesabilitarControlesDuranteProcessamento();
            cancelarProcessamento = false;

            Task.Run(() => ProcessarFechamento());
        }

        private void BtnProcessarCadastroDeclarante_Click(object sender, EventArgs e)
        {
            string erro = ValidarDadosCadastroDeclarante();
            if (!string.IsNullOrEmpty(erro))
            {
                string mensagemFinal = erro.Contains(CAMPO_PREFIXO) || erro.Contains("Aba") 
                    ? $"Não é possível processar o evento de Cadastro de Declarante.\n\n{MSG_CAMPOS_FALTANDO}{erro}\n\nPor favor, preencha os campos indicados acima na aba Configuração."
                    : $"Não é possível processar o evento de Cadastro de Declarante.\n\n{erro}";
                MessageBox.Show(mensagemFinal, 
                    TITULO_CAMPOS_FALTANDO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DesabilitarControlesDuranteProcessamento();
            cancelarProcessamento = false;

            Task.Run(() => ProcessarCadastroDeclarante());
        }

        private void BtnCancelar_Click(object sender, EventArgs e)
        {
            cancelarProcessamento = true;
            AdicionarLog("Processamento cancelado pelo usuário.");
            FinalizarProcessamento(cancelado: true);
            ReabilitarControlesAposProcessamento();
        }

        private string ValidarDadosAbertura()
        {
            List<string> camposFaltando = new List<string>();

            if (ConfigForm == null)
            {
                camposFaltando.Add(MSG_ACESSE_CONFIG);
                camposFaltando.Add(MSG_CAMPO_CNPJ);
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);
                camposFaltando.Add(MSG_CAMPO_PERIODO);
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);
                camposFaltando.Add("Aba 'Dados de Abertura': Preencha todos os campos obrigatórios");
                return string.Join("\n", camposFaltando);
            }

            if (ConfigForm.Config == null)
            {
                camposFaltando.Add(MSG_CONFIG_NAO_INICIALIZADA);
                camposFaltando.Add(MSG_CAMPO_CNPJ);
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);
                camposFaltando.Add(MSG_CAMPO_PERIODO);
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);
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
                camposFaltando.Add(MSG_CAMPO_CNPJ);

            if (string.IsNullOrWhiteSpace(config.CertThumbprint))
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);

            if (string.IsNullOrWhiteSpace(config.CertServidorThumbprint))
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);

            if (string.IsNullOrWhiteSpace(config.Periodo))
                camposFaltando.Add(MSG_CAMPO_PERIODO);

            if (string.IsNullOrWhiteSpace(config.DiretorioLotes))
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);

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
                return MSG_CAMPOS_FALTANDO + string.Join("\n", camposFaltando);
            }

            return null;
        }

        private string ValidarDadosMovimentacao()
        {
            List<string> camposFaltando = new List<string>();

            if (ConfigForm == null)
            {
                camposFaltando.Add(MSG_ACESSE_CONFIG);
                camposFaltando.Add(MSG_CAMPO_CNPJ);
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);
                camposFaltando.Add(MSG_CAMPO_PERIODO);
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);
                camposFaltando.Add("Campo: 'Page Size' na seção 'Configuração de Processamento'");
                return string.Join("\n", camposFaltando);
            }

            if (ConfigForm.Config == null)
            {
                camposFaltando.Add(MSG_CONFIG_NAO_INICIALIZADA);
                camposFaltando.Add(MSG_CAMPO_CNPJ);
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);
                camposFaltando.Add(MSG_CAMPO_PERIODO);
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);
                camposFaltando.Add("Campo: 'Page Size' na seção 'Configuração de Processamento'");
                return string.Join("\n", camposFaltando);
            }

            var config = ConfigForm.Config;

            // Validações gerais obrigatórias para movimentação
            if (string.IsNullOrWhiteSpace(config.CnpjDeclarante))
                camposFaltando.Add(MSG_CAMPO_CNPJ);

            if (string.IsNullOrWhiteSpace(config.CertThumbprint))
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);

            if (string.IsNullOrWhiteSpace(config.CertServidorThumbprint))
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);

            if (string.IsNullOrWhiteSpace(config.Periodo))
                camposFaltando.Add(MSG_CAMPO_PERIODO);

            if (string.IsNullOrWhiteSpace(config.DiretorioLotes))
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);

            if (config.PageSize <= 0)
                camposFaltando.Add("Campo: 'Page Size' na seção 'Configuração de Processamento' (deve ser maior que zero)");

            if (camposFaltando.Count > 0)
            {
                return MSG_CAMPOS_FALTANDO + string.Join("\n", camposFaltando);
            }

            return null;
        }

        private string ValidarDadosFechamento()
        {
            List<string> camposFaltando = new List<string>();

            if (ConfigForm == null)
            {
                camposFaltando.Add(MSG_ACESSE_CONFIG);
                camposFaltando.Add(MSG_CAMPO_CNPJ);
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);
                camposFaltando.Add(MSG_CAMPO_PERIODO);
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);
                camposFaltando.Add("Aba 'Dados de Fechamento': Preencha 'Data de Início'");
                camposFaltando.Add("Aba 'Dados de Fechamento': Preencha 'Data de Fim'");
                return string.Join("\n", camposFaltando);
            }

            if (ConfigForm.Config == null)
            {
                camposFaltando.Add(MSG_CONFIG_NAO_INICIALIZADA);
                camposFaltando.Add(MSG_CAMPO_CNPJ);
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);
                camposFaltando.Add(MSG_CAMPO_PERIODO);
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);
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
                camposFaltando.Add(MSG_CAMPO_CNPJ);

            if (string.IsNullOrWhiteSpace(config.CertThumbprint))
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);

            if (string.IsNullOrWhiteSpace(config.CertServidorThumbprint))
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);

            if (string.IsNullOrWhiteSpace(config.Periodo))
                camposFaltando.Add(MSG_CAMPO_PERIODO);

            if (string.IsNullOrWhiteSpace(config.DiretorioLotes))
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);

            // Validações específicas de fechamento
            if (string.IsNullOrWhiteSpace(dadosFechamento.DtInicio))
                camposFaltando.Add("Campo: 'Data de Início' na aba 'Dados de Fechamento'");

            if (string.IsNullOrWhiteSpace(dadosFechamento.DtFim))
                camposFaltando.Add("Campo: 'Data de Fim' na aba 'Dados de Fechamento'");

            if (camposFaltando.Count > 0)
            {
                return MSG_CAMPOS_FALTANDO + string.Join("\n", camposFaltando);
            }

            return null;
        }

        private string ValidarDadosCadastroDeclarante()
        {
            List<string> camposFaltando = new List<string>();

            if (ConfigForm == null)
            {
                camposFaltando.Add("Acesse a aba 'Configuração' e preencha todos os campos obrigatórios");
                camposFaltando.Add(MSG_CAMPO_CNPJ);
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);
                camposFaltando.Add("Aba 'Cadastro Declarante': Preencha 'Nome (Razão Social)'");
                camposFaltando.Add("Aba 'Cadastro Declarante': Preencha 'Endereço Livre'");
                camposFaltando.Add("Aba 'Cadastro Declarante': Preencha 'Município'");
                camposFaltando.Add("Aba 'Cadastro Declarante': Preencha 'UF'");
                camposFaltando.Add("Aba 'Cadastro Declarante': Preencha 'CEP'");
                return string.Join("\n", camposFaltando);
            }

            if (ConfigForm.Config == null)
            {
                camposFaltando.Add("Configuração geral não foi inicializada. Preencha os seguintes campos:");
                camposFaltando.Add(MSG_CAMPO_CNPJ);
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);
                return string.Join("\n", camposFaltando);
            }

            var config = ConfigForm.Config;
            var dadosCadastro = ConfigForm.DadosCadastroDeclarante;

            if (dadosCadastro == null)
            {
                camposFaltando.Add("Dados de cadastro de declarante não foram configurados. Preencha os seguintes campos:");
                camposFaltando.Add("Aba 'Cadastro Declarante': Preencha 'Nome (Razão Social)'");
                camposFaltando.Add("Aba 'Cadastro Declarante': Preencha 'Endereço Livre'");
                camposFaltando.Add("Aba 'Cadastro Declarante': Preencha 'Município'");
                camposFaltando.Add("Aba 'Cadastro Declarante': Preencha 'UF'");
                camposFaltando.Add("Aba 'Cadastro Declarante': Preencha 'CEP'");
                return string.Join("\n", camposFaltando);
            }

            // Validações gerais
            if (string.IsNullOrWhiteSpace(config.CnpjDeclarante))
                camposFaltando.Add(MSG_CAMPO_CNPJ);

            if (string.IsNullOrWhiteSpace(config.CertThumbprint))
                camposFaltando.Add(MSG_CAMPO_CERT_ASSINATURA);

            if (string.IsNullOrWhiteSpace(config.CertServidorThumbprint))
                camposFaltando.Add(MSG_CAMPO_CERT_SERVIDOR);

            if (string.IsNullOrWhiteSpace(config.DiretorioLotes))
                camposFaltando.Add(MSG_CAMPO_DIRETORIO);

            // Validações específicas de cadastro de declarante
            if (string.IsNullOrWhiteSpace(dadosCadastro.Nome))
                camposFaltando.Add("Campo: 'Nome (Razão Social)' na aba 'Cadastro Declarante'");

            if (string.IsNullOrWhiteSpace(dadosCadastro.EnderecoLivre))
                camposFaltando.Add("Campo: 'Endereço Livre' na aba 'Cadastro Declarante'");

            if (string.IsNullOrWhiteSpace(dadosCadastro.Municipio))
                camposFaltando.Add("Campo: 'Município (Código IBGE)' na aba 'Cadastro Declarante'");

            if (string.IsNullOrWhiteSpace(dadosCadastro.UF))
                camposFaltando.Add("Campo: 'UF' na aba 'Cadastro Declarante'");

            if (string.IsNullOrWhiteSpace(dadosCadastro.CEP))
                camposFaltando.Add("Campo: 'CEP' na aba 'Cadastro Declarante'");

            if (dadosCadastro.PaisResid == null || dadosCadastro.PaisResid.Count == 0 || !dadosCadastro.PaisResid.Contains("BR"))
                camposFaltando.Add("Campo: 'País Residência Fiscal' na aba 'Cadastro Declarante' (deve conter 'BR')");

            if (camposFaltando.Count > 0)
            {
                return MSG_CAMPOS_FALTANDO + string.Join("\n", camposFaltando);
            }

            return null;
        }

        /// <summary>
        /// Converte datas (DtInicio e DtFim) em período YYYYMM
        /// </summary>
        private string ConverterDatasParaPeriodo(string dtInicio, string dtFim)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(dtInicio) || string.IsNullOrWhiteSpace(dtFim))
                    return null;

                DateTime inicio = DateTime.Parse(dtInicio, CULTURE_INFO_PT_BR);
                DateTime fim = DateTime.Parse(dtFim, CULTURE_INFO_PT_BR);

                int ano = inicio.Year;
                int mesInicio = inicio.Month;
                int mesFim = fim.Month;

                // Determinar período baseado nas datas
                // Primeiro semestre: Jan-Jun (meses 1-6)
                // Segundo semestre: Jul-Dez (meses 7-12)
                if (mesInicio >= 1 && mesInicio <= 6 && mesFim >= 1 && mesFim <= 6)
                {
                    // Primeiro semestre - usar 01 ou 06
                    return $"{ano}01";
                }
                else if (mesInicio >= 7 && mesInicio <= 12 && mesFim >= 7 && mesFim <= 12)
                {
                    // Segundo semestre - usar 02 ou 12
                    return $"{ano}02";
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Valida se abertura, movimentação e fechamento estão no mesmo período
        /// </summary>
        private string ValidarPeriodosCompletos()
        {
            if (ConfigForm == null || ConfigForm.Config == null)
                return "Configuração não inicializada.";

            var config = ConfigForm.Config;
            var dadosAbertura = ConfigForm.DadosAbertura;
            var dadosFechamento = ConfigForm.DadosFechamento;

            if (dadosAbertura == null)
                return "Dados de abertura não configurados.";

            if (dadosFechamento == null)
                return "Dados de fechamento não configurados.";

            if (string.IsNullOrWhiteSpace(config.Periodo))
                return "Período não configurado na configuração geral.";

            // Converter datas de abertura em período
            string periodoAbertura = ConverterDatasParaPeriodo(dadosAbertura.DtInicio, dadosAbertura.DtFim);
            if (periodoAbertura == null)
                return "Não foi possível determinar o período das datas de abertura. Verifique se as datas estão corretas.";

            // Converter datas de fechamento em período
            string periodoFechamento = ConverterDatasParaPeriodo(dadosFechamento.DtInicio, dadosFechamento.DtFim);
            if (periodoFechamento == null)
                return "Não foi possível determinar o período das datas de fechamento. Verifique se as datas estão corretas.";

            // Comparar períodos
            string periodoMovimentacao = config.Periodo.Trim();

            List<string> periodosDiferentes = new List<string>();

            if (periodoAbertura != periodoMovimentacao)
                periodosDiferentes.Add($"Abertura: {periodoAbertura} (diferente de Movimentação: {periodoMovimentacao})");

            if (periodoFechamento != periodoMovimentacao)
                periodosDiferentes.Add($"Fechamento: {periodoFechamento} (diferente de Movimentação: {periodoMovimentacao})");

            if (periodoAbertura != periodoFechamento)
                periodosDiferentes.Add($"Abertura: {periodoAbertura} (diferente de Fechamento: {periodoFechamento})");

            if (periodosDiferentes.Count > 0)
            {
                return "Os períodos não estão alinhados!\n\n" +
                       $"Período Configurado (Movimentação): {periodoMovimentacao}\n" +
                       $"Período Abertura: {periodoAbertura}\n" +
                       $"Período Fechamento: {periodoFechamento}\n\n" +
                       "Por favor, corrija os períodos para que todos sejam iguais antes de usar o processamento completo.\n\n" +
                       string.Join("\n", periodosDiferentes);
            }

            return null;
        }

        private async Task ProcessarAbertura()
        {
            try
            {
                IniciarProcessamento();
                AtualizarEtapa("Iniciando processamento de abertura...");
                status.TotalLotes = 1;

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
                    string ambienteStr = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : AMBIENTE_HOMOLOG;
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
                    persistenceService.RegistrarLogLote(idLoteBanco, LOG_GERACAO, $"XML gerado: {Path.GetFileName(arquivoXml)}");
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
                string arquivoAssinado = arquivoXml.Replace(".xml", SUFIXO_ASSINADO);
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
                        persistenceService.AtualizarLote(idLoteBanco, STATUS_ASSINADO);
                        persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ASSINATURA, $"XML assinado: {Path.GetFileName(arquivoAssinado)}");
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
                        persistenceService.AtualizarLote(idLoteBanco, STATUS_CRIPTOGRAFADO);
                        persistenceService.RegistrarLogLote(idLoteBanco, STATUS_CRIPTOGRAFIA, $"XML criptografado: {Path.GetFileName(arquivoCriptografado)}");
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
                        
                        // SEMPRE tentar extrair o protocolo, independentemente do código de resposta
                        string protocoloFinal = ExtrairProtocolo(resposta);
                        
                        // Determinar status e mensagem baseado no código de resposta
                        string statusLote = "ENVIADO_COM_RESPOSTA";
                        string mensagemStatus = "";
                        
                        if (resposta.CodigoResposta == 1)
                        {
                            statusLote = STATUS_ENVIADO;
                            mensagemStatus = "Lote enviado com sucesso";
                        }
                        else if (resposta.CodigoResposta == 7)
                        {
                            statusLote = STATUS_REJEITADO;
                            mensagemStatus = "Lote REJEITADO";
                            AdicionarLog($"✗ Lote de abertura REJEITADO - Código: {resposta.CodigoResposta}");
                            AdicionarLog($"  Descrição: {resposta.Descricao}");
                            
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
                            AdicionarLog($"⚠ Lote de abertura - Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}");
                        }
                        
                        // Se não encontrou protocolo, aguardar um pouco e tentar novamente
                        if (string.IsNullOrEmpty(protocoloFinal))
                        {
                            AdicionarLog($"⚠ Aguardando resposta do servidor para obter protocolo...");
                            System.Threading.Thread.Sleep(2000);
                            protocoloFinal = ExtrairProtocolo(resposta);
                            
                            if (string.IsNullOrEmpty(protocoloFinal))
                            {
                                statusLote = "ENVIADO_SEM_PROTOCOLO";
                                mensagemStatus = "Protocolo não retornado pelo servidor";
                                AdicionarLog($"⚠ ATENÇÃO: Protocolo não foi retornado pelo servidor!");
                            }
                        }
                        
                        // Se encontrou protocolo, atualizar objeto resposta e adicionar à lista
                        if (!string.IsNullOrEmpty(protocoloFinal))
                        {
                            resposta.Protocolo = protocoloFinal;
                            
                            if (!status.ProtocolosEnviados.Contains(protocoloFinal))
                            {
                                status.ProtocolosEnviados.Add(protocoloFinal);
                            }
                            
                            AdicionarLog($"✓ Protocolo obtido: {protocoloFinal}");
                            
                            // Atualizar protocolo no lote já registrado (sistema antigo)
                                ProtocoloPersistenciaService.RegistrarProtocolo(
                                    TipoLote.Abertura, 
                                    arquivoCriptografado, 
                                protocoloFinal,
                                config.Periodo,
                                quantidadeEventos
                            );
                        }
                        
                        // Atualizar lote no banco com resposta e protocolo (se encontrado)
                        if (idLoteBanco > 0)
                        {
                            try
                            {
                                var persistenceService = new EfinanceiraDatabasePersistenceService();
                                string xmlResposta = resposta.XmlCompleto ?? "";
                                persistenceService.AtualizarLote(
                                    idLoteBanco,
                                    statusLote,
                                    protocoloFinal,
                                    resposta.CodigoResposta,
                                    resposta.Descricao,
                                    xmlResposta,
                                    null,
                                    null,
                                    null,
                                    DateTime.Now,
                                    null,
                                    string.IsNullOrEmpty(protocoloFinal) ? mensagemStatus : null
                                );
                                
                                if (!string.IsNullOrEmpty(protocoloFinal))
                                {
                                    persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ENVIO, 
                                        $"Lote enviado. Protocolo: {protocoloFinal}, Código: {resposta.CodigoResposta}");
                                }
                                else
                                {
                                    persistenceService.RegistrarLogLote(idLoteBanco, "ENVIO_AVISO", 
                                        $"Lote enviado, mas protocolo não retornado. Código: {resposta.CodigoResposta}");
                                }
                            }
                            catch (Exception exDb)
                            {
                                AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                            }
                        }
                        
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
                        
                        // Exibir MessageBox com resultado
                        if (!string.IsNullOrEmpty(protocoloFinal))
                        {
                            MessageBox.Show(
                                $"Lote de abertura enviado com sucesso!\n\n" +
                                $"PROTOCOLO: {protocoloFinal}\n\n" +
                                $"Código de Resposta: {resposta.CodigoResposta}\n" +
                                $"Descrição: {resposta.Descricao}\n\n" +
                                $"Este protocolo foi salvo e pode ser consultado na aba 'Consulta'.",
                                "Envio Concluído",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information
                            );
                        }
                        else
                        {
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
                        
                        AtualizarEstatisticas();
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

                FinalizarProcessamento();
                ReabilitarControlesAposProcessamento();
            }
            catch (Exception ex)
            {
                AdicionarLog($"ERRO: {ex.Message}");
                status.LotesComErro = 1;
                FinalizarProcessamento(mensagemErro: ex.Message);
                ReabilitarControlesAposProcessamento();
                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show($"Erro ao processar: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }

        private async Task ProcessarMovimentacao()
        {
            try
            {
                // Verificar se está em modo completo
                bool modoCompleto = status.ModoCompleto;
                
                if (!modoCompleto)
                {
                    IniciarProcessamento();
                }
                
                AtualizarEtapa("Iniciando processamento de movimentação...");
                
                if (!modoCompleto)
                {
                    status.TotalLotes = 0;
                }
                else
                {
                    // Em modo completo, manter contadores mas resetar apenas os de movimentação
                    status.LotesMovimentacaoProcessados = 0;
                    status.TotalLotesMovimentacao = 0;
                    status.TemposLotesMovimentacao.Clear();
                    status.TemposLotesMovimentacao.Add(DateTime.Now);
                }

                var config = ConfigForm.Config;

                // Validar e calcular período
                string periodoStr = config.Periodo;
                if (string.IsNullOrWhiteSpace(periodoStr))
                {
                    throw new ArgumentException("Período não configurado. Configure o campo 'Período' na seção Configuração Geral.", nameof(config.Periodo));
                }

                // Validar formato do período
                if (!EfinanceiraPeriodoUtil.ValidarPeriodo(periodoStr))
                {
                    throw new ArgumentException($"Período inválido: {periodoStr}. Deve estar no formato YYYYMM onde MM deve ser:\n" +
                        $"  • 01 ou 06 = Primeiro semestre (Janeiro a Junho)\n" +
                        $"  • 02 ou 12 = Segundo semestre (Julho a Dezembro)\n" +
                        $"Exemplos: 202301 (Jan-Jun/2023) ou 202302 (Jul-Dez/2023)", nameof(config.Periodo));
                }

                // Calcular datas do período semestral
                var (dataInicio, dataFim) = EfinanceiraPeriodoUtil.CalcularPeriodoSemestral(periodoStr);
                
                AdicionarLog($"Período configurado: {periodoStr}");
                AdicionarLog($"Datas calculadas pelo EfinanceiraPeriodoUtil: {dataInicio} a {dataFim}");
                
                DateTime dtInicio = DateTime.Parse(dataInicio, CULTURE_INFO_PT_BR);
                DateTime dtFim = DateTime.Parse(dataFim, CULTURE_INFO_PT_BR);
                
                // Indicar que está consultando o banco
                AtualizarEtapa("Consultando banco de dados...");
                this.Invoke((MethodInvoker)delegate
                {
                    if (progressBarGeral.Style != ProgressBarStyle.Marquee)
                    {
                        progressBarGeral.Style = ProgressBarStyle.Marquee;
                        progressBarGeral.MarqueeAnimationSpeed = 30;
                    }
                    lblProgressoGeral.Text = "Consultando banco de dados...";
                });
                
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
                    throw new InvalidOperationException($"Meses inválidos calculados: Mês Inicial={mesInicial}, Mês Final={mesFinal}. Verifique o período configurado.");
                }

                // Para períodos semestrais, validar:
                // - Se mesInicial = 1, mesFinal deve ser 6 (Jan-Jun)
                // - Se mesInicial = 7, mesFinal deve ser 12 (Jul-Dez)
                if ((mesInicial == 1 && mesFinal != 6) || (mesInicial == 7 && mesFinal != 12))
                {
                    throw new InvalidOperationException($"Período semestral inválido: Mês Inicial={mesInicial}, Mês Final={mesFinal}. " +
                        $"Esperado: 1-6 (Jan-Jun) ou 7-12 (Jul-Dez). Verifique o período '{periodoStr}' configurado.");
                }

                // Garantir que mesInicial <= mesFinal (para a query SQL funcionar corretamente)
                if (mesInicial > mesFinal)
                {
                    throw new InvalidOperationException($"Erro: Mês Inicial ({mesInicial}) é maior que Mês Final ({mesFinal}). " +
                        $"Isso não é válido para um período semestral. Período configurado: {periodoStr}");
                }

                // Testar conexão com banco
                AtualizarEtapa("Testando conexão com banco de dados...");
                var dbService = new EfinanceiraDatabaseService();
                try
                {
                    if (!dbService.TestarConexao())
                    {
                        FinalizarProcessamento(mensagemErro: "Não foi possível conectar ao banco de dados");
                        throw new InvalidOperationException("Não foi possível conectar ao banco de dados. Verifique as credenciais.");
                    }
                    AdicionarLog("Conexão com banco de dados estabelecida.");
                }
                catch (Exception exConexao)
                {
                    FinalizarProcessamento(mensagemErro: $"Erro ao conectar: {exConexao.Message}");
                    AdicionarLog($"✗ Erro ao conectar ao banco de dados: {exConexao.Message}");
                    throw new InvalidOperationException($"Erro ao conectar ao banco de dados: {exConexao.Message}", exConexao);
                }

                // Buscar pessoas com contas (paginado)
                AtualizarEtapa("Buscando dados do banco de dados...");
                int pageSize = config.PageSize;
                int offset = config.OffsetRegistros;
                int maxLotes = config.MaxLotes ?? int.MaxValue;
                int lotesGerados = 0;
                int eventosOffset = config.EventoOffset - 1; // Ajustar para índice base 0

                while (lotesGerados < maxLotes && !cancelarProcessamento)
                {
                    AtualizarEtapa($"Consultando banco de dados (página {lotesGerados + 1})...");
                    this.Invoke((MethodInvoker)delegate
                    {
                        if (progressBarGeral.Style != ProgressBarStyle.Marquee)
                        {
                            progressBarGeral.Style = ProgressBarStyle.Marquee;
                            progressBarGeral.MarqueeAnimationSpeed = 30;
                        }
                        lblProgressoGeral.Text = $"Consultando banco de dados (página {lotesGerados + 1})...";
                    });
                    
                    AdicionarLog($"Buscando página: offset={offset}, limit={pageSize}...");
                    
                    List<DadosPessoaConta> pessoas;
                    try
                    {
                        pessoas = dbService.BuscarPessoasComContas(ano, mesInicial, mesFinal, pageSize, offset);
                    }
                    catch (Exception exConsulta)
                    {
                        FinalizarProcessamento(mensagemErro: $"Erro ao consultar banco: {exConsulta.Message}");
                        AdicionarLog($"✗ Erro ao consultar banco de dados: {exConsulta.Message}");
                        this.Invoke((MethodInvoker)delegate
                        {
                            MessageBox.Show(
                                $"Erro ao consultar banco de dados:\n\n{exConsulta.Message}\n\n" +
                                "O processamento foi interrompido.",
                                "Erro na Consulta ao Banco",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        });
                        throw new InvalidOperationException($"Erro ao consultar banco de dados: {exConsulta.Message}", exConsulta);
                    }
                    
                    // Atualizar barra de progresso após consulta
                    this.Invoke((MethodInvoker)delegate
                    {
                        if (progressBarGeral.Style == ProgressBarStyle.Marquee && status.TotalLotes > 0)
                        {
                            progressBarGeral.Style = ProgressBarStyle.Continuous;
                        }
                    });
                    
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
                        
                        // Rastrear tempo se estiver em modo completo
                        DateTime inicioLote = DateTime.Now;
                        if (modoCompleto)
                        {
                            status.TotalLotesMovimentacao = lotesGerados;
                            if (status.TemposLotesMovimentacao.Count == 0)
                            {
                                status.TemposLotesMovimentacao.Add(inicioLote);
                            }
                        }

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
                            string ambienteStr = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : AMBIENTE_HOMOLOG;
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
                            persistenceService.RegistrarLogLote(idLoteBanco, LOG_GERACAO, $"XML gerado: {Path.GetFileName(arquivoXml)}");
                            
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
                        string arquivoAssinado = arquivoXml.Replace(".xml", SUFIXO_ASSINADO);
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
                                            System.Xml.XmlNodeList protocoloList = doc.GetElementsByTagName(PROTOCOLO_ENVIO);
                                            if (protocoloList != null && protocoloList.Count > 0)
                                            {
                                                protocoloFinal = protocoloList[0].InnerText.Trim();
                                            }
                                            else
                                            {
                                                // Tentar com XPath (com namespace)
                                                System.Xml.XmlNamespaceManager nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                                                nsmgr.AddNamespace("ns", NAMESPACE_ENVIO_LOTE);
                                                
                                                System.Xml.XmlNode protocoloNode = doc.SelectSingleNode(PROTOCOLO_XPATH)
                                                    ?? doc.SelectSingleNode(PROTOCOLO_NS_XPATH, nsmgr)
                                                    ?? doc.SelectSingleNode(PROTOCOLO_SIMPLE_XPATH)
                                                    ?? doc.SelectSingleNode(PROTOCOLO_NS_SIMPLE_XPATH, nsmgr);
                                                
                                                if (protocoloNode != null)
                                                {
                                                    protocoloFinal = protocoloNode.InnerText.Trim();
                                                }
                                                else
                                                {
                                                    // Tentar buscar por numeroProtocolo (como no método extrairProtocolo do Java)
                                                    System.Xml.XmlNodeList numeroProtocoloList = doc.GetElementsByTagName(NUMERO_PROTOCOLO);
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
                                                    STATUS_ENVIADO,
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
                                                persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ENVIO, $"Lote enviado com sucesso. Protocolo: {protocoloFinal}");
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
                                        
                                        // Rastrear tempo final do lote se estiver em modo completo
                                        if (modoCompleto)
                                        {
                                            status.LotesMovimentacaoProcessados++;
                                            DateTime fimLote = DateTime.Now;
                                            TimeSpan tempoLote = fimLote - inicioLote;
                                            if (status.TemposLotesMovimentacao.Count > 0)
                                            {
                                                // Calcular tempo médio atualizado
                                                TimeSpan tempoTotal = fimLote - status.TemposLotesMovimentacao[0];
                                                status.TempoMedioPorLote = TimeSpan.FromMilliseconds(tempoTotal.TotalMilliseconds / status.LotesMovimentacaoProcessados);
                                            }
                                            AtualizarEstatisticas();
                                        }
                                    }
                                    else
                                    {
                                        // Se não recebeu protocolo, aguardar e tentar novamente ou informar erro
                                        AdicionarLog($"⚠ ATENÇÃO: Lote {lotesGerados} enviado mas protocolo não foi retornado!");
                                        
                                        // Rastrear tempo mesmo em caso de erro
                                        if (modoCompleto)
                                        {
                                            status.LotesMovimentacaoProcessados++;
                                            DateTime fimLote = DateTime.Now;
                                            if (status.TemposLotesMovimentacao.Count > 0)
                                            {
                                                TimeSpan tempoTotal = fimLote - status.TemposLotesMovimentacao[0];
                                                status.TempoMedioPorLote = TimeSpan.FromMilliseconds(tempoTotal.TotalMilliseconds / status.LotesMovimentacaoProcessados);
                                            }
                                            AtualizarEstatisticas();
                                        }
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
                                                STATUS_REJEITADO,
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
                                            nsmgr.AddNamespace("ns", NAMESPACE_ENVIO_LOTE);
                                            
                                            System.Xml.XmlNode protocoloNode = doc.SelectSingleNode(PROTOCOLO_XPATH)
                                                ?? doc.SelectSingleNode(PROTOCOLO_NS_XPATH, nsmgr)
                                                ?? doc.SelectSingleNode(PROTOCOLO_SIMPLE_XPATH)
                                                ?? doc.SelectSingleNode(PROTOCOLO_NS_SIMPLE_XPATH, nsmgr);
                                            
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
                                                persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ENVIO, 
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

                if (!status.ModoCompleto)
                {
                    FinalizarProcessamento();
                }
                ReabilitarControlesAposProcessamento();
            }
            catch (Exception ex)
            {
                AdicionarLog($"ERRO: {ex.Message}");
                if (ex.InnerException != null)
                {
                    AdicionarLog($"Detalhes: {ex.InnerException.Message}");
                }
                status.LotesComErro++;
                if (!status.ModoCompleto)
                {
                    FinalizarProcessamento(mensagemErro: ex.Message);
                }
                ReabilitarControlesAposProcessamento();
                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show($"Erro ao processar movimentação: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }

        private async Task ProcessarFechamento()
        {
            try
            {
                IniciarProcessamento();
                AtualizarEtapa("Iniciando processamento de fechamento...");
                status.TotalLotes = 1;

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
                    string ambienteStr = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : AMBIENTE_HOMOLOG;
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
                    persistenceService.RegistrarLogLote(idLoteBanco, LOG_GERACAO, $"XML gerado: {Path.GetFileName(arquivoXml)}");
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
                string arquivoAssinado = arquivoXml.Replace(".xml", SUFIXO_ASSINADO);
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
                        persistenceService.AtualizarLote(idLoteBanco, STATUS_ASSINADO);
                        persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ASSINATURA, $"XML assinado: {Path.GetFileName(arquivoAssinado)}");
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
                        persistenceService.AtualizarLote(idLoteBanco, STATUS_CRIPTOGRAFADO);
                        persistenceService.RegistrarLogLote(idLoteBanco, STATUS_CRIPTOGRAFIA, $"XML criptografado: {Path.GetFileName(arquivoCriptografado)}");
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
                                        nsmgr.AddNamespace("ns", NAMESPACE_ENVIO_LOTE);
                                        
                                        System.Xml.XmlNode protocoloNode = doc.SelectSingleNode(PROTOCOLO_XPATH)
                                            ?? doc.SelectSingleNode(PROTOCOLO_NS_XPATH, nsmgr)
                                            ?? doc.SelectSingleNode(PROTOCOLO_SIMPLE_XPATH)
                                            ?? doc.SelectSingleNode(PROTOCOLO_NS_SIMPLE_XPATH, nsmgr);
                                        
                                        if (protocoloNode != null)
                                        {
                                            protocoloFinal = protocoloNode.InnerText.Trim();
                                        }
                                        else
                                        {
                                            // Tentar buscar por numeroProtocolo (como no método extrairProtocolo do Java)
                                            System.Xml.XmlNodeList numeroProtocoloList = doc.GetElementsByTagName(NUMERO_PROTOCOLO);
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
                                            STATUS_ENVIADO,
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
                                        persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ENVIO, $"Lote enviado com sucesso. Protocolo: {protocoloFinal}");
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
                                        STATUS_REJEITADO,
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
                                        nsmgr.AddNamespace("ns", NAMESPACE_ENVIO_LOTE);
                                        
                                        System.Xml.XmlNode protocoloNode = doc.SelectSingleNode(PROTOCOLO_XPATH)
                                            ?? doc.SelectSingleNode(PROTOCOLO_NS_XPATH, nsmgr)
                                            ?? doc.SelectSingleNode(PROTOCOLO_SIMPLE_XPATH)
                                            ?? doc.SelectSingleNode(PROTOCOLO_NS_SIMPLE_XPATH, nsmgr);
                                        
                                        if (protocoloNode != null)
                                        {
                                            protocoloExtraido = protocoloNode.InnerText.Trim();
                                        }
                                        else
                                        {
                                            // Tentar buscar por numeroProtocolo (como no método extrairProtocolo do Java)
                                            System.Xml.XmlNodeList numeroProtocoloList = doc.GetElementsByTagName(NUMERO_PROTOCOLO);
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
                                        persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ENVIO, 
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

                FinalizarProcessamento();
                ReabilitarControlesAposProcessamento();
            }
            catch (Exception ex)
            {
                AdicionarLog($"ERRO: {ex.Message}");
                status.LotesComErro = 1;
                FinalizarProcessamento(mensagemErro: ex.Message);
                ReabilitarControlesAposProcessamento();
                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show($"Erro ao processar: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }

        /// <summary>
        /// Processa tudo automaticamente: Abertura -> Movimentações -> Fechamento
        /// Sequência: Processar Abertura -> Enviar Abertura -> Processar Movimentações -> Enviar Movimentações -> Processar Fechamento -> Enviar Fechamento
        /// </summary>
        private async Task ProcessarCompleto()
        {
            try
            {
                IniciarProcessamento();
                status.ModoCompleto = true;
                status.StatusEtapa = 0; // Aguardando início
                status.AberturaFinalizada = false;
                status.LotesMovimentacaoProcessados = 0;
                status.TotalLotesMovimentacao = 0;
                status.TemposLotesMovimentacao.Clear();
                mapArquivoParaIdLote.Clear();
                arquivoAberturaCriptografado = null;
                idLoteAberturaBanco = 0;

                AdicionarLog("════════════════════════════════════════");
                AdicionarLog("INICIANDO PROCESSAMENTO COMPLETO");
                AdicionarLog("════════════════════════════════════════");

                // ETAPA 1: PROCESSAR ABERTURA
                AdicionarLog("\n>>> ETAPA 1: PROCESSANDO ABERTURA <<<");
                status.StatusEtapa = 1; // Processando Abertura
                AtualizarEstatisticas();

                await ProcessarAberturaCompleto();

                if (cancelarProcessamento)
                {
                    AdicionarLog("Processamento cancelado pelo usuário.");
                    status.ModoCompleto = false;
                    FinalizarProcessamento(cancelado: true);
                    ReabilitarControlesAposProcessamento();
                    return;
                }
                
                AdicionarLog("✓ Abertura processada com sucesso!");

                // ETAPA 2: ENVIAR ABERTURA E AGUARDAR SUCESSO
                if (!chkApenasProcessar.Checked)
                {
                    AdicionarLog("\n>>> ETAPA 2: ENVIANDO ABERTURA <<<");
                    status.StatusEtapa = 2; // Abertura Enviada
                    AtualizarEstatisticas();
                    
                    if (string.IsNullOrEmpty(arquivoAberturaCriptografado))
                    {
                        throw new Exception("Arquivo de abertura criptografado não foi gerado. Não é possível enviar.");
                    }
                    
                    if (!File.Exists(arquivoAberturaCriptografado))
                    {
                        throw new Exception($"Arquivo de abertura não encontrado: {arquivoAberturaCriptografado}");
                    }
                    
                    AdicionarLog($"Arquivo pronto para envio: {Path.GetFileName(arquivoAberturaCriptografado)} ({new FileInfo(arquivoAberturaCriptografado).Length} bytes)");
                    
                    bool aberturaEnviada = await EnviarAberturaCompleto();

                    if (!aberturaEnviada)
                    {
                        throw new Exception("Falha ao enviar abertura. Processamento completo interrompido.");
                    }
                    
                    AdicionarLog("✓ Abertura enviada com sucesso!");
                }
                else
                {
                    AdicionarLog("\n>>> ETAPA 2: ENVIO DE ABERTURA PULADO (modo 'Apenas Processar' ativo) <<<");
                    status.StatusEtapa = 2; // Abertura processada (sem envio)
                    AtualizarEstatisticas();
                    AdicionarLog("✓ Abertura processada (envio pulado - modo 'Apenas Processar')");
                }

                if (cancelarProcessamento)
                {
                    AdicionarLog("Processamento cancelado pelo usuário.");
                    status.ModoCompleto = false;
                    FinalizarProcessamento(cancelado: true);
                    ReabilitarControlesAposProcessamento();
                    return;
                }

                status.AberturaFinalizada = true;
                AtualizarEstatisticas();

                // ETAPA 3: VERIFICAR SE ABERTURA FOI ENVIADA ANTES DE PROCESSAR MOVIMENTAÇÕES
                if (!chkApenasProcessar.Checked)
                {
                    AdicionarLog("\n>>> VERIFICANDO SE ABERTURA FOI ENVIADA <<<");
                    var config = ConfigForm.Config;
                    string ambienteStr = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : AMBIENTE_HOMOLOG;
                    
                    AtualizarEtapa("Verificando status da abertura no banco de dados...");
                    this.Invoke((MethodInvoker)delegate
                    {
                        if (progressBarGeral.Style != ProgressBarStyle.Marquee)
                        {
                            progressBarGeral.Style = ProgressBarStyle.Marquee;
                            progressBarGeral.MarqueeAnimationSpeed = 30;
                        }
                        lblProgressoGeral.Text = "Verificando status da abertura...";
                    });
                    
                    try
                    {
                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                        bool aberturaEnviada = persistenceService.VerificarAberturaEnviadaParaPeriodo(config.Periodo, ambienteStr);
                        
                        if (!aberturaEnviada)
                        {
                            // Aguardar um pouco e verificar novamente (pode ser que ainda esteja processando)
                            AdicionarLog("Aguardando confirmação da abertura...");
                            await Task.Delay(2000); // Aguardar 2 segundos
                            aberturaEnviada = persistenceService.VerificarAberturaEnviadaParaPeriodo(config.Periodo, ambienteStr);
                        }
                        
                        // Atualizar barra de progresso após consulta
                        this.Invoke((MethodInvoker)delegate
                        {
                            if (progressBarGeral.Style == ProgressBarStyle.Marquee && status.TotalLotes > 0)
                            {
                                progressBarGeral.Style = ProgressBarStyle.Continuous;
                            }
                        });
                        
                        if (!aberturaEnviada)
                        {
                            AdicionarLog("⚠ AVISO: Abertura ainda não foi confirmada como enviada no banco.");
                            AdicionarLog("Continuando processamento de movimentações...");
                            AdicionarLog("Se ocorrer erro 'Não existe e-Financeira aberta', verifique se a abertura foi realmente enviada e aceita.");
                        }
                        else
                        {
                            AdicionarLog("✓ Abertura confirmada como enviada. Prosseguindo com movimentações...");
                        }
                    }
                    catch (Exception exConsulta)
                    {
                        FinalizarProcessamento(mensagemErro: $"Erro ao verificar abertura: {exConsulta.Message}");
                        AdicionarLog($"✗ Erro ao verificar abertura no banco de dados: {exConsulta.Message}");
                        this.Invoke((MethodInvoker)delegate
                        {
                            MessageBox.Show(
                                $"Erro ao verificar abertura no banco de dados:\n\n{exConsulta.Message}\n\n" +
                                "O processamento foi interrompido.",
                                "Erro na Consulta ao Banco",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        });
                        throw new InvalidOperationException($"Erro ao verificar abertura no banco de dados: {exConsulta.Message}", exConsulta);
                    }
                }

                // ETAPA 3: PROCESSAR TODOS OS LOTES DE MOVIMENTAÇÃO
                AdicionarLog("\n>>> ETAPA 3: PROCESSANDO LOTES DE MOVIMENTAÇÃO FINANCEIRA <<<");
                status.StatusEtapa = 3; // Processando lotes de movimentação financeira
                DateTime inicioMovimentacoes = DateTime.Now;
                status.TemposLotesMovimentacao.Clear();
                status.TemposLotesMovimentacao.Add(inicioMovimentacoes);
                AtualizarEstatisticas();

                List<string> arquivosMovimentacao = await ProcessarMovimentacaoCompleto();

                if (cancelarProcessamento)
                {
                    AdicionarLog("Processamento cancelado pelo usuário.");
                    status.ModoCompleto = false;
                    FinalizarProcessamento(cancelado: true);
                    ReabilitarControlesAposProcessamento();
                    return;
                }

                AdicionarLog($"✓ {arquivosMovimentacao.Count} lote(s) de movimentação processado(s)!");

                // ETAPA 4: ENVIAR TODOS OS LOTES DE MOVIMENTAÇÃO
                if (arquivosMovimentacao.Count > 0 && !chkApenasProcessar.Checked)
                {
                    AdicionarLog("\n>>> ETAPA 4: ENVIANDO LOTES DE MOVIMENTAÇÃO FINANCEIRA <<<");
                    status.StatusEtapa = 4; // Enviando lotes de movimentação financeira
                    AtualizarEstatisticas();

                    await EnviarMovimentacoesCompleto(arquivosMovimentacao);

                    if (cancelarProcessamento)
                    {
                        AdicionarLog("Processamento cancelado pelo usuário.");
                        status.ModoCompleto = false;
                        FinalizarProcessamento(cancelado: true);
                        ReabilitarControlesAposProcessamento();
                        return;
                    }

                    // Calcular tempo médio por lote
                    if (status.TotalLotesMovimentacao > 0 && status.LotesMovimentacaoProcessados > 0)
                    {
                        TimeSpan tempoTotalMovimentacoes = DateTime.Now - inicioMovimentacoes;
                        status.TempoMedioPorLote = TimeSpan.FromMilliseconds(tempoTotalMovimentacoes.TotalMilliseconds / status.LotesMovimentacaoProcessados);
                    }
                    
                    AdicionarLog($"✓ {arquivosMovimentacao.Count} lote(s) de movimentação enviado(s)!");
                }
                else if (chkApenasProcessar.Checked)
                {
                    AdicionarLog("Envio de movimentações não realizado (modo 'Apenas Processar' ativo).");
                }

                AtualizarEstatisticas();

                // ETAPA 5: PROCESSAR FECHAMENTO
                AdicionarLog("\n>>> ETAPA 5: PROCESSANDO LOTE DE FECHAMENTO <<<");
                status.StatusEtapa = 5; // Processando lote de fechamento
                AtualizarEstatisticas();

                string arquivoFechamento = await ProcessarFechamentoCompleto();

                if (cancelarProcessamento)
                {
                    AdicionarLog("Processamento cancelado pelo usuário.");
                    status.ModoCompleto = false;
                    FinalizarProcessamento(cancelado: true);
                    ReabilitarControlesAposProcessamento();
                    return;
                }

                if (string.IsNullOrEmpty(arquivoFechamento))
                {
                    throw new Exception("Erro: Arquivo de fechamento não foi gerado!");
                }
                
                AdicionarLog("✓ Fechamento processado!");

                // ETAPA 6: ENVIAR FECHAMENTO
                if (!string.IsNullOrEmpty(arquivoFechamento) && !chkApenasProcessar.Checked)
                {
                    AdicionarLog("\n>>> ETAPA 6: ENVIANDO FECHAMENTO <<<");
                    status.StatusEtapa = 6; // Fechamento Enviado
                    AtualizarEstatisticas();

                    await EnviarFechamentoCompleto(arquivoFechamento);

                    if (cancelarProcessamento)
                    {
                        AdicionarLog("Processamento cancelado pelo usuário.");
                        status.ModoCompleto = false;
                        FinalizarProcessamento(cancelado: true);
                        ReabilitarControlesAposProcessamento();
                        return;
                    }
                    
                    AdicionarLog("✓ Fechamento enviado com sucesso!");
                }
                else if (chkApenasProcessar.Checked)
                {
                    AdicionarLog("Envio de fechamento não realizado (modo 'Apenas Processar' ativo).");
                }

                status.ModoCompleto = false;
                FinalizarProcessamento();
                AtualizarEstatisticas();

                // Atualizar lista na aba Consulta
                if (ConsultaForm != null)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        ConsultaForm.AtualizarListaLotes();
                    });
                }

                AdicionarLog("\n════════════════════════════════════════");
                AdicionarLog("PROCESSAMENTO COMPLETO FINALIZADO!");
                AdicionarLog("════════════════════════════════════════");
                AdicionarLog($"Total de lotes processados: {status.TotalLotes}");
                AdicionarLog($"Lotes enviados: {status.LotesEnviados}");
                AdicionarLog($"Lotes com erro: {status.LotesComErro}");
                if (status.TempoMedioPorLote.HasValue)
                {
                    AdicionarLog($"Tempo médio por lote: {status.TempoMedioPorLote.Value:mm\\:ss}");
                }

                ReabilitarControlesAposProcessamento();
                this.Invoke((MethodInvoker)delegate
                {
                    string mensagemEstatisticas = "Processamento completo finalizado com sucesso!\n\n";
                    mensagemEstatisticas += $"════════════════════════════════════════\n";
                    mensagemEstatisticas += $"ESTATÍSTICAS DO PROCESSAMENTO\n";
                    mensagemEstatisticas += $"════════════════════════════════════════\n\n";
                    mensagemEstatisticas += $"Total de Lotes: {status.TotalLotes}\n";
                    mensagemEstatisticas += $"Lotes Enviados: {status.LotesEnviados}\n";
                    mensagemEstatisticas += $"Lotes com Erro: {status.LotesComErro}\n";
                    mensagemEstatisticas += $"Lotes Assinados: {status.LotesAssinados}\n";
                    mensagemEstatisticas += $"Lotes Criptografados: {status.LotesCriptografados}\n";
                    if (status.TempoMedioPorLote.HasValue)
                    {
                        mensagemEstatisticas += $"Tempo Médio por Lote: {status.TempoMedioPorLote.Value:mm\\:ss}\n";
                    }
                    TimeSpan tempoTotal = DateTime.Now - status.InicioProcessamento;
                    mensagemEstatisticas += $"Tempo Total: {tempoTotal:hh\\:mm\\:ss}\n";
                    
                    MessageBox.Show(mensagemEstatisticas, "Processamento Completo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                });
            }
            catch (Exception ex)
            {
                AdicionarLog($"ERRO no processamento completo: {ex.Message}");
                if (ex.InnerException != null)
                {
                    AdicionarLog($"Detalhes: {ex.InnerException.Message}");
                }
                status.LotesComErro++;
                status.ModoCompleto = false;
                FinalizarProcessamento(mensagemErro: ex.Message);
                AtualizarEstatisticas();

                ReabilitarControlesAposProcessamento();
                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show($"Erro no processamento completo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }

        // Variáveis temporárias para armazenar arquivos durante processamento completo
        private string arquivoAberturaCriptografado = null;
        private long idLoteAberturaBanco = 0; // ID do lote de abertura no banco
        private Dictionary<string, long> mapArquivoParaIdLote = new Dictionary<string, long>(); // Mapeia arquivo criptografado -> ID do lote no banco

        /// <summary>
        /// Processa abertura (gera, assina, criptografa) sem enviar - para modo completo
        /// </summary>
        private async Task ProcessarAberturaCompleto()
        {
            var config = ConfigForm.Config;
            var dadosAbertura = ConfigForm.DadosAbertura;

            // 1. Gerar XML
            AtualizarEtapa("Gerando XML de abertura...");
            var geradorService = new EfinanceiraGeradorXmlService();
            string arquivoXml = geradorService.GerarXmlAbertura(dadosAbertura, config.DiretorioLotes);
            AdicionarLog($"XML gerado: {arquivoXml}");

            if (cancelarProcessamento) return;

            // Contar eventos
            int quantidadeEventos = ProtocoloPersistenciaService.ContarEventosNoXml(arquivoXml);

            // 2. Assinar
            AtualizarEtapa("Assinando XML...");
            var assinaturaService = new EfinanceiraAssinaturaService();
            X509Certificate2 cert = BuscarCertificado(config.CertThumbprint);
            var xmlAssinado = assinaturaService.AssinarEventosDoArquivo(arquivoXml, cert);
            string arquivoAssinado = arquivoXml.Replace(".xml", SUFIXO_ASSINADO);
            xmlAssinado.Save(arquivoAssinado);
            status.LotesAssinados = 1;
            AtualizarEstatisticas();
            AdicionarLog($"XML assinado: {arquivoAssinado}");

            if (cancelarProcessamento) return;

            // Registrar lote no banco após gerar XML
            try
            {
                var persistenceService = new EfinanceiraDatabasePersistenceService();
                string ambienteStr = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : AMBIENTE_HOMOLOG;
                idLoteAberturaBanco = persistenceService.RegistrarLote(
                    TipoLote.Abertura,
                    config.Periodo,
                    quantidadeEventos,
                    config.CnpjDeclarante,
                    arquivoXml,
                    null, // Ainda não assinado
                    null, // Ainda não criptografado
                    ambienteStr
                );
                persistenceService.RegistrarLogLote(idLoteAberturaBanco, LOG_GERACAO, $"XML gerado: {Path.GetFileName(arquivoXml)}");
            }
            catch (Exception exDb)
            {
                AdicionarLog($"⚠ Aviso: Erro ao registrar lote no banco: {exDb.Message}");
            }

            if (cancelarProcessamento) return;

            // 3. Criptografar
            AtualizarEtapa("Criptografando XML...");
            var criptografiaService = new EfinanceiraCriptografiaService();
            arquivoAberturaCriptografado = criptografiaService.CriptografarLote(arquivoAssinado, config.CertServidorThumbprint);
            status.LotesCriptografados = 1;
            AtualizarEstatisticas();
            AdicionarLog($"XML criptografado: {arquivoAberturaCriptografado}");
            
            // Verificar se o arquivo foi criado corretamente
            if (string.IsNullOrEmpty(arquivoAberturaCriptografado))
            {
                throw new Exception("Erro: Arquivo criptografado não foi gerado!");
            }
            
            if (!File.Exists(arquivoAberturaCriptografado))
            {
                throw new Exception($"Erro: Arquivo criptografado não encontrado: {arquivoAberturaCriptografado}");
            }
            
            AdicionarLog($"✓ Arquivo criptografado verificado: {Path.GetFileName(arquivoAberturaCriptografado)} ({new FileInfo(arquivoAberturaCriptografado).Length} bytes)");

            // Atualizar lote no banco após assinar e criptografar
            if (idLoteAberturaBanco > 0)
            {
                try
                {
                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                    persistenceService.AtualizarLote(idLoteAberturaBanco, STATUS_CRIPTOGRAFADO);
                    persistenceService.RegistrarLogLote(idLoteAberturaBanco, STATUS_ASSINATURA, $"XML assinado: {Path.GetFileName(arquivoAssinado)}");
                    persistenceService.RegistrarLogLote(idLoteAberturaBanco, STATUS_CRIPTOGRAFIA, $"XML criptografado: {Path.GetFileName(arquivoAberturaCriptografado)}");
                }
                catch (Exception exDb)
                {
                    AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                }
            }
        }

        /// <summary>
        /// Envia abertura e retorna true se foi bem-sucedido - para modo completo
        /// </summary>
        private async Task<bool> EnviarAberturaCompleto()
        {
            if (string.IsNullOrEmpty(arquivoAberturaCriptografado))
            {
                AdicionarLog("✗ ERRO: Arquivo de abertura criptografado está vazio!");
                return false;
            }

            if (!File.Exists(arquivoAberturaCriptografado))
            {
                AdicionarLog($"✗ ERRO: Arquivo de abertura não encontrado: {arquivoAberturaCriptografado}");
                return false;
            }

            string arquivoCriptografado = arquivoAberturaCriptografado;
            var config = ConfigForm.Config;
            X509Certificate2 cert = BuscarCertificado(config.CertThumbprint);

            try
            {
                AdicionarLog($"Enviando arquivo: {Path.GetFileName(arquivoCriptografado)}");
                AdicionarLog($"Tamanho do arquivo: {new FileInfo(arquivoCriptografado).Length} bytes");
                AtualizarEtapa("Enviando abertura para e-Financeira...");
                
                var envioService = new EfinanceiraEnvioService();
                var resposta = envioService.EnviarLote(arquivoCriptografado, config, cert);
                
                AdicionarLog($"Resposta recebida - Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}");
                
                string protocoloFinal = ExtrairProtocolo(resposta);
                
                // Se há protocolo, o envio foi bem-sucedido (lote foi recebido pela e-Financeira)
                // O protocolo é o indicador principal de sucesso, não apenas o código de resposta
                if (!string.IsNullOrEmpty(protocoloFinal))
                {
                    status.LotesEnviados = 1;
                    if (!status.ProtocolosEnviados.Contains(protocoloFinal))
                    {
                        status.ProtocolosEnviados.Add(protocoloFinal);
                    }
                    
                    // Atualizar lote no banco após envio
                    if (idLoteAberturaBanco > 0)
                    {
                        try
                        {
                            var persistenceService = new EfinanceiraDatabasePersistenceService();
                            persistenceService.AtualizarLote(
                                idLoteAberturaBanco,
                                STATUS_ENVIADO,
                                protocoloFinal,
                                resposta.CodigoResposta,
                                resposta.Descricao,
                                resposta.XmlCompleto ?? "",
                                null, null, null,
                                DateTime.Now,
                                null,
                                null
                            );
                            persistenceService.RegistrarLogLote(idLoteAberturaBanco, STATUS_ENVIO, $"Abertura enviada. Protocolo: {protocoloFinal}");
                        }
                        catch (Exception exDb)
                        {
                            AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                        }
                    }
                    
                    AtualizarEstatisticas();
                    AdicionarLog($"✓ Abertura enviada com sucesso! Protocolo: {protocoloFinal}");
                    
                    // Limpar apenas após confirmar sucesso
                    arquivoAberturaCriptografado = null;
                    idLoteAberturaBanco = 0;
                    
                    return true;
                }
                else
                {
                    // Sem protocolo = falha no envio
                    // Atualizar lote no banco com erro
                    if (idLoteAberturaBanco > 0)
                    {
                        try
                        {
                            var persistenceService = new EfinanceiraDatabasePersistenceService();
                            persistenceService.AtualizarLote(
                                idLoteAberturaBanco,
                                STATUS_REJEITADO,
                                null,
                                resposta.CodigoResposta,
                                resposta.Descricao,
                                resposta.XmlCompleto ?? "",
                                null, null, null,
                                DateTime.Now,
                                null,
                                $"Falha no envio: {resposta.Descricao}"
                            );
                        }
                        catch (Exception exDb)
                        {
                            AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                        }
                    }
                    
                    AdicionarLog($"✗ Falha ao enviar abertura. Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}");
                    AdicionarLog($"✗ Nenhum protocolo foi retornado pelo servidor.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                AdicionarLog($"✗ Erro ao enviar abertura: {ex.Message}");
                if (ex.InnerException != null)
                {
                    AdicionarLog($"  Detalhes: {ex.InnerException.Message}");
                }
                return false;
            }
        }

        /// <summary>
        /// Processa todos os lotes de movimentação e retorna lista de arquivos criptografados - para modo completo
        /// </summary>
        private async Task<List<string>> ProcessarMovimentacaoCompleto()
        {
            List<string> arquivosCriptografados = new List<string>();
            var config = ConfigForm.Config;

            // Validar e calcular período
            string periodoStr = config.Periodo;
            var (dataInicio, dataFim) = EfinanceiraPeriodoUtil.CalcularPeriodoSemestral(periodoStr);
            DateTime dtInicio = DateTime.Parse(dataInicio, CULTURE_INFO_PT_BR);
            DateTime dtFim = DateTime.Parse(dataFim, CULTURE_INFO_PT_BR);
            int ano = dtInicio.Year;
            int mesInicial = dtInicio.Month;
            int mesFinal = dtFim.Month;

            // Buscar pessoas com contas
            var dbService = new EfinanceiraDatabaseService();
            int pageSize = config.PageSize;
            int offset = config.OffsetRegistros;
            int maxLotes = config.MaxLotes ?? int.MaxValue;
            int lotesGerados = 0;
            int eventosOffset = config.EventoOffset - 1;
            int eventosPorLote = config.EventosPorLote > 0 && config.EventosPorLote <= 50 ? config.EventosPorLote : 50;

            X509Certificate2 cert = BuscarCertificado(config.CertThumbprint);
            var assinaturaService = new EfinanceiraAssinaturaService();
            var criptografiaService = new EfinanceiraCriptografiaService();

            while (lotesGerados < maxLotes && !cancelarProcessamento)
            {
                var pessoas = dbService.BuscarPessoasComContas(ano, mesInicial, mesFinal, pageSize, offset);
                
                if (pessoas.Count == 0)
                {
                    AdicionarLog("Não há mais registros para processar.");
                    break;
                }

                for (int i = eventosOffset; i < pessoas.Count && lotesGerados < maxLotes; i += eventosPorLote)
                {
                    if (cancelarProcessamento) break;

                    var pessoasLote = pessoas.Skip(i).Take(eventosPorLote).ToList();
                    if (pessoasLote.Count == 0) break;

                    lotesGerados++;
                    status.TotalLotes = lotesGerados;
                    status.TotalLotesMovimentacao = lotesGerados;

                    AtualizarEtapa($"Processando lote {lotesGerados} de movimentação ({pessoasLote.Count} eventos)...");

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

                    // Registrar lote no banco após gerar XML
                    long idLoteBanco = 0;
                    try
                    {
                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                        string ambienteStr = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : AMBIENTE_HOMOLOG;
                        idLoteBanco = persistenceService.RegistrarLote(
                            TipoLote.Movimentacao,
                            periodoStr,
                            pessoasLote.Count,
                            config.CnpjDeclarante,
                            arquivoXml,
                            null, // Ainda não assinado
                            null, // Ainda não criptografado
                            ambienteStr
                        );
                        persistenceService.RegistrarEventosDoLote(idLoteBanco, pessoasLote);
                        persistenceService.RegistrarLogLote(idLoteBanco, LOG_GERACAO, $"XML gerado: {Path.GetFileName(arquivoXml)}");
                    }
                    catch (Exception exDb)
                    {
                        AdicionarLog($"⚠ Aviso: Erro ao registrar lote no banco: {exDb.Message}");
                    }

                    // Assinar
                    var xmlAssinado = assinaturaService.AssinarEventosDoArquivo(arquivoXml, cert);
                    string arquivoAssinado = arquivoXml.Replace(".xml", SUFIXO_ASSINADO);
                    xmlAssinado.Save(arquivoAssinado);
                    status.LotesAssinados++;
                    AtualizarEstatisticas();

                    // Atualizar lote no banco após assinar
                    if (idLoteBanco > 0)
                    {
                        try
                        {
                            var persistenceService = new EfinanceiraDatabasePersistenceService();
                            persistenceService.AtualizarLote(idLoteBanco, STATUS_ASSINADO);
                            persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ASSINATURA, $"XML assinado: {Path.GetFileName(arquivoAssinado)}");
                        }
                        catch (Exception exDb)
                        {
                            AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                        }
                    }

                    // Criptografar
                    string arquivoCriptografado = criptografiaService.CriptografarLote(arquivoAssinado, config.CertServidorThumbprint);
                    status.LotesCriptografados++;
                    arquivosCriptografados.Add(arquivoCriptografado);
                    
                    // Mapear arquivo para ID do lote
                    if (idLoteBanco > 0)
                    {
                        mapArquivoParaIdLote[arquivoCriptografado] = idLoteBanco;
                    }
                    
                    AtualizarEstatisticas();

                    // Atualizar lote no banco após criptografar
                    if (idLoteBanco > 0)
                    {
                        try
                        {
                            var persistenceService = new EfinanceiraDatabasePersistenceService();
                            persistenceService.AtualizarLote(idLoteBanco, STATUS_CRIPTOGRAFADO);
                            persistenceService.RegistrarLogLote(idLoteBanco, STATUS_CRIPTOGRAFIA, $"XML criptografado: {Path.GetFileName(arquivoCriptografado)}");
                        }
                        catch (Exception exDb)
                        {
                            AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                        }
                    }

                    AdicionarLog($"Lote {lotesGerados} processado: {Path.GetFileName(arquivoCriptografado)}");
                    eventosOffset += pessoasLote.Count;
                }

                offset += pageSize;
            }

            return arquivosCriptografados;
        }

        /// <summary>
        /// Envia todos os lotes de movimentação - para modo completo
        /// </summary>
        private async Task EnviarMovimentacoesCompleto(List<string> arquivosCriptografados)
        {
            var config = ConfigForm.Config;
            X509Certificate2 cert = BuscarCertificado(config.CertThumbprint);
            var envioService = new EfinanceiraEnvioService();

            for (int i = 0; i < arquivosCriptografados.Count; i++)
            {
                if (cancelarProcessamento) break;

                string arquivoCriptografado = arquivosCriptografados[i];
                long idLoteBanco = mapArquivoParaIdLote.ContainsKey(arquivoCriptografado) ? mapArquivoParaIdLote[arquivoCriptografado] : 0;
                AtualizarEtapa($"Enviando lote {i + 1} de {arquivosCriptografados.Count} de movimentação...");

                try
                {
                    var resposta = envioService.EnviarLote(arquivoCriptografado, config, cert);
                    string protocoloFinal = ExtrairProtocolo(resposta);

                    // Se há protocolo, o envio foi bem-sucedido (lote foi recebido pela e-Financeira)
                    if (!string.IsNullOrEmpty(protocoloFinal))
                    {
                        status.LotesEnviados++;
                        if (!status.ProtocolosEnviados.Contains(protocoloFinal))
                        {
                            status.ProtocolosEnviados.Add(protocoloFinal);
                        }
                        
                        // Atualizar lote no banco após envio
                        if (idLoteBanco > 0)
                        {
                            try
                            {
                                var persistenceService = new EfinanceiraDatabasePersistenceService();
                                persistenceService.AtualizarLote(
                                    idLoteBanco,
                                    STATUS_ENVIADO,
                                    protocoloFinal,
                                    resposta.CodigoResposta,
                                    resposta.Descricao,
                                    resposta.XmlCompleto ?? "",
                                    null, null, null,
                                    DateTime.Now,
                                    null,
                                    null
                                );
                                persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ENVIO, $"Lote enviado. Protocolo: {protocoloFinal}");
                            }
                            catch (Exception exDb)
                            {
                                AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                            }
                        }
                        
                        AdicionarLog($"✓ Lote {i + 1} enviado com sucesso! Protocolo: {protocoloFinal}");
                    }
                    else
                    {
                        status.LotesComErro++;
                        
                        // Atualizar lote no banco com erro
                        if (idLoteBanco > 0)
                        {
                            try
                            {
                                var persistenceService = new EfinanceiraDatabasePersistenceService();
                                persistenceService.AtualizarLote(
                                    idLoteBanco,
                                    STATUS_REJEITADO,
                                    null,
                                    resposta.CodigoResposta,
                                    resposta.Descricao,
                                    resposta.XmlCompleto ?? "",
                                    null, null, null,
                                    DateTime.Now,
                                    null,
                                    $"Falha no envio: {resposta.Descricao}"
                                );
                            }
                            catch (Exception exDb)
                            {
                                AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                            }
                        }
                        
                        AdicionarLog($"✗ Lote {i + 1} falhou. Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}");
                        AdicionarLog($"✗ Nenhum protocolo foi retornado pelo servidor.");
                    }

                    status.LotesMovimentacaoProcessados++;
                    AtualizarEstatisticas();
                }
                catch (Exception ex)
                {
                    status.LotesComErro++;
                    AdicionarLog($"✗ Erro ao enviar lote {i + 1}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Processa fechamento (gera, assina, criptografa) sem enviar - para modo completo
        /// </summary>
        private async Task<string> ProcessarFechamentoCompleto()
        {
            var config = ConfigForm.Config;
            var dadosFechamento = ConfigForm.DadosFechamento;

            // 1. Gerar XML
            AtualizarEtapa("Gerando XML de fechamento...");
            var geradorService = new EfinanceiraGeradorXmlService();
            string arquivoXml = geradorService.GerarXmlFechamento(dadosFechamento, config.DiretorioLotes);
            AdicionarLog($"XML gerado: {arquivoXml}");

            if (cancelarProcessamento) return null;

            // Registrar lote no banco após gerar XML
            int quantidadeEventos = ProtocoloPersistenciaService.ContarEventosNoXml(arquivoXml);
            long idLoteBanco = 0;
            try
            {
                var persistenceService = new EfinanceiraDatabasePersistenceService();
                string ambienteStr = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : AMBIENTE_HOMOLOG;
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
                persistenceService.RegistrarLogLote(idLoteBanco, LOG_GERACAO, $"XML gerado: {Path.GetFileName(arquivoXml)}");
            }
            catch (Exception exDb)
            {
                AdicionarLog($"⚠ Aviso: Erro ao registrar lote no banco: {exDb.Message}");
            }

            // 2. Assinar
            AtualizarEtapa("Assinando XML...");
            var assinaturaService = new EfinanceiraAssinaturaService();
            X509Certificate2 cert = BuscarCertificado(config.CertThumbprint);
            var xmlAssinado = assinaturaService.AssinarEventosDoArquivo(arquivoXml, cert);
            string arquivoAssinado = arquivoXml.Replace(".xml", SUFIXO_ASSINADO);
            xmlAssinado.Save(arquivoAssinado);
            status.LotesAssinados++;
            AtualizarEstatisticas();
            AdicionarLog($"XML assinado: {arquivoAssinado}");

            // Atualizar lote no banco após assinar
            if (idLoteBanco > 0)
            {
                try
                {
                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                    persistenceService.AtualizarLote(idLoteBanco, STATUS_ASSINADO);
                    persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ASSINATURA, $"XML assinado: {Path.GetFileName(arquivoAssinado)}");
                }
                catch (Exception exDb)
                {
                    AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                }
            }

            if (cancelarProcessamento) return null;

            // 3. Criptografar
            AtualizarEtapa("Criptografando XML...");
            var criptografiaService = new EfinanceiraCriptografiaService();
            string arquivoCriptografado = criptografiaService.CriptografarLote(arquivoAssinado, config.CertServidorThumbprint);
            status.LotesCriptografados++;
            AtualizarEstatisticas();
            AdicionarLog($"XML criptografado: {arquivoCriptografado}");

            // Atualizar lote no banco após criptografar
            if (idLoteBanco > 0)
            {
                try
                {
                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                    persistenceService.AtualizarLote(idLoteBanco, STATUS_CRIPTOGRAFADO);
                    persistenceService.RegistrarLogLote(idLoteBanco, STATUS_CRIPTOGRAFIA, $"XML criptografado: {Path.GetFileName(arquivoCriptografado)}");
                    // Armazenar ID do lote para atualização após envio
                    mapArquivoParaIdLote[arquivoCriptografado] = idLoteBanco;
                }
                catch (Exception exDb)
                {
                    AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                }
            }

            return arquivoCriptografado;
        }

        /// <summary>
        /// Envia fechamento - para modo completo
        /// </summary>
        private async Task EnviarFechamentoCompleto(string arquivoCriptografado)
        {
            var config = ConfigForm.Config;
            X509Certificate2 cert = BuscarCertificado(config.CertThumbprint);
            long idLoteBanco = mapArquivoParaIdLote.ContainsKey(arquivoCriptografado) ? mapArquivoParaIdLote[arquivoCriptografado] : 0;

            try
            {
                AtualizarEtapa("Enviando fechamento para e-Financeira...");
                var envioService = new EfinanceiraEnvioService();
                var resposta = envioService.EnviarLote(arquivoCriptografado, config, cert);
                
                AdicionarLog($"Resposta recebida - Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}");
                
                string protocoloFinal = ExtrairProtocolo(resposta);
                
                // Se há protocolo, o envio foi bem-sucedido (lote foi recebido pela e-Financeira)
                if (!string.IsNullOrEmpty(protocoloFinal))
                {
                    status.LotesEnviados++;
                    if (!status.ProtocolosEnviados.Contains(protocoloFinal))
                    {
                        status.ProtocolosEnviados.Add(protocoloFinal);
                    }
                    
                    // Atualizar lote no banco após envio
                    if (idLoteBanco > 0)
                    {
                        try
                        {
                            var persistenceService = new EfinanceiraDatabasePersistenceService();
                            persistenceService.AtualizarLote(
                                idLoteBanco,
                                STATUS_ENVIADO,
                                protocoloFinal,
                                resposta.CodigoResposta,
                                resposta.Descricao,
                                resposta.XmlCompleto ?? "",
                                null, null, null,
                                DateTime.Now,
                                null,
                                null
                            );
                            persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ENVIO, $"Fechamento enviado. Protocolo: {protocoloFinal}");
                        }
                        catch (Exception exDb)
                        {
                            AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                        }
                    }
                    
                    AtualizarEstatisticas();
                    AdicionarLog($"✓ Fechamento enviado com sucesso! Protocolo: {protocoloFinal}");
                }
                else
                {
                    status.LotesComErro++;
                    
                    // Atualizar lote no banco com erro
                    if (idLoteBanco > 0)
                    {
                        try
                        {
                            var persistenceService = new EfinanceiraDatabasePersistenceService();
                            persistenceService.AtualizarLote(
                                idLoteBanco,
                                STATUS_REJEITADO,
                                null,
                                resposta.CodigoResposta,
                                resposta.Descricao,
                                resposta.XmlCompleto ?? "",
                                null, null, null,
                                DateTime.Now,
                                null,
                                $"Falha no envio: {resposta.Descricao}"
                            );
                        }
                        catch (Exception exDb)
                        {
                            AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                        }
                    }
                    
                    AdicionarLog($"✗ Falha ao enviar fechamento. Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}");
                    AdicionarLog($"✗ Nenhum protocolo foi retornado pelo servidor.");
                }
            }
            catch (Exception ex)
            {
                status.LotesComErro++;
                AdicionarLog($"✗ Erro ao enviar fechamento: {ex.Message}");
            }
        }

        private async Task ProcessarCadastroDeclarante()
        {
            try
            {
                AtualizarEtapa("Iniciando processamento de cadastro de declarante...");
                status.InicioProcessamento = DateTime.Now;
                status.TotalLotes = 1;
                status.ProtocolosEnviados.Clear();

                var config = ConfigForm.Config;
                var dadosCadastro = ConfigForm.DadosCadastroDeclarante;

                // 1. Gerar XML
                AtualizarEtapa("Gerando XML de cadastro de declarante...");
                var geradorService = new EfinanceiraGeradorXmlService();
                string arquivoXml = geradorService.GerarXmlCadastroDeclarante(dadosCadastro, config.DiretorioLotes);
                AdicionarLog($"XML gerado: {arquivoXml}");

                if (cancelarProcessamento) return;

                // Registrar lote no banco após gerar XML
                int quantidadeEventos = 1; // Cadastro de declarante sempre tem 1 evento
                long idLoteBanco = 0;
                try
                {
                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                    string ambienteStr = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : AMBIENTE_HOMOLOG;
                    idLoteBanco = persistenceService.RegistrarLote(
                        TipoLote.CadastroDeclarante,
                        null, // Cadastro não tem período
                        quantidadeEventos,
                        config.CnpjDeclarante,
                        arquivoXml,
                        null, // Ainda não assinado
                        null, // Ainda não criptografado
                        ambienteStr
                    );
                    persistenceService.RegistrarLogLote(idLoteBanco, LOG_GERACAO, $"XML gerado: {Path.GetFileName(arquivoXml)}");
                    AdicionarLog($"Lote registrado no banco (ID: {idLoteBanco}).");
                }
                catch (Exception exDb)
                {
                    string erroCompleto = $"⚠ ERRO ao registrar lote no banco: {exDb.Message}";
                    AdicionarLog(erroCompleto);
                    System.Diagnostics.Debug.WriteLine(erroCompleto);
                    System.Diagnostics.Debug.WriteLine($"Stack Trace: {exDb.StackTrace}");
                }

                if (cancelarProcessamento) return;

                // 2. Assinar XML
                AtualizarEtapa("Assinando XML...");
                var assinaturaService = new EfinanceiraAssinaturaService();
                X509Certificate2 cert = BuscarCertificado(config.CertThumbprint);
                var xmlAssinado = assinaturaService.AssinarEventosDoArquivo(arquivoXml, cert);
                string arquivoAssinado = arquivoXml.Replace(".xml", SUFIXO_ASSINADO);
                xmlAssinado.Save(arquivoAssinado);
                AdicionarLog($"XML assinado: {arquivoAssinado}");

                if (idLoteBanco > 0)
                {
                    try
                    {
                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                        persistenceService.AtualizarLote(idLoteBanco, STATUS_ASSINADO);
                        persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ASSINATURA, $"XML assinado: {Path.GetFileName(arquivoAssinado)}");
                    }
                    catch (Exception exDb)
                    {
                        AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após assinatura: {exDb.Message}");
                    }
                }

                if (cancelarProcessamento) return;

                // 3. Criptografar XML
                AtualizarEtapa("Criptografando XML...");
                var criptografiaService = new EfinanceiraCriptografiaService();
                string arquivoCriptografado = criptografiaService.CriptografarLote(arquivoAssinado, config.CertServidorThumbprint);
                AdicionarLog($"XML criptografado: {arquivoCriptografado}");

                if (idLoteBanco > 0)
                {
                    try
                    {
                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                        persistenceService.AtualizarLote(idLoteBanco, STATUS_CRIPTOGRAFADO);
                        persistenceService.RegistrarLogLote(idLoteBanco, STATUS_CRIPTOGRAFIA, $"XML criptografado: {Path.GetFileName(arquivoCriptografado)}");
                    }
                    catch (Exception exDb)
                    {
                        AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após criptografia: {exDb.Message}");
                    }
                }

                if (cancelarProcessamento) return;

                // 4. Enviar (se habilitado)
                if (!chkApenasProcessar.Checked && config.TestEnvioHabilitado)
                {
                    AtualizarEtapa("Enviando lote para e-Financeira...");
                    X509Certificate2 certificado = BuscarCertificado(config.CertThumbprint);
                    var envioService = new EfinanceiraEnvioService();
                    var resposta = envioService.EnviarLote(arquivoCriptografado, config, certificado);

                    AdicionarLog($"Código de Resposta: {resposta.CodigoResposta}");
                    AdicionarLog($"Descrição: {resposta.Descricao}");

                    // Extrair protocolo
                    string protocoloFinal = ExtrairProtocolo(resposta);

                    if (resposta.CodigoResposta == 1)
                    {
                        if (!string.IsNullOrEmpty(protocoloFinal))
                        {
                            if (!status.ProtocolosEnviados.Contains(protocoloFinal))
                            {
                                status.ProtocolosEnviados.Add(protocoloFinal);
                            }
                            
                            AdicionarLog($"✓ Lote de cadastro enviado com sucesso! Protocolo: {protocoloFinal}");
                            AdicionarLog($"════════════════════════════════════════");
                            AdicionarLog($"PROTOCOLO DO LOTE DE CADASTRO:");
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
                                    persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ENVIO, $"Lote enviado com sucesso. Protocolo: {protocoloFinal}");
                                }
                                catch (Exception exDb)
                                {
                                    AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco após envio: {exDb.Message}");
                                }
                            }
                            
                            // Atualizar protocolo no lote já registrado (sistema antigo)
                            ProtocoloPersistenciaService.RegistrarProtocolo(
                                TipoLote.CadastroDeclarante, 
                                arquivoCriptografado, 
                                protocoloFinal,
                                null,
                                quantidadeEventos
                            );
                        }
                        else
                        {
                            AdicionarLog($"⚠ Lote enviado, mas protocolo não foi extraído da resposta.");
                        }
                    }
                    else if (resposta.CodigoResposta == 7)
                    {
                        AdicionarLog($"✗ Lote de cadastro REJEITADO - Código: {resposta.CodigoResposta}");
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
                                    STATUS_REJEITADO,
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
                    }
                    else
                    {
                        AdicionarLog($"⚠ Lote enviado, mas código de resposta não é 1 (sucesso) nem 7 (rejeição). Código: {resposta.CodigoResposta}");
                        
                        // Tentar extrair protocolo mesmo assim
                        if (!string.IsNullOrEmpty(protocoloFinal))
                        {
                            if (!status.ProtocolosEnviados.Contains(protocoloFinal))
                            {
                                status.ProtocolosEnviados.Add(protocoloFinal);
                            }
                            
                            AdicionarLog($"✓ Protocolo extraído: {protocoloFinal}");
                            
                            // Atualizar lote no banco
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
                                    persistenceService.RegistrarLogLote(idLoteBanco, STATUS_ENVIO, $"Lote enviado. Protocolo: {protocoloFinal}");
                                }
                                catch (Exception exDb)
                                {
                                    AdicionarLog($"⚠ Aviso: Erro ao atualizar lote no banco: {exDb.Message}");
                                }
                            }
                            
                            // Atualizar protocolo no lote já registrado (sistema antigo)
                            ProtocoloPersistenciaService.RegistrarProtocolo(
                                TipoLote.CadastroDeclarante, 
                                arquivoCriptografado, 
                                protocoloFinal,
                                null,
                                quantidadeEventos
                            );
                        }
                    }
                }
                else
                {
                    AdicionarLog("ℹ Envio desabilitado ou modo 'Apenas Processar' ativado. XML gerado, assinado e criptografado, mas não enviado.");
                }

                status.LotesProcessados = 1;
                status.LotesAssinados = 1;
                status.LotesCriptografados = 1;
                if (!chkApenasProcessar.Checked && config.TestEnvioHabilitado)
                {
                    status.LotesEnviados = 1;
                }

                AtualizarEtapa("Processamento concluído!");
                AdicionarLog("✓ Processamento de cadastro de declarante concluído com sucesso!");

                ReabilitarControlesAposProcessamento();
            }
            catch (Exception ex)
            {
                AdicionarLog($"ERRO: {ex.Message}");
                status.LotesComErro = 1;
                ReabilitarControlesAposProcessamento();
                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show($"Erro ao processar: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }

        private static X509Certificate2 BuscarCertificado(string thumbprint)
        {
            if (string.IsNullOrEmpty(thumbprint))
            {
                throw new ArgumentException("Thumbprint do certificado não configurado.", nameof(thumbprint));
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

            throw new InvalidOperationException($"Certificado com thumbprint '{thumbprint}' não encontrado.");
        }

        private void AtualizarEtapa(string etapa)
        {
            this.Invoke((MethodInvoker)delegate
            {
                status.EtapaAtual = etapa;
                lblEtapaAtual.Text = $"Etapa: {etapa}";
                status.TempoDecorrido = DateTime.UtcNow - status.InicioProcessamento;
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

                // Atualizar status de etapa (0-6)
                string textoEtapa = "";
                Color corEtapa = Color.Black;
                
                if (status.ModoCompleto)
                {
                    switch (status.StatusEtapa)
                    {
                        case 0:
                            textoEtapa = "0 - Aguardando início do processamento";
                            corEtapa = Color.Gray;
                            break;
                        case 1:
                            textoEtapa = "1 - Processando Abertura";
                            corEtapa = Color.Blue;
                            break;
                        case 2:
                            textoEtapa = "2 - Abertura Enviada";
                            corEtapa = Color.Green;
                            break;
                        case 3:
                            textoEtapa = "3 - Processando lotes de movimentação financeira";
                            corEtapa = Color.Orange;
                            break;
                        case 4:
                            textoEtapa = "4 - Enviando lotes de movimentação financeira";
                            corEtapa = Color.DarkOrange;
                            break;
                        case 5:
                            textoEtapa = "5 - Processando lote de fechamento";
                            corEtapa = Color.Purple;
                            break;
                        case 6:
                            textoEtapa = "6 - Fechamento Enviado";
                            corEtapa = Color.Green;
                            break;
                        default:
                            textoEtapa = "Status desconhecido";
                            corEtapa = Color.Black;
                            break;
                    }
                }
                else
                {
                    // Quando não está em modo completo, usar EtapaAtual (string)
                    textoEtapa = string.IsNullOrEmpty(status.EtapaAtual) ? "Aguardando..." : status.EtapaAtual;
                    corEtapa = Color.Black;
                }
                
                lblStatusEtapa.Text = $"Status Etapa: {textoEtapa}";
                lblStatusEtapa.ForeColor = corEtapa;

                // Atualizar status de abertura
                if (status.AberturaFinalizada || status.StatusEtapa >= 2)
                {
                    lblAberturaFinalizada.Text = "Abertura: ✓ Finalizada";
                    lblAberturaFinalizada.ForeColor = Color.Green;
                }
                else if (status.StatusEtapa == 1)
                {
                    lblAberturaFinalizada.Text = "Abertura: Em processamento...";
                    lblAberturaFinalizada.ForeColor = Color.Blue;
                }
                else
                {
                    lblAberturaFinalizada.Text = "Abertura: Não iniciada";
                    lblAberturaFinalizada.ForeColor = Color.Gray;
                }

                // Atualizar tempo médio por lote
                if (status.TempoMedioPorLote.HasValue && status.TempoMedioPorLote.Value.TotalMilliseconds > 0)
                {
                    lblTempoMedioPorLote.Text = $"Tempo Médio/Lote: {status.TempoMedioPorLote.Value:mm\\:ss}";
                }
                else if ((status.StatusEtapa == 3 || status.StatusEtapa == 4) && status.TotalLotesMovimentacao > 0)
                {
                    // Calcular tempo médio em tempo real durante processamento
                    if (status.TemposLotesMovimentacao.Count > 0 && status.LotesMovimentacaoProcessados > 0)
                    {
                        TimeSpan tempoTotal = DateTime.Now - status.TemposLotesMovimentacao[0];
                        TimeSpan tempoMedio = TimeSpan.FromMilliseconds(tempoTotal.TotalMilliseconds / status.LotesMovimentacaoProcessados);
                        lblTempoMedioPorLote.Text = $"Tempo Médio/Lote: {tempoMedio:mm\\:ss} (calculando...)";
                    }
                    else
                    {
                        lblTempoMedioPorLote.Text = $"Tempo Médio/Lote: Calculando...";
                    }
                }
                else
                {
                    lblTempoMedioPorLote.Text = "Tempo Médio/Lote: -";
                }

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
