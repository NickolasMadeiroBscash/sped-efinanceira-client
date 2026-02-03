using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExemploAssinadorXML.Models;
using ExemploAssinadorXML.Services;
using ConfiguracaoCompleta = ExemploAssinadorXML.Services.ConfiguracaoCompleta;

namespace ExemploAssinadorXML.Forms
{
    public partial class ConfiguracaoForm : Form
    {
        // Constantes para mensagens e tooltips
        private const string TITULO_VALIDACAO = "Validação";
        private const string TOOLTIP_TIPO_AMBIENTE = "Tipo de ambiente:\n• 1 - Produção: Ambiente oficial da Receita Federal\n• 2 - Homologação: Ambiente de testes";
        private const string TOOLTIP_APLICACAO_EMISSORA = "Aplicação que está gerando o evento:\n• 1 - Aplicação do contribuinte: Sistema próprio da empresa\n• 2 - Outros: Outras aplicações";
        private const string TOOLTIP_IND_RETIFICACAO = "Indicador de retificação:\n• 1 - Original: Declaração original\n• 2 - Retificação espontânea: Correção feita pela empresa\n• 3 - Retificação a pedido: Correção solicitada pela Receita\n\nSe selecionar 2 ou 3, preencha o Nº Recibo.";
        private const string TOOLTIP_NR_RECIBO = "Número do recibo da declaração original que está sendo retificada.\nObrigatório apenas quando 'Ind Retificação' for 2 ou 3.";
        private const string TOOLTIP_COMPLEMENTO_ENDERECO = "Complemento do endereço (apto, sala, bloco, etc.) - opcional.";
        
        // Constantes para URLs
        private const string URL_TESTE = "https://pre-efinanceira.receita.fazenda.gov.br/recepcao/lotes/cripto";
        private const string URL_PRODUCAO = "https://efinanceira.receita.fazenda.gov.br/recepcao/lotes/cripto";
        private const string URL_CONSULTA_TESTE = "https://pre-efinanceira.receita.fazenda.gov.br/consulta/lotes/";
        private const string URL_CONSULTA_PRODUCAO = "https://efinanceira.receita.fazenda.gov.br/consulta/lotes/";
        
        // Constante para formato de data
        private static readonly System.Globalization.CultureInfo CULTURE_INFO_PT_BR = System.Globalization.CultureInfo.GetCultureInfo("pt-BR");
        private GroupBox grpConfigGeral;
        private Label lblCnpjDeclarante;
        private TextBox txtCnpjDeclarante;
        private Label lblCertThumbprint;
        private TextBox txtCertThumbprint;
        private Button btnSelecionarCert;
        private Label lblCertServidorThumbprint;
        private TextBox txtCertServidorThumbprint;
        private Button btnSelecionarCertServidor;
        private ComboBox cmbAmbiente;
        private Label lblAmbiente;
        private CheckBox chkModoTeste;
        private CheckBox chkHabilitarEnvio;
        private TextBox txtDiretorioLotes;
        private Label lblDiretorioLotes;
        private Button btnSelecionarDiretorio;
        private TextBox txtPeriodo;
        private Label lblPeriodo;
        private ComboBox cmbSemestre;
        private Label lblSemestre;
        private NumericUpDown numAno;
        private Label lblAno;
        private CheckBox chkCalcularPeriodoAutomatico;
        private Button btnCarregarConfig;
        private Button btnTestarConexao;

        // Configurações de Processamento
        private GroupBox grpProcessamento;
        private NumericUpDown numPageSize;
        private NumericUpDown numEventoOffset;
        private NumericUpDown numOffsetRegistros;
        private NumericUpDown numMaxLotes;
        private CheckBox chkMaxLotesIlimitado;
        private NumericUpDown numEventosPorLote;
        private ToolTip toolTip;

        private TabControl tabConfig;
        private TabPage tabAbertura;
        private TabPage tabFechamento;
        private TabPage tabCadastroDeclarante;
        private ScrollableControl scrollAbertura;
        private ScrollableControl scrollFechamento;
        private ScrollableControl scrollCadastroDeclarante;

        // Abertura - Básicos
        private TextBox txtDtInicioAbertura;
        private TextBox txtDtFimAbertura;
        private ComboBox cmbTipoAmbienteAbertura;
        private ComboBox cmbAplicacaoEmissoraAbertura;
        private ComboBox cmbIndRetificacaoAbertura;
        private TextBox txtNrReciboAbertura;
        private CheckBox chkIndicarMovOpFin;
        private CheckBox chkCalcularPeriodoAbertura;
        private ComboBox cmbSemestreAbertura;
        private NumericUpDown numAnoAbertura;

        // ResponsavelRMF
        private TextBox txtRMF_CNPJ;
        private TextBox txtRMF_CPF;
        private TextBox txtRMF_Nome;
        private TextBox txtRMF_Setor;
        private TextBox txtRMF_DDD;
        private TextBox txtRMF_Telefone;
        private TextBox txtRMF_Ramal;
        private TextBox txtRMF_Logradouro;
        private TextBox txtRMF_Numero;
        private TextBox txtRMF_Complemento;
        private TextBox txtRMF_Bairro;
        private TextBox txtRMF_CEP;
        private TextBox txtRMF_Municipio;
        private TextBox txtRMF_UF;

        // RespeFin
        private TextBox txtRespeFin_CPF;
        private TextBox txtRespeFin_Nome;
        private TextBox txtRespeFin_Setor;
        private TextBox txtRespeFin_DDD;
        private TextBox txtRespeFin_Telefone;
        private TextBox txtRespeFin_Ramal;
        private TextBox txtRespeFin_Logradouro;
        private TextBox txtRespeFin_Numero;
        private TextBox txtRespeFin_Complemento;
        private TextBox txtRespeFin_Bairro;
        private TextBox txtRespeFin_CEP;
        private TextBox txtRespeFin_Municipio;
        private TextBox txtRespeFin_UF;
        private TextBox txtRespeFin_Email;

        // RepresLegal
        private TextBox txtRepresLegal_CPF;
        private TextBox txtRepresLegal_Setor;
        private TextBox txtRepresLegal_DDD;
        private TextBox txtRepresLegal_Telefone;
        private TextBox txtRepresLegal_Ramal;

        // Fechamento
        private TextBox txtDtInicioFechamento;
        private TextBox txtDtFimFechamento;
        private ComboBox cmbTipoAmbienteFechamento;
        private ComboBox cmbAplicacaoEmissoraFechamento;
        private ComboBox cmbIndRetificacaoFechamento;
        private TextBox txtNrReciboFechamento;
        private ComboBox cmbSitEspecial;
        private CheckBox chkNadaADeclarar;
        private CheckBox chkFechamentoPP;
        private CheckBox chkFechamentoMovOpFin;
        private CheckBox chkFechamentoMovOpFinAnual;
        private CheckBox chkCalcularPeriodoFechamento;
        private ComboBox cmbSemestreFechamento;
        private NumericUpDown numAnoFechamento;

        // Cadastro Declarante
        private ComboBox cmbTipoAmbienteCadastro;
        private ComboBox cmbAplicacaoEmissoraCadastro;
        private ComboBox cmbIndRetificacaoCadastro;
        private TextBox txtNrReciboCadastro;
        private TextBox txtGIIN;
        private TextBox txtCategoriaDeclarante;
        private TextBox txtNomeCadastro;
        private TextBox txtTpNome;
        private TextBox txtEnderecoLivreCadastro;
        private TextBox txtTpEnderecoCadastro;
        private TextBox txtMunicipioCadastro;
        private TextBox txtUFCadastro;
        private TextBox txtCEPCadastro;
        private TextBox txtPaisCadastro;
        private TextBox txtPaisResid;

        private Button btnSalvarConfig;
        private Button btnLimparDadosTeste;
        private readonly ConfiguracaoPersistenciaService persistenciaService;

        public DadosAbertura DadosAbertura { get; private set; }
        public DadosFechamento DadosFechamento { get; private set; }
        public DadosCadastroDeclarante DadosCadastroDeclarante { get; private set; }
        public EfinanceiraConfig Config { get; private set; }

        public ConfiguracaoForm()
        {
            persistenciaService = new ConfiguracaoPersistenciaService();
            InitializeComponent();
            CarregarConfiguracoesPadrao();
            // Carregar configuração salva após definir valores padrão
            CarregarConfiguracaoSalva();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Text = "Configuração e-Financeira";
            this.Size = new Size(1250, 800);
            this.MinimumSize = new Size(1200, 700);

            // ToolTip
            toolTip = new ToolTip();
            toolTip.IsBalloon = true;
            toolTip.ToolTipTitle = "Ajuda";
            toolTip.ToolTipIcon = ToolTipIcon.Info;

            // Configuração Geral
            grpConfigGeral = new GroupBox();
            grpConfigGeral.Text = "Configuração Geral";
            grpConfigGeral.Location = new Point(10, 10);
            grpConfigGeral.Size = new Size(1230, 400);
            grpConfigGeral.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            int yPos = 25;
            int espacamentoVertical = 35;
            int labelWidth = 180;
            int campoX = 200;
            int campoWidth = 400;

            // Linha 1: CNPJ Declarante
            lblCnpjDeclarante = new Label { Text = "CNPJ Declarante:", Location = new Point(10, yPos), Size = new Size(labelWidth, 20) };
            txtCnpjDeclarante = new TextBox { Location = new Point(campoX, yPos - 3), Size = new Size(200, 23) };
            txtCnpjDeclarante.Leave += ValidarCNPJ;
            toolTip.SetToolTip(lblCnpjDeclarante, "CNPJ da empresa declarante (sem pontos, barras ou hífens).\nExemplo: 12345678000190");
            toolTip.SetToolTip(txtCnpjDeclarante, "CNPJ da empresa declarante (sem pontos, barras ou hífens).\nExemplo: 12345678000190");
            yPos += espacamentoVertical;

            // Linha 2: Certificado para Assinatura
            lblCertThumbprint = new Label { Text = "Certificado para Assinatura:", Location = new Point(10, yPos), Size = new Size(labelWidth, 20) };
            txtCertThumbprint = new TextBox { Location = new Point(campoX, yPos - 3), Size = new Size(campoWidth, 23), ReadOnly = true };
            btnSelecionarCert = new Button { Text = "Selecionar...", Location = new Point(campoX + campoWidth + 10, yPos - 3), Size = new Size(100, 25) };
            btnSelecionarCert.Click += BtnSelecionarCert_Click;
            toolTip.SetToolTip(lblCertThumbprint, "Certificado digital A1 (instalado no Windows) usado para assinar os XMLs.\nClique em 'Selecionar...' para escolher o certificado.");
            toolTip.SetToolTip(txtCertThumbprint, "Certificado digital A1 (instalado no Windows) usado para assinar os XMLs.\nClique em 'Selecionar...' para escolher o certificado.");
            toolTip.SetToolTip(btnSelecionarCert, "Abre a lista de certificados instalados no Windows para seleção.");
            yPos += espacamentoVertical;

            // Linha 3: Certificado do Servidor
            lblCertServidorThumbprint = new Label { Text = "Certificado do Servidor:", Location = new Point(10, yPos), Size = new Size(labelWidth, 20) };
            txtCertServidorThumbprint = new TextBox { Location = new Point(campoX, yPos - 3), Size = new Size(campoWidth, 23), ReadOnly = true };
            btnSelecionarCertServidor = new Button { Text = "Selecionar...", Location = new Point(campoX + campoWidth + 10, yPos - 3), Size = new Size(100, 25) };
            btnSelecionarCertServidor.Click += BtnSelecionarCertServidor_Click;
            toolTip.SetToolTip(lblCertServidorThumbprint, "Certificado público do servidor da e-Financeira usado para criptografar os XMLs.\nEste certificado é fornecido pela Receita Federal.");
            toolTip.SetToolTip(txtCertServidorThumbprint, "Certificado público do servidor da e-Financeira usado para criptografar os XMLs.\nEste certificado é fornecido pela Receita Federal.");
            toolTip.SetToolTip(btnSelecionarCertServidor, "Abre a lista de certificados para selecionar o certificado público do servidor.");
            yPos += espacamentoVertical;

            // Linha 4: Seleção Automática de Semestre
            chkCalcularPeriodoAutomatico = new CheckBox 
            { 
                Text = "Calcular período automaticamente", 
                Location = new Point(10, yPos), 
                Size = new Size(250, 20),
                Checked = true
            };
            toolTip.SetToolTip(chkCalcularPeriodoAutomatico, "Quando marcado, calcula automaticamente o período baseado no semestre e ano selecionados.\nO campo de período fica somente leitura.");
            chkCalcularPeriodoAutomatico.CheckedChanged += (s, e) =>
            {
                bool calcularAutomatico = chkCalcularPeriodoAutomatico.Checked;
                cmbSemestre.Enabled = calcularAutomatico;
                numAno.Enabled = calcularAutomatico;
                txtPeriodo.ReadOnly = calcularAutomatico;
                txtPeriodo.BackColor = calcularAutomatico ? SystemColors.Control : SystemColors.Window;
                
                if (calcularAutomatico)
                {
                    CalcularPeriodoAutomatico();
                }
            };
            yPos += espacamentoVertical;

            // Linha 5: Semestre e Ano (para cálculo automático)
            lblSemestre = new Label { Text = "Semestre:", Location = new Point(10, yPos), Size = new Size(100, 20) };
            cmbSemestre = new ComboBox 
            { 
                Location = new Point(115, yPos - 3), 
                Size = new Size(150, 23), 
                DropDownStyle = ComboBoxStyle.DropDownList 
            };
            cmbSemestre.Items.AddRange(new[] { "1º Semestre (Jan-Jun)", "2º Semestre (Jul-Dez)" });
            cmbSemestre.SelectedIndex = 0;
            cmbSemestre.SelectedIndexChanged += (s, e) => { if (chkCalcularPeriodoAutomatico.Checked) CalcularPeriodoAutomatico(); };
            toolTip.SetToolTip(lblSemestre, "Selecione o semestre para cálculo automático do período.");
            toolTip.SetToolTip(cmbSemestre, "Selecione o semestre:\n• 1º Semestre = Janeiro a Junho (período YYYY01)\n• 2º Semestre = Julho a Dezembro (período YYYY02)");

            lblAno = new Label { Text = "Ano:", Location = new Point(275, yPos), Size = new Size(50, 20) };
            numAno = new NumericUpDown 
            { 
                Location = new Point(330, yPos - 3), 
                Size = new Size(80, 23),
                Minimum = 2000,
                Maximum = 2100,
                Value = DateTime.Now.Year
            };
            numAno.ValueChanged += (s, e) => { if (chkCalcularPeriodoAutomatico.Checked) CalcularPeriodoAutomatico(); };
            toolTip.SetToolTip(lblAno, "Ano para cálculo automático do período.");
            toolTip.SetToolTip(numAno, "Ano para cálculo automático do período (2000-2100).");
            yPos += espacamentoVertical;

            // Linha 6: Período
            lblPeriodo = new Label { Text = "Período (YYYYMM - calculado automaticamente):", Location = new Point(10, yPos), Size = new Size(labelWidth, 20) };
            txtPeriodo = new TextBox { Location = new Point(campoX, yPos - 3), Size = new Size(200, 23), ReadOnly = true, BackColor = SystemColors.Control };
            txtPeriodo.MaxLength = 6;
            toolTip.SetToolTip(lblPeriodo, "Período semestral no formato YYYYMM (calculado automaticamente quando a opção 'Calcular período automaticamente' está marcada):\n• 01 = Primeiro semestre (Janeiro a Junho)\n• 02 = Segundo semestre (Julho a Dezembro)\nExemplos: 202301 (Jan-Jun/2023) ou 202302 (Jul-Dez/2023)");
            toolTip.SetToolTip(txtPeriodo, "Período semestral no formato YYYYMM.\nEste campo é calculado automaticamente quando a opção 'Calcular período automaticamente' está marcada.");
            txtPeriodo.Leave += (s, e) => 
            {
                if (!string.IsNullOrWhiteSpace(txtPeriodo.Text) && txtPeriodo.Text.Length == 6)
                {
                    // Validar formato YYYYMM
                    if (int.TryParse(txtPeriodo.Text, out int periodo))
                    {
                        int ano = periodo / 100;
                        int mes = periodo % 100;
                        
                        // Validar ano
                        if (ano < 2000 || ano > 2100)
                        {
                            MessageBox.Show($"Ano inválido: {ano}. Deve estar entre 2000 e 2100.", 
                                TITULO_VALIDACAO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            txtPeriodo.Focus();
                            return;
                        }
                        
                        // Validar mês: aceita 01, 02, 06 ou 12
                        // 01 ou 06 = Primeiro semestre (Jan-Jun)
                        // 02 ou 12 = Segundo semestre (Jul-Dez)
                        bool mesValido = (mes == 1 || mes == 6) || (mes == 2 || mes == 12);
                        if (!mesValido)
                        {
                            MessageBox.Show($"Mês inválido no período: {mes:00}.\n\n" +
                                $"Use:\n" +
                                $"  • 01 ou 06 = Primeiro semestre (Janeiro a Junho)\n" +
                                $"  • 02 ou 12 = Segundo semestre (Julho a Dezembro)\n\n" +
                                $"Exemplos: 202301 (Jan-Jun/2023) ou 202302 (Jul-Dez/2023)", 
                                TITULO_VALIDACAO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            txtPeriodo.Focus();
                            return;
                        }
                        
                        // Validação passou
                        string semestre = (mes == 1 || mes == 6) ? "Jan-Jun" : "Jul-Dez";
                        System.Diagnostics.Debug.WriteLine($"Período validado: {txtPeriodo.Text} = {semestre}/{ano}");
                    }
                    else
                    {
                        MessageBox.Show("Período inválido. Use apenas números no formato YYYYMM (ex: 202301 ou 202302).", 
                            "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        txtPeriodo.Focus();
                    }
                }
            };
            yPos += espacamentoVertical;

            // Linha 5: Diretório de Lotes
            lblDiretorioLotes = new Label { Text = "Diretório de Lotes:", Location = new Point(10, yPos), Size = new Size(labelWidth, 20) };
            txtDiretorioLotes = new TextBox { Location = new Point(campoX, yPos - 3), Size = new Size(campoWidth, 23) };
            btnSelecionarDiretorio = new Button { Text = "...", Location = new Point(campoX + campoWidth + 10, yPos - 3), Size = new Size(30, 25) };
            btnSelecionarDiretorio.Click += BtnSelecionarDiretorio_Click;
            toolTip.SetToolTip(lblDiretorioLotes, "Diretório onde os arquivos XML serão salvos.\nOs arquivos gerados (original, assinado e criptografado) serão salvos neste diretório.");
            toolTip.SetToolTip(txtDiretorioLotes, "Diretório onde os arquivos XML serão salvos.\nOs arquivos gerados (original, assinado e criptografado) serão salvos neste diretório.");
            toolTip.SetToolTip(btnSelecionarDiretorio, "Abre o diálogo para selecionar o diretório onde os lotes serão salvos.");
            yPos += espacamentoVertical;

            // Linha 6: Ambiente e Opções
            lblAmbiente = new Label { Text = "Ambiente:", Location = new Point(10, yPos), Size = new Size(labelWidth, 20) };
            cmbAmbiente = new ComboBox { Location = new Point(campoX, yPos - 3), Size = new Size(100, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbAmbiente.Items.AddRange(new[] { "TEST", "PROD" });
            cmbAmbiente.SelectedIndex = 0;
            toolTip.SetToolTip(lblAmbiente, "Ambiente de trabalho:\n• TEST = Ambiente de testes/homologação\n• PROD = Ambiente de produção");
            toolTip.SetToolTip(cmbAmbiente, "Ambiente de trabalho:\n• TEST = Ambiente de testes/homologação\n• PROD = Ambiente de produção");

            chkModoTeste = new CheckBox { Text = "Modo Teste", Location = new Point(campoX + 120, yPos), Size = new Size(120, 20), Checked = true };
            toolTip.SetToolTip(chkModoTeste, "Quando marcado, usa valores menores para processamento (menos registros por página, menos eventos por lote).\nÚtil para testes rápidos sem processar muitos dados.");
            
            chkHabilitarEnvio = new CheckBox { Text = "Habilitar Envio", Location = new Point(campoX + 250, yPos), Size = new Size(150, 20) };
            toolTip.SetToolTip(chkHabilitarEnvio, "Quando marcado, permite o envio automático dos lotes para a e-Financeira após processamento.\nSe desmarcado, apenas processa e salva os arquivos localmente.");
            yPos += espacamentoVertical;

            // Linha 7: Botões
            btnCarregarConfig = new Button { Text = "Carregar Configuração", Location = new Point(10, yPos), Size = new Size(150, 30) };
            btnCarregarConfig.Click += BtnCarregarConfig_Click;

            btnTestarConexao = new Button { Text = "Testar Conexão BD", Location = new Point(170, yPos), Size = new Size(150, 30) };
            btnTestarConexao.Click += BtnTestarConexao_Click;

            btnSalvarConfig = new Button { Text = "Salvar Configurações", Location = new Point(330, yPos), Size = new Size(150, 30) };
            btnSalvarConfig.Click += BtnSalvarConfig_Click;

            btnLimparDadosTeste = new Button { Text = "Limpar Dados de Teste", Location = new Point(490, yPos), Size = new Size(150, 30) };
            btnLimparDadosTeste.Click += BtnLimparDadosTeste_Click;
            toolTip.SetToolTip(btnLimparDadosTeste, "Limpa os dados de teste no ambiente de Produção Restrita da e-Financeira.\nRequer certificado válido e CNPJ configurado.");

            grpConfigGeral.Controls.AddRange(new Control[] {
                lblCnpjDeclarante, txtCnpjDeclarante,
                lblCertThumbprint, txtCertThumbprint, btnSelecionarCert,
                lblCertServidorThumbprint, txtCertServidorThumbprint, btnSelecionarCertServidor,
                chkCalcularPeriodoAutomatico,
                lblSemestre, cmbSemestre, lblAno, numAno,
                lblPeriodo, txtPeriodo,
                lblDiretorioLotes, txtDiretorioLotes, btnSelecionarDiretorio,
                lblAmbiente, cmbAmbiente, chkModoTeste, chkHabilitarEnvio,
                btnCarregarConfig, btnTestarConexao, btnSalvarConfig, btnLimparDadosTeste
            });
            
            // Calcular período inicial
            CalcularPeriodoAutomatico();

            // Configurações de Processamento
            CriarGrupoProcessamento();

            // TabControl
            tabConfig = new TabControl();
            tabConfig.Location = new Point(10, 530);
            tabConfig.Size = new Size(1230, 250);
            tabConfig.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            tabAbertura = new TabPage { Text = "Abertura", UseVisualStyleBackColor = true };
            tabFechamento = new TabPage { Text = "Fechamento", UseVisualStyleBackColor = true };
            tabCadastroDeclarante = new TabPage { Text = "Cadastro Declarante", UseVisualStyleBackColor = true };
            tabConfig.TabPages.Add(tabAbertura);
            tabConfig.TabPages.Add(tabFechamento);
            tabConfig.TabPages.Add(tabCadastroDeclarante);

            CriarAbaAbertura();
            CriarAbaFechamento();
            CriarAbaCadastroDeclarante();

            // Evento para ajustar valores padrão quando ambiente mudar
            cmbAmbiente.SelectedIndexChanged += CmbAmbiente_SelectedIndexChanged;

            this.Controls.AddRange(new Control[] { grpConfigGeral, grpProcessamento, tabConfig });
            this.ResumeLayout(false);
            
            // Aplicar valores padrão após todos os controles serem criados
            if (grpProcessamento != null && numPageSize != null)
            {
                CmbAmbiente_SelectedIndexChanged(null, null);
            }
        }

        private void CriarGrupoProcessamento()
        {
            grpProcessamento = new GroupBox();
            grpProcessamento.Text = "Configurações de Processamento";
            grpProcessamento.Location = new Point(10, 420);
            grpProcessamento.Size = new Size(1230, 100);
            grpProcessamento.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            int xInicio = 10;
            int yInicio = 25;
            int espacamento = 240;

            // Page Size
            var lblPageSize = new Label { Text = "Page Size:", Location = new Point(xInicio, yInicio), Size = new Size(100, 20) };
            numPageSize = new NumericUpDown { Location = new Point(xInicio + 105, yInicio - 3), Size = new Size(80, 23), Minimum = 1, Maximum = 10000, Value = 500 };
            var btnHelpPageSize = CriarBotaoAjuda(xInicio + 190, yInicio - 3);
            toolTip.SetToolTip(btnHelpPageSize, 
                "Tamanho da página para consultas ao banco de dados.\n\n" +
                "Produção: 500 ou maior (conforme memória disponível)\n" +
                "Teste: 50-100 (para processar menos dados)\n\n" +
                "Valores maiores melhoram performance mas consomem mais memória.");

            // Evento Offset
            var lblEventoOffset = new Label { Text = "Evento Offset:", Location = new Point(xInicio + espacamento, yInicio), Size = new Size(100, 20) };
            numEventoOffset = new NumericUpDown { Location = new Point(xInicio + espacamento + 105, yInicio - 3), Size = new Size(80, 23), Minimum = 0, Maximum = 10000, Value = 1 };
            var btnHelpEventoOffset = CriarBotaoAjuda(xInicio + espacamento + 190, yInicio - 3);
            toolTip.SetToolTip(btnHelpEventoOffset,
                "Controle de onde começar a gerar eventos.\n\n" +
                "Produção: 0 ou 1 (começar do início)\n" +
                "Teste: Pode ser maior para pular eventos iniciais\n\n" +
                "Útil para retomar processamento ou testar eventos específicos.");

            // Offset Registros
            var lblOffsetRegistros = new Label { Text = "Offset Registros:", Location = new Point(xInicio + espacamento * 2, yInicio), Size = new Size(120, 20) };
            numOffsetRegistros = new NumericUpDown { Location = new Point(xInicio + espacamento * 2 + 125, yInicio - 3), Size = new Size(80, 23), Minimum = 0, Maximum = 1000000, Value = 0 };
            var btnHelpOffsetRegistros = CriarBotaoAjuda(xInicio + espacamento * 2 + 210, yInicio - 3);
            toolTip.SetToolTip(btnHelpOffsetRegistros,
                "Pular registros iniciais nas consultas ao banco.\n\n" +
                "Produção: 0 (não usar)\n" +
                "Teste: Pode ter valor para pular registros já processados\n\n" +
                "Útil apenas em modo de teste para processar registros específicos.");

            // Max Lotes
            var lblMaxLotes = new Label { Text = "Max Lotes:", Location = new Point(xInicio, yInicio + 35), Size = new Size(100, 20) };
            numMaxLotes = new NumericUpDown { Location = new Point(xInicio + 105, yInicio + 32), Size = new Size(80, 23), Minimum = 1, Maximum = 10000, Value = 1, Enabled = true };
            chkMaxLotesIlimitado = new CheckBox { Text = "Ilimitado", Location = new Point(xInicio + 190, yInicio + 35), Size = new Size(80, 20), Checked = false };
            chkMaxLotesIlimitado.CheckedChanged += (s, e) => 
            {
                numMaxLotes.Enabled = !chkMaxLotesIlimitado.Checked;
                if (chkMaxLotesIlimitado.Checked) numMaxLotes.Value = 1;
            };
            var btnHelpMaxLotes = CriarBotaoAjuda(xInicio + 275, yInicio + 32);
            toolTip.SetToolTip(btnHelpMaxLotes,
                "Limitar quantidade de lotes gerados.\n\n" +
                "Produção: Ilimitado ou valor alto (não limitar)\n" +
                "Teste: Valor baixo (ex: 1-5) para processar poucos lotes\n\n" +
                "Útil em modo de teste para processar apenas alguns lotes.");

            // Eventos por Lote
            var lblEventosPorLote = new Label { Text = "Eventos por Lote:", Location = new Point(xInicio + espacamento * 2, yInicio + 35), Size = new Size(120, 20) };
            numEventosPorLote = new NumericUpDown { Location = new Point(xInicio + espacamento * 2 + 125, yInicio + 32), Size = new Size(80, 23), Minimum = 1, Maximum = 50, Value = 50 };
            var btnHelpEventosPorLote = CriarBotaoAjuda(xInicio + espacamento * 2 + 210, yInicio + 32);
            toolTip.SetToolTip(btnHelpEventosPorLote,
                "Quantidade de eventos por lote.\n\n" +
                "Produção: 50 (máximo permitido pelo e-Financeira)\n" +
                "Teste: 1-10 (para gerar lotes menores e testar mais rapidamente)\n\n" +
                "Valores válidos: 1 a 50 eventos por lote (conforme manual e-Financeira).\n" +
                "Para gerar um lote com apenas 1 evento, defina este valor como 1.");

            grpProcessamento.Controls.AddRange(new Control[] {
                lblPageSize, numPageSize, btnHelpPageSize,
                lblEventoOffset, numEventoOffset, btnHelpEventoOffset,
                lblOffsetRegistros, numOffsetRegistros, btnHelpOffsetRegistros,
                lblMaxLotes, numMaxLotes, chkMaxLotesIlimitado, btnHelpMaxLotes,
                lblEventosPorLote, numEventosPorLote, btnHelpEventosPorLote
            });
        }

        private Button CriarBotaoAjuda(int x, int y)
        {
            var btn = new Button
            {
                Text = "?",
                Location = new Point(x, y),
                Size = new Size(25, 23),
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold),
                BackColor = Color.LightBlue,
                FlatStyle = FlatStyle.Flat
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void CmbAmbiente_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool isProducao = cmbAmbiente.SelectedItem.ToString() == "PROD";
            
            // Desabilitar botão de limpar dados se não for ambiente de testes
            if (btnLimparDadosTeste != null)
            {
                btnLimparDadosTeste.Enabled = !isProducao;
            }
            
            if (isProducao)
            {
                // Valores padrão para produção
                numPageSize.Value = 500;
                numEventoOffset.Value = 1;
                numOffsetRegistros.Value = 0;
                chkMaxLotesIlimitado.Checked = true;
                if (numEventosPorLote != null) numEventosPorLote.Value = 50;
            }
            else
            {
                // Valores padrão para teste
                numPageSize.Value = 50;
                numEventoOffset.Value = 1;
                numOffsetRegistros.Value = 0;
                chkMaxLotesIlimitado.Checked = false;
                numMaxLotes.Value = 1;
                if (numEventosPorLote != null) numEventosPorLote.Value = 1; // Teste: 1 evento por lote
            }
        }

        private void CriarAbaAbertura()
        {
            scrollAbertura = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            int yPos = 10;

            // Dados Básicos
            var grpBasicos = new GroupBox { Text = "Dados Básicos", Location = new Point(10, yPos), Size = new Size(scrollAbertura.Width - 30, 180) };
            grpBasicos.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            yPos += 190;

            // Cálculo Automático de Período
            chkCalcularPeriodoAbertura = new CheckBox 
            { 
                Text = "Calcular período automaticamente", 
                Location = new Point(10, 25), 
                Size = new Size(250, 20),
                Checked = true
            };
            toolTip.SetToolTip(chkCalcularPeriodoAbertura, "Quando marcado, calcula automaticamente as datas de início e fim baseado no semestre e ano selecionados.\nOs campos de data ficam somente leitura.");
            chkCalcularPeriodoAbertura.CheckedChanged += (s, e) =>
            {
                bool calcularAutomatico = chkCalcularPeriodoAbertura.Checked;
                cmbSemestreAbertura.Enabled = calcularAutomatico;
                numAnoAbertura.Enabled = calcularAutomatico;
                txtDtInicioAbertura.ReadOnly = calcularAutomatico;
                txtDtFimAbertura.ReadOnly = calcularAutomatico;
                txtDtInicioAbertura.BackColor = calcularAutomatico ? SystemColors.Control : SystemColors.Window;
                txtDtFimAbertura.BackColor = calcularAutomatico ? SystemColors.Control : SystemColors.Window;
                
                if (calcularAutomatico)
                {
                    CalcularPeriodoAbertura();
                }
            };

            var lblSemestreAbertura = new Label { Text = "Semestre:", Location = new Point(10, 50), Size = new Size(100, 20) };
            cmbSemestreAbertura = new ComboBox 
            { 
                Location = new Point(115, 47), 
                Size = new Size(150, 23), 
                DropDownStyle = ComboBoxStyle.DropDownList 
            };
            cmbSemestreAbertura.Items.AddRange(new[] { "1º Semestre (Jan-Jun)", "2º Semestre (Jul-Dez)" });
            cmbSemestreAbertura.SelectedIndex = 0;
            cmbSemestreAbertura.SelectedIndexChanged += (s, e) => { if (chkCalcularPeriodoAbertura.Checked) CalcularPeriodoAbertura(); };
            toolTip.SetToolTip(lblSemestreAbertura, "Selecione o semestre para cálculo automático das datas.");
            toolTip.SetToolTip(cmbSemestreAbertura, "Selecione o semestre:\n• 1º Semestre = Janeiro a Junho (01/01 a 30/06)\n• 2º Semestre = Julho a Dezembro (01/07 a 31/12)");

            var lblAnoAbertura = new Label { Text = "Ano:", Location = new Point(275, 50), Size = new Size(50, 20) };
            numAnoAbertura = new NumericUpDown 
            { 
                Location = new Point(330, 47), 
                Size = new Size(80, 23),
                Minimum = 2000,
                Maximum = 2100,
                Value = DateTime.Now.Year
            };
            numAnoAbertura.ValueChanged += (s, e) => { if (chkCalcularPeriodoAbertura.Checked) CalcularPeriodoAbertura(); };
            toolTip.SetToolTip(lblAnoAbertura, "Ano para cálculo automático das datas.");
            toolTip.SetToolTip(numAnoAbertura, "Ano para cálculo automático das datas (2000-2100).");

            var lblDtInicio = new Label { Text = "Data Início (AAAA-MM-DD):", Location = new Point(10, 80), Size = new Size(150, 20) };
            txtDtInicioAbertura = new TextBox { Location = new Point(165, 77), Size = new Size(150, 23), ReadOnly = true, BackColor = SystemColors.Control };
            txtDtInicioAbertura.Leave += ValidarData;
            toolTip.SetToolTip(lblDtInicio, "Data de início do período de abertura no formato AAAA-MM-DD.\nCalculada automaticamente quando a opção 'Calcular período automaticamente' está marcada.\nExemplo: 2023-01-01 (1º semestre) ou 2023-07-01 (2º semestre)");
            toolTip.SetToolTip(txtDtInicioAbertura, "Data de início do período de abertura no formato AAAA-MM-DD.\nCalculada automaticamente quando a opção 'Calcular período automaticamente' está marcada.");

            var lblDtFim = new Label { Text = "Data Fim (AAAA-MM-DD):", Location = new Point(330, 80), Size = new Size(150, 20) };
            txtDtFimAbertura = new TextBox { Location = new Point(485, 77), Size = new Size(150, 23), ReadOnly = true, BackColor = SystemColors.Control };
            txtDtFimAbertura.Leave += ValidarData;
            toolTip.SetToolTip(lblDtFim, "Data de fim do período de abertura no formato AAAA-MM-DD.\nCalculada automaticamente quando a opção 'Calcular período automaticamente' está marcada.\nExemplo: 2023-06-30 (1º semestre) ou 2023-12-31 (2º semestre)");
            toolTip.SetToolTip(txtDtFimAbertura, "Data de fim do período de abertura no formato AAAA-MM-DD.\nCalculada automaticamente quando a opção 'Calcular período automaticamente' está marcada.");

            var lblTipoAmbiente = new Label { Text = "Tipo Ambiente:", Location = new Point(650, 80), Size = new Size(100, 20) };
            cmbTipoAmbienteAbertura = new ComboBox { Location = new Point(755, 77), Size = new Size(150, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTipoAmbienteAbertura.Items.AddRange(new[] { "1 - Produção", "2 - Homologação" });
            cmbTipoAmbienteAbertura.SelectedIndex = 1;
            toolTip.SetToolTip(lblTipoAmbiente, TOOLTIP_TIPO_AMBIENTE);
            toolTip.SetToolTip(cmbTipoAmbienteAbertura, TOOLTIP_TIPO_AMBIENTE);

            var lblAplicacaoEmissora = new Label { Text = "Aplicação Emissora:", Location = new Point(920, 80), Size = new Size(120, 20) };
            cmbAplicacaoEmissoraAbertura = new ComboBox { Location = new Point(1045, 77), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbAplicacaoEmissoraAbertura.Items.AddRange(new[] { "1 - Aplicação do contribuinte", "2 - Outros" });
            cmbAplicacaoEmissoraAbertura.SelectedIndex = 0;
            toolTip.SetToolTip(lblAplicacaoEmissora, TOOLTIP_APLICACAO_EMISSORA);
            toolTip.SetToolTip(cmbAplicacaoEmissoraAbertura, TOOLTIP_APLICACAO_EMISSORA);

            var lblIndRetificacao = new Label { Text = "Ind Retificação:", Location = new Point(650, 110), Size = new Size(100, 20) };
            cmbIndRetificacaoAbertura = new ComboBox { Location = new Point(755, 107), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbIndRetificacaoAbertura.Items.AddRange(new[] { "1 - Original", "2 - Retificação espontânea", "3 - Retificação a pedido" });
            cmbIndRetificacaoAbertura.SelectedIndex = 0;
            cmbIndRetificacaoAbertura.SelectedIndexChanged += CmbIndRetificacaoAbertura_SelectedIndexChanged;
            toolTip.SetToolTip(lblIndRetificacao, TOOLTIP_IND_RETIFICACAO);
            toolTip.SetToolTip(cmbIndRetificacaoAbertura, TOOLTIP_IND_RETIFICACAO);

            var lblNrRecibo = new Label { Text = "Nº Recibo:", Location = new Point(970, 110), Size = new Size(80, 20) };
            txtNrReciboAbertura = new TextBox { Location = new Point(1055, 107), Size = new Size(150, 23), Enabled = false };
            toolTip.SetToolTip(lblNrRecibo, TOOLTIP_NR_RECIBO);
            toolTip.SetToolTip(txtNrReciboAbertura, TOOLTIP_NR_RECIBO);

            chkIndicarMovOpFin = new CheckBox { Text = "Indicar MovOpFin", Location = new Point(650, 140), Size = new Size(200, 20) };
            toolTip.SetToolTip(chkIndicarMovOpFin, "Marque esta opção se a empresa possui movimentações de operações financeiras a declarar.\nSe marcado, indica que haverá eventos de movimentação financeira.");

            grpBasicos.Controls.AddRange(new Control[] {
                chkCalcularPeriodoAbertura,
                lblSemestreAbertura, cmbSemestreAbertura, lblAnoAbertura, numAnoAbertura,
                lblDtInicio, txtDtInicioAbertura, lblDtFim, txtDtFimAbertura,
                lblTipoAmbiente, cmbTipoAmbienteAbertura, lblAplicacaoEmissora, cmbAplicacaoEmissoraAbertura,
                lblIndRetificacao, cmbIndRetificacaoAbertura, lblNrRecibo, txtNrReciboAbertura,
                chkIndicarMovOpFin
            });
            
            // Calcular período inicial
            CalcularPeriodoAbertura();

            // ResponsavelRMF
            var grpRMF = new GroupBox { Text = "Responsável RMF", Location = new Point(10, yPos), Size = new Size(scrollAbertura.Width - 30, 200) };
            grpRMF.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            yPos += 210;

            CriarCamposResponsavelRMF(grpRMF, 10);

            // RespeFin
            var grpRespeFin = new GroupBox { Text = "Responsável e-Financeira", Location = new Point(10, yPos), Size = new Size(scrollAbertura.Width - 30, 220) };
            grpRespeFin.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            yPos += 230;

            CriarCamposRespeFin(grpRespeFin, 10);

            // RepresLegal
            var grpRepresLegal = new GroupBox { Text = "Representante Legal", Location = new Point(10, yPos), Size = new Size(scrollAbertura.Width - 30, 140) };
            grpRepresLegal.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            CriarCamposRepresLegal(grpRepresLegal, 10);

            scrollAbertura.Controls.AddRange(new Control[] { grpBasicos, grpRMF, grpRespeFin, grpRepresLegal });
            tabAbertura.Controls.Add(scrollAbertura);
        }

        private void CriarCamposResponsavelRMF(GroupBox grp, int xInicio)
        {
            int y = 25;
            int x = xInicio;

            var lblCNPJ = new Label { Text = "CNPJ:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRMF_CNPJ = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(150, 23) };
            txtRMF_CNPJ.Leave += ValidarCNPJ;
            toolTip.SetToolTip(lblCNPJ, "CNPJ do responsável RMF (sem pontos, barras ou hífens).\nObrigatório se for pessoa jurídica.");
            toolTip.SetToolTip(txtRMF_CNPJ, "CNPJ do responsável RMF (sem pontos, barras ou hífens).\nObrigatório se for pessoa jurídica.");

            var lblCPF = new Label { Text = "CPF:", Location = new Point(x + 250, y), Size = new Size(80, 20) };
            txtRMF_CPF = new TextBox { Location = new Point(x + 335, y - 3), Size = new Size(150, 23) };
            txtRMF_CPF.Leave += ValidarCPF;
            toolTip.SetToolTip(lblCPF, "CPF do responsável RMF (sem pontos ou hífens).\nObrigatório se for pessoa física.");
            toolTip.SetToolTip(txtRMF_CPF, "CPF do responsável RMF (sem pontos ou hífens).\nObrigatório se for pessoa física.");

            var lblNome = new Label { Text = "Nome:", Location = new Point(x + 500, y), Size = new Size(80, 20) };
            txtRMF_Nome = new TextBox { Location = new Point(x + 585, y - 3), Size = new Size(300, 23) };
            toolTip.SetToolTip(lblNome, "Nome completo do responsável RMF.");
            toolTip.SetToolTip(txtRMF_Nome, "Nome completo do responsável RMF.");

            y += 35;
            var lblSetor = new Label { Text = "Setor:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRMF_Setor = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(200, 23) };
            toolTip.SetToolTip(lblSetor, "Setor/departamento do responsável RMF na empresa.");
            toolTip.SetToolTip(txtRMF_Setor, "Setor/departamento do responsável RMF na empresa.");

            var lblDDD = new Label { Text = "DDD:", Location = new Point(x + 300, y), Size = new Size(50, 20) };
            txtRMF_DDD = new TextBox { Location = new Point(x + 355, y - 3), Size = new Size(50, 23) };
            toolTip.SetToolTip(lblDDD, "DDD do telefone do responsável RMF (apenas números).\nExemplo: 11");
            toolTip.SetToolTip(txtRMF_DDD, "DDD do telefone do responsável RMF (apenas números).\nExemplo: 11");

            var lblTelefone = new Label { Text = "Telefone:", Location = new Point(x + 420, y), Size = new Size(70, 20) };
            txtRMF_Telefone = new TextBox { Location = new Point(x + 495, y - 3), Size = new Size(100, 23) };
            toolTip.SetToolTip(lblTelefone, "Número do telefone do responsável RMF (apenas números).\nExemplo: 987654321");
            toolTip.SetToolTip(txtRMF_Telefone, "Número do telefone do responsável RMF (apenas números).\nExemplo: 987654321");

            var lblRamal = new Label { Text = "Ramal:", Location = new Point(x + 610, y), Size = new Size(60, 20) };
            txtRMF_Ramal = new TextBox { Location = new Point(x + 675, y - 3), Size = new Size(80, 23) };
            toolTip.SetToolTip(lblRamal, "Ramal do telefone do responsável RMF (opcional).");
            toolTip.SetToolTip(txtRMF_Ramal, "Ramal do telefone do responsável RMF (opcional).");

            y += 35;
            var lblLogradouro = new Label { Text = "Logradouro:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRMF_Logradouro = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(300, 23) };
            toolTip.SetToolTip(lblLogradouro, "Nome da rua, avenida ou logradouro do endereço do responsável RMF.");
            toolTip.SetToolTip(txtRMF_Logradouro, "Nome da rua, avenida ou logradouro do endereço do responsável RMF.");

            var lblNumero = new Label { Text = "Número:", Location = new Point(x + 400, y), Size = new Size(60, 20) };
            txtRMF_Numero = new TextBox { Location = new Point(x + 465, y - 3), Size = new Size(80, 23) };
            toolTip.SetToolTip(lblNumero, "Número do endereço do responsável RMF.");
            toolTip.SetToolTip(txtRMF_Numero, "Número do endereço do responsável RMF.");

            var lblComplemento = new Label { Text = "Complemento:", Location = new Point(x + 560, y), Size = new Size(90, 20) };
            txtRMF_Complemento = new TextBox { Location = new Point(x + 655, y - 3), Size = new Size(200, 23) };
            toolTip.SetToolTip(lblComplemento, TOOLTIP_COMPLEMENTO_ENDERECO);
            toolTip.SetToolTip(txtRMF_Complemento, TOOLTIP_COMPLEMENTO_ENDERECO);

            y += 35;
            var lblBairro = new Label { Text = "Bairro:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRMF_Bairro = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(200, 23) };
            toolTip.SetToolTip(lblBairro, "Bairro do endereço do responsável RMF.");
            toolTip.SetToolTip(txtRMF_Bairro, "Bairro do endereço do responsável RMF.");

            var lblCEP = new Label { Text = "CEP:", Location = new Point(x + 300, y), Size = new Size(50, 20) };
            txtRMF_CEP = new TextBox { Location = new Point(x + 355, y - 3), Size = new Size(100, 23) };
            txtRMF_CEP.Leave += ValidarCEP;
            toolTip.SetToolTip(lblCEP, "CEP do endereço do responsável RMF (com ou sem hífen).\nExemplo: 01310-100 ou 01310100");
            toolTip.SetToolTip(txtRMF_CEP, "CEP do endereço do responsável RMF (com ou sem hífen).\nExemplo: 01310-100 ou 01310100");

            var lblMunicipio = new Label { Text = "Município:", Location = new Point(x + 470, y), Size = new Size(70, 20) };
            txtRMF_Municipio = new TextBox { Location = new Point(x + 545, y - 3), Size = new Size(200, 23) };
            toolTip.SetToolTip(lblMunicipio, "Nome do município do endereço do responsável RMF.");
            toolTip.SetToolTip(txtRMF_Municipio, "Nome do município do endereço do responsável RMF.");

            var lblUF = new Label { Text = "UF:", Location = new Point(x + 760, y), Size = new Size(30, 20) };
            txtRMF_UF = new TextBox { Location = new Point(x + 795, y - 3), Size = new Size(50, 23), MaxLength = 2 };
            txtRMF_UF.CharacterCasing = CharacterCasing.Upper;
            toolTip.SetToolTip(lblUF, "Sigla do estado (UF) do endereço do responsável RMF.\nExemplo: SP, RJ, MG");
            toolTip.SetToolTip(txtRMF_UF, "Sigla do estado (UF) do endereço do responsável RMF.\nExemplo: SP, RJ, MG");

            grp.Controls.AddRange(new Control[] {
                lblCNPJ, txtRMF_CNPJ, lblCPF, txtRMF_CPF, lblNome, txtRMF_Nome,
                lblSetor, txtRMF_Setor, lblDDD, txtRMF_DDD, lblTelefone, txtRMF_Telefone, lblRamal, txtRMF_Ramal,
                lblLogradouro, txtRMF_Logradouro, lblNumero, txtRMF_Numero, lblComplemento, txtRMF_Complemento,
                lblBairro, txtRMF_Bairro, lblCEP, txtRMF_CEP, lblMunicipio, txtRMF_Municipio, lblUF, txtRMF_UF
            });
        }

        private void CriarCamposRespeFin(GroupBox grp, int xInicio)
        {
            int y = 25;
            int x = xInicio;

            var lblCPF = new Label { Text = "CPF:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRespeFin_CPF = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(150, 23) };
            txtRespeFin_CPF.Leave += ValidarCPF;
            toolTip.SetToolTip(lblCPF, "CPF do responsável pela e-Financeira (sem pontos ou hífens).\nObrigatório.");
            toolTip.SetToolTip(txtRespeFin_CPF, "CPF do responsável pela e-Financeira (sem pontos ou hífens).\nObrigatório.");

            var lblNome = new Label { Text = "Nome:", Location = new Point(x + 250, y), Size = new Size(80, 20) };
            txtRespeFin_Nome = new TextBox { Location = new Point(x + 335, y - 3), Size = new Size(300, 23) };
            toolTip.SetToolTip(lblNome, "Nome completo do responsável pela e-Financeira.\nObrigatório.");
            toolTip.SetToolTip(txtRespeFin_Nome, "Nome completo do responsável pela e-Financeira.\nObrigatório.");

            var lblEmail = new Label { Text = "Email:", Location = new Point(x + 650, y), Size = new Size(60, 20) };
            txtRespeFin_Email = new TextBox { Location = new Point(x + 715, y - 3), Size = new Size(200, 23) };
            txtRespeFin_Email.Leave += ValidarEmail;
            toolTip.SetToolTip(lblEmail, "Email do responsável pela e-Financeira.\nFormato: nome@dominio.com");
            toolTip.SetToolTip(txtRespeFin_Email, "Email do responsável pela e-Financeira.\nFormato: nome@dominio.com");

            y += 35;
            var lblSetor = new Label { Text = "Setor:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRespeFin_Setor = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(200, 23) };
            toolTip.SetToolTip(lblSetor, "Setor/departamento do responsável pela e-Financeira na empresa.");
            toolTip.SetToolTip(txtRespeFin_Setor, "Setor/departamento do responsável pela e-Financeira na empresa.");

            var lblDDD = new Label { Text = "DDD:", Location = new Point(x + 300, y), Size = new Size(50, 20) };
            txtRespeFin_DDD = new TextBox { Location = new Point(x + 355, y - 3), Size = new Size(50, 23) };
            toolTip.SetToolTip(lblDDD, "DDD do telefone do responsável pela e-Financeira (apenas números).\nExemplo: 11");
            toolTip.SetToolTip(txtRespeFin_DDD, "DDD do telefone do responsável pela e-Financeira (apenas números).\nExemplo: 11");

            var lblTelefone = new Label { Text = "Telefone:", Location = new Point(x + 420, y), Size = new Size(70, 20) };
            txtRespeFin_Telefone = new TextBox { Location = new Point(x + 495, y - 3), Size = new Size(100, 23) };
            toolTip.SetToolTip(lblTelefone, "Número do telefone do responsável pela e-Financeira (apenas números).\nExemplo: 987654321");
            toolTip.SetToolTip(txtRespeFin_Telefone, "Número do telefone do responsável pela e-Financeira (apenas números).\nExemplo: 987654321");

            var lblRamal = new Label { Text = "Ramal:", Location = new Point(x + 610, y), Size = new Size(60, 20) };
            txtRespeFin_Ramal = new TextBox { Location = new Point(x + 675, y - 3), Size = new Size(80, 23) };
            toolTip.SetToolTip(lblRamal, "Ramal do telefone do responsável pela e-Financeira (opcional).");
            toolTip.SetToolTip(txtRespeFin_Ramal, "Ramal do telefone do responsável pela e-Financeira (opcional).");

            y += 35;
            var lblLogradouro = new Label { Text = "Logradouro:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRespeFin_Logradouro = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(300, 23) };
            toolTip.SetToolTip(lblLogradouro, "Nome da rua, avenida ou logradouro do endereço do responsável pela e-Financeira.");
            toolTip.SetToolTip(txtRespeFin_Logradouro, "Nome da rua, avenida ou logradouro do endereço do responsável pela e-Financeira.");

            var lblNumero = new Label { Text = "Número:", Location = new Point(x + 400, y), Size = new Size(60, 20) };
            txtRespeFin_Numero = new TextBox { Location = new Point(x + 465, y - 3), Size = new Size(80, 23) };
            toolTip.SetToolTip(lblNumero, "Número do endereço do responsável pela e-Financeira.");
            toolTip.SetToolTip(txtRespeFin_Numero, "Número do endereço do responsável pela e-Financeira.");

            var lblComplemento = new Label { Text = "Complemento:", Location = new Point(x + 560, y), Size = new Size(90, 20) };
            txtRespeFin_Complemento = new TextBox { Location = new Point(x + 655, y - 3), Size = new Size(200, 23) };
            toolTip.SetToolTip(lblComplemento, TOOLTIP_COMPLEMENTO_ENDERECO);
            toolTip.SetToolTip(txtRespeFin_Complemento, TOOLTIP_COMPLEMENTO_ENDERECO);

            y += 35;
            var lblBairro = new Label { Text = "Bairro:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRespeFin_Bairro = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(200, 23) };
            toolTip.SetToolTip(lblBairro, "Bairro do endereço do responsável pela e-Financeira.");
            toolTip.SetToolTip(txtRespeFin_Bairro, "Bairro do endereço do responsável pela e-Financeira.");

            var lblCEP = new Label { Text = "CEP:", Location = new Point(x + 300, y), Size = new Size(50, 20) };
            txtRespeFin_CEP = new TextBox { Location = new Point(x + 355, y - 3), Size = new Size(100, 23) };
            txtRespeFin_CEP.Leave += ValidarCEP;
            toolTip.SetToolTip(lblCEP, "CEP do endereço do responsável pela e-Financeira (com ou sem hífen).\nExemplo: 01310-100 ou 01310100");
            toolTip.SetToolTip(txtRespeFin_CEP, "CEP do endereço do responsável pela e-Financeira (com ou sem hífen).\nExemplo: 01310-100 ou 01310100");

            var lblMunicipio = new Label { Text = "Município:", Location = new Point(x + 470, y), Size = new Size(70, 20) };
            txtRespeFin_Municipio = new TextBox { Location = new Point(x + 545, y - 3), Size = new Size(200, 23) };
            toolTip.SetToolTip(lblMunicipio, "Nome do município do endereço do responsável pela e-Financeira.");
            toolTip.SetToolTip(txtRespeFin_Municipio, "Nome do município do endereço do responsável pela e-Financeira.");

            var lblUF = new Label { Text = "UF:", Location = new Point(x + 760, y), Size = new Size(30, 20) };
            txtRespeFin_UF = new TextBox { Location = new Point(x + 795, y - 3), Size = new Size(50, 23), MaxLength = 2 };
            txtRespeFin_UF.CharacterCasing = CharacterCasing.Upper;
            toolTip.SetToolTip(lblUF, "Sigla do estado (UF) do endereço do responsável pela e-Financeira.\nExemplo: SP, RJ, MG");
            toolTip.SetToolTip(txtRespeFin_UF, "Sigla do estado (UF) do endereço do responsável pela e-Financeira.\nExemplo: SP, RJ, MG");

            grp.Controls.AddRange(new Control[] {
                lblCPF, txtRespeFin_CPF, lblNome, txtRespeFin_Nome, lblEmail, txtRespeFin_Email,
                lblSetor, txtRespeFin_Setor, lblDDD, txtRespeFin_DDD, lblTelefone, txtRespeFin_Telefone, lblRamal, txtRespeFin_Ramal,
                lblLogradouro, txtRespeFin_Logradouro, lblNumero, txtRespeFin_Numero, lblComplemento, txtRespeFin_Complemento,
                lblBairro, txtRespeFin_Bairro, lblCEP, txtRespeFin_CEP, lblMunicipio, txtRespeFin_Municipio, lblUF, txtRespeFin_UF
            });
        }

        private void CriarCamposRepresLegal(GroupBox grp, int xInicio)
        {
            int y = 25;
            int x = xInicio;

            var lblCPF = new Label { Text = "CPF:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRepresLegal_CPF = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(150, 23) };
            txtRepresLegal_CPF.Leave += ValidarCPF;
            toolTip.SetToolTip(lblCPF, "CPF do representante legal (sem pontos ou hífens).\nObrigatório.");
            toolTip.SetToolTip(txtRepresLegal_CPF, "CPF do representante legal (sem pontos ou hífens).\nObrigatório.");

            var lblSetor = new Label { Text = "Setor:", Location = new Point(x + 250, y), Size = new Size(80, 20) };
            txtRepresLegal_Setor = new TextBox { Location = new Point(x + 335, y - 3), Size = new Size(200, 23) };
            toolTip.SetToolTip(lblSetor, "Setor/departamento do representante legal na empresa.");
            toolTip.SetToolTip(txtRepresLegal_Setor, "Setor/departamento do representante legal na empresa.");

            y += 35;
            var lblDDD = new Label { Text = "DDD:", Location = new Point(x, y), Size = new Size(50, 20) };
            txtRepresLegal_DDD = new TextBox { Location = new Point(x + 55, y - 3), Size = new Size(50, 23) };
            toolTip.SetToolTip(lblDDD, "DDD do telefone do representante legal (apenas números).\nExemplo: 11");
            toolTip.SetToolTip(txtRepresLegal_DDD, "DDD do telefone do representante legal (apenas números).\nExemplo: 11");

            var lblTelefone = new Label { Text = "Telefone:", Location = new Point(x + 120, y), Size = new Size(70, 20) };
            txtRepresLegal_Telefone = new TextBox { Location = new Point(x + 195, y - 3), Size = new Size(100, 23) };
            toolTip.SetToolTip(lblTelefone, "Número do telefone do representante legal (apenas números).\nExemplo: 987654321");
            toolTip.SetToolTip(txtRepresLegal_Telefone, "Número do telefone do representante legal (apenas números).\nExemplo: 987654321");

            var lblRamal = new Label { Text = "Ramal:", Location = new Point(x + 310, y), Size = new Size(60, 20) };
            txtRepresLegal_Ramal = new TextBox { Location = new Point(x + 375, y - 3), Size = new Size(80, 23) };
            toolTip.SetToolTip(lblRamal, "Ramal do telefone do representante legal (opcional).");
            toolTip.SetToolTip(txtRepresLegal_Ramal, "Ramal do telefone do representante legal (opcional).");

            grp.Controls.AddRange(new Control[] {
                lblCPF, txtRepresLegal_CPF, lblSetor, txtRepresLegal_Setor,
                lblDDD, txtRepresLegal_DDD, lblTelefone, txtRepresLegal_Telefone, lblRamal, txtRepresLegal_Ramal
            });
        }

        private void CriarAbaFechamento()
        {
            scrollFechamento = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            int yPos = 10;

            var grpBasicos = new GroupBox { Text = "Dados Básicos", Location = new Point(10, yPos), Size = new Size(scrollFechamento.Width - 30, 360) };
            grpBasicos.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Cálculo Automático de Período
            chkCalcularPeriodoFechamento = new CheckBox 
            { 
                Text = "Calcular período automaticamente", 
                Location = new Point(10, 25), 
                Size = new Size(250, 20),
                Checked = true
            };
            toolTip.SetToolTip(chkCalcularPeriodoFechamento, "Quando marcado, calcula automaticamente as datas de início e fim baseado no semestre e ano selecionados.\nOs campos de data ficam somente leitura.");
            chkCalcularPeriodoFechamento.CheckedChanged += (s, e) =>
            {
                bool calcularAutomatico = chkCalcularPeriodoFechamento.Checked;
                cmbSemestreFechamento.Enabled = calcularAutomatico;
                numAnoFechamento.Enabled = calcularAutomatico;
                txtDtInicioFechamento.ReadOnly = calcularAutomatico;
                txtDtFimFechamento.ReadOnly = calcularAutomatico;
                txtDtInicioFechamento.BackColor = calcularAutomatico ? SystemColors.Control : SystemColors.Window;
                txtDtFimFechamento.BackColor = calcularAutomatico ? SystemColors.Control : SystemColors.Window;
                
                if (calcularAutomatico)
                {
                    CalcularPeriodoFechamento();
                }
            };

            var lblSemestreFechamento = new Label { Text = "Semestre:", Location = new Point(10, 50), Size = new Size(100, 20) };
            cmbSemestreFechamento = new ComboBox 
            { 
                Location = new Point(115, 47), 
                Size = new Size(150, 23), 
                DropDownStyle = ComboBoxStyle.DropDownList 
            };
            cmbSemestreFechamento.Items.AddRange(new[] { "1º Semestre (Jan-Jun)", "2º Semestre (Jul-Dez)" });
            cmbSemestreFechamento.SelectedIndex = 0;
            cmbSemestreFechamento.SelectedIndexChanged += (s, e) => { if (chkCalcularPeriodoFechamento.Checked) CalcularPeriodoFechamento(); };
            toolTip.SetToolTip(lblSemestreFechamento, "Selecione o semestre para cálculo automático das datas.");
            toolTip.SetToolTip(cmbSemestreFechamento, "Selecione o semestre:\n• 1º Semestre = Janeiro a Junho (01/01 a 30/06)\n• 2º Semestre = Julho a Dezembro (01/07 a 31/12)");

            var lblAnoFechamento = new Label { Text = "Ano:", Location = new Point(275, 50), Size = new Size(50, 20) };
            numAnoFechamento = new NumericUpDown 
            { 
                Location = new Point(330, 47), 
                Size = new Size(80, 23),
                Minimum = 2000,
                Maximum = 2100,
                Value = DateTime.Now.Year
            };
            numAnoFechamento.ValueChanged += (s, e) => { if (chkCalcularPeriodoFechamento.Checked) CalcularPeriodoFechamento(); };
            toolTip.SetToolTip(lblAnoFechamento, "Ano para cálculo automático das datas.");
            toolTip.SetToolTip(numAnoFechamento, "Ano para cálculo automático das datas (2000-2100).");

            var lblDtInicio = new Label { Text = "Data Início (AAAA-MM-DD):", Location = new Point(10, 80), Size = new Size(150, 20) };
            txtDtInicioFechamento = new TextBox { Location = new Point(165, 77), Size = new Size(150, 23), ReadOnly = true, BackColor = SystemColors.Control };
            txtDtInicioFechamento.Leave += ValidarData;
            toolTip.SetToolTip(lblDtInicio, "Data de início do período de fechamento no formato AAAA-MM-DD.\nCalculada automaticamente quando a opção 'Calcular período automaticamente' está marcada.\nExemplo: 2023-01-01 (1º semestre) ou 2023-07-01 (2º semestre)");
            toolTip.SetToolTip(txtDtInicioFechamento, "Data de início do período de fechamento no formato AAAA-MM-DD.\nCalculada automaticamente quando a opção 'Calcular período automaticamente' está marcada.");

            var lblDtFim = new Label { Text = "Data Fim (AAAA-MM-DD):", Location = new Point(330, 80), Size = new Size(150, 20) };
            txtDtFimFechamento = new TextBox { Location = new Point(485, 77), Size = new Size(150, 23), ReadOnly = true, BackColor = SystemColors.Control };
            txtDtFimFechamento.Leave += ValidarData;
            toolTip.SetToolTip(lblDtFim, "Data de fim do período de fechamento no formato AAAA-MM-DD.\nCalculada automaticamente quando a opção 'Calcular período automaticamente' está marcada.\nExemplo: 2023-06-30 (1º semestre) ou 2023-12-31 (2º semestre)");
            toolTip.SetToolTip(txtDtFimFechamento, "Data de fim do período de fechamento no formato AAAA-MM-DD.\nCalculada automaticamente quando a opção 'Calcular período automaticamente' está marcada.");

            var lblTipoAmbiente = new Label { Text = "Tipo Ambiente:", Location = new Point(650, 80), Size = new Size(100, 20) };
            cmbTipoAmbienteFechamento = new ComboBox { Location = new Point(755, 77), Size = new Size(150, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTipoAmbienteFechamento.Items.AddRange(new[] { "1 - Produção", "2 - Homologação" });
            cmbTipoAmbienteFechamento.SelectedIndex = 1;
            toolTip.SetToolTip(lblTipoAmbiente, TOOLTIP_TIPO_AMBIENTE);
            toolTip.SetToolTip(cmbTipoAmbienteFechamento, TOOLTIP_TIPO_AMBIENTE);

            var lblAplicacaoEmissora = new Label { Text = "Aplicação Emissora:", Location = new Point(920, 80), Size = new Size(120, 20) };
            cmbAplicacaoEmissoraFechamento = new ComboBox { Location = new Point(1045, 77), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbAplicacaoEmissoraFechamento.Items.AddRange(new[] { "1 - Aplicação do contribuinte", "2 - Outros" });
            cmbAplicacaoEmissoraFechamento.SelectedIndex = 0;
            toolTip.SetToolTip(lblAplicacaoEmissora, TOOLTIP_APLICACAO_EMISSORA);
            toolTip.SetToolTip(cmbAplicacaoEmissoraFechamento, TOOLTIP_APLICACAO_EMISSORA);

            var lblIndRetificacao = new Label { Text = "Ind Retificação:", Location = new Point(650, 110), Size = new Size(100, 20) };
            cmbIndRetificacaoFechamento = new ComboBox { Location = new Point(755, 107), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbIndRetificacaoFechamento.Items.AddRange(new[] { "1 - Original", "2 - Retificação espontânea", "3 - Retificação a pedido" });
            cmbIndRetificacaoFechamento.SelectedIndex = 0;
            cmbIndRetificacaoFechamento.SelectedIndexChanged += CmbIndRetificacaoFechamento_SelectedIndexChanged;
            toolTip.SetToolTip(lblIndRetificacao, TOOLTIP_IND_RETIFICACAO);
            toolTip.SetToolTip(cmbIndRetificacaoFechamento, TOOLTIP_IND_RETIFICACAO);

            var lblNrRecibo = new Label { Text = "Nº Recibo:", Location = new Point(970, 110), Size = new Size(80, 20) };
            txtNrReciboFechamento = new TextBox { Location = new Point(1055, 107), Size = new Size(150, 23), Enabled = false };
            toolTip.SetToolTip(lblNrRecibo, TOOLTIP_NR_RECIBO);
            toolTip.SetToolTip(txtNrReciboFechamento, TOOLTIP_NR_RECIBO);

            var lblSitEspecial = new Label { Text = "Situação Especial:", Location = new Point(650, 140), Size = new Size(120, 20) };
            cmbSitEspecial = new ComboBox { Location = new Point(775, 137), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbSitEspecial.Items.AddRange(new[] { "0 - Não se aplica", "1 - Extinção", "2 - Fusão", "3 - Incorporação", "5 - Cisão" });
            cmbSitEspecial.SelectedIndex = 0;
            toolTip.SetToolTip(lblSitEspecial, "Situação especial da empresa:\n• 0 - Não se aplica: Situação normal\n• 1 - Extinção: Empresa extinta\n• 2 - Fusão: Empresa fundida\n• 3 - Incorporação: Empresa incorporada\n• 5 - Cisão: Empresa cindida");
            toolTip.SetToolTip(cmbSitEspecial, "Situação especial da empresa:\n• 0 - Não se aplica: Situação normal\n• 1 - Extinção: Empresa extinta\n• 2 - Fusão: Empresa fundida\n• 3 - Incorporação: Empresa incorporada\n• 5 - Cisão: Empresa cindida");

            chkNadaADeclarar = new CheckBox { Text = "Nada a Declarar", Location = new Point(990, 140), Size = new Size(150, 20) };
            toolTip.SetToolTip(chkNadaADeclarar, "Marque esta opção se não houver nada a declarar no período.\nSe marcado, não é necessário preencher os campos de fechamento.");

            // FechamentoPP - Toggle
            var lblFechamentoPP = new Label { Text = "FechamentoPP:", Location = new Point(10, 200), Size = new Size(120, 20) };
            chkFechamentoPP = new CheckBox 
            { 
                Text = "Com movimento de Previdência Privada", 
                Location = new Point(135, 197), 
                Size = new Size(280, 23),
                Appearance = Appearance.Button,
                FlatStyle = FlatStyle.Flat
            };
            chkFechamentoPP.FlatAppearance.BorderSize = 1;
            chkFechamentoPP.FlatAppearance.BorderColor = Color.Gray;
            chkFechamentoPP.CheckedChanged += (s, e) => 
            {
                chkFechamentoPP.Text = chkFechamentoPP.Checked 
                    ? "✓ Com movimento de Previdência Privada" 
                    : "✗ Sem movimento de Previdência Privada";
                chkFechamentoPP.BackColor = chkFechamentoPP.Checked ? Color.LightGreen : Color.LightGray;
            };
            chkFechamentoPP.Checked = false; // Inicializar (dispara o evento CheckedChanged)
            toolTip.SetToolTip(chkFechamentoPP, "Marque se há movimento de Previdência Privada no período.\nDesmarque se não houver movimento de Previdência Privada.");
            toolTip.SetToolTip(lblFechamentoPP, "Indica se há movimento de Previdência Privada no período.\nMarque = Com movimento | Desmarque = Sem movimento");

            // FechamentoMovOpFin - Toggle
            var lblFechamentoMovOpFin = new Label { Text = "FechamentoMovOpFin:", Location = new Point(430, 200), Size = new Size(180, 20) };
            chkFechamentoMovOpFin = new CheckBox 
            { 
                Text = "Com movimento de Operação Financeira", 
                Location = new Point(615, 197), 
                Size = new Size(280, 23),
                Appearance = Appearance.Button,
                FlatStyle = FlatStyle.Flat
            };
            chkFechamentoMovOpFin.FlatAppearance.BorderSize = 1;
            chkFechamentoMovOpFin.FlatAppearance.BorderColor = Color.Gray;
            chkFechamentoMovOpFin.CheckedChanged += (s, e) => 
            {
                chkFechamentoMovOpFin.Text = chkFechamentoMovOpFin.Checked 
                    ? "✓ Com movimento de Operação Financeira" 
                    : "✗ Sem movimento de Operação Financeira";
                chkFechamentoMovOpFin.BackColor = chkFechamentoMovOpFin.Checked ? Color.LightGreen : Color.LightGray;
            };
            chkFechamentoMovOpFin.Checked = false; // Inicializar (dispara o evento CheckedChanged)
            toolTip.SetToolTip(chkFechamentoMovOpFin, "Marque se há movimento de Operação Financeira no período.\nDesmarque se não houver movimento de Operação Financeira.");
            toolTip.SetToolTip(lblFechamentoMovOpFin, "Indica se há movimento de Operação Financeira no período.\nMarque = Com movimento | Desmarque = Sem movimento");

            // FechamentoMovOpFinAnual - Toggle
            var lblFechamentoMovOpFinAnual = new Label { Text = "FechamentoMovOpFinAnual:", Location = new Point(910, 200), Size = new Size(200, 20) };
            chkFechamentoMovOpFinAnual = new CheckBox 
            { 
                Text = "Com movimento de Operação Financeira Anual", 
                Location = new Point(1115, 197), 
                Size = new Size(320, 23),
                Appearance = Appearance.Button,
                FlatStyle = FlatStyle.Flat
            };
            chkFechamentoMovOpFinAnual.FlatAppearance.BorderSize = 1;
            chkFechamentoMovOpFinAnual.FlatAppearance.BorderColor = Color.Gray;
            chkFechamentoMovOpFinAnual.CheckedChanged += (s, e) => 
            {
                chkFechamentoMovOpFinAnual.Text = chkFechamentoMovOpFinAnual.Checked 
                    ? "✓ Com movimento de Operação Financeira Anual" 
                    : "✗ Sem movimento de Operação Financeira Anual";
                chkFechamentoMovOpFinAnual.BackColor = chkFechamentoMovOpFinAnual.Checked ? Color.LightGreen : Color.LightGray;
            };
            chkFechamentoMovOpFinAnual.Checked = false; // Inicializar (dispara o evento CheckedChanged)
            toolTip.SetToolTip(chkFechamentoMovOpFinAnual, "Marque se há movimento de Operação Financeira Anual no período.\nDesmarque se não houver movimento de Operação Financeira Anual.");
            toolTip.SetToolTip(lblFechamentoMovOpFinAnual, "Indica se há movimento de Operação Financeira Anual no período.\nMarque = Com movimento | Desmarque = Sem movimento");

            // Legenda explicativa
            var lblLegenda = new Label 
            { 
                Text = "ℹ️ IMPORTANTE: Se 'Nada a Declarar' NÃO estiver marcado, você DEVE marcar pelo menos um dos campos de fechamento acima.\n" +
                       "Use os toggles acima para indicar se há movimento de Previdência Privada, Operação Financeira ou Operação Financeira Anual.\n" +
                       "Exemplo: Se você enviou movimentações financeiras, marque 'FechamentoMovOpFin'.",
                Location = new Point(10, 240), 
                Size = new Size(1180, 50),
                ForeColor = Color.DarkBlue,
                Font = new Font("Microsoft Sans Serif", 8.5f, FontStyle.Regular)
            };

            grpBasicos.Controls.AddRange(new Control[] {
                chkCalcularPeriodoFechamento,
                lblSemestreFechamento, cmbSemestreFechamento, lblAnoFechamento, numAnoFechamento,
                lblDtInicio, txtDtInicioFechamento, lblDtFim, txtDtFimFechamento,
                lblTipoAmbiente, cmbTipoAmbienteFechamento, lblAplicacaoEmissora, cmbAplicacaoEmissoraFechamento,
                lblIndRetificacao, cmbIndRetificacaoFechamento, lblNrRecibo, txtNrReciboFechamento,
                lblSitEspecial, cmbSitEspecial, chkNadaADeclarar,
                lblFechamentoPP, chkFechamentoPP, lblFechamentoMovOpFin, chkFechamentoMovOpFin,
                lblFechamentoMovOpFinAnual, chkFechamentoMovOpFinAnual,
                lblLegenda
            });
            
            // Calcular período inicial
            CalcularPeriodoFechamento();

            scrollFechamento.Controls.Add(grpBasicos);
            tabFechamento.Controls.Add(scrollFechamento);
        }

        private void CriarAbaCadastroDeclarante()
        {
            scrollCadastroDeclarante = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            int yPos = 10;

            var grpBasicos = new GroupBox { Text = "Dados Básicos", Location = new Point(10, yPos), Size = new Size(scrollCadastroDeclarante.Width - 30, 400) };
            grpBasicos.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Campos básicos do ideEvento
            var lblTipoAmbiente = new Label { Text = "Tipo Ambiente:", Location = new Point(10, 25), Size = new Size(100, 20) };
            cmbTipoAmbienteCadastro = new ComboBox { Location = new Point(115, 22), Size = new Size(150, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTipoAmbienteCadastro.Items.AddRange(new[] { "1 - Produção", "2 - Homologação" });
            cmbTipoAmbienteCadastro.SelectedIndex = 1;
            toolTip.SetToolTip(lblTipoAmbiente, TOOLTIP_TIPO_AMBIENTE);
            toolTip.SetToolTip(cmbTipoAmbienteCadastro, TOOLTIP_TIPO_AMBIENTE);

            var lblAplicacaoEmissora = new Label { Text = "Aplicação Emissora:", Location = new Point(280, 25), Size = new Size(120, 20) };
            cmbAplicacaoEmissoraCadastro = new ComboBox { Location = new Point(405, 22), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbAplicacaoEmissoraCadastro.Items.AddRange(new[] { "1 - Aplicação do contribuinte", "2 - Outros" });
            cmbAplicacaoEmissoraCadastro.SelectedIndex = 0;
            toolTip.SetToolTip(lblAplicacaoEmissora, TOOLTIP_APLICACAO_EMISSORA);
            toolTip.SetToolTip(cmbAplicacaoEmissoraCadastro, TOOLTIP_APLICACAO_EMISSORA);

            var lblIndRetificacao = new Label { Text = "Ind Retificação:", Location = new Point(620, 25), Size = new Size(100, 20) };
            cmbIndRetificacaoCadastro = new ComboBox { Location = new Point(725, 22), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbIndRetificacaoCadastro.Items.AddRange(new[] { "1 - Original", "2 - Retificação espontânea", "3 - Retificação a pedido" });
            cmbIndRetificacaoCadastro.SelectedIndex = 0;
            cmbIndRetificacaoCadastro.SelectedIndexChanged += CmbIndRetificacaoCadastro_SelectedIndexChanged;
            toolTip.SetToolTip(lblIndRetificacao, TOOLTIP_IND_RETIFICACAO);
            toolTip.SetToolTip(cmbIndRetificacaoCadastro, TOOLTIP_IND_RETIFICACAO);

            var lblNrRecibo = new Label { Text = "Nº Recibo:", Location = new Point(940, 25), Size = new Size(80, 20) };
            txtNrReciboCadastro = new TextBox { Location = new Point(1025, 22), Size = new Size(150, 23), Enabled = false };
            toolTip.SetToolTip(lblNrRecibo, TOOLTIP_NR_RECIBO);
            toolTip.SetToolTip(txtNrReciboCadastro, TOOLTIP_NR_RECIBO);

            // Campos do infoCadastro
            yPos = 60;
            var lblGIIN = new Label { Text = "GIIN:", Location = new Point(10, yPos), Size = new Size(80, 20) };
            txtGIIN = new TextBox { Location = new Point(95, yPos - 3), Size = new Size(200, 23) };
            toolTip.SetToolTip(lblGIIN, "GIIN (Global Intermediary Identification Number) da Entidade Declarante.\nFormato: 6 caracteres alfanuméricos + '.' + 5 caracteres + '.' + 2 caracteres + '.' + 3 caracteres numéricos");
            toolTip.SetToolTip(txtGIIN, "GIIN da Entidade Declarante (opcional).");

            var lblCategoriaDeclarante = new Label { Text = "Categoria Declarante:", Location = new Point(310, yPos), Size = new Size(130, 20) };
            txtCategoriaDeclarante = new TextBox { Location = new Point(445, yPos - 3), Size = new Size(150, 23) };
            toolTip.SetToolTip(lblCategoriaDeclarante, "Código da categoria de declarante conforme tabela vigente.\nExemplo: FATCA602 para Instituições Financeiras Brasileiras Informantes.");
            toolTip.SetToolTip(txtCategoriaDeclarante, "Código da categoria de declarante (opcional).");

            yPos += 35;
            var lblNome = new Label { Text = "Nome (Razão Social):", Location = new Point(10, yPos), Size = new Size(130, 20) };
            txtNomeCadastro = new TextBox { Location = new Point(145, yPos - 3), Size = new Size(400, 23) };
            toolTip.SetToolTip(lblNome, "Razão social da Entidade Declarante, nome empresarial ou denominação.\nDeve ser idêntico ao que consta no Cadastro CNPJ.");
            toolTip.SetToolTip(txtNomeCadastro, "Razão social da Entidade Declarante (obrigatório).");

            var lblTpNome = new Label { Text = "Tipo Nome:", Location = new Point(560, yPos), Size = new Size(80, 20) };
            txtTpNome = new TextBox { Location = new Point(645, yPos - 3), Size = new Size(150, 23) };
            toolTip.SetToolTip(lblTpNome, "Classificação do nome apresentado (opcional).\nConforme tabela de Tipos de Nome vigente.");
            toolTip.SetToolTip(txtTpNome, "Tipo do nome (opcional).");

            yPos += 35;
            var lblEnderecoLivre = new Label { Text = "Endereço Livre:", Location = new Point(10, yPos), Size = new Size(100, 20) };
            txtEnderecoLivreCadastro = new TextBox { Location = new Point(115, yPos - 3), Size = new Size(500, 23) };
            toolTip.SetToolTip(lblEnderecoLivre, "Endereço principal da Entidade Declarante em formato livre.\nFormato: endereço/cep/cidade/UF");
            toolTip.SetToolTip(txtEnderecoLivreCadastro, "Endereço principal em formato livre (obrigatório).");

            var lblTpEndereco = new Label { Text = "Tipo Endereço:", Location = new Point(630, yPos), Size = new Size(100, 20) };
            txtTpEnderecoCadastro = new TextBox { Location = new Point(735, yPos - 3), Size = new Size(150, 23) };
            toolTip.SetToolTip(lblTpEndereco, "Tipo de endereço principal conforme tabela vigente (opcional).");
            toolTip.SetToolTip(txtTpEnderecoCadastro, "Tipo de endereço (opcional).");

            yPos += 35;
            var lblMunicipio = new Label { Text = "Município (Código IBGE):", Location = new Point(10, yPos), Size = new Size(140, 20) };
            txtMunicipioCadastro = new TextBox { Location = new Point(155, yPos - 3), Size = new Size(150, 23) };
            toolTip.SetToolTip(lblMunicipio, "Código do município conforme tabela do IBGE (7 dígitos).");
            toolTip.SetToolTip(txtMunicipioCadastro, "Código do município (obrigatório).");

            var lblUF = new Label { Text = "UF:", Location = new Point(320, yPos), Size = new Size(50, 20) };
            txtUFCadastro = new TextBox { Location = new Point(375, yPos - 3), Size = new Size(50, 23) };
            toolTip.SetToolTip(lblUF, "Sigla da Unidade da Federação (UF).");
            toolTip.SetToolTip(txtUFCadastro, "UF (obrigatório).");

            var lblCEP = new Label { Text = "CEP:", Location = new Point(440, yPos), Size = new Size(50, 20) };
            txtCEPCadastro = new TextBox { Location = new Point(495, yPos - 3), Size = new Size(100, 23) };
            toolTip.SetToolTip(lblCEP, "CEP do endereço (8 dígitos, sem hífen).");
            toolTip.SetToolTip(txtCEPCadastro, "CEP (obrigatório).");

            var lblPais = new Label { Text = "País:", Location = new Point(610, yPos), Size = new Size(50, 20) };
            txtPaisCadastro = new TextBox { Location = new Point(665, yPos - 3), Size = new Size(50, 23) };
            txtPaisCadastro.Text = "BR";
            toolTip.SetToolTip(lblPais, "Código do país conforme tabela ISO-3166-1 alfa 2.");
            toolTip.SetToolTip(txtPaisCadastro, "Código do país (obrigatório, padrão: BR).");

            yPos += 35;
            var lblPaisResid = new Label { Text = "País Residência Fiscal:", Location = new Point(10, yPos), Size = new Size(140, 20) };
            txtPaisResid = new TextBox { Location = new Point(155, yPos - 3), Size = new Size(200, 23) };
            txtPaisResid.Text = "BR";
            toolTip.SetToolTip(lblPaisResid, "País(es) de residência fiscal da entidade declarante.\nDeve conter pelo menos 'BR'. Pode conter múltiplos países separados por vírgula.");
            toolTip.SetToolTip(txtPaisResid, "País(es) de residência fiscal (obrigatório, deve conter BR).\nExemplo: BR ou BR,US");

            grpBasicos.Controls.AddRange(new Control[] {
                lblTipoAmbiente, cmbTipoAmbienteCadastro,
                lblAplicacaoEmissora, cmbAplicacaoEmissoraCadastro,
                lblIndRetificacao, cmbIndRetificacaoCadastro, lblNrRecibo, txtNrReciboCadastro,
                lblGIIN, txtGIIN, lblCategoriaDeclarante, txtCategoriaDeclarante,
                lblNome, txtNomeCadastro, lblTpNome, txtTpNome,
                lblEnderecoLivre, txtEnderecoLivreCadastro, lblTpEndereco, txtTpEnderecoCadastro,
                lblMunicipio, txtMunicipioCadastro, lblUF, txtUFCadastro, lblCEP, txtCEPCadastro, lblPais, txtPaisCadastro,
                lblPaisResid, txtPaisResid
            });

            scrollCadastroDeclarante.Controls.Add(grpBasicos);
            tabCadastroDeclarante.Controls.Add(scrollCadastroDeclarante);
        }

        private void CmbIndRetificacaoCadastro_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool habilitar = cmbIndRetificacaoCadastro.SelectedIndex == 1 || cmbIndRetificacaoCadastro.SelectedIndex == 2;
            txtNrReciboCadastro.Enabled = habilitar;
            txtNrReciboCadastro.BackColor = habilitar ? SystemColors.Window : SystemColors.Control;
        }

        // Validações
        private void ValidarCNPJ(object sender, EventArgs e)
        {
            var txt = sender as TextBox;
            if (txt == null) return;

            string cnpj = txt.Text.Replace(".", "").Replace("/", "").Replace("-", "").Trim();
            if (string.IsNullOrEmpty(cnpj)) return;

            if (cnpj.Length != 14 || !cnpj.All(char.IsDigit))
            {
                MessageBox.Show("CNPJ inválido. Deve conter 14 dígitos.", TITULO_VALIDACAO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txt.Focus();
            }
        }

        private void ValidarCPF(object sender, EventArgs e)
        {
            var txt = sender as TextBox;
            if (txt == null) return;

            string cpf = txt.Text.Replace(".", "").Replace("-", "").Trim();
            if (string.IsNullOrEmpty(cpf)) return;

            if (cpf.Length != 11 || !cpf.All(char.IsDigit))
            {
                MessageBox.Show("CPF inválido. Deve conter 11 dígitos.", TITULO_VALIDACAO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txt.Focus();
            }
        }

        private void ValidarCEP(object sender, EventArgs e)
        {
            var txt = sender as TextBox;
            if (txt == null) return;

            string cep = txt.Text.Replace("-", "").Trim();
            if (string.IsNullOrEmpty(cep)) return;

            if (cep.Length != 8 || !cep.All(char.IsDigit))
            {
                MessageBox.Show("CEP inválido. Deve conter 8 dígitos.", TITULO_VALIDACAO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txt.Focus();
            }
        }

        private void ValidarData(object sender, EventArgs e)
        {
            var txt = sender as TextBox;
            if (txt == null) return;

            if (string.IsNullOrEmpty(txt.Text)) return;

            if (!DateTime.TryParseExact(txt.Text, "yyyy-MM-dd", CULTURE_INFO_PT_BR, System.Globalization.DateTimeStyles.None, out _))
            {
                MessageBox.Show("Data inválida. Use o formato AAAA-MM-DD.", TITULO_VALIDACAO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txt.Focus();
            }
        }

        private void ValidarEmail(object sender, EventArgs e)
        {
            var txt = sender as TextBox;
            if (txt == null) return;

            if (string.IsNullOrEmpty(txt.Text)) return;

            var regex = new Regex(@"^[^@\s]+@[^@\s]+\.[^@\s]+$");
            if (!regex.IsMatch(txt.Text))
            {
                MessageBox.Show("Email inválido.", TITULO_VALIDACAO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txt.Focus();
            }
        }

        /// <summary>
        /// Calcula o período automaticamente baseado no semestre e ano selecionados
        /// </summary>
        private void CalcularPeriodoAutomatico()
        {
            if (!chkCalcularPeriodoAutomatico.Checked)
                return;

            int ano = (int)numAno.Value;
            int semestre = cmbSemestre.SelectedIndex + 1; // 1 ou 2
            
            // Calcular período: YYYY01 para 1º semestre, YYYY02 para 2º semestre
            string periodo = $"{ano}{semestre:00}";
            txtPeriodo.Text = periodo;
        }

        /// <summary>
        /// Preenche os campos de semestre e ano baseado no período informado
        /// </summary>
        private void PreencherSemestreAnoDoPeriodo(string periodo)
        {
            if (string.IsNullOrWhiteSpace(periodo) || periodo.Length != 6)
                return;

            if (!int.TryParse(periodo, out int periodoInt))
                return;

            int ano = periodoInt / 100;
            int mes = periodoInt % 100;

            // Determinar semestre baseado no mês
            // 01 ou 06 = 1º semestre, 02 ou 12 = 2º semestre
            int semestre;
            if (mes == 1 || mes == 6)
            {
                semestre = 1;
            }
            else if (mes == 2 || mes == 12)
            {
                semestre = 2;
            }
            else
            {
                return; // Período inválido
            }

            // Preencher campos (sem disparar eventos para evitar loop)
            bool calcularAutomatico = chkCalcularPeriodoAutomatico.Checked;
            
            numAno.Value = ano;
            cmbSemestre.SelectedIndex = semestre - 1;
            
            if (calcularAutomatico)
            {
                txtPeriodo.Text = periodo;
            }
        }

        /// <summary>
        /// Calcula as datas de abertura automaticamente baseado no semestre e ano selecionados
        /// </summary>
        private void CalcularPeriodoAbertura()
        {
            if (!chkCalcularPeriodoAbertura.Checked)
                return;

            int ano = (int)numAnoAbertura.Value;
            int semestre = cmbSemestreAbertura.SelectedIndex + 1; // 1 ou 2
            
            // Calcular datas: 1º semestre = 01/01 a 30/06, 2º semestre = 01/07 a 31/12
            if (semestre == 1)
            {
                txtDtInicioAbertura.Text = $"{ano}-01-01";
                txtDtFimAbertura.Text = $"{ano}-06-30";
            }
            else
            {
                txtDtInicioAbertura.Text = $"{ano}-07-01";
                txtDtFimAbertura.Text = $"{ano}-12-31";
            }
        }

        /// <summary>
        /// Calcula as datas de fechamento automaticamente baseado no semestre e ano selecionados
        /// </summary>
        private void CalcularPeriodoFechamento()
        {
            if (!chkCalcularPeriodoFechamento.Checked)
                return;

            int ano = (int)numAnoFechamento.Value;
            int semestre = cmbSemestreFechamento.SelectedIndex + 1; // 1 ou 2
            
            // Calcular datas: 1º semestre = 01/01 a 30/06, 2º semestre = 01/07 a 31/12
            if (semestre == 1)
            {
                txtDtInicioFechamento.Text = $"{ano}-01-01";
                txtDtFimFechamento.Text = $"{ano}-06-30";
            }
            else
            {
                txtDtInicioFechamento.Text = $"{ano}-07-01";
                txtDtFimFechamento.Text = $"{ano}-12-31";
            }
        }

        /// <summary>
        /// Preenche os campos de semestre e ano na aba de Abertura baseado nas datas
        /// </summary>
        private void PreencherSemestreAnoAbertura(string dtInicio, string dtFim)
        {
            if (string.IsNullOrWhiteSpace(dtInicio) || string.IsNullOrWhiteSpace(dtFim))
                return;

            if (!DateTime.TryParse(dtInicio, CULTURE_INFO_PT_BR, System.Globalization.DateTimeStyles.None, out DateTime inicio) || 
                !DateTime.TryParse(dtFim, CULTURE_INFO_PT_BR, System.Globalization.DateTimeStyles.None, out DateTime fim))
                return;

            int ano = inicio.Year;
            int mesInicio = inicio.Month;
            int mesFim = fim.Month;

            // Determinar semestre: 1º = Jan-Jun (mes 1-6), 2º = Jul-Dez (mes 7-12)
            int semestre;
            if (mesInicio >= 1 && mesInicio <= 6 && mesFim >= 1 && mesFim <= 6)
            {
                semestre = 1;
            }
            else if (mesInicio >= 7 && mesInicio <= 12 && mesFim >= 7 && mesFim <= 12)
            {
                semestre = 2;
            }
            else
            {
                return; // Datas inválidas para um semestre
            }

            // Preencher campos (sem disparar eventos para evitar loop)
            numAnoAbertura.Value = ano;
            cmbSemestreAbertura.SelectedIndex = semestre - 1;
        }

        /// <summary>
        /// Preenche os campos de semestre e ano na aba de Fechamento baseado nas datas
        /// </summary>
        private void PreencherSemestreAnoFechamento(string dtInicio, string dtFim)
        {
            if (string.IsNullOrWhiteSpace(dtInicio) || string.IsNullOrWhiteSpace(dtFim))
                return;

            if (!DateTime.TryParse(dtInicio, CULTURE_INFO_PT_BR, System.Globalization.DateTimeStyles.None, out DateTime inicio) || 
                !DateTime.TryParse(dtFim, CULTURE_INFO_PT_BR, System.Globalization.DateTimeStyles.None, out DateTime fim))
                return;

            int ano = inicio.Year;
            int mesInicio = inicio.Month;
            int mesFim = fim.Month;

            // Determinar semestre: 1º = Jan-Jun (mes 1-6), 2º = Jul-Dez (mes 7-12)
            int semestre;
            if (mesInicio >= 1 && mesInicio <= 6 && mesFim >= 1 && mesFim <= 6)
            {
                semestre = 1;
            }
            else if (mesInicio >= 7 && mesInicio <= 12 && mesFim >= 7 && mesFim <= 12)
            {
                semestre = 2;
            }
            else
            {
                return; // Datas inválidas para um semestre
            }

            // Preencher campos (sem disparar eventos para evitar loop)
            numAnoFechamento.Value = ano;
            cmbSemestreFechamento.SelectedIndex = semestre - 1;
        }

        private void CmbIndRetificacaoAbertura_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool habilitar = cmbIndRetificacaoAbertura.SelectedIndex > 0;
            txtNrReciboAbertura.Enabled = habilitar;
            if (!habilitar) txtNrReciboAbertura.Text = "";
        }

        private void CmbIndRetificacaoFechamento_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool habilitar = cmbIndRetificacaoFechamento.SelectedIndex > 0;
            txtNrReciboFechamento.Enabled = habilitar;
            if (!habilitar) txtNrReciboFechamento.Text = "";
        }

        private void BtnSelecionarCert_Click(object sender, EventArgs e)
        {
            X509Certificate2 cert = SelecionarCertificado();
            if (cert != null)
            {
                txtCertThumbprint.Text = cert.Thumbprint;
            }
        }

        private void BtnSelecionarCertServidor_Click(object sender, EventArgs e)
        {
            X509Certificate2 cert = SelecionarCertificado();
            if (cert != null)
            {
                txtCertServidorThumbprint.Text = cert.Thumbprint;
            }
        }

        private static X509Certificate2 SelecionarCertificado()
        {
            X509Store store = new X509Store("MY", StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
            X509Certificate2Collection certs = store.Certificates;
            X509Certificate2Collection certsParaAssinatura = certs.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, false);
            X509Certificate2Collection certsParaSelecionar = X509Certificate2UI.SelectFromCollection(certsParaAssinatura,
                "Certificado(s) Digital(is) disponível(is)", "Selecione o certificado digital para uso no aplicativo", X509SelectionFlag.SingleSelection);

            store.Close();

            if (certsParaSelecionar.Count == 0)
            {
                return null;
            }

            return certsParaSelecionar[0];
        }

        private void BtnSelecionarDiretorio_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtDiretorioLotes.Text = dialog.SelectedPath;
            }
        }

        private void BtnCarregarConfig_Click(object sender, EventArgs e)
        {
            try
            {
                var configCompleta = persistenciaService.CarregarConfiguracao();
                if (configCompleta == null)
                {
                    MessageBox.Show("Nenhuma configuração salva encontrada.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                CarregarConfiguracaoNaTela(configCompleta);
                MessageBox.Show("Configuração carregada com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao carregar configuração: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnTestarConexao_Click(object sender, EventArgs e)
        {
            try
            {
                var dbService = new EfinanceiraDatabaseService();
                
                btnTestarConexao.Enabled = false;
                btnTestarConexao.Text = "Testando...";
                Application.DoEvents();
                
                bool conectado = dbService.TestarConexao();
                
                if (conectado)
                {
                    // Teste adicional: tentar fazer uma consulta simples
                    var pessoas = dbService.BuscarPessoasComContas(2024, 1, 6, 1, 0);
                    MessageBox.Show($"✓ Conexão estabelecida com sucesso!\n\nTeste de consulta: {pessoas.Count} registro(s) encontrado(s).", 
                        "Conexão OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("✗ Falha na conexão com o banco de dados.\n\nVerifique as credenciais e conectividade de rede.", 
                        "Erro de Conexão", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao testar conexão:\n\n{ex.Message}", 
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnTestarConexao.Enabled = true;
                btnTestarConexao.Text = "Testar Conexão BD";
            }
        }

        private void BtnLimparDadosTeste_Click(object sender, EventArgs e)
        {
            try
            {
                // Validar CNPJ
                if (string.IsNullOrWhiteSpace(txtCnpjDeclarante.Text))
                {
                    MessageBox.Show("CNPJ do declarante não informado.\n\nPor favor, preencha o CNPJ antes de executar a limpeza.", 
                        "CNPJ Obrigatório", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Validar certificado
                if (string.IsNullOrWhiteSpace(txtCertThumbprint.Text))
                {
                    MessageBox.Show("Certificado não configurado.\n\nPor favor, selecione um certificado antes de executar a limpeza.", 
                        "Certificado Obrigatório", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Confirmar ação
                DialogResult resultado = MessageBox.Show(
                    $"Deseja realmente limpar os dados de teste para o CNPJ {txtCnpjDeclarante.Text}?\n\n" +
                    "⚠️ ATENÇÃO: Esta ação é irreversível e excluirá todos os dados de teste no ambiente de Produção Restrita.\n\n" +
                    "Certifique-se de que:\n" +
                    "• O certificado possui permissão para o CNPJ informado\n" +
                    "• Você realmente deseja excluir todos os dados de teste",
                    "Confirmar Limpeza de Dados",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                if (resultado != DialogResult.Yes)
                {
                    return;
                }

                // Desabilitar botão durante execução
                btnLimparDadosTeste.Enabled = false;
                btnLimparDadosTeste.Text = "Limpando...";
                Application.DoEvents();

                // Buscar certificado
                X509Certificate2 certificado = BuscarCertificado(txtCertThumbprint.Text);

                // Executar limpeza
                var limpezaService = new EfinanceiraLimpezaService();
                var resposta = limpezaService.LimparDadosTeste(txtCnpjDeclarante.Text, certificado);

                // Exibir resultado
                if (resposta.Sucesso)
                {
                    MessageBox.Show(
                        $"✓ Limpeza executada com sucesso!\n\n" +
                        $"Código HTTP: {resposta.CodigoHttp}\n" +
                        $"Mensagem: {resposta.Mensagem}",
                        "Limpeza Concluída",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        $"✗ Erro ao executar limpeza.\n\n" +
                        $"Código HTTP: {resposta.CodigoHttp}\n" +
                        $"Mensagem: {resposta.Mensagem}",
                        "Erro na Limpeza",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Erro ao executar limpeza de dados:\n\n{ex.Message}",
                    "Erro",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                btnLimparDadosTeste.Enabled = true;
                btnLimparDadosTeste.Text = "Limpar Dados de Teste";
            }
        }

        private static X509Certificate2 BuscarCertificado(string thumbprint)
        {
            if (string.IsNullOrEmpty(thumbprint))
            {
                throw new ArgumentException("Thumbprint do certificado não configurado.", nameof(thumbprint));
            }

            string thumbprintNormalizado = thumbprint.Replace(" ", "").Replace("-", "").ToUpper();

            // Buscar no repositório CurrentUser\My
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

            // Buscar no repositório LocalMachine\My
            store = new X509Store(StoreName.My, StoreLocation.LocalMachine);
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

            throw new InvalidOperationException($"Certificado com thumbprint '{thumbprint}' não encontrado no repositório do Windows.");
        }

        private void BtnSalvarConfig_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidarCamposObrigatorios())
                {
                    return;
                }

                // Salvar configurações
                Config = new EfinanceiraConfig
                {
                    CnpjDeclarante = txtCnpjDeclarante.Text,
                    CertThumbprint = txtCertThumbprint.Text,
                    CertServidorThumbprint = txtCertServidorThumbprint.Text,
                    Ambiente = cmbAmbiente.SelectedItem.ToString() == "PROD" ? EfinanceiraAmbiente.PROD : EfinanceiraAmbiente.TEST,
                    ModoTeste = chkModoTeste.Checked,
                    TestEnvioHabilitado = chkHabilitarEnvio.Checked,
                    DiretorioLotes = txtDiretorioLotes.Text,
                    Periodo = txtPeriodo.Text,
                    PageSize = (int)numPageSize.Value,
                    EventoOffset = (int)numEventoOffset.Value,
                    OffsetRegistros = (int)numOffsetRegistros.Value,
                    MaxLotes = chkMaxLotesIlimitado.Checked ? null : (int?)numMaxLotes.Value,
                    EventosPorLote = numEventosPorLote != null ? (int)numEventosPorLote.Value : 50,
                    UrlTeste = URL_TESTE,
                    UrlProducao = URL_PRODUCAO,
                    UrlConsultaTeste = URL_CONSULTA_TESTE,
                    UrlConsultaProducao = URL_CONSULTA_PRODUCAO
                };

                // Salvar dados de abertura
                DadosAbertura = new DadosAbertura
                {
                    CnpjDeclarante = txtCnpjDeclarante.Text,
                    DtInicio = txtDtInicioAbertura.Text,
                    DtFim = txtDtFimAbertura.Text,
                    TipoAmbiente = cmbTipoAmbienteAbertura.SelectedIndex + 1,
                    AplicacaoEmissora = cmbAplicacaoEmissoraAbertura.SelectedIndex + 1,
                    IndRetificacao = cmbIndRetificacaoAbertura.SelectedIndex + 1,
                    NrRecibo = txtNrReciboAbertura.Text,
                    IndicarMovOpFin = chkIndicarMovOpFin.Checked,
                    ResponsavelRMF = CriarDadosResponsavelRMF(),
                    RespeFin = CriarDadosRespeFin(),
                    RepresLegal = CriarDadosRepresLegal()
                };

                // Salvar dados de fechamento
                DadosFechamento = new DadosFechamento
                {
                    CnpjDeclarante = txtCnpjDeclarante.Text,
                    DtInicio = txtDtInicioFechamento.Text,
                    DtFim = txtDtFimFechamento.Text,
                    TipoAmbiente = cmbTipoAmbienteFechamento.SelectedIndex + 1,
                    AplicacaoEmissora = cmbAplicacaoEmissoraFechamento.SelectedIndex + 1,
                    IndRetificacao = cmbIndRetificacaoFechamento.SelectedIndex + 1,
                    NrRecibo = txtNrReciboFechamento.Text,
                    SitEspecial = int.Parse(cmbSitEspecial.SelectedItem.ToString().Split('-')[0].Trim()),
                    NadaADeclarar = chkNadaADeclarar.Checked ? "1" : null,
                    FechamentoPP = chkFechamentoPP.Checked ? (int?)1 : null,
                    FechamentoMovOpFin = chkFechamentoMovOpFin.Checked ? (int?)1 : null,
                    FechamentoMovOpFinAnual = chkFechamentoMovOpFinAnual.Checked ? (int?)1 : null
                };

                // Salvar dados de cadastro de declarante
                DadosCadastroDeclarante = CriarDadosCadastroDeclarante();

                // Persistir
                persistenciaService.SalvarConfiguracao(Config, DadosAbertura, DadosFechamento, DadosCadastroDeclarante);

                MessageBox.Show("Configurações salvas com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar configurações: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool ValidarCamposObrigatorios()
        {
            if (string.IsNullOrWhiteSpace(txtCnpjDeclarante.Text))
            {
                MessageBox.Show("CNPJ Declarante é obrigatório.", TITULO_VALIDACAO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtCnpjDeclarante.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtDtInicioAbertura.Text))
            {
                MessageBox.Show("Data de início da abertura é obrigatória.", TITULO_VALIDACAO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtDtInicioAbertura.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtDtFimAbertura.Text))
            {
                MessageBox.Show("Data de fim da abertura é obrigatória.", TITULO_VALIDACAO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtDtFimAbertura.Focus();
                return false;
            }

            if (chkIndicarMovOpFin.Checked && (string.IsNullOrWhiteSpace(txtRMF_CPF.Text) || string.IsNullOrWhiteSpace(txtRespeFin_CPF.Text)))
            {
                MessageBox.Show("Ao indicar MovOpFin, é necessário preencher CPF do Responsável RMF e RespeFin.", TITULO_VALIDACAO, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
            }

            // Validação de fechamento: se não for "nada a declarar", deve ter pelo menos um fechamento
            if (!chkNadaADeclarar.Checked)
            {
                bool temFechamentoPP = chkFechamentoPP.Checked;
                bool temFechamentoMovOpFin = chkFechamentoMovOpFin.Checked;
                bool temFechamentoMovOpFinAnual = chkFechamentoMovOpFinAnual.Checked;

                if (!temFechamentoPP && !temFechamentoMovOpFin && !temFechamentoMovOpFinAnual)
                {
                    MessageBox.Show(
                        "Se 'Nada a Declarar' não estiver marcado, você DEVE preencher pelo menos um dos campos de fechamento:\n\n" +
                        "• FechamentoPP (marque se houver movimento de Previdência Privada)\n" +
                        "• FechamentoMovOpFin (marque se houver movimento de Operação Financeira)\n" +
                        "• FechamentoMovOpFinAnual (marque se houver movimento de Operação Financeira Anual)\n\n" +
                        "Exemplo: Se você enviou movimentações financeiras, marque o toggle 'FechamentoMovOpFin'.",
                        "Validação de Fechamento",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    chkFechamentoMovOpFin.Focus();
                    return false;
                }
            }

            return true;
        }

        private DadosResponsavelRMF CriarDadosResponsavelRMF()
        {
            if (string.IsNullOrWhiteSpace(txtRMF_CPF.Text)) return null;

            return new DadosResponsavelRMF
            {
                Cnpj = txtRMF_CNPJ.Text,
                Cpf = txtRMF_CPF.Text,
                Nome = txtRMF_Nome.Text,
                Setor = txtRMF_Setor.Text,
                TelefoneDDD = txtRMF_DDD.Text,
                TelefoneNumero = txtRMF_Telefone.Text,
                TelefoneRamal = txtRMF_Ramal.Text,
                EnderecoLogradouro = txtRMF_Logradouro.Text,
                EnderecoNumero = txtRMF_Numero.Text,
                EnderecoComplemento = txtRMF_Complemento.Text,
                EnderecoBairro = txtRMF_Bairro.Text,
                EnderecoCEP = txtRMF_CEP.Text,
                EnderecoMunicipio = txtRMF_Municipio.Text,
                EnderecoUF = txtRMF_UF.Text
            };
        }

        private DadosRespeFin CriarDadosRespeFin()
        {
            if (string.IsNullOrWhiteSpace(txtRespeFin_CPF.Text)) return null;

            return new DadosRespeFin
            {
                Cpf = txtRespeFin_CPF.Text,
                Nome = txtRespeFin_Nome.Text,
                Setor = txtRespeFin_Setor.Text,
                TelefoneDDD = txtRespeFin_DDD.Text,
                TelefoneNumero = txtRespeFin_Telefone.Text,
                TelefoneRamal = txtRespeFin_Ramal.Text,
                EnderecoLogradouro = txtRespeFin_Logradouro.Text,
                EnderecoNumero = txtRespeFin_Numero.Text,
                EnderecoComplemento = txtRespeFin_Complemento.Text,
                EnderecoBairro = txtRespeFin_Bairro.Text,
                EnderecoCEP = txtRespeFin_CEP.Text,
                EnderecoMunicipio = txtRespeFin_Municipio.Text,
                EnderecoUF = txtRespeFin_UF.Text,
                Email = txtRespeFin_Email.Text
            };
        }

        private DadosRepresLegal CriarDadosRepresLegal()
        {
            if (string.IsNullOrWhiteSpace(txtRepresLegal_CPF.Text)) return null;

            return new DadosRepresLegal
            {
                Cpf = txtRepresLegal_CPF.Text,
                Setor = txtRepresLegal_Setor.Text,
                TelefoneDDD = txtRepresLegal_DDD.Text,
                TelefoneNumero = txtRepresLegal_Telefone.Text,
                TelefoneRamal = txtRepresLegal_Ramal.Text
            };
        }

        private DadosCadastroDeclarante CriarDadosCadastroDeclarante()
        {
            var dados = new DadosCadastroDeclarante
            {
                CnpjDeclarante = txtCnpjDeclarante.Text,
                TipoAmbiente = cmbTipoAmbienteCadastro.SelectedIndex + 1,
                AplicacaoEmissora = cmbAplicacaoEmissoraCadastro.SelectedIndex + 1,
                IndRetificacao = cmbIndRetificacaoCadastro.SelectedIndex + 1,
                NrRecibo = txtNrReciboCadastro.Text,
                GIIN = string.IsNullOrWhiteSpace(txtGIIN.Text) ? null : txtGIIN.Text,
                CategoriaDeclarante = string.IsNullOrWhiteSpace(txtCategoriaDeclarante.Text) ? null : txtCategoriaDeclarante.Text,
                Nome = txtNomeCadastro.Text,
                TpNome = string.IsNullOrWhiteSpace(txtTpNome.Text) ? null : txtTpNome.Text,
                EnderecoLivre = txtEnderecoLivreCadastro.Text,
                TpEndereco = string.IsNullOrWhiteSpace(txtTpEnderecoCadastro.Text) ? null : txtTpEnderecoCadastro.Text,
                Municipio = txtMunicipioCadastro.Text,
                UF = txtUFCadastro.Text,
                CEP = txtCEPCadastro.Text,
                Pais = string.IsNullOrWhiteSpace(txtPaisCadastro.Text) ? "BR" : txtPaisCadastro.Text
            };

            // Processar países de residência fiscal
            if (!string.IsNullOrWhiteSpace(txtPaisResid.Text))
            {
                dados.PaisResid = txtPaisResid.Text.Split(',').Select(p => p.Trim()).ToList();
            }
            else
            {
                dados.PaisResid = new List<string> { "BR" };
            }

            // Por enquanto, não implementamos NIFs, TiposInstPgto e EnderecosOutros
            // Podem ser adicionados posteriormente se necessário
            dados.NIFs = new List<DadosNIF>();
            dados.TiposInstPgto = new List<DadosTipoInstPgto>();
            dados.EnderecosOutros = new List<DadosEnderecoOutros>();

            return dados;
        }

        private void CarregarConfiguracaoNaTela(ConfiguracaoCompleta configCompleta)
        {
            if (configCompleta.Config != null)
            {
                CarregarConfiguracaoGeral(configCompleta.Config);
            }

            if (configCompleta.DadosAbertura != null)
            {
                CarregarDadosAbertura(configCompleta.DadosAbertura);
            }

            if (configCompleta.DadosFechamento != null)
            {
                CarregarDadosFechamento(configCompleta.DadosFechamento);
            }

            if (configCompleta.DadosCadastroDeclarante != null)
            {
                CarregarDadosCadastroDeclarante(configCompleta.DadosCadastroDeclarante);
            }
        }

        private void CarregarConfiguracaoGeral(EfinanceiraConfig config)
        {
            txtCnpjDeclarante.Text = config.CnpjDeclarante ?? "";
            txtCertThumbprint.Text = config.CertThumbprint ?? "";
            txtCertServidorThumbprint.Text = config.CertServidorThumbprint ?? "";
            cmbAmbiente.SelectedItem = config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : "TEST";
            chkModoTeste.Checked = config.ModoTeste;
            chkHabilitarEnvio.Checked = config.TestEnvioHabilitado;
            txtDiretorioLotes.Text = config.DiretorioLotes ?? "";
            
            // Desabilitar botão de limpar dados se não for ambiente de testes
            if (btnLimparDadosTeste != null)
            {
                bool isProducao = config.Ambiente == EfinanceiraAmbiente.PROD;
                btnLimparDadosTeste.Enabled = !isProducao;
            }
            
            // Carregar período e preencher semestre/ano se houver
            string periodoSalvo = config.Periodo ?? "";
            if (!string.IsNullOrEmpty(periodoSalvo))
            {
                PreencherSemestreAnoDoPeriodo(periodoSalvo);
            }
            else
            {
                CalcularPeriodoAutomatico();
            }
                
                // Carregar configurações de processamento
            numPageSize.Value = config.PageSize > 0 ? config.PageSize : 500;
            numEventoOffset.Value = config.EventoOffset >= 0 ? config.EventoOffset : 1;
            numOffsetRegistros.Value = config.OffsetRegistros >= 0 ? config.OffsetRegistros : 0;
                if (numEventosPorLote != null)
                {
                numEventosPorLote.Value = config.EventosPorLote > 0 && config.EventosPorLote <= 50 
                    ? config.EventosPorLote : 50;
                }
            if (config.MaxLotes.HasValue)
                {
                    chkMaxLotesIlimitado.Checked = false;
                numMaxLotes.Value = config.MaxLotes.Value;
                }
                else
                {
                    chkMaxLotesIlimitado.Checked = true;
                }
            }

        private void CarregarDadosAbertura(DadosAbertura dados)
        {
            // Preencher semestre e ano baseado nas datas
            if (!string.IsNullOrEmpty(dados.DtInicio) && !string.IsNullOrEmpty(dados.DtFim))
            {
                PreencherSemestreAnoAbertura(dados.DtInicio, dados.DtFim);
            }
            
                txtDtInicioAbertura.Text = dados.DtInicio ?? "";
                txtDtFimAbertura.Text = dados.DtFim ?? "";
                cmbTipoAmbienteAbertura.SelectedIndex = dados.TipoAmbiente > 0 ? dados.TipoAmbiente - 1 : 1;
                cmbAplicacaoEmissoraAbertura.SelectedIndex = dados.AplicacaoEmissora > 0 ? dados.AplicacaoEmissora - 1 : 0;
                cmbIndRetificacaoAbertura.SelectedIndex = dados.IndRetificacao > 0 ? dados.IndRetificacao - 1 : 0;
                txtNrReciboAbertura.Text = dados.NrRecibo ?? "";
                chkIndicarMovOpFin.Checked = dados.IndicarMovOpFin;

                if (dados.ResponsavelRMF != null)
                {
                CarregarDadosResponsavelRMF(dados.ResponsavelRMF);
                }

                if (dados.RespeFin != null)
                {
                CarregarDadosRespeFin(dados.RespeFin);
                }

                if (dados.RepresLegal != null)
                {
                CarregarDadosRepresLegal(dados.RepresLegal);
            }
        }

        private void CarregarDadosResponsavelRMF(DadosResponsavelRMF responsavelRMF)
        {
            txtRMF_CNPJ.Text = responsavelRMF.Cnpj ?? "";
            txtRMF_CPF.Text = responsavelRMF.Cpf ?? "";
            txtRMF_Nome.Text = responsavelRMF.Nome ?? "";
            txtRMF_Setor.Text = responsavelRMF.Setor ?? "";
            txtRMF_DDD.Text = responsavelRMF.TelefoneDDD ?? "";
            txtRMF_Telefone.Text = responsavelRMF.TelefoneNumero ?? "";
            txtRMF_Ramal.Text = responsavelRMF.TelefoneRamal ?? "";
            txtRMF_Logradouro.Text = responsavelRMF.EnderecoLogradouro ?? "";
            txtRMF_Numero.Text = responsavelRMF.EnderecoNumero ?? "";
            txtRMF_Complemento.Text = responsavelRMF.EnderecoComplemento ?? "";
            txtRMF_Bairro.Text = responsavelRMF.EnderecoBairro ?? "";
            txtRMF_CEP.Text = responsavelRMF.EnderecoCEP ?? "";
            txtRMF_Municipio.Text = responsavelRMF.EnderecoMunicipio ?? "";
            txtRMF_UF.Text = responsavelRMF.EnderecoUF ?? "";
        }

        private void CarregarDadosRespeFin(DadosRespeFin respeFin)
        {
            txtRespeFin_CPF.Text = respeFin.Cpf ?? "";
            txtRespeFin_Nome.Text = respeFin.Nome ?? "";
            txtRespeFin_Setor.Text = respeFin.Setor ?? "";
            txtRespeFin_DDD.Text = respeFin.TelefoneDDD ?? "";
            txtRespeFin_Telefone.Text = respeFin.TelefoneNumero ?? "";
            txtRespeFin_Ramal.Text = respeFin.TelefoneRamal ?? "";
            txtRespeFin_Logradouro.Text = respeFin.EnderecoLogradouro ?? "";
            txtRespeFin_Numero.Text = respeFin.EnderecoNumero ?? "";
            txtRespeFin_Complemento.Text = respeFin.EnderecoComplemento ?? "";
            txtRespeFin_Bairro.Text = respeFin.EnderecoBairro ?? "";
            txtRespeFin_CEP.Text = respeFin.EnderecoCEP ?? "";
            txtRespeFin_Municipio.Text = respeFin.EnderecoMunicipio ?? "";
            txtRespeFin_UF.Text = respeFin.EnderecoUF ?? "";
            txtRespeFin_Email.Text = respeFin.Email ?? "";
        }

        private void CarregarDadosRepresLegal(DadosRepresLegal represLegal)
        {
            txtRepresLegal_CPF.Text = represLegal.Cpf ?? "";
            txtRepresLegal_Setor.Text = represLegal.Setor ?? "";
            txtRepresLegal_DDD.Text = represLegal.TelefoneDDD ?? "";
            txtRepresLegal_Telefone.Text = represLegal.TelefoneNumero ?? "";
            txtRepresLegal_Ramal.Text = represLegal.TelefoneRamal ?? "";
        }

        private void CarregarDadosFechamento(DadosFechamento dados)
        {
            // Preencher semestre e ano baseado nas datas
            if (!string.IsNullOrEmpty(dados.DtInicio) && !string.IsNullOrEmpty(dados.DtFim))
            {
                PreencherSemestreAnoFechamento(dados.DtInicio, dados.DtFim);
            }
            
                txtDtInicioFechamento.Text = dados.DtInicio ?? "";
                txtDtFimFechamento.Text = dados.DtFim ?? "";
                cmbTipoAmbienteFechamento.SelectedIndex = dados.TipoAmbiente > 0 ? dados.TipoAmbiente - 1 : 1;
                cmbAplicacaoEmissoraFechamento.SelectedIndex = dados.AplicacaoEmissora > 0 ? dados.AplicacaoEmissora - 1 : 0;
                cmbIndRetificacaoFechamento.SelectedIndex = dados.IndRetificacao > 0 ? dados.IndRetificacao - 1 : 0;
                txtNrReciboFechamento.Text = dados.NrRecibo ?? "";
                cmbSitEspecial.SelectedIndex = Array.IndexOf(new[] { 0, 1, 2, 3, 5 }, dados.SitEspecial);
                chkNadaADeclarar.Checked = dados.NadaADeclarar == "1";
            chkFechamentoPP.Checked = dados.FechamentoPP > 0;
            chkFechamentoMovOpFin.Checked = dados.FechamentoMovOpFin > 0;
            chkFechamentoMovOpFinAnual.Checked = dados.FechamentoMovOpFinAnual > 0;
        }

        private void CarregarDadosCadastroDeclarante(DadosCadastroDeclarante dados)
        {
            cmbTipoAmbienteCadastro.SelectedIndex = dados.TipoAmbiente > 0 ? dados.TipoAmbiente - 1 : 1;
            cmbAplicacaoEmissoraCadastro.SelectedIndex = dados.AplicacaoEmissora > 0 ? dados.AplicacaoEmissora - 1 : 0;
            cmbIndRetificacaoCadastro.SelectedIndex = dados.IndRetificacao > 0 ? dados.IndRetificacao - 1 : 0;
            txtNrReciboCadastro.Text = dados.NrRecibo ?? "";
            txtGIIN.Text = dados.GIIN ?? "";
            txtCategoriaDeclarante.Text = dados.CategoriaDeclarante ?? "";
            txtNomeCadastro.Text = dados.Nome ?? "";
            txtTpNome.Text = dados.TpNome ?? "";
            txtEnderecoLivreCadastro.Text = dados.EnderecoLivre ?? "";
            txtTpEnderecoCadastro.Text = dados.TpEndereco ?? "";
            txtMunicipioCadastro.Text = dados.Municipio ?? "";
            txtUFCadastro.Text = dados.UF ?? "";
            txtCEPCadastro.Text = dados.CEP ?? "";
            txtPaisCadastro.Text = dados.Pais ?? "BR";
            
            if (dados.PaisResid != null && dados.PaisResid.Count > 0)
            {
                txtPaisResid.Text = string.Join(",", dados.PaisResid);
            }
            else
            {
                txtPaisResid.Text = "BR";
            }
        }

        private void CarregarConfiguracaoSalva()
        {
            try
            {
                var configCompleta = persistenciaService.CarregarConfiguracao();
                if (configCompleta != null)
                {
                    CarregarConfiguracaoNaTela(configCompleta);
                }
            }
            catch
            {
                // Ignora erros ao carregar configuração salva
            }
        }

        private void CarregarConfiguracoesPadrao()
        {
            // Configuração Geral
            if (string.IsNullOrEmpty(txtCnpjDeclarante.Text))
                txtCnpjDeclarante.Text = "40994589000105";
            if (string.IsNullOrEmpty(txtCertServidorThumbprint.Text))
                txtCertServidorThumbprint.Text = "cc242988a739caa7757b29e2a900ae35519cdb39";
            if (string.IsNullOrEmpty(txtDiretorioLotes.Text))
                txtDiretorioLotes.Text = Path.Combine(Application.StartupPath, "lotes");
            if (string.IsNullOrEmpty(txtPeriodo.Text))
            {
                // Usar período atual calculado automaticamente
                txtPeriodo.Text = EfinanceiraPeriodoUtil.CalcularPeriodoAtual();
            }

            // Abertura - Dados Básicos
            if (string.IsNullOrEmpty(txtDtInicioAbertura.Text))
                txtDtInicioAbertura.Text = DateTime.Now.AddMonths(-6).ToString("yyyy-MM-01");
            if (string.IsNullOrEmpty(txtDtFimAbertura.Text))
                txtDtFimAbertura.Text = DateTime.Now.ToString("yyyy-MM-dd");
            
            // Abertura - Responsável RMF (dados padrão do exemplo)
            if (txtRMF_CNPJ != null && string.IsNullOrEmpty(txtRMF_CNPJ.Text))
            {
                txtRMF_CNPJ.Text = "40994589000105";
                txtRMF_CPF.Text = "78325121726";
                txtRMF_Nome.Text = "Responsável RMF";
                txtRMF_Setor.Text = "Contabilidade";
                txtRMF_DDD.Text = "85";
                txtRMF_Telefone.Text = "30523500";
                txtRMF_Ramal.Text = "";
                txtRMF_Logradouro.Text = "Avenida Desembargador Moreira";
                txtRMF_Numero.Text = "1300";
                txtRMF_Complemento.Text = "Sala 1307 T- Sul";
                txtRMF_Bairro.Text = "Aldeota";
                txtRMF_CEP.Text = "60170002";
                txtRMF_Municipio.Text = "Fortaleza";
                txtRMF_UF.Text = "CE";
            }

            // Abertura - RespeFin (dados padrão do exemplo)
            if (txtRespeFin_CPF != null && string.IsNullOrEmpty(txtRespeFin_CPF.Text))
            {
                txtRespeFin_CPF.Text = "39768628197";
                txtRespeFin_Nome.Text = "Responsável e-Financeira";
                txtRespeFin_Setor.Text = "Contabilidade";
                txtRespeFin_DDD.Text = "85";
                txtRespeFin_Telefone.Text = "30523500";
                txtRespeFin_Ramal.Text = "";
                txtRespeFin_Logradouro.Text = "Avenida Desembargador Moreira";
                txtRespeFin_Numero.Text = "1300";
                txtRespeFin_Complemento.Text = "Sala 1307 T- Sul";
                txtRespeFin_Bairro.Text = "Aldeota";
                txtRespeFin_CEP.Text = "60170002";
                txtRespeFin_Municipio.Text = "Fortaleza";
                txtRespeFin_UF.Text = "CE";
                txtRespeFin_Email.Text = "contabilidade@bscash.com.br";
            }

            // Abertura - RepresLegal (dados padrão do exemplo)
            if (txtRepresLegal_CPF != null && string.IsNullOrEmpty(txtRepresLegal_CPF.Text))
            {
                txtRepresLegal_CPF.Text = "09684527870";
                txtRepresLegal_Setor.Text = "Diretoria";
                txtRepresLegal_DDD.Text = "11";
                txtRepresLegal_Telefone.Text = "29760435";
                txtRepresLegal_Ramal.Text = "";
            }

            // Marcar IndicarMovOpFin como true se houver dados preenchidos
            if (chkIndicarMovOpFin != null && !chkIndicarMovOpFin.Checked && 
                (!string.IsNullOrEmpty(txtRMF_CPF.Text) || !string.IsNullOrEmpty(txtRespeFin_CPF.Text)))
                {
                    chkIndicarMovOpFin.Checked = true;
            }

            // Fechamento
            if (string.IsNullOrEmpty(txtDtInicioFechamento.Text))
                txtDtInicioFechamento.Text = DateTime.Now.AddMonths(-6).ToString("yyyy-MM-01");
            if (string.IsNullOrEmpty(txtDtFimFechamento.Text))
                txtDtFimFechamento.Text = DateTime.Now.ToString("yyyy-MM-dd");
            
            // Aplicar valores padrão baseado no ambiente selecionado (apenas se os controles já existirem)
            if (numPageSize != null && numEventoOffset != null)
            {
                CmbAmbiente_SelectedIndexChanged(null, null);
            }
        }
    }
}
