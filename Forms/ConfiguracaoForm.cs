using System;
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
        private ScrollableControl scrollAbertura;
        private ScrollableControl scrollFechamento;

        // Abertura - Básicos
        private TextBox txtDtInicioAbertura;
        private TextBox txtDtFimAbertura;
        private ComboBox cmbTipoAmbienteAbertura;
        private ComboBox cmbAplicacaoEmissoraAbertura;
        private ComboBox cmbIndRetificacaoAbertura;
        private TextBox txtNrReciboAbertura;
        private CheckBox chkIndicarMovOpFin;

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
        private NumericUpDown numFechamentoPP;
        private NumericUpDown numFechamentoMovOpFin;
        private NumericUpDown numFechamentoMovOpFinAnual;

        private Button btnSalvarConfig;
        private ConfiguracaoPersistenciaService persistenciaService;

        public DadosAbertura DadosAbertura { get; private set; }
        public DadosFechamento DadosFechamento { get; private set; }
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
            this.Size = new Size(1000, 700);
            this.MinimumSize = new Size(800, 600);

            // ToolTip
            toolTip = new ToolTip();
            toolTip.IsBalloon = true;
            toolTip.ToolTipTitle = "Ajuda";
            toolTip.ToolTipIcon = ToolTipIcon.Info;

            // Configuração Geral
            grpConfigGeral = new GroupBox();
            grpConfigGeral.Text = "Configuração Geral";
            grpConfigGeral.Location = new Point(10, 10);
            grpConfigGeral.Size = new Size(980, 280);
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
            yPos += espacamentoVertical;

            // Linha 2: Certificado para Assinatura
            lblCertThumbprint = new Label { Text = "Certificado para Assinatura:", Location = new Point(10, yPos), Size = new Size(labelWidth, 20) };
            txtCertThumbprint = new TextBox { Location = new Point(campoX, yPos - 3), Size = new Size(campoWidth, 23), ReadOnly = true };
            btnSelecionarCert = new Button { Text = "Selecionar...", Location = new Point(campoX + campoWidth + 10, yPos - 3), Size = new Size(100, 25) };
            btnSelecionarCert.Click += BtnSelecionarCert_Click;
            yPos += espacamentoVertical;

            // Linha 3: Certificado do Servidor
            lblCertServidorThumbprint = new Label { Text = "Certificado do Servidor:", Location = new Point(10, yPos), Size = new Size(labelWidth, 20) };
            txtCertServidorThumbprint = new TextBox { Location = new Point(campoX, yPos - 3), Size = new Size(campoWidth, 23), ReadOnly = true };
            btnSelecionarCertServidor = new Button { Text = "Selecionar...", Location = new Point(campoX + campoWidth + 10, yPos - 3), Size = new Size(100, 25) };
            btnSelecionarCertServidor.Click += BtnSelecionarCertServidor_Click;
            yPos += espacamentoVertical;

            // Linha 4: Período
            lblPeriodo = new Label { Text = "Período (YYYYMM - 01/06=Jan-Jun, 02/12=Jul-Dez):", Location = new Point(10, yPos), Size = new Size(labelWidth, 20) };
            txtPeriodo = new TextBox { Location = new Point(campoX, yPos - 3), Size = new Size(200, 23) };
            txtPeriodo.MaxLength = 6;
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
                                "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                                "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            yPos += espacamentoVertical;

            // Linha 6: Ambiente e Opções
            lblAmbiente = new Label { Text = "Ambiente:", Location = new Point(10, yPos), Size = new Size(labelWidth, 20) };
            cmbAmbiente = new ComboBox { Location = new Point(campoX, yPos - 3), Size = new Size(100, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbAmbiente.Items.AddRange(new[] { "TEST", "PROD" });
            cmbAmbiente.SelectedIndex = 0;

            chkModoTeste = new CheckBox { Text = "Modo Teste", Location = new Point(campoX + 120, yPos), Size = new Size(120, 20), Checked = true };
            chkHabilitarEnvio = new CheckBox { Text = "Habilitar Envio", Location = new Point(campoX + 250, yPos), Size = new Size(150, 20) };
            yPos += espacamentoVertical;

            // Linha 7: Botões
            btnCarregarConfig = new Button { Text = "Carregar Configuração", Location = new Point(10, yPos), Size = new Size(150, 30) };
            btnCarregarConfig.Click += BtnCarregarConfig_Click;

            btnTestarConexao = new Button { Text = "Testar Conexão BD", Location = new Point(170, yPos), Size = new Size(150, 30) };
            btnTestarConexao.Click += BtnTestarConexao_Click;

            grpConfigGeral.Controls.AddRange(new Control[] {
                lblCnpjDeclarante, txtCnpjDeclarante,
                lblCertThumbprint, txtCertThumbprint, btnSelecionarCert,
                lblCertServidorThumbprint, txtCertServidorThumbprint, btnSelecionarCertServidor,
                lblPeriodo, txtPeriodo,
                lblDiretorioLotes, txtDiretorioLotes, btnSelecionarDiretorio,
                lblAmbiente, cmbAmbiente, chkModoTeste, chkHabilitarEnvio,
                btnCarregarConfig, btnTestarConexao
            });

            // Configurações de Processamento
            CriarGrupoProcessamento();

            // TabControl
            tabConfig = new TabControl();
            tabConfig.Location = new Point(10, 410);
            tabConfig.Size = new Size(980, 210);
            tabConfig.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            tabAbertura = new TabPage { Text = "Abertura", UseVisualStyleBackColor = true };
            tabFechamento = new TabPage { Text = "Fechamento", UseVisualStyleBackColor = true };
            tabConfig.TabPages.Add(tabAbertura);
            tabConfig.TabPages.Add(tabFechamento);

            CriarAbaAbertura();
            CriarAbaFechamento();

            // Botão Salvar
            btnSalvarConfig = new Button { Text = "Salvar Configurações", Location = new Point(10, 630), Size = new Size(150, 35), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            btnSalvarConfig.Click += BtnSalvarConfig_Click;

            // Evento para ajustar valores padrão quando ambiente mudar
            cmbAmbiente.SelectedIndexChanged += CmbAmbiente_SelectedIndexChanged;

            this.Controls.AddRange(new Control[] { grpConfigGeral, grpProcessamento, tabConfig, btnSalvarConfig });
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
            grpProcessamento.Location = new Point(10, 300);
            grpProcessamento.Size = new Size(980, 100);
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
            var grpBasicos = new GroupBox { Text = "Dados Básicos", Location = new Point(10, yPos), Size = new Size(940, 120) };
            yPos += 130;

            var lblDtInicio = new Label { Text = "Data Início (AAAA-MM-DD):", Location = new Point(10, 25), Size = new Size(150, 20) };
            txtDtInicioAbertura = new TextBox { Location = new Point(165, 22), Size = new Size(150, 23) };
            txtDtInicioAbertura.Leave += ValidarData;

            var lblDtFim = new Label { Text = "Data Fim (AAAA-MM-DD):", Location = new Point(330, 25), Size = new Size(150, 20) };
            txtDtFimAbertura = new TextBox { Location = new Point(485, 22), Size = new Size(150, 23) };
            txtDtFimAbertura.Leave += ValidarData;

            var lblTipoAmbiente = new Label { Text = "Tipo Ambiente:", Location = new Point(10, 55), Size = new Size(100, 20) };
            cmbTipoAmbienteAbertura = new ComboBox { Location = new Point(115, 52), Size = new Size(150, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTipoAmbienteAbertura.Items.AddRange(new[] { "1 - Produção", "2 - Homologação" });
            cmbTipoAmbienteAbertura.SelectedIndex = 1;

            var lblAplicacaoEmissora = new Label { Text = "Aplicação Emissora:", Location = new Point(280, 55), Size = new Size(120, 20) };
            cmbAplicacaoEmissoraAbertura = new ComboBox { Location = new Point(405, 52), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbAplicacaoEmissoraAbertura.Items.AddRange(new[] { "1 - Aplicação do contribuinte", "2 - Outros" });
            cmbAplicacaoEmissoraAbertura.SelectedIndex = 0;

            var lblIndRetificacao = new Label { Text = "Ind Retificação:", Location = new Point(10, 85), Size = new Size(100, 20) };
            cmbIndRetificacaoAbertura = new ComboBox { Location = new Point(115, 82), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbIndRetificacaoAbertura.Items.AddRange(new[] { "1 - Original", "2 - Retificação espontânea", "3 - Retificação a pedido" });
            cmbIndRetificacaoAbertura.SelectedIndex = 0;
            cmbIndRetificacaoAbertura.SelectedIndexChanged += CmbIndRetificacaoAbertura_SelectedIndexChanged;

            var lblNrRecibo = new Label { Text = "Nº Recibo:", Location = new Point(330, 85), Size = new Size(80, 20) };
            txtNrReciboAbertura = new TextBox { Location = new Point(415, 82), Size = new Size(150, 23), Enabled = false };

            chkIndicarMovOpFin = new CheckBox { Text = "Indicar MovOpFin", Location = new Point(580, 25), Size = new Size(150, 20) };

            grpBasicos.Controls.AddRange(new Control[] {
                lblDtInicio, txtDtInicioAbertura, lblDtFim, txtDtFimAbertura,
                lblTipoAmbiente, cmbTipoAmbienteAbertura, lblAplicacaoEmissora, cmbAplicacaoEmissoraAbertura,
                lblIndRetificacao, cmbIndRetificacaoAbertura, lblNrRecibo, txtNrReciboAbertura,
                chkIndicarMovOpFin
            });

            // ResponsavelRMF
            var grpRMF = new GroupBox { Text = "Responsável RMF", Location = new Point(10, yPos), Size = new Size(940, 200) };
            yPos += 210;

            CriarCamposResponsavelRMF(grpRMF, 10);

            // RespeFin
            var grpRespeFin = new GroupBox { Text = "Responsável e-Financeira", Location = new Point(10, yPos), Size = new Size(940, 220) };
            yPos += 230;

            CriarCamposRespeFin(grpRespeFin, 10);

            // RepresLegal
            var grpRepresLegal = new GroupBox { Text = "Representante Legal", Location = new Point(10, yPos), Size = new Size(940, 100) };
            yPos += 110;

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

            var lblCPF = new Label { Text = "CPF:", Location = new Point(x + 250, y), Size = new Size(80, 20) };
            txtRMF_CPF = new TextBox { Location = new Point(x + 335, y - 3), Size = new Size(150, 23) };
            txtRMF_CPF.Leave += ValidarCPF;

            var lblNome = new Label { Text = "Nome:", Location = new Point(x + 500, y), Size = new Size(80, 20) };
            txtRMF_Nome = new TextBox { Location = new Point(x + 585, y - 3), Size = new Size(300, 23) };

            y += 35;
            var lblSetor = new Label { Text = "Setor:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRMF_Setor = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(200, 23) };

            var lblDDD = new Label { Text = "DDD:", Location = new Point(x + 300, y), Size = new Size(50, 20) };
            txtRMF_DDD = new TextBox { Location = new Point(x + 355, y - 3), Size = new Size(50, 23) };

            var lblTelefone = new Label { Text = "Telefone:", Location = new Point(x + 420, y), Size = new Size(70, 20) };
            txtRMF_Telefone = new TextBox { Location = new Point(x + 495, y - 3), Size = new Size(100, 23) };

            var lblRamal = new Label { Text = "Ramal:", Location = new Point(x + 610, y), Size = new Size(60, 20) };
            txtRMF_Ramal = new TextBox { Location = new Point(x + 675, y - 3), Size = new Size(80, 23) };

            y += 35;
            var lblLogradouro = new Label { Text = "Logradouro:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRMF_Logradouro = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(300, 23) };

            var lblNumero = new Label { Text = "Número:", Location = new Point(x + 400, y), Size = new Size(60, 20) };
            txtRMF_Numero = new TextBox { Location = new Point(x + 465, y - 3), Size = new Size(80, 23) };

            var lblComplemento = new Label { Text = "Complemento:", Location = new Point(x + 560, y), Size = new Size(90, 20) };
            txtRMF_Complemento = new TextBox { Location = new Point(x + 655, y - 3), Size = new Size(200, 23) };

            y += 35;
            var lblBairro = new Label { Text = "Bairro:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRMF_Bairro = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(200, 23) };

            var lblCEP = new Label { Text = "CEP:", Location = new Point(x + 300, y), Size = new Size(50, 20) };
            txtRMF_CEP = new TextBox { Location = new Point(x + 355, y - 3), Size = new Size(100, 23) };
            txtRMF_CEP.Leave += ValidarCEP;

            var lblMunicipio = new Label { Text = "Município:", Location = new Point(x + 470, y), Size = new Size(70, 20) };
            txtRMF_Municipio = new TextBox { Location = new Point(x + 545, y - 3), Size = new Size(200, 23) };

            var lblUF = new Label { Text = "UF:", Location = new Point(x + 760, y), Size = new Size(30, 20) };
            txtRMF_UF = new TextBox { Location = new Point(x + 795, y - 3), Size = new Size(50, 23), MaxLength = 2 };
            txtRMF_UF.CharacterCasing = CharacterCasing.Upper;

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

            var lblNome = new Label { Text = "Nome:", Location = new Point(x + 250, y), Size = new Size(80, 20) };
            txtRespeFin_Nome = new TextBox { Location = new Point(x + 335, y - 3), Size = new Size(300, 23) };

            var lblEmail = new Label { Text = "Email:", Location = new Point(x + 650, y), Size = new Size(60, 20) };
            txtRespeFin_Email = new TextBox { Location = new Point(x + 715, y - 3), Size = new Size(200, 23) };
            txtRespeFin_Email.Leave += ValidarEmail;

            y += 35;
            var lblSetor = new Label { Text = "Setor:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRespeFin_Setor = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(200, 23) };

            var lblDDD = new Label { Text = "DDD:", Location = new Point(x + 300, y), Size = new Size(50, 20) };
            txtRespeFin_DDD = new TextBox { Location = new Point(x + 355, y - 3), Size = new Size(50, 23) };

            var lblTelefone = new Label { Text = "Telefone:", Location = new Point(x + 420, y), Size = new Size(70, 20) };
            txtRespeFin_Telefone = new TextBox { Location = new Point(x + 495, y - 3), Size = new Size(100, 23) };

            var lblRamal = new Label { Text = "Ramal:", Location = new Point(x + 610, y), Size = new Size(60, 20) };
            txtRespeFin_Ramal = new TextBox { Location = new Point(x + 675, y - 3), Size = new Size(80, 23) };

            y += 35;
            var lblLogradouro = new Label { Text = "Logradouro:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRespeFin_Logradouro = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(300, 23) };

            var lblNumero = new Label { Text = "Número:", Location = new Point(x + 400, y), Size = new Size(60, 20) };
            txtRespeFin_Numero = new TextBox { Location = new Point(x + 465, y - 3), Size = new Size(80, 23) };

            var lblComplemento = new Label { Text = "Complemento:", Location = new Point(x + 560, y), Size = new Size(90, 20) };
            txtRespeFin_Complemento = new TextBox { Location = new Point(x + 655, y - 3), Size = new Size(200, 23) };

            y += 35;
            var lblBairro = new Label { Text = "Bairro:", Location = new Point(x, y), Size = new Size(80, 20) };
            txtRespeFin_Bairro = new TextBox { Location = new Point(x + 85, y - 3), Size = new Size(200, 23) };

            var lblCEP = new Label { Text = "CEP:", Location = new Point(x + 300, y), Size = new Size(50, 20) };
            txtRespeFin_CEP = new TextBox { Location = new Point(x + 355, y - 3), Size = new Size(100, 23) };
            txtRespeFin_CEP.Leave += ValidarCEP;

            var lblMunicipio = new Label { Text = "Município:", Location = new Point(x + 470, y), Size = new Size(70, 20) };
            txtRespeFin_Municipio = new TextBox { Location = new Point(x + 545, y - 3), Size = new Size(200, 23) };

            var lblUF = new Label { Text = "UF:", Location = new Point(x + 760, y), Size = new Size(30, 20) };
            txtRespeFin_UF = new TextBox { Location = new Point(x + 795, y - 3), Size = new Size(50, 23), MaxLength = 2 };
            txtRespeFin_UF.CharacterCasing = CharacterCasing.Upper;

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

            var lblSetor = new Label { Text = "Setor:", Location = new Point(x + 250, y), Size = new Size(80, 20) };
            txtRepresLegal_Setor = new TextBox { Location = new Point(x + 335, y - 3), Size = new Size(200, 23) };

            var lblDDD = new Label { Text = "DDD:", Location = new Point(x + 550, y), Size = new Size(50, 20) };
            txtRepresLegal_DDD = new TextBox { Location = new Point(x + 605, y - 3), Size = new Size(50, 23) };

            var lblTelefone = new Label { Text = "Telefone:", Location = new Point(x + 670, y), Size = new Size(70, 20) };
            txtRepresLegal_Telefone = new TextBox { Location = new Point(x + 745, y - 3), Size = new Size(100, 23) };

            var lblRamal = new Label { Text = "Ramal:", Location = new Point(x + 860, y), Size = new Size(60, 20) };
            txtRepresLegal_Ramal = new TextBox { Location = new Point(x + 925, y - 3), Size = new Size(80, 23) };

            grp.Controls.AddRange(new Control[] {
                lblCPF, txtRepresLegal_CPF, lblSetor, txtRepresLegal_Setor,
                lblDDD, txtRepresLegal_DDD, lblTelefone, txtRepresLegal_Telefone, lblRamal, txtRepresLegal_Ramal
            });
        }

        private void CriarAbaFechamento()
        {
            scrollFechamento = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            int yPos = 10;

            var grpBasicos = new GroupBox { Text = "Dados Básicos", Location = new Point(10, yPos), Size = new Size(940, 250) };
            yPos += 260;

            var lblDtInicio = new Label { Text = "Data Início (AAAA-MM-DD):", Location = new Point(10, 25), Size = new Size(150, 20) };
            txtDtInicioFechamento = new TextBox { Location = new Point(165, 22), Size = new Size(150, 23) };
            txtDtInicioFechamento.Leave += ValidarData;

            var lblDtFim = new Label { Text = "Data Fim (AAAA-MM-DD):", Location = new Point(330, 25), Size = new Size(150, 20) };
            txtDtFimFechamento = new TextBox { Location = new Point(485, 22), Size = new Size(150, 23) };
            txtDtFimFechamento.Leave += ValidarData;

            var lblTipoAmbiente = new Label { Text = "Tipo Ambiente:", Location = new Point(10, 55), Size = new Size(100, 20) };
            cmbTipoAmbienteFechamento = new ComboBox { Location = new Point(115, 52), Size = new Size(150, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTipoAmbienteFechamento.Items.AddRange(new[] { "1 - Produção", "2 - Homologação" });
            cmbTipoAmbienteFechamento.SelectedIndex = 1;

            var lblAplicacaoEmissora = new Label { Text = "Aplicação Emissora:", Location = new Point(280, 55), Size = new Size(120, 20) };
            cmbAplicacaoEmissoraFechamento = new ComboBox { Location = new Point(405, 52), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbAplicacaoEmissoraFechamento.Items.AddRange(new[] { "1 - Aplicação do contribuinte", "2 - Outros" });
            cmbAplicacaoEmissoraFechamento.SelectedIndex = 0;

            var lblIndRetificacao = new Label { Text = "Ind Retificação:", Location = new Point(10, 85), Size = new Size(100, 20) };
            cmbIndRetificacaoFechamento = new ComboBox { Location = new Point(115, 82), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbIndRetificacaoFechamento.Items.AddRange(new[] { "1 - Original", "2 - Retificação espontânea", "3 - Retificação a pedido" });
            cmbIndRetificacaoFechamento.SelectedIndex = 0;
            cmbIndRetificacaoFechamento.SelectedIndexChanged += CmbIndRetificacaoFechamento_SelectedIndexChanged;

            var lblNrRecibo = new Label { Text = "Nº Recibo:", Location = new Point(330, 85), Size = new Size(80, 20) };
            txtNrReciboFechamento = new TextBox { Location = new Point(415, 82), Size = new Size(150, 23), Enabled = false };

            var lblSitEspecial = new Label { Text = "Situação Especial:", Location = new Point(10, 115), Size = new Size(120, 20) };
            cmbSitEspecial = new ComboBox { Location = new Point(135, 112), Size = new Size(200, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbSitEspecial.Items.AddRange(new[] { "0 - Não se aplica", "1 - Extinção", "2 - Fusão", "3 - Incorporação", "5 - Cisão" });
            cmbSitEspecial.SelectedIndex = 0;

            chkNadaADeclarar = new CheckBox { Text = "Nada a Declarar", Location = new Point(350, 115), Size = new Size(150, 20) };

            var lblFechamentoPP = new Label { Text = "FechamentoPP:", Location = new Point(10, 145), Size = new Size(120, 20) };
            numFechamentoPP = new NumericUpDown { Location = new Point(135, 142), Size = new Size(60, 23), Minimum = 0, Maximum = 1, DecimalPlaces = 0 };
            toolTip.SetToolTip(numFechamentoPP, "0 = Sem movimento de Previdência Privada\n1 = Com movimento de Previdência Privada");
            toolTip.SetToolTip(lblFechamentoPP, "0 = Sem movimento de Previdência Privada\n1 = Com movimento de Previdência Privada");

            var lblFechamentoMovOpFin = new Label { Text = "FechamentoMovOpFin:", Location = new Point(210, 145), Size = new Size(150, 20) };
            numFechamentoMovOpFin = new NumericUpDown { Location = new Point(365, 142), Size = new Size(60, 23), Minimum = 0, Maximum = 1, DecimalPlaces = 0 };
            toolTip.SetToolTip(numFechamentoMovOpFin, "0 = Sem movimento de Operação Financeira\n1 = Com movimento de Operação Financeira");
            toolTip.SetToolTip(lblFechamentoMovOpFin, "0 = Sem movimento de Operação Financeira\n1 = Com movimento de Operação Financeira");

            var lblFechamentoMovOpFinAnual = new Label { Text = "FechamentoMovOpFinAnual:", Location = new Point(440, 145), Size = new Size(180, 20) };
            numFechamentoMovOpFinAnual = new NumericUpDown { Location = new Point(625, 142), Size = new Size(60, 23), Minimum = 0, Maximum = 1, DecimalPlaces = 0 };
            toolTip.SetToolTip(numFechamentoMovOpFinAnual, "0 = Sem movimento de Operação Financeira Anual\n1 = Com movimento de Operação Financeira Anual");
            toolTip.SetToolTip(lblFechamentoMovOpFinAnual, "0 = Sem movimento de Operação Financeira Anual\n1 = Com movimento de Operação Financeira Anual");

            // Legenda explicativa
            var lblLegenda = new Label 
            { 
                Text = "ℹ️ IMPORTANTE: Se 'Nada a Declarar' NÃO estiver marcado, você DEVE preencher pelo menos um dos campos de fechamento acima (FechamentoPP, FechamentoMovOpFin ou FechamentoMovOpFinAnual).\n" +
                       "Valores: 0 = Sem movimento | 1 = Com movimento\n" +
                       "Exemplo: Se você enviou movimentações financeiras, deve preencher FechamentoMovOpFin com 0 (sem movimento adicional) ou 1 (com movimento adicional).",
                Location = new Point(10, 175), 
                Size = new Size(920, 60),
                ForeColor = Color.DarkBlue,
                Font = new Font("Microsoft Sans Serif", 8.5f, FontStyle.Regular)
            };

            toolTip.SetToolTip(chkNadaADeclarar, "Marque esta opção se não houver nada a declarar no período.\nSe marcado, não é necessário preencher os campos de fechamento.");

            grpBasicos.Controls.AddRange(new Control[] {
                lblDtInicio, txtDtInicioFechamento, lblDtFim, txtDtFimFechamento,
                lblTipoAmbiente, cmbTipoAmbienteFechamento, lblAplicacaoEmissora, cmbAplicacaoEmissoraFechamento,
                lblIndRetificacao, cmbIndRetificacaoFechamento, lblNrRecibo, txtNrReciboFechamento,
                lblSitEspecial, cmbSitEspecial, chkNadaADeclarar,
                lblFechamentoPP, numFechamentoPP, lblFechamentoMovOpFin, numFechamentoMovOpFin,
                lblFechamentoMovOpFinAnual, numFechamentoMovOpFinAnual,
                lblLegenda
            });

            scrollFechamento.Controls.Add(grpBasicos);
            tabFechamento.Controls.Add(scrollFechamento);
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
                MessageBox.Show("CNPJ inválido. Deve conter 14 dígitos.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MessageBox.Show("CPF inválido. Deve conter 11 dígitos.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MessageBox.Show("CEP inválido. Deve conter 8 dígitos.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txt.Focus();
            }
        }

        private void ValidarData(object sender, EventArgs e)
        {
            var txt = sender as TextBox;
            if (txt == null) return;

            if (string.IsNullOrEmpty(txt.Text)) return;

            if (!DateTime.TryParseExact(txt.Text, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _))
            {
                MessageBox.Show("Data inválida. Use o formato AAAA-MM-DD.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MessageBox.Show("Email inválido.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txt.Focus();
            }
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

        private X509Certificate2 SelecionarCertificado()
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
                    UrlTeste = "https://pre-efinanceira.receita.fazenda.gov.br/recepcao/lotes/cripto",
                    UrlProducao = "https://efinanceira.receita.fazenda.gov.br/recepcao/lotes/cripto",
                    UrlConsultaTeste = "https://pre-efinanceira.receita.fazenda.gov.br/consulta/lotes/",
                    UrlConsultaProducao = "https://efinanceira.receita.fazenda.gov.br/consulta/lotes/"
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
                    FechamentoPP = numFechamentoPP.Value > 0 ? (int?)numFechamentoPP.Value : null,
                    FechamentoMovOpFin = numFechamentoMovOpFin.Value > 0 ? (int?)numFechamentoMovOpFin.Value : null,
                    FechamentoMovOpFinAnual = numFechamentoMovOpFinAnual.Value > 0 ? (int?)numFechamentoMovOpFinAnual.Value : null
                };

                // Persistir
                persistenciaService.SalvarConfiguracao(Config, DadosAbertura, DadosFechamento);

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
                MessageBox.Show("CNPJ Declarante é obrigatório.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtCnpjDeclarante.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtDtInicioAbertura.Text))
            {
                MessageBox.Show("Data de início da abertura é obrigatória.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtDtInicioAbertura.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtDtFimAbertura.Text))
            {
                MessageBox.Show("Data de fim da abertura é obrigatória.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtDtFimAbertura.Focus();
                return false;
            }

            if (chkIndicarMovOpFin.Checked)
            {
                if (string.IsNullOrWhiteSpace(txtRMF_CPF.Text) || string.IsNullOrWhiteSpace(txtRespeFin_CPF.Text))
                {
                    MessageBox.Show("Ao indicar MovOpFin, é necessário preencher CPF do Responsável RMF e RespeFin.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            // Validação de fechamento: se não for "nada a declarar", deve ter pelo menos um fechamento
            if (!chkNadaADeclarar.Checked)
            {
                bool temFechamentoPP = numFechamentoPP.Value > 0;
                bool temFechamentoMovOpFin = numFechamentoMovOpFin.Value > 0;
                bool temFechamentoMovOpFinAnual = numFechamentoMovOpFinAnual.Value > 0;

                if (!temFechamentoPP && !temFechamentoMovOpFin && !temFechamentoMovOpFinAnual)
                {
                    MessageBox.Show(
                        "Se 'Nada a Declarar' não estiver marcado, você DEVE preencher pelo menos um dos campos de fechamento:\n\n" +
                        "• FechamentoPP (0 = sem movimento, 1 = com movimento)\n" +
                        "• FechamentoMovOpFin (0 = sem movimento, 1 = com movimento)\n" +
                        "• FechamentoMovOpFinAnual (0 = sem movimento, 1 = com movimento)\n\n" +
                        "Exemplo: Se você enviou movimentações financeiras, deve preencher FechamentoMovOpFin com 0 ou 1.",
                        "Validação de Fechamento",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    numFechamentoMovOpFin.Focus();
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

        private void CarregarConfiguracaoNaTela(ConfiguracaoCompleta configCompleta)
        {
            if (configCompleta.Config != null)
            {
                txtCnpjDeclarante.Text = configCompleta.Config.CnpjDeclarante ?? "";
                txtCertThumbprint.Text = configCompleta.Config.CertThumbprint ?? "";
                txtCertServidorThumbprint.Text = configCompleta.Config.CertServidorThumbprint ?? "";
                cmbAmbiente.SelectedItem = configCompleta.Config.Ambiente == EfinanceiraAmbiente.PROD ? "PROD" : "TEST";
                chkModoTeste.Checked = configCompleta.Config.ModoTeste;
                chkHabilitarEnvio.Checked = configCompleta.Config.TestEnvioHabilitado;
                txtDiretorioLotes.Text = configCompleta.Config.DiretorioLotes ?? "";
                txtPeriodo.Text = configCompleta.Config.Periodo ?? "";
                
                // Carregar configurações de processamento
                numPageSize.Value = configCompleta.Config.PageSize > 0 ? configCompleta.Config.PageSize : 500;
                numEventoOffset.Value = configCompleta.Config.EventoOffset >= 0 ? configCompleta.Config.EventoOffset : 1;
                numOffsetRegistros.Value = configCompleta.Config.OffsetRegistros >= 0 ? configCompleta.Config.OffsetRegistros : 0;
                if (numEventosPorLote != null)
                {
                    numEventosPorLote.Value = configCompleta.Config.EventosPorLote > 0 && configCompleta.Config.EventosPorLote <= 50 
                        ? configCompleta.Config.EventosPorLote : 50;
                }
                if (configCompleta.Config.MaxLotes.HasValue)
                {
                    chkMaxLotesIlimitado.Checked = false;
                    numMaxLotes.Value = configCompleta.Config.MaxLotes.Value;
                }
                else
                {
                    chkMaxLotesIlimitado.Checked = true;
                }
            }

            if (configCompleta.DadosAbertura != null)
            {
                var dados = configCompleta.DadosAbertura;
                txtDtInicioAbertura.Text = dados.DtInicio ?? "";
                txtDtFimAbertura.Text = dados.DtFim ?? "";
                cmbTipoAmbienteAbertura.SelectedIndex = dados.TipoAmbiente > 0 ? dados.TipoAmbiente - 1 : 1;
                cmbAplicacaoEmissoraAbertura.SelectedIndex = dados.AplicacaoEmissora > 0 ? dados.AplicacaoEmissora - 1 : 0;
                cmbIndRetificacaoAbertura.SelectedIndex = dados.IndRetificacao > 0 ? dados.IndRetificacao - 1 : 0;
                txtNrReciboAbertura.Text = dados.NrRecibo ?? "";
                chkIndicarMovOpFin.Checked = dados.IndicarMovOpFin;

                if (dados.ResponsavelRMF != null)
                {
                    txtRMF_CNPJ.Text = dados.ResponsavelRMF.Cnpj ?? "";
                    txtRMF_CPF.Text = dados.ResponsavelRMF.Cpf ?? "";
                    txtRMF_Nome.Text = dados.ResponsavelRMF.Nome ?? "";
                    txtRMF_Setor.Text = dados.ResponsavelRMF.Setor ?? "";
                    txtRMF_DDD.Text = dados.ResponsavelRMF.TelefoneDDD ?? "";
                    txtRMF_Telefone.Text = dados.ResponsavelRMF.TelefoneNumero ?? "";
                    txtRMF_Ramal.Text = dados.ResponsavelRMF.TelefoneRamal ?? "";
                    txtRMF_Logradouro.Text = dados.ResponsavelRMF.EnderecoLogradouro ?? "";
                    txtRMF_Numero.Text = dados.ResponsavelRMF.EnderecoNumero ?? "";
                    txtRMF_Complemento.Text = dados.ResponsavelRMF.EnderecoComplemento ?? "";
                    txtRMF_Bairro.Text = dados.ResponsavelRMF.EnderecoBairro ?? "";
                    txtRMF_CEP.Text = dados.ResponsavelRMF.EnderecoCEP ?? "";
                    txtRMF_Municipio.Text = dados.ResponsavelRMF.EnderecoMunicipio ?? "";
                    txtRMF_UF.Text = dados.ResponsavelRMF.EnderecoUF ?? "";
                }

                if (dados.RespeFin != null)
                {
                    txtRespeFin_CPF.Text = dados.RespeFin.Cpf ?? "";
                    txtRespeFin_Nome.Text = dados.RespeFin.Nome ?? "";
                    txtRespeFin_Setor.Text = dados.RespeFin.Setor ?? "";
                    txtRespeFin_DDD.Text = dados.RespeFin.TelefoneDDD ?? "";
                    txtRespeFin_Telefone.Text = dados.RespeFin.TelefoneNumero ?? "";
                    txtRespeFin_Ramal.Text = dados.RespeFin.TelefoneRamal ?? "";
                    txtRespeFin_Logradouro.Text = dados.RespeFin.EnderecoLogradouro ?? "";
                    txtRespeFin_Numero.Text = dados.RespeFin.EnderecoNumero ?? "";
                    txtRespeFin_Complemento.Text = dados.RespeFin.EnderecoComplemento ?? "";
                    txtRespeFin_Bairro.Text = dados.RespeFin.EnderecoBairro ?? "";
                    txtRespeFin_CEP.Text = dados.RespeFin.EnderecoCEP ?? "";
                    txtRespeFin_Municipio.Text = dados.RespeFin.EnderecoMunicipio ?? "";
                    txtRespeFin_UF.Text = dados.RespeFin.EnderecoUF ?? "";
                    txtRespeFin_Email.Text = dados.RespeFin.Email ?? "";
                }

                if (dados.RepresLegal != null)
                {
                    txtRepresLegal_CPF.Text = dados.RepresLegal.Cpf ?? "";
                    txtRepresLegal_Setor.Text = dados.RepresLegal.Setor ?? "";
                    txtRepresLegal_DDD.Text = dados.RepresLegal.TelefoneDDD ?? "";
                    txtRepresLegal_Telefone.Text = dados.RepresLegal.TelefoneNumero ?? "";
                    txtRepresLegal_Ramal.Text = dados.RepresLegal.TelefoneRamal ?? "";
                }
            }

            if (configCompleta.DadosFechamento != null)
            {
                var dados = configCompleta.DadosFechamento;
                txtDtInicioFechamento.Text = dados.DtInicio ?? "";
                txtDtFimFechamento.Text = dados.DtFim ?? "";
                cmbTipoAmbienteFechamento.SelectedIndex = dados.TipoAmbiente > 0 ? dados.TipoAmbiente - 1 : 1;
                cmbAplicacaoEmissoraFechamento.SelectedIndex = dados.AplicacaoEmissora > 0 ? dados.AplicacaoEmissora - 1 : 0;
                cmbIndRetificacaoFechamento.SelectedIndex = dados.IndRetificacao > 0 ? dados.IndRetificacao - 1 : 0;
                txtNrReciboFechamento.Text = dados.NrRecibo ?? "";
                cmbSitEspecial.SelectedIndex = Array.IndexOf(new[] { 0, 1, 2, 3, 5 }, dados.SitEspecial);
                chkNadaADeclarar.Checked = dados.NadaADeclarar == "1";
                numFechamentoPP.Value = dados.FechamentoPP ?? 0;
                numFechamentoMovOpFin.Value = dados.FechamentoMovOpFin ?? 0;
                numFechamentoMovOpFinAnual.Value = dados.FechamentoMovOpFinAnual ?? 0;
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
            if (chkIndicarMovOpFin != null && !chkIndicarMovOpFin.Checked)
            {
                if (!string.IsNullOrEmpty(txtRMF_CPF.Text) || !string.IsNullOrEmpty(txtRespeFin_CPF.Text))
                {
                    chkIndicarMovOpFin.Checked = true;
                }
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
