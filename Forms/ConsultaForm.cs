using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;
using ExemploAssinadorXML.Models;
using ExemploAssinadorXML.Services;

namespace ExemploAssinadorXML.Forms
{
    public partial class ConsultaForm : Form
    {
        private GroupBox grpConsulta;
        private TextBox txtProtocolo;
        private Button btnConsultar;
        private RichTextBox rtbResultado;

        private GroupBox grpLotes;
        private ListBox lstLotes;
        private Button btnAtualizarLotes;
        private Button btnGerarFechamento;

        private GroupBox grpDetalhes;
        private RichTextBox rtbDetalhes;
        private List<LoteInfo> lotesCarregados;

        public ConfiguracaoForm ConfigForm { get; set; }

        public ConsultaForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Consulta por Protocolo
            grpConsulta = new GroupBox();
            grpConsulta.Text = "Consulta por Protocolo";
            grpConsulta.Location = new Point(10, 10);
            grpConsulta.Size = new Size(750, 150);
            grpConsulta.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            var lblProtocolo = new Label();
            lblProtocolo.Text = "Protocolo:";
            lblProtocolo.Location = new Point(10, 25);
            lblProtocolo.Size = new Size(80, 20);

            txtProtocolo = new TextBox();
            txtProtocolo.Location = new Point(95, 22);
            txtProtocolo.Size = new Size(500, 23);

            btnConsultar = new Button();
            btnConsultar.Text = "Consultar";
            btnConsultar.Location = new Point(605, 21);
            btnConsultar.Size = new Size(100, 25);
            btnConsultar.Click += BtnConsultar_Click;

            rtbResultado = new RichTextBox();
            rtbResultado.Location = new Point(10, 55);
            rtbResultado.Size = new Size(730, 90);
            rtbResultado.ReadOnly = true;
            rtbResultado.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            grpConsulta.Controls.AddRange(new Control[] {
                lblProtocolo, txtProtocolo, btnConsultar, rtbResultado
            });

            // Lista de Lotes
            grpLotes = new GroupBox();
            grpLotes.Text = "Lotes Processados";
            grpLotes.Location = new Point(10, 170);
            grpLotes.Size = new Size(350, 300);
            grpLotes.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;

            lstLotes = new ListBox();
            lstLotes.Location = new Point(10, 25);
            lstLotes.Size = new Size(330, 200);
            lstLotes.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            lstLotes.SelectedIndexChanged += LstLotes_SelectedIndexChanged;
            lstLotes.DoubleClick += LstLotes_DoubleClick;

            btnAtualizarLotes = new Button();
            btnAtualizarLotes.Text = "Atualizar Lista";
            btnAtualizarLotes.Location = new Point(10, 235);
            btnAtualizarLotes.Size = new Size(150, 30);
            btnAtualizarLotes.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnAtualizarLotes.Click += BtnAtualizarLotes_Click;

            btnGerarFechamento = new Button();
            btnGerarFechamento.Text = "Gerar Fechamento";
            btnGerarFechamento.Location = new Point(170, 235);
            btnGerarFechamento.Size = new Size(170, 30);
            btnGerarFechamento.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnGerarFechamento.Click += BtnGerarFechamento_Click;

            grpLotes.Controls.AddRange(new Control[] {
                lstLotes, btnAtualizarLotes, btnGerarFechamento
            });

            // Detalhes
            grpDetalhes = new GroupBox();
            grpDetalhes.Text = "Detalhes do Lote";
            grpDetalhes.Location = new Point(370, 170);
            grpDetalhes.Size = new Size(390, 300);
            grpDetalhes.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            rtbDetalhes = new RichTextBox();
            rtbDetalhes.Location = new Point(10, 25);
            rtbDetalhes.Size = new Size(370, 270);
            rtbDetalhes.ReadOnly = true;
            rtbDetalhes.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            grpDetalhes.Controls.Add(rtbDetalhes);

            this.Controls.AddRange(new Control[] {
                grpConsulta, grpLotes, grpDetalhes
            });

            this.Load += ConsultaForm_Load;

            this.ResumeLayout(false);
        }

        private void ConsultaForm_Load(object sender, EventArgs e)
        {
            // Carregar protocolos automaticamente ao abrir a tela
            BtnAtualizarLotes_Click(null, null);
        }

        /// <summary>
        /// Atualiza a lista de lotes processados (pode ser chamado externamente)
        /// </summary>
        public void AtualizarListaLotes()
        {
            BtnAtualizarLotes_Click(null, null);
        }

        private void LstLotes_DoubleClick(object sender, EventArgs e)
        {
            if (lstLotes.SelectedIndex < 0) return;

            try
            {
                // Tentar extrair protocolo do item selecionado
                if (lotesCarregados != null && lstLotes.SelectedIndex < lotesCarregados.Count)
                {
                    var lote = lotesCarregados[lstLotes.SelectedIndex];
                    if (!string.IsNullOrEmpty(lote.Protocolo))
                    {
                        txtProtocolo.Text = lote.Protocolo;
                        BtnConsultar_Click(null, null);
                    }
                    else
                    {
                        MessageBox.Show("Este lote não possui protocolo registrado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao consultar protocolo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnConsultar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtProtocolo.Text))
            {
                MessageBox.Show("Informe o protocolo para consulta.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (ConfigForm == null || ConfigForm.Config == null)
            {
                MessageBox.Show("Configure as opções primeiro na aba Configuração.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var consultaService = new EfinanceiraConsultaService();
                var config = ConfigForm.Config;
                var cert = BuscarCertificado(config.CertThumbprint);
                var resposta = consultaService.ConsultarProtocolo(txtProtocolo.Text, config, cert);

                rtbResultado.Clear();
                rtbResultado.AppendText("═══════════════════════════════════════════════════════════\n");
                rtbResultado.AppendText("                    RESULTADO DA CONSULTA\n");
                rtbResultado.AppendText("═══════════════════════════════════════════════════════════\n\n");
                
                rtbResultado.AppendText($"📋 Protocolo: {txtProtocolo.Text}\n");
                rtbResultado.AppendText($"🌐 Código HTTP: {resposta.CodigoHttp}\n");
                rtbResultado.AppendText($"📊 Código Resposta: {resposta.CodigoResposta}\n");
                rtbResultado.AppendText($"📝 Descrição: {resposta.Descricao}\n\n");

                // Informações adicionais do lote
                if (!string.IsNullOrEmpty(resposta.ProtocoloEnvio))
                {
                    rtbResultado.AppendText($"📤 Protocolo de Envio: {resposta.ProtocoloEnvio}\n");
                }
                if (resposta.DataRecepcao.HasValue)
                {
                    rtbResultado.AppendText($"📥 Data/Hora Recepção: {resposta.DataRecepcao.Value:dd/MM/yyyy HH:mm:ss}\n");
                }
                if (resposta.DataProcessamento.HasValue)
                {
                    rtbResultado.AppendText($"⚙️  Data/Hora Processamento: {resposta.DataProcessamento.Value:dd/MM/yyyy HH:mm:ss}\n");
                }
                if (!string.IsNullOrEmpty(resposta.VersaoAplicativoRecepcao))
                {
                    rtbResultado.AppendText($"🔢 Versão Aplicativo Recepção: {resposta.VersaoAplicativoRecepcao}\n");
                }
                if (!string.IsNullOrEmpty(resposta.VersaoAplicativoProcessamento))
                {
                    rtbResultado.AppendText($"🔢 Versão Aplicativo Processamento: {resposta.VersaoAplicativoProcessamento}\n");
                }

                rtbResultado.AppendText("\n");
                rtbResultado.AppendText("═══════════════════════════════════════════════════════════\n");
                rtbResultado.AppendText("                         STATUS DO LOTE\n");
                rtbResultado.AppendText("═══════════════════════════════════════════════════════════\n\n");

                // Interpretar códigos de resposta baseado no Java de referência
                if (resposta.CodigoResposta == 1)
                {
                    rtbResultado.AppendText("⏳ Status: Lote ainda está em processamento.\n");
                }
                else if (resposta.CodigoResposta == 2)
                {
                    rtbResultado.AppendText("✅ Status: Lote processado com sucesso! Todos os eventos foram processados.\n");
                }
                else if (resposta.CodigoResposta == 3)
                {
                    rtbResultado.AppendText("⚠️  Status: Lote processado, mas possui um ou mais eventos com ocorrências de erro.\n");
                }
                else if (resposta.CodigoResposta == 4)
                {
                    rtbResultado.AppendText("❌ Status: A consulta possui ocorrências. Verifique os parâmetros informados.\n");
                }
                else if (resposta.CodigoResposta == 5)
                {
                    rtbResultado.AppendText("❌ Status: Lote não encontrado com o protocolo informado.\n");
                }
                else if (resposta.CodigoResposta == 9)
                {
                    rtbResultado.AppendText("❌ Status: Erro interno na e-Financeira.\n");
                }
                else
                {
                    rtbResultado.AppendText("❓ Status: Resposta inesperada do servidor.\n");
                }

                // Ocorrências gerais do lote
                if (resposta.Ocorrencias != null && resposta.Ocorrencias.Count > 0)
                {
                    rtbResultado.AppendText("\n");
                    rtbResultado.AppendText("═══════════════════════════════════════════════════════════\n");
                    rtbResultado.AppendText("              OCORRÊNCIAS GERAIS DO LOTE\n");
                    rtbResultado.AppendText("═══════════════════════════════════════════════════════════\n\n");
                    
                    for (int i = 0; i < resposta.Ocorrencias.Count; i++)
                    {
                        var ocorrencia = resposta.Ocorrencias[i];
                        rtbResultado.AppendText($"🔴 OCORRÊNCIA {i + 1}:\n");
                        if (!string.IsNullOrEmpty(ocorrencia.Codigo))
                        {
                            rtbResultado.AppendText($"   Código: {ocorrencia.Codigo}\n");
                        }
                        if (!string.IsNullOrEmpty(ocorrencia.Descricao))
                        {
                            rtbResultado.AppendText($"   Descrição: {ocorrencia.Descricao}\n");
                        }
                        if (!string.IsNullOrEmpty(ocorrencia.Tipo))
                        {
                            rtbResultado.AppendText($"   Tipo: {ocorrencia.Tipo}\n");
                        }
                        rtbResultado.AppendText("\n");
                    }
                }

                // Detalhes de cada evento individual
                if (resposta.DetalhesEventos != null && resposta.DetalhesEventos.Count > 0)
                {
                    rtbResultado.AppendText("\n");
                    rtbResultado.AppendText("═══════════════════════════════════════════════════════════\n");
                    rtbResultado.AppendText("                    DETALHES DOS EVENTOS\n");
                    rtbResultado.AppendText("═══════════════════════════════════════════════════════════\n\n");
                    
                    for (int i = 0; i < resposta.DetalhesEventos.Count; i++)
                    {
                        var evento = resposta.DetalhesEventos[i];
                        rtbResultado.AppendText($"📦 EVENTO {i + 1}:\n");
                        rtbResultado.AppendText($"   ID: {evento.IdEvento ?? "N/A"}\n");
                        
                        if (!string.IsNullOrEmpty(evento.TipoEvento))
                        {
                            string tipoEventoDesc = ObterDescricaoTipoEvento(evento.TipoEvento);
                            rtbResultado.AppendText($"   Tipo: {evento.TipoEvento} - {tipoEventoDesc}\n");
                        }
                        
                        if (!string.IsNullOrEmpty(evento.CodigoRetorno))
                        {
                            rtbResultado.AppendText($"   Código Retorno: {evento.CodigoRetorno}\n");
                        }
                        
                        if (!string.IsNullOrEmpty(evento.DescricaoRetorno))
                        {
                            string statusIcon = evento.DescricaoRetorno.ToUpper().Contains("ERRO") ? "❌" : "✅";
                            rtbResultado.AppendText($"   Status: {statusIcon} {evento.DescricaoRetorno}\n");
                        }
                        
                        if (evento.DataRecepcao.HasValue)
                        {
                            rtbResultado.AppendText($"   Data Recepção: {evento.DataRecepcao.Value:dd/MM/yyyy HH:mm:ss}\n");
                        }
                        
                        if (evento.DataProcessamento.HasValue)
                        {
                            rtbResultado.AppendText($"   Data Processamento: {evento.DataProcessamento.Value:dd/MM/yyyy HH:mm:ss}\n");
                        }
                        
                        if (!string.IsNullOrEmpty(evento.Hash))
                        {
                            rtbResultado.AppendText($"   Hash: {evento.Hash}\n");
                        }
                        
                        // Ocorrências do evento
                        if (evento.Ocorrencias != null && evento.Ocorrencias.Count > 0)
                        {
                            rtbResultado.AppendText($"\n   ⚠️  ERROS ENCONTRADOS NESTE EVENTO:\n");
                            for (int j = 0; j < evento.Ocorrencias.Count; j++)
                            {
                                var ocorrencia = evento.Ocorrencias[j];
                                rtbResultado.AppendText($"\n      🔴 Erro {j + 1}:\n");
                                if (!string.IsNullOrEmpty(ocorrencia.Codigo))
                                {
                                    rtbResultado.AppendText($"         Código: {ocorrencia.Codigo}\n");
                                }
                                if (!string.IsNullOrEmpty(ocorrencia.Descricao))
                                {
                                    rtbResultado.AppendText($"         Descrição: {ocorrencia.Descricao}\n");
                                }
                                if (!string.IsNullOrEmpty(ocorrencia.Tipo))
                                {
                                    rtbResultado.AppendText($"         Tipo: {ocorrencia.Tipo}\n");
                                }
                            }
                        }
                        else if (!string.IsNullOrEmpty(evento.DescricaoRetorno) && 
                                 evento.DescricaoRetorno.ToUpper().Contains("SUCESSO"))
                        {
                            rtbResultado.AppendText($"\n   ✅ Nenhum erro encontrado neste evento.\n");
                        }
                        
                        rtbResultado.AppendText("\n");
                        rtbResultado.AppendText("───────────────────────────────────────────────────────────\n\n");
                    }
                }

                // Exibir XML completo apenas se solicitado (opcional - pode ser removido ou colocado em botão separado)
                // Removido para tornar a exibição mais limpa e focada nas informações importantes
            }
            catch (Exception ex)
            {
                rtbResultado.Clear();
                rtbResultado.AppendText($"ERRO: {ex.Message}");
                MessageBox.Show($"Erro ao consultar protocolo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAtualizarLotes_Click(object sender, EventArgs e)
        {
            try
            {
                lstLotes.Items.Clear();
                
                // Carregar lotes com protocolos registrados
                var lotes = ProtocoloPersistenciaService.CarregarLotes();
                
                // Inicializar lista (mesmo se vazia)
                lotesCarregados = (lotes != null && lotes.Count > 0) 
                    ? lotes.OrderByDescending(l => l.DataProcessamento).ToList()
                    : new List<LoteInfo>();
                
                if (lotesCarregados.Count > 0)
                {
                    foreach (var lote in lotesCarregados)
                    {
                        string tipoStr = lote.Tipo.ToString();
                        string protocoloStr = !string.IsNullOrEmpty(lote.Protocolo) 
                            ? $"Protocolo: {lote.Protocolo}" 
                            : "Sem protocolo";
                        string periodoStr = !string.IsNullOrEmpty(lote.Periodo) 
                            ? $"Período: {lote.Periodo}" 
                            : "";
                        string dataStr = lote.DataProcessamento.ToString("dd/MM/yyyy HH:mm:ss");
                        
                        string item = $"[{tipoStr}] {protocoloStr} | {periodoStr} | {dataStr}";
                        lstLotes.Items.Add(item);
                    }
                }
                
                // Também carregar arquivos do diretório (para compatibilidade)
                if (ConfigForm != null && ConfigForm.Config != null)
                {
                    string diretorio = ConfigForm.Config.DiretorioLotes;
                    if (Directory.Exists(diretorio))
                    {
                        var arquivos = Directory.GetFiles(diretorio, "*-Criptografado.xml", SearchOption.AllDirectories);
                        foreach (var arquivo in arquivos)
                        {
                            // Verificar se já está na lista
                            string nomeArquivo = Path.GetFileName(arquivo);
                            bool jaExiste = lotes.Any(l => 
                                !string.IsNullOrEmpty(l.ArquivoCriptografado) && 
                                Path.GetFileName(l.ArquivoCriptografado) == nomeArquivo);
                            
                            if (!jaExiste)
                            {
                                var info = new FileInfo(arquivo);
                                lstLotes.Items.Add($"[Arquivo] {info.Name} - {info.LastWriteTime:dd/MM/yyyy HH:mm:ss}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao atualizar lista: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LstLotes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstLotes.SelectedIndex < 0) return;

            try
            {
                string itemSelecionado = lstLotes.SelectedItem.ToString();
                rtbDetalhes.Clear();
                
                // Verificar se é um lote registrado (tem protocolo)
                if (lotesCarregados != null && lstLotes.SelectedIndex < lotesCarregados.Count)
                {
                    var lote = lotesCarregados[lstLotes.SelectedIndex];
                    
                    // Tipo do lote
                    string tipoStr = lote.Tipo.ToString();
                    rtbDetalhes.AppendText($"════════════════════════════════════════\n");
                    rtbDetalhes.AppendText($"TIPO DE LOTE: {tipoStr}\n");
                    rtbDetalhes.AppendText($"════════════════════════════════════════\n\n");
                    
                    // Quantidade de eventos
                    rtbDetalhes.AppendText($"Quantidade de Eventos: {lote.QuantidadeEventos}\n");
                    
                    // Período
                    string periodoStr = !string.IsNullOrEmpty(lote.Periodo) ? lote.Periodo : "Não informado";
                    rtbDetalhes.AppendText($"Período: {periodoStr}\n");
                    
                    // Protocolo
                    string protocoloStr = !string.IsNullOrEmpty(lote.Protocolo) ? lote.Protocolo : "Não disponível";
                    rtbDetalhes.AppendText($"Protocolo: {protocoloStr}\n");
                    
                    // Status
                    rtbDetalhes.AppendText($"Status: {lote.Status}\n");
                    
                    // Data Processamento
                    rtbDetalhes.AppendText($"Data Processamento: {lote.DataProcessamento:dd/MM/yyyy HH:mm:ss}\n\n");
                    
                    // Arquivos
                    rtbDetalhes.AppendText($"════════════════════════════════════════\n");
                    rtbDetalhes.AppendText($"ARQUIVOS:\n");
                    rtbDetalhes.AppendText($"════════════════════════════════════════\n");
                    
                    if (!string.IsNullOrEmpty(lote.ArquivoOriginal))
                    {
                        rtbDetalhes.AppendText($"Original: {Path.GetFileName(lote.ArquivoOriginal)}\n");
                    }
                    if (!string.IsNullOrEmpty(lote.ArquivoAssinado))
                    {
                        rtbDetalhes.AppendText($"Assinado: {Path.GetFileName(lote.ArquivoAssinado)}\n");
                    }
                    if (!string.IsNullOrEmpty(lote.ArquivoCriptografado))
                    {
                        rtbDetalhes.AppendText($"Criptografado: {Path.GetFileName(lote.ArquivoCriptografado)}\n");
                        
                        // Tentar obter informações do arquivo
                        if (File.Exists(lote.ArquivoCriptografado))
                        {
                            var info = new FileInfo(lote.ArquivoCriptografado);
                            rtbDetalhes.AppendText($"Tamanho: {info.Length / 1024.0:F2} KB\n");
                        }
                    }
                    
                    // Preencher campo de protocolo automaticamente
                    if (!string.IsNullOrEmpty(lote.Protocolo))
                    {
                        txtProtocolo.Text = lote.Protocolo;
                        rtbDetalhes.AppendText($"\n[Protocolo preenchido automaticamente - clique em 'Consultar' para verificar status]\n");
                    }
                    else
                    {
                        txtProtocolo.Text = "";
                        rtbDetalhes.AppendText($"\n[Este lote ainda não possui protocolo - não foi enviado ou aguardando resposta]\n");
                    }
                }
                else
                {
                    // É um arquivo do diretório (sem registro)
                    if (ConfigForm != null && ConfigForm.Config != null)
                    {
                        string diretorio = ConfigForm.Config.DiretorioLotes;
                        string nomeArquivo = itemSelecionado.Contains("-Criptografado.xml") 
                            ? itemSelecionado.Substring(itemSelecionado.IndexOf("]") + 2).Split('-')[0] + "-Criptografado.xml"
                            : itemSelecionado.Split('-')[0] + "-Criptografado.xml";
                        string caminhoCompleto = Path.Combine(diretorio, nomeArquivo);

                        if (File.Exists(caminhoCompleto))
                        {
                            var info = new FileInfo(caminhoCompleto);
                            rtbDetalhes.AppendText($"Arquivo: {info.Name}\n");
                            rtbDetalhes.AppendText($"Tamanho: {info.Length / 1024.0:F2} KB\n");
                            rtbDetalhes.AppendText($"Data: {info.LastWriteTime:dd/MM/yyyy HH:mm:ss}\n");
                            rtbDetalhes.AppendText($"\n[Este arquivo não possui protocolo registrado]\n");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                rtbDetalhes.Clear();
                rtbDetalhes.AppendText($"Erro ao carregar detalhes: {ex.Message}");
            }
        }

        private void BtnGerarFechamento_Click(object sender, EventArgs e)
        {
            if (ConfigForm == null || ConfigForm.DadosFechamento == null)
            {
                MessageBox.Show("Configure os dados de fechamento primeiro na aba Configuração.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var form = new GerarFechamentoForm(ConfigForm);
            form.ShowDialog();
        }

        private string ObterDescricaoTipoEvento(string tipoEvento)
        {
            switch (tipoEvento)
            {
                case "001": return "Cadastro de Declarante";
                case "002": return "Abertura e-Financeira";
                case "003": return "Cadastro de Intermediário";
                case "004": return "Cadastro de Patrocinado";
                case "005": return "Exclusão e-Financeira";
                case "006": return "Exclusão";
                case "007": return "Fechamento e-Financeira";
                case "008": return "Movimentação de Operação Financeira";
                case "009": return "Movimentação de Previdência Privada";
                default: return "Tipo desconhecido";
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
    }

    public class GerarFechamentoForm : Form
    {
        private Label lblPeriodo;
        private TextBox txtPeriodo;
        private Button btnCalcularPeriodo;
        private Label lblDataInicio;
        private Label lblDataFim;
        private TextBox txtDataInicio;
        private TextBox txtDataFim;
        private Label lblInfoPeriodo;
        private Button btnGerar;
        private Button btnCancelar;
        private ProgressBar progressBar;
        private Label lblStatus;

        public ConfiguracaoForm ConfigForm { get; set; }

        public GerarFechamentoForm(ConfiguracaoForm configForm)
        {
            ConfigForm = configForm;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Gerar Fechamento por Período";
            this.Size = new Size(500, 280);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            int yPos = 15;
            int xInicio = 15;

            // Período
            lblPeriodo = new Label();
            lblPeriodo.Text = "Período (YYYYMM):";
            lblPeriodo.Location = new Point(xInicio, yPos);
            lblPeriodo.Size = new Size(120, 20);

            txtPeriodo = new TextBox();
            txtPeriodo.Location = new Point(xInicio + 125, yPos - 3);
            txtPeriodo.Size = new Size(120, 23);
            txtPeriodo.Text = EfinanceiraPeriodoUtil.CalcularPeriodoAtual();
            txtPeriodo.Leave += TxtPeriodo_Leave;

            btnCalcularPeriodo = new Button();
            btnCalcularPeriodo.Text = "Calcular Período Atual";
            btnCalcularPeriodo.Location = new Point(xInicio + 255, yPos - 3);
            btnCalcularPeriodo.Size = new Size(150, 25);
            btnCalcularPeriodo.Click += BtnCalcularPeriodo_Click;

            yPos += 35;

            // Datas calculadas
            lblDataInicio = new Label();
            lblDataInicio.Text = "Data Início:";
            lblDataInicio.Location = new Point(xInicio, yPos);
            lblDataInicio.Size = new Size(100, 20);

            txtDataInicio = new TextBox();
            txtDataInicio.Location = new Point(xInicio + 105, yPos - 3);
            txtDataInicio.Size = new Size(120, 23);
            txtDataInicio.ReadOnly = true;
            txtDataInicio.BackColor = Color.WhiteSmoke;

            lblDataFim = new Label();
            lblDataFim.Text = "Data Fim:";
            lblDataFim.Location = new Point(xInicio + 240, yPos);
            lblDataFim.Size = new Size(80, 20);

            txtDataFim = new TextBox();
            txtDataFim.Location = new Point(xInicio + 325, yPos - 3);
            txtDataFim.Size = new Size(120, 23);
            txtDataFim.ReadOnly = true;
            txtDataFim.BackColor = Color.WhiteSmoke;

            yPos += 35;

            // Info
            lblInfoPeriodo = new Label();
            lblInfoPeriodo.Text = "Info: Período 06 = Jan-Jun | Período 12 = Jul-Dez";
            lblInfoPeriodo.Location = new Point(xInicio, yPos);
            lblInfoPeriodo.Size = new Size(450, 20);
            lblInfoPeriodo.ForeColor = Color.Gray;
            lblInfoPeriodo.Font = new Font(lblInfoPeriodo.Font, FontStyle.Italic);

            yPos += 35;

            // Progress
            progressBar = new ProgressBar();
            progressBar.Location = new Point(xInicio, yPos);
            progressBar.Size = new Size(450, 23);
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.Visible = false;

            yPos += 30;

            lblStatus = new Label();
            lblStatus.Text = "";
            lblStatus.Location = new Point(xInicio, yPos);
            lblStatus.Size = new Size(450, 20);
            lblStatus.ForeColor = Color.Blue;

            yPos += 30;

            // Botões
            btnGerar = new Button();
            btnGerar.Text = "Gerar Fechamento";
            btnGerar.Location = new Point(xInicio + 200, yPos);
            btnGerar.Size = new Size(130, 35);
            btnGerar.Click += BtnGerar_Click;

            btnCancelar = new Button();
            btnCancelar.Text = "Cancelar";
            btnCancelar.Location = new Point(xInicio + 340, yPos);
            btnCancelar.Size = new Size(110, 35);
            btnCancelar.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[] {
                lblPeriodo, txtPeriodo, btnCalcularPeriodo,
                lblDataInicio, txtDataInicio, lblDataFim, txtDataFim,
                lblInfoPeriodo, progressBar, lblStatus, btnGerar, btnCancelar
            });

            // Calcular período inicial
            TxtPeriodo_Leave(null, null);
        }

        private void TxtPeriodo_Leave(object sender, EventArgs e)
        {
            try
            {
                string periodo = txtPeriodo.Text.Trim();
                if (EfinanceiraPeriodoUtil.ValidarPeriodo(periodo))
                {
                    var (dataInicio, dataFim) = EfinanceiraPeriodoUtil.CalcularPeriodoSemestral(periodo);
                    txtDataInicio.Text = dataInicio;
                    txtDataFim.Text = dataFim;
                    lblStatus.Text = $"Período válido: {dataInicio} até {dataFim}";
                    lblStatus.ForeColor = Color.Green;
                }
                else
                {
                    txtDataInicio.Text = "";
                    txtDataFim.Text = "";
                    lblStatus.Text = "Período inválido. Use formato YYYYMM (ex: 202312 ou 202406)";
                    lblStatus.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                txtDataInicio.Text = "";
                txtDataFim.Text = "";
                lblStatus.Text = $"Erro: {ex.Message}";
                lblStatus.ForeColor = Color.Red;
            }
        }

        private void BtnCalcularPeriodo_Click(object sender, EventArgs e)
        {
            txtPeriodo.Text = EfinanceiraPeriodoUtil.CalcularPeriodoAtual();
            TxtPeriodo_Leave(null, null);
        }

        private async void BtnGerar_Click(object sender, EventArgs e)
        {
            try
            {
                // Validar período
                string periodo = txtPeriodo.Text.Trim();
                if (!EfinanceiraPeriodoUtil.ValidarPeriodo(periodo))
                {
                    MessageBox.Show("Período inválido. Use formato YYYYMM (ex: 202312 para Jul-Dez ou 202406 para Jan-Jun).", 
                        "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (ConfigForm == null || ConfigForm.Config == null)
                {
                    MessageBox.Show("Configure as opções primeiro na aba Configuração.", "Aviso", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (ConfigForm.DadosFechamento == null)
                {
                    MessageBox.Show("Configure os dados de fechamento primeiro na aba Configuração.", "Aviso", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Calcular datas
                var (dataInicio, dataFim) = EfinanceiraPeriodoUtil.CalcularPeriodoSemestral(periodo);

                // Desabilitar controles
                btnGerar.Enabled = false;
                btnCancelar.Enabled = false;
                progressBar.Visible = true;
                lblStatus.Text = "Gerando fechamento...";
                lblStatus.ForeColor = Color.Blue;
                Application.DoEvents();

                // Criar dados de fechamento baseado no período
                var dadosFechamento = new DadosFechamento
                {
                    CnpjDeclarante = ConfigForm.Config.CnpjDeclarante,
                    DtInicio = dataInicio,
                    DtFim = dataFim,
                    TipoAmbiente = ConfigForm.DadosFechamento.TipoAmbiente,
                    AplicacaoEmissora = ConfigForm.DadosFechamento.AplicacaoEmissora,
                    IndRetificacao = ConfigForm.DadosFechamento.IndRetificacao,
                    NrRecibo = ConfigForm.DadosFechamento.NrRecibo,
                    SitEspecial = ConfigForm.DadosFechamento.SitEspecial,
                    NadaADeclarar = ConfigForm.DadosFechamento.NadaADeclarar,
                    FechamentoPP = ConfigForm.DadosFechamento.FechamentoPP,
                    FechamentoMovOpFin = ConfigForm.DadosFechamento.FechamentoMovOpFin,
                    FechamentoMovOpFinAnual = ConfigForm.DadosFechamento.FechamentoMovOpFinAnual,
                    ContasAReportarEntDecExterior = ConfigForm.DadosFechamento.ContasAReportarEntDecExterior,
                    EntidadesPatrocinadas = ConfigForm.DadosFechamento.EntidadesPatrocinadas
                };

                // Gerar XML
                lblStatus.Text = "Gerando XML de fechamento...";
                Application.DoEvents();

                var geradorService = new EfinanceiraGeradorXmlService();
                string arquivoXml = geradorService.GerarXmlFechamento(dadosFechamento, ConfigForm.Config.DiretorioLotes);

                lblStatus.Text = $"XML gerado: {Path.GetFileName(arquivoXml)}";
                lblStatus.ForeColor = Color.Green;

                MessageBox.Show($"Fechamento gerado com sucesso!\n\nArquivo: {arquivoXml}\n\n" +
                    $"Período: {dataInicio} até {dataFim}\n" +
                    $"Próximo passo: Processar o arquivo na aba Processamento.",
                    "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.Close();
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"Erro: {ex.Message}";
                lblStatus.ForeColor = Color.Red;
                MessageBox.Show($"Erro ao gerar fechamento: {ex.Message}", "Erro", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Visible = false;
                btnGerar.Enabled = true;
                btnCancelar.Enabled = true;
            }
        }
    }
}
