using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
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
        private DateTimePicker dtpFiltroDataInicio;
        private DateTimePicker dtpFiltroDataFim;
        private Label lblFiltroDataInicio;
        private Label lblFiltroDataFim;
        private CheckBox chkUsarFiltroData;
        private CheckBox chkUsarFiltroPeriodo;
        private ComboBox cmbFiltroPeriodo;
        private Label lblFiltroPeriodo;
        private ComboBox cmbFiltroAmbiente;
        private Label lblFiltroAmbiente;

        private GroupBox grpDetalhes;
        private RichTextBox rtbDetalhes;
        private List<LoteInfo> lotesCarregados;
        private List<LoteBancoInfo> lotesBancoCarregados;

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
            grpLotes.Size = new Size(350, 420);
            grpLotes.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;

            int yPos = 25;

            // Filtro de data
            chkUsarFiltroData = new CheckBox();
            chkUsarFiltroData.Text = "Filtrar por data";
            chkUsarFiltroData.Location = new Point(10, yPos);
            chkUsarFiltroData.Size = new Size(120, 20);
            chkUsarFiltroData.Checked = false; // Por padrão, não filtrar por data

            yPos += 25;

            lblFiltroDataInicio = new Label();
            lblFiltroDataInicio.Text = "Data Início:";
            lblFiltroDataInicio.Location = new Point(10, yPos);
            lblFiltroDataInicio.Size = new Size(70, 20);
            lblFiltroDataInicio.Enabled = chkUsarFiltroData.Checked;

            dtpFiltroDataInicio = new DateTimePicker();
            dtpFiltroDataInicio.Location = new Point(85, yPos - 3);
            dtpFiltroDataInicio.Size = new Size(110, 23);
            dtpFiltroDataInicio.Format = DateTimePickerFormat.Short;
            dtpFiltroDataInicio.Value = DateTime.Today;
            dtpFiltroDataInicio.Enabled = chkUsarFiltroData.Checked;

            lblFiltroDataFim = new Label();
            lblFiltroDataFim.Text = "Data Fim:";
            lblFiltroDataFim.Location = new Point(200, yPos);
            lblFiltroDataFim.Size = new Size(60, 20);
            lblFiltroDataFim.Enabled = chkUsarFiltroData.Checked;

            dtpFiltroDataFim = new DateTimePicker();
            dtpFiltroDataFim.Location = new Point(265, yPos - 3);
            dtpFiltroDataFim.Size = new Size(110, 23);
            dtpFiltroDataFim.Format = DateTimePickerFormat.Short;
            dtpFiltroDataFim.Value = DateTime.Today;
            dtpFiltroDataFim.Enabled = chkUsarFiltroData.Checked;

            chkUsarFiltroData.CheckedChanged += (s, e) => 
            {
                bool enabled = chkUsarFiltroData.Checked;
                dtpFiltroDataInicio.Enabled = enabled;
                dtpFiltroDataFim.Enabled = enabled;
                lblFiltroDataInicio.Enabled = enabled;
                lblFiltroDataFim.Enabled = enabled;
            };

            yPos += 30;

            // Filtro de período
            chkUsarFiltroPeriodo = new CheckBox();
            chkUsarFiltroPeriodo.Text = "Filtrar por período";
            chkUsarFiltroPeriodo.Location = new Point(10, yPos);
            chkUsarFiltroPeriodo.Size = new Size(130, 20);
            chkUsarFiltroPeriodo.Checked = false;

            yPos += 25;

            lblFiltroPeriodo = new Label();
            lblFiltroPeriodo.Text = "Período:";
            lblFiltroPeriodo.Location = new Point(10, yPos);
            lblFiltroPeriodo.Size = new Size(70, 20);
            lblFiltroPeriodo.Enabled = chkUsarFiltroPeriodo.Checked;

            cmbFiltroPeriodo = new ComboBox();
            cmbFiltroPeriodo.Location = new Point(85, yPos - 3);
            cmbFiltroPeriodo.Size = new Size(200, 23);
            cmbFiltroPeriodo.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbFiltroPeriodo.Enabled = chkUsarFiltroPeriodo.Checked;
            PopularComboPeriodos();

            chkUsarFiltroPeriodo.CheckedChanged += (s, e) => 
            {
                bool enabled = chkUsarFiltroPeriodo.Checked;
                cmbFiltroPeriodo.Enabled = enabled;
                lblFiltroPeriodo.Enabled = enabled;
            };

            yPos += 30;

            // Filtro de ambiente
            lblFiltroAmbiente = new Label();
            lblFiltroAmbiente.Text = "Ambiente:";
            lblFiltroAmbiente.Location = new Point(10, yPos);
            lblFiltroAmbiente.Size = new Size(70, 20);

            cmbFiltroAmbiente = new ComboBox();
            cmbFiltroAmbiente.Location = new Point(85, yPos - 3);
            cmbFiltroAmbiente.Size = new Size(150, 23);
            cmbFiltroAmbiente.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbFiltroAmbiente.Items.AddRange(new object[] { "Ambos", "Teste", "Produção" });
            cmbFiltroAmbiente.SelectedIndex = 0; // Ambos por padrão

            yPos += 30;

            Button btnFiltrar = new Button();
            btnFiltrar.Text = "Filtrar";
            btnFiltrar.Location = new Point(10, yPos);
            btnFiltrar.Size = new Size(100, 25);
            btnFiltrar.Click += BtnFiltrar_Click;

            yPos += 35;

            lstLotes = new ListBox();
            lstLotes.Location = new Point(10, yPos);
            lstLotes.Size = new Size(330, 200);
            lstLotes.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            lstLotes.SelectedIndexChanged += LstLotes_SelectedIndexChanged;
            lstLotes.DoubleClick += LstLotes_DoubleClick;

            btnAtualizarLotes = new Button();
            btnAtualizarLotes.Text = "Atualizar Lista";
            btnAtualizarLotes.Location = new Point(10, yPos + 210);
            btnAtualizarLotes.Size = new Size(150, 30);
            btnAtualizarLotes.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnAtualizarLotes.Click += BtnAtualizarLotes_Click;

            btnGerarFechamento = new Button();
            btnGerarFechamento.Text = "Gerar Fechamento";
            btnGerarFechamento.Location = new Point(170, yPos + 210);
            btnGerarFechamento.Size = new Size(170, 30);
            btnGerarFechamento.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnGerarFechamento.Click += BtnGerarFechamento_Click;

            grpLotes.Controls.AddRange(new Control[] {
                chkUsarFiltroData, lblFiltroDataInicio, dtpFiltroDataInicio, lblFiltroDataFim, dtpFiltroDataFim,
                chkUsarFiltroPeriodo, lblFiltroPeriodo, cmbFiltroPeriodo,
                lblFiltroAmbiente, cmbFiltroAmbiente, btnFiltrar,
                lstLotes, btnAtualizarLotes, btnGerarFechamento
            });

            // Detalhes
            grpDetalhes = new GroupBox();
            grpDetalhes.Text = "Detalhes do Lote";
            grpDetalhes.Location = new Point(370, 170);
            grpDetalhes.Size = new Size(390, 420);
            grpDetalhes.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            rtbDetalhes = new RichTextBox();
            rtbDetalhes.Location = new Point(10, 25);
            rtbDetalhes.Size = new Size(370, 390);
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
                // Priorizar lotes do banco
                if (lotesBancoCarregados != null && lstLotes.SelectedIndex < lotesBancoCarregados.Count)
                {
                    var lote = lotesBancoCarregados[lstLotes.SelectedIndex];
                    if (!string.IsNullOrEmpty(lote.ProtocoloEnvio))
                    {
                        txtProtocolo.Text = lote.ProtocoloEnvio;
                        BtnConsultar_Click(null, null);
                    }
                    else
                    {
                        MessageBox.Show("Este lote não possui protocolo registrado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                // Fallback para sistema antigo
                else if (lotesCarregados != null && lstLotes.SelectedIndex < lotesCarregados.Count)
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
                // Buscar período do lote no banco de dados
                string periodoLote = null;
                try
                {
                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                    var loteBanco = persistenceService.BuscarLotePorProtocolo(txtProtocolo.Text);
                    if (loteBanco != null && !string.IsNullOrEmpty(loteBanco.Periodo))
                    {
                        periodoLote = loteBanco.Periodo;
                    }
                }
                catch (Exception exPeriodo)
                {
                    // Não interromper a consulta se houver erro ao buscar período
                    System.Diagnostics.Debug.WriteLine($"Erro ao buscar período do lote: {exPeriodo.Message}");
                }

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
                rtbResultado.AppendText($"📝 Descrição: {resposta.Descricao}\n");
                
                // Exibir período do lote
                if (!string.IsNullOrEmpty(periodoLote))
                {
                    // Formatar período de forma mais amigável
                    string periodoFormatado = periodoLote;
                    if (periodoLote.Length == 6)
                    {
                        int ano = int.Parse(periodoLote.Substring(0, 4));
                        int mes = int.Parse(periodoLote.Substring(4, 2));
                        string semestre = "";
                        if (mes == 1 || mes == 6)
                        {
                            semestre = "Jan-Jun";
                        }
                        else if (mes == 2 || mes == 12)
                        {
                            semestre = "Jul-Dez";
                        }
                        periodoFormatado = $"{periodoLote} ({semestre}/{ano})";
                    }
                    rtbResultado.AppendText($"📅 Período do Lote: {periodoFormatado}\n");
                }
                
                rtbResultado.AppendText("\n");

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
                                
                                // Adicionar orientação de tratamento do erro
                                string orientacao = ObterOrientacaoTratamentoErro(ocorrencia.Codigo, ocorrencia.Descricao);
                                if (!string.IsNullOrEmpty(orientacao))
                                {
                                    rtbResultado.AppendText($"\n         💡 COMO TRATAR ESTE ERRO:\n");
                                    rtbResultado.AppendText($"         {orientacao}\n");
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

                // Registrar consulta no banco de dados
                try
                {
                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                    var loteBanco = persistenceService.BuscarLotePorProtocolo(txtProtocolo.Text);
                    
                    if (loteBanco != null)
                    {
                        // Serializar XML de resposta para JSON (simplificado)
                        string xmlRespostaJson = JsonSerializer.Serialize(resposta);
                        
                        // Atualizar lote com resultado da consulta
                        persistenceService.AtualizarLote(
                            loteBanco.IdLote,
                            resposta.CodigoResposta == 2 ? "CONSULTADO_SUCESSO" : 
                            resposta.CodigoResposta == 3 ? "CONSULTADO_COM_ERRO" : 
                            resposta.CodigoResposta == 1 ? "EM_PROCESSAMENTO" : "CONSULTADO",
                            null,
                            null,
                            null,
                            null,
                            resposta.CodigoResposta,
                            resposta.Descricao,
                            xmlRespostaJson,
                            null,
                            resposta.DataProcessamento ?? resposta.DataRecepcao,
                            null
                        );
                        
                        persistenceService.RegistrarLogLote(loteBanco.IdLote, "CONSULTA", 
                            $"Consulta realizada. Código: {resposta.CodigoResposta}, Descrição: {resposta.Descricao}");
                        
                        // Atualizar status dos eventos
                        if (resposta.DetalhesEventos != null && resposta.DetalhesEventos.Count > 0)
                        {
                            var eventosBanco = persistenceService.BuscarEventosDoLote(loteBanco.IdLote);
                            
                            foreach (var eventoConsulta in resposta.DetalhesEventos)
                            {
                                // Tentar encontrar evento correspondente pelo ID
                                var eventoBanco = eventosBanco.FirstOrDefault(evt => 
                                    !string.IsNullOrEmpty(evt.IdEventoXml) && 
                                    evt.IdEventoXml.Contains(eventoConsulta.IdEvento ?? ""));
                                
                                if (eventoBanco != null)
                                {
                                    string statusEvento = eventoConsulta.DescricaoRetorno ?? "CONSULTADO";
                                    string ocorrenciasJson = null;
                                    
                                    if (eventoConsulta.Ocorrencias != null && eventoConsulta.Ocorrencias.Count > 0)
                                    {
                                        ocorrenciasJson = JsonSerializer.Serialize(eventoConsulta.Ocorrencias);
                                    }
                                    
                                    persistenceService.AtualizarEvento(
                                        eventoBanco.IdEvento,
                                        statusEvento,
                                        ocorrenciasJson,
                                        null // numeroRecibo pode ser extraído se disponível
                                    );
                                }
                            }
                        }
                    }
                }
                catch (Exception exDb)
                {
                    // Não interromper o fluxo se houver erro ao registrar no banco
                    System.Diagnostics.Debug.WriteLine($"Erro ao registrar consulta no banco: {exDb.Message}");
                }
            }
            catch (Exception ex)
            {
                rtbResultado.Clear();
                rtbResultado.AppendText($"ERRO: {ex.Message}");
                MessageBox.Show($"Erro ao consultar protocolo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnFiltrar_Click(object sender, EventArgs e)
        {
            BtnAtualizarLotes_Click(sender, e);
        }

        private void BtnAtualizarLotes_Click(object sender, EventArgs e)
        {
            try
            {
                lstLotes.Items.Clear();
                lotesCarregados = new List<LoteInfo>();
                lotesBancoCarregados = new List<LoteBancoInfo>();
                
                // Buscar lotes do banco de dados
                try
                {
                    var persistenceService = new EfinanceiraDatabasePersistenceService();
                    DateTime? dataInicio = null;
                    DateTime? dataFim = null;
                    string periodo = null;
                    
                    // Aplicar filtro de data se marcado
                    if (chkUsarFiltroData.Checked)
                    {
                        dataInicio = dtpFiltroDataInicio.Value.Date;
                        dataFim = dtpFiltroDataFim.Value.Date.AddDays(1).AddSeconds(-1);
                    }
                    
                    // Aplicar filtro de período se marcado
                    if (chkUsarFiltroPeriodo.Checked && cmbFiltroPeriodo.SelectedItem != null)
                    {
                        // Extrair apenas o período (YYYYMM) do item selecionado
                        // Formato: "202301 - Jan-Jun/2023" -> "202301"
                        string itemSelecionado = cmbFiltroPeriodo.SelectedItem.ToString();
                        periodo = itemSelecionado.Split(' ')[0]; // Pega apenas a primeira parte (YYYYMM)
                    }
                    
                    // Aplicar filtro de ambiente
                    string ambiente = null;
                    if (cmbFiltroAmbiente.SelectedItem != null)
                    {
                        string ambienteSelecionado = cmbFiltroAmbiente.SelectedItem.ToString();
                        if (ambienteSelecionado == "Teste")
                        {
                            ambiente = "TEST";
                        }
                        else if (ambienteSelecionado == "Produção")
                        {
                            ambiente = "PROD";
                        }
                        // Se for "Ambos", ambiente permanece null
                    }
                    
                    // Buscar lotes com os filtros aplicados
                    lotesBancoCarregados = persistenceService.BuscarLotes(dataInicio, dataFim, periodo, ambiente);
                    
                    if (lotesBancoCarregados.Count > 0)
                    {
                        foreach (var lote in lotesBancoCarregados.OrderByDescending(l => l.DataCriacao))
                        {
                            string tipoStr = lote.TipoLote.ToString();
                            string protocoloStr = !string.IsNullOrEmpty(lote.ProtocoloEnvio) 
                                ? $"Protocolo: {lote.ProtocoloEnvio}" 
                                : "Sem protocolo";
                            string periodoStr = !string.IsNullOrEmpty(lote.Periodo) 
                                ? $"Período: {lote.Periodo}" 
                                : "";
                            string dataStr = lote.DataCriacao.ToString("dd/MM/yyyy HH:mm:ss");
                            string retificacaoStr = lote.EhRetificacao ? " [RETIFICAÇÃO]" : "";
                            
                            string item = $"[{tipoStr}]{retificacaoStr} {protocoloStr} | {periodoStr} | {dataStr}";
                            lstLotes.Items.Add(item);
                        }
                        
                        string filtroInfo = "";
                        if (chkUsarFiltroData.Checked)
                        {
                            filtroInfo = $"Data: {dtpFiltroDataInicio.Value:dd/MM/yyyy} a {dtpFiltroDataFim.Value:dd/MM/yyyy}";
                        }
                        if (chkUsarFiltroPeriodo.Checked && !string.IsNullOrWhiteSpace(periodo))
                        {
                            if (!string.IsNullOrEmpty(filtroInfo)) filtroInfo += " | ";
                            filtroInfo += $"Período: {periodo}";
                        }
                        if (!string.IsNullOrEmpty(ambiente))
                        {
                            if (!string.IsNullOrEmpty(filtroInfo)) filtroInfo += " | ";
                            filtroInfo += $"Ambiente: {ambiente}";
                        }
                        else if (cmbFiltroAmbiente.SelectedItem != null && cmbFiltroAmbiente.SelectedItem.ToString() == "Ambos")
                        {
                            if (!string.IsNullOrEmpty(filtroInfo)) filtroInfo += " | ";
                            filtroInfo += "Ambiente: Ambos";
                        }
                        if (string.IsNullOrEmpty(filtroInfo))
                        {
                            filtroInfo = "Todos os lotes";
                        }
                        
                        MessageBox.Show($"Encontrados {lotesBancoCarregados.Count} lote(s) no banco de dados.\n\nFiltro: {filtroInfo}", 
                            "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string mensagem = "Nenhum lote encontrado";
                        if (chkUsarFiltroData.Checked)
                        {
                            mensagem += $" para o período de {dtpFiltroDataInicio.Value:dd/MM/yyyy} a {dtpFiltroDataFim.Value:dd/MM/yyyy}";
                        }
                        if (chkUsarFiltroPeriodo.Checked && !string.IsNullOrWhiteSpace(periodo))
                        {
                            mensagem += $" com período {periodo}";
                        }
                        if (!string.IsNullOrEmpty(ambiente))
                        {
                            mensagem += $" no ambiente {ambiente}";
                        }
                        mensagem += ".";
                        MessageBox.Show(mensagem, "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception exDb)
                {
                    MessageBox.Show($"Erro ao buscar lotes do banco: {exDb.Message}\n\nTentando carregar do sistema antigo...", 
                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    
                    // Fallback: carregar do sistema antigo
                var lotes = ProtocoloPersistenciaService.CarregarLotes();
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
                rtbDetalhes.Clear();
                
                // Priorizar lotes do banco
                if (lotesBancoCarregados != null && lstLotes.SelectedIndex < lotesBancoCarregados.Count)
                {
                    var lote = lotesBancoCarregados[lstLotes.SelectedIndex];
                    
                    // Tipo do lote - Formatação melhorada
                    string tipoStr = lote.TipoLote.ToString();
                    rtbDetalhes.AppendText("═══════════════════════════════════════════════════════════\n");
                    rtbDetalhes.AppendText("                    DETALHES DO LOTE\n");
                    rtbDetalhes.AppendText("═══════════════════════════════════════════════════════════\n\n");
                    
                    rtbDetalhes.AppendText($"📋 Tipo de Lote: {tipoStr}\n");
                    if (lote.EhRetificacao)
                    {
                        rtbDetalhes.AppendText($"⚠️  RETIFICAÇÃO (ID Lote Original: {lote.IdLoteOriginal ?? 0})\n");
                    }
                    rtbDetalhes.AppendText("\n");
                    
                    // Protocolo - Destacado
                    string protocoloStr = !string.IsNullOrEmpty(lote.ProtocoloEnvio) ? lote.ProtocoloEnvio : "Não disponível";
                    rtbDetalhes.AppendText($"📤 Protocolo de Envio: {protocoloStr}\n");
                    
                    // Período
                    string periodoStr = !string.IsNullOrEmpty(lote.Periodo) ? lote.Periodo : "Não informado";
                    rtbDetalhes.AppendText($"📅 Período: {periodoStr}\n");
                    rtbDetalhes.AppendText($"📊 Semestre: {lote.Semestre}\n");
                    rtbDetalhes.AppendText($"🔢 Número do Lote: {lote.NumeroLote}\n");
                    rtbDetalhes.AppendText($"🏢 CNPJ Declarante: {lote.CnpjDeclarante ?? "N/A"}\n");
                    rtbDetalhes.AppendText($"🌐 Ambiente: {lote.Ambiente ?? "N/A"}\n\n");
                    
                    // Estatísticas de eventos
                    rtbDetalhes.AppendText("═══════════════════════════════════════════════════════════\n");
                    rtbDetalhes.AppendText("              ESTATÍSTICAS DE EVENTOS\n");
                    rtbDetalhes.AppendText("═══════════════════════════════════════════════════════════\n\n");
                    
                    rtbDetalhes.AppendText($"📦 Quantidade de Eventos: {lote.QuantidadeEventos}\n");
                    rtbDetalhes.AppendText($"📝 Total Eventos Registrados: {lote.TotalEventosRegistrados}\n");
                    rtbDetalhes.AppendText($"👤 Eventos com CPF: {lote.TotalEventosComCpf}\n");
                    rtbDetalhes.AppendText($"✅ Eventos com Sucesso: {lote.TotalEventosSucesso}\n");
                    rtbDetalhes.AppendText($"❌ Eventos com Erro: {lote.TotalEventosComErro}\n\n");
                    
                    // Status
                    rtbDetalhes.AppendText("═══════════════════════════════════════════════════════════\n");
                    rtbDetalhes.AppendText("                         STATUS\n");
                    rtbDetalhes.AppendText("═══════════════════════════════════════════════════════════\n\n");
                    
                    rtbDetalhes.AppendText($"📊 Status Atual: {lote.Status ?? "N/A"}\n\n");
                    
                    // Datas
                    rtbDetalhes.AppendText("═══════════════════════════════════════════════════════════\n");
                    rtbDetalhes.AppendText("                         DATAS\n");
                    rtbDetalhes.AppendText("═══════════════════════════════════════════════════════════\n\n");
                    
                    rtbDetalhes.AppendText($"📅 Data Criação: {lote.DataCriacao:dd/MM/yyyy HH:mm:ss}\n");
                    if (lote.DataEnvio.HasValue)
                    {
                        rtbDetalhes.AppendText($"📤 Data Envio: {lote.DataEnvio.Value:dd/MM/yyyy HH:mm:ss}\n");
                    }
                    if (lote.DataConfirmacao.HasValue)
                    {
                        rtbDetalhes.AppendText($"✅ Data Confirmação: {lote.DataConfirmacao.Value:dd/MM/yyyy HH:mm:ss}\n");
                    }
                    
                    // Respostas
                    if (lote.CodigoRespostaEnvio.HasValue || lote.CodigoRespostaConsulta.HasValue)
                    {
                        rtbDetalhes.AppendText("\n");
                        rtbDetalhes.AppendText("═══════════════════════════════════════════════════════════\n");
                        rtbDetalhes.AppendText("                    RESPOSTAS DO SERVIDOR\n");
                        rtbDetalhes.AppendText("═══════════════════════════════════════════════════════════\n\n");
                    }
                    
                    if (lote.CodigoRespostaEnvio.HasValue)
                    {
                        rtbDetalhes.AppendText($"📤 Resposta do Envio:\n");
                        rtbDetalhes.AppendText($"   Código: {lote.CodigoRespostaEnvio}\n");
                        rtbDetalhes.AppendText($"   Descrição: {lote.DescricaoRespostaEnvio ?? "N/A"}\n\n");
                    }
                    if (lote.CodigoRespostaConsulta.HasValue)
                    {
                        rtbDetalhes.AppendText($"🔍 Resposta da Consulta:\n");
                        rtbDetalhes.AppendText($"   Código: {lote.CodigoRespostaConsulta}\n");
                        rtbDetalhes.AppendText($"   Descrição: {lote.DescricaoRespostaConsulta ?? "N/A"}\n\n");
                    }
                    
                    // Buscar eventos do lote
                    try
                    {
                        var persistenceService = new EfinanceiraDatabasePersistenceService();
                        var eventos = persistenceService.BuscarEventosDoLote(lote.IdLote);
                        
                        if (eventos.Count > 0)
                        {
                            rtbDetalhes.AppendText($"\n════════════════════════════════════════\n");
                            rtbDetalhes.AppendText($"EVENTOS DO LOTE ({eventos.Count}):\n");
                            rtbDetalhes.AppendText($"════════════════════════════════════════\n");
                            
                            int eventosComCpf = 0;
                            foreach (var evento in eventos.Take(10)) // Limitar a 10 para não ficar muito grande
                            {
                                rtbDetalhes.AppendText($"\nEvento ID: {evento.IdEventoXml ?? "N/A"}\n");
                                if (!string.IsNullOrEmpty(evento.Cpf))
                                {
                                    rtbDetalhes.AppendText($"  CPF: {evento.Cpf}\n");
                                    eventosComCpf++;
                                }
                                if (!string.IsNullOrEmpty(evento.Nome))
                                {
                                    rtbDetalhes.AppendText($"  Nome: {evento.Nome}\n");
                                }
                                rtbDetalhes.AppendText($"  Status: {evento.StatusEvento ?? "N/A"}\n");
                                if (evento.EhRetificacao)
                                {
                                    rtbDetalhes.AppendText($"  ⚠️  RETIFICAÇÃO\n");
                                }
                            }
                            
                            if (eventos.Count > 10)
                            {
                                rtbDetalhes.AppendText($"\n... e mais {eventos.Count - 10} evento(s)\n");
                            }
                            
                            rtbDetalhes.AppendText($"\nTotal de eventos com CPF: {eventosComCpf}\n");
                        }
                    }
                    catch (Exception exEventos)
                    {
                        rtbDetalhes.AppendText($"\n⚠️  Erro ao buscar eventos: {exEventos.Message}\n");
                    }
                    
                    // Preencher campo de protocolo automaticamente
                    if (!string.IsNullOrEmpty(lote.ProtocoloEnvio))
                    {
                        txtProtocolo.Text = lote.ProtocoloEnvio;
                        rtbDetalhes.AppendText($"\n[Protocolo preenchido automaticamente - clique em 'Consultar' para verificar status]\n");
                }
                else
                {
                        txtProtocolo.Text = "";
                        rtbDetalhes.AppendText($"\n[Este lote ainda não possui protocolo - não foi enviado ou aguardando resposta]\n");
                    }
                }
                // Fallback para sistema antigo
                else if (lotesCarregados != null && lstLotes.SelectedIndex < lotesCarregados.Count)
                {
                    var lote = lotesCarregados[lstLotes.SelectedIndex];
                    
                    string tipoStr = lote.Tipo.ToString();
                    rtbDetalhes.AppendText($"════════════════════════════════════════\n");
                    rtbDetalhes.AppendText($"TIPO DE LOTE: {tipoStr}\n");
                    rtbDetalhes.AppendText($"════════════════════════════════════════\n\n");
                    
                    rtbDetalhes.AppendText($"Quantidade de Eventos: {lote.QuantidadeEventos}\n");
                    
                    string periodoStr = !string.IsNullOrEmpty(lote.Periodo) ? lote.Periodo : "Não informado";
                    rtbDetalhes.AppendText($"Período: {periodoStr}\n");
                    
                    string protocoloStr = !string.IsNullOrEmpty(lote.Protocolo) ? lote.Protocolo : "Não disponível";
                    rtbDetalhes.AppendText($"Protocolo: {protocoloStr}\n");
                    
                    rtbDetalhes.AppendText($"Status: {lote.Status}\n");
                    rtbDetalhes.AppendText($"Data Processamento: {lote.DataProcessamento:dd/MM/yyyy HH:mm:ss}\n\n");
                    
                    if (!string.IsNullOrEmpty(lote.Protocolo))
                    {
                        txtProtocolo.Text = lote.Protocolo;
                    }
                    else
                    {
                        txtProtocolo.Text = "";
                    }
                }
            }
            catch (Exception ex)
            {
                rtbDetalhes.Clear();
                rtbDetalhes.AppendText($"Erro ao carregar detalhes: {ex.Message}");
            }
        }

        /// <summary>
        /// Popula o ComboBox de períodos com os últimos 5 anos e próximos 5 anos, cada um com 2 semestres
        /// </summary>
        private void PopularComboPeriodos()
        {
            cmbFiltroPeriodo.Items.Clear();
            
            int anoAtual = DateTime.Now.Year;
            
            // Adicionar períodos dos últimos 5 anos (5 anos para trás)
            for (int ano = anoAtual - 5; ano < anoAtual; ano++)
            {
                // Primeiro semestre (Jan-Jun) - mês 01 ou 06
                cmbFiltroPeriodo.Items.Add($"{ano}01 - Jan-Jun/{ano}");
                cmbFiltroPeriodo.Items.Add($"{ano}06 - Jan-Jun/{ano}");
                
                // Segundo semestre (Jul-Dez) - mês 02 ou 12
                cmbFiltroPeriodo.Items.Add($"{ano}02 - Jul-Dez/{ano}");
                cmbFiltroPeriodo.Items.Add($"{ano}12 - Jul-Dez/{ano}");
            }
            
            // Adicionar períodos do ano atual
            cmbFiltroPeriodo.Items.Add($"{anoAtual}01 - Jan-Jun/{anoAtual}");
            cmbFiltroPeriodo.Items.Add($"{anoAtual}06 - Jan-Jun/{anoAtual}");
            cmbFiltroPeriodo.Items.Add($"{anoAtual}02 - Jul-Dez/{anoAtual}");
            cmbFiltroPeriodo.Items.Add($"{anoAtual}12 - Jul-Dez/{anoAtual}");
            
            // Adicionar períodos dos próximos 5 anos (5 anos para frente)
            for (int ano = anoAtual + 1; ano <= anoAtual + 5; ano++)
            {
                // Primeiro semestre (Jan-Jun) - mês 01 ou 06
                cmbFiltroPeriodo.Items.Add($"{ano}01 - Jan-Jun/{ano}");
                cmbFiltroPeriodo.Items.Add($"{ano}06 - Jan-Jun/{ano}");
                
                // Segundo semestre (Jul-Dez) - mês 02 ou 12
                cmbFiltroPeriodo.Items.Add($"{ano}02 - Jul-Dez/{ano}");
                cmbFiltroPeriodo.Items.Add($"{ano}12 - Jul-Dez/{ano}");
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

        /// <summary>
        /// Obtém orientações sobre como tratar erros específicos da e-Financeira
        /// </summary>
        private string ObterOrientacaoTratamentoErro(string codigoErro, string descricaoErro)
        {
            if (string.IsNullOrWhiteSpace(codigoErro) && string.IsNullOrWhiteSpace(descricaoErro))
                return null;

            string codigo = codigoErro?.ToUpper().Trim() ?? "";
            string descricao = descricaoErro?.ToUpper() ?? "";

            // Erros comuns e suas orientações
            if (codigo.Contains("MS1034") || descricao.Contains("JÁ EXISTE E-FINANCEIRA"))
            {
                return "Este erro indica que já existe uma e-Financeira aberta para este período. " +
                       "SOLUÇÃO: Verifique se já foi enviada uma abertura para este período. " +
                       "Se sim, você deve enviar movimentações ou fechamento. " +
                       "Se não, verifique se o período está correto e se há conflito com outro lote.";
            }

            if (codigo.Contains("MS1047") || descricao.Contains("NÃO EXISTE E-FINANCEIRA ABERTA"))
            {
                return "Este erro indica que não existe e-Financeira aberta para o período informado. " +
                       "SOLUÇÃO: Você deve enviar primeiro um lote de ABERTURA para este período antes de enviar movimentações. " +
                       "Verifique se a abertura foi enviada e aceita pela Receita Federal.";
            }

            if (codigo.Contains("MS1001") || descricao.Contains("CNPJ"))
            {
                return "Erro relacionado ao CNPJ do declarante. " +
                       "SOLUÇÃO: Verifique se o CNPJ está correto na configuração e se corresponde ao certificado digital utilizado.";
            }

            if (codigo.Contains("MS1002") || descricao.Contains("CPF"))
            {
                return "Erro relacionado ao CPF do declarado. " +
                       "SOLUÇÃO: Verifique se o CPF está correto, sem pontos ou hífens, e se possui 11 dígitos. " +
                       "Corrija o CPF no banco de dados e reenvie o lote.";
            }

            if (codigo.Contains("MS1003") || descricao.Contains("PERÍODO"))
            {
                return "Erro relacionado ao período informado. " +
                       "SOLUÇÃO: Verifique se o período está no formato correto (YYYYMM, ex: 202301). " +
                       "O período deve corresponder a um semestre válido (01-06 ou 07-12).";
            }

            if (codigo.Contains("MS1004") || descricao.Contains("ASSINATURA"))
            {
                return "Erro na assinatura digital do XML. " +
                       "SOLUÇÃO: Verifique se o certificado digital está válido e tem permissão de assinatura. " +
                       "Regenere o XML e assine novamente com um certificado válido.";
            }

            if (codigo.Contains("MS1005") || descricao.Contains("CRIPTOGRAFIA"))
            {
                return "Erro na criptografia do XML. " +
                       "SOLUÇÃO: Verifique se o certificado do servidor está correto e válido. " +
                       "Regenere o XML e criptografe novamente.";
            }

            if (codigo.Contains("MS1006") || descricao.Contains("VALIDAÇÃO"))
            {
                return "Erro de validação do XML. " +
                       "SOLUÇÃO: Verifique se todos os campos obrigatórios estão preenchidos corretamente. " +
                       "Consulte o manual da e-Financeira para verificar as regras de validação.";
            }

            if (descricao.Contains("RETIFICAÇÃO") || descricao.Contains("RETIFIC"))
            {
                return "Erro relacionado a retificação. " +
                       "SOLUÇÃO: Se este é um lote de retificação, verifique se o número do recibo do lote original está correto. " +
                       "Certifique-se de que o lote original foi processado com sucesso antes de enviar a retificação.";
            }

            if (descricao.Contains("DUPLICADO") || descricao.Contains("JÁ EXISTE"))
            {
                return "O evento ou lote já foi enviado anteriormente. " +
                       "SOLUÇÃO: Verifique se este lote já foi processado. Se sim, não é necessário reenviar. " +
                       "Se precisar corrigir, envie uma retificação do lote original.";
            }

            // Orientação genérica para outros erros
            return "Para tratar este erro: " +
                   "1. Verifique a descrição do erro acima para entender o problema específico. " +
                   "2. Corrija os dados no banco de dados ou na configuração conforme necessário. " +
                   "3. Regenere o XML com os dados corrigidos. " +
                   "4. Se for um erro de retificação, verifique se o lote original foi aceito. " +
                   "5. Consulte o manual da e-Financeira para mais detalhes sobre este código de erro.";
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
