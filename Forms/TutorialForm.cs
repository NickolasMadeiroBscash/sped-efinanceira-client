using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExemploAssinadorXML.Forms
{
    public partial class TutorialForm : Form
    {
        private RichTextBox rtbTutorial;
        private Panel panelHeader;

        public TutorialForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Header
            panelHeader = new Panel();
            panelHeader.Dock = DockStyle.Top;
            panelHeader.Height = 60;
            panelHeader.BackColor = Color.FromArgb(0, 102, 204);
            
            Label lblTitulo = new Label();
            lblTitulo.Text = "📚 Tutorial e-Financeira - Guia de Uso";
            lblTitulo.Font = new Font("Microsoft Sans Serif", 16F, FontStyle.Bold);
            lblTitulo.ForeColor = Color.White;
            lblTitulo.Location = new Point(20, 15);
            lblTitulo.AutoSize = true;
            panelHeader.Controls.Add(lblTitulo);

            // RichTextBox para o tutorial
            rtbTutorial = new RichTextBox();
            rtbTutorial.Dock = DockStyle.Fill;
            rtbTutorial.ReadOnly = true;
            rtbTutorial.BackColor = Color.White;
            rtbTutorial.Font = new Font("Microsoft Sans Serif", 10F);
            rtbTutorial.Margin = new Padding(10);

            this.Controls.Add(rtbTutorial);
            this.Controls.Add(panelHeader);

            // Carregar conteúdo
            CarregarConteudoTutorial();

            this.ResumeLayout(false);
        }

        private void CarregarConteudoTutorial()
        {
            string conteudo = @"
╔══════════════════════════════════════════════════════════════════════════════╗
║                    SISTEMA E-FINANCEIRA - GUIA COMPLETO                      ║
╚══════════════════════════════════════════════════════════════════════════════╝

┌──────────────────────────────────────────────────────────────────────────────┐
│ 1. O QUE É ESTE SISTEMA?                                                     │
└──────────────────────────────────────────────────────────────────────────────┘

Este sistema automatiza o processo de geração, assinatura digital, criptografia e 
envio de lotes para a e-Financeira da Receita Federal do Brasil.

O sistema realiza automaticamente:
  ✓ Geração de XMLs de abertura, movimentação e fechamento
  ✓ Assinatura digital com certificado A1 ou A3
  ✓ Criptografia dos lotes (AES + RSA)
  ✓ Envio para a Receita Federal
  ✓ Consulta de protocolos e status


┌──────────────────────────────────────────────────────────────────────────────┐
│ 2. FLUXO COMPLETO DO PROCESSO                                                │
└──────────────────────────────────────────────────────────────────────────────┘

O processo da e-Financeira segue uma sequência OBRIGATÓRIA:

  PASSO 1: CONFIGURAÇÃO INICIAL
  ─────────────────────────────
  • Configure os dados da empresa (CNPJ, certificados)
  • Preencha os dados de abertura (Responsável RMF, RespeFin, Representante Legal)
  • Configure os dados de fechamento
  • Selecione o ambiente (TESTE ou PRODUÇÃO)
  • Salve as configurações

  PASSO 2: ABERTURA DA E-FINANCEIRA
  ──────────────────────────────────
  • Vá para a aba ""Processamento""
  • Clique em ""Processar Abertura""
  • O sistema irá:
    → Gerar o XML de abertura
    → Assinar digitalmente
    → Criptografar
    → Enviar para a Receita (se habilitado)
  • Aguarde o protocolo de recebimento

  PASSO 3: ENVIO DE LOTES DE MOVIMENTAÇÃO
  ────────────────────────────────────────
  • Após a abertura ser aceita, envie os lotes de movimentação financeira
  • Vá para a aba ""Processamento""
  • Clique em ""Processar Movimentação""
  • O sistema processará todos os lotes do período configurado
  • Cada lote será assinado, criptografado e enviado

  PASSO 4: FECHAMENTO DA E-FINANCEIRA
  ─────────────────────────────────────
  • Após enviar todos os lotes de movimentação, gere o fechamento
  • Vá para a aba ""Consulta""
  • Clique em ""Gerar Fechamento""
  • Informe o período (formato YYYYMM: 202406 = Jan-Jun, 202412 = Jul-Dez)
  • O sistema gerará o XML de fechamento
  • Processe o fechamento na aba ""Processamento""


┌──────────────────────────────────────────────────────────────────────────────┐
│ 3. COMO USAR CADA ABA                                                        │
└──────────────────────────────────────────────────────────────────────────────┘

  ┌─ ABA: CONFIGURAÇÃO ─────────────────────────────────────────────────────┐
  │                                                                          │
  │ Esta aba é usada para configurar todos os dados necessários:            │
  │                                                                          │
  │ • Configuração Geral:                                                   │
  │   - CNPJ da empresa declarante                                          │
  │   - Certificado para assinatura (thumbprint)                            │
  │   - Certificado do servidor para criptografia                            │
  │   - Ambiente (TESTE ou PRODUÇÃO)                                         │
  │   - Diretório onde os lotes serão salvos                                │
  │                                                                          │
  │ • Aba Abertura:                                                         │
  │   - Datas de início e fim do período semestral                          │
  │   - Dados do Responsável RMF (CNPJ, CPF, Nome, Telefone, Endereço)     │
  │   - Dados do Responsável e-Financeira (CPF, Nome, Email, etc.)          │
  │   - Dados do Representante Legal (CPF, Setor, Telefone)                 │
  │                                                                          │
  │ • Aba Fechamento:                                                       │
  │   - Datas de início e fim do período                                     │
  │   - Situação especial (se aplicável)                                     │
  │   - Indicadores de fechamento (PP, MovOpFin, MovOpFinAnual)              │
  │                                                                          │
  │ • Configurações de Processamento:                                        │
  │   - Page Size: Tamanho da página para consultas (padrão: 500 produção)  │
  │   - Evento Offset: Onde começar a gerar eventos (padrão: 1)              │
  │   - Offset Registros: Pular registros iniciais (padrão: 0 produção)    │
  │   - Max Lotes: Limitar quantidade de lotes (padrão: ilimitado produção) │
  │                                                                          │
  │ IMPORTANTE: Sempre clique em ""Salvar Configurações"" após alterar!     │
  └──────────────────────────────────────────────────────────────────────────┘

  ┌─ ABA: PROCESSAMENTO ───────────────────────────────────────────────────┐
  │                                                                          │
  │ Esta aba é usada para processar os lotes:                              │
  │                                                                          │
  │ • Processar Abertura:                                                  │
  │   - Gera, assina, criptografa e envia o lote de abertura               │
  │   - Deve ser feito PRIMEIRO                                             │
  │                                                                          │
  │ • Processar Movimentação:                                               │
  │   - Processa todos os lotes de movimentação financeira do período     │
  │   - Só funciona após a abertura ser aceita                               │
  │                                                                          │
  │ • Processar Fechamento:                                                 │
  │   - Processa o lote de fechamento gerado                                │
  │   - Só funciona após todos os lotes de movimentação serem enviados      │
  │                                                                          │
  │ • Opções:                                                               │
  │   ☐ Apenas Processar: Marque para NÃO enviar, apenas gerar arquivos    │
  │                                                                          │
  │ Durante o processamento você verá:                                      │
  │   - Etapa atual (Assinando, Criptografando, Enviando...)                │
  │   - Progresso geral                                                     │
  │   - Estatísticas (quantos processados, enviados, com erro)              │
  │   - Log detalhado de cada operação                                      │
  └──────────────────────────────────────────────────────────────────────────┘

  ┌─ ABA: CONSULTA ─────────────────────────────────────────────────────────┐
  │                                                                          │
  │ Esta aba permite consultar lotes processados e gerar fechamento:       │
  │                                                                          │
  │ • Consultar Protocolo:                                                  │
  │   - Informe o número do protocolo recebido                              │
  │   - Veja o status do lote (Processado, Rejeitado, etc.)                 │
  │                                                                          │
  │ • Gerar Fechamento:                                                     │
  │   - Clique no botão ""Gerar Fechamento""                                │
  │   - Informe o período no formato YYYYMM:                                │
  │     * 202406 = Janeiro a Junho de 2024                                  │
  │     * 202412 = Julho a Dezembro de 2024                                 │
  │   - O sistema calculará automaticamente as datas                        │
  │   - Clique em ""Gerar Fechamento"" para criar o XML                     │
  └──────────────────────────────────────────────────────────────────────────┘


┌──────────────────────────────────────────────────────────────────────────────┐
│ 4. INFORMAÇÕES IMPORTANTES SOBRE PREENCHIMENTO                               │
└──────────────────────────────────────────────────────────────────────────────┘

  PERÍODOS SEMESTRAIS:
  ────────────────────
  A e-Financeira trabalha com períodos semestrais:
  
  • 1º Semestre: 01/01 até 30/06 (período: YYYY06)
  • 2º Semestre: 01/07 até 31/12 (período: YYYY12)
  
  Exemplos:
  • Período 202406 = 01/01/2024 até 30/06/2024
  • Período 202412 = 01/07/2024 até 31/12/2024


  ORDEM OBRIGATÓRIA:
  ──────────────────
  1. Primeiro: Enviar evento de CADASTRO da empresa declarante (se necessário)
  2. Segundo: Enviar evento de ABERTURA do período
  3. Terceiro: Enviar lotes de MOVIMENTAÇÃO financeira
  4. Quarto: Enviar evento de FECHAMENTO do período
  
  ⚠️ ATENÇÃO: Não é possível enviar movimentação sem abertura aceita!
  ⚠️ ATENÇÃO: Não é possível enviar fechamento sem todas as movimentações!


  AMBIENTES:
  ──────────
  • TESTE (Homologação): Use para validar antes de enviar dados reais
  • PRODUÇÃO: Use apenas quando tiver certeza de que tudo está correto
  
  ⚠️ IMPORTANTE: Dados de produção não podem ser enviados para teste e vice-versa!


  CERTIFICADOS:
  ────────────
  • Certificado de Assinatura: Usado para assinar digitalmente os XMLs
  • Certificado do Servidor: Usado para criptografar os lotes
  
  ⚠️ Os certificados devem estar instalados no Windows e ter permissão de uso!


  VALIDAÇÕES IMPORTANTES:
  ───────────────────────
  • CNPJ e CPF devem ser válidos (com dígitos verificadores corretos)
  • Datas devem estar no formato AAAA-MM-DD (ex: 2024-01-01)
  • Período de abertura e fechamento devem corresponder ao mesmo semestre
  • Todos os campos obrigatórios devem ser preenchidos


┌──────────────────────────────────────────────────────────────────────────────┐
│ 5. DICAS E BOAS PRÁTICAS                                                     │
└──────────────────────────────────────────────────────────────────────────────┘

  ✓ Sempre teste primeiro no ambiente de TESTE antes de usar PRODUÇÃO
  ✓ Salve as configurações após qualquer alteração
  ✓ Verifique os logs durante o processamento para identificar problemas
  ✓ Mantenha backup dos arquivos XML gerados
  ✓ Consulte os protocolos após o envio para confirmar o recebimento
  ✓ Use o modo ""Apenas Processar"" para validar sem enviar
  ✓ Verifique se os certificados estão válidos antes de processar


┌──────────────────────────────────────────────────────────────────────────────┐
│ 6. REFERÊNCIA AO MANUAL OFICIAL                                             │
└──────────────────────────────────────────────────────────────────────────────┘

Este sistema segue o Manual de Preenchimento e-Financeira - Anexo II - Versão 2.0

Para informações detalhadas sobre:
  • Leiautes dos eventos (Cadastro, Abertura, Fechamento, Exclusão)
  • Regras de validação e mensagens de erro
  • Tabelas de referência (Países, Municípios, UF, etc.)
  • Formato dos campos e valores permitidos

Consulte o manual oficial disponível no site da Receita Federal:
  http://sped.rfb.gov.br/


┌──────────────────────────────────────────────────────────────────────────────┐
│ 7. RESOLUÇÃO DE PROBLEMAS COMUNS                                            │
└──────────────────────────────────────────────────────────────────────────────┘

  PROBLEMA: ""Certificado não encontrado""
  SOLUÇÃO: Verifique se o certificado está instalado e o thumbprint está correto

  PROBLEMA: ""Erro ao assinar XML""
  SOLUÇÃO: Verifique se o certificado tem permissão de assinatura digital

  PROBLEMA: ""Lote rejeitado pela Receita""
  SOLUÇÃO: Consulte o protocolo para ver a mensagem de erro específica

  PROBLEMA: ""Não é possível processar movimentação""
  SOLUÇÃO: Verifique se a abertura foi enviada e aceita primeiro

  PROBLEMA: ""Período inválido""
  SOLUÇÃO: Use formato YYYYMM (ex: 202406 ou 202412)


╔══════════════════════════════════════════════════════════════════════════════╗
║                         FIM DO TUTORIAL                                      ║
║                                                                              ║
║  Em caso de dúvidas, consulte o manual oficial ou entre em contato com      ║
║  a equipe de desenvolvimento.                                                ║
╚══════════════════════════════════════════════════════════════════════════════╝
";

            rtbTutorial.Text = conteudo;
            rtbTutorial.SelectionStart = 0;
            rtbTutorial.SelectionLength = 0;
        }
    }
}
