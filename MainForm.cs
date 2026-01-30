using System;
using System.Windows.Forms;
using ExemploAssinadorXML.Forms;

namespace ExemploAssinadorXML
{
    public partial class MainForm : Form
    {
        private TabControl tabControl;
        private TabPage tabTutorial;
        private TabPage tabConfiguracao;
        private TabPage tabProcessamento;
        private TabPage tabConsulta;
        
        private TutorialForm tutorialForm;
        private ConfiguracaoForm configuracaoForm;
        private ProcessamentoForm processamentoForm;
        private ConsultaForm consultaForm;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Form
            this.Text = "Sistema e-Financeira - Assinatura e Envio";
            this.Size = new System.Drawing.Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new System.Drawing.Size(1000, 600);

            // TabControl
            tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;
            tabControl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            // Tab Tutorial (primeira aba)
            tabTutorial = new TabPage();
            tabTutorial.Text = "📚 Tutorial";
            tabTutorial.UseVisualStyleBackColor = true;
            tabControl.TabPages.Add(tabTutorial);

            // Tab Configuração
            tabConfiguracao = new TabPage();
            tabConfiguracao.Text = "Configuração";
            tabConfiguracao.UseVisualStyleBackColor = true;
            tabControl.TabPages.Add(tabConfiguracao);

            // Tab Processamento
            tabProcessamento = new TabPage();
            tabProcessamento.Text = "Processamento";
            tabProcessamento.UseVisualStyleBackColor = true;
            tabControl.TabPages.Add(tabProcessamento);

            // Tab Consulta
            tabConsulta = new TabPage();
            tabConsulta.Text = "Consulta";
            tabConsulta.UseVisualStyleBackColor = true;
            tabControl.TabPages.Add(tabConsulta);

            this.Controls.Add(tabControl);

            // Criar formulários filhos
            tutorialForm = new TutorialForm();
            tutorialForm.Dock = DockStyle.Fill;
            tutorialForm.TopLevel = false;
            tutorialForm.FormBorderStyle = FormBorderStyle.None;
            tabTutorial.Controls.Add(tutorialForm);
            tutorialForm.Show();

            configuracaoForm = new ConfiguracaoForm();
            configuracaoForm.Dock = DockStyle.Fill;
            configuracaoForm.TopLevel = false;
            configuracaoForm.FormBorderStyle = FormBorderStyle.None;
            tabConfiguracao.Controls.Add(configuracaoForm);
            configuracaoForm.Show();

            processamentoForm = new ProcessamentoForm();
            processamentoForm.Dock = DockStyle.Fill;
            processamentoForm.TopLevel = false;
            processamentoForm.FormBorderStyle = FormBorderStyle.None;
            processamentoForm.ConfigForm = configuracaoForm;
            tabProcessamento.Controls.Add(processamentoForm);
            processamentoForm.Show();

            consultaForm = new ConsultaForm();
            consultaForm.Dock = DockStyle.Fill;
            consultaForm.TopLevel = false;
            consultaForm.FormBorderStyle = FormBorderStyle.None;
            consultaForm.ConfigForm = configuracaoForm;
            tabConsulta.Controls.Add(consultaForm);
            consultaForm.Show();

            this.ResumeLayout(false);
        }
    }
}
