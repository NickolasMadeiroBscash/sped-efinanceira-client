namespace ExemploAssinadorXML
{
    partial class ExemploAssinadorXML
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.bntAssinar = new System.Windows.Forms.Button();
            this.btnValidarAssinatura = new System.Windows.Forms.Button();
            this.txtCaminhoArquivo = new System.Windows.Forms.TextBox();
            this.lblCaminho = new System.Windows.Forms.Label();
            this.ofdArquivoXML = new System.Windows.Forms.OpenFileDialog();
            this.btnSelecionarArquivo = new System.Windows.Forms.Button();
            this.lblSHA256 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // bntAssinar
            // 
            this.bntAssinar.Location = new System.Drawing.Point(178, 65);
            this.bntAssinar.Name = "bntAssinar";
            this.bntAssinar.Size = new System.Drawing.Size(108, 23);
            this.bntAssinar.TabIndex = 0;
            this.bntAssinar.Text = "Assinar";
            this.bntAssinar.UseVisualStyleBackColor = true;
            this.bntAssinar.Click += new System.EventHandler(this.bntAssinar_Click);
            // 
            // btnValidarAssinatura
            // 
            this.btnValidarAssinatura.Location = new System.Drawing.Point(313, 65);
            this.btnValidarAssinatura.Name = "btnValidarAssinatura";
            this.btnValidarAssinatura.Size = new System.Drawing.Size(111, 23);
            this.btnValidarAssinatura.TabIndex = 1;
            this.btnValidarAssinatura.Text = "Validar Assinatura";
            this.btnValidarAssinatura.UseVisualStyleBackColor = true;
            this.btnValidarAssinatura.Click += new System.EventHandler(this.btnValidarAssinatura_Click);
            // 
            // txtCaminhoArquivo
            // 
            this.txtCaminhoArquivo.Location = new System.Drawing.Point(9, 30);
            this.txtCaminhoArquivo.Name = "txtCaminhoArquivo";
            this.txtCaminhoArquivo.Size = new System.Drawing.Size(541, 20);
            this.txtCaminhoArquivo.TabIndex = 2;
            // 
            // lblCaminho
            // 
            this.lblCaminho.AutoSize = true;
            this.lblCaminho.Location = new System.Drawing.Point(6, 15);
            this.lblCaminho.Name = "lblCaminho";
            this.lblCaminho.Size = new System.Drawing.Size(118, 13);
            this.lblCaminho.TabIndex = 3;
            this.lblCaminho.Text = "Caminho Arquivo XML :";
            // 
            // ofdArquivoXML
            // 
            this.ofdArquivoXML.FileName = "evento.xml";
            // 
            // btnSelecionarArquivo
            // 
            this.btnSelecionarArquivo.Location = new System.Drawing.Point(562, 29);
            this.btnSelecionarArquivo.Name = "btnSelecionarArquivo";
            this.btnSelecionarArquivo.Size = new System.Drawing.Size(126, 23);
            this.btnSelecionarArquivo.TabIndex = 4;
            this.btnSelecionarArquivo.Text = "Selecionar Arquivo...";
            this.btnSelecionarArquivo.UseVisualStyleBackColor = true;
            this.btnSelecionarArquivo.Click += new System.EventHandler(this.btnSelecionarArquivo_Click);
            // 
            // lblSHA256
            // 
            this.lblSHA256.AutoSize = true;
            this.lblSHA256.Location = new System.Drawing.Point(612, 81);
            this.lblSHA256.Name = "lblSHA256";
            this.lblSHA256.Size = new System.Drawing.Size(86, 13);
            this.lblSHA256.TabIndex = 5;
            this.lblSHA256.Text = "Versão: SHA256";
            // 
            // ExemploAssinadorXML
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(705, 103);
            this.Controls.Add(this.lblSHA256);
            this.Controls.Add(this.btnSelecionarArquivo);
            this.Controls.Add(this.lblCaminho);
            this.Controls.Add(this.txtCaminhoArquivo);
            this.Controls.Add(this.btnValidarAssinatura);
            this.Controls.Add(this.bntAssinar);
            this.Name = "ExemploAssinadorXML";
            this.Text = "Exemplo Assinador XML e-Financeira";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bntAssinar;
        private System.Windows.Forms.Button btnValidarAssinatura;
        private System.Windows.Forms.TextBox txtCaminhoArquivo;
        private System.Windows.Forms.Label lblCaminho;
        private System.Windows.Forms.OpenFileDialog ofdArquivoXML;
        private System.Windows.Forms.Button btnSelecionarArquivo;
        private System.Windows.Forms.Label lblSHA256;
    }
}

