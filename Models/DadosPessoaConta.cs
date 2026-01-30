using System;

namespace ExemploAssinadorXML.Models
{
    public class DadosPessoaConta
    {
        public long IdPessoa { get; set; }
        public string Documento { get; set; }
        public string Nome { get; set; }
        public string Cpf { get; set; }
        public string Nacionalidade { get; set; }
        public string Telefone { get; set; }
        public string Email { get; set; }
        public long IdConta { get; set; }
        public string NumeroConta { get; set; }
        public string DigitoConta { get; set; }
        public decimal SaldoAtual { get; set; }
        public string Logradouro { get; set; }
        public string Numero { get; set; }
        public string Complemento { get; set; }
        public string Bairro { get; set; }
        public string Cep { get; set; }
        public string TipoLogradouro { get; set; }
        public string EnderecoLivre { get; set; }
        public decimal TotCreditos { get; set; }
        public decimal TotDebitos { get; set; }
    }
}
