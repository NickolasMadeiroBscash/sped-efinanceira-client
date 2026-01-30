using System;
using System.Xml.Serialization;

namespace ExemploAssinadorXML.Models
{
    [Serializable]
    [XmlRoot("DadosAbertura")]
    public class DadosAbertura
    {
        public string CnpjDeclarante { get; set; }
        public string DtInicio { get; set; } // Formato: AAAA-MM-DD
        public string DtFim { get; set; } // Formato: AAAA-MM-DD
        public int TipoAmbiente { get; set; } // 1 = Produção, 2 = Homologação
        public int AplicacaoEmissora { get; set; } // 1 = Aplicação do contribuinte, 2 = Outros
        public int IndRetificacao { get; set; } // 1 = Original, 2 = Retificação espontânea, 3 = Retificação a pedido
        public string NrRecibo { get; set; }
        public bool IndicarMovOpFin { get; set; }
        
        // Dados de AberturaMovOpFin
        public DadosResponsavelRMF ResponsavelRMF { get; set; }
        public DadosRespeFin RespeFin { get; set; }
        public DadosRepresLegal RepresLegal { get; set; }
        
        // Dados de AberturaPP
        public string[] TiposEmpresaPP { get; set; }
    }

    [Serializable]
    public class DadosResponsavelRMF
    {
        public string Cnpj { get; set; }
        public string Cpf { get; set; }
        public string Nome { get; set; }
        public string Setor { get; set; }
        public string TelefoneDDD { get; set; }
        public string TelefoneNumero { get; set; }
        public string TelefoneRamal { get; set; }
        public string EnderecoLogradouro { get; set; }
        public string EnderecoNumero { get; set; }
        public string EnderecoComplemento { get; set; }
        public string EnderecoBairro { get; set; }
        public string EnderecoCEP { get; set; }
        public string EnderecoMunicipio { get; set; }
        public string EnderecoUF { get; set; }
    }

    [Serializable]
    public class DadosRespeFin
    {
        public string Cpf { get; set; }
        public string Nome { get; set; }
        public string Setor { get; set; }
        public string TelefoneDDD { get; set; }
        public string TelefoneNumero { get; set; }
        public string TelefoneRamal { get; set; }
        public string EnderecoLogradouro { get; set; }
        public string EnderecoNumero { get; set; }
        public string EnderecoComplemento { get; set; }
        public string EnderecoBairro { get; set; }
        public string EnderecoCEP { get; set; }
        public string EnderecoMunicipio { get; set; }
        public string EnderecoUF { get; set; }
        public string Email { get; set; }
    }

    [Serializable]
    public class DadosRepresLegal
    {
        public string Cpf { get; set; }
        public string Setor { get; set; }
        public string TelefoneDDD { get; set; }
        public string TelefoneNumero { get; set; }
        public string TelefoneRamal { get; set; }
    }
}
