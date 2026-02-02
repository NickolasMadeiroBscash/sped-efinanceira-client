using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace ExemploAssinadorXML.Models
{
    [Serializable]
    [XmlRoot("DadosCadastroDeclarante")]
    public class DadosCadastroDeclarante
    {
        public string CnpjDeclarante { get; set; }
        public int TipoAmbiente { get; set; } // 1 = Produção, 2 = Homologação
        public int AplicacaoEmissora { get; set; } // 1 = Aplicação do contribuinte, 2 = Outros
        public int IndRetificacao { get; set; } // 1 = Original, 2 = Retificação espontânea, 3 = Retificação a pedido
        public string NrRecibo { get; set; }
        
        // infoCadastro
        public string GIIN { get; set; }
        public string CategoriaDeclarante { get; set; }
        public List<DadosTipoInstPgto> TiposInstPgto { get; set; }
        public List<DadosNIF> NIFs { get; set; }
        public string Nome { get; set; }
        public string TpNome { get; set; }
        public string EnderecoLivre { get; set; }
        public string TpEndereco { get; set; }
        public string Municipio { get; set; }
        public string UF { get; set; }
        public string CEP { get; set; }
        public string Pais { get; set; }
        public List<string> PaisResid { get; set; } // Lista de países de residência fiscal (deve conter "BR")
        public List<DadosEnderecoOutros> EnderecosOutros { get; set; }
    }

    [Serializable]
    public class DadosTipoInstPgto
    {
        public string TpInstPgto { get; set; } // 1, 2 ou 3
    }

    [Serializable]
    public class DadosNIF
    {
        public string NumeroNIF { get; set; }
        public string PaisEmissao { get; set; }
        public string TpNIF { get; set; }
    }

    [Serializable]
    public class DadosEnderecoOutros
    {
        public string TpEndereco { get; set; }
        public string EnderecoLivre { get; set; }
        public DadosEnderecoEstrutura EnderecoEstrutura { get; set; }
        public string Pais { get; set; }
    }

    [Serializable]
    public class DadosEnderecoEstrutura
    {
        public string EnderecoLivre { get; set; }
        public DadosEndereco Endereco { get; set; }
        public string CEP { get; set; }
        public string Municipio { get; set; }
        public string UF { get; set; }
    }

    [Serializable]
    public class DadosEndereco
    {
        public string Logradouro { get; set; }
        public string Numero { get; set; }
        public string Complemento { get; set; }
        public string Andar { get; set; }
        public string Bairro { get; set; }
        public string CaixaPostal { get; set; }
    }
}
