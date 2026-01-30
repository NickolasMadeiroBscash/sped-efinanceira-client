using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace ExemploAssinadorXML.Models
{
    [Serializable]
    [XmlRoot("DadosFechamento")]
    public class DadosFechamento
    {
        public string CnpjDeclarante { get; set; }
        public string DtInicio { get; set; } // Formato: AAAA-MM-DD
        public string DtFim { get; set; } // Formato: AAAA-MM-DD
        public int TipoAmbiente { get; set; } // 1 = Produção, 2 = Homologação
        public int AplicacaoEmissora { get; set; } // 1 = Aplicação do contribuinte, 2 = Outros
        public int IndRetificacao { get; set; } // 1 = Original, 2 = Retificação espontânea, 3 = Retificação a pedido
        public string NrRecibo { get; set; }
        public int SitEspecial { get; set; } // 0 = Não se aplica, 1 = Extinção, 2 = Fusão, 3 = Incorporação, 5 = Cisão
        public string NadaADeclarar { get; set; } // "1" = nada a declarar
        public int? FechamentoPP { get; set; } // 0 = sem movimento, 1 = com movimento
        public int? FechamentoMovOpFin { get; set; } // 0 = sem movimento, 1 = com movimento
        public int? FechamentoMovOpFinAnual { get; set; } // 0 = sem movimento, 1 = com movimento
        public int? ContasAReportarEntDecExterior { get; set; } // 0 = não há contas a reportar
        public List<DadosEntPatDecExterior> EntidadesPatrocinadas { get; set; }
    }

    [Serializable]
    public class DadosEntPatDecExterior
    {
        public string Giin { get; set; }
        public string Cnpj { get; set; }
        public int? ContasAReportar { get; set; } // 0 = não há contas a reportar
        public int? InCadPatrocinadoEncerrado { get; set; } // 1 = Sim
        public int? InGIINEncerrado { get; set; } // 1 = Sim
    }
}
