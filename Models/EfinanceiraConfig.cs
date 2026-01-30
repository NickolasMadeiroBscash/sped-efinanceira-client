using System;
using System.Xml.Serialization;

namespace ExemploAssinadorXML.Models
{
    [Serializable]
    [XmlRoot("EfinanceiraConfig")]
    public class EfinanceiraConfig
    {
        public string CnpjDeclarante { get; set; }
        public string CertThumbprint { get; set; }
        public string CertServidorThumbprint { get; set; }
        public EfinanceiraAmbiente Ambiente { get; set; }
        public bool TestEnvioHabilitado { get; set; }
        public bool ModoTeste { get; set; }
        public string Periodo { get; set; }
        public string DiretorioLotes { get; set; }
        
        // Configurações de Processamento
        public int PageSize { get; set; } = 500; // Tamanho da página para consultas ao banco
        public int EventoOffset { get; set; } = 1; // Controle de onde começar a gerar eventos
        public int OffsetRegistros { get; set; } = 0; // Pular registros iniciais (modo teste)
        public int? MaxLotes { get; set; } // Limitar quantidade de lotes gerados (null = ilimitado)
        public int EventosPorLote { get; set; } = 50; // Quantidade de eventos por lote (1 a 50, conforme manual)
        
        public string UrlTeste { get; set; }
        public string UrlProducao { get; set; }
        public string UrlConsultaTeste { get; set; }
        public string UrlConsultaProducao { get; set; }
    }

    public enum EfinanceiraAmbiente
    {
        TEST,
        PROD
    }
}
