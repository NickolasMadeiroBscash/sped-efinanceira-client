using System;
using System.Collections.Generic;

namespace ExemploAssinadorXML.Models
{
    public class StatusProcessamento
    {
        public int TotalLotes { get; set; }
        public int LotesProcessados { get; set; }
        public int LotesAssinados { get; set; }
        public int LotesCriptografados { get; set; }
        public int LotesEnviados { get; set; }
        public int LotesComErro { get; set; }
        public string EtapaAtual { get; set; }
        public string MensagemAtual { get; set; }
        public DateTime InicioProcessamento { get; set; }
        public TimeSpan TempoDecorrido { get; set; }
        public TimeSpan? TempoEstimadoRestante { get; set; }
        public List<string> ProtocolosEnviados { get; set; }

        public StatusProcessamento()
        {
            ProtocolosEnviados = new List<string>();
        }
    }

    public enum TipoLote
    {
        Abertura,
        Movimentacao,
        Fechamento
    }

    [Serializable]
    public class LoteInfo
    {
        public TipoLote Tipo { get; set; }
        public string ArquivoOriginal { get; set; }
        public string ArquivoAssinado { get; set; }
        public string ArquivoCriptografado { get; set; }
        public string Protocolo { get; set; }
        public string Status { get; set; }
        public string Erro { get; set; }
        public DateTime DataProcessamento { get; set; }
        public string Periodo { get; set; }
    }
}
