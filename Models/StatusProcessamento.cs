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
        
        // Campos para processamento completo
        public int StatusEtapa { get; set; } // 0=Aguardando, 1=Processando Abertura, 2=Abertura Enviada, 3=Processando Movimentação, 4=Enviando Movimentação, 5=Processando Fechamento, 6=Fechamento Enviado
        public bool AberturaFinalizada { get; set; }
        public int LotesMovimentacaoProcessados { get; set; }
        public int TotalLotesMovimentacao { get; set; }
        public TimeSpan? TempoMedioPorLote { get; set; }
        public List<DateTime> TemposLotesMovimentacao { get; set; }
        public bool ModoCompleto { get; set; } // Indica se está em modo de processamento completo

        public StatusProcessamento()
        {
            ProtocolosEnviados = new List<string>();
            TemposLotesMovimentacao = new List<DateTime>();
            StatusEtapa = 0; // Aguardando início
            ModoCompleto = false;
        }
    }

    public enum TipoLote
    {
        Abertura,
        Movimentacao,
        Fechamento,
        CadastroDeclarante
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
        public int QuantidadeEventos { get; set; }
    }
}
