using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;
using ExemploAssinadorXML.Models;

namespace ExemploAssinadorXML.Services
{
    public class ProtocoloPersistenciaService
    {
        private static string GetArquivoProtocolos()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string pastaApp = Path.Combine(appData, "AssinadorEFinanceira");
            if (!Directory.Exists(pastaApp))
            {
                Directory.CreateDirectory(pastaApp);
            }
            return Path.Combine(pastaApp, "protocolos.xml");
        }

        public static void RegistrarProtocolo(TipoLote tipo, string arquivoCriptografado, string protocolo, string periodo = null, int quantidadeEventos = 0)
        {
            try
            {
                var lotes = CarregarLotes();
                
                // Verificar se já existe (evitar duplicatas)
                var loteExistente = lotes.FirstOrDefault(l => 
                    l.ArquivoCriptografado == arquivoCriptografado || 
                    (!string.IsNullOrEmpty(protocolo) && l.Protocolo == protocolo));
                
                if (loteExistente != null)
                {
                    // Atualizar protocolo se não existir
                    if (string.IsNullOrEmpty(loteExistente.Protocolo) && !string.IsNullOrEmpty(protocolo))
                    {
                        loteExistente.Protocolo = protocolo;
                        loteExistente.Status = "Enviado";
                        loteExistente.DataProcessamento = DateTime.Now;
                    }
                    // Atualizar quantidade de eventos se não estiver definida
                    if (loteExistente.QuantidadeEventos == 0 && quantidadeEventos > 0)
                    {
                        loteExistente.QuantidadeEventos = quantidadeEventos;
                    }
                }
                else
                {
                    // Adicionar novo lote
                    var novoLote = new LoteInfo
                    {
                        Tipo = tipo,
                        ArquivoCriptografado = arquivoCriptografado,
                        ArquivoAssinado = arquivoCriptografado?.Replace("-Criptografado.xml", "-ASSINADO.xml"),
                        ArquivoOriginal = arquivoCriptografado?.Replace("-Criptografado.xml", ".xml"),
                        Protocolo = protocolo,
                        Status = !string.IsNullOrEmpty(protocolo) ? "Enviado" : "Processado",
                        DataProcessamento = DateTime.Now,
                        Periodo = periodo,
                        QuantidadeEventos = quantidadeEventos
                    };
                    lotes.Add(novoLote);
                }

                SalvarLotes(lotes);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao registrar protocolo: {ex.Message}");
            }
        }

        public static List<LoteInfo> CarregarLotes()
        {
            try
            {
                string arquivo = GetArquivoProtocolos();
                if (!File.Exists(arquivo))
                {
                    return new List<LoteInfo>();
                }

                XmlSerializer serializer = new XmlSerializer(typeof(ListaProtocolos));
                using (FileStream stream = new FileStream(arquivo, FileMode.Open))
                {
                    var lista = (ListaProtocolos)serializer.Deserialize(stream);
                    return lista.Lotes ?? new List<LoteInfo>();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao carregar lotes: {ex.Message}");
                return new List<LoteInfo>();
            }
        }

        private static void SalvarLotes(List<LoteInfo> lotes)
        {
            try
            {
                string arquivo = GetArquivoProtocolos();
                XmlSerializer serializer = new XmlSerializer(typeof(ListaProtocolos));
                using (FileStream stream = new FileStream(arquivo, FileMode.Create))
                {
                    var lista = new ListaProtocolos { Lotes = lotes };
                    serializer.Serialize(stream, lista);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao salvar lotes: {ex.Message}");
            }
        }

        /// <summary>
        /// Conta a quantidade de eventos em um arquivo XML
        /// </summary>
        public static int ContarEventosNoXml(string caminhoArquivo)
        {
            try
            {
                if (!File.Exists(caminhoArquivo))
                {
                    return 0;
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(caminhoArquivo);

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("eFinanceira", doc.DocumentElement.NamespaceURI);

                // Tentar encontrar eventos na estrutura de lote
                XmlNodeList eventos = doc.SelectNodes("//eFinanceira:loteEventosAssincrono/eFinanceira:eventos/eFinanceira:evento", nsmgr);
                
                if (eventos.Count == 0)
                {
                    eventos = doc.SelectNodes("//eFinanceira:loteEventosAssincrono/eFinanceira:evento", nsmgr);
                }

                // Se não encontrou na estrutura de lote, buscar eventos diretamente
                if (eventos.Count == 0)
                {
                    eventos = doc.SelectNodes("//*[local-name()='evtAberturaeFinanceira' or local-name()='evtCadDeclarante' or local-name()='evtCadIntermediario' or local-name()='evtCadPatrocinado' or local-name()='evtExclusaoeFinanceira' or local-name()='evtExclusao' or local-name()='evtFechamentoeFinanceira' or local-name()='evtMovOpFin' or local-name()='evtMovPP']");
                }

                return eventos.Count;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao contar eventos no XML: {ex.Message}");
                return 0;
            }
        }
    }

    [XmlRoot("ListaProtocolos")]
    public class ListaProtocolos
    {
        [XmlElement("Lote")]
        public List<LoteInfo> Lotes { get; set; }
    }
}
