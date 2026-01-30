using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml;
using ExemploAssinadorXML.Models;

namespace ExemploAssinadorXML.Services
{
    public class EfinanceiraConsultaService
    {
        public RespostaConsultaEfinanceira ConsultarProtocolo(string protocolo, EfinanceiraConfig config, X509Certificate2 certificado)
        {
            try
            {
                string url = config.Ambiente == EfinanceiraAmbiente.PROD
                    ? config.UrlConsultaProducao
                    : config.UrlConsultaTeste;

                if (string.IsNullOrEmpty(url))
                {
                    throw new Exception("URL de consulta não configurada.");
                }

                url = url.TrimEnd('/') + "/" + protocolo;

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "GET";
                request.ClientCertificates.Add(certificado);
                request.Timeout = 60000; // 1 minuto

                HttpWebResponse response = null;
                int codigoHttp = 0;
                string respostaXml = "";

                try
                {
                    response = (HttpWebResponse)request.GetResponse();
                    codigoHttp = (int)response.StatusCode;

                    using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                    {
                        respostaXml = reader.ReadToEnd();
                    }
                }
                catch (WebException wex)
                {
                    if (wex.Response != null)
                    {
                        response = (HttpWebResponse)wex.Response;
                        codigoHttp = (int)response.StatusCode;

                        using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                        {
                            respostaXml = reader.ReadToEnd();
                        }
                    }
                    else
                    {
                        throw new Exception($"Erro de rede ao consultar protocolo: {wex.Message}", wex);
                    }
                }
                finally
                {
                    if (response != null)
                    {
                        response.Close();
                    }
                }

                RespostaConsultaEfinanceira resposta = new RespostaConsultaEfinanceira
                {
                    CodigoHttp = codigoHttp,
                    XmlCompleto = respostaXml
                };

                if (codigoHttp == 200 || codigoHttp == 404 || codigoHttp == 422)
                {
                    resposta = ProcessarRespostaConsulta(resposta);
                }

                return resposta;
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao consultar protocolo: {ex.Message}", ex);
            }
        }

        private RespostaConsultaEfinanceira ProcessarRespostaConsulta(RespostaConsultaEfinanceira resposta)
        {
            try
            {
                if (string.IsNullOrEmpty(resposta.XmlCompleto))
                {
                    return resposta;
                }

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(resposta.XmlCompleto);

                // Namespace manager
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("ef", "http://www.eFinanceira.gov.br/schemas/retornoLoteEventosAssincrono/v1_0_0");
                nsmgr.AddNamespace("evt", "http://www.eFinanceira.gov.br/schemas/retornoEvento/v1_3_0");

                // Buscar cdResposta (código de resposta do lote)
                XmlNodeList cdRespostaList = doc.GetElementsByTagName("cdResposta");
                if (cdRespostaList.Count > 0)
                {
                    string cdRespostaStr = cdRespostaList[0].InnerText.Trim();
                    if (int.TryParse(cdRespostaStr, out int codigoResposta))
                    {
                        resposta.CodigoResposta = codigoResposta;
                    }
                }

                // Buscar descResposta (descrição da resposta do lote)
                XmlNodeList descRespostaList = doc.GetElementsByTagName("descResposta");
                if (descRespostaList.Count > 0)
                {
                    resposta.Descricao = descRespostaList[0].InnerText.Trim();
                }

                // Buscar dados de recepção do lote
                XmlNodeList dadosRecepcaoList = doc.GetElementsByTagName("dadosRecepcaoLote");
                if (dadosRecepcaoList.Count > 0)
                {
                    XmlElement dadosRecepcao = (XmlElement)dadosRecepcaoList[0];
                    
                    XmlNodeList dhRecepcaoList = dadosRecepcao.GetElementsByTagName("dhRecepcao");
                    if (dhRecepcaoList.Count > 0 && DateTime.TryParse(dhRecepcaoList[0].InnerText.Trim(), out DateTime dtRecepcao))
                    {
                        resposta.DataRecepcao = dtRecepcao;
                    }
                    
                    XmlNodeList versaoRecepcaoList = dadosRecepcao.GetElementsByTagName("versaoAplicativoRecepcao");
                    if (versaoRecepcaoList.Count > 0)
                    {
                        resposta.VersaoAplicativoRecepcao = versaoRecepcaoList[0].InnerText.Trim();
                    }
                    
                    XmlNodeList protocoloEnvioList = dadosRecepcao.GetElementsByTagName("protocoloEnvio");
                    if (protocoloEnvioList.Count > 0)
                    {
                        resposta.ProtocoloEnvio = protocoloEnvioList[0].InnerText.Trim();
                    }
                }

                // Buscar dados de processamento do lote
                XmlNodeList dadosProcessamentoList = doc.GetElementsByTagName("dadosProcessamentoLote");
                if (dadosProcessamentoList.Count > 0)
                {
                    XmlElement dadosProcessamento = (XmlElement)dadosProcessamentoList[0];
                    
                    XmlNodeList dhProcessamentoList = dadosProcessamento.GetElementsByTagName("dhProcessamento");
                    if (dhProcessamentoList.Count > 0 && DateTime.TryParse(dhProcessamentoList[0].InnerText.Trim(), out DateTime dtProcessamento))
                    {
                        resposta.DataProcessamento = dtProcessamento;
                    }
                    
                    XmlNodeList versaoProcessamentoList = dadosProcessamento.GetElementsByTagName("versaoAplicativoProcessamentoLote");
                    if (versaoProcessamentoList.Count > 0)
                    {
                        resposta.VersaoAplicativoProcessamento = versaoProcessamentoList[0].InnerText.Trim();
                    }
                }

                // Buscar ocorrências do lote (gerais)
                resposta.Ocorrencias = new List<OcorrenciaEfinanceira>();
                XmlNodeList ocorrenciasList = doc.GetElementsByTagName("ocorrencia");
                for (int i = 0; i < ocorrenciasList.Count; i++)
                {
                    XmlElement ocorrenciaElement = (XmlElement)ocorrenciasList[i];
                    
                    // Verificar se é ocorrência do lote (não está dentro de um evento)
                    if (ocorrenciaElement.ParentNode != null && 
                        ocorrenciaElement.ParentNode.Name != "dadosRegistroOcorrenciaEvento")
                    {
                        OcorrenciaEfinanceira ocorrencia = new OcorrenciaEfinanceira();

                        XmlNodeList codigoList = ocorrenciaElement.GetElementsByTagName("codigo");
                        if (codigoList.Count > 0)
                        {
                            ocorrencia.Codigo = codigoList[0].InnerText.Trim();
                        }

                        XmlNodeList descricaoList = ocorrenciaElement.GetElementsByTagName("descricao");
                        if (descricaoList.Count > 0)
                        {
                            ocorrencia.Descricao = descricaoList[0].InnerText.Trim();
                        }

                        XmlNodeList tipoList = ocorrenciaElement.GetElementsByTagName("tipo");
                        if (tipoList.Count > 0)
                        {
                            ocorrencia.Tipo = tipoList[0].InnerText.Trim();
                        }

                        resposta.Ocorrencias.Add(ocorrencia);
                    }
                }

                // Buscar detalhes de cada evento individual
                resposta.DetalhesEventos = new List<DetalheEventoConsulta>();
                XmlNodeList eventosList = doc.GetElementsByTagName("evento");
                for (int i = 0; i < eventosList.Count; i++)
                {
                    XmlElement eventoElement = (XmlElement)eventosList[i];
                    DetalheEventoConsulta detalheEvento = new DetalheEventoConsulta();
                    
                    // ID do evento
                    if (eventoElement.HasAttribute("id"))
                    {
                        detalheEvento.IdEvento = eventoElement.GetAttribute("id");
                    }
                    
                    // Buscar retornoEvento dentro do evento
                    XmlNodeList retornoEventoList = eventoElement.GetElementsByTagName("retornoEvento");
                    if (retornoEventoList.Count > 0)
                    {
                        XmlElement retornoEvento = (XmlElement)retornoEventoList[0];
                        
                        // Dados de recepção do evento
                        XmlNodeList dadosRecepcaoEventoList = retornoEvento.GetElementsByTagName("dadosRecepcaoEvento");
                        if (dadosRecepcaoEventoList.Count > 0)
                        {
                            XmlElement dadosRecepcaoEvento = (XmlElement)dadosRecepcaoEventoList[0];
                            
                            XmlNodeList idEventoList = dadosRecepcaoEvento.GetElementsByTagName("idEvento");
                            if (idEventoList.Count > 0)
                            {
                                detalheEvento.IdEvento = idEventoList[0].InnerText.Trim();
                            }
                            
                            XmlNodeList tipoEventoList = dadosRecepcaoEvento.GetElementsByTagName("tipoEvento");
                            if (tipoEventoList.Count > 0)
                            {
                                detalheEvento.TipoEvento = tipoEventoList[0].InnerText.Trim();
                            }
                            
                            XmlNodeList hashList = dadosRecepcaoEvento.GetElementsByTagName("hash");
                            if (hashList.Count > 0)
                            {
                                detalheEvento.Hash = hashList[0].InnerText.Trim();
                            }
                            
                            XmlNodeList dhRecepcaoEventoList = dadosRecepcaoEvento.GetElementsByTagName("dhRecepcao");
                            if (dhRecepcaoEventoList.Count > 0 && DateTime.TryParse(dhRecepcaoEventoList[0].InnerText.Trim(), out DateTime dtRecepcaoEvento))
                            {
                                detalheEvento.DataRecepcao = dtRecepcaoEvento;
                            }
                            
                            XmlNodeList dhProcessamentoEventoList = dadosRecepcaoEvento.GetElementsByTagName("dhProcessamento");
                            if (dhProcessamentoEventoList.Count > 0 && DateTime.TryParse(dhProcessamentoEventoList[0].InnerText.Trim(), out DateTime dtProcessamentoEvento))
                            {
                                detalheEvento.DataProcessamento = dtProcessamentoEvento;
                            }
                        }
                        
                        // Status do evento
                        XmlNodeList statusEventoList = retornoEvento.GetElementsByTagName("status");
                        if (statusEventoList.Count > 0)
                        {
                            XmlElement statusEvento = (XmlElement)statusEventoList[0];
                            
                            XmlNodeList cdRetornoList = statusEvento.GetElementsByTagName("cdRetorno");
                            if (cdRetornoList.Count > 0)
                            {
                                detalheEvento.CodigoRetorno = cdRetornoList[0].InnerText.Trim();
                            }
                            
                            XmlNodeList descRetornoList = statusEvento.GetElementsByTagName("descRetorno");
                            if (descRetornoList.Count > 0)
                            {
                                detalheEvento.DescricaoRetorno = descRetornoList[0].InnerText.Trim();
                            }
                            
                            // Ocorrências do evento
                            XmlNodeList dadosRegistroList = statusEvento.GetElementsByTagName("dadosRegistroOcorrenciaEvento");
                            if (dadosRegistroList.Count > 0)
                            {
                                XmlElement dadosRegistro = (XmlElement)dadosRegistroList[0];
                                
                                // Buscar todas as ocorrências dentro de dadosRegistroOcorrenciaEvento
                                // Pode haver múltiplos elementos <ocorrencias> ou <ocorrencia>
                                XmlNodeList ocorrenciasEventoList = dadosRegistro.GetElementsByTagName("ocorrencias");
                                
                                // Se não encontrou com "ocorrencias", tenta buscar "ocorrencia" (singular)
                                if (ocorrenciasEventoList.Count == 0)
                                {
                                    ocorrenciasEventoList = dadosRegistro.GetElementsByTagName("ocorrencia");
                                }
                                
                                // Processar cada ocorrência encontrada
                                for (int j = 0; j < ocorrenciasEventoList.Count; j++)
                                {
                                    XmlElement ocorrenciaEvento = (XmlElement)ocorrenciasEventoList[j];
                                    OcorrenciaEfinanceira ocorrencia = new OcorrenciaEfinanceira
                                    {
                                        IdEvento = detalheEvento.IdEvento
                                    };
                                    
                                    // Buscar código, descrição e tipo dentro do elemento ocorrência
                                    XmlNodeList codigoOcorrList = ocorrenciaEvento.GetElementsByTagName("codigo");
                                    if (codigoOcorrList.Count > 0)
                                    {
                                        ocorrencia.Codigo = codigoOcorrList[0].InnerText.Trim();
                                    }
                                    
                                    XmlNodeList descricaoOcorrList = ocorrenciaEvento.GetElementsByTagName("descricao");
                                    if (descricaoOcorrList.Count > 0)
                                    {
                                        ocorrencia.Descricao = descricaoOcorrList[0].InnerText.Trim();
                                    }
                                    
                                    XmlNodeList tipoOcorrList = ocorrenciaEvento.GetElementsByTagName("tipo");
                                    if (tipoOcorrList.Count > 0)
                                    {
                                        ocorrencia.Tipo = tipoOcorrList[0].InnerText.Trim();
                                    }
                                    
                                    // Só adiciona se tiver pelo menos código ou descrição
                                    if (!string.IsNullOrEmpty(ocorrencia.Codigo) || !string.IsNullOrEmpty(ocorrencia.Descricao))
                                    {
                                        detalheEvento.Ocorrencias.Add(ocorrencia);
                                    }
                                }
                            }
                        }
                    }
                    
                    resposta.DetalhesEventos.Add(detalheEvento);
                }

                return resposta;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Aviso: Não foi possível parsear o XML de resposta: {ex.Message}");
                return resposta;
            }
        }
    }

    public class RespostaConsultaEfinanceira
    {
        public int CodigoHttp { get; set; }
        public int CodigoResposta { get; set; }
        public string Descricao { get; set; }
        public List<OcorrenciaEfinanceira> Ocorrencias { get; set; } = new List<OcorrenciaEfinanceira>();
        public List<DetalheEventoConsulta> DetalhesEventos { get; set; } = new List<DetalheEventoConsulta>();
        public string ProtocoloEnvio { get; set; }
        public DateTime? DataRecepcao { get; set; }
        public DateTime? DataProcessamento { get; set; }
        public string VersaoAplicativoRecepcao { get; set; }
        public string VersaoAplicativoProcessamento { get; set; }
        public string XmlCompleto { get; set; }

        // Propriedades para compatibilidade com código existente
        public string Codigo => CodigoResposta.ToString();
        public string Mensagem => Descricao;
    }

    public class OcorrenciaEfinanceira
    {
        public string Codigo { get; set; }
        public string Descricao { get; set; }
        public string Tipo { get; set; }

        // Propriedade para compatibilidade com código existente
        public string Mensagem => Descricao;
        public string IdEvento { get; set; }
    }

    public class DetalheEventoConsulta
    {
        public string IdEvento { get; set; }
        public string TipoEvento { get; set; }
        public string CodigoRetorno { get; set; }
        public string DescricaoRetorno { get; set; }
        public string Hash { get; set; }
        public DateTime? DataRecepcao { get; set; }
        public DateTime? DataProcessamento { get; set; }
        public List<OcorrenciaEfinanceira> Ocorrencias { get; set; } = new List<OcorrenciaEfinanceira>();
    }
}
