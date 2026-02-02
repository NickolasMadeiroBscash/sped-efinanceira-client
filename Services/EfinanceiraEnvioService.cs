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
    public class EfinanceiraEnvioException : Exception
    {
        public EfinanceiraEnvioException(string message) : base(message)
        {
        }

        public EfinanceiraEnvioException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }

    public class EfinanceiraEnvioService
    {
        private const string EFINANCEIRA_XML_NAMESPACE = "http://www.eFinanceira.gov.br/schemas/envioLoteCriptografado/v1_2_0";

        public RespostaEnvioEfinanceira EnviarLote(string caminhoArquivoCriptografado, EfinanceiraConfig config, X509Certificate2 certificado)
        {
            try
            {
                string url = config.Ambiente == EfinanceiraAmbiente.PROD 
                    ? config.UrlProducao 
                    : config.UrlTeste;

                if (string.IsNullOrEmpty(url))
                {
                    throw new EfinanceiraEnvioException("URL de envio não configurada.");
                }

                string xmlCriptografado = File.ReadAllText(caminhoArquivoCriptografado, Encoding.UTF8);

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";
                request.ContentType = "application/xml";
                request.ClientCertificates.Add(certificado);
                request.Timeout = 300000; // 5 minutos

                byte[] data = Encoding.UTF8.GetBytes(xmlCriptografado);
                request.ContentLength = data.Length;

                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                    stream.Flush();
                }

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
                    // Capturar resposta mesmo em caso de erro HTTP
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
                        throw new EfinanceiraEnvioException($"Erro ao enviar lote: {wex.Message}", wex);
                    }
                }
                finally
                {
                    if (response != null)
                    {
                        response.Close();
                    }
                }

                RespostaEnvioEfinanceira resposta = new RespostaEnvioEfinanceira
                {
                    CodigoHttp = codigoHttp,
                    XmlCompleto = respostaXml
                };

                // Processar resposta apenas se for código HTTP válido (200, 201, 422)
                if (codigoHttp == 200 || codigoHttp == 201 || codigoHttp == 422)
                {
                    resposta = ProcessarRespostaEnvio(resposta);
                }
                else
                {
                    resposta.Descricao = $"Erro HTTP {codigoHttp}";
                }

                return resposta;
            }
            catch (WebException ex)
            {
                if (ex.Response != null)
                {
                    using (StreamReader reader = new StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8))
                    {
                        string erro = reader.ReadToEnd();
                        throw new EfinanceiraEnvioException($"Erro ao enviar lote: {erro}", ex);
                    }
                }
                throw new EfinanceiraEnvioException($"Erro ao enviar lote: {ex.Message}", ex);
            }
            catch (EfinanceiraEnvioException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new EfinanceiraEnvioException($"Erro ao enviar lote: {ex.Message}", ex);
            }
        }

        private static RespostaEnvioEfinanceira ProcessarRespostaEnvio(RespostaEnvioEfinanceira resposta)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(resposta.XmlCompleto);

                // Tentar diferentes namespaces e nomes de tags (baseado no código Java)
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("ns", EFINANCEIRA_XML_NAMESPACE);
                nsmgr.AddNamespace("", ""); // Namespace vazio também

                // Buscar cdResposta (código de resposta)
                XmlNode cdRespostaNode = doc.SelectSingleNode("//cdResposta") 
                    ?? doc.SelectSingleNode("//ns:cdResposta", nsmgr);
                if (cdRespostaNode != null)
                {
                    string cdRespostaStr = cdRespostaNode.InnerText.Trim();
                    if (int.TryParse(cdRespostaStr, out int codigo))
                    {
                        resposta.CodigoResposta = codigo;
                    }
                }

                // Buscar descResposta (descrição)
                XmlNode descRespostaNode = doc.SelectSingleNode("//descResposta")
                    ?? doc.SelectSingleNode("//ns:descResposta", nsmgr);
                if (descRespostaNode != null)
                {
                    resposta.Descricao = descRespostaNode.InnerText.Trim();
                }

                // Buscar protocoloEnvio (protocolo) - seguindo a lógica do Java
                // Primeiro tentar com GetElementsByTagName (sem namespace, como no Java)
                XmlNodeList protocoloList = doc.GetElementsByTagName("protocoloEnvio");
                if (protocoloList != null && protocoloList.Count > 0)
                {
                    resposta.Protocolo = protocoloList[0].InnerText.Trim();
                    System.Diagnostics.Debug.WriteLine($"Protocolo encontrado via GetElementsByTagName: {resposta.Protocolo}");
                }
                else
                {
                    // Tentar com XPath (com namespace)
                    XmlNode protocoloNode = doc.SelectSingleNode("//protocoloEnvio")
                        ?? doc.SelectSingleNode("//ns:protocoloEnvio", nsmgr)
                        ?? doc.SelectSingleNode("//protocolo")
                        ?? doc.SelectSingleNode("//ns:protocolo", nsmgr);
                    if (protocoloNode != null)
                    {
                        resposta.Protocolo = protocoloNode.InnerText.Trim();
                        System.Diagnostics.Debug.WriteLine($"Protocolo encontrado via XPath: {resposta.Protocolo}");
                    }
                    else
                    {
                        // Tentar buscar por numeroProtocolo (como no método extrairProtocolo do Java)
                        XmlNodeList numeroProtocoloList = doc.GetElementsByTagName("numeroProtocolo");
                        if (numeroProtocoloList != null && numeroProtocoloList.Count > 0)
                        {
                            resposta.Protocolo = numeroProtocoloList[0].InnerText.Trim();
                            System.Diagnostics.Debug.WriteLine($"Protocolo encontrado via numeroProtocolo: {resposta.Protocolo}");
                        }
                        else
                        {
                            // Log para debug se não encontrou protocolo
                            System.Diagnostics.Debug.WriteLine("AVISO: Protocolo não encontrado no XML de resposta. XML: " + resposta.XmlCompleto.Substring(0, Math.Min(500, resposta.XmlCompleto.Length)));
                        }
                    }
                }

                // Buscar ocorrências
                resposta.Ocorrencias = new List<OcorrenciaEnvio>();
                XmlNodeList ocorrenciasNodes = doc.SelectNodes("//ocorrencia") 
                    ?? doc.SelectNodes("//ns:ocorrencia", nsmgr);
                
                if (ocorrenciasNodes != null && ocorrenciasNodes.Count > 0)
                {
                    foreach (XmlNode ocorrNode in ocorrenciasNodes)
                    {
                        OcorrenciaEnvio ocorr = new OcorrenciaEnvio();
                        
                        XmlNode codigoNode = ocorrNode.SelectSingleNode("codigo")
                            ?? ocorrNode.SelectSingleNode("ns:codigo", nsmgr);
                        if (codigoNode != null)
                        {
                            ocorr.Codigo = codigoNode.InnerText.Trim();
                        }

                        XmlNode descricaoNode = ocorrNode.SelectSingleNode("descricao")
                            ?? ocorrNode.SelectSingleNode("ns:descricao", nsmgr);
                        if (descricaoNode != null)
                        {
                            ocorr.Descricao = descricaoNode.InnerText.Trim();
                        }

                        XmlNode tipoNode = ocorrNode.SelectSingleNode("tipo")
                            ?? ocorrNode.SelectSingleNode("ns:tipo", nsmgr);
                        if (tipoNode != null)
                        {
                            ocorr.Tipo = tipoNode.InnerText.Trim();
                        }

                        resposta.Ocorrencias.Add(ocorr);
                    }
                }
            }
            catch (Exception ex)
            {
                // Não lançar exceção, apenas logar aviso (como no código Java)
                System.Diagnostics.Debug.WriteLine($"Aviso: Não foi possível parsear o XML de resposta: {ex.Message}");
            }

            return resposta;
        }
    }

    public class RespostaEnvioEfinanceira
    {
        public int CodigoHttp { get; set; }
        public int? CodigoResposta { get; set; }
        public string Descricao { get; set; }
        public string Protocolo { get; set; }
        public List<OcorrenciaEnvio> Ocorrencias { get; set; }
        public string XmlCompleto { get; set; }
        
        // Propriedades legadas para compatibilidade
        public string Codigo => CodigoResposta?.ToString() ?? "";
        public string Mensagem => Descricao ?? "";
    }

    public class OcorrenciaEnvio
    {
        public string Codigo { get; set; }
        public string Descricao { get; set; }
        public string Tipo { get; set; }
    }
}
