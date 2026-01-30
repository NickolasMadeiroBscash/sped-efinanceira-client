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

                // Buscar cdResposta (código de resposta)
                XmlNodeList cdRespostaList = doc.GetElementsByTagName("cdResposta");
                if (cdRespostaList.Count > 0)
                {
                    string cdRespostaStr = cdRespostaList[0].InnerText.Trim();
                    if (int.TryParse(cdRespostaStr, out int codigoResposta))
                    {
                        resposta.CodigoResposta = codigoResposta;
                    }
                }

                // Buscar descResposta (descrição da resposta)
                XmlNodeList descRespostaList = doc.GetElementsByTagName("descResposta");
                if (descRespostaList.Count > 0)
                {
                    resposta.Descricao = descRespostaList[0].InnerText.Trim();
                }

                // Buscar ocorrências
                resposta.Ocorrencias = new List<OcorrenciaEfinanceira>();
                XmlNodeList ocorrenciasList = doc.GetElementsByTagName("ocorrencia");
                for (int i = 0; i < ocorrenciasList.Count; i++)
                {
                    XmlElement ocorrenciaElement = (XmlElement)ocorrenciasList[i];
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
}
