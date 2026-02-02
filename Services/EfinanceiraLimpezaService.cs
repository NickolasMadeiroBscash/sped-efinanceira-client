using System;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace ExemploAssinadorXML.Services
{
    public class EfinanceiraLimpezaService
    {
        public RespostaLimpezaDadosTeste LimparDadosTeste(string cnpjDeclarante, X509Certificate2 certificado)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(cnpjDeclarante))
                {
                    throw new ArgumentException("CNPJ do declarante não informado.", nameof(cnpjDeclarante));
                }

                // Remover formatação do CNPJ (apenas números)
                string cnpjLimpo = cnpjDeclarante.Replace(".", "").Replace("/", "").Replace("-", "").Trim();

                if (cnpjLimpo.Length != 14 || !System.Text.RegularExpressions.Regex.IsMatch(cnpjLimpo, @"^\d+$"))
                {
                    throw new ArgumentException("CNPJ inválido. Deve conter 14 dígitos numéricos.", nameof(cnpjDeclarante));
                }

                // URL do endpoint de limpeza (apenas ambiente de testes)
                string url = $"https://pre-efinanceira.receita.fazenda.gov.br/recepcao/limpezaDadosTesteProducaoRestrita?cnpjDeclarante={cnpjLimpo}";

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "DELETE";
                request.ClientCertificates.Add(certificado);
                request.Timeout = 60000; // 1 minuto
                request.ContentType = "application/xml";

                HttpWebResponse response = null;
                int codigoHttp = 0;
                string respostaBody = "";

                try
                {
                    response = (HttpWebResponse)request.GetResponse();
                    codigoHttp = (int)response.StatusCode;

                    using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                    {
                        respostaBody = reader.ReadToEnd();
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
                            respostaBody = reader.ReadToEnd();
                        }
                    }
                    else
                    {
                        throw new InvalidOperationException($"Erro na comunicação com a e-Financeira: {wex.Message}", wex);
                    }
                }
                finally
                {
                    response?.Close();
                }

                return new RespostaLimpezaDadosTeste
                {
                    CodigoHttp = codigoHttp,
                    Sucesso = codigoHttp == 200,
                    Mensagem = ObterMensagemPorCodigo(codigoHttp, respostaBody)
                };
            }
            catch (Exception ex)
            {
                return new RespostaLimpezaDadosTeste
                {
                    CodigoHttp = 0,
                    Sucesso = false,
                    Mensagem = $"Erro ao executar limpeza: {ex.Message}"
                };
            }
        }

        private static string ObterMensagemPorCodigo(int codigoHttp, string respostaBody)
        {
            switch (codigoHttp)
            {
                case 200:
                    return string.IsNullOrWhiteSpace(respostaBody) 
                        ? "Limpeza efetuada com sucesso." 
                        : $"Limpeza efetuada com sucesso. {respostaBody}";
                
                case 400:
                    return $"Parâmetro CNPJ não informado ou inválido. {respostaBody}";
                
                case 401:
                    return $"O certificado usado na conexão é inválido ou não possui permissão para limpar dados do CNPJ informado. {respostaBody}";
                
                case 404:
                    return $"URL informada incorretamente. Serviço não encontrado. {respostaBody}";
                
                case 405:
                    return $"Não foi utilizado o método HTTP DELETE na chamada ao endpoint. {respostaBody}";
                
                case 495:
                case 496:
                    return $"Não foi utilizado certificado válido na chamada ao endpoint. {respostaBody}";
                
                case 500:
                    return $"Erro interno na e-Financeira. {respostaBody}";
                
                default:
                    return string.IsNullOrWhiteSpace(respostaBody) 
                        ? $"Erro HTTP {codigoHttp}" 
                        : $"Erro HTTP {codigoHttp}: {respostaBody}";
            }
        }
    }

    public class RespostaLimpezaDadosTeste
    {
        public int CodigoHttp { get; set; }
        public bool Sucesso { get; set; }
        public string Mensagem { get; set; }
    }
}
