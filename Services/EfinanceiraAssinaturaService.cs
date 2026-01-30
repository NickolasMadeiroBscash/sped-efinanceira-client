using System;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Xml;

namespace ExemploAssinadorXML.Services
{
    public class EfinanceiraAssinaturaService
    {
        private const string SIGNATURE_METHOD = @"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
        private const string DIGEST_METHOD = @"http://www.w3.org/2001/04/xmlenc#sha256";
        private const string ATRIBUTO_ID = "id";
        private const int KEY_SIZE = 2048;
        private static bool algoritmoRegistrado = false;
        private static readonly object lockObject = new object();

        // Construtor estático para registrar o algoritmo de assinatura
        static EfinanceiraAssinaturaService()
        {
            RegistrarAlgoritmo();
        }

        // Garante que o algoritmo seja registrado (chamado no construtor estático e nos métodos públicos)
        private static void RegistrarAlgoritmo()
        {
            if (!algoritmoRegistrado)
            {
                lock (lockObject)
                {
                    if (!algoritmoRegistrado)
                    {
                        CryptoConfig.AddAlgorithm(typeof(RSAPKCS1SHA256SignatureDescription), SIGNATURE_METHOD);
                        algoritmoRegistrado = true;
                    }
                }
            }
        }

        public XmlDocument AssinarEventosDoArquivo(string caminhoArquivo, X509Certificate2 certificadoAssinatura)
        {
            // Garante que o algoritmo esteja registrado
            RegistrarAlgoritmo();

            // Carrega XML (sem PreserveWhitespace inicial, como no exemplo)
            XmlDocument arquivoXml = new XmlDocument();
            try
            {
                arquivoXml.Load(caminhoArquivo);
            }
            catch (Exception ex)
            {
                throw new Exception($"Não foi possível carregar XML indicado: {ex.Message}", ex);
            }

            // Verifica se XML possui eventos
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(arquivoXml.NameTable);
            nsmgr.AddNamespace("eFinanceira", arquivoXml.DocumentElement.NamespaceURI);

            // Primeiro tenta encontrar eventos na estrutura de lote (com elemento <eventos> intermediário)
            XmlNodeList eventos = arquivoXml.SelectNodes("//eFinanceira:loteEventosAssincrono/eFinanceira:eventos/eFinanceira:evento", nsmgr);
            bool estruturaLote = eventos.Count > 0;

            // Se não encontrou, tenta estrutura de lote sem elemento <eventos> intermediário
            if (!estruturaLote)
            {
                eventos = arquivoXml.SelectNodes("//eFinanceira:loteEventosAssincrono/eFinanceira:evento", nsmgr);
                estruturaLote = eventos.Count > 0;
            }

            // Se não encontrou na estrutura de lote, tenta encontrar eventos diretamente no eFinanceira
            if (!estruturaLote)
            {
                // Busca eventos diretamente dentro do elemento raiz eFinanceira
                // Usa namespace local para buscar eventos que podem ter namespaces diferentes
                eventos = arquivoXml.SelectNodes("//*[local-name()='evtAberturaeFinanceira' or local-name()='evtCadDeclarante' or local-name()='evtCadIntermediario' or local-name()='evtCadPatrocinado' or local-name()='evtExclusaoeFinanceira' or local-name()='evtExclusao' or local-name()='evtFechamentoeFinanceira' or local-name()='evtMovOpFin' or local-name()='evtMovPP']");
            }

            if (eventos.Count <= 0)
            {
                throw new Exception("Não encontrou eventos no arquivo selecionado.");
            }

            // Assina cada evento do arquivo            
            foreach (XmlNode node in eventos)
            {
                XmlDocument xmlDocEvento;
                string tagEventoParaAssinar;
                bool temEFinanceiraAninhado = false;

                if (estruturaLote)
                {
                    // Estrutura de lote: o evento está dentro de um nó <evento>
                    // O InnerXml pode conter outro <eFinanceira> aninhado com o evento real
                    xmlDocEvento = new XmlDocument(); 
                    xmlDocEvento.PreserveWhitespace = true;
                    xmlDocEvento.LoadXml(node.InnerXml);
                    
                    // Verifica se há um eFinanceira aninhado (com namespace diferente)
                    if (xmlDocEvento.DocumentElement != null && 
                        xmlDocEvento.DocumentElement.LocalName == "eFinanceira" &&
                        xmlDocEvento.DocumentElement.NamespaceURI != arquivoXml.DocumentElement.NamespaceURI)
                    {
                        temEFinanceiraAninhado = true;
                        // O documento já está correto, o evento está dentro do eFinanceira
                        // O método ObtemTagEventoAssinar vai encontrar usando Contains na OuterXml
                    }
                    
                    // Busca o evento real dentro do documento
                    tagEventoParaAssinar = ObtemTagEventoAssinar(xmlDocEvento);
                }
                else
                {
                    // Estrutura direta: o evento está diretamente no eFinanceira
                    xmlDocEvento = new XmlDocument();
                    xmlDocEvento.PreserveWhitespace = true;
                    xmlDocEvento.LoadXml(node.OuterXml);
                    tagEventoParaAssinar = ObtemTagEventoAssinar(xmlDocEvento);
                }

                if (string.IsNullOrWhiteSpace(tagEventoParaAssinar))
                {
                    throw new Exception($"Tipo Evento inválido para a e-Financeira: '{tagEventoParaAssinar}'");
                }

                XmlDocument xmlDocEventoAssinado = AssinarXmlEvento(xmlDocEvento, certificadoAssinatura, tagEventoParaAssinar);
                    
                if (xmlDocEventoAssinado == null)
                {
                    throw new Exception("Erro ao assinar evento.");
                }

                if (estruturaLote)
                {
                    // Para estrutura de lote, substitui o InnerXml do nó evento
                    if (temEFinanceiraAninhado)
                    {
                        // Preserva a estrutura aninhada com eFinanceira
                        // O documento assinado deve ter o eFinanceira como DocumentElement
                        node.InnerXml = xmlDocEventoAssinado.DocumentElement.OuterXml;
                    }
                    else
                    {
                        // Estrutura simples, usa InnerXml
                        node.InnerXml = xmlDocEventoAssinado.InnerXml;
                    }
                }
                else
                {
                    // Para estrutura direta, substitui o nó completo preservando atributos e assinatura
                    // Importa todos os nós filhos do DocumentElement (incluindo a assinatura)
                    XmlNode parentNode = node.ParentNode;
                    
                    // Remove o nó antigo
                    parentNode.RemoveChild(node);
                    
                    // Importa o elemento assinado (sem ser DocumentElement, apenas o elemento em si)
                    XmlNode importedNode = arquivoXml.ImportNode(xmlDocEventoAssinado.DocumentElement, true);
                    
                    // Adiciona o nó importado ao parent
                    parentNode.AppendChild(importedNode);
                }
            }

            return arquivoXml;
        }

        private string ObtemTagEventoAssinar(XmlDocument arquivo)
        {
            string tipoEvento = null;
            if (arquivo.OuterXml.Contains("evtCadDeclarante")) tipoEvento = "evtCadDeclarante";
            else if (arquivo.OuterXml.Contains("evtAberturaeFinanceira")) tipoEvento = "evtAberturaeFinanceira";
            else if (arquivo.OuterXml.Contains("evtCadIntermediario")) tipoEvento = "evtCadIntermediario";
            else if (arquivo.OuterXml.Contains("evtCadPatrocinado")) tipoEvento = "evtCadPatrocinado";
            else if (arquivo.OuterXml.Contains("evtExclusaoeFinanceira")) tipoEvento = "evtExclusaoeFinanceira";
            else if (arquivo.OuterXml.Contains("evtExclusao")) tipoEvento = "evtExclusao";
            else if (arquivo.OuterXml.Contains("evtFechamentoeFinanceira")) tipoEvento = "evtFechamentoeFinanceira";
            else if (arquivo.OuterXml.Contains("evtMovOpFin")) tipoEvento = "evtMovOpFin";
            else if (arquivo.OuterXml.Contains("evtMovPP")) tipoEvento = "evtMovPP";
            return tipoEvento;
        }

        public XmlDocument AssinarXmlEvento(XmlDocument xmlDocEvento, X509Certificate2 certificado, string tagEventoParaAssinar)
        {
            // Garante que o algoritmo esteja registrado
            RegistrarAlgoritmo();

            try
            {                
                XmlNodeList nodeParaAssinatura = xmlDocEvento.GetElementsByTagName(tagEventoParaAssinar);
                SignedXml signedXml = new SignedXml((XmlElement)nodeParaAssinatura[0]);
                signedXml.SignedInfo.SignatureMethod = SIGNATURE_METHOD;

                // Adicionando a chave privada para assinar o documento
                using (RSA chavePrivada = ObterChavePrivada(certificado))
                {
                    signedXml.SigningKey = chavePrivada;

                    Reference reference = new Reference("#" + nodeParaAssinatura[0].Attributes[ATRIBUTO_ID].Value);
                    reference.AddTransform(new XmlDsigEnvelopedSignatureTransform(false));
                    reference.AddTransform(new XmlDsigC14NTransform(false));
                    reference.DigestMethod = DIGEST_METHOD;
                    signedXml.AddReference(reference);

                    KeyInfo keyInfo = new KeyInfo();
                    keyInfo.AddClause(new KeyInfoX509Data(certificado));
                    signedXml.KeyInfo = keyInfo;

                    signedXml.ComputeSignature();


                    // Adiciona xml assinatura ao evento
                    XmlElement xmlElementAssinado = signedXml.GetXml();
                    XmlElement elementoEvento = (XmlElement)xmlDocEvento.GetElementsByTagName(tagEventoParaAssinar)[0];
                    
                    // Se o elemento é o DocumentElement, adiciona diretamente a ele
                    // Caso contrário, adiciona ao ParentNode
                    if (elementoEvento == xmlDocEvento.DocumentElement)
                    {
                        elementoEvento.AppendChild(xmlElementAssinado);
                    }
                    else
                    {
                        elementoEvento.ParentNode.AppendChild(xmlElementAssinado);
                    }

                    XmlDocument xmlAssinado = new XmlDocument();
                    xmlAssinado.PreserveWhitespace = true;
                    xmlAssinado.LoadXml(xmlDocEvento.OuterXml);

                    return xmlAssinado;
                }
            }
            catch (Exception ex)
            {                
                throw new Exception($"Falha ao assinar xml evento: {ex.Message}", ex);
            }
        }

        private RSA ObterChavePrivada(X509Certificate2 certificado)
        {
            RSACryptoServiceProvider privateKeyCertificado = (RSACryptoServiceProvider)certificado.PrivateKey;
            
            CspKeyContainerInfo enhCsp = new RSACryptoServiceProvider(KEY_SIZE).CspKeyContainerInfo;

            CngProvider provider = new CngProvider(enhCsp.ProviderName);
            using (CngKey key = CngKey.Open(privateKeyCertificado.CspKeyContainerInfo.KeyContainerName, provider))
            {
                RSA rsa = new RSACng(key)
                {
                    KeySize = KEY_SIZE
                };

                return rsa;
            }
        }
    }
}
