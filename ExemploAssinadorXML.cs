using System;
using System.Deployment.Internal.CodeSigning;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Windows.Forms;
using System.Xml;

namespace ExemploAssinadorXML
{
    public partial class ExemploAssinadorXML : Form
    {

        private const string SIGNATURE_METHOD = @"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
        private const string DIGEST_METHOD = @"http://www.w3.org/2001/04/xmlenc#sha256";
        private const string ATRIBUTO_ID = "id";
        
        public ExemploAssinadorXML()
        {
            InitializeComponent();
            CryptoConfig.AddAlgorithm(typeof(RSAPKCS1SHA256SignatureDescription), SIGNATURE_METHOD);
        }

        private void btnSelecionarArquivo_Click(object sender, EventArgs e)
        {
            ofdArquivoXML.Filter = "xml|*.xml";
            ofdArquivoXML.Title = "Abrir arquivo xml";
            if (ofdArquivoXML.ShowDialog() == DialogResult.OK)
            {
                txtCaminhoArquivo.Text = ofdArquivoXML.FileName;
            }
        }

        private void bntAssinar_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(txtCaminhoArquivo.Text.Trim()))
            {
                MessageBox.Show("Selecione um arquivo XML para assinar.");
                return;
            }

            if (!System.IO.File.Exists(txtCaminhoArquivo.Text))
            {
                MessageBox.Show("Arquivo não existe no disco.");
                return;
            }

            // Seleciona o certificado de assinatura e assina o arquivo
            X509Certificate2 cert = SelecionarCertificadoParaAssinatura();
            if (cert == null)
            {
                MessageBox.Show("Certificado não selecionado.");
                return;
            }


            // Verifica se certificado tem chave privada
            try
            {
                bool temChavePrivada = cert.HasPrivateKey;
                if (!temChavePrivada)
                {
                    MessageBox.Show("Certificado não possui chave privada.");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao ler chave privada do certificado : " + ex.Message);
                return;
            }


            // Assina eventos do arquivo
            XmlDocument xmlLoteAssinado = AssinarEventosDoArquivo(cert);


            if (xmlLoteAssinado != null)
            {
                string pathArquivoAssinado = txtCaminhoArquivo.Text.Trim().Replace(".xml", "-ASSINADO.xml");
                xmlLoteAssinado.Save(pathArquivoAssinado);
                MessageBox.Show("Arquivo Xml assinado : " + pathArquivoAssinado);
            }
            else
            {
                MessageBox.Show("Ocorreu erro ao assinar Xml");
            }
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



        private XmlDocument AssinarEventosDoArquivo(X509Certificate2 certificadoAssinatura)
        {
            // Carrega XML
            XmlDocument arquivoXml = new XmlDocument();
            try
            {
                arquivoXml.Load(txtCaminhoArquivo.Text.Trim());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Não foi possivel carregar XML indicado : " + ex.Message);
                return null;
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
                MessageBox.Show("Não encontrou eventos no arquivo selecionado.");
                return null;
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
                    MessageBox.Show($"Tipo Evento invalido para a e-Financeira : '{tagEventoParaAssinar}'");
                    return null;
                }

                XmlDocument xmlDocEventoAssinado = AssinarXmlEvento(xmlDocEvento, certificadoAssinatura, tagEventoParaAssinar);
                    
                if (xmlDocEventoAssinado == null)
                {
                    return null;
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




        public X509Certificate2 SelecionarCertificadoParaAssinatura()
        {
            X509Store store = new X509Store("MY", StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
            X509Certificate2Collection certs = store.Certificates;
            X509Certificate2Collection certsParaAssinatura = certs.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, false);
            X509Certificate2Collection certsParaSelecionar = X509Certificate2UI.SelectFromCollection(certsParaAssinatura, 
                "Certificado(s) Digital(is) disponível(is)", "Selecione o certificado digital para uso no aplicativo", X509SelectionFlag.SingleSelection);

            if (certsParaSelecionar.Count == 0)
            {
                string msgResultado = "Nenhum certificado digital foi selecionado ou o certificado selecionado está com problemas.";
                MessageBox.Show(msgResultado, "Advertência", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
            else
            {
                return certsParaSelecionar[0];                
            }            
        }



        public XmlDocument AssinarXmlEvento(XmlDocument xmlDocEvento, X509Certificate2 certificado, string tagEventoParaAssinar)
        {
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
                MessageBox.Show("Falha ao assinar xml evento : " + ex.Message);
                return null;
            }
        }



        private const int KEY_SIZE = 2048;

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


        private void btnValidarAssinatura_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCaminhoArquivo.Text.Trim()))
            {
                MessageBox.Show("Selecione um arquivo XML assinado.");
                return;
            }
            if (!System.IO.File.Exists(txtCaminhoArquivo.Text))
            {
                MessageBox.Show("Arquivo não existe no disco.");
                return;
            }

            XmlDocument xmlLoteAssinado = new XmlDocument();
            try
            {
                xmlLoteAssinado.Load(txtCaminhoArquivo.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Não foi possivel carregar o XML : " + ex.Message);
                return;
            }

            if (!VerificarExistenciaAssinatura(xmlLoteAssinado))
            {
                MessageBox.Show("O XML não está assinado.");
                return;
            }

            ValidarAssinatura(xmlLoteAssinado);
        }


        private bool VerificarExistenciaAssinatura(XmlDocument xml)
        {            
            XmlNodeList nodeList = xml.GetElementsByTagName("Signature", "*");
                        
            return nodeList != null && nodeList.Count > 0;
        }



        /// <summary>
        /// Valida assinatura de cada evento do xml
        /// </summary>        
        private void ValidarAssinatura(XmlDocument xmlAssinado)
        {
            try
            {
                // Selecionar as tags de assinatura.
                XmlNodeList nodeList = xmlAssinado.GetElementsByTagName("Signature", "*");

                // Verifica se o xml do evento foi assinado.
                if (nodeList.Count <= 0)
                {
                    MessageBox.Show("O evento não esta assinado.");
                    return;
                }

                bool assinaturasValidas = true;
                foreach (XmlNode assinatura in nodeList)
                {
                    XmlDocument evento = new XmlDocument() { PreserveWhitespace = true };
                    evento.LoadXml(assinatura.ParentNode.OuterXml);

                    SignedXml signedXml = new SignedXml(evento);
                    signedXml.LoadXml((XmlElement)assinatura);

                    // Carregar certificado do Xml num objeto X509Certificate2
                    string pubKey = signedXml.KeyInfo.GetXml().InnerText;
                    byte[] pubKeyBytes = Convert.FromBase64String(pubKey);
                    X509Certificate2 x509 = new X509Certificate2(pubKeyBytes);

                    // Verifica se a assinatura é rsa-sha256 
                    if (!signedXml.SignatureMethod.Equals("http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"))
                    {
                        MessageBox.Show("O documento " + ((Reference)signedXml.SignedInfo.References[0]).Uri + " não possui uma assinatura digital rsa-sha256.");
                    }
                    else
                    {
                        // Se a assinatura XML e o certificado digital são válidos
                        if (!signedXml.CheckSignature(x509, false))
                        {
                            // Se somente a assinatura é válida
                            if (signedXml.CheckSignature(x509, true))
                            {
                                // Assinatura é válida e o certificado é inválido.
                                MessageBox.Show("Certificado Digital X509 extraído do documento XML é inválido. id = " + ((Reference)signedXml.SignedInfo.References[0]).Uri);
                            }
                            else
                            {
                                MessageBox.Show("Assinatura Digital do documento XML é inválida. id = " + ((Reference)signedXml.SignedInfo.References[0]).Uri);
                            }
                            assinaturasValidas = false;
                        }
                    }
                }

                if (assinaturasValidas)
                {
                    MessageBox.Show("Assinatura Válida.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha ao verificar a assinatura do documento XML : " + ex.Message);
            }
        }





    }
}
