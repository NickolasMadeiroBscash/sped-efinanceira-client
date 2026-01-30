using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml;
using ExemploAssinadorXML.Models;

namespace ExemploAssinadorXML.Services
{
    public class EfinanceiraCriptografiaService
    {
        private const int AES_KEY_SIZE = 128;
        // AES/CBC/PKCS7Padding é equivalente a AES/CBC/PKCS5Padding no .NET
        // O .NET usa PKCS7Padding que é compatível com PKCS5Padding
        // RSA/ECB/PKCS1Padding - não usado diretamente, usamos RSA.Encrypt com RSAEncryptionPadding.Pkcs1

        public string CriptografarLote(string caminhoArquivoXml, string thumbprintCertificado)
        {
            try
            {
                // Ler e normalizar XML
                string xmlOriginal = LerENormalizarXml(caminhoArquivoXml);

                // Gerar chave AES e IV
                byte[] chaveAESBytes = GerarChaveAES();
                byte[] vetorAES = GerarVetorInicializacao();

                // Criptografar XML com AES
                byte[] xmlCriptografado = CriptografarComAES(xmlOriginal, chaveAESBytes, vetorAES);
                string xmlCriptografadoBase64 = Convert.ToBase64String(xmlCriptografado);

                // Buscar certificado do servidor
                X509Certificate2 certificadoServidor = BuscarCertificadoNoWindows(thumbprintCertificado);
                RSA rsaPublica = certificadoServidor.GetRSAPublicKey();

                // Concatenar chave AES + IV
                byte[] chaveConcat = ConcatenarBytes(chaveAESBytes, vetorAES);

                // Criptografar chave com RSA
                byte[] chaveCriptografada = rsaPublica.Encrypt(chaveConcat, RSAEncryptionPadding.Pkcs1);
                string chaveCriptografadaBase64 = Convert.ToBase64String(chaveCriptografada);

                // Gerar XML criptografado
                string caminhoArquivoSaida = GerarXmlCriptografado(
                    caminhoArquivoXml,
                    xmlCriptografadoBase64,
                    thumbprintCertificado,
                    chaveCriptografadaBase64
                );

                return caminhoArquivoSaida;
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao criptografar lote: {ex.Message}", ex);
            }
        }

        private string LerENormalizarXml(string caminho)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.PreserveWhitespace = true;
                doc.Load(caminho);

                using (StringWriter sw = new StringWriter())
                {
                    using (XmlTextWriter xw = new XmlTextWriter(sw))
                    {
                        xw.Formatting = Formatting.None;
                        xw.Indentation = 0;
                        doc.WriteTo(xw);
                    }
                    return sw.ToString();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao ler e normalizar XML: {ex.Message}", ex);
            }
        }

        private byte[] GerarChaveAES()
        {
            using (Aes aes = Aes.Create())
            {
                aes.KeySize = AES_KEY_SIZE;
                aes.GenerateKey();
                return aes.Key;
            }
        }

        private byte[] GerarVetorInicializacao()
        {
            byte[] iv = new byte[16];
            using (RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider())
            {
                rng.GetBytes(iv);
            }
            return iv;
        }

        private byte[] CriptografarComAES(string xml, byte[] chave, byte[] iv)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = chave;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                using (ICryptoTransform encryptor = aes.CreateEncryptor())
                {
                    byte[] xmlBytes = Encoding.UTF8.GetBytes(xml);
                    return encryptor.TransformFinalBlock(xmlBytes, 0, xmlBytes.Length);
                }
            }
        }

        private X509Certificate2 BuscarCertificadoNoWindows(string thumbprint)
        {
            string thumbprintNormalizado = thumbprint.Replace(" ", "").Replace("-", "").ToUpper();

            // Buscar no repositório CurrentUser\My
            X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            try
            {
                store.Open(OpenFlags.ReadOnly);
                foreach (X509Certificate2 cert in store.Certificates)
                {
                    string certThumbprint = cert.Thumbprint.Replace(" ", "").Replace("-", "").ToUpper();
                    if (certThumbprint == thumbprintNormalizado)
                    {
                        return cert;
                    }
                }
            }
            finally
            {
                store.Close();
            }

            // Buscar no repositório LocalMachine\My
            store = new X509Store(StoreName.My, StoreLocation.LocalMachine);
            try
            {
                store.Open(OpenFlags.ReadOnly);
                foreach (X509Certificate2 cert in store.Certificates)
                {
                    string certThumbprint = cert.Thumbprint.Replace(" ", "").Replace("-", "").ToUpper();
                    if (certThumbprint == thumbprintNormalizado)
                    {
                        return cert;
                    }
                }
            }
            finally
            {
                store.Close();
            }

            // Buscar no repositório LocalMachine\Root
            store = new X509Store(StoreName.Root, StoreLocation.LocalMachine);
            try
            {
                store.Open(OpenFlags.ReadOnly);
                foreach (X509Certificate2 cert in store.Certificates)
                {
                    string certThumbprint = cert.Thumbprint.Replace(" ", "").Replace("-", "").ToUpper();
                    if (certThumbprint == thumbprintNormalizado)
                    {
                        return cert;
                    }
                }
            }
            finally
            {
                store.Close();
            }

            throw new Exception($"Certificado com thumbprint '{thumbprint}' não encontrado no repositório do Windows.");
        }

        private byte[] ConcatenarBytes(byte[] a, byte[] b)
        {
            byte[] resultado = new byte[a.Length + b.Length];
            Buffer.BlockCopy(a, 0, resultado, 0, a.Length);
            Buffer.BlockCopy(b, 0, resultado, a.Length, b.Length);
            return resultado;
        }

        private string GerarXmlCriptografado(
            string caminhoArquivoOriginal,
            string xmlCriptografadoBase64,
            string thumbprintCertificado,
            string chaveCriptografadaBase64)
        {
            string idLote = Guid.NewGuid().ToString();
            string nomeArquivoOriginal = Path.GetFileName(caminhoArquivoOriginal);
            string nomeSemExtensao = Path.GetFileNameWithoutExtension(nomeArquivoOriginal);
            string extensao = Path.GetExtension(caminhoArquivoOriginal);
            string diretorio = Path.GetDirectoryName(caminhoArquivoOriginal);
            string caminhoSaida = Path.Combine(diretorio, nomeSemExtensao + "-Criptografado" + extensao);

            StringBuilder xml = new StringBuilder();
            xml.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            xml.AppendLine("<eFinanceira xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns=\"http://www.eFinanceira.gov.br/schemas/envioLoteCriptografado/v1_2_0\">");
            xml.AppendLine("  <loteCriptografado>");
            xml.AppendLine("    <id>" + idLote + "</id>");
            xml.AppendLine("    <idCertificado>" + thumbprintCertificado + "</idCertificado>");
            xml.AppendLine("    <chave>" + chaveCriptografadaBase64 + "</chave>");
            xml.AppendLine("    <lote>" + xmlCriptografadoBase64 + "</lote>");
            xml.AppendLine("  </loteCriptografado>");
            xml.AppendLine("</eFinanceira>");

            File.WriteAllText(caminhoSaida, xml.ToString(), Encoding.UTF8);

            return caminhoSaida;
        }
    }
}
