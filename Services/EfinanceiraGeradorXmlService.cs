using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using ExemploAssinadorXML.Models;

namespace ExemploAssinadorXML.Services
{
    public class EfinanceiraGeradorXmlService
    {
        private const string NAMESPACE_LOTE = "http://www.eFinanceira.gov.br/schemas/envioLoteEventosAssincrono/v1_0_0";
        private const string NAMESPACE_EVENTO_ABERTURA = "http://www.eFinanceira.gov.br/schemas/evtAberturaeFinanceira/v1_2_1";
        private const string NAMESPACE_EVENTO_FECHAMENTO = "http://www.eFinanceira.gov.br/schemas/evtFechamentoeFinanceira/v1_3_0";
        private const string NAMESPACE_EVENTO_CADASTRO = "http://www.eFinanceira.gov.br/schemas/evtCadDeclarante/v1_2_0";
        private const string VERSAO_APLICATIVO = "00000000000000000001";
        
        // Constantes para literais XML
        private const string ENCODING_UTF8 = "utf-8";
        private const string ELEMENTO_EFINANCEIRA = "eFinanceira";
        private const string ATTR_XMLNS_XSI = "xmlns:xsi";
        private const string XMLNS_XSI_VALUE = "http://www.w3.org/2001/XMLSchema-instance";
        private const string ATTR_XMLNS_XSD = "xmlns:xsd";
        private const string XMLNS_XSD_VALUE = "http://www.w3.org/2001/XMLSchema";
        private const string ELEMENTO_LOTE_EVENTOS_ASSINCRONO = "loteEventosAssincrono";
        private const string ELEMENTO_CNPJ_DECLARANTE = "cnpjDeclarante";
        private const string ELEMENTO_EVENTOS = "eventos";
        private const string ELEMENTO_EVENTO = "evento";
        private const string ELEMENTO_IDE_EVENTO = "ideEvento";
        private const string ELEMENTO_IND_RETIFICACAO = "indRetificacao";
        private const string ELEMENTO_TP_AMB = "tpAmb";
        private const string ELEMENTO_APLIC_EMI = "aplicEmi";
        private const string ELEMENTO_VER_APLIC = "verAplic";
        private const string ELEMENTO_NUMERO = "Numero";

        public string GerarXmlAbertura(DadosAbertura dados, string diretorioSaida)
        {
            if (!Directory.Exists(diretorioSaida))
            {
                Directory.CreateDirectory(diretorioSaida);
            }

            string idEvento = GerarIdEvento();
            string nomeArquivo = $"abertura-efinanceira-{dados.DtInicio.Replace("-", "")}-{dados.DtFim.Replace("-", "")}_{DateTime.Now:yyyyMMdd_HHmmss}.xml";
            string caminhoCompleto = Path.Combine(diretorioSaida, nomeArquivo);

            XmlDocument doc = new XmlDocument();
            XmlDeclaration decl = doc.CreateXmlDeclaration("1.0", ENCODING_UTF8, "no");
            doc.AppendChild(decl);

            // Elemento raiz eFinanceira
            XmlElement eFinanceiraRaiz = doc.CreateElement(ELEMENTO_EFINANCEIRA, NAMESPACE_LOTE);
            eFinanceiraRaiz.SetAttribute(ATTR_XMLNS_XSI, XMLNS_XSI_VALUE);
            eFinanceiraRaiz.SetAttribute(ATTR_XMLNS_XSD, XMLNS_XSD_VALUE);
            doc.AppendChild(eFinanceiraRaiz);

            // loteEventosAssincrono (deve estar no mesmo namespace)
            XmlElement loteEventosAssincrono = doc.CreateElement(ELEMENTO_LOTE_EVENTOS_ASSINCRONO, NAMESPACE_LOTE);
            eFinanceiraRaiz.AppendChild(loteEventosAssincrono);

            // cnpjDeclarante (deve estar no mesmo namespace)
            XmlElement cnpjDeclaranteLote = doc.CreateElement(ELEMENTO_CNPJ_DECLARANTE, NAMESPACE_LOTE);
            cnpjDeclaranteLote.InnerText = dados.CnpjDeclarante;
            loteEventosAssincrono.AppendChild(cnpjDeclaranteLote);

            // eventos (deve estar no mesmo namespace)
            XmlElement eventos = doc.CreateElement(ELEMENTO_EVENTOS, NAMESPACE_LOTE);
            loteEventosAssincrono.AppendChild(eventos);

            // evento (deve estar no mesmo namespace)
            XmlElement evento = doc.CreateElement(ELEMENTO_EVENTO, NAMESPACE_LOTE);
            evento.SetAttribute("id", "ID0");
            eventos.AppendChild(evento);

            // eFinanceira interno
            XmlElement eFinanceiraInterno = doc.CreateElement(ELEMENTO_EFINANCEIRA, NAMESPACE_EVENTO_ABERTURA);
            evento.AppendChild(eFinanceiraInterno);

            // evtAberturaeFinanceira (deve estar no namespace do evento)
            XmlElement evtAberturaeFinanceira = doc.CreateElement("evtAberturaeFinanceira", NAMESPACE_EVENTO_ABERTURA);
            evtAberturaeFinanceira.SetAttribute("id", idEvento);
            eFinanceiraInterno.AppendChild(evtAberturaeFinanceira);

            // ideEvento
            CriarIdeEvento(doc, evtAberturaeFinanceira, dados);
            
            // ideDeclarante
            CriarIdeDeclarante(doc, evtAberturaeFinanceira, dados.CnpjDeclarante);
            
            // infoAbertura
            CriarInfoAbertura(doc, evtAberturaeFinanceira, dados);

            // AberturaMovOpFin (se necessário)
            if (dados.IndicarMovOpFin || dados.ResponsavelRMF != null || dados.RespeFin != null || dados.RepresLegal != null)
            {
                CriarAberturaMovOpFin(doc, evtAberturaeFinanceira, dados);
            }

            // AberturaPP (se necessário)
            if (dados.TiposEmpresaPP != null && dados.TiposEmpresaPP.Length > 0)
            {
                CriarAberturaPP(doc, evtAberturaeFinanceira, dados.TiposEmpresaPP);
            }

            // Salvar
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  ",
                Encoding = Encoding.UTF8,
                OmitXmlDeclaration = false
            };

            using (XmlWriter writer = XmlWriter.Create(caminhoCompleto, settings))
            {
                doc.Save(writer);
            }

            return caminhoCompleto;
        }

        public string GerarXmlFechamento(DadosFechamento dados, string diretorioSaida)
        {
            if (!Directory.Exists(diretorioSaida))
            {
                Directory.CreateDirectory(diretorioSaida);
            }

            string idEvento = GerarIdEvento();
            string nomeArquivo = $"fechamento-efinanceira-{dados.DtInicio.Replace("-", "")}-{dados.DtFim.Replace("-", "")}_{DateTime.Now:yyyyMMdd_HHmmss}.xml";
            string caminhoCompleto = Path.Combine(diretorioSaida, nomeArquivo);

            XmlDocument doc = new XmlDocument();
            XmlDeclaration decl = doc.CreateXmlDeclaration("1.0", ENCODING_UTF8, "no");
            doc.AppendChild(decl);

            // Elemento raiz eFinanceira
            XmlElement eFinanceiraRaiz = doc.CreateElement(ELEMENTO_EFINANCEIRA, NAMESPACE_LOTE);
            eFinanceiraRaiz.SetAttribute(ATTR_XMLNS_XSI, XMLNS_XSI_VALUE);
            eFinanceiraRaiz.SetAttribute(ATTR_XMLNS_XSD, XMLNS_XSD_VALUE);
            doc.AppendChild(eFinanceiraRaiz);

            // loteEventosAssincrono (deve estar no mesmo namespace)
            XmlElement loteEventosAssincrono = doc.CreateElement(ELEMENTO_LOTE_EVENTOS_ASSINCRONO, NAMESPACE_LOTE);
            eFinanceiraRaiz.AppendChild(loteEventosAssincrono);

            // cnpjDeclarante (deve estar no mesmo namespace)
            XmlElement cnpjDeclaranteLote = doc.CreateElement(ELEMENTO_CNPJ_DECLARANTE, NAMESPACE_LOTE);
            cnpjDeclaranteLote.InnerText = dados.CnpjDeclarante;
            loteEventosAssincrono.AppendChild(cnpjDeclaranteLote);

            // eventos (deve estar no mesmo namespace)
            XmlElement eventos = doc.CreateElement(ELEMENTO_EVENTOS, NAMESPACE_LOTE);
            loteEventosAssincrono.AppendChild(eventos);

            // evento (deve estar no mesmo namespace)
            XmlElement evento = doc.CreateElement(ELEMENTO_EVENTO, NAMESPACE_LOTE);
            evento.SetAttribute("id", "ID0");
            eventos.AppendChild(evento);

            // eFinanceira interno
            XmlElement eFinanceiraInterno = doc.CreateElement(ELEMENTO_EFINANCEIRA, NAMESPACE_EVENTO_FECHAMENTO);
            evento.AppendChild(eFinanceiraInterno);

            // evtFechamentoeFinanceira (deve estar no namespace do evento)
            XmlElement evtFechamentoeFinanceira = doc.CreateElement("evtFechamentoeFinanceira", NAMESPACE_EVENTO_FECHAMENTO);
            evtFechamentoeFinanceira.SetAttribute("id", idEvento);
            eFinanceiraInterno.AppendChild(evtFechamentoeFinanceira);

            // ideEvento
            CriarIdeEventoFechamento(doc, evtFechamentoeFinanceira, dados);
            
            // ideDeclarante
            CriarIdeDeclarante(doc, evtFechamentoeFinanceira, dados.CnpjDeclarante);
            
            // infoFechamento
            CriarInfoFechamento(doc, evtFechamentoeFinanceira, dados);

            // FechamentoPP (se necessário)
            if (dados.FechamentoPP.HasValue)
            {
                CriarFechamentoPP(doc, evtFechamentoeFinanceira, dados.FechamentoPP.Value);
            }

            // FechamentoMovOpFin (se necessário)
            if (dados.FechamentoMovOpFin.HasValue)
            {
                CriarFechamentoMovOpFin(doc, evtFechamentoeFinanceira, dados);
            }

            // FechamentoMovOpFinAnual (se necessário)
            if (dados.FechamentoMovOpFinAnual.HasValue)
            {
                CriarFechamentoMovOpFinAnual(doc, evtFechamentoeFinanceira, dados.FechamentoMovOpFinAnual.Value);
            }

            // Salvar
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  ",
                Encoding = Encoding.UTF8,
                OmitXmlDeclaration = false
            };

            using (XmlWriter writer = XmlWriter.Create(caminhoCompleto, settings))
            {
                doc.Save(writer);
            }

            return caminhoCompleto;
        }

        private static void CriarIdeEvento(XmlDocument doc, XmlElement pai, DadosAbertura dados)
        {
            XmlElement ideEvento = doc.CreateElement(ELEMENTO_IDE_EVENTO, NAMESPACE_EVENTO_ABERTURA);
            pai.AppendChild(ideEvento);

            XmlElement indRetificacao = doc.CreateElement(ELEMENTO_IND_RETIFICACAO, NAMESPACE_EVENTO_ABERTURA);
            indRetificacao.InnerText = dados.IndRetificacao.ToString();
            ideEvento.AppendChild(indRetificacao);

            if ((dados.IndRetificacao == 2 || dados.IndRetificacao == 3) && !string.IsNullOrEmpty(dados.NrRecibo))
            {
                XmlElement nrRecibo = doc.CreateElement("nrRecibo", NAMESPACE_EVENTO_ABERTURA);
                nrRecibo.InnerText = dados.NrRecibo;
                ideEvento.AppendChild(nrRecibo);
            }

            XmlElement tpAmb = doc.CreateElement(ELEMENTO_TP_AMB, NAMESPACE_EVENTO_ABERTURA);
            tpAmb.InnerText = dados.TipoAmbiente.ToString();
            ideEvento.AppendChild(tpAmb);

            XmlElement aplicEmi = doc.CreateElement(ELEMENTO_APLIC_EMI, NAMESPACE_EVENTO_ABERTURA);
            aplicEmi.InnerText = dados.AplicacaoEmissora.ToString();
            ideEvento.AppendChild(aplicEmi);

            XmlElement verAplic = doc.CreateElement(ELEMENTO_VER_APLIC, NAMESPACE_EVENTO_ABERTURA);
            verAplic.InnerText = VERSAO_APLICATIVO;
            ideEvento.AppendChild(verAplic);
        }

        private static void CriarIdeEventoFechamento(XmlDocument doc, XmlElement pai, DadosFechamento dados)
        {
            string namespaceEvento = pai.NamespaceURI;
            
            XmlElement ideEvento = doc.CreateElement(ELEMENTO_IDE_EVENTO, namespaceEvento);
            pai.AppendChild(ideEvento);

            XmlElement indRetificacao = doc.CreateElement(ELEMENTO_IND_RETIFICACAO, namespaceEvento);
            indRetificacao.InnerText = dados.IndRetificacao.ToString();
            ideEvento.AppendChild(indRetificacao);

            if ((dados.IndRetificacao == 2 || dados.IndRetificacao == 3) && !string.IsNullOrEmpty(dados.NrRecibo))
            {
                XmlElement nrRecibo = doc.CreateElement("nrRecibo", namespaceEvento);
                nrRecibo.InnerText = dados.NrRecibo;
                ideEvento.AppendChild(nrRecibo);
            }

            XmlElement tpAmb = doc.CreateElement(ELEMENTO_TP_AMB, namespaceEvento);
            tpAmb.InnerText = dados.TipoAmbiente.ToString();
            ideEvento.AppendChild(tpAmb);

            XmlElement aplicEmi = doc.CreateElement(ELEMENTO_APLIC_EMI, namespaceEvento);
            aplicEmi.InnerText = dados.AplicacaoEmissora.ToString();
            ideEvento.AppendChild(aplicEmi);

            XmlElement verAplic = doc.CreateElement(ELEMENTO_VER_APLIC, namespaceEvento);
            verAplic.InnerText = VERSAO_APLICATIVO;
            ideEvento.AppendChild(verAplic);
        }

        private static void CriarIdeDeclarante(XmlDocument doc, XmlElement pai, string cnpj)
        {
            // Verificar o namespace do elemento pai para usar o mesmo namespace
            string namespaceEvento = pai.NamespaceURI;
            
            XmlElement ideDeclarante = doc.CreateElement("ideDeclarante", namespaceEvento);
            pai.AppendChild(ideDeclarante);

            XmlElement cnpjDeclarante = doc.CreateElement("cnpjDeclarante", namespaceEvento);
            cnpjDeclarante.InnerText = cnpj;
            ideDeclarante.AppendChild(cnpjDeclarante);
        }

        private static void CriarInfoAbertura(XmlDocument doc, XmlElement pai, DadosAbertura dados)
        {
            XmlElement infoAbertura = doc.CreateElement("infoAbertura", NAMESPACE_EVENTO_ABERTURA);
            pai.AppendChild(infoAbertura);

            XmlElement dtInicio = doc.CreateElement("dtInicio", NAMESPACE_EVENTO_ABERTURA);
            dtInicio.InnerText = dados.DtInicio;
            infoAbertura.AppendChild(dtInicio);

            XmlElement dtFim = doc.CreateElement("dtFim", NAMESPACE_EVENTO_ABERTURA);
            dtFim.InnerText = dados.DtFim;
            infoAbertura.AppendChild(dtFim);
        }

        private void CriarAberturaMovOpFin(XmlDocument doc, XmlElement pai, DadosAbertura dados)
        {
            XmlElement aberturaMovOpFin = doc.CreateElement("AberturaMovOpFin", NAMESPACE_EVENTO_ABERTURA);
            pai.AppendChild(aberturaMovOpFin);

            if (dados.ResponsavelRMF != null)
            {
                CriarResponsavelRMF(doc, aberturaMovOpFin, dados.ResponsavelRMF);
            }

            if (dados.RespeFin != null)
            {
                CriarRespeFin(doc, aberturaMovOpFin, dados.RespeFin);
            }

            if (dados.RepresLegal != null)
            {
                CriarRepresLegal(doc, aberturaMovOpFin, dados.RepresLegal);
            }
        }

        private static void CriarResponsavelRMF(XmlDocument doc, XmlElement pai, DadosResponsavelRMF dados)
        {
            string namespaceEvento = pai.NamespaceURI;
            
            XmlElement responsavelRMF = doc.CreateElement("ResponsavelRMF", namespaceEvento);
            pai.AppendChild(responsavelRMF);

            XmlElement cnpj = doc.CreateElement("CNPJ", namespaceEvento);
            cnpj.InnerText = dados.Cnpj;
            responsavelRMF.AppendChild(cnpj);

            XmlElement cpf = doc.CreateElement("CPF", namespaceEvento);
            cpf.InnerText = dados.Cpf;
            responsavelRMF.AppendChild(cpf);

            XmlElement nome = doc.CreateElement("Nome", namespaceEvento);
            nome.InnerText = dados.Nome;
            responsavelRMF.AppendChild(nome);

            XmlElement setor = doc.CreateElement("Setor", namespaceEvento);
            setor.InnerText = dados.Setor;
            responsavelRMF.AppendChild(setor);

            // Telefone
            XmlElement telefone = doc.CreateElement("Telefone", namespaceEvento);
            responsavelRMF.AppendChild(telefone);

            XmlElement ddd = doc.CreateElement("DDD", namespaceEvento);
            ddd.InnerText = dados.TelefoneDDD;
            telefone.AppendChild(ddd);

            XmlElement numero = doc.CreateElement(ELEMENTO_NUMERO, namespaceEvento);
            numero.InnerText = dados.TelefoneNumero;
            telefone.AppendChild(numero);

            if (!string.IsNullOrEmpty(dados.TelefoneRamal))
            {
                XmlElement ramal = doc.CreateElement("Ramal", namespaceEvento);
                ramal.InnerText = dados.TelefoneRamal;
                telefone.AppendChild(ramal);
            }

            // Endereco
            XmlElement endereco = doc.CreateElement("Endereco", namespaceEvento);
            responsavelRMF.AppendChild(endereco);

            XmlElement logradouro = doc.CreateElement("Logradouro", namespaceEvento);
            logradouro.InnerText = dados.EnderecoLogradouro;
            endereco.AppendChild(logradouro);

            XmlElement numeroEndereco = doc.CreateElement(ELEMENTO_NUMERO, namespaceEvento);
            numeroEndereco.InnerText = dados.EnderecoNumero;
            endereco.AppendChild(numeroEndereco);

            if (!string.IsNullOrEmpty(dados.EnderecoComplemento))
            {
                XmlElement complemento = doc.CreateElement("Complemento", namespaceEvento);
                complemento.InnerText = dados.EnderecoComplemento;
                endereco.AppendChild(complemento);
            }

            XmlElement bairro = doc.CreateElement("Bairro", namespaceEvento);
            bairro.InnerText = dados.EnderecoBairro;
            endereco.AppendChild(bairro);

            XmlElement cep = doc.CreateElement("CEP", namespaceEvento);
            cep.InnerText = dados.EnderecoCEP;
            endereco.AppendChild(cep);

            XmlElement municipio = doc.CreateElement("Municipio", namespaceEvento);
            municipio.InnerText = dados.EnderecoMunicipio;
            endereco.AppendChild(municipio);

            XmlElement uf = doc.CreateElement("UF", namespaceEvento);
            uf.InnerText = dados.EnderecoUF;
            endereco.AppendChild(uf);
        }

        private static void CriarRespeFin(XmlDocument doc, XmlElement pai, DadosRespeFin dados)
        {
            string namespaceEvento = pai.NamespaceURI;
            
            XmlElement respeFin = doc.CreateElement("RespeFin", namespaceEvento);
            pai.AppendChild(respeFin);

            XmlElement cpf = doc.CreateElement("CPF", namespaceEvento);
            cpf.InnerText = dados.Cpf;
            respeFin.AppendChild(cpf);

            XmlElement nome = doc.CreateElement("Nome", namespaceEvento);
            nome.InnerText = dados.Nome;
            respeFin.AppendChild(nome);

            XmlElement setor = doc.CreateElement("Setor", namespaceEvento);
            setor.InnerText = dados.Setor;
            respeFin.AppendChild(setor);

            // Telefone
            XmlElement telefone = doc.CreateElement("Telefone", namespaceEvento);
            respeFin.AppendChild(telefone);

            XmlElement ddd = doc.CreateElement("DDD", namespaceEvento);
            ddd.InnerText = dados.TelefoneDDD;
            telefone.AppendChild(ddd);

            XmlElement numero = doc.CreateElement(ELEMENTO_NUMERO, namespaceEvento);
            numero.InnerText = dados.TelefoneNumero;
            telefone.AppendChild(numero);

            if (!string.IsNullOrEmpty(dados.TelefoneRamal))
            {
                XmlElement ramal = doc.CreateElement("Ramal", namespaceEvento);
                ramal.InnerText = dados.TelefoneRamal;
                telefone.AppendChild(ramal);
            }

            // Endereco
            XmlElement endereco = doc.CreateElement("Endereco", namespaceEvento);
            respeFin.AppendChild(endereco);

            XmlElement logradouro = doc.CreateElement("Logradouro", namespaceEvento);
            logradouro.InnerText = dados.EnderecoLogradouro;
            endereco.AppendChild(logradouro);

            XmlElement numeroEndereco = doc.CreateElement(ELEMENTO_NUMERO, namespaceEvento);
            numeroEndereco.InnerText = dados.EnderecoNumero;
            endereco.AppendChild(numeroEndereco);

            if (!string.IsNullOrEmpty(dados.EnderecoComplemento))
            {
                XmlElement complemento = doc.CreateElement("Complemento", namespaceEvento);
                complemento.InnerText = dados.EnderecoComplemento;
                endereco.AppendChild(complemento);
            }

            XmlElement bairro = doc.CreateElement("Bairro", namespaceEvento);
            bairro.InnerText = dados.EnderecoBairro;
            endereco.AppendChild(bairro);

            XmlElement cep = doc.CreateElement("CEP", namespaceEvento);
            cep.InnerText = dados.EnderecoCEP;
            endereco.AppendChild(cep);

            XmlElement municipio = doc.CreateElement("Municipio", namespaceEvento);
            municipio.InnerText = dados.EnderecoMunicipio;
            endereco.AppendChild(municipio);

            XmlElement uf = doc.CreateElement("UF", namespaceEvento);
            uf.InnerText = dados.EnderecoUF;
            endereco.AppendChild(uf);

            XmlElement email = doc.CreateElement("Email", namespaceEvento);
            email.InnerText = dados.Email;
            respeFin.AppendChild(email);
        }

        private static void CriarRepresLegal(XmlDocument doc, XmlElement pai, DadosRepresLegal dados)
        {
            string namespaceEvento = pai.NamespaceURI;
            
            XmlElement represLegal = doc.CreateElement("RepresLegal", namespaceEvento);
            pai.AppendChild(represLegal);

            XmlElement cpf = doc.CreateElement("CPF", namespaceEvento);
            cpf.InnerText = dados.Cpf;
            represLegal.AppendChild(cpf);

            XmlElement setor = doc.CreateElement("Setor", namespaceEvento);
            setor.InnerText = dados.Setor;
            represLegal.AppendChild(setor);

            // Telefone
            XmlElement telefone = doc.CreateElement("Telefone", namespaceEvento);
            represLegal.AppendChild(telefone);

            XmlElement ddd = doc.CreateElement("DDD", namespaceEvento);
            ddd.InnerText = dados.TelefoneDDD;
            telefone.AppendChild(ddd);

            XmlElement numero = doc.CreateElement(ELEMENTO_NUMERO, namespaceEvento);
            numero.InnerText = dados.TelefoneNumero;
            telefone.AppendChild(numero);

            if (!string.IsNullOrEmpty(dados.TelefoneRamal))
            {
                XmlElement ramal = doc.CreateElement("Ramal", namespaceEvento);
                ramal.InnerText = dados.TelefoneRamal;
                telefone.AppendChild(ramal);
            }
        }

        private static void CriarAberturaPP(XmlDocument doc, XmlElement pai, string[] tiposEmpresaPP)
        {
            XmlElement aberturaPP = doc.CreateElement("AberturaPP", NAMESPACE_EVENTO_ABERTURA);
            pai.AppendChild(aberturaPP);

            foreach (string tpPrevPriv in tiposEmpresaPP)
            {
                XmlElement tpEmpresa = doc.CreateElement("tpEmpresa", NAMESPACE_EVENTO_ABERTURA);
                aberturaPP.AppendChild(tpEmpresa);

                XmlElement tpPrevPrivElement = doc.CreateElement("tpPrevPriv", NAMESPACE_EVENTO_ABERTURA);
                tpPrevPrivElement.InnerText = tpPrevPriv;
                tpEmpresa.AppendChild(tpPrevPrivElement);
            }
        }

        private static void CriarInfoFechamento(XmlDocument doc, XmlElement pai, DadosFechamento dados)
        {
            string namespaceEvento = pai.NamespaceURI;
            
            XmlElement infoFechamento = doc.CreateElement("infoFechamento", namespaceEvento);
            pai.AppendChild(infoFechamento);

            XmlElement dtInicio = doc.CreateElement("dtInicio", namespaceEvento);
            dtInicio.InnerText = dados.DtInicio;
            infoFechamento.AppendChild(dtInicio);

            XmlElement dtFim = doc.CreateElement("dtFim", namespaceEvento);
            dtFim.InnerText = dados.DtFim;
            infoFechamento.AppendChild(dtFim);

            XmlElement sitEspecial = doc.CreateElement("sitEspecial", namespaceEvento);
            sitEspecial.InnerText = dados.SitEspecial.ToString();
            infoFechamento.AppendChild(sitEspecial);

            if (!string.IsNullOrEmpty(dados.NadaADeclarar) && dados.NadaADeclarar == "1")
            {
                XmlElement nadaADeclarar = doc.CreateElement("nadaADeclarar", namespaceEvento);
                nadaADeclarar.InnerText = "1";
                infoFechamento.AppendChild(nadaADeclarar);
            }
        }

        private static void CriarFechamentoPP(XmlDocument doc, XmlElement pai, int fechamentoPP)
        {
            string namespaceEvento = pai.NamespaceURI;
            
            XmlElement fechamentoPPElement = doc.CreateElement("FechamentoPP", namespaceEvento);
            pai.AppendChild(fechamentoPPElement);

            XmlElement fechamentoPPValue = doc.CreateElement("FechamentoPP", namespaceEvento);
            fechamentoPPValue.InnerText = fechamentoPP.ToString();
            fechamentoPPElement.AppendChild(fechamentoPPValue);
        }

        private static void CriarFechamentoMovOpFin(XmlDocument doc, XmlElement pai, DadosFechamento dados)
        {
            string namespaceEvento = pai.NamespaceURI;
            
            XmlElement fechamentoMovOpFinElement = doc.CreateElement("FechamentoMovOpFin", namespaceEvento);
            pai.AppendChild(fechamentoMovOpFinElement);

            XmlElement fechamentoMovOpFinValue = doc.CreateElement("FechamentoMovOpFin", namespaceEvento);
            fechamentoMovOpFinValue.InnerText = dados.FechamentoMovOpFin.Value.ToString();
            fechamentoMovOpFinElement.AppendChild(fechamentoMovOpFinValue);

            if (dados.ContasAReportarEntDecExterior.HasValue)
            {
                XmlElement entDecExterior = doc.CreateElement("EntDecExterior", namespaceEvento);
                fechamentoMovOpFinElement.AppendChild(entDecExterior);

                XmlElement contasAReportar = doc.CreateElement("ContasAReportar", namespaceEvento);
                contasAReportar.InnerText = dados.ContasAReportarEntDecExterior.Value.ToString();
                entDecExterior.AppendChild(contasAReportar);
            }

            if (dados.EntidadesPatrocinadas != null && dados.EntidadesPatrocinadas.Count > 0)
            {
                foreach (var entidade in dados.EntidadesPatrocinadas)
                {
                    XmlElement entPatDecExterior = doc.CreateElement("EntPatDecExterior", namespaceEvento);
                    fechamentoMovOpFinElement.AppendChild(entPatDecExterior);

                    XmlElement giin = doc.CreateElement("GIIN", namespaceEvento);
                    giin.InnerText = entidade.Giin;
                    entPatDecExterior.AppendChild(giin);

                    XmlElement cnpj = doc.CreateElement("CNPJ", namespaceEvento);
                    cnpj.InnerText = entidade.Cnpj;
                    entPatDecExterior.AppendChild(cnpj);

                    if (entidade.ContasAReportar.HasValue)
                    {
                        XmlElement contasAReportar = doc.CreateElement("ContasAReportar", namespaceEvento);
                        contasAReportar.InnerText = entidade.ContasAReportar.Value.ToString();
                        entPatDecExterior.AppendChild(contasAReportar);
                    }

                    if (entidade.InCadPatrocinadoEncerrado.HasValue)
                    {
                        XmlElement inCadPatrocinadoEncerrado = doc.CreateElement("inCadPatrocinadoEncerrado", namespaceEvento);
                        inCadPatrocinadoEncerrado.InnerText = entidade.InCadPatrocinadoEncerrado.Value.ToString();
                        entPatDecExterior.AppendChild(inCadPatrocinadoEncerrado);
                    }

                    if (entidade.InGIINEncerrado.HasValue)
                    {
                        XmlElement inGIINEncerrado = doc.CreateElement("inGIINEncerrado", namespaceEvento);
                        inGIINEncerrado.InnerText = entidade.InGIINEncerrado.Value.ToString();
                        entPatDecExterior.AppendChild(inGIINEncerrado);
                    }
                }
            }
        }

        private static void CriarFechamentoMovOpFinAnual(XmlDocument doc, XmlElement pai, int fechamentoMovOpFinAnual)
        {
            string namespaceEvento = pai.NamespaceURI;
            
            XmlElement fechamentoMovOpFinAnualElement = doc.CreateElement("FechamentoMovOpFinAnual", namespaceEvento);
            pai.AppendChild(fechamentoMovOpFinAnualElement);

            XmlElement fechamentoMovOpFinAnualValue = doc.CreateElement("FechamentoMovOpFinAnual", namespaceEvento);
            fechamentoMovOpFinAnualValue.InnerText = fechamentoMovOpFinAnual.ToString();
            fechamentoMovOpFinAnualElement.AppendChild(fechamentoMovOpFinAnualValue);
        }

        /// <summary>
        /// Gera lote de movimentação financeira a partir de lista de pessoas com contas
        /// </summary>
        public string GerarXmlMovimentacao(List<DadosPessoaConta> pessoas, string cnpjDeclarante, string periodoStr, 
            int tipoAmbiente, int eventoOffset, string diretorioSaida)
        {
            if (!Directory.Exists(diretorioSaida))
            {
                Directory.CreateDirectory(diretorioSaida);
            }

            const string NAMESPACE_EVENTO_MOV = "http://www.eFinanceira.gov.br/schemas/evtMovOpFin/v1_2_1";

            // Gerar um único lote com as pessoas fornecidas (quantidade configurável, máximo 50 eventos conforme manual)
            string nomeArquivo = $"movimentacao-efinanceira-{periodoStr}_{DateTime.Now:yyyyMMdd_HHmmss}.xml";
            string caminhoCompleto = Path.Combine(diretorioSaida, nomeArquivo);

            XmlDocument doc = new XmlDocument();
            XmlDeclaration decl = doc.CreateXmlDeclaration("1.0", ENCODING_UTF8, "no");
            doc.AppendChild(decl);

            // Elemento raiz eFinanceira
            XmlElement eFinanceiraRaiz = doc.CreateElement(ELEMENTO_EFINANCEIRA, NAMESPACE_LOTE);
            eFinanceiraRaiz.SetAttribute(ATTR_XMLNS_XSI, XMLNS_XSI_VALUE);
            eFinanceiraRaiz.SetAttribute(ATTR_XMLNS_XSD, XMLNS_XSD_VALUE);
            doc.AppendChild(eFinanceiraRaiz);

            // loteEventosAssincrono (deve estar no mesmo namespace)
            XmlElement loteEventosAssincrono = doc.CreateElement(ELEMENTO_LOTE_EVENTOS_ASSINCRONO, NAMESPACE_LOTE);
            eFinanceiraRaiz.AppendChild(loteEventosAssincrono);

            // cnpjDeclarante (deve estar no mesmo namespace)
            XmlElement cnpjDeclaranteLote = doc.CreateElement(ELEMENTO_CNPJ_DECLARANTE, NAMESPACE_LOTE);
            cnpjDeclaranteLote.InnerText = cnpjDeclarante;
            loteEventosAssincrono.AppendChild(cnpjDeclaranteLote);

            // eventos (deve estar no mesmo namespace)
            XmlElement eventos = doc.CreateElement(ELEMENTO_EVENTOS, NAMESPACE_LOTE);
            loteEventosAssincrono.AppendChild(eventos);

            // Criar eventos para cada pessoa (aplicar offset se necessário)
            // A quantidade de pessoas já foi limitada pelo ProcessamentoForm conforme EventosPorLote configurado
            int indiceEvento = 0;
            
            // Validar eventoOffset para evitar índice fora do intervalo
            if (eventoOffset < 0)
            {
                eventoOffset = 0;
            }
            if (eventoOffset >= pessoas.Count)
            {
                throw new ArgumentException($"EventoOffset ({eventoOffset}) é maior ou igual ao número de pessoas ({pessoas.Count}). " +
                    $"Ajuste o EventoOffset para um valor menor.");
            }
            
            // Processar todas as pessoas fornecidas (já limitadas pelo ProcessamentoForm)
            int maxEventos = pessoas.Count - eventoOffset;
            for (int j = eventoOffset; j < pessoas.Count && indiceEvento < maxEventos; j++)
            {
                var pessoa = pessoas[j];
                string idEvento = FormatIdEvento(pessoa.IdPessoa);
                
                // evento (deve estar no mesmo namespace)
                XmlElement evento = doc.CreateElement("evento", NAMESPACE_LOTE);
                evento.SetAttribute("id", $"ID{indiceEvento}");
                eventos.AppendChild(evento);

                // eFinanceira interno
                XmlElement eFinanceiraInterno = doc.CreateElement(ELEMENTO_EFINANCEIRA, NAMESPACE_EVENTO_MOV);
                evento.AppendChild(eFinanceiraInterno);

                // evtMovOpFin
                CriarEventoMovOpFin(doc, eFinanceiraInterno, pessoa, cnpjDeclarante, periodoStr, idEvento, tipoAmbiente);
                indiceEvento++;
            }

            // Salvar
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  ",
                Encoding = Encoding.UTF8,
                OmitXmlDeclaration = false
            };

            using (XmlWriter writer = XmlWriter.Create(caminhoCompleto, settings))
            {
                doc.Save(writer);
            }

            return caminhoCompleto;
        }

        private void CriarEventoMovOpFin(XmlDocument doc, XmlElement eFinanceiraInterno, DadosPessoaConta pessoa, 
            string cnpjDeclarante, string periodoStr, string idEvento, int tipoAmbiente)
        {
            const string NAMESPACE_EVENTO_MOV = "http://www.eFinanceira.gov.br/schemas/evtMovOpFin/v1_2_1";
            
            // evtMovOpFin (deve estar no namespace do evento)
            XmlElement evtMovOpFin = doc.CreateElement("evtMovOpFin", NAMESPACE_EVENTO_MOV);
            evtMovOpFin.SetAttribute("id", idEvento);
            eFinanceiraInterno.AppendChild(evtMovOpFin);

            // ideEvento
            XmlElement ideEvento = doc.CreateElement(ELEMENTO_IDE_EVENTO, NAMESPACE_EVENTO_MOV);
            evtMovOpFin.AppendChild(ideEvento);

            XmlElement indRetificacao = doc.CreateElement(ELEMENTO_IND_RETIFICACAO, NAMESPACE_EVENTO_MOV);
            indRetificacao.InnerText = "1"; // Original
            ideEvento.AppendChild(indRetificacao);

            XmlElement tpAmb = doc.CreateElement(ELEMENTO_TP_AMB, NAMESPACE_EVENTO_MOV);
            tpAmb.InnerText = tipoAmbiente.ToString();
            ideEvento.AppendChild(tpAmb);

            XmlElement aplicEmi = doc.CreateElement(ELEMENTO_APLIC_EMI, NAMESPACE_EVENTO_MOV);
            aplicEmi.InnerText = "1"; // Aplicativo da empresa
            ideEvento.AppendChild(aplicEmi);

            XmlElement verAplic = doc.CreateElement(ELEMENTO_VER_APLIC, NAMESPACE_EVENTO_MOV);
            verAplic.InnerText = VERSAO_APLICATIVO;
            ideEvento.AppendChild(verAplic);

            // ideDeclarante
            CriarIdeDeclarante(doc, evtMovOpFin, cnpjDeclarante);

            // ideDeclarado
            XmlElement ideDeclarado = doc.CreateElement("ideDeclarado", NAMESPACE_EVENTO_MOV);
            evtMovOpFin.AppendChild(ideDeclarado);

            XmlElement tpNI = doc.CreateElement("tpNI", NAMESPACE_EVENTO_MOV);
            tpNI.InnerText = "1"; // CPF
            ideDeclarado.AppendChild(tpNI);

            XmlElement nIDeclarado = doc.CreateElement("NIDeclarado", NAMESPACE_EVENTO_MOV);
            nIDeclarado.InnerText = NormalizarCpf(pessoa.Cpf);
            ideDeclarado.AppendChild(nIDeclarado);

            XmlElement nomeDeclarado = doc.CreateElement("NomeDeclarado", NAMESPACE_EVENTO_MOV);
            nomeDeclarado.InnerText = pessoa.Nome ?? "";
            ideDeclarado.AppendChild(nomeDeclarado);

            // EnderecoLivre
            string enderecoLivre = ConstruirEnderecoLivre(pessoa);
            if (!string.IsNullOrEmpty(enderecoLivre))
            {
                XmlElement enderecoLivreElem = doc.CreateElement("EnderecoLivre", NAMESPACE_EVENTO_MOV);
                enderecoLivreElem.InnerText = enderecoLivre;
                ideDeclarado.AppendChild(enderecoLivreElem);
            }

            // PaisEndereco
            XmlElement paisEndereco = doc.CreateElement("PaisEndereco", NAMESPACE_EVENTO_MOV);
            ideDeclarado.AppendChild(paisEndereco);
            XmlElement paisEnd = doc.CreateElement("Pais", NAMESPACE_EVENTO_MOV);
            paisEnd.InnerText = "BR";
            paisEndereco.AppendChild(paisEnd);

            // paisResid
            XmlElement paisResid = doc.CreateElement("paisResid", NAMESPACE_EVENTO_MOV);
            ideDeclarado.AppendChild(paisResid);
            XmlElement paisRes = doc.CreateElement("Pais", NAMESPACE_EVENTO_MOV);
            paisRes.InnerText = NormalizarCodigoPais(pessoa.Nacionalidade);
            paisResid.AppendChild(paisRes);

            // PaisNacionalidade
            XmlElement paisNacionalidade = doc.CreateElement("PaisNacionalidade", NAMESPACE_EVENTO_MOV);
            ideDeclarado.AppendChild(paisNacionalidade);
            XmlElement paisNac = doc.CreateElement("Pais", NAMESPACE_EVENTO_MOV);
            paisNac.InnerText = NormalizarCodigoPais(pessoa.Nacionalidade);
            paisNacionalidade.AppendChild(paisNac);

            // mesCaixa
            XmlElement mesCaixa = doc.CreateElement("mesCaixa", NAMESPACE_EVENTO_MOV);
            evtMovOpFin.AppendChild(mesCaixa);

            XmlElement anoMesCaixa = doc.CreateElement("anoMesCaixa", NAMESPACE_EVENTO_MOV);
            anoMesCaixa.InnerText = periodoStr;
            mesCaixa.AppendChild(anoMesCaixa);

            // movOpFin
            XmlElement movOpFin = doc.CreateElement("movOpFin", NAMESPACE_EVENTO_MOV);
            mesCaixa.AppendChild(movOpFin);

            // Conta
            XmlElement conta = doc.CreateElement("Conta", NAMESPACE_EVENTO_MOV);
            movOpFin.AppendChild(conta);

            // infoConta
            XmlElement infoConta = doc.CreateElement("infoConta", NAMESPACE_EVENTO_MOV);
            conta.AppendChild(infoConta);

            // Reportavel
            XmlElement reportavel = doc.CreateElement("Reportavel", NAMESPACE_EVENTO_MOV);
            infoConta.AppendChild(reportavel);
            XmlElement paisReportavel = doc.CreateElement("Pais", NAMESPACE_EVENTO_MOV);
            paisReportavel.InnerText = "BR";
            reportavel.AppendChild(paisReportavel);

            XmlElement tpConta = doc.CreateElement("tpConta", NAMESPACE_EVENTO_MOV);
            tpConta.InnerText = "1";
            infoConta.AppendChild(tpConta);

            XmlElement subTpConta = doc.CreateElement("subTpConta", NAMESPACE_EVENTO_MOV);
            subTpConta.InnerText = "105";
            infoConta.AppendChild(subTpConta);

            XmlElement tpNumConta = doc.CreateElement("tpNumConta", NAMESPACE_EVENTO_MOV);
            tpNumConta.InnerText = "OECD605";
            infoConta.AppendChild(tpNumConta);

            XmlElement numConta = doc.CreateElement("numConta", NAMESPACE_EVENTO_MOV);
            numConta.InnerText = ConstruirNumeroConta(pessoa);
            infoConta.AppendChild(numConta);

            XmlElement tpRelacaoDeclarado = doc.CreateElement("tpRelacaoDeclarado", NAMESPACE_EVENTO_MOV);
            tpRelacaoDeclarado.InnerText = "1";
            infoConta.AppendChild(tpRelacaoDeclarado);

            XmlElement noTitulares = doc.CreateElement("NoTitulares", NAMESPACE_EVENTO_MOV);
            noTitulares.InnerText = "1";
            infoConta.AppendChild(noTitulares);

            // BalancoConta
            XmlElement balancoConta = doc.CreateElement("BalancoConta", NAMESPACE_EVENTO_MOV);
            infoConta.AppendChild(balancoConta);

            XmlElement totCreditos = doc.CreateElement("totCreditos", NAMESPACE_EVENTO_MOV);
            totCreditos.InnerText = FormatarValorMonetario(pessoa.TotCreditos);
            balancoConta.AppendChild(totCreditos);

            XmlElement totDebitos = doc.CreateElement("totDebitos", NAMESPACE_EVENTO_MOV);
            totDebitos.InnerText = FormatarValorMonetario(pessoa.TotDebitos);
            balancoConta.AppendChild(totDebitos);

            XmlElement totCreditosMesmaTitularidade = doc.CreateElement("totCreditosMesmaTitularidade", NAMESPACE_EVENTO_MOV);
            totCreditosMesmaTitularidade.InnerText = "0,00";
            balancoConta.AppendChild(totCreditosMesmaTitularidade);

            XmlElement totDebitosMesmaTitularidade = doc.CreateElement("totDebitosMesmaTitularidade", NAMESPACE_EVENTO_MOV);
            totDebitosMesmaTitularidade.InnerText = "0,00";
            balancoConta.AppendChild(totDebitosMesmaTitularidade);

            XmlElement vlrUltDia = doc.CreateElement("vlrUltDia", NAMESPACE_EVENTO_MOV);
            vlrUltDia.InnerText = FormatarValorMonetario(pessoa.SaldoAtual);
            balancoConta.AppendChild(vlrUltDia);

            // PgtosAcum
            XmlElement pgtosAcum = doc.CreateElement("PgtosAcum", NAMESPACE_EVENTO_MOV);
            infoConta.AppendChild(pgtosAcum);

            XmlElement tpPgto = doc.CreateElement("tpPgto", NAMESPACE_EVENTO_MOV);
            tpPgto.InnerText = "999";
            pgtosAcum.AppendChild(tpPgto);

            XmlElement totPgtosAcum = doc.CreateElement("totPgtosAcum", NAMESPACE_EVENTO_MOV);
            totPgtosAcum.InnerText = "0,00";
            pgtosAcum.AppendChild(totPgtosAcum);
        }

        private static string FormatIdEvento(long idPessoa)
        {
            string sequencial = idPessoa.ToString().PadLeft(18, '0');
            if (sequencial.Length > 18)
            {
                sequencial = sequencial.Substring(sequencial.Length - 18);
            }
            return "ID" + sequencial;
        }

        private static string NormalizarCodigoPais(string nacionalidade)
        {
            if (string.IsNullOrWhiteSpace(nacionalidade))
            {
                return "BR";
            }

            // Remove espaços e converte para maiúsculas
            string codigo = nacionalidade.Trim().ToUpper();

            // Se já é um código de 2 letras válido, retorna
            if (codigo.Length == 2 && codigo.All(char.IsLetter))
            {
                return codigo;
            }

            // Mapeia nomes comuns para códigos ISO
            if (codigo.Contains("BRASIL") || codigo == "BRASILEIRA" || codigo == "BRASILEIRO")
            {
                return "BR";
            }

            // Se não conseguir mapear, retorna "BR" como padrão
            return "BR";
        }

        private static string NormalizarCpf(string cpf)
        {
            if (string.IsNullOrEmpty(cpf)) return "";
            return cpf.Replace(".", "").Replace("-", "").Trim();
        }

        private static string ConstruirEnderecoLivre(DadosPessoaConta pessoa)
        {
            if (string.IsNullOrEmpty(pessoa.EnderecoLivre))
            {
                List<string> partes = new List<string>();
                if (!string.IsNullOrEmpty(pessoa.TipoLogradouro)) partes.Add(pessoa.TipoLogradouro);
                if (!string.IsNullOrEmpty(pessoa.Logradouro)) partes.Add(pessoa.Logradouro);
                if (!string.IsNullOrEmpty(pessoa.Numero)) partes.Add(pessoa.Numero);
                if (!string.IsNullOrEmpty(pessoa.Complemento)) partes.Add(pessoa.Complemento);
                if (!string.IsNullOrEmpty(pessoa.Bairro)) partes.Add(pessoa.Bairro);
                if (!string.IsNullOrEmpty(pessoa.Cep)) partes.Add(pessoa.Cep);

                return string.Join(" ", partes);
            }
            return pessoa.EnderecoLivre;
        }

        private static string ConstruirNumeroConta(DadosPessoaConta pessoa)
        {
            string numConta = pessoa.NumeroConta;
            if (!string.IsNullOrEmpty(pessoa.DigitoConta))
            {
                numConta += pessoa.DigitoConta;
            }
            return numConta;
        }

        private static string FormatarValorMonetario(decimal valor)
        {
            // Formata sem separador de milhar, apenas com vírgula como decimal (como no Java)
            // Exemplo: 18000.00 -> "18000,00" (não "18.000,00")
            return valor.ToString("0.00", System.Globalization.CultureInfo.GetCultureInfo("en-US")).Replace(".", ",");
        }

        public string GerarXmlCadastroDeclarante(DadosCadastroDeclarante dados, string diretorioSaida)
        {
            if (!Directory.Exists(diretorioSaida))
            {
                Directory.CreateDirectory(diretorioSaida);
            }

            string idEvento = GerarIdEvento();
            string nomeArquivo = $"cadastro-declarante-{dados.CnpjDeclarante}_{DateTime.Now:yyyyMMdd_HHmmss}.xml";
            string caminhoCompleto = Path.Combine(diretorioSaida, nomeArquivo);

            XmlDocument doc = new XmlDocument();
            XmlDeclaration decl = doc.CreateXmlDeclaration("1.0", ENCODING_UTF8, "no");
            doc.AppendChild(decl);

            // Elemento raiz eFinanceira
            XmlElement eFinanceiraRaiz = doc.CreateElement(ELEMENTO_EFINANCEIRA, NAMESPACE_LOTE);
            eFinanceiraRaiz.SetAttribute(ATTR_XMLNS_XSI, XMLNS_XSI_VALUE);
            eFinanceiraRaiz.SetAttribute(ATTR_XMLNS_XSD, XMLNS_XSD_VALUE);
            doc.AppendChild(eFinanceiraRaiz);

            // loteEventosAssincrono
            XmlElement loteEventosAssincrono = doc.CreateElement("loteEventosAssincrono", NAMESPACE_LOTE);
            eFinanceiraRaiz.AppendChild(loteEventosAssincrono);

            // cnpjDeclarante
            XmlElement cnpjDeclaranteLote = doc.CreateElement("cnpjDeclarante", NAMESPACE_LOTE);
            cnpjDeclaranteLote.InnerText = dados.CnpjDeclarante;
            loteEventosAssincrono.AppendChild(cnpjDeclaranteLote);

            // eventos
            XmlElement eventos = doc.CreateElement("eventos", NAMESPACE_LOTE);
            loteEventosAssincrono.AppendChild(eventos);

            // evento
            XmlElement evento = doc.CreateElement("evento", NAMESPACE_LOTE);
            evento.SetAttribute("id", "ID0");
            eventos.AppendChild(evento);

            // eFinanceira interno
            XmlElement eFinanceiraInterno = doc.CreateElement(ELEMENTO_EFINANCEIRA, NAMESPACE_EVENTO_CADASTRO);
            evento.AppendChild(eFinanceiraInterno);

            // evtCadDeclarante
            XmlElement evtCadDeclarante = doc.CreateElement("evtCadDeclarante", NAMESPACE_EVENTO_CADASTRO);
            evtCadDeclarante.SetAttribute("id", idEvento);
            eFinanceiraInterno.AppendChild(evtCadDeclarante);

            // ideEvento
            CriarIdeEventoCadastro(doc, evtCadDeclarante, dados);
            
            // ideDeclarante
            CriarIdeDeclarante(doc, evtCadDeclarante, dados.CnpjDeclarante);
            
            // infoCadastro
            CriarInfoCadastro(doc, evtCadDeclarante, dados);

            // Salvar
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  ",
                Encoding = Encoding.UTF8,
                OmitXmlDeclaration = false
            };

            using (XmlWriter writer = XmlWriter.Create(caminhoCompleto, settings))
            {
                doc.Save(writer);
            }

            return caminhoCompleto;
        }

        private static void CriarIdeEventoCadastro(XmlDocument doc, XmlElement pai, DadosCadastroDeclarante dados)
        {
            XmlElement ideEvento = doc.CreateElement(ELEMENTO_IDE_EVENTO, NAMESPACE_EVENTO_CADASTRO);
            pai.AppendChild(ideEvento);

            XmlElement indRetificacao = doc.CreateElement(ELEMENTO_IND_RETIFICACAO, NAMESPACE_EVENTO_CADASTRO);
            indRetificacao.InnerText = dados.IndRetificacao.ToString();
            ideEvento.AppendChild(indRetificacao);

            if ((dados.IndRetificacao == 2 || dados.IndRetificacao == 3) && !string.IsNullOrEmpty(dados.NrRecibo))
            {
                XmlElement nrRecibo = doc.CreateElement("nrRecibo", NAMESPACE_EVENTO_CADASTRO);
                nrRecibo.InnerText = dados.NrRecibo;
                ideEvento.AppendChild(nrRecibo);
            }

            XmlElement tpAmb = doc.CreateElement(ELEMENTO_TP_AMB, NAMESPACE_EVENTO_CADASTRO);
            tpAmb.InnerText = dados.TipoAmbiente.ToString();
            ideEvento.AppendChild(tpAmb);

            XmlElement aplicEmi = doc.CreateElement(ELEMENTO_APLIC_EMI, NAMESPACE_EVENTO_CADASTRO);
            aplicEmi.InnerText = dados.AplicacaoEmissora.ToString();
            ideEvento.AppendChild(aplicEmi);

            XmlElement verAplic = doc.CreateElement(ELEMENTO_VER_APLIC, NAMESPACE_EVENTO_CADASTRO);
            verAplic.InnerText = VERSAO_APLICATIVO;
            ideEvento.AppendChild(verAplic);
        }

        private static void CriarInfoCadastro(XmlDocument doc, XmlElement pai, DadosCadastroDeclarante dados)
        {
            XmlElement infoCadastro = doc.CreateElement("infoCadastro", NAMESPACE_EVENTO_CADASTRO);
            pai.AppendChild(infoCadastro);

            CriarElementosOpcionaisIniciais(doc, infoCadastro, dados);
            CriarElementosObrigatoriosBasicos(doc, infoCadastro, dados);
            CriarPaisResid(doc, infoCadastro, dados);
            CriarEnderecosOutros(doc, infoCadastro, dados);
        }

        private static void CriarElementosOpcionaisIniciais(XmlDocument doc, XmlElement infoCadastro, DadosCadastroDeclarante dados)
        {
            // GIIN (opcional)
            if (!string.IsNullOrWhiteSpace(dados.GIIN))
            {
                XmlElement giin = doc.CreateElement("GIIN", NAMESPACE_EVENTO_CADASTRO);
                giin.InnerText = dados.GIIN;
                infoCadastro.AppendChild(giin);
            }

            // CategoriaDeclarante (opcional)
            if (!string.IsNullOrWhiteSpace(dados.CategoriaDeclarante))
            {
                XmlElement categoriaDeclarante = doc.CreateElement("CategoriaDeclarante", NAMESPACE_EVENTO_CADASTRO);
                categoriaDeclarante.InnerText = dados.CategoriaDeclarante;
                infoCadastro.AppendChild(categoriaDeclarante);
            }

            // infoTpInstPgto (opcional)
            if (dados.TiposInstPgto != null && dados.TiposInstPgto.Count > 0)
            {
                foreach (var tipoInstPgto in dados.TiposInstPgto)
                {
                    XmlElement infoTpInstPgto = doc.CreateElement("infoTpInstPgto", NAMESPACE_EVENTO_CADASTRO);
                    infoCadastro.AppendChild(infoTpInstPgto);

                    XmlElement tpInstPgto = doc.CreateElement("tpInstPgto", NAMESPACE_EVENTO_CADASTRO);
                    tpInstPgto.InnerText = tipoInstPgto.TpInstPgto;
                    infoTpInstPgto.AppendChild(tpInstPgto);
                }
            }

            // NIF (opcional)
            if (dados.NIFs != null && dados.NIFs.Count > 0)
            {
                foreach (var nif in dados.NIFs)
                {
                    XmlElement nifElement = doc.CreateElement("NIF", NAMESPACE_EVENTO_CADASTRO);
                    infoCadastro.AppendChild(nifElement);

                    XmlElement numeroNIF = doc.CreateElement("NumeroNIF", NAMESPACE_EVENTO_CADASTRO);
                    numeroNIF.InnerText = nif.NumeroNIF;
                    nifElement.AppendChild(numeroNIF);

                    XmlElement paisEmissao = doc.CreateElement("PaisEmissao", NAMESPACE_EVENTO_CADASTRO);
                    paisEmissao.InnerText = nif.PaisEmissao;
                    nifElement.AppendChild(paisEmissao);

                    if (!string.IsNullOrWhiteSpace(nif.TpNIF))
                    {
                        XmlElement tpNIF = doc.CreateElement("tpNIF", NAMESPACE_EVENTO_CADASTRO);
                        tpNIF.InnerText = nif.TpNIF;
                        nifElement.AppendChild(tpNIF);
                    }
                }
            }
        }

        private static void CriarElementosObrigatoriosBasicos(XmlDocument doc, XmlElement infoCadastro, DadosCadastroDeclarante dados)
        {
            // Nome (obrigatório)
            XmlElement nome = doc.CreateElement("nome", NAMESPACE_EVENTO_CADASTRO);
            nome.InnerText = dados.Nome;
            infoCadastro.AppendChild(nome);

            // TpNome (opcional)
            if (!string.IsNullOrWhiteSpace(dados.TpNome))
            {
                XmlElement tpNome = doc.CreateElement("tpNome", NAMESPACE_EVENTO_CADASTRO);
                tpNome.InnerText = dados.TpNome;
                infoCadastro.AppendChild(tpNome);
            }

            // EnderecoLivre (obrigatório)
            XmlElement enderecoLivre = doc.CreateElement("enderecoLivre", NAMESPACE_EVENTO_CADASTRO);
            enderecoLivre.InnerText = dados.EnderecoLivre;
            infoCadastro.AppendChild(enderecoLivre);

            // TpEndereco (opcional)
            if (!string.IsNullOrWhiteSpace(dados.TpEndereco))
            {
                XmlElement tpEndereco = doc.CreateElement("tpEndereco", NAMESPACE_EVENTO_CADASTRO);
                tpEndereco.InnerText = dados.TpEndereco;
                infoCadastro.AppendChild(tpEndereco);
            }

            // Municipio (obrigatório)
            XmlElement municipio = doc.CreateElement("municipio", NAMESPACE_EVENTO_CADASTRO);
            municipio.InnerText = dados.Municipio;
            infoCadastro.AppendChild(municipio);

            // UF (obrigatório)
            XmlElement uf = doc.CreateElement("UF", NAMESPACE_EVENTO_CADASTRO);
            uf.InnerText = dados.UF;
            infoCadastro.AppendChild(uf);

            // CEP (obrigatório)
            XmlElement cep = doc.CreateElement("CEP", NAMESPACE_EVENTO_CADASTRO);
            cep.InnerText = dados.CEP;
            infoCadastro.AppendChild(cep);

            // Pais (obrigatório)
            XmlElement pais = doc.CreateElement("Pais", NAMESPACE_EVENTO_CADASTRO);
            pais.InnerText = dados.Pais;
            infoCadastro.AppendChild(pais);
        }

        private static void CriarPaisResid(XmlDocument doc, XmlElement infoCadastro, DadosCadastroDeclarante dados)
        {
            // paisResid (obrigatório - deve conter pelo menos "BR")
            if (dados.PaisResid != null && dados.PaisResid.Count > 0)
            {
                foreach (var paisResid in dados.PaisResid)
                {
                    XmlElement paisResidElement = doc.CreateElement("paisResid", NAMESPACE_EVENTO_CADASTRO);
                    infoCadastro.AppendChild(paisResidElement);

                    XmlElement paisElement = doc.CreateElement("Pais", NAMESPACE_EVENTO_CADASTRO);
                    paisElement.InnerText = paisResid;
                    paisResidElement.AppendChild(paisElement);
                }
            }
            else
            {
                // Se não houver, adicionar "BR" por padrão
                XmlElement paisResidElement = doc.CreateElement("paisResid", NAMESPACE_EVENTO_CADASTRO);
                infoCadastro.AppendChild(paisResidElement);

                XmlElement paisElement = doc.CreateElement("Pais", NAMESPACE_EVENTO_CADASTRO);
                paisElement.InnerText = "BR";
                paisResidElement.AppendChild(paisElement);
            }
        }

        private static void CriarEnderecosOutros(XmlDocument doc, XmlElement infoCadastro, DadosCadastroDeclarante dados)
        {
            // EnderecoOutros (opcional)
            if (dados.EnderecosOutros == null || dados.EnderecosOutros.Count == 0)
            {
                return;
            }

            foreach (var enderecoOutros in dados.EnderecosOutros)
            {
                XmlElement enderecoOutrosElement = doc.CreateElement("EnderecoOutros", NAMESPACE_EVENTO_CADASTRO);
                infoCadastro.AppendChild(enderecoOutrosElement);

                if (!string.IsNullOrWhiteSpace(enderecoOutros.TpEndereco))
                {
                    XmlElement tpEnderecoOutros = doc.CreateElement("tpEndereco", NAMESPACE_EVENTO_CADASTRO);
                    tpEnderecoOutros.InnerText = enderecoOutros.TpEndereco;
                    enderecoOutrosElement.AppendChild(tpEnderecoOutros);
                }

                if (!string.IsNullOrWhiteSpace(enderecoOutros.EnderecoLivre))
                {
                    XmlElement enderecoLivreOutros = doc.CreateElement("EnderecoLivre", NAMESPACE_EVENTO_CADASTRO);
                    enderecoLivreOutros.InnerText = enderecoOutros.EnderecoLivre;
                    enderecoOutrosElement.AppendChild(enderecoLivreOutros);
                }

                if (enderecoOutros.EnderecoEstrutura != null)
                {
                    CriarEnderecoEstrutura(doc, enderecoOutrosElement, enderecoOutros.EnderecoEstrutura);
                }

                if (!string.IsNullOrWhiteSpace(enderecoOutros.Pais))
                {
                    XmlElement paisOutros = doc.CreateElement("Pais", NAMESPACE_EVENTO_CADASTRO);
                    paisOutros.InnerText = enderecoOutros.Pais;
                    enderecoOutrosElement.AppendChild(paisOutros);
                }
            }
        }

        private static void CriarEnderecoEstrutura(XmlDocument doc, XmlElement enderecoOutrosElement, DadosEnderecoEstrutura enderecoEstrutura)
        {
            XmlElement enderecoEstruturaElement = doc.CreateElement("EnderecoEstrutura", NAMESPACE_EVENTO_CADASTRO);
            enderecoOutrosElement.AppendChild(enderecoEstruturaElement);

            if (!string.IsNullOrWhiteSpace(enderecoEstrutura.EnderecoLivre))
            {
                XmlElement enderecoLivreEstrutura = doc.CreateElement("EnderecoLivre", NAMESPACE_EVENTO_CADASTRO);
                enderecoLivreEstrutura.InnerText = enderecoEstrutura.EnderecoLivre;
                enderecoEstruturaElement.AppendChild(enderecoLivreEstrutura);
            }

            if (enderecoEstrutura.Endereco != null)
            {
                CriarEnderecoDetalhado(doc, enderecoEstruturaElement, enderecoEstrutura.Endereco);
            }

            if (!string.IsNullOrWhiteSpace(enderecoEstrutura.CEP))
            {
                XmlElement cepEstrutura = doc.CreateElement("CEP", NAMESPACE_EVENTO_CADASTRO);
                cepEstrutura.InnerText = enderecoEstrutura.CEP;
                enderecoEstruturaElement.AppendChild(cepEstrutura);
            }

            if (!string.IsNullOrWhiteSpace(enderecoEstrutura.Municipio))
            {
                XmlElement municipioEstrutura = doc.CreateElement("Município", NAMESPACE_EVENTO_CADASTRO);
                municipioEstrutura.InnerText = enderecoEstrutura.Municipio;
                enderecoEstruturaElement.AppendChild(municipioEstrutura);
            }

            if (!string.IsNullOrWhiteSpace(enderecoEstrutura.UF))
            {
                XmlElement ufEstrutura = doc.CreateElement("UF", NAMESPACE_EVENTO_CADASTRO);
                ufEstrutura.InnerText = enderecoEstrutura.UF;
                enderecoEstruturaElement.AppendChild(ufEstrutura);
            }
        }

        private static void CriarEnderecoDetalhado(XmlDocument doc, XmlElement enderecoEstruturaElement, DadosEndereco endereco)
        {
            XmlElement enderecoElement = doc.CreateElement("Endereco", NAMESPACE_EVENTO_CADASTRO);
            enderecoEstruturaElement.AppendChild(enderecoElement);

            if (!string.IsNullOrWhiteSpace(endereco.Logradouro))
            {
                XmlElement logradouro = doc.CreateElement("Logradouro", NAMESPACE_EVENTO_CADASTRO);
                logradouro.InnerText = endereco.Logradouro;
                enderecoElement.AppendChild(logradouro);
            }

            if (!string.IsNullOrWhiteSpace(endereco.Numero))
            {
                XmlElement numero = doc.CreateElement(ELEMENTO_NUMERO, NAMESPACE_EVENTO_CADASTRO);
                numero.InnerText = endereco.Numero;
                enderecoElement.AppendChild(numero);
            }

            if (!string.IsNullOrWhiteSpace(endereco.Complemento))
            {
                XmlElement complemento = doc.CreateElement("Complemento", NAMESPACE_EVENTO_CADASTRO);
                complemento.InnerText = endereco.Complemento;
                enderecoElement.AppendChild(complemento);
            }

            if (!string.IsNullOrWhiteSpace(endereco.Andar))
            {
                XmlElement andar = doc.CreateElement("Andar", NAMESPACE_EVENTO_CADASTRO);
                andar.InnerText = endereco.Andar;
                enderecoElement.AppendChild(andar);
            }

            if (!string.IsNullOrWhiteSpace(endereco.Bairro))
            {
                XmlElement bairro = doc.CreateElement("Bairro", NAMESPACE_EVENTO_CADASTRO);
                bairro.InnerText = endereco.Bairro;
                enderecoElement.AppendChild(bairro);
            }

            if (!string.IsNullOrWhiteSpace(endereco.CaixaPostal))
            {
                XmlElement caixaPostal = doc.CreateElement("CaixaPostal", NAMESPACE_EVENTO_CADASTRO);
                caixaPostal.InnerText = endereco.CaixaPostal;
                enderecoElement.AppendChild(caixaPostal);
            }
        }

        private static string GerarIdEvento()
        {
            long timestamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            string sequencial = timestamp.ToString().PadLeft(18, '0');
            return "ID" + sequencial;
        }
    }
}
