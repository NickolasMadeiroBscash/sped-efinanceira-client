using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using ExemploAssinadorXML.Models;

namespace ExemploAssinadorXML.Services
{
    public class EfinanceiraExportacaoExcelService
    {
        /// <summary>
        /// Obtém o código do tipo de evento baseado no TipoLote
        /// </summary>
        private string ObterCodigoTipoEvento(TipoLote tipoLote, LoteBancoInfo lote = null)
        {
            // Verificar pelo caminho do arquivo para casos especiais
            if (lote != null && !string.IsNullOrEmpty(lote.CaminhoArquivoXml))
            {
                string nomeArquivo = lote.CaminhoArquivoXml.ToLower();
                
                if (nomeArquivo.Contains("exclusao") && nomeArquivo.Contains("efinanceira"))
                    return "009"; // Exclusão de eFinanceira
                
                if (nomeArquivo.Contains("exclusao") || nomeArquivo.Contains("exclusão"))
                    return "006"; // Exclusão de Evento Enviado Indevidamente
                
                if (nomeArquivo.Contains("cadastro") && nomeArquivo.Contains("intermediario"))
                    return "007"; // Cadastro de Intermediário
                
                if (nomeArquivo.Contains("cadastro") && nomeArquivo.Contains("patrocinado"))
                    return "008"; // Cadastro de Patrocinado
                
                if ((nomeArquivo.Contains("info") && nomeArquivo.Contains("empresa")) || nomeArquivo.Contains("declarante"))
                    return "001"; // Informações da Empresa Declarante
            }
            
            // Mapeamento padrão por TipoLote
            switch (tipoLote)
            {
                case TipoLote.Abertura:
                    return "002"; // Abertura
                case TipoLote.Movimentacao:
                    return "003"; // Movimento de Operações Financeiras
                case TipoLote.Fechamento:
                    return "005"; // Fechamento
                case TipoLote.CadastroDeclarante:
                    return "001"; // Informações da Empresa Declarante
                default:
                    return "003"; // Padrão: Movimento de Operações Financeiras
            }
        }

        /// <summary>
        /// Exporta todos os dados de um período (abertura, movimentações, fechamento) para Excel
        /// </summary>
        public string ExportarPeriodoParaExcel(string periodo, string ambiente, string diretorioExportacao)
        {
            var persistenceService = new EfinanceiraDatabasePersistenceService();
            
            // Buscar todos os lotes do período, ordenados por data (mais recente primeiro)
            var lotes = persistenceService.BuscarLotes(null, null, periodo, ambiente);
            
            if (lotes.Count == 0)
            {
                throw new Exception($"Nenhum lote encontrado para o período {periodo} no ambiente {ambiente}");
            }

            // Garantir que sempre pegue apenas os lotes mais recentes de cada tipo
            // Isso é importante para casos de teste onde pode haver múltiplos lotes do mesmo período
            // Quando há múltiplos lotes do mesmo tipo, pegar apenas o mais recente (ignorar os passados)
            var lotesAgrupadosPorTipo = lotes
                .GroupBy(l => l.TipoLote)
                .ToDictionary(g => g.Key, g => g.OrderByDescending(l => l.DataCriacao).ToList());
            
            // Separar por tipo - sempre pegar apenas o mais recente de cada tipo
            var loteAbertura = lotesAgrupadosPorTipo.ContainsKey(TipoLote.Abertura) 
                ? lotesAgrupadosPorTipo[TipoLote.Abertura].FirstOrDefault() 
                : null;
            
            // Pegar apenas o lote de movimentação mais recente (não todos)
            var loteMovimentacao = lotesAgrupadosPorTipo.ContainsKey(TipoLote.Movimentacao)
                ? lotesAgrupadosPorTipo[TipoLote.Movimentacao].FirstOrDefault()
                : null;
            
            // Converter para lista para manter compatibilidade com o código existente
            var lotesMovimentacao = loteMovimentacao != null 
                ? new List<LoteBancoInfo> { loteMovimentacao }
                : new List<LoteBancoInfo>();
            
            var loteFechamento = lotesAgrupadosPorTipo.ContainsKey(TipoLote.Fechamento)
                ? lotesAgrupadosPorTipo[TipoLote.Fechamento].FirstOrDefault()
                : null;

            // Criar arquivo Excel
            string nomeArquivo = $"eFinanceira_Periodo_{periodo}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            string caminhoArquivo = Path.Combine(diretorioExportacao, nomeArquivo);

            // EPPlus 4.5.3.3 é gratuito para uso não comercial, não precisa configurar LicenseContext

            using (var package = new ExcelPackage())
            {
                // Aba 1: Resumo do Período
                CriarAbaResumo(package, periodo, ambiente, loteAbertura, lotesMovimentacao, loteFechamento);

                // Aba 2: Abertura
                if (loteAbertura != null)
                {
                    CriarAbaAbertura(package, loteAbertura, persistenceService);
                }
                else
                {
                    CriarAbaVazia(package, "Abertura", "Abertura ainda não foi enviada para este período.");
                }

                // Aba 3: Movimentações
                if (lotesMovimentacao.Count > 0)
                {
                    CriarAbaMovimentacoes(package, lotesMovimentacao, persistenceService);
                }
                else
                {
                    CriarAbaVazia(package, "Movimentações", "Movimentações ainda não foram enviadas para este período.");
                }

                // Aba 4: Fechamento
                if (loteFechamento != null)
                {
                    CriarAbaFechamento(package, loteFechamento, persistenceService);
                }
                else
                {
                    CriarAbaVazia(package, "Fechamento", "Fechamento ainda não foi enviado para este período.");
                }

                // Aba 5: Eventos Detalhados (apenas dos lotes mais recentes)
                var lotesMaisRecentes = new List<LoteBancoInfo>();
                if (loteAbertura != null) lotesMaisRecentes.Add(loteAbertura);
                lotesMaisRecentes.AddRange(lotesMovimentacao);
                if (loteFechamento != null) lotesMaisRecentes.Add(loteFechamento);
                
                CriarAbaEventosDetalhados(package, lotesMaisRecentes, persistenceService);

                // Salvar arquivo
                package.SaveAs(new FileInfo(caminhoArquivo));
            }

            return caminhoArquivo;
        }

        /// <summary>
        /// Exporta dados de um lote específico para Excel
        /// </summary>
        public string ExportarLoteParaExcel(long idLote, string diretorioExportacao)
        {
            var persistenceService = new EfinanceiraDatabasePersistenceService();
            
            // Buscar lote
            var lotes = persistenceService.BuscarLotes(null, null, null, null);
            var lote = lotes.FirstOrDefault(l => l.IdLote == idLote);
            
            if (lote == null)
            {
                throw new Exception($"Lote com ID {idLote} não encontrado.");
            }

            // Buscar eventos do lote
            var eventos = persistenceService.BuscarEventosDoLote(idLote);

            // Criar arquivo Excel
            string nomeArquivo = $"eFinanceira_Lote_{idLote}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            string caminhoArquivo = Path.Combine(diretorioExportacao, nomeArquivo);

            // EPPlus 4.5.3.3 é gratuito para uso não comercial, não precisa configurar LicenseContext

            using (var package = new ExcelPackage())
            {
                // Aba 1: Informações do Lote
                CriarAbaInformacoesLote(package, lote);

                // Aba 2: Eventos
                CriarAbaEventos(package, eventos);

                package.SaveAs(new FileInfo(caminhoArquivo));
            }

            return caminhoArquivo;
        }

        private void CriarAbaResumo(ExcelPackage package, string periodo, string ambiente, 
            LoteBancoInfo loteAbertura, List<LoteBancoInfo> lotesMovimentacao, LoteBancoInfo loteFechamento)
        {
            var worksheet = package.Workbook.Worksheets.Add("Resumo do Período");
            int linha = 1;

            // Título
            worksheet.Cells[linha, 1].Value = "RELATÓRIO DETALHADO DO PERÍODO E-FINANCEIRA";
            worksheet.Cells[linha, 1, linha, 8].Merge = true;
            worksheet.Cells[linha, 1].Style.Font.Size = 16;
            worksheet.Cells[linha, 1].Style.Font.Bold = true;
            worksheet.Cells[linha, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            linha += 2;

            // Informações do Período
            worksheet.Cells[linha, 1].Value = "Período:";
            worksheet.Cells[linha, 1].Style.Font.Bold = true;
            worksheet.Cells[linha, 2].Value = periodo;
            linha++;

            worksheet.Cells[linha, 1].Value = "Ambiente:";
            worksheet.Cells[linha, 1].Style.Font.Bold = true;
            worksheet.Cells[linha, 2].Value = ambiente ?? "Não informado";
            linha++;

            // CNPJ Declarante da Abertura
            if (loteAbertura != null && !string.IsNullOrEmpty(loteAbertura.CnpjDeclarante))
            {
                worksheet.Cells[linha, 1].Value = "CNPJ Declarante:";
                worksheet.Cells[linha, 1].Style.Font.Bold = true;
                worksheet.Cells[linha, 2].Value = loteAbertura.CnpjDeclarante;
                linha++;
            }

            worksheet.Cells[linha, 1].Value = "Data de Exportação:";
            worksheet.Cells[linha, 1].Style.Font.Bold = true;
            worksheet.Cells[linha, 2].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            linha += 2;

            // Cabeçalho da tabela de lotes
            worksheet.Cells[linha, 1].Value = "Tipo de Evento";
            worksheet.Cells[linha, 2].Value = "Código";
            worksheet.Cells[linha, 3].Value = "ID Lote";
            worksheet.Cells[linha, 4].Value = "Número Lote";
            worksheet.Cells[linha, 5].Value = "Protocolo";
            worksheet.Cells[linha, 6].Value = "Status";
            worksheet.Cells[linha, 7].Value = "Qtd. Eventos";
            worksheet.Cells[linha, 8].Value = "Eventos Sucesso";
            worksheet.Cells[linha, 9].Value = "Eventos Erro";
            worksheet.Cells[linha, 10].Value = "Data Criação";
            worksheet.Cells[linha, 11].Value = "Data Envio";
            worksheet.Cells[linha, 12].Value = "CNPJ Declarante";

            var headerRange = worksheet.Cells[linha, 1, linha, 12];
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            headerRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            linha++;

            // Abertura
            if (loteAbertura != null)
            {
                worksheet.Cells[linha, 1].Value = "Abertura";
                worksheet.Cells[linha, 2].Value = ObterCodigoTipoEvento(TipoLote.Abertura, loteAbertura);
                worksheet.Cells[linha, 3].Value = loteAbertura.IdLote;
                worksheet.Cells[linha, 4].Value = loteAbertura.NumeroLote;
                worksheet.Cells[linha, 5].Value = loteAbertura.ProtocoloEnvio ?? "Não disponível";
                worksheet.Cells[linha, 6].Value = loteAbertura.Status ?? "N/A";
                worksheet.Cells[linha, 7].Value = loteAbertura.QuantidadeEventos;
                worksheet.Cells[linha, 8].Value = loteAbertura.TotalEventosSucesso;
                worksheet.Cells[linha, 9].Value = loteAbertura.TotalEventosComErro;
                worksheet.Cells[linha, 10].Value = loteAbertura.DataCriacao.ToString("dd/MM/yyyy HH:mm:ss");
                worksheet.Cells[linha, 11].Value = loteAbertura.DataEnvio?.ToString("dd/MM/yyyy HH:mm:ss") ?? "Não enviado";
                worksheet.Cells[linha, 12].Value = loteAbertura.CnpjDeclarante ?? "N/A";
                
                // Colorir status
                if (loteAbertura.Status?.ToUpper().Contains("SUCESSO") == true || loteAbertura.Status?.ToUpper().Contains("ACEITO") == true)
                    worksheet.Cells[linha, 6].Style.Font.Color.SetColor(Color.Green);
                else if (loteAbertura.Status?.ToUpper().Contains("ERRO") == true || loteAbertura.Status?.ToUpper().Contains("REJEITADO") == true)
                    worksheet.Cells[linha, 6].Style.Font.Color.SetColor(Color.Red);
                
                linha++;
            }

            // Movimentações
            foreach (var loteMov in lotesMovimentacao.OrderByDescending(l => l.DataCriacao))
            {
                worksheet.Cells[linha, 1].Value = "Movimento de Operações Financeiras";
                worksheet.Cells[linha, 2].Value = ObterCodigoTipoEvento(TipoLote.Movimentacao, loteMov);
                worksheet.Cells[linha, 3].Value = loteMov.IdLote;
                worksheet.Cells[linha, 4].Value = loteMov.NumeroLote;
                worksheet.Cells[linha, 5].Value = loteMov.ProtocoloEnvio ?? "Não disponível";
                worksheet.Cells[linha, 6].Value = loteMov.Status ?? "N/A";
                worksheet.Cells[linha, 7].Value = loteMov.QuantidadeEventos;
                worksheet.Cells[linha, 8].Value = loteMov.TotalEventosSucesso;
                worksheet.Cells[linha, 9].Value = loteMov.TotalEventosComErro;
                worksheet.Cells[linha, 10].Value = loteMov.DataCriacao.ToString("dd/MM/yyyy HH:mm:ss");
                worksheet.Cells[linha, 11].Value = loteMov.DataEnvio?.ToString("dd/MM/yyyy HH:mm:ss") ?? "Não enviado";
                worksheet.Cells[linha, 12].Value = loteMov.CnpjDeclarante ?? "N/A";
                
                // Colorir status
                if (loteMov.Status?.ToUpper().Contains("SUCESSO") == true || loteMov.Status?.ToUpper().Contains("ACEITO") == true)
                    worksheet.Cells[linha, 6].Style.Font.Color.SetColor(Color.Green);
                else if (loteMov.Status?.ToUpper().Contains("ERRO") == true || loteMov.Status?.ToUpper().Contains("REJEITADO") == true)
                    worksheet.Cells[linha, 6].Style.Font.Color.SetColor(Color.Red);
                
                linha++;
            }

            // Fechamento
            if (loteFechamento != null)
            {
                worksheet.Cells[linha, 1].Value = "Fechamento";
                worksheet.Cells[linha, 2].Value = ObterCodigoTipoEvento(TipoLote.Fechamento, loteFechamento);
                worksheet.Cells[linha, 3].Value = loteFechamento.IdLote;
                worksheet.Cells[linha, 4].Value = loteFechamento.NumeroLote;
                worksheet.Cells[linha, 5].Value = loteFechamento.ProtocoloEnvio ?? "Não disponível";
                worksheet.Cells[linha, 6].Value = loteFechamento.Status ?? "N/A";
                worksheet.Cells[linha, 7].Value = loteFechamento.QuantidadeEventos;
                worksheet.Cells[linha, 8].Value = loteFechamento.TotalEventosSucesso;
                worksheet.Cells[linha, 9].Value = loteFechamento.TotalEventosComErro;
                worksheet.Cells[linha, 10].Value = loteFechamento.DataCriacao.ToString("dd/MM/yyyy HH:mm:ss");
                worksheet.Cells[linha, 11].Value = loteFechamento.DataEnvio?.ToString("dd/MM/yyyy HH:mm:ss") ?? "Não enviado";
                worksheet.Cells[linha, 12].Value = loteFechamento.CnpjDeclarante ?? "N/A";
                
                // Colorir status
                if (loteFechamento.Status?.ToUpper().Contains("SUCESSO") == true || loteFechamento.Status?.ToUpper().Contains("ACEITO") == true)
                    worksheet.Cells[linha, 6].Style.Font.Color.SetColor(Color.Green);
                else if (loteFechamento.Status?.ToUpper().Contains("ERRO") == true || loteFechamento.Status?.ToUpper().Contains("REJEITADO") == true)
                    worksheet.Cells[linha, 6].Style.Font.Color.SetColor(Color.Red);
                
                linha++;
            }

            linha += 2;

            // Estatísticas Gerais
            if (loteAbertura != null || lotesMovimentacao.Count > 0 || loteFechamento != null)
            {
                worksheet.Cells[linha, 1].Value = "ESTATÍSTICAS GERAIS";
                worksheet.Cells[linha, 1, linha, 3].Merge = true;
                worksheet.Cells[linha, 1].Style.Font.Bold = true;
                worksheet.Cells[linha, 1].Style.Font.Size = 12;
                linha++;

                int totalEventos = 0;
                int totalEventosSucesso = 0;
                int totalEventosErro = 0;

                if (loteAbertura != null)
                {
                    totalEventos += loteAbertura.QuantidadeEventos;
                    totalEventosSucesso += loteAbertura.TotalEventosSucesso;
                    totalEventosErro += loteAbertura.TotalEventosComErro;
                }

                foreach (var loteMov in lotesMovimentacao)
                {
                    totalEventos += loteMov.QuantidadeEventos;
                    totalEventosSucesso += loteMov.TotalEventosSucesso;
                    totalEventosErro += loteMov.TotalEventosComErro;
                }

                if (loteFechamento != null)
                {
                    totalEventos += loteFechamento.QuantidadeEventos;
                    totalEventosSucesso += loteFechamento.TotalEventosSucesso;
                    totalEventosErro += loteFechamento.TotalEventosComErro;
                }

                worksheet.Cells[linha, 1].Value = "Total de Eventos (Abertura, movimentações e fechamento):";
                worksheet.Cells[linha, 1].Style.Font.Bold = true;
                worksheet.Cells[linha, 2].Value = totalEventos;
                linha++;

                worksheet.Cells[linha, 1].Value = "Eventos de movimentação com Sucesso:";
                worksheet.Cells[linha, 1].Style.Font.Bold = true;
                worksheet.Cells[linha, 2].Value = totalEventosSucesso;
                worksheet.Cells[linha, 2].Style.Font.Color.SetColor(Color.Green);
                linha++;

                worksheet.Cells[linha, 1].Value = "Eventos de movimentação com Erro:";
                worksheet.Cells[linha, 1].Style.Font.Bold = true;
                worksheet.Cells[linha, 2].Value = totalEventosErro;
                worksheet.Cells[linha, 2].Style.Font.Color.SetColor(totalEventosErro > 0 ? Color.Red : Color.Green);
            }

            // Ajustar largura das colunas
            worksheet.Column(1).Width = 35;  // Tipo de Evento
            worksheet.Column(2).Width = 10;  // Código
            worksheet.Column(3).Width = 12;  // ID Lote
            worksheet.Column(4).Width = 15;  // Número Lote
            worksheet.Column(5).Width = 30;  // Protocolo
            worksheet.Column(6).Width = 20;  // Status
            worksheet.Column(7).Width = 15;  // Qtd. Eventos
            worksheet.Column(8).Width = 18; // Eventos Sucesso
            worksheet.Column(9).Width = 15; // Eventos Erro
            worksheet.Column(10).Width = 20; // Data Criação
            worksheet.Column(11).Width = 20; // Data Envio
            worksheet.Column(12).Width = 20; // CNPJ Declarante

            // Aplicar bordas na tabela
            if (linha > 3)
            {
                var dataRange = worksheet.Cells[3, 1, linha - 1, 12];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }
        }

        private void CriarAbaAbertura(ExcelPackage package, LoteBancoInfo lote, EfinanceiraDatabasePersistenceService persistenceService)
        {
            var worksheet = package.Workbook.Worksheets.Add("Abertura");
            int linha = 1;

            // Título
            worksheet.Cells[linha, 1].Value = "DADOS DA ABERTURA";
            worksheet.Cells[linha, 1, linha, 6].Merge = true;
            worksheet.Cells[linha, 1].Style.Font.Size = 14;
            worksheet.Cells[linha, 1].Style.Font.Bold = true;
            worksheet.Cells[linha, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            linha += 2;

            // Informações do Lote
            AdicionarLinhaInfo(worksheet, ref linha, "Tipo de Evento:", "Abertura");
            AdicionarLinhaInfo(worksheet, ref linha, "Código do Tipo:", ObterCodigoTipoEvento(TipoLote.Abertura, lote));
            AdicionarLinhaInfo(worksheet, ref linha, "ID do Lote:", lote.IdLote.ToString());
            AdicionarLinhaInfo(worksheet, ref linha, "Período:", lote.Periodo ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Número do Lote:", lote.NumeroLote.ToString());
            AdicionarLinhaInfo(worksheet, ref linha, "CNPJ Declarante:", lote.CnpjDeclarante ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Ambiente:", lote.Ambiente ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Protocolo de Envio:", lote.ProtocoloEnvio ?? "Não disponível");
            AdicionarLinhaInfo(worksheet, ref linha, "Status:", lote.Status ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Data de Criação:", lote.DataCriacao.ToString("dd/MM/yyyy HH:mm:ss"));
            if (lote.DataEnvio.HasValue)
                AdicionarLinhaInfo(worksheet, ref linha, "Data de Envio:", lote.DataEnvio.Value.ToString("dd/MM/yyyy HH:mm:ss"));
            if (lote.DataConfirmacao.HasValue)
                AdicionarLinhaInfo(worksheet, ref linha, "Data de Confirmação:", lote.DataConfirmacao.Value.ToString("dd/MM/yyyy HH:mm:ss"));
            AdicionarLinhaInfo(worksheet, ref linha, "É Retificação:", lote.EhRetificacao ? "Sim" : "Não");
            if (lote.EhRetificacao && lote.IdLoteOriginal.HasValue)
                AdicionarLinhaInfo(worksheet, ref linha, "ID Lote Original:", lote.IdLoteOriginal.Value.ToString());

            linha += 2;

            // Respostas do Servidor
            if (lote.CodigoRespostaEnvio.HasValue || lote.CodigoRespostaConsulta.HasValue)
            {
                worksheet.Cells[linha, 1].Value = "RESPOSTAS DO SERVIDOR";
                worksheet.Cells[linha, 1, linha, 6].Merge = true;
                worksheet.Cells[linha, 1].Style.Font.Bold = true;
                worksheet.Cells[linha, 1].Style.Font.Size = 12;
                linha++;

                if (lote.CodigoRespostaEnvio.HasValue)
                {
                    AdicionarLinhaInfo(worksheet, ref linha, "Código Resposta Envio:", lote.CodigoRespostaEnvio.Value.ToString());
                    AdicionarLinhaInfo(worksheet, ref linha, "Descrição Resposta Envio:", lote.DescricaoRespostaEnvio ?? "N/A");
                }

                if (lote.CodigoRespostaConsulta.HasValue)
                {
                    AdicionarLinhaInfo(worksheet, ref linha, "Código Resposta Consulta:", lote.CodigoRespostaConsulta.Value.ToString());
                    AdicionarLinhaInfo(worksheet, ref linha, "Descrição Resposta Consulta:", lote.DescricaoRespostaConsulta ?? "N/A");
                }

                linha += 2;
            }

            // Eventos
            var eventos = persistenceService.BuscarEventosDoLote(lote.IdLote);
            if (eventos.Count > 0)
            {
                worksheet.Cells[linha, 1].Value = "EVENTOS DA ABERTURA";
                worksheet.Cells[linha, 1, linha, 8].Merge = true;
                worksheet.Cells[linha, 1].Style.Font.Bold = true;
                worksheet.Cells[linha, 1].Style.Font.Size = 12;
                linha++;

                // Cabeçalhos
                CriarCabecalhoEventos(worksheet, linha);
                linha++;

                // Dados dos eventos
                foreach (var evento in eventos)
                {
                    AdicionarLinhaEvento(worksheet, linha, evento);
                    linha++;
                }
            }

            // Ajustar largura das colunas
            AjustarLarguraColunas(worksheet);
        }

        private void CriarAbaMovimentacoes(ExcelPackage package, List<LoteBancoInfo> lotes, EfinanceiraDatabasePersistenceService persistenceService)
        {
            var worksheet = package.Workbook.Worksheets.Add("Movimentações");
            int linha = 1;

            // Título
            worksheet.Cells[linha, 1].Value = $"MOVIMENTAÇÕES FINANCEIRAS - EVENTOS DETALHADOS";
            worksheet.Cells[linha, 1, linha, 18].Merge = true;
            worksheet.Cells[linha, 1].Style.Font.Size = 14;
            worksheet.Cells[linha, 1].Style.Font.Bold = true;
            worksheet.Cells[linha, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            linha += 2;

            // Cabeçalhos com todos os campos da tabela tb_efinanceira_evento
            worksheet.Cells[linha, 1].Value = "ID Evento";
            worksheet.Cells[linha, 2].Value = "ID Lote";
            worksheet.Cells[linha, 3].Value = "ID Pessoa";
            worksheet.Cells[linha, 4].Value = "ID Conta";
            worksheet.Cells[linha, 5].Value = "CPF";
            worksheet.Cells[linha, 6].Value = "Nome";
            worksheet.Cells[linha, 7].Value = "Número Conta";
            worksheet.Cells[linha, 8].Value = "Dígito Conta";
            worksheet.Cells[linha, 9].Value = "Saldo Atual";
            worksheet.Cells[linha, 10].Value = "Tot. Créditos";
            worksheet.Cells[linha, 11].Value = "Tot. Débitos";
            worksheet.Cells[linha, 12].Value = "ID Evento XML";
            worksheet.Cells[linha, 13].Value = "Status Evento";
            worksheet.Cells[linha, 14].Value = "Ocorrências JSON";
            worksheet.Cells[linha, 15].Value = "Data Criação";
            worksheet.Cells[linha, 16].Value = "Número Recibo";
            worksheet.Cells[linha, 17].Value = "Ind. Retificação";
            worksheet.Cells[linha, 18].Value = "Tipo Evento";

            var headerRange = worksheet.Cells[linha, 1, linha, 18];
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            headerRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            linha++;

            // Para cada lote de movimentação, buscar todos os eventos e exibir em linhas
            foreach (var lote in lotes.OrderByDescending(l => l.DataCriacao))
            {
                var eventos = persistenceService.BuscarEventosDoLote(lote.IdLote);
                
                foreach (var evento in eventos)
                {
                    worksheet.Cells[linha, 1].Value = evento.IdEvento;
                    worksheet.Cells[linha, 2].Value = evento.IdLote;
                    worksheet.Cells[linha, 3].Value = evento.IdPessoa?.ToString() ?? "";
                    worksheet.Cells[linha, 4].Value = evento.IdConta?.ToString() ?? "";
                    worksheet.Cells[linha, 5].Value = evento.Cpf ?? "";
                    worksheet.Cells[linha, 6].Value = evento.Nome ?? "";
                    worksheet.Cells[linha, 7].Value = evento.NumeroConta?.ToString() ?? "";
                    worksheet.Cells[linha, 8].Value = evento.DigitoConta ?? "";
                    
                    // Saldo Atual
                    if (evento.SaldoAtual.HasValue)
                    {
                        worksheet.Cells[linha, 9].Value = evento.SaldoAtual.Value;
                        worksheet.Cells[linha, 9].Style.Numberformat.Format = "#,##0.00";
                    }
                    else
                    {
                        worksheet.Cells[linha, 9].Value = "";
                    }
                    
                    // Tot. Créditos
                    if (evento.TotCreditos.HasValue)
                    {
                        worksheet.Cells[linha, 10].Value = evento.TotCreditos.Value;
                        worksheet.Cells[linha, 10].Style.Numberformat.Format = "#,##0.00";
                    }
                    else
                    {
                        worksheet.Cells[linha, 10].Value = "";
                    }
                    
                    // Tot. Débitos
                    if (evento.TotDebitos.HasValue)
                    {
                        worksheet.Cells[linha, 11].Value = evento.TotDebitos.Value;
                        worksheet.Cells[linha, 11].Style.Numberformat.Format = "#,##0.00";
                    }
                    else
                    {
                        worksheet.Cells[linha, 11].Value = "";
                    }
                    
                    worksheet.Cells[linha, 12].Value = evento.IdEventoXml ?? "";
                    worksheet.Cells[linha, 13].Value = evento.StatusEvento ?? "";
                    worksheet.Cells[linha, 14].Value = evento.OcorrenciasJson ?? "";
                    worksheet.Cells[linha, 15].Value = evento.DataCriacao.ToString("dd/MM/yyyy HH:mm:ss");
                    worksheet.Cells[linha, 16].Value = evento.NumeroRecibo ?? "";
                    worksheet.Cells[linha, 17].Value = evento.IndRetificacao;
                    worksheet.Cells[linha, 18].Value = ObterCodigoTipoEvento(lote.TipoLote, lote);

                    // Colorir status
                    if (evento.StatusEvento?.ToUpper().Contains("ERRO") == true)
                        worksheet.Cells[linha, 13].Style.Font.Color.SetColor(Color.Red);
                    else if (evento.StatusEvento?.ToUpper().Contains("SUCESSO") == true)
                        worksheet.Cells[linha, 13].Style.Font.Color.SetColor(Color.Green);
                    
                    linha++;
                }
            }

            // Aplicar bordas e filtro
            if (linha > 3)
            {
                var dataRange = worksheet.Cells[3, 1, linha - 1, 18];
                dataRange.AutoFilter = true;
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }

            AjustarLarguraColunasEventos(worksheet);
        }

        private void CriarAbaFechamento(ExcelPackage package, LoteBancoInfo lote, EfinanceiraDatabasePersistenceService persistenceService)
        {
            var worksheet = package.Workbook.Worksheets.Add("Fechamento");
            int linha = 1;

            // Título
            worksheet.Cells[linha, 1].Value = "DADOS DO FECHAMENTO";
            worksheet.Cells[linha, 1, linha, 6].Merge = true;
            worksheet.Cells[linha, 1].Style.Font.Size = 14;
            worksheet.Cells[linha, 1].Style.Font.Bold = true;
            worksheet.Cells[linha, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            linha += 2;

            // Informações do Lote (similar à abertura)
            AdicionarLinhaInfo(worksheet, ref linha, "Tipo de Evento:", "Fechamento");
            AdicionarLinhaInfo(worksheet, ref linha, "Código do Tipo:", ObterCodigoTipoEvento(TipoLote.Fechamento, lote));
            AdicionarLinhaInfo(worksheet, ref linha, "ID do Lote:", lote.IdLote.ToString());
            AdicionarLinhaInfo(worksheet, ref linha, "Período:", lote.Periodo ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Número do Lote:", lote.NumeroLote.ToString());
            AdicionarLinhaInfo(worksheet, ref linha, "CNPJ Declarante:", lote.CnpjDeclarante ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Ambiente:", lote.Ambiente ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Protocolo de Envio:", lote.ProtocoloEnvio ?? "Não disponível");
            AdicionarLinhaInfo(worksheet, ref linha, "Status:", lote.Status ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Data de Criação:", lote.DataCriacao.ToString("dd/MM/yyyy HH:mm:ss"));
            if (lote.DataEnvio.HasValue)
                AdicionarLinhaInfo(worksheet, ref linha, "Data de Envio:", lote.DataEnvio.Value.ToString("dd/MM/yyyy HH:mm:ss"));
            if (lote.DataConfirmacao.HasValue)
                AdicionarLinhaInfo(worksheet, ref linha, "Data de Confirmação:", lote.DataConfirmacao.Value.ToString("dd/MM/yyyy HH:mm:ss"));
            AdicionarLinhaInfo(worksheet, ref linha, "É Retificação:", lote.EhRetificacao ? "Sim" : "Não");
            if (lote.EhRetificacao && lote.IdLoteOriginal.HasValue)
                AdicionarLinhaInfo(worksheet, ref linha, "ID Lote Original:", lote.IdLoteOriginal.Value.ToString());

            linha += 2;

            // Respostas do Servidor
            if (lote.CodigoRespostaEnvio.HasValue || lote.CodigoRespostaConsulta.HasValue)
            {
                worksheet.Cells[linha, 1].Value = "RESPOSTAS DO SERVIDOR";
                worksheet.Cells[linha, 1, linha, 6].Merge = true;
                worksheet.Cells[linha, 1].Style.Font.Bold = true;
                linha++;

                if (lote.CodigoRespostaEnvio.HasValue)
                {
                    AdicionarLinhaInfo(worksheet, ref linha, "Código Resposta Envio:", lote.CodigoRespostaEnvio.Value.ToString());
                    AdicionarLinhaInfo(worksheet, ref linha, "Descrição:", lote.DescricaoRespostaEnvio ?? "N/A");
                }

                if (lote.CodigoRespostaConsulta.HasValue)
                {
                    AdicionarLinhaInfo(worksheet, ref linha, "Código Resposta Consulta:", lote.CodigoRespostaConsulta.Value.ToString());
                    AdicionarLinhaInfo(worksheet, ref linha, "Descrição:", lote.DescricaoRespostaConsulta ?? "N/A");
                }
            }

            AjustarLarguraColunas(worksheet);
        }

        private void CriarAbaEventosDetalhados(ExcelPackage package, List<LoteBancoInfo> lotes, EfinanceiraDatabasePersistenceService persistenceService)
        {
            var worksheet = package.Workbook.Worksheets.Add("Eventos Detalhados");
            int linha = 1;

            // Cabeçalhos
            CriarCabecalhoEventosCompleto(worksheet, linha);
            linha++;

            // Eventos apenas dos lotes mais recentes (já filtrados na chamada)
            foreach (var lote in lotes.OrderByDescending(l => l.DataCriacao))
            {
                var eventos = persistenceService.BuscarEventosDoLote(lote.IdLote);
                foreach (var evento in eventos)
                {
                    AdicionarLinhaEventoCompleto(worksheet, linha, evento, lote);
                    linha++;
                }
            }

            // Formatação
            var range = worksheet.Cells[1, 1, linha - 1, 18];
            range.AutoFilter = true;
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

            AjustarLarguraColunasEventos(worksheet);
        }

        private void CriarAbaVazia(ExcelPackage package, string nomeAba, string mensagem)
        {
            var worksheet = package.Workbook.Worksheets.Add(nomeAba);
            worksheet.Cells[1, 1].Value = mensagem;
            worksheet.Cells[1, 1].Style.Font.Size = 12;
            worksheet.Cells[1, 1].Style.Font.Color.SetColor(Color.Gray);
            worksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Column(1).Width = 50;
        }

        private void CriarAbaInformacoesLote(ExcelPackage package, LoteBancoInfo lote)
        {
            var worksheet = package.Workbook.Worksheets.Add("Informações do Lote");
            int linha = 1;

            worksheet.Cells[linha, 1].Value = "INFORMAÇÕES DO LOTE";
            worksheet.Cells[linha, 1, linha, 2].Merge = true;
            worksheet.Cells[linha, 1].Style.Font.Size = 14;
            worksheet.Cells[linha, 1].Style.Font.Bold = true;
            linha += 2;

            AdicionarLinhaInfo(worksheet, ref linha, "ID do Lote:", lote.IdLote.ToString());
            AdicionarLinhaInfo(worksheet, ref linha, "Tipo:", lote.TipoLote.ToString());
            AdicionarLinhaInfo(worksheet, ref linha, "Período:", lote.Periodo ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Número do Lote:", lote.NumeroLote.ToString());
            AdicionarLinhaInfo(worksheet, ref linha, "CNPJ Declarante:", lote.CnpjDeclarante ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Ambiente:", lote.Ambiente ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Protocolo:", lote.ProtocoloEnvio ?? "Não disponível");
            AdicionarLinhaInfo(worksheet, ref linha, "Status:", lote.Status ?? "N/A");
            AdicionarLinhaInfo(worksheet, ref linha, "Quantidade Eventos:", lote.QuantidadeEventos.ToString());
            AdicionarLinhaInfo(worksheet, ref linha, "Data Criação:", lote.DataCriacao.ToString("dd/MM/yyyy HH:mm:ss"));
            if (lote.DataEnvio.HasValue)
                AdicionarLinhaInfo(worksheet, ref linha, "Data Envio:", lote.DataEnvio.Value.ToString("dd/MM/yyyy HH:mm:ss"));
            if (lote.DataConfirmacao.HasValue)
                AdicionarLinhaInfo(worksheet, ref linha, "Data Confirmação:", lote.DataConfirmacao.Value.ToString("dd/MM/yyyy HH:mm:ss"));
            AdicionarLinhaInfo(worksheet, ref linha, "É Retificação:", lote.EhRetificacao ? "Sim" : "Não");
            if (lote.EhRetificacao && lote.IdLoteOriginal.HasValue)
                AdicionarLinhaInfo(worksheet, ref linha, "ID Lote Original:", lote.IdLoteOriginal.Value.ToString());

            AjustarLarguraColunas(worksheet);
        }

        private void CriarAbaEventos(ExcelPackage package, List<EventoBancoInfo> eventos)
        {
            var worksheet = package.Workbook.Worksheets.Add("Eventos");
            int linha = 1;

            CriarCabecalhoEventosCompleto(worksheet, linha);
            linha++;

            foreach (var evento in eventos)
            {
                AdicionarLinhaEventoCompleto(worksheet, linha, evento, null);
                linha++;
            }

            var range = worksheet.Cells[1, 1, linha - 1, 18];
            range.AutoFilter = true;
            AjustarLarguraColunasEventos(worksheet);
        }

        private void CriarCabecalhoEventos(ExcelWorksheet worksheet, int linha)
        {
            worksheet.Cells[linha, 1].Value = "ID Evento";
            worksheet.Cells[linha, 2].Value = "CPF";
            worksheet.Cells[linha, 3].Value = "Nome";
            worksheet.Cells[linha, 4].Value = "Número Conta";
            worksheet.Cells[linha, 5].Value = "Saldo Atual";
            worksheet.Cells[linha, 6].Value = "Tot. Créditos";
            worksheet.Cells[linha, 7].Value = "Tot. Débitos";
            worksheet.Cells[linha, 8].Value = "Status";

            var headerRange = worksheet.Cells[linha, 1, linha, 8];
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            headerRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private void CriarCabecalhoEventosCompleto(ExcelWorksheet worksheet, int linha)
        {
            // Todos os campos conforme o JSON fornecido
            worksheet.Cells[linha, 1].Value = "ID Evento";
            worksheet.Cells[linha, 2].Value = "ID Lote";
            worksheet.Cells[linha, 3].Value = "ID Pessoa";
            worksheet.Cells[linha, 4].Value = "ID Conta";
            worksheet.Cells[linha, 5].Value = "CPF";
            worksheet.Cells[linha, 6].Value = "Nome";
            worksheet.Cells[linha, 7].Value = "Número Conta";
            worksheet.Cells[linha, 8].Value = "Dígito Conta";
            worksheet.Cells[linha, 9].Value = "Saldo Atual";
            worksheet.Cells[linha, 10].Value = "Tot. Créditos";
            worksheet.Cells[linha, 11].Value = "Tot. Débitos";
            worksheet.Cells[linha, 12].Value = "ID Evento XML";
            worksheet.Cells[linha, 13].Value = "Status Evento";
            worksheet.Cells[linha, 14].Value = "Ocorrências JSON";
            worksheet.Cells[linha, 15].Value = "Data Criação";
            worksheet.Cells[linha, 16].Value = "Número Recibo";
            worksheet.Cells[linha, 17].Value = "Ind. Retificação";
            worksheet.Cells[linha, 18].Value = "Tipo Lote";

            var headerRange = worksheet.Cells[linha, 1, linha, 18];
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            headerRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private void AdicionarLinhaEvento(ExcelWorksheet worksheet, int linha, EventoBancoInfo evento)
        {
            worksheet.Cells[linha, 1].Value = evento.IdEventoXml ?? "N/A";
            worksheet.Cells[linha, 2].Value = evento.Cpf ?? "N/A";
            worksheet.Cells[linha, 3].Value = evento.Nome ?? "N/A";
            worksheet.Cells[linha, 4].Value = evento.NumeroConta?.ToString() ?? "N/A";
            worksheet.Cells[linha, 5].Value = evento.SaldoAtual?.ToString("N2") ?? "0,00";
            worksheet.Cells[linha, 5].Style.Numberformat.Format = "#,##0.00";
            worksheet.Cells[linha, 6].Value = evento.TotCreditos?.ToString("N2") ?? "0,00";
            worksheet.Cells[linha, 6].Style.Numberformat.Format = "#,##0.00";
            worksheet.Cells[linha, 7].Value = evento.TotDebitos?.ToString("N2") ?? "0,00";
            worksheet.Cells[linha, 7].Style.Numberformat.Format = "#,##0.00";
            worksheet.Cells[linha, 8].Value = evento.StatusEvento ?? "N/A";
            
            // Colorir status
            if (evento.StatusEvento?.ToUpper().Contains("ERRO") == true)
                worksheet.Cells[linha, 8].Style.Font.Color.SetColor(Color.Red);
            else if (evento.StatusEvento?.ToUpper().Contains("SUCESSO") == true)
                worksheet.Cells[linha, 8].Style.Font.Color.SetColor(Color.Green);
        }

        private void AdicionarLinhaEventoCompleto(ExcelWorksheet worksheet, int linha, EventoBancoInfo evento, LoteBancoInfo lote)
        {
            // Todos os campos conforme o JSON fornecido
            worksheet.Cells[linha, 1].Value = evento.IdEvento;
            worksheet.Cells[linha, 2].Value = evento.IdLote;
            worksheet.Cells[linha, 3].Value = evento.IdPessoa?.ToString() ?? "";
            worksheet.Cells[linha, 4].Value = evento.IdConta?.ToString() ?? "";
            worksheet.Cells[linha, 5].Value = evento.Cpf ?? "";
            worksheet.Cells[linha, 6].Value = evento.Nome ?? "";
            worksheet.Cells[linha, 7].Value = evento.NumeroConta?.ToString() ?? "";
            worksheet.Cells[linha, 8].Value = evento.DigitoConta ?? "";
            
            // Saldo Atual - usar o valor real do banco, não um padrão
            if (evento.SaldoAtual.HasValue)
            {
                worksheet.Cells[linha, 9].Value = evento.SaldoAtual.Value;
                worksheet.Cells[linha, 9].Style.Numberformat.Format = "#,##0.00";
            }
            else
            {
                worksheet.Cells[linha, 9].Value = "";
            }
            
            // Tot. Créditos
            if (evento.TotCreditos.HasValue)
            {
                worksheet.Cells[linha, 10].Value = evento.TotCreditos.Value;
                worksheet.Cells[linha, 10].Style.Numberformat.Format = "#,##0.00";
            }
            else
            {
                worksheet.Cells[linha, 10].Value = "";
            }
            
            // Tot. Débitos
            if (evento.TotDebitos.HasValue)
            {
                worksheet.Cells[linha, 11].Value = evento.TotDebitos.Value;
                worksheet.Cells[linha, 11].Style.Numberformat.Format = "#,##0.00";
            }
            else
            {
                worksheet.Cells[linha, 11].Value = "";
            }
            
            worksheet.Cells[linha, 12].Value = evento.IdEventoXml ?? "";
            worksheet.Cells[linha, 13].Value = evento.StatusEvento ?? "";
            worksheet.Cells[linha, 14].Value = evento.OcorrenciasJson ?? "";
            worksheet.Cells[linha, 15].Value = evento.DataCriacao.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            worksheet.Cells[linha, 16].Value = evento.NumeroRecibo ?? "";
            worksheet.Cells[linha, 17].Value = evento.IndRetificacao;
            worksheet.Cells[linha, 18].Value = lote != null ? ObterCodigoTipoEvento(lote.TipoLote, lote) : "";

            // Colorir status
            if (evento.StatusEvento?.ToUpper().Contains("ERRO") == true)
                worksheet.Cells[linha, 13].Style.Font.Color.SetColor(Color.Red);
            else if (evento.StatusEvento?.ToUpper().Contains("SUCESSO") == true)
                worksheet.Cells[linha, 13].Style.Font.Color.SetColor(Color.Green);
        }

        private void AdicionarLinhaInfo(ExcelWorksheet worksheet, ref int linha, string label, string value)
        {
            worksheet.Cells[linha, 1].Value = label;
            worksheet.Cells[linha, 1].Style.Font.Bold = true;
            worksheet.Cells[linha, 2].Value = value;
            linha++;
        }

        private void AjustarLarguraColunas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 25;
            worksheet.Column(2).Width = 30;
            worksheet.Column(3).Width = 20;
            worksheet.Column(4).Width = 20;
        }

        private void AjustarLarguraColunasEventos(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 12;  // ID Evento
            worksheet.Column(2).Width = 12;  // ID Lote
            worksheet.Column(3).Width = 12;  // ID Pessoa
            worksheet.Column(4).Width = 12;  // ID Conta
            worksheet.Column(5).Width = 15;  // CPF
            worksheet.Column(6).Width = 30;  // Nome
            worksheet.Column(7).Width = 15;  // Número Conta
            worksheet.Column(8).Width = 12;  // Dígito Conta
            worksheet.Column(9).Width = 15;  // Saldo Atual
            worksheet.Column(10).Width = 15; // Tot. Créditos
            worksheet.Column(11).Width = 15; // Tot. Débitos
            worksheet.Column(12).Width = 20; // ID Evento XML
            worksheet.Column(13).Width = 20; // Status Evento
            worksheet.Column(14).Width = 30; // Ocorrências JSON
            worksheet.Column(15).Width = 25; // Data Criação
            worksheet.Column(16).Width = 20; // Número Recibo
            worksheet.Column(17).Width = 15; // Ind. Retificação
            worksheet.Column(18).Width = 15; // Tipo Lote
        }
    }
}
