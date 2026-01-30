using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Npgsql;
using ExemploAssinadorXML.Models;

namespace ExemploAssinadorXML.Services
{
    public class EfinanceiraDatabasePersistenceService
    {
        private readonly string _connectionString;

        public EfinanceiraDatabasePersistenceService()
        {
            // Usar as mesmas credenciais do EfinanceiraDatabaseService
            const string DB_HOST = "10.30.0.21";
            const string DB_PORT = "5432";
            const string DB_NAME = "bscash";
            const string DB_USER = "nickolas.oliveira";
            const string DB_PASSWORD = "1QclT+-IVB2B";

            _connectionString = $"Host={DB_HOST};Port={DB_PORT};Database={DB_NAME};Username={DB_USER};Password={DB_PASSWORD};Timeout=30;Command Timeout=60;";
        }

        public EfinanceiraDatabasePersistenceService(string connectionString)
        {
            _connectionString = connectionString;
        }

        /// <summary>
        /// Registra um lote no banco de dados
        /// </summary>
        public long RegistrarLote(TipoLote tipo, string periodo, int quantidadeEventos, string cnpjDeclarante, 
            string caminhoArquivoXml, string caminhoArquivoAssinado, string caminhoArquivoCriptografado,
            string ambiente, int? numeroLote = null, long? idLoteOriginal = null)
        {
            try
            {
                int semestre = CalcularSemestre(periodo);
                string status = "GERADO";
                string hashConteudo = CalcularHashArquivo(caminhoArquivoXml);

                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();

                    // Se não informou número do lote, buscar o próximo
                    if (!numeroLote.HasValue)
                    {
                        numeroLote = ObterProximoNumeroLote(connection, periodo, semestre);
                    }

                    string sql = @"
                        INSERT INTO efinanceira.tb_efinanceira_lote 
                        (periodo, semestre, numerolote, quantidadeeventos, cnpjdeclarante, hashconteudo,
                         caminhoarquivolotexml, caminhoarquivoloteassinadoxml, caminhoarquivolotecriptografadoxml,
                         status, ambiente, datacriacao, version, id_lote_original)
                        VALUES 
                        (@periodo, @semestre, @numerolote, @quantidadeeventos, @cnpjdeclarante, @hashconteudo,
                         @caminhoarquivolotexml, @caminhoarquivoloteassinadoxml, @caminhoarquivolotecriptografadoxml,
                         @status, @ambiente, @datacriacao, 0, @idloteoriginal)
                        RETURNING idlote";

                    using (var command = new NpgsqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@periodo", periodo ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@semestre", semestre);
                        command.Parameters.AddWithValue("@numerolote", numeroLote.Value);
                        command.Parameters.AddWithValue("@quantidadeeventos", quantidadeEventos);
                        command.Parameters.AddWithValue("@cnpjdeclarante", cnpjDeclarante ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@hashconteudo", hashConteudo ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@caminhoarquivolotexml", caminhoArquivoXml ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@caminhoarquivoloteassinadoxml", caminhoArquivoAssinado ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@caminhoarquivolotecriptografadoxml", caminhoArquivoCriptografado ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@status", status);
                        command.Parameters.AddWithValue("@ambiente", ambiente);
                        command.Parameters.AddWithValue("@datacriacao", DateTime.Now);
                        command.Parameters.AddWithValue("@idloteoriginal", idLoteOriginal ?? (object)DBNull.Value);

                        long idLote = (long)command.ExecuteScalar();
                        return idLote;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao registrar lote no banco: {ex.Message}");
                throw new Exception($"Erro ao registrar lote no banco de dados: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Atualiza o status e informações de um lote
        /// </summary>
        public void AtualizarLote(long idLote, string status, string protocoloEnvio = null, 
            int? codigoRespostaEnvio = null, string descricaoRespostaEnvio = null, string xmlRespostaEnvio = null,
            int? codigoRespostaConsulta = null, string descricaoRespostaConsulta = null, string xmlRespostaConsulta = null,
            DateTime? dataEnvio = null, DateTime? dataConfirmacao = null, string ultimoErro = null)
        {
            try
            {
                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();

                    string sql = @"
                        UPDATE efinanceira.tb_efinanceira_lote 
                        SET status = @status,
                            dataatualizacao = @dataatualizacao,
                            version = version + 1";

                    var parameters = new List<NpgsqlParameter>
                    {
                        new NpgsqlParameter("@status", status),
                        new NpgsqlParameter("@dataatualizacao", DateTime.Now)
                    };

                    if (!string.IsNullOrEmpty(protocoloEnvio))
                    {
                        sql += ", protocoloenvio = @protocoloenvio";
                        parameters.Add(new NpgsqlParameter("@protocoloenvio", protocoloEnvio));
                    }

                    if (codigoRespostaEnvio.HasValue)
                    {
                        sql += ", codigorespostaenvio = @codigorespostaenvio";
                        parameters.Add(new NpgsqlParameter("@codigorespostaenvio", codigoRespostaEnvio.Value));
                    }

                    if (!string.IsNullOrEmpty(descricaoRespostaEnvio))
                    {
                        sql += ", descricaorespostaenvio = @descricaorespostaenvio";
                        parameters.Add(new NpgsqlParameter("@descricaorespostaenvio", descricaoRespostaEnvio));
                    }

                    if (!string.IsNullOrEmpty(xmlRespostaEnvio))
                    {
                        sql += ", xmlrespostaenvio = @xmlrespostaenvio";
                        parameters.Add(new NpgsqlParameter("@xmlrespostaenvio", xmlRespostaEnvio));
                    }

                    if (codigoRespostaConsulta.HasValue)
                    {
                        sql += ", codigorespostaconsulta = @codigorespostaconsulta";
                        parameters.Add(new NpgsqlParameter("@codigorespostaconsulta", codigoRespostaConsulta.Value));
                    }

                    if (!string.IsNullOrEmpty(descricaoRespostaConsulta))
                    {
                        sql += ", descricaorespostaconsulta = @descricaorespostaconsulta";
                        parameters.Add(new NpgsqlParameter("@descricaorespostaconsulta", descricaoRespostaConsulta));
                    }

                    if (!string.IsNullOrEmpty(xmlRespostaConsulta))
                    {
                        sql += ", xmlrespostaconsulta = @xmlrespostaconsulta";
                        parameters.Add(new NpgsqlParameter("@xmlrespostaconsulta", xmlRespostaConsulta));
                    }

                    if (dataEnvio.HasValue)
                    {
                        sql += ", dataenvio = @dataenvio";
                        parameters.Add(new NpgsqlParameter("@dataenvio", dataEnvio.Value));
                    }

                    if (dataConfirmacao.HasValue)
                    {
                        sql += ", dataconfirmacao = @dataconfirmacao";
                        parameters.Add(new NpgsqlParameter("@dataconfirmacao", dataConfirmacao.Value));
                    }

                    if (!string.IsNullOrEmpty(ultimoErro))
                    {
                        sql += ", ultimoerro = @ultimoerro";
                        parameters.Add(new NpgsqlParameter("@ultimoerro", ultimoErro));
                    }

                    sql += " WHERE idlote = @idlote";

                    using (var command = new NpgsqlCommand(sql, connection))
                    {
                        command.Parameters.AddRange(parameters.ToArray());
                        command.Parameters.AddWithValue("@idlote", idLote);
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao atualizar lote no banco: {ex.Message}");
                throw new Exception($"Erro ao atualizar lote no banco de dados: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Normaliza o CPF removendo caracteres especiais e limitando a 11 caracteres
        /// </summary>
        private string NormalizarCpf(string cpf)
        {
            if (string.IsNullOrWhiteSpace(cpf))
                return null;

            // Remove caracteres não numéricos
            string cpfLimpo = new string(cpf.Where(char.IsDigit).ToArray());

            // Limita a 11 caracteres (tamanho máximo do campo no banco)
            if (cpfLimpo.Length > 11)
            {
                cpfLimpo = cpfLimpo.Substring(0, 11);
            }

            return string.IsNullOrWhiteSpace(cpfLimpo) ? null : cpfLimpo;
        }

        /// <summary>
        /// Registra múltiplos eventos de uma lista de pessoas
        /// </summary>
        public void RegistrarEventosDoLote(long idLote, List<DadosPessoaConta> pessoas, string idEventoPrefix = "ID", int indRetificacao = 0)
        {
            foreach (var pessoa in pessoas)
            {
                string idEventoXml = $"{idEventoPrefix}{pessoa.IdPessoa.ToString().PadLeft(18, '0')}";
                if (idEventoXml.Length > 50) idEventoXml = idEventoXml.Substring(idEventoXml.Length - 50);

                long? numeroConta = null;
                if (!string.IsNullOrEmpty(pessoa.NumeroConta) && long.TryParse(pessoa.NumeroConta, out long numConta))
                {
                    numeroConta = numConta;
                }

                // Normalizar CPF antes de registrar (já será normalizado novamente em RegistrarEvento, mas é bom fazer aqui também)
                string cpfNormalizado = NormalizarCpf(pessoa.Cpf);

                RegistrarEvento(
                    idLote,
                    pessoa.IdPessoa,
                    pessoa.IdConta,
                    cpfNormalizado,
                    pessoa.Nome,
                    numeroConta,
                    pessoa.DigitoConta,
                    pessoa.SaldoAtual,
                    pessoa.TotCreditos,
                    pessoa.TotDebitos,
                    idEventoXml,
                    "GERADO",
                    null,
                    null,
                    indRetificacao
                );
            }
        }

        /// <summary>
        /// Registra um evento no banco de dados
        /// </summary>
        public long RegistrarEvento(long idLote, long? idPessoa, long? idConta, string cpf, string nome,
            long? numeroConta, string digitoConta, decimal? saldoAtual, decimal? totCreditos, decimal? totDebitos,
            string idEventoXml, string statusEvento, string ocorrenciasJson = null, string numeroRecibo = null, int indRetificacao = 0)
        {
            try
            {
                // Normalizar CPF antes de salvar
                string cpfNormalizado = NormalizarCpf(cpf);

                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();

                    string sql = @"
                        INSERT INTO efinanceira.tb_efinanceira_evento 
                        (idlote, idpessoa, idconta, cpf, nome, numeroconta, digitoconta, saldoatual,
                         totcreditos, totdebitos, ideventoxml, statusevento, ocorrenciasefinanceirajson,
                         datacriacao, numerorecibo, indretificacao)
                        VALUES 
                        (@idlote, @idpessoa, @idconta, @cpf, @nome, @numeroconta, @digitoconta, @saldoatual,
                         @totcreditos, @totdebitos, @ideventoxml, @statusevento, @ocorrenciasjson,
                         @datacriacao, @numerorecibo, @indretificacao)
                        RETURNING idevento";

                    using (var command = new NpgsqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@idlote", idLote);
                        command.Parameters.AddWithValue("@idpessoa", idPessoa ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@idconta", idConta ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@cpf", cpfNormalizado ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@nome", nome ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@numeroconta", numeroConta ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@digitoconta", digitoConta ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@saldoatual", saldoAtual ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@totcreditos", totCreditos ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@totdebitos", totDebitos ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@ideventoxml", idEventoXml ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@statusevento", statusEvento ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@ocorrenciasjson", ocorrenciasJson ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@datacriacao", DateTime.Now);
                        command.Parameters.AddWithValue("@numerorecibo", numeroRecibo ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@indretificacao", indRetificacao);

                        long idEvento = (long)command.ExecuteScalar();
                        return idEvento;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao registrar evento no banco: {ex.Message}");
                throw new Exception($"Erro ao registrar evento no banco de dados: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Atualiza o status e ocorrências de um evento
        /// </summary>
        public void AtualizarEvento(long idEvento, string statusEvento, string ocorrenciasJson = null, string numeroRecibo = null)
        {
            try
            {
                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();

                    string sql = @"
                        UPDATE efinanceira.tb_efinanceira_evento 
                        SET statusevento = @statusevento";

                    var parameters = new List<NpgsqlParameter>
                    {
                        new NpgsqlParameter("@statusevento", statusEvento ?? (object)DBNull.Value)
                    };

                    if (!string.IsNullOrEmpty(ocorrenciasJson))
                    {
                        sql += ", ocorrenciasefinanceirajson = @ocorrenciasjson";
                        parameters.Add(new NpgsqlParameter("@ocorrenciasjson", ocorrenciasJson));
                    }

                    if (!string.IsNullOrEmpty(numeroRecibo))
                    {
                        sql += ", numerorecibo = @numerorecibo";
                        parameters.Add(new NpgsqlParameter("@numerorecibo", numeroRecibo));
                    }

                    sql += " WHERE idevento = @idevento";

                    using (var command = new NpgsqlCommand(sql, connection))
                    {
                        command.Parameters.AddRange(parameters.ToArray());
                        command.Parameters.AddWithValue("@idevento", idEvento);
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao atualizar evento no banco: {ex.Message}");
                throw new Exception($"Erro ao atualizar evento no banco de dados: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Registra um log de processamento do lote
        /// </summary>
        public void RegistrarLogLote(long idLote, string etapa, string mensagem, string payloadCurto = null)
        {
            try
            {
                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();

                    string sql = @"
                        INSERT INTO efinanceira.tb_efinanceira_lote_log 
                        (idlote, etapa, mensagem, payloadcurto, timestamp)
                        VALUES 
                        (@idlote, @etapa, @mensagem, @payloadcurto, @timestamp)";

                    using (var command = new NpgsqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@idlote", idLote);
                        command.Parameters.AddWithValue("@etapa", etapa);
                        command.Parameters.AddWithValue("@mensagem", mensagem);
                        command.Parameters.AddWithValue("@payloadcurto", payloadCurto ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@timestamp", DateTime.Now);
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao registrar log do lote: {ex.Message}");
                // Não lançar exceção para não interromper o processamento
            }
        }

        /// <summary>
        /// Busca lotes do banco de dados com filtro de data
        /// </summary>
        public List<LoteBancoInfo> BuscarLotes(DateTime? dataInicio = null, DateTime? dataFim = null, int? limite = null)
        {
            List<LoteBancoInfo> lotes = new List<LoteBancoInfo>();

            try
            {
                // Ajustar datas se necessário
                if (dataInicio.HasValue && !dataFim.HasValue)
                {
                    dataFim = dataInicio.Value.AddDays(1).AddSeconds(-1);
                }
                else if (!dataInicio.HasValue && dataFim.HasValue)
                {
                    dataInicio = dataFim.Value.Date;
                }

                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();

                    string sql = @"
                        SELECT 
                            l.idlote,
                            l.periodo,
                            l.semestre,
                            l.numerolote,
                            l.quantidadeeventos,
                            l.cnpjdeclarante,
                            l.protocoloenvio,
                            l.status,
                            l.ambiente,
                            l.codigorespostaenvio,
                            l.descricaorespostaenvio,
                            l.codigorespostaconsulta,
                            l.descricaorespostaconsulta,
                            l.datacriacao,
                            l.dataenvio,
                            l.dataconfirmacao,
                            l.id_lote_original,
                            l.caminhoarquivolotexml,
                            COUNT(DISTINCT e.idevento) as total_eventos_registrados,
                            COUNT(DISTINCT CASE WHEN e.cpf IS NOT NULL AND e.cpf != '' THEN e.cpf END) as total_eventos_com_cpf,
                            COUNT(DISTINCT CASE WHEN e.statusevento = 'ERRO' THEN e.idevento END) as total_eventos_com_erro,
                            COUNT(DISTINCT CASE WHEN e.statusevento = 'SUCESSO' THEN e.idevento END) as total_eventos_sucesso
                        FROM efinanceira.tb_efinanceira_lote l
                        LEFT JOIN efinanceira.tb_efinanceira_evento e ON e.idlote = l.idlote";

                    // Adicionar filtro de data apenas se informado
                    if (dataInicio.HasValue && dataFim.HasValue)
                    {
                        sql += @"
                        WHERE l.datacriacao >= @datainicio 
                          AND l.datacriacao <= @datafim";
                    }

                    sql += @"
                        GROUP BY l.idlote, l.periodo, l.semestre, l.numerolote, l.quantidadeeventos,
                                 l.cnpjdeclarante, l.protocoloenvio, l.status, l.ambiente,
                                 l.codigorespostaenvio, l.descricaorespostaenvio, l.codigorespostaconsulta,
                                 l.descricaorespostaconsulta, l.datacriacao, l.dataenvio, l.dataconfirmacao,
                                 l.id_lote_original, l.caminhoarquivolotexml
                        ORDER BY l.datacriacao DESC";

                    if (limite.HasValue)
                    {
                        sql += " LIMIT @limite";
                    }

                    using (var command = new NpgsqlCommand(sql, connection))
                    {
                        if (dataInicio.HasValue && dataFim.HasValue)
                        {
                            command.Parameters.AddWithValue("@datainicio", dataInicio.Value);
                            command.Parameters.AddWithValue("@datafim", dataFim.Value);
                        }
                        if (limite.HasValue)
                        {
                            command.Parameters.AddWithValue("@limite", limite.Value);
                        }

                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var lote = new LoteBancoInfo
                                {
                                    IdLote = GetSafeLong(reader, "idlote"),
                                    Periodo = GetSafeString(reader, "periodo"),
                                    Semestre = GetSafeInt(reader, "semestre"),
                                    NumeroLote = GetSafeInt(reader, "numerolote"),
                                    QuantidadeEventos = GetSafeInt(reader, "quantidadeeventos"),
                                    CnpjDeclarante = GetSafeString(reader, "cnpjdeclarante"),
                                    ProtocoloEnvio = GetSafeString(reader, "protocoloenvio"),
                                    Status = GetSafeString(reader, "status"),
                                    Ambiente = GetSafeString(reader, "ambiente"),
                                    CodigoRespostaEnvio = GetSafeIntNullable(reader, "codigorespostaenvio"),
                                    DescricaoRespostaEnvio = GetSafeString(reader, "descricaorespostaenvio"),
                                    CodigoRespostaConsulta = GetSafeIntNullable(reader, "codigorespostaconsulta"),
                                    DescricaoRespostaConsulta = GetSafeString(reader, "descricaorespostaconsulta"),
                                    DataCriacao = GetSafeDateTime(reader, "datacriacao"),
                                    DataEnvio = GetSafeDateTimeNullable(reader, "dataenvio"),
                                    DataConfirmacao = GetSafeDateTimeNullable(reader, "dataconfirmacao"),
                                    IdLoteOriginal = GetSafeLongNullable(reader, "id_lote_original"),
                                    CaminhoArquivoXml = GetSafeString(reader, "caminhoarquivolotexml"),
                                    TotalEventosRegistrados = GetSafeInt(reader, "total_eventos_registrados"),
                                    TotalEventosComCpf = GetSafeInt(reader, "total_eventos_com_cpf"),
                                    TotalEventosComErro = GetSafeInt(reader, "total_eventos_com_erro"),
                                    TotalEventosSucesso = GetSafeInt(reader, "total_eventos_sucesso")
                                };

                                // Determinar tipo do lote baseado no status e dados
                                lote.TipoLote = DeterminarTipoLote(lote);

                                lotes.Add(lote);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao buscar lotes do banco: {ex.Message}");
                throw new Exception($"Erro ao buscar lotes do banco de dados: {ex.Message}", ex);
            }

            return lotes;
        }

        /// <summary>
        /// Busca eventos de um lote específico
        /// </summary>
        public List<EventoBancoInfo> BuscarEventosDoLote(long idLote)
        {
            List<EventoBancoInfo> eventos = new List<EventoBancoInfo>();

            try
            {
                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();

                    string sql = @"
                        SELECT 
                            e.idevento,
                            e.idlote,
                            e.idpessoa,
                            e.idconta,
                            e.cpf,
                            e.nome,
                            e.numeroconta,
                            e.digitoconta,
                            e.saldoatual,
                            e.totcreditos,
                            e.totdebitos,
                            e.ideventoxml,
                            e.statusevento,
                            e.ocorrenciasefinanceirajson,
                            e.datacriacao,
                            e.numerorecibo,
                            e.indretificacao
                        FROM efinanceira.tb_efinanceira_evento e
                        WHERE e.idlote = @idlote
                        ORDER BY e.idevento";

                    using (var command = new NpgsqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@idlote", idLote);

                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var evento = new EventoBancoInfo
                                {
                                    IdEvento = GetSafeLong(reader, "idevento"),
                                    IdLote = GetSafeLong(reader, "idlote"),
                                    IdPessoa = GetSafeLongNullable(reader, "idpessoa"),
                                    IdConta = GetSafeLongNullable(reader, "idconta"),
                                    Cpf = GetSafeString(reader, "cpf"),
                                    Nome = GetSafeString(reader, "nome"),
                                    NumeroConta = GetSafeLongNullable(reader, "numeroconta"),
                                    DigitoConta = GetSafeString(reader, "digitoconta"),
                                    SaldoAtual = GetSafeDecimalNullable(reader, "saldoatual"),
                                    TotCreditos = GetSafeDecimalNullable(reader, "totcreditos"),
                                    TotDebitos = GetSafeDecimalNullable(reader, "totdebitos"),
                                    IdEventoXml = GetSafeString(reader, "ideventoxml"),
                                    StatusEvento = GetSafeString(reader, "statusevento"),
                                    OcorrenciasJson = GetSafeString(reader, "ocorrenciasefinanceirajson"),
                                    DataCriacao = GetSafeDateTime(reader, "datacriacao"),
                                    NumeroRecibo = GetSafeString(reader, "numerorecibo"),
                                    IndRetificacao = GetSafeInt(reader, "indretificacao")
                                };

                                eventos.Add(evento);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao buscar eventos do lote: {ex.Message}");
                throw new Exception($"Erro ao buscar eventos do lote: {ex.Message}", ex);
            }

            return eventos;
        }

        /// <summary>
        /// Busca um lote pelo protocolo
        /// </summary>
        public LoteBancoInfo BuscarLotePorProtocolo(string protocolo)
        {
            try
            {
                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();

                    string sql = @"
                        SELECT 
                            l.idlote,
                            l.periodo,
                            l.semestre,
                            l.numerolote,
                            l.quantidadeeventos,
                            l.cnpjdeclarante,
                            l.protocoloenvio,
                            l.status,
                            l.ambiente,
                            l.codigorespostaenvio,
                            l.descricaorespostaenvio,
                            l.codigorespostaconsulta,
                            l.descricaorespostaconsulta,
                            l.datacriacao,
                            l.dataenvio,
                            l.dataconfirmacao,
                            l.id_lote_original,
                            l.caminhoarquivolotexml
                        FROM efinanceira.tb_efinanceira_lote l
                        WHERE l.protocoloenvio = @protocolo
                        ORDER BY l.datacriacao DESC
                        LIMIT 1";

                    using (var command = new NpgsqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@protocolo", protocolo);

                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                var lote = new LoteBancoInfo
                                {
                                    IdLote = GetSafeLong(reader, "idlote"),
                                    Periodo = GetSafeString(reader, "periodo"),
                                    Semestre = GetSafeInt(reader, "semestre"),
                                    NumeroLote = GetSafeInt(reader, "numerolote"),
                                    QuantidadeEventos = GetSafeInt(reader, "quantidadeeventos"),
                                    CnpjDeclarante = GetSafeString(reader, "cnpjdeclarante"),
                                    ProtocoloEnvio = GetSafeString(reader, "protocoloenvio"),
                                    Status = GetSafeString(reader, "status"),
                                    Ambiente = GetSafeString(reader, "ambiente"),
                                    CodigoRespostaEnvio = GetSafeIntNullable(reader, "codigorespostaenvio"),
                                    DescricaoRespostaEnvio = GetSafeString(reader, "descricaorespostaenvio"),
                                    CodigoRespostaConsulta = GetSafeIntNullable(reader, "codigorespostaconsulta"),
                                    DescricaoRespostaConsulta = GetSafeString(reader, "descricaorespostaconsulta"),
                                    DataCriacao = GetSafeDateTime(reader, "datacriacao"),
                                    DataEnvio = GetSafeDateTimeNullable(reader, "dataenvio"),
                                    DataConfirmacao = GetSafeDateTimeNullable(reader, "dataconfirmacao"),
                                    IdLoteOriginal = GetSafeLongNullable(reader, "id_lote_original"),
                                    CaminhoArquivoXml = GetSafeString(reader, "caminhoarquivolotexml")
                                };
                                
                                lote.TipoLote = DeterminarTipoLote(lote);
                                return lote;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao buscar lote por protocolo: {ex.Message}");
            }

            return null;
        }

        // Métodos auxiliares
        private int CalcularSemestre(string periodo)
        {
            if (string.IsNullOrEmpty(periodo) || periodo.Length < 6)
                return 0;

            int mes = int.Parse(periodo.Substring(4, 2));
            if (mes == 1 || mes == 6) return 1; // Primeiro semestre
            if (mes == 2 || mes == 12) return 2; // Segundo semestre
            return 0;
        }

        private int ObterProximoNumeroLote(NpgsqlConnection connection, string periodo, int semestre)
        {
            string sql = @"
                SELECT COALESCE(MAX(numerolote), 0) + 1
                FROM efinanceira.tb_efinanceira_lote
                WHERE periodo = @periodo AND semestre = @semestre";

            using (var command = new NpgsqlCommand(sql, connection))
            {
                command.Parameters.AddWithValue("@periodo", periodo);
                command.Parameters.AddWithValue("@semestre", semestre);
                object result = command.ExecuteScalar();
                return result != null && result != DBNull.Value ? Convert.ToInt32(result) : 1;
            }
        }

        private string CalcularHashArquivo(string caminhoArquivo)
        {
            try
            {
                if (string.IsNullOrEmpty(caminhoArquivo) || !File.Exists(caminhoArquivo))
                    return null;

                using (var sha256 = SHA256.Create())
                {
                    using (var stream = File.OpenRead(caminhoArquivo))
                    {
                        byte[] hash = sha256.ComputeHash(stream);
                        return BitConverter.ToString(hash).Replace("-", "").ToLower();
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        private TipoLote DeterminarTipoLote(LoteBancoInfo lote)
        {
            // Tentar determinar pelo caminho do arquivo
            if (!string.IsNullOrEmpty(lote.CaminhoArquivoXml))
            {
                string nomeArquivo = Path.GetFileName(lote.CaminhoArquivoXml).ToLower();
                if (nomeArquivo.Contains("abertura"))
                {
                    return TipoLote.Abertura;
                }
                else if (nomeArquivo.Contains("fechamento"))
                {
                    return TipoLote.Fechamento;
                }
                else if (nomeArquivo.Contains("movimentacao") || nomeArquivo.Contains("movimentação"))
                {
                    return TipoLote.Movimentacao;
                }
            }
            
            // Se não conseguir determinar, retorna Movimentacao como padrão (mais comum)
            return TipoLote.Movimentacao;
        }

        // Métodos auxiliares para leitura segura
        private string GetSafeString(NpgsqlDataReader reader, string columnName, string defaultValue = "")
        {
            try
            {
                int ordinal = reader.GetOrdinal(columnName);
                return reader.IsDBNull(ordinal) ? defaultValue : reader.GetString(ordinal);
            }
            catch
            {
                return defaultValue;
            }
        }

        private long GetSafeLong(NpgsqlDataReader reader, string columnName, long defaultValue = 0)
        {
            try
            {
                int ordinal = reader.GetOrdinal(columnName);
                return reader.IsDBNull(ordinal) ? defaultValue : reader.GetInt64(ordinal);
            }
            catch
            {
                return defaultValue;
            }
        }

        private long? GetSafeLongNullable(NpgsqlDataReader reader, string columnName)
        {
            try
            {
                int ordinal = reader.GetOrdinal(columnName);
                return reader.IsDBNull(ordinal) ? (long?)null : reader.GetInt64(ordinal);
            }
            catch
            {
                return null;
            }
        }

        private int GetSafeInt(NpgsqlDataReader reader, string columnName, int defaultValue = 0)
        {
            try
            {
                int ordinal = reader.GetOrdinal(columnName);
                return reader.IsDBNull(ordinal) ? defaultValue : reader.GetInt32(ordinal);
            }
            catch
            {
                return defaultValue;
            }
        }

        private int? GetSafeIntNullable(NpgsqlDataReader reader, string columnName)
        {
            try
            {
                int ordinal = reader.GetOrdinal(columnName);
                return reader.IsDBNull(ordinal) ? (int?)null : reader.GetInt32(ordinal);
            }
            catch
            {
                return null;
            }
        }

        private decimal? GetSafeDecimalNullable(NpgsqlDataReader reader, string columnName)
        {
            try
            {
                int ordinal = reader.GetOrdinal(columnName);
                return reader.IsDBNull(ordinal) ? (decimal?)null : reader.GetDecimal(ordinal);
            }
            catch
            {
                return null;
            }
        }

        private DateTime GetSafeDateTime(NpgsqlDataReader reader, string columnName)
        {
            try
            {
                int ordinal = reader.GetOrdinal(columnName);
                return reader.IsDBNull(ordinal) ? DateTime.MinValue : reader.GetDateTime(ordinal);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        private DateTime? GetSafeDateTimeNullable(NpgsqlDataReader reader, string columnName)
        {
            try
            {
                int ordinal = reader.GetOrdinal(columnName);
                return reader.IsDBNull(ordinal) ? (DateTime?)null : reader.GetDateTime(ordinal);
            }
            catch
            {
                return null;
            }
        }
    }

    public class LoteBancoInfo
    {
        public long IdLote { get; set; }
        public string Periodo { get; set; }
        public int Semestre { get; set; }
        public int NumeroLote { get; set; }
        public int QuantidadeEventos { get; set; }
        public string CnpjDeclarante { get; set; }
        public string ProtocoloEnvio { get; set; }
        public string Status { get; set; }
        public string Ambiente { get; set; }
        public int? CodigoRespostaEnvio { get; set; }
        public string DescricaoRespostaEnvio { get; set; }
        public int? CodigoRespostaConsulta { get; set; }
        public string DescricaoRespostaConsulta { get; set; }
        public DateTime DataCriacao { get; set; }
        public DateTime? DataEnvio { get; set; }
        public DateTime? DataConfirmacao { get; set; }
        public long? IdLoteOriginal { get; set; }
        public string CaminhoArquivoXml { get; set; }
        public TipoLote TipoLote { get; set; }
        public int TotalEventosRegistrados { get; set; }
        public int TotalEventosComCpf { get; set; }
        public int TotalEventosComErro { get; set; }
        public int TotalEventosSucesso { get; set; }
        public bool EhRetificacao => IdLoteOriginal.HasValue;
    }

    public class EventoBancoInfo
    {
        public long IdEvento { get; set; }
        public long IdLote { get; set; }
        public long? IdPessoa { get; set; }
        public long? IdConta { get; set; }
        public string Cpf { get; set; }
        public string Nome { get; set; }
        public long? NumeroConta { get; set; }
        public string DigitoConta { get; set; }
        public decimal? SaldoAtual { get; set; }
        public decimal? TotCreditos { get; set; }
        public decimal? TotDebitos { get; set; }
        public string IdEventoXml { get; set; }
        public string StatusEvento { get; set; }
        public string OcorrenciasJson { get; set; }
        public DateTime DataCriacao { get; set; }
        public string NumeroRecibo { get; set; }
        public int IndRetificacao { get; set; }
        public bool EhRetificacao => IndRetificacao > 0;
    }
}
