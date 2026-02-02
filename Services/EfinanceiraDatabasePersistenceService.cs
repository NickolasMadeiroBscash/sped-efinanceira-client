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
    public class EfinanceiraDatabasePersistenceException : InvalidOperationException
    {
        public EfinanceiraDatabasePersistenceException(string message) : base(message)
        {
        }

        public EfinanceiraDatabasePersistenceException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }

    public class RegistrarLoteDto
    {
        public TipoLote Tipo { get; set; }
        public string Periodo { get; set; }
        public int QuantidadeEventos { get; set; }
        public string CnpjDeclarante { get; set; }
        public string CaminhoArquivoXml { get; set; }
        public string CaminhoArquivoAssinado { get; set; }
        public string CaminhoArquivoCriptografado { get; set; }
        public string Ambiente { get; set; }
        public int? NumeroLote { get; set; }
        public long? IdLoteOriginal { get; set; }
    }

    public class AtualizarLoteDto
    {
        public string Status { get; set; }
        public string ProtocoloEnvio { get; set; }
        public int? CodigoRespostaEnvio { get; set; }
        public string DescricaoRespostaEnvio { get; set; }
        public string XmlRespostaEnvio { get; set; }
        public int? CodigoRespostaConsulta { get; set; }
        public string DescricaoRespostaConsulta { get; set; }
        public string XmlRespostaConsulta { get; set; }
        public DateTime? DataEnvio { get; set; }
        public DateTime? DataConfirmacao { get; set; }
        public string UltimoErro { get; set; }
    }

    public class RegistrarEventoDto
    {
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
        public string NumeroRecibo { get; set; }
        public int IndRetificacao { get; set; }
    }

    public class EfinanceiraDatabasePersistenceService
    {
        private readonly string _connectionString;
        private const string PARAM_ID_LOTE = "@idlote";

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
        /// Registra um lote no banco de dados (sobrecarga para compatibilidade)
        /// </summary>
        public long RegistrarLote(TipoLote tipo, string periodo, int quantidadeEventos, string cnpjDeclarante, 
            string caminhoArquivoXml, string caminhoArquivoAssinado, string caminhoArquivoCriptografado,
            string ambiente, int? numeroLote = null, long? idLoteOriginal = null)
        {
            var dto = new RegistrarLoteDto
            {
                Tipo = tipo,
                Periodo = periodo,
                QuantidadeEventos = quantidadeEventos,
                CnpjDeclarante = cnpjDeclarante,
                CaminhoArquivoXml = caminhoArquivoXml,
                CaminhoArquivoAssinado = caminhoArquivoAssinado,
                CaminhoArquivoCriptografado = caminhoArquivoCriptografado,
                Ambiente = ambiente,
                NumeroLote = numeroLote,
                IdLoteOriginal = idLoteOriginal
            };
            return RegistrarLote(dto);
        }

        /// <summary>
        /// Registra um lote no banco de dados
        /// </summary>
        public long RegistrarLote(RegistrarLoteDto dto)
        {
            try
            {
                int semestre = CalcularSemestre(dto.Periodo);
                string status = "GERADO";
                string hashConteudo = CalcularHashArquivo(dto.CaminhoArquivoXml);

                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();

                    // Se não informou número do lote, buscar o próximo
                    int? numeroLote = dto.NumeroLote;
                    if (!numeroLote.HasValue)
                    {
                        numeroLote = ObterProximoNumeroLote(connection, dto.Periodo, semestre);
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
                        command.Parameters.AddWithValue("@periodo", dto.Periodo ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@semestre", semestre);
                        command.Parameters.AddWithValue("@numerolote", numeroLote.Value);
                        command.Parameters.AddWithValue("@quantidadeeventos", dto.QuantidadeEventos);
                        command.Parameters.AddWithValue("@cnpjdeclarante", dto.CnpjDeclarante ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@hashconteudo", hashConteudo ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@caminhoarquivolotexml", dto.CaminhoArquivoXml ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@caminhoarquivoloteassinadoxml", dto.CaminhoArquivoAssinado ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@caminhoarquivolotecriptografadoxml", dto.CaminhoArquivoCriptografado ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@status", status);
                        command.Parameters.AddWithValue("@ambiente", dto.Ambiente);
                        command.Parameters.AddWithValue("@datacriacao", DateTime.Now);
                        command.Parameters.AddWithValue("@idloteoriginal", dto.IdLoteOriginal ?? (object)DBNull.Value);

                        long idLote = (long)command.ExecuteScalar();
                        return idLote;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao registrar lote no banco: {ex.Message}");
                throw new EfinanceiraDatabasePersistenceException($"Erro ao registrar lote no banco de dados: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Atualiza o status e informações de um lote (sobrecarga para compatibilidade - apenas status)
        /// </summary>
        public void AtualizarLote(long idLote, string status)
        {
            var dto = new AtualizarLoteDto { Status = status };
            AtualizarLote(idLote, dto);
        }

        /// <summary>
        /// Atualiza o status e informações de um lote (sobrecarga para compatibilidade)
        /// </summary>
        public void AtualizarLote(long idLote, string status, string protocoloEnvio = null, 
            int? codigoRespostaEnvio = null, string descricaoRespostaEnvio = null, string xmlRespostaEnvio = null,
            int? codigoRespostaConsulta = null, string descricaoRespostaConsulta = null, string xmlRespostaConsulta = null,
            DateTime? dataEnvio = null, DateTime? dataConfirmacao = null, string ultimoErro = null)
        {
            var dto = new AtualizarLoteDto
            {
                Status = status,
                ProtocoloEnvio = protocoloEnvio,
                CodigoRespostaEnvio = codigoRespostaEnvio,
                DescricaoRespostaEnvio = descricaoRespostaEnvio,
                XmlRespostaEnvio = xmlRespostaEnvio,
                CodigoRespostaConsulta = codigoRespostaConsulta,
                DescricaoRespostaConsulta = descricaoRespostaConsulta,
                XmlRespostaConsulta = xmlRespostaConsulta,
                DataEnvio = dataEnvio,
                DataConfirmacao = dataConfirmacao,
                UltimoErro = ultimoErro
            };
            AtualizarLote(idLote, dto);
        }

        /// <summary>
        /// Atualiza o status e informações de um lote
        /// </summary>
        public void AtualizarLote(long idLote, AtualizarLoteDto dto)
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
                        new NpgsqlParameter("@status", dto.Status),
                        new NpgsqlParameter("@dataatualizacao", DateTime.Now)
                    };

                    if (!string.IsNullOrEmpty(dto.ProtocoloEnvio))
                    {
                        sql += ", protocoloenvio = @protocoloenvio";
                        parameters.Add(new NpgsqlParameter("@protocoloenvio", dto.ProtocoloEnvio));
                    }

                    if (dto.CodigoRespostaEnvio.HasValue)
                    {
                        sql += ", codigorespostaenvio = @codigorespostaenvio";
                        parameters.Add(new NpgsqlParameter("@codigorespostaenvio", dto.CodigoRespostaEnvio.Value));
                    }

                    if (!string.IsNullOrEmpty(dto.DescricaoRespostaEnvio))
                    {
                        sql += ", descricaorespostaenvio = @descricaorespostaenvio";
                        parameters.Add(new NpgsqlParameter("@descricaorespostaenvio", dto.DescricaoRespostaEnvio));
                    }

                    if (!string.IsNullOrEmpty(dto.XmlRespostaEnvio))
                    {
                        sql += ", xmlrespostaenvio = @xmlrespostaenvio";
                        parameters.Add(new NpgsqlParameter("@xmlrespostaenvio", dto.XmlRespostaEnvio));
                    }

                    if (dto.CodigoRespostaConsulta.HasValue)
                    {
                        sql += ", codigorespostaconsulta = @codigorespostaconsulta";
                        parameters.Add(new NpgsqlParameter("@codigorespostaconsulta", dto.CodigoRespostaConsulta.Value));
                    }

                    if (!string.IsNullOrEmpty(dto.DescricaoRespostaConsulta))
                    {
                        sql += ", descricaorespostaconsulta = @descricaorespostaconsulta";
                        parameters.Add(new NpgsqlParameter("@descricaorespostaconsulta", dto.DescricaoRespostaConsulta));
                    }

                    if (!string.IsNullOrEmpty(dto.XmlRespostaConsulta))
                    {
                        sql += ", xmlrespostaconsulta = @xmlrespostaconsulta";
                        parameters.Add(new NpgsqlParameter("@xmlrespostaconsulta", dto.XmlRespostaConsulta));
                    }

                    if (dto.DataEnvio.HasValue)
                    {
                        sql += ", dataenvio = @dataenvio";
                        parameters.Add(new NpgsqlParameter("@dataenvio", dto.DataEnvio.Value));
                    }

                    if (dto.DataConfirmacao.HasValue)
                    {
                        sql += ", dataconfirmacao = @dataconfirmacao";
                        parameters.Add(new NpgsqlParameter("@dataconfirmacao", dto.DataConfirmacao.Value));
                    }

                    if (!string.IsNullOrEmpty(dto.UltimoErro))
                    {
                        sql += ", ultimoerro = @ultimoerro";
                        parameters.Add(new NpgsqlParameter("@ultimoerro", dto.UltimoErro));
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
                throw new EfinanceiraDatabasePersistenceException($"Erro ao atualizar lote no banco de dados: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Normaliza o CPF removendo caracteres especiais e limitando a 11 caracteres
        /// </summary>
        private static string NormalizarCpf(string cpf)
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

                var eventoDto = new RegistrarEventoDto
                {
                    IdLote = idLote,
                    IdPessoa = pessoa.IdPessoa,
                    IdConta = pessoa.IdConta,
                    Cpf = cpfNormalizado,
                    Nome = pessoa.Nome,
                    NumeroConta = numeroConta,
                    DigitoConta = pessoa.DigitoConta,
                    SaldoAtual = pessoa.SaldoAtual,
                    TotCreditos = pessoa.TotCreditos,
                    TotDebitos = pessoa.TotDebitos,
                    IdEventoXml = idEventoXml,
                    StatusEvento = "GERADO",
                    OcorrenciasJson = null,
                    NumeroRecibo = null,
                    IndRetificacao = indRetificacao
                };

                RegistrarEvento(eventoDto);
            }
        }

        /// <summary>
        /// Registra um evento no banco de dados
        /// </summary>
        public long RegistrarEvento(RegistrarEventoDto dto)
        {
            try
            {
                // Normalizar CPF antes de salvar
                string cpfNormalizado = NormalizarCpf(dto.Cpf);

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
                        command.Parameters.AddWithValue(PARAM_ID_LOTE, dto.IdLote);
                        command.Parameters.AddWithValue("@idpessoa", dto.IdPessoa ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@idconta", dto.IdConta ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@cpf", cpfNormalizado ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@nome", dto.Nome ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@numeroconta", dto.NumeroConta ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@digitoconta", dto.DigitoConta ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@saldoatual", dto.SaldoAtual ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@totcreditos", dto.TotCreditos ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@totdebitos", dto.TotDebitos ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@ideventoxml", dto.IdEventoXml ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@statusevento", dto.StatusEvento ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@ocorrenciasjson", dto.OcorrenciasJson ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@datacriacao", DateTime.Now);
                        command.Parameters.AddWithValue("@numerorecibo", dto.NumeroRecibo ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@indretificacao", dto.IndRetificacao);

                        long idEvento = (long)command.ExecuteScalar();
                        return idEvento;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao registrar evento no banco: {ex.Message}");
                throw new EfinanceiraDatabasePersistenceException($"Erro ao registrar evento no banco de dados: {ex.Message}", ex);
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
                throw new EfinanceiraDatabasePersistenceException($"Erro ao atualizar evento no banco de dados: {ex.Message}", ex);
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
                        command.Parameters.AddWithValue(PARAM_ID_LOTE, idLote);
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
        /// Busca lotes do banco de dados com filtro de data e/ou período
        /// </summary>
        public List<LoteBancoInfo> BuscarLotes(DateTime? dataInicio = null, DateTime? dataFim = null, string periodo = null, int? limite = null)
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

                    var whereConditions = new List<string>();

                    // Adicionar filtro de data se informado
                    if (dataInicio.HasValue && dataFim.HasValue)
                    {
                        whereConditions.Add("l.datacriacao >= @datainicio");
                        whereConditions.Add("l.datacriacao <= @datafim");
                    }

                    // Adicionar filtro de período se informado
                    if (!string.IsNullOrWhiteSpace(periodo))
                    {
                        whereConditions.Add("l.periodo = @periodo");
                    }

                    // Adicionar WHERE se houver condições
                    if (whereConditions.Count > 0)
                    {
                        sql += "\n                        WHERE " + string.Join(" AND ", whereConditions);
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
                        if (!string.IsNullOrWhiteSpace(periodo))
                        {
                            command.Parameters.AddWithValue("@periodo", periodo);
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
                throw new EfinanceiraDatabasePersistenceException($"Erro ao buscar lotes do banco de dados: {ex.Message}", ex);
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
                        command.Parameters.AddWithValue(PARAM_ID_LOTE, idLote);

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
                throw new EfinanceiraDatabasePersistenceException($"Erro ao buscar eventos do lote: {ex.Message}", ex);
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
        private static int CalcularSemestre(string periodo)
        {
            if (string.IsNullOrEmpty(periodo) || periodo.Length < 6)
                return 0; // Retorna 0 para Cadastro de Declarante que não tem período

            int mes = int.Parse(periodo.Substring(4, 2));
            if (mes == 1 || mes == 6) return 1; // Primeiro semestre
            if (mes == 2 || mes == 12) return 2; // Segundo semestre
            return 0;
        }

        private int ObterProximoNumeroLote(NpgsqlConnection connection, string periodo, int semestre)
        {
            // Se período for null (Cadastro de Declarante), usar lógica diferente
            if (string.IsNullOrEmpty(periodo))
            {
                string sql = @"
                    SELECT COALESCE(MAX(numerolote), 0) + 1
                    FROM efinanceira.tb_efinanceira_lote
                    WHERE periodo IS NULL";

                using (var command = new NpgsqlCommand(sql, connection))
                {
                    object result = command.ExecuteScalar();
                    return result != null && result != DBNull.Value ? Convert.ToInt32(result) : 1;
                }
            }

            string sqlComPeriodo = @"
                SELECT COALESCE(MAX(numerolote), 0) + 1
                FROM efinanceira.tb_efinanceira_lote
                WHERE periodo = @periodo AND semestre = @semestre";

            using (var command = new NpgsqlCommand(sqlComPeriodo, connection))
            {
                command.Parameters.AddWithValue("@periodo", periodo);
                command.Parameters.AddWithValue("@semestre", semestre);
                object result = command.ExecuteScalar();
                return result != null && result != DBNull.Value ? Convert.ToInt32(result) : 1;
            }
        }

        private static string CalcularHashArquivo(string caminhoArquivo)
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

        private static TipoLote DeterminarTipoLote(LoteBancoInfo lote)
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
                else if (nomeArquivo.Contains("cadastro") || nomeArquivo.Contains("declarante"))
                {
                    return TipoLote.CadastroDeclarante;
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
        private static string GetSafeString(NpgsqlDataReader reader, string columnName, string defaultValue = "")
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

        private static long GetSafeLong(NpgsqlDataReader reader, string columnName, long defaultValue = 0)
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

        private static long? GetSafeLongNullable(NpgsqlDataReader reader, string columnName)
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

        private static int GetSafeInt(NpgsqlDataReader reader, string columnName, int defaultValue = 0)
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

        private static int? GetSafeIntNullable(NpgsqlDataReader reader, string columnName)
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

        private static decimal? GetSafeDecimalNullable(NpgsqlDataReader reader, string columnName)
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

        private static DateTime GetSafeDateTime(NpgsqlDataReader reader, string columnName)
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

        private static DateTime? GetSafeDateTimeNullable(NpgsqlDataReader reader, string columnName)
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
