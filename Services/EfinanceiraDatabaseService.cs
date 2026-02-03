using System;
using System.Collections.Generic;
using System.Data;
using Npgsql;
using NpgsqlTypes;
using ExemploAssinadorXML.Models;

namespace ExemploAssinadorXML.Services
{
    public class EfinanceiraDatabaseService
    {
        private readonly string _connectionString;

        public EfinanceiraDatabaseService()
        {
            // Credenciais do banco de dados fornecidas pelo usuário
            const string DB_HOST = "10.30.0.21";
            const string DB_PORT = "5432";
            const string DB_NAME = "bscash";
            const string DB_USER = "nickolas.oliveira";
            const string DB_PASSWORD = "1QclT+-IVB2B";

            // Timeout de conexão otimizado: 15 segundos (suficiente para conexão)
            // Command Timeout será definido por comando individualmente
            _connectionString = $"Host={DB_HOST};Port={DB_PORT};Database={DB_NAME};Username={DB_USER};Password={DB_PASSWORD};Timeout=15;Pooling=true;MinPoolSize=1;MaxPoolSize=20;";
        }

        public EfinanceiraDatabaseService(string connectionString)
        {
            _connectionString = connectionString;
        }

        /// <summary>
        /// Testa a conexão com o banco de dados
        /// </summary>
        public bool TestarConexao()
        {
            try
            {
                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Erro ao conectar ao banco de dados: {ex.Message}", 
                    "Erro de Conexão", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Busca pessoas com contas e movimentações no período informado
        /// </summary>
        public List<DadosPessoaConta> BuscarPessoasComContas(int ano, int mesInicial, int mesFinal, int limit, int offset)
        {
            List<DadosPessoaConta> pessoas = new List<DadosPessoaConta>();
            
            // Validação dos parâmetros recebidos
            if (mesInicial < 1 || mesInicial > 12)
            {
                throw new ArgumentException($"Mês Inicial inválido: {mesInicial}. Deve estar entre 1 e 12.");
            }
            if (mesFinal < 1 || mesFinal > 12)
            {
                throw new ArgumentException($"Mês Final inválido: {mesFinal}. Deve estar entre 1 e 12.");
            }
            if (mesInicial > mesFinal)
            {
                throw new ArgumentException($"Mês Inicial ({mesInicial}) não pode ser maior que Mês Final ({mesFinal}). " +
                    $"Para períodos semestrais, use: 1-6 (Jan-Jun) ou 7-12 (Jul-Dez).");
            }
            
            // Log dos filtros aplicados
            System.Diagnostics.Debug.WriteLine("=== FILTROS APLICADOS NA BUSCA ===");
            System.Diagnostics.Debug.WriteLine($"Ano: {ano}");
            System.Diagnostics.Debug.WriteLine($"Mês Inicial: {mesInicial} ({GetNomeMes(mesInicial)})");
            System.Diagnostics.Debug.WriteLine($"Mês Final: {mesFinal} ({GetNomeMes(mesFinal)})");
            System.Diagnostics.Debug.WriteLine($"Limit (máximo de registros): {limit}");
            System.Diagnostics.Debug.WriteLine($"Offset (registros a pular): {offset}");
            System.Diagnostics.Debug.WriteLine("");
            System.Diagnostics.Debug.WriteLine("=== CONDIÇÕES DE FILTRO ===");
            System.Diagnostics.Debug.WriteLine("✓ Pessoa ativa (p.situacao = '1')");
            System.Diagnostics.Debug.WriteLine("✓ Conta ativa (c.situacao = '1')");
            System.Diagnostics.Debug.WriteLine("✓ CPF não nulo e não vazio");
            System.Diagnostics.Debug.WriteLine("✓ CPF com pelo menos 11 caracteres");
            System.Diagnostics.Debug.WriteLine($"✓ Ano da operação = {ano}");
            System.Diagnostics.Debug.WriteLine($"✓ Mês da operação entre {mesInicial} ({GetNomeMes(mesInicial)}) e {mesFinal} ({GetNomeMes(mesFinal)})");
            System.Diagnostics.Debug.WriteLine($"✓ Limite de {limit} registros");
            System.Diagnostics.Debug.WriteLine($"✓ Pulando {offset} registros iniciais");
            System.Diagnostics.Debug.WriteLine("");
            
            string sql = @"
                WITH contas_com_movimentacao AS (
                    -- Primeiro: identificar contas que têm movimentação no período (mais eficiente)
                    SELECT DISTINCT ex.idconta
                    FROM conta.tb_extrato ex
                    WHERE ex.dataoperacao >= @dataInicio
                      AND ex.dataoperacao <= @dataFim
                ),
                totais_extrato AS (
                    -- Segundo: calcular totais apenas para contas que têm movimentação
                    SELECT 
                        ex.idconta,
                        COALESCE(SUM(CASE WHEN ex.naturezaoperacao = 'C' THEN ex.valoroperacao ELSE 0 END), 0) as TotCreditos,
                        COALESCE(SUM(CASE WHEN ex.naturezaoperacao = 'D' THEN ex.valoroperacao ELSE 0 END), 0) as TotDebitos
                    FROM conta.tb_extrato ex
                    WHERE ex.dataoperacao >= @dataInicio
                      AND ex.dataoperacao <= @dataFim
                    GROUP BY ex.idconta
                )
                SELECT 
                    p.idpessoa as IdPessoa,
                    p.documento as Documento,
                    p.nome as Nome,
                    pf.cpf as Cpf,
                    COALESCE(pf.nacionalidade, 'BR') as Nacionalidade,
                    COALESCE(p.telefone, '') as Telefone,
                    COALESCE(p.email, '') as Email,
                    c.idconta as IdConta,
                    COALESCE(CAST(c.numeroconta AS TEXT), '') as NumeroConta,
                    COALESCE(c.digitoconta, '') as DigitoConta,
                    COALESCE(c.saldoatual, 0) as SaldoAtual,
                    COALESCE(e.logradouro, '') as Logradouro,
                    COALESCE(e.numero, '') as Numero,
                    COALESCE(e.complemento, '') as Complemento,
                    COALESCE(e.bairro, '') as Bairro,
                    COALESCE(e.cep, '') as Cep,
                    COALESCE(e.tipologradouro, '') as TipoLogradouro,
                    '' as EnderecoLivre,
                    COALESCE(te.TotCreditos, 0) as TotCreditos,
                    COALESCE(te.TotDebitos, 0) as TotDebitos
                FROM manager.tb_pessoa p
                INNER JOIN manager.tb_pessoafisica pf ON pf.idpessoa = p.idpessoa
                INNER JOIN conta.tb_conta c ON c.idpessoa = p.idpessoa
                INNER JOIN contas_com_movimentacao ccm ON ccm.idconta = c.idconta
                LEFT JOIN totais_extrato te ON te.idconta = c.idconta
                LEFT JOIN manager.tb_endereco e ON e.idpessoa = p.idpessoa AND e.situacao = '1'
                WHERE p.situacao = '1'
                  AND c.situacao = '1'
                  AND pf.cpf IS NOT NULL
                  AND TRIM(pf.cpf) != ''
                  AND LENGTH(TRIM(pf.cpf)) >= 11
                ORDER BY p.idpessoa
                LIMIT @limit OFFSET @offset";

            try
            {
                System.Diagnostics.Debug.WriteLine("=== EXECUTANDO CONSULTA ===");
                System.Diagnostics.Debug.WriteLine($"Conectando ao banco: {_connectionString.Split(';')[0]}...");
                
                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();
                    System.Diagnostics.Debug.WriteLine("✓ Conexão estabelecida com sucesso");

                    using (var command = new NpgsqlCommand(sql, connection))
                    {
                        // Timeout otimizado: reduzido para 120 segundos (2 minutos) já que a consulta está otimizada
                        command.CommandTimeout = 120;
                        
                        // Calcular datas para filtro eficiente (permite uso de índices)
                        DateTime dataInicio = new DateTime(ano, mesInicial, 1);
                        DateTime dataFim = new DateTime(ano, mesFinal, DateTime.DaysInMonth(ano, mesFinal), 23, 59, 59);
                        
                        // Especificar tipos explicitamente para evitar problemas de inferência
                        command.Parameters.Add(new NpgsqlParameter("@dataInicio", NpgsqlDbType.Timestamp) { Value = dataInicio });
                        command.Parameters.Add(new NpgsqlParameter("@dataFim", NpgsqlDbType.Timestamp) { Value = dataFim });
                        command.Parameters.Add(new NpgsqlParameter("@limit", NpgsqlDbType.Integer) { Value = limit });
                        command.Parameters.Add(new NpgsqlParameter("@offset", NpgsqlDbType.Integer) { Value = offset });

                        System.Diagnostics.Debug.WriteLine("Executando query SQL...");
                        using (var reader = command.ExecuteReader())
                        {
                            int contador = 0;
                            try
                            {
                                while (reader.Read())
                                {
                                    try
                                    {
                                        var pessoa = new DadosPessoaConta
                                        {
                                            IdPessoa = GetSafeLong(reader, "IdPessoa"),
                                            Documento = GetSafeString(reader, "Documento"),
                                            Nome = GetSafeString(reader, "Nome"),
                                            Cpf = GetSafeString(reader, "Cpf"),
                                            Nacionalidade = GetSafeString(reader, "Nacionalidade", "BR"),
                                            Telefone = GetSafeString(reader, "Telefone"),
                                            Email = GetSafeString(reader, "Email"),
                                            IdConta = GetSafeLong(reader, "IdConta"),
                                            NumeroConta = GetSafeNumeroConta(reader, "NumeroConta"),
                                            DigitoConta = GetSafeString(reader, "DigitoConta"),
                                            SaldoAtual = GetSafeDecimal(reader, "SaldoAtual"),
                                            Logradouro = GetSafeString(reader, "Logradouro"),
                                            Numero = GetSafeString(reader, "Numero"),
                                            Complemento = GetSafeString(reader, "Complemento"),
                                            Bairro = GetSafeString(reader, "Bairro"),
                                            Cep = GetSafeString(reader, "Cep"),
                                            TipoLogradouro = GetSafeString(reader, "TipoLogradouro"),
                                            EnderecoLivre = GetSafeString(reader, "EnderecoLivre"),
                                            TotCreditos = GetSafeDecimal(reader, "TotCreditos"),
                                            TotDebitos = GetSafeDecimal(reader, "TotDebitos")
                                        };

                                        pessoas.Add(pessoa);
                                        contador++;
                                    }
                                    catch (Exception exLinha)
                                    {
                                        // Log do erro mas continua processando outras linhas
                                        System.Diagnostics.Debug.WriteLine($"✗ Erro ao processar linha {contador + 1}: {exLinha.Message}");
                                        System.Diagnostics.Debug.WriteLine($"   Stack Trace: {exLinha.StackTrace}");
                                        if (exLinha.InnerException != null)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"   Inner Exception: {exLinha.InnerException.Message}");
                                        }
                                    }
                                }
                                
                                System.Diagnostics.Debug.WriteLine($"✓ Consulta concluída: {contador} registro(s) encontrado(s) e processado(s) com sucesso");
                                System.Diagnostics.Debug.WriteLine($"✓ Total retornado: {pessoas.Count} pessoa(s)");
                                System.Diagnostics.Debug.WriteLine("");
                            }
                            catch (Exception exRead)
                            {
                                System.Diagnostics.Debug.WriteLine($"✗ ERRO ao ler dados do DataReader: {exRead.Message}");
                                System.Diagnostics.Debug.WriteLine($"   Stack Trace: {exRead.StackTrace}");
                                if (exRead.InnerException != null)
                                {
                                    System.Diagnostics.Debug.WriteLine($"   Inner Exception: {exRead.InnerException.Message}");
                                }
                                throw;
                            }
                        }
                    }
                }
            }
            catch (NpgsqlException npgEx)
            {
                string filtrosInfo = $"Filtros aplicados:\n" +
                    $"  • Ano: {ano}\n" +
                    $"  • Mês Inicial: {mesInicial}\n" +
                    $"  • Mês Final: {mesFinal}\n" +
                    $"  • Limit: {limit}\n" +
                    $"  • Offset: {offset}";
                
                string mensagemDetalhada = $"Erro PostgreSQL ao buscar pessoas com contas:\n\n{npgEx.Message}\n\n{filtrosInfo}";
                if (npgEx.InnerException != null)
                {
                    mensagemDetalhada += $"\n\nDetalhes: {npgEx.InnerException.Message}";
                }
                
                System.Diagnostics.Debug.WriteLine($"✗ ERRO PostgreSQL: {npgEx.Message}");
                System.Diagnostics.Debug.WriteLine(filtrosInfo);
                
                System.Windows.Forms.MessageBox.Show(mensagemDetalhada, 
                    "Erro de Banco de Dados", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                string filtrosInfo = $"Filtros aplicados:\n" +
                    $"  • Ano: {ano}\n" +
                    $"  • Mês Inicial: {mesInicial}\n" +
                    $"  • Mês Final: {mesFinal}\n" +
                    $"  • Limit: {limit}\n" +
                    $"  • Offset: {offset}";
                
                string mensagemDetalhada = $"Erro ao buscar pessoas com contas:\n\n{ex.Message}\n\n{filtrosInfo}";
                if (ex.InnerException != null)
                {
                    mensagemDetalhada += $"\n\nDetalhes: {ex.InnerException.Message}";
                }
                mensagemDetalhada += $"\n\nStack Trace: {ex.StackTrace}";
                
                System.Diagnostics.Debug.WriteLine($"✗ ERRO: {ex.Message}");
                System.Diagnostics.Debug.WriteLine(filtrosInfo);
                System.Diagnostics.Debug.WriteLine($"Stack Trace: {ex.StackTrace}");
                
                System.Windows.Forms.MessageBox.Show(mensagemDetalhada, 
                    "Erro de Banco de Dados", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            return pessoas;
        }

        /// <summary>
        /// Calcula totais de movimentação para uma conta específica
        /// </summary>
        public Tuple<decimal, decimal> CalcularTotaisMovimentacao(long idConta, int ano, int mesInicial, int mesFinal)
        {
            decimal totCreditos = 0;
            decimal totDebitos = 0;

            System.Diagnostics.Debug.WriteLine("=== CALCULANDO TOTAIS DE MOVIMENTAÇÃO ===");
            System.Diagnostics.Debug.WriteLine($"ID Conta: {idConta}");
            System.Diagnostics.Debug.WriteLine($"Ano: {ano}");
            System.Diagnostics.Debug.WriteLine($"Mês Inicial: {mesInicial}");
            System.Diagnostics.Debug.WriteLine($"Mês Final: {mesFinal}");
            System.Diagnostics.Debug.WriteLine("");

            // Otimização: usar comparação direta de datas ao invés de EXTRACT para permitir uso de índices
            DateTime dataInicio = new DateTime(ano, mesInicial, 1);
            DateTime dataFim = new DateTime(ano, mesFinal, DateTime.DaysInMonth(ano, mesFinal), 23, 59, 59);
            
            string sql = @"
                SELECT 
                    COALESCE(SUM(CASE WHEN e.naturezaoperacao = 'C' THEN e.valoroperacao ELSE 0 END), 0) as TotCreditos,
                    COALESCE(SUM(CASE WHEN e.naturezaoperacao = 'D' THEN e.valoroperacao ELSE 0 END), 0) as TotDebitos
                FROM conta.tb_extrato e
                WHERE e.idconta = @idConta
                  AND e.dataoperacao >= @dataInicio
                  AND e.dataoperacao <= @dataFim";

            try
            {
                using (var connection = new NpgsqlConnection(_connectionString))
                {
                    connection.Open();

                    using (var command = new NpgsqlCommand(sql, connection))
                    {
                        // Timeout otimizado
                        command.CommandTimeout = 60;
                        
                        command.Parameters.Add(new NpgsqlParameter("@idConta", NpgsqlDbType.Bigint) { Value = idConta });
                        command.Parameters.Add(new NpgsqlParameter("@dataInicio", NpgsqlDbType.Timestamp) { Value = dataInicio });
                        command.Parameters.Add(new NpgsqlParameter("@dataFim", NpgsqlDbType.Timestamp) { Value = dataFim });

                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                totCreditos = GetSafeDecimal(reader, "TotCreditos");
                                totDebitos = GetSafeDecimal(reader, "TotDebitos");
                                System.Diagnostics.Debug.WriteLine($"✓ Totais calculados - Créditos: R$ {totCreditos:F2}, Débitos: R$ {totDebitos:F2}");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine("⚠ Nenhum registro encontrado para calcular totais");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Erro ao calcular totais de movimentação: {ex.Message}", 
                    "Erro de Banco de Dados", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            return Tuple.Create(totCreditos, totDebitos);
        }

        // Métodos auxiliares para leitura segura dos dados
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

        private static decimal GetSafeDecimal(NpgsqlDataReader reader, string columnName, decimal defaultValue = 0)
        {
            try
            {
                int ordinal = reader.GetOrdinal(columnName);
                if (reader.IsDBNull(ordinal))
                    return defaultValue;
                
                // Tenta ler como decimal primeiro
                if (reader.GetFieldType(ordinal) == typeof(decimal) || reader.GetFieldType(ordinal) == typeof(double) || reader.GetFieldType(ordinal) == typeof(float))
                {
                    return reader.GetDecimal(ordinal);
                }
                
                // Se não for decimal, tenta converter
                object value = reader.GetValue(ordinal);
                if (value != null && value != DBNull.Value)
                {
                    return Convert.ToDecimal(value);
                }
                
                return defaultValue;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao ler decimal da coluna {columnName}: {ex.Message}");
                return defaultValue;
            }
        }

        private static string GetSafeNumeroConta(NpgsqlDataReader reader, string columnName, string defaultValue = "")
        {
            try
            {
                int ordinal = reader.GetOrdinal(columnName);
                if (reader.IsDBNull(ordinal))
                    return defaultValue;

                // Verifica o tipo do campo
                Type fieldType = reader.GetFieldType(ordinal);
                
                // Se for string, lê diretamente
                if (fieldType == typeof(string))
                {
                    return reader.GetString(ordinal);
                }
                
                // Se for numérico (int, long, decimal, etc), converte para string
                if (fieldType == typeof(long) || fieldType == typeof(int) || fieldType == typeof(short) || 
                    fieldType == typeof(decimal) || fieldType == typeof(double) || fieldType == typeof(float))
                {
                    object value = reader.GetValue(ordinal);
                    if (value != null && value != DBNull.Value)
                    {
                        // Remove decimais se houver (ex: 12345.00 -> 12345)
                        if (value is decimal dec)
                        {
                            return ((long)dec).ToString();
                        }
                        return value.ToString();
                    }
                }
                
                // Para outros tipos, tenta converter para string
                object objValue = reader.GetValue(ordinal);
                return objValue?.ToString() ?? defaultValue;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao ler NumeroConta da coluna {columnName}: {ex.Message}");
                return defaultValue;
            }
        }

        private static string GetNomeMes(int mes)
        {
            string[] meses = { "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro" };
            return mes >= 1 && mes <= 12 ? meses[mes] : $"Mês {mes}";
        }
    }
}