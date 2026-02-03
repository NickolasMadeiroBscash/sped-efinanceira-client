# Otimizações de Performance - Banco de Dados

## Melhorias Implementadas

### 1. Consulta Principal (`BuscarPessoasComContas`)

**Problemas Identificados:**
- JOIN com `tb_extrato` antes de filtrar por data, gerando milhões de linhas desnecessárias
- Uso de `EXTRACT(YEAR/MONTH)` na cláusula WHERE, impedindo uso de índices
- GROUP BY com muitos campos, causando lentidão
- Timeout muito alto (5 minutos), indicando consulta ineficiente

**Soluções Implementadas:**
- ✅ Uso de CTE (Common Table Expression) para filtrar extratos primeiro
- ✅ Comparação direta de datas (`dataoperacao >= @dataInicio AND dataoperacao <= @dataFim`) ao invés de EXTRACT
- ✅ Separação da lógica em duas CTEs: uma para identificar contas com movimentação e outra para calcular totais
- ✅ Timeout reduzido para 120 segundos (2 minutos) - consulta agora é mais rápida
- ✅ Connection pooling habilitado para reutilização de conexões

### 2. Consulta de Totais (`CalcularTotaisMovimentacao`)

**Problemas Identificados:**
- Uso de `EXTRACT(YEAR/MONTH)` na cláusula WHERE

**Soluções Implementadas:**
- ✅ Comparação direta de datas ao invés de EXTRACT
- ✅ Timeout otimizado para 60 segundos

### 3. Configuração de Conexão

**Melhorias:**
- ✅ Connection pooling habilitado (`Pooling=true;MinPoolSize=1;MaxPoolSize=20`)
- ✅ Timeout de conexão reduzido para 15 segundos (suficiente para conexão)
- ✅ Command timeout definido individualmente por comando

## Índices Recomendados para PostgreSQL

Para melhorar ainda mais a performance, execute os seguintes comandos SQL no banco de dados:

```sql
-- Índice para filtro de data na tabela de extrato (CRÍTICO para performance)
CREATE INDEX IF NOT EXISTS idx_extrato_dataoperacao 
ON conta.tb_extrato(dataoperacao);

-- Índice composto para consultas com idconta e data
CREATE INDEX IF NOT EXISTS idx_extrato_idconta_data 
ON conta.tb_extrato(idconta, dataoperacao);

-- Índice para natureza da operação (usado nos cálculos de totais)
CREATE INDEX IF NOT EXISTS idx_extrato_idconta_natureza_data 
ON conta.tb_extrato(idconta, naturezaoperacao, dataoperacao);

-- Índice para filtro de situação na tabela de pessoa
CREATE INDEX IF NOT EXISTS idx_pessoa_situacao 
ON manager.tb_pessoa(situacao) 
WHERE situacao = '1';

-- Índice para filtro de situação na tabela de conta
CREATE INDEX IF NOT EXISTS idx_conta_situacao 
ON conta.tb_conta(situacao) 
WHERE situacao = '1';

-- Índice para CPF (usado no filtro)
CREATE INDEX IF NOT EXISTS idx_pessoafisica_cpf 
ON manager.tb_pessoafisica(cpf) 
WHERE cpf IS NOT NULL AND TRIM(cpf) != '';

-- Índice para endereço ativo
CREATE INDEX IF NOT EXISTS idx_endereco_pessoa_situacao 
ON manager.tb_endereco(idpessoa, situacao) 
WHERE situacao = '1';
```

### Verificar Índices Existentes

Para verificar se os índices já existem:

```sql
-- Listar índices da tabela tb_extrato
SELECT indexname, indexdef 
FROM pg_indexes 
WHERE tablename = 'tb_extrato' AND schemaname = 'conta';

-- Listar índices da tabela tb_pessoa
SELECT indexname, indexdef 
FROM pg_indexes 
WHERE tablename = 'tb_pessoa' AND schemaname = 'manager';

-- Listar índices da tabela tb_conta
SELECT indexname, indexdef 
FROM pg_indexes 
WHERE tablename = 'tb_conta' AND schemaname = 'conta';
```

## Ganhos Esperados de Performance

- **Antes:** Consultas levavam 5+ minutos ou falhavam por timeout
- **Depois:** Consultas devem executar em menos de 30 segundos (com índices)
- **Redução estimada:** 90-95% no tempo de execução

## Monitoramento

Para monitorar a performance das consultas, execute:

```sql
-- Ver consultas lentas em execução
SELECT pid, now() - pg_stat_activity.query_start AS duration, query 
FROM pg_stat_activity 
WHERE (now() - pg_stat_activity.query_start) > interval '5 seconds'
AND state = 'active';

-- Ver estatísticas de uso de índices
SELECT schemaname, tablename, indexname, idx_scan, idx_tup_read, idx_tup_fetch
FROM pg_stat_user_indexes
WHERE schemaname IN ('conta', 'manager')
ORDER BY idx_scan DESC;
```

## Notas Importantes

1. **Índices Parciais:** Os índices com `WHERE` são índices parciais, que ocupam menos espaço e são mais eficientes para filtros específicos.

2. **Manutenção:** Após criar os índices, execute `ANALYZE` nas tabelas para atualizar as estatísticas:
   ```sql
   ANALYZE conta.tb_extrato;
   ANALYZE manager.tb_pessoa;
   ANALYZE conta.tb_conta;
   ANALYZE manager.tb_pessoafisica;
   ANALYZE manager.tb_endereco;
   ```

3. **Espaço em Disco:** Os índices ocupam espaço adicional. Monitore o uso de disco após criar os índices.

4. **Backup:** Sempre faça backup antes de criar índices em produção.
