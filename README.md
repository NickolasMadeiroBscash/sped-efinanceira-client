# Sistema e-Financeira - Assinatura e Envio

## 📋 Índice

- [Visão Geral](#visão-geral)
- [Para que Serve](#para-que-serve)
- [Funcionalidades](#funcionalidades)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Tecnologias Utilizadas](#tecnologias-utilizadas)
- [Requisitos](#requisitos)
- [Como Usar](#como-usar)
- [Arquitetura e Funcionamento](#arquitetura-e-funcionamento)
- [Melhorias Implementadas](#melhorias-implementadas)
- [Como Evoluir](#como-evoluir)
- [Troubleshooting](#troubleshooting)

---

## 🎯 Visão Geral

O **Sistema e-Financeira - Assinatura e Envio** é uma aplicação desktop desenvolvida em C# (.NET Framework 4.8) com Windows Forms, projetada para automatizar o processo de geração, assinatura digital, criptografia e envio de eventos para a plataforma e-Financeira da Receita Federal do Brasil.

O sistema integra-se com um banco de dados PostgreSQL para extrair dados de movimentações financeiras e gerar os XMLs conforme as especificações técnicas da e-Financeira.

---

## 🎯 Para que Serve

Este sistema foi desenvolvido para:

1. **Automatizar a Declaração e-Financeira**: Elimina a necessidade de processamento manual de grandes volumes de dados de movimentações financeiras.

2. **Gerar Eventos XML Conformes**: Cria arquivos XML que seguem rigorosamente os schemas XSD da e-Financeira, incluindo:
   - Eventos de Abertura (`evtAberturaeFinanceira`)
   - Eventos de Movimentação (`evtMovOpFin`)
   - Eventos de Fechamento (`evtFechamentoeFinanceira`)

3. **Assinar Digitalmente**: Aplica assinatura digital XML-DSig usando certificados A1 (token ou smartcard) com algoritmo RSA-SHA256.

4. **Criptografar Lotes**: Criptografa os lotes usando criptografia híbrida (AES-128-CBC + RSA) conforme especificação da Receita Federal.

5. **Enviar para e-Financeira**: Realiza o envio automático dos lotes criptografados para os endpoints da Receita Federal (ambiente de teste ou produção).

6. **Consultar Status**: Permite consultar o status de processamento dos lotes enviados através dos protocolos retornados.

7. **Integração com Banco de Dados**: Conecta-se a um banco PostgreSQL para extrair dados de pessoas, contas e movimentações financeiras.

---

## ✨ Funcionalidades

### 1. **Configuração**
- Configuração de certificados digitais (assinatura e servidor)
- Configuração de ambiente (Teste/Produção)
- Configuração de dados de abertura e fechamento
- Configuração de parâmetros de processamento (paginamento, eventos por lote, etc.)
- Persistência de configurações em arquivo XML

### 2. **Processamento de Abertura**
- Geração de XML de evento de abertura
- Assinatura digital do evento
- Criptografia do lote
- Envio opcional para e-Financeira
- Registro automático de protocolos

### 3. **Processamento de Movimentação**
- Consulta paginada ao banco de dados PostgreSQL
- Geração de lotes com múltiplos eventos (até 50 eventos por lote)
- Processamento em lote com controle de progresso
- Suporte a períodos semestrais (Jan-Jun ou Jul-Dez)
- Configuração flexível de eventos por lote e paginação

### 4. **Processamento de Fechamento**
- Geração de XML de evento de fechamento
- Suporte a diferentes tipos de fechamento (PP, MovOpFin, MovOpFinAnual)
- Assinatura e criptografia
- Envio opcional para e-Financeira

### 5. **Consulta de Protocolos**
- Consulta de status de lotes enviados
- Visualização de ocorrências e erros
- Lista de lotes processados
- Geração de fechamento por período

### 6. **Tutorial Integrado**
- Documentação e guia de uso dentro da aplicação

---

## 📁 Estrutura do Projeto

```
assinadorEfinanceira/
├── Forms/                          # Formulários Windows Forms
│   ├── ConfiguracaoForm.cs         # Tela de configuração
│   ├── ProcessamentoForm.cs        # Tela de processamento
│   ├── ConsultaForm.cs             # Tela de consulta de protocolos
│   └── TutorialForm.cs             # Tela de tutorial
│
├── Services/                        # Serviços de negócio
│   ├── EfinanceiraAssinaturaService.cs      # Assinatura digital XML
│   ├── EfinanceiraCriptografiaService.cs     # Criptografia de lotes
│   ├── EfinanceiraEnvioService.cs           # Envio para e-Financeira
│   ├── EfinanceiraConsultaService.cs        # Consulta de protocolos
│   ├── EfinanceiraGeradorXmlService.cs      # Geração de XMLs
│   ├── EfinanceiraDatabaseService.cs        # Acesso ao banco PostgreSQL
│   ├── EfinanceiraPeriodoUtil.cs            # Utilitários de período
│   ├── ConfiguracaoPersistenciaService.cs    # Persistência de configurações
│   ├── ProtocoloPersistenciaService.cs       # Persistência de protocolos
│   └── RSAPKCS1SHA256SignatureDescription.cs # Algoritmo de assinatura
│
├── Models/                          # Modelos de dados
│   ├── EfinanceiraConfig.cs        # Configuração geral
│   ├── DadosAbertura.cs            # Dados de evento de abertura
│   ├── DadosFechamento.cs          # Dados de evento de fechamento
│   ├── DadosPessoaConta.cs        # Dados de pessoa e conta
│   └── StatusProcessamento.cs      # Status do processamento
│
├── MainForm.cs                      # Formulário principal (com abas)
├── Program.cs                       # Ponto de entrada da aplicação
├── ExemploAssinadorXML.csproj       # Arquivo de projeto
└── App.config                       # Configuração da aplicação
```

### Descrição dos Componentes

#### **Forms/**
- **ConfiguracaoForm**: Interface completa para configuração de todos os parâmetros do sistema, incluindo certificados, dados de abertura/fechamento e parâmetros de processamento.
- **ProcessamentoForm**: Interface de processamento com controle de progresso, logs e estatísticas em tempo real.
- **ConsultaForm**: Interface para consulta de protocolos e visualização de lotes processados.
- **TutorialForm**: Documentação integrada.

#### **Services/**
- **EfinanceiraAssinaturaService**: Implementa assinatura digital XML-DSig com algoritmo RSA-SHA256, suportando múltiplos tipos de eventos e estruturas de lote.
- **EfinanceiraCriptografiaService**: Implementa criptografia híbrida (AES-128-CBC para o XML + RSA para a chave AES).
- **EfinanceiraEnvioService**: Gerencia comunicação HTTP com os endpoints da e-Financeira, incluindo tratamento de respostas e ocorrências.
- **EfinanceiraConsultaService**: Realiza consultas de status de protocolos enviados.
- **EfinanceiraGeradorXmlService**: Gera XMLs conformes aos schemas XSD da e-Financeira para todos os tipos de eventos.
- **EfinanceiraDatabaseService**: Acesso ao banco PostgreSQL com consultas otimizadas e paginação.
- **EfinanceiraPeriodoUtil**: Utilitários para cálculo e validação de períodos semestrais.

#### **Models/**
- Modelos de dados serializáveis para configuração, eventos e status.

---

## 🛠 Tecnologias Utilizadas

### Framework e Linguagem
- **.NET Framework 4.8**
- **C#**
- **Windows Forms**

### Bibliotecas e Pacotes NuGet
- **Npgsql 6.0.11**: Driver PostgreSQL para .NET
- **System.Security.Cryptography.Xml 6.0.1**: Assinatura digital XML
- **System.Text.Json 10.0.2**: Serialização JSON (para configurações)
- **System.Security.Cryptography**: Criptografia AES e RSA

### Banco de Dados
- **PostgreSQL**: Banco de dados relacional para armazenar dados de pessoas, contas e movimentações

### Certificados Digitais
- **Certificados A1**: Token ou smartcard instalados no repositório do Windows
- **Algoritmo de Assinatura**: RSA-SHA256 (XML-DSig)

---

## 📋 Requisitos

### Requisitos de Sistema
- Windows 7 ou superior
- .NET Framework 4.8
- Certificado digital A1 instalado no Windows (para assinatura)
- Certificado do servidor e-Financeira instalado (para criptografia)
- Acesso à internet (para envio e consulta)
- Acesso ao banco de dados PostgreSQL

### Requisitos de Certificados
1. **Certificado para Assinatura**: Certificado A1 com chave privada, instalado no repositório `CurrentUser\My` do Windows, com permissão de assinatura digital.
2. **Certificado do Servidor**: Certificado público da Receita Federal para criptografia, instalado no repositório do Windows.

### Requisitos de Banco de Dados
- PostgreSQL com acesso às tabelas:
  - `manager.tb_pessoa`
  - `manager.tb_pessoafisica`
  - `conta.tb_conta`
  - `conta.tb_extrato`
  - `manager.tb_endereco`

---

## 🚀 Como Usar

### 1. Configuração Inicial

#### Passo 1: Configurar Certificados
1. Abra a aplicação
2. Vá para a aba **"Configuração"**
3. Clique em **"Selecionar..."** ao lado de "Certificado para Assinatura"
4. Selecione seu certificado digital A1
5. Clique em **"Selecionar..."** ao lado de "Certificado do Servidor"
6. Selecione o certificado público da Receita Federal

#### Passo 2: Configurar Dados Gerais
1. Preencha o **CNPJ Declarante**
2. Configure o **Período** no formato `YYYYMM`:
   - `01` ou `06` = Primeiro semestre (Janeiro a Junho)
   - `02` ou `12` = Segundo semestre (Julho a Dezembro)
   - Exemplo: `202301` = Jan-Jun/2023
3. Selecione o **Diretório de Lotes** onde os XMLs serão salvos
4. Escolha o **Ambiente** (TEST ou PROD)

#### Passo 3: Configurar Dados de Abertura
1. Na aba **"Abertura"** dentro de "Configuração"
2. Preencha:
   - Data Início e Data Fim (formato: `AAAA-MM-DD`)
   - Tipo Ambiente (1 = Produção, 2 = Homologação)
   - Aplicação Emissora
   - Indicação de Retificação
3. Se marcar **"Indicar MovOpFin"**, preencha:
   - Responsável RMF (CNPJ, CPF, Nome, Setor, Telefone, Endereço completo)
   - Responsável e-Financeira (CPF, Nome, Setor, Telefone, Endereço, Email)
   - Representante Legal (CPF, Setor, Telefone)

#### Passo 4: Configurar Dados de Fechamento
1. Na aba **"Fechamento"** dentro de "Configuração"
2. Preencha:
   - Data Início e Data Fim
   - Tipo Ambiente
   - Situação Especial
   - Se não marcar "Nada a Declarar", preencha pelo menos um:
     - FechamentoPP (0 = sem movimento, 1 = com movimento)
     - FechamentoMovOpFin (0 = sem movimento, 1 = com movimento)
     - FechamentoMovOpFinAnual (0 = sem movimento, 1 = com movimento)

#### Passo 5: Configurar Parâmetros de Processamento
1. Na seção **"Configurações de Processamento"**:
   - **Page Size**: Tamanho da página para consultas ao banco (Produção: 500+, Teste: 50-100)
   - **Evento Offset**: Onde começar a gerar eventos (normalmente 0 ou 1)
   - **Offset Registros**: Pular registros iniciais (usar apenas em teste)
   - **Max Lotes**: Limitar quantidade de lotes (ou "Ilimitado")
   - **Eventos por Lote**: Quantidade de eventos por lote (1 a 50, conforme manual e-Financeira)

#### Passo 6: Salvar Configuração
1. Clique em **"Salvar Configurações"**
2. As configurações serão salvas em arquivo XML e carregadas automaticamente na próxima execução

### 2. Processar Abertura

1. Vá para a aba **"Processamento"**
2. Clique em **"Processar Abertura"**
3. O sistema irá:
   - Gerar o XML de abertura
   - Assinar digitalmente
   - Criptografar
   - Enviar para e-Financeira (se não estiver marcado "Apenas Processar")
4. O protocolo retornado será exibido e salvo automaticamente

### 3. Processar Movimentação

1. Certifique-se de que o **Período** está configurado corretamente
2. Clique em **"Processar Movimentação"**
3. O sistema irá:
   - Conectar ao banco PostgreSQL
   - Buscar pessoas com contas e movimentações no período
   - Gerar lotes com até 50 eventos cada (conforme configuração)
   - Assinar e criptografar cada lote
   - Enviar para e-Financeira (se não estiver marcado "Apenas Processar")
4. O progresso será exibido em tempo real
5. Os protocolos serão salvos automaticamente

### 4. Processar Fechamento

1. Clique em **"Processar Fechamento"**
2. O sistema irá gerar, assinar, criptografar e enviar o evento de fechamento

### 5. Consultar Protocolos

1. Vá para a aba **"Consulta"**
2. Digite o protocolo ou selecione um lote da lista
3. Clique em **"Consultar"**
4. O status do lote será exibido com detalhes e ocorrências (se houver)

### 6. Modo "Apenas Processar"

- Marque a opção **"Apenas Processar (não enviar)"** para gerar os XMLs sem enviar para a e-Financeira
- Útil para validação antes do envio real

---

## 🏗 Arquitetura e Funcionamento

### Fluxo de Processamento

#### 1. **Geração de XML**
```
Dados de Configuração → EfinanceiraGeradorXmlService → XML Conforme XSD
```

O serviço `EfinanceiraGeradorXmlService` gera XMLs seguindo os namespaces e estruturas definidas nos schemas XSD da e-Financeira:
- Namespace de lote: `http://www.eFinanceira.gov.br/schemas/envioLoteEventosAssincrono/v1_0_0`
- Namespace de abertura: `http://www.eFinanceira.gov.br/schemas/evtAberturaeFinanceira/v1_2_1`
- Namespace de fechamento: `http://www.eFinanceira.gov.br/schemas/evtFechamentoeFinanceira/v1_3_0`
- Namespace de movimentação: `http://www.eFinanceira.gov.br/schemas/evtMovOpFin/v1_2_1`

#### 2. **Assinatura Digital**
```
XML Gerado → EfinanceiraAssinaturaService → XML Assinado (XML-DSig)
```

- Usa algoritmo **RSA-SHA256** (`http://www.w3.org/2001/04/xmldsig-more#rsa-sha256`)
- Digest method: **SHA256** (`http://www.w3.org/2001/04/xmlenc#sha256`)
- Assina cada evento individualmente dentro do lote
- Suporta estruturas de lote com ou sem elemento `<eventos>` intermediário

#### 3. **Criptografia**
```
XML Assinado → EfinanceiraCriptografiaService → XML Criptografado
```

Processo de criptografia híbrida:
1. Gera chave AES-128 aleatória
2. Gera IV (vetor de inicialização) aleatório
3. Criptografa o XML com AES-128-CBC-PKCS7
4. Concatena chave AES + IV
5. Criptografa a chave concatenada com RSA usando o certificado público do servidor
6. Gera XML final com estrutura `loteCriptografado`

#### 4. **Envio**
```
XML Criptografado → EfinanceiraEnvioService → Resposta com Protocolo
```

- Envia via HTTP POST para o endpoint da e-Financeira
- Usa certificado A1 para autenticação SSL/TLS
- Processa resposta XML e extrai:
  - Código de resposta
  - Descrição
  - Protocolo de envio
  - Ocorrências (se houver)

#### 5. **Consulta**
```
Protocolo → EfinanceiraConsultaService → Status do Lote
```

- Consulta via HTTP GET no endpoint de consulta
- Interpreta códigos de resposta:
  - `1`: Lote em processamento
  - `2`: Lote processado com sucesso
  - `3`: Lote processado com ocorrências
  - `4`: Ocorrências na consulta
  - `5`: Lote não encontrado
  - `9`: Erro interno

### Integração com Banco de Dados

O `EfinanceiraDatabaseService` realiza consultas otimizadas ao PostgreSQL:

```sql
SELECT 
    p.idpessoa, p.documento, p.nome, pf.cpf, pf.nacionalidade,
    c.idconta, c.numeroconta, c.digitoconta, c.saldoatual,
    e.logradouro, e.numero, e.complemento, e.bairro, e.cep,
    SUM(CASE WHEN ex.naturezaoperacao = 'C' THEN ex.valoroperacao ELSE 0 END) as TotCreditos,
    SUM(CASE WHEN ex.naturezaoperacao = 'D' THEN ex.valoroperacao ELSE 0 END) as TotDebitos
FROM manager.tb_pessoa p
INNER JOIN manager.tb_pessoafisica pf ON pf.idpessoa = p.idpessoa
INNER JOIN conta.tb_conta c ON c.idpessoa = p.idpessoa
INNER JOIN conta.tb_extrato ex ON ex.idconta = c.idconta
LEFT JOIN manager.tb_endereco e ON e.idpessoa = p.idpessoa AND e.situacao = '1'
WHERE p.situacao = '1'
  AND c.situacao = '1'
  AND pf.cpf IS NOT NULL
  AND EXTRACT(YEAR FROM ex.dataoperacao) = @ano
  AND EXTRACT(MONTH FROM ex.dataoperacao) BETWEEN @mesInicial AND @mesFinal
GROUP BY ...
ORDER BY p.idpessoa
LIMIT @limit OFFSET @offset
```

### Gerenciamento de Períodos

O `EfinanceiraPeriodoUtil` calcula períodos semestrais automaticamente:
- **Período 01 ou 06**: Janeiro a Junho
- **Período 02 ou 12**: Julho a Dezembro
- Calcula datas de início e fim automaticamente
- Valida formato `YYYYMM`

---

## ✨ Melhorias Implementadas

### 1. **Processamento em Lote Otimizado**
- Paginação configurável para evitar sobrecarga de memória
- Controle de eventos por lote (1 a 50)
- Suporte a processamento parcial (offset de registros e eventos)

### 2. **Interface de Usuário Aprimorada**
- Abas organizadas (Tutorial, Configuração, Processamento, Consulta)
- Controles de progresso em tempo real
- Logs detalhados de processamento
- Estatísticas de lotes processados

### 3. **Persistência de Configurações**
- Salva configurações em arquivo XML
- Carrega automaticamente na inicialização
- Suporta múltiplas configurações

### 4. **Registro de Protocolos**
- Salva protocolos retornados automaticamente
- Permite consulta posterior
- Lista de lotes processados

### 5. **Validações Robustas**
- Validação de CNPJ, CPF, CEP, Email
- Validação de datas e períodos
- Validação de campos obrigatórios antes do processamento

### 6. **Tratamento de Erros**
- Mensagens de erro descritivas
- Logs detalhados para debugging
- Tratamento de exceções em todas as camadas

### 7. **Suporte a Modo Teste**
- Configurações específicas para ambiente de teste
- Processamento limitado para validação
- Opção "Apenas Processar" sem envio

---

## 🔄 Como Evoluir

### 1. **Adicionar Novos Tipos de Eventos**

Para adicionar suporte a novos tipos de eventos (ex: `evtCadDeclarante`, `evtCadIntermediario`):

1. **Criar Modelo de Dados**:
   ```csharp
   // Models/DadosCadastroDeclarante.cs
   public class DadosCadastroDeclarante
   {
       public string CnpjDeclarante { get; set; }
       // ... outros campos
   }
   ```

2. **Adicionar Método no Gerador**:
   ```csharp
   // Services/EfinanceiraGeradorXmlService.cs
   public string GerarXmlCadastroDeclarante(DadosCadastroDeclarante dados, string diretorioSaida)
   {
       // Implementar geração de XML conforme schema XSD
   }
   ```

3. **Adicionar Suporte na Assinatura**:
   ```csharp
   // Services/EfinanceiraAssinaturaService.cs
   private string ObtemTagEventoAssinar(XmlDocument arquivo)
   {
       // Adicionar: if (arquivo.OuterXml.Contains("evtCadDeclarante")) ...
   }
   ```

4. **Adicionar Interface no ProcessamentoForm**:
   - Botão para processar novo tipo de evento
   - Validação de dados específicos

### 2. **Melhorar Performance**

- **Processamento Assíncrono**: Já implementado com `Task.Run()`, pode ser melhorado com `async/await` completo
- **Cache de Certificados**: Cachear certificados carregados para evitar buscas repetidas
- **Otimização de Consultas**: Índices no banco de dados para campos usados em WHERE e JOIN
- **Processamento Paralelo**: Processar múltiplos lotes em paralelo (com cuidado para não sobrecarregar)

### 3. **Adicionar Funcionalidades**

#### **Retry Automático**
```csharp
// Adicionar em EfinanceiraEnvioService
public RespostaEnvioEfinanceira EnviarLoteComRetry(string caminhoArquivo, EfinanceiraConfig config, X509Certificate2 certificado, int maxTentativas = 3)
{
    for (int i = 0; i < maxTentativas; i++)
    {
        try
        {
            return EnviarLote(caminhoArquivo, config, certificado);
        }
        catch (Exception ex)
        {
            if (i == maxTentativas - 1) throw;
            Thread.Sleep(1000 * (i + 1)); // Backoff exponencial
        }
    }
}
```

#### **Validação de XML contra XSD**
```csharp
// Adicionar validação antes do envio
public bool ValidarXmlContraXsd(string caminhoXml, string caminhoXsd)
{
    // Usar XmlSchemaSet para validar
}
```

#### **Relatórios e Estatísticas**
- Exportar relatórios em PDF/Excel
- Dashboard com estatísticas de envios
- Histórico de processamentos

#### **Notificações**
- Notificações por email quando lote for processado
- Alertas para erros críticos

### 4. **Melhorar Segurança**

- **Criptografia de Configurações**: Criptografar arquivo de configuração com senha
- **Logs de Auditoria**: Registrar todas as operações críticas
- **Validação de Certificados**: Verificar validade e revogação de certificados

### 5. **Migrar para .NET Core/.NET 6+**

Para modernizar e permitir multiplataforma:

1. Criar novo projeto .NET 6+
2. Migrar Windows Forms para alternativa multiplataforma (ex: Avalonia, MAUI)
3. Atualizar dependências NuGet
4. Ajustar APIs que mudaram entre .NET Framework e .NET 6+

### 6. **Adicionar Testes**

```csharp
// Exemplo de teste unitário
[Test]
public void TestarGeracaoXmlAbertura()
{
    var dados = new DadosAbertura { /* ... */ };
    var service = new EfinanceiraGeradorXmlService();
    var xml = service.GerarXmlAbertura(dados, @"C:\temp");
    Assert.IsTrue(File.Exists(xml));
}
```

### 7. **Documentação de API**

Adicionar documentação XML nos métodos públicos:

```csharp
/// <summary>
/// Gera XML de evento de abertura conforme schema XSD da e-Financeira.
/// </summary>
/// <param name="dados">Dados de abertura preenchidos</param>
/// <param name="diretorioSaida">Diretório onde o XML será salvo</param>
/// <returns>Caminho completo do arquivo XML gerado</returns>
/// <exception cref="ArgumentException">Quando dados obrigatórios estão faltando</exception>
public string GerarXmlAbertura(DadosAbertura dados, string diretorioSaida)
{
    // ...
}
```

---

## 🔧 Troubleshooting

### Problema: Certificado não encontrado

**Solução**:
1. Verifique se o certificado está instalado no repositório correto (`CurrentUser\My`)
2. Verifique se o thumbprint está correto (sem espaços ou hífens)
3. Certifique-se de que o certificado tem chave privada (A1)
4. Tente reinstalar o certificado

### Problema: Erro ao conectar ao banco de dados

**Solução**:
1. Verifique as credenciais no código (`EfinanceiraDatabaseService.cs`)
2. Teste a conexão usando o botão "Testar Conexão BD"
3. Verifique firewall e permissões de rede
4. Confirme que o PostgreSQL está rodando

### Problema: XML rejeitado pela e-Financeira

**Solução**:
1. Verifique se o XML está conforme o schema XSD (valide manualmente)
2. Verifique se a assinatura digital está correta
3. Verifique se os dados obrigatórios estão preenchidos
4. Consulte as ocorrências retornadas na resposta

### Problema: Erro de criptografia

**Solução**:
1. Verifique se o certificado do servidor está instalado corretamente
2. Verifique se o thumbprint do certificado do servidor está correto
3. Tente reinstalar o certificado público da Receita Federal

### Problema: Processamento lento

**Solução**:
1. Aumente o `Page Size` se tiver memória disponível
2. Reduza `Eventos por Lote` se houver problemas de timeout
3. Verifique a performance do banco de dados (índices, estatísticas)
4. Considere processar em horários de menor carga

### Problema: Período inválido

**Solução**:
1. Use formato `YYYYMM` (ex: `202301` ou `202302`)
2. Use `01` ou `06` para primeiro semestre (Jan-Jun)
3. Use `02` ou `12` para segundo semestre (Jul-Dez)
4. Use o botão "Calcular Período Atual" na tela de consulta

---

## 📝 Notas Importantes

1. **Certificados**: Sempre mantenha backups dos certificados e senhas em local seguro.

2. **Ambiente de Teste**: Sempre teste no ambiente de homologação antes de enviar para produção.

3. **Backup de XMLs**: Mantenha backups dos XMLs gerados, assinados e criptografados.

4. **Protocolos**: Guarde os protocolos retornados, pois são necessários para consultas futuras.

5. **Períodos**: O sistema processa períodos semestrais. Certifique-se de processar ambos os semestres do ano.

6. **Validação**: Sempre valide os dados antes do processamento em produção.

7. **Logs**: Monitore os logs para identificar problemas rapidamente.

---

## 📞 Suporte

Para questões técnicas ou problemas:
1. Consulte os logs da aplicação
2. Verifique as mensagens de erro detalhadas
3. Consulte a documentação oficial da e-Financeira
4. Revise os schemas XSD fornecidos pela Receita Federal

---

## 📄 Licença

Este projeto é de uso interno. Consulte a política de licenciamento da organização.

---

## 🔄 Histórico de Versões

- **v1.0**: Versão inicial com suporte a abertura, movimentação e fechamento
- Funcionalidades de processamento em lote
- Interface com abas
- Persistência de configurações
- Consulta de protocolos

---

**Desenvolvido para automatizar e simplificar o processo de declaração e-Financeira.**
