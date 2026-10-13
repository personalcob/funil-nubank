/* =====================================================================================
   FUNIL DE COBRANÇA NUBANK — CC + LL (consulta única)                                VERSÃO 8
   -------------------------------------------------------------------------------------
   Junta as duas consultas originais (CONSULTA_CC e CONSULTA_LL) sem mudar o conceito dos indicadores:
     • CC (cartão, carteira 5): regras da CONSULTA_CC  — chave do contrato = collection__id.
     • LL (empréstimo, carteira 6): regras da CONSULTA_LL — foto FOTO_MATCHING_EMPRESTIMO, filtro por
       DATA_MATCHING, saldo = soma dos loans do contrato no dia (pico do período), bucket por cliente+contrato;
       pagamento do arquivo do LL ligado ao contrato pelo loan__id (DIV_AUX.CAMPO028 -> CAMPO001).
     • Faixas de bucket padronizadas no de/para do CC (+ rótulos do LL, os mesmos da aba AUX do Excel).
   Resultado: as 12 colunas de sempre da aba "Dados" + a coluna FUNIL no final ('CC' ou 'LL').
   Performance: as leituras pesadas (fotos e discagens) usam o cache incremental; discagens, acordos e o
   arquivo de empréstimo são lidos UMA vez por período e servem às duas carteiras.
   Períodos: automáticos e comparáveis em dias úteis (atual x anterior); o mesmo período vale para CC e LL.

   VALIDAR CONTRA AS ORIGINAIS (setembro inteiro, como nas consultas do projeto):
     @DATA_BASE = '2026-10-01', @DATA_CORTE_MODO = 'D1', @PERIODOS_RODAR = 'ATUAL', @USAR_CACHE = 0
   ===================================================================================== */
SET NOCOUNT ON;

/* =====================================================================================
   0) PARÂMETROS
   ===================================================================================== */

-- 0.A) PERÍODO E EXECUÇÃO -------------------------------------------------------------
DECLARE @PERIODOS_RODAR          VARCHAR(10) = 'AMBOS';      -- 'AMBOS' (comparativo) | 'ATUAL' | 'ANTERIOR'
DECLARE @CARTEIRAS_RODAR         VARCHAR(10) = 'AMBAS';      -- 'AMBAS' | 'CC' | 'LL'
DECLARE @DIAS_UTEIS_FORCADO      INT         = NULL;         -- NULL = automático | N = força N dias úteis nos dois meses
DECLARE @DATA_CORTE_MODO         VARCHAR(10) = 'D1';  -- 'PAGAMENTO' = último pagamento (arquivo do CC; só LL -> do LL) | 'D1' = ontem
DECLARE @CORTE_SO_DOWNPAYMENT    BIT         = 1;            -- CC: 1 = corte pelo último direct+downpayment | 0 = qualquer direct
DECLARE @CONSIDERAR_FERIADOS     BIT         = 1;            -- 1 = feriados da lista 0.D não contam como dia útil
DECLARE @ESTENDER_FIM_NAO_UTIL   BIT         = 0;            -- 1 = se o atual termina em dia não útil, o anterior vai até a véspera do próximo d.u.
DECLARE @DATA_BASE               DATE        = NULL ;         -- NULL = hoje | data = simula a execução nesse dia (períodos anteriores)
DECLARE @APENAS_MOSTRAR_PERIODOS BIT         = 0;            -- 1 = só mostra as datas calculadas e para (segundos)
DECLARE @MOSTRAR_RESUMO          BIT         = 1;            -- 1 = grade extra com o total por etapa, período e carteira
DECLARE @MOSTRAR_DIAGNOSTICO     BIT         = 1;            -- 1 = avisos extras (contrato repetido, loan sem contrato...) | 0 no dia a dia
DECLARE @DECIMAL_VIRGULA         BIT         = 1;            -- 1 = BALANCE e CASH com 2 casas e VÍRGULA (colar no Excel pt-BR) | 0 = com ponto
DECLARE @EXCLUIR_BUCKETS_INVALIDOS BIT      = 1;            -- 1 = tira "(vazio)" e "00.xx" de TODAS as etapas, como no Excel | 0 = entrega tudo
DECLARE @USAR_CACHE              BIT         = 1;            -- 1 = cache incremental (1ª rodada monta, as seguintes só leem o dia novo) | 0 = lê tudo direto
DECLARE @RECRIAR_CACHE           BIT         = 0;            -- 1 = apaga e refaz o cache dos períodos desta rodada (dado antigo corrigido)

-- 0.B) ARQUIVOS ------------------------------------------------------------------------
-- Sempre caminho de rede (\\servidor\...): o SQL Server não enxerga letra de unidade nem o Desktop.
DECLARE @ORIGEM_PAGAMENTOS       VARCHAR(10)   = 'MANUAL';     -- 'MANUAL' (pasta MAIO) | 'ROBO' (pasta TESTE_ROBO) — vale para CC e LL
DECLARE @PASTA_PAGAMENTOS_MANUAL NVARCHAR(400) = N'\\10.10.2.55\planejamento\0. Time Plan - Mis (Lucas,Mikas,Pam e Italo)\02. Mikas\NUBANK\PAGAMENTOS\MAIO\';
DECLARE @PASTA_PAGAMENTOS_ROBO   NVARCHAR(400) = N'\\10.10.2.55\planejamento\0. Time Plan - Mis (Lucas,Mikas,Pam e Italo)\02. Mikas\NUBANK\PAGAMENTOS\TESTE_ROBO\';
DECLARE @ARQUIVO_PAGAMENTOS_CC   NVARCHAR(100) = N'pagamentos_CC_122.xlsx';
DECLARE @ARQUIVO_PAGAMENTOS_LL   NVARCHAR(100) = N'pagamentos_LL_122.xlsx';
DECLARE @ABA_PAGAMENTOS          NVARCHAR(50)  = N'Planilha1';
DECLARE @PASTA_DISCAGENS         NVARCHAR(400) = N'\\10.10.2.55\planejamento\2 - MIS\1 - Relatórios\10 - Nubank\Arquivos_Envio_Cliente\Enviados\';
DECLARE @PREFIXO_DISCAGENS       NVARCHAR(50)  = N'personalcob_';
DECLARE @DIAS_FOLGA_DISCAGENS    INT           = 3;          -- lê até N dias após o fim do período (o arquivo de um dia traz as ligações do anterior)

-- 0.C) REGRAS DE NEGÓCIO ---------------------------------------------------------------
DECLARE @CARTEIRA_CC               INT  = 5;
DECLARE @CARTEIRA_LL               INT  = 6;
DECLARE @MARCADOR_DESENROLA        INT  = 77;
DECLARE @MARCADOR_CAMPANHA_INTERNA INT  = 119;
DECLARE @MARCADORES_DESDE          DATE = '2026-04-01';      -- marcadores contam a partir desta data
DECLARE @DIAS_ANTES_VENC           INT  = 2;                 -- pagamento indireto conta se cair entre D-2 ...
DECLARE @DIAS_APOS_VENC            INT  = 2;                 -- ... e D+2 do vencimento do acordo da mesma carteira
DECLARE @LL_VINCULO_PAGAMENTO      VARCHAR(1)  = 'C';          -- LL: a que contrato vai o pagamento de um loan que existe em mais de um contrato (DIV_AUX CAMPO028 -> CAMPO001):
                                                               --   'A' = a TODOS os contratos do loan (igual à consulta original do LL: o pagamento conta em dobro)
                                                               --   'B' = só aos contratos que estão na base do período
                                                               --   'C' = a UM contrato só (o da base; se houver mais de um na base, o mais recente = maior ID_DIVIDA)
DECLARE @LL_DEAL_TYPE_IGNORADO     VARCHAR(50) = 'not-commissioned';   -- LL: tipo de pagamento que não entra (o robô já tira)

IF OBJECT_ID('tempdb..#USUARIOS_EXCLUIDOS') IS NOT NULL DROP TABLE #USUARIOS_EXCLUIDOS;
CREATE TABLE #USUARIOS_EXCLUIDOS (ID_USUARIO INT PRIMARY KEY);   -- usuários ignorados nos acordos
INSERT INTO #USUARIOS_EXCLUIDOS (ID_USUARIO) VALUES (1), (818), (888), (2199), (3064);

-- De/para de bucket = o do CC (CONSULTA_CC) + os rótulos do LL (os mesmos da aba AUX do Excel).
-- O LL escreve alguns rótulos com ':' (ex.: 01:1-30): a consulta troca ':' por '.' antes de procurar aqui.
-- Bucket fora da lista fica como veio (e aparece no aviso "buckets fora do padrão").
IF OBJECT_ID('tempdb..#DE_PARA_BUCKET') IS NOT NULL DROP TABLE #DE_PARA_BUCKET;
CREATE TABLE #DE_PARA_BUCKET (ORIGEM VARCHAR(30) NOT NULL PRIMARY KEY, DESTINO VARCHAR(30) NOT NULL);
INSERT INTO #DE_PARA_BUCKET (ORIGEM, DESTINO) VALUES
('01.1-30','0.0-60'),('02.31-60','0.0-60'),('01.0-60','0.0-60'),('0.0-60','0.0-60'),('0.1-60','0.0-60'),
('03.61-90','1.61-90'),('1.61-90','1.61-90'),('01.61-90','1.61-90'),('02.61-90','1.61-90'),
('04.91-120','2.91-120'),('2.91-120','2.91-120'),('02.91-120','2.91-120'),('03.91-120','2.91-120'),
('05.121-180','3.121-180'),('3.121-180','3.121-180'),('03.121-180','3.121-180'),('04.121-150','3.121-180'),('05.151-180','3.121-180'),
('06.181-360','4.181-360'),('4.181-360','4.181-360'),('06.181-210','4.181-360'),('07.211-240','4.181-360'),('08.241-270','4.181-360'),
('09.271-300','4.181-360'),('10.301-330','4.181-360'),('11.331-360','4.181-360'),('05.181-210','4.181-360'),('06.211-240','4.181-360'),
('07.241-270','4.181-360'),('08.271-300','4.181-360'),('09.301-330','4.181-360'),('10.331-360','4.181-360'),('12.361-390','4.181-360'),
('07.361-720','5.361-720'),('5.361-720','5.361-720'),('05.361-720','5.361-720'),('13.391-720','5.361-720'),('11.361-390','5.361-720'),('12.391-720','5.361-720'),
('08.721-1080','6.721-1080'),('6.721-1080','6.721-1080'),('06.721-1080','6.721-1080'),('13.721-1080','6.721-1080'),('14.721-1080','6.721-1080'),
('09.1081-1440','7.1081-1440'),('7.1081-1440','7.1081-1440'),('07.1081-1440','7.1081-1440'),
('10.1440+','8.1441+'),('15.1081+','8.1441+'),('14.1081+','8.1441+'),('8.1441+','8.1441+'),('8.1440+','8.1441+'),('09.1081+','8.1441+'),
-- rótulos do LL (aba AUX do Excel)
('06.181-270','4.181-360'),('07.271-360','4.181-360'),('08.361-720','5.361-720'),
('08.721+','6.721-1080'),('09.721+','6.721-1080'),('09.1080+','7.1081-1440');

-- 0.D) CALENDÁRIO ----------------------------------------------------------------------
DECLARE @FERIADOS TABLE (DATA DATE PRIMARY KEY);   -- mesma lista do config.yaml do robô
INSERT INTO @FERIADOS (DATA) VALUES
('2026-01-01'),('2026-02-16'),('2026-02-17'),('2026-04-03'),('2026-04-21'),('2026-05-01'),('2026-06-04'),
('2026-09-07'),('2026-10-12'),('2026-11-02'),('2026-11-15'),('2026-11-20'),('2026-12-25'),
('2027-01-01'),('2027-02-08'),('2027-02-09'),('2027-03-26'),('2027-04-21'),('2027-05-01'),('2027-05-27'),
('2027-09-07'),('2027-10-12'),('2027-11-02'),('2027-11-15'),('2027-11-20'),('2027-12-25');

-- Nome do mês: rótulo do MES ("Setembro/2026"), tabelas FOTO_MATCHING_<Mês>_<Ano> e pastas de discagem.
-- Março é testado com e sem cedilha automaticamente.
DECLARE @MESES TABLE (NUM INT PRIMARY KEY, NOME NVARCHAR(20));
INSERT INTO @MESES (NUM, NOME) VALUES
(1,N'Janeiro'),(2,N'Fevereiro'),(3,N'Março'),(4,N'Abril'),(5,N'Maio'),(6,N'Junho'),
(7,N'Julho'),(8,N'Agosto'),(9,N'Setembro'),(10,N'Outubro'),(11,N'Novembro'),(12,N'Dezembro');

/* =====================================================================================
   DAQUI PARA BAIXO NÃO HÁ O QUE EDITAR
   ===================================================================================== */
DECLARE @MSG     NVARCHAR(2000);   -- mensagens em tempo real: RAISERROR(@MSG, 0, 1) WITH NOWAIT
DECLARE @SQL     NVARCHAR(MAX);
DECLARE @LINHAS  BIGINT;
DECLARE @T_GERAL DATETIME2 = SYSDATETIME();
DECLARE @T_SECAO DATETIME2 = SYSDATETIME();

/* =====================================================================================
   1) VALIDAÇÃO E PAGAMENTOS (CC e LL) — o corte do período depende deles
      Os dois arquivos viram UMA tabela (#PAG), já com o contrato (collection__id) de cada pagamento:
        CC: o arquivo já traz o collection__id;
        LL: o arquivo traz o loan__id -> contrato pela DIV_AUX (CAMPO028 = loan__id, CAMPO001 = contrato), carteira 6.
   ===================================================================================== */
IF @PERIODOS_RODAR NOT IN ('AMBOS', 'ATUAL', 'ANTERIOR')
   OR @CARTEIRAS_RODAR NOT IN ('AMBAS', 'CC', 'LL')
   OR @DATA_CORTE_MODO NOT IN ('PAGAMENTO', 'D1')
   OR @ORIGEM_PAGAMENTOS NOT IN ('MANUAL', 'ROBO')
   OR @LL_VINCULO_PAGAMENTO NOT IN ('A', 'B', 'C')
BEGIN
    RAISERROR('Parâmetro inválido. Use @PERIODOS_RODAR = AMBOS/ATUAL/ANTERIOR, @CARTEIRAS_RODAR = AMBAS/CC/LL, @DATA_CORTE_MODO = PAGAMENTO/D1, @ORIGEM_PAGAMENTOS = MANUAL/ROBO, @LL_VINCULO_PAGAMENTO = A/B/C.', 16, 1);
    RETURN;
END;

DECLARE @CARTEIRAS TABLE (ORDEM INT PRIMARY KEY, CARTEIRA VARCHAR(2), ID INT);
INSERT INTO @CARTEIRAS (ORDEM, CARTEIRA, ID)
SELECT C.ORDEM, C.CARTEIRA, C.ID
FROM (VALUES (1, 'CC', @CARTEIRA_CC), (2, 'LL', @CARTEIRA_LL)) C (ORDEM, CARTEIRA, ID)
WHERE @CARTEIRAS_RODAR IN ('AMBAS', C.CARTEIRA);

DECLARE @PASTA_PAG   NVARCHAR(400) = CASE WHEN @ORIGEM_PAGAMENTOS = 'ROBO' THEN @PASTA_PAGAMENTOS_ROBO ELSE @PASTA_PAGAMENTOS_MANUAL END;
DECLARE @ARQUIVO_PAG NVARCHAR(600);

IF OBJECT_ID('tempdb..#PAG') IS NOT NULL DROP TABLE #PAG;
CREATE TABLE #PAG (
    CARTEIRA                 VARCHAR(2) NOT NULL,
    customer__id             VARCHAR(36),
    collection__id           VARCHAR(36),
    DT_HORA                  DATETIME2(0),   -- data e hora: desempata o "último pagamento" do contrato
    DT_PAGAMENTO             DATE,
    payment_type             VARCHAR(50),    -- CC: payment_type | LL: deal_type
    new_old                  VARCHAR(50),    -- só CC
    finance_days_late_bucket VARCHAR(50),    -- CC: finance_days_late_bucket | LL: bucket
    VALOR                    DECIMAL(18,2),  -- CC: total_payment | LL: payment
    loan__id                 VARCHAR(36),    -- só LL (o pagamento conta em um contrato por loan, conforme @LL_VINCULO_PAGAMENTO)
    ID_DIVIDA_REC            BIGINT          -- só LL: maior ID_DIVIDA do par loan-contrato na DIV_AUX (quanto maior, mais recente)
);

-- 1.1) CC ----------------------------------------------------------------------------------------
IF EXISTS (SELECT 1 FROM @CARTEIRAS WHERE CARTEIRA = 'CC')
BEGIN
    SET @ARQUIVO_PAG = @PASTA_PAG + @ARQUIVO_PAGAMENTOS_CC;
    SET @MSG = REPLACE(CONCAT(N'Lendo pagamentos CC: ', @ARQUIVO_PAG), N'%', N'%%');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;
    -- O caminho do OPENROWSET não aceita variável: o comando é montado como texto.
    -- Data e valor aceitam o texto do robô ou data/número do Excel. Coluna que faltar: "Invalid column name".
    SET @SQL = N'
    INSERT INTO #PAG WITH (TABLOCK)
        (CARTEIRA, customer__id, collection__id, DT_HORA, DT_PAGAMENTO, payment_type, new_old, finance_days_late_bucket, VALOR)
    SELECT ''CC'',
           CAST(X.customer__id   AS VARCHAR(36)),
           CAST(X.collection__id AS VARCHAR(36)),
           D.DT_HORA,
           CAST(D.DT_HORA AS DATE),
           CAST(X.payment_type             AS VARCHAR(50)),
           CAST(X.new_old                  AS VARCHAR(50)),
           CAST(X.finance_days_late_bucket AS VARCHAR(50)),
           COALESCE(TRY_CONVERT(DECIMAL(18,2), X.total_payment),
                    TRY_CONVERT(DECIMAL(18,2), REPLACE(CAST(X.total_payment AS NVARCHAR(100)), N'','', N''.'')))
    FROM OPENROWSET(''MICROSOFT.ACE.OLEDB.12.0'',
                    ''EXCEL 12.0;DATABASE=' + REPLACE(@ARQUIVO_PAG, N'''', N'''''') + N';'',
                    ''SELECT * FROM [' + @ABA_PAGAMENTOS + N'$]'') X
    CROSS APPLY (SELECT COALESCE(TRY_CONVERT(DATETIME2(0), X.[date]),
                                 TRY_CONVERT(DATETIME2(0), LEFT(LTRIM(CAST(X.[date] AS NVARCHAR(50))), 10), 103)) AS DT_HORA) D;';
    EXEC sp_executesql @SQL;
END;

-- 1.2) LL ----------------------------------------------------------------------------------------
IF EXISTS (SELECT 1 FROM @CARTEIRAS WHERE CARTEIRA = 'LL')
BEGIN
    SET @ARQUIVO_PAG = @PASTA_PAG + @ARQUIVO_PAGAMENTOS_LL;
    SET @MSG = REPLACE(CONCAT(N'Lendo pagamentos LL: ', @ARQUIVO_PAG), N'%', N'%%');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;
    IF OBJECT_ID('tempdb..#PAG_LL') IS NOT NULL DROP TABLE #PAG_LL;
    CREATE TABLE #PAG_LL (customer__id VARCHAR(36), loan__id VARCHAR(36), DT_HORA DATETIME2(0), DT_PAGAMENTO DATE,
                          deal_type VARCHAR(50), bucket VARCHAR(50), VALOR DECIMAL(18,2));
    SET @SQL = N'
    INSERT INTO #PAG_LL WITH (TABLOCK) (customer__id, loan__id, DT_HORA, DT_PAGAMENTO, deal_type, bucket, VALOR)
    SELECT CAST(X.customer__id AS VARCHAR(36)),
           CAST(X.loan__id     AS VARCHAR(36)),
           D.DT_HORA,
           CAST(D.DT_HORA AS DATE),
           CAST(X.deal_type AS VARCHAR(50)),
           CAST(X.bucket    AS VARCHAR(50)),
           COALESCE(TRY_CONVERT(DECIMAL(18,2), X.payment),
                    TRY_CONVERT(DECIMAL(18,2), REPLACE(CAST(X.payment AS NVARCHAR(100)), N'','', N''.'')))
    FROM OPENROWSET(''MICROSOFT.ACE.OLEDB.12.0'',
                    ''EXCEL 12.0;DATABASE=' + REPLACE(@ARQUIVO_PAG, N'''', N'''''') + N';'',
                    ''SELECT * FROM [' + @ABA_PAGAMENTOS + N'$]'') X
    CROSS APPLY (SELECT COALESCE(TRY_CONVERT(DATETIME2(0), X.[date]),
                                 TRY_CONVERT(DATETIME2(0), LEFT(LTRIM(CAST(X.[date] AS NVARCHAR(50))), 10), 103)) AS DT_HORA) D
    WHERE ISNULL(CAST(X.deal_type AS VARCHAR(50)), '''') <> @P_IGNORADO;';
    EXEC sp_executesql @SQL, N'@P_IGNORADO VARCHAR(50)', @P_IGNORADO = @LL_DEAL_TYPE_IGNORADO;

    -- loan__id -> contrato (regra da CONSULTA_LL: DIV_AUX.CAMPO028 = loan__id, DIV_MASTER na carteira 6).
    -- Lido UMA vez e sem repetir linha: se a DIV_AUX tiver o mesmo loan em mais de uma linha, o pagamento conta uma vez.
    IF OBJECT_ID('tempdb..#LOANS') IS NOT NULL DROP TABLE #LOANS;
    SELECT DISTINCT loan__id INTO #LOANS FROM #PAG_LL WHERE loan__id IS NOT NULL;
    CREATE UNIQUE CLUSTERED INDEX IX ON #LOANS (loan__id);
    IF OBJECT_ID('tempdb..#LOAN_COLL') IS NOT NULL DROP TABLE #LOAN_COLL;
    SELECT CAST(FA.CAMPO028 AS VARCHAR(36)) AS loan__id,
           CAST(FA.CAMPO001 AS VARCHAR(36)) AS collection__id,
           COUNT(*)                         AS LINHAS_DIV,
           MAX(TRY_CAST(FA.ID_DIVIDA AS BIGINT)) AS ID_DIVIDA_REC
    INTO #LOAN_COLL
    FROM [COBwebNUBANK].[dbo].[DIV_AUX] FA WITH (NOLOCK)
    INNER JOIN [COBwebNUBANK].[dbo].[DIV_MASTER] DM WITH (NOLOCK) ON DM.ID_DIVIDA = FA.ID_DIVIDA
    WHERE DM.ID_CARTEIRA = @CARTEIRA_LL
      AND EXISTS (SELECT 1 FROM #LOANS L WHERE L.loan__id = FA.CAMPO028)
    GROUP BY CAST(FA.CAMPO028 AS VARCHAR(36)), CAST(FA.CAMPO001 AS VARCHAR(36))
    OPTION (RECOMPILE);
    CREATE CLUSTERED INDEX IX ON #LOAN_COLL (loan__id);

    INSERT INTO #PAG WITH (TABLOCK)
        (CARTEIRA, customer__id, collection__id, DT_HORA, DT_PAGAMENTO, payment_type, new_old, finance_days_late_bucket, VALOR, loan__id, ID_DIVIDA_REC)
    SELECT 'LL', P.customer__id, M.collection__id, P.DT_HORA, P.DT_PAGAMENTO, P.deal_type, NULL, P.bucket, P.VALOR, P.loan__id, M.ID_DIVIDA_REC
    FROM #PAG_LL P
    INNER JOIN #LOAN_COLL M ON M.loan__id = P.loan__id;

    SET @MSG = CONCAT(N'Pagamentos LL: ', (SELECT COUNT(*) FROM #PAG_LL), N' linhas no arquivo | ',
                      (SELECT COUNT(*) FROM #PAG_LL P WHERE NOT EXISTS (SELECT 1 FROM #LOAN_COLL M WHERE M.loan__id = P.loan__id)),
                      N' sem contrato na carteira 6 (ficam de fora, como na consulta do LL)');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;
    IF @MOSTRAR_DIAGNOSTICO = 1
    BEGIN
        SET @MSG = CONCAT(N'  DIAGNÓSTICO LL: ',
                          (SELECT COUNT(*) FROM (SELECT loan__id FROM #LOAN_COLL GROUP BY loan__id HAVING COUNT(*) > 1) X),
                          N' loan(s) ligados a mais de um contrato na DIV_AUX (regra do vínculo: ', @LL_VINCULO_PAGAMENTO, N') | ',
                          (SELECT COUNT(*) FROM #LOAN_COLL WHERE LINHAS_DIV > 1),
                          N' loan(s) com linha repetida na DIV_AUX (na consulta original o pagamento somaria em dobro; aqui conta 1x)');
        RAISERROR(@MSG, 0, 1) WITH NOWAIT;
    END;
    DROP TABLE #PAG_LL;
    DROP TABLE #LOANS;
    -- #LOAN_COLL fica: a escolha do contrato de cada loan é feita por período (seção 4.5)
END;

CREATE CLUSTERED INDEX IX ON #PAG (CARTEIRA, DT_PAGAMENTO);
SELECT @MSG = STUFF((
    SELECT NCHAR(10) + CONCAT(N'Pagamentos ', CARTEIRA, N': ', COUNT(*), N' linhas | de ',
                              ISNULL(CONVERT(VARCHAR(10), MIN(DT_PAGAMENTO), 103), N'?'), N' a ',
                              ISNULL(CONVERT(VARCHAR(10), MAX(DT_PAGAMENTO), 103), N'?'),
                              CASE WHEN SUM(CASE WHEN DT_PAGAMENTO IS NULL OR VALOR IS NULL THEN 1 ELSE 0 END) > 0
                                   THEN CONCAT(N' | AVISO: ', SUM(CASE WHEN DT_PAGAMENTO IS NULL OR VALOR IS NULL THEN 1 ELSE 0 END),
                                               N' sem data ou valor legível (ficam de fora)') ELSE N'' END)
    FROM #PAG GROUP BY CARTEIRA ORDER BY CARTEIRA
    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(2000)'), 1, 1, N'');
SET @MSG = REPLACE(ISNULL(@MSG, N'Nenhum pagamento carregado.'), N'%', N'%%');
RAISERROR(@MSG, 0, 1) WITH NOWAIT;

/* =====================================================================================
   2) PERÍODOS — data de corte e dias úteis
      ATUAL    = 1º dia do mês do corte até o corte (N dias úteis)
      ANTERIOR = 1º dia do mês anterior até o N-ésimo dia útil dele
   ===================================================================================== */
DECLARE @HOJE      DATE = ISNULL(@DATA_BASE, CAST(GETDATE() AS DATE));
DECLARE @D_MENOS_1 DATE = DATEADD(DAY, -1, @HOJE);

-- O corte sai do arquivo do CC (como sempre); rodando só o LL, sai do arquivo do LL (pagamento direct).
-- Os dois funis usam o MESMO período, para o comparativo ser igual nas duas carteiras.
DECLARE @CART_CORTE VARCHAR(2) = CASE WHEN @CARTEIRAS_RODAR = 'LL' THEN 'LL' ELSE 'CC' END;
DECLARE @MAX_PAGAMENTO DATE = (
    SELECT MAX(DT_PAGAMENTO) FROM #PAG
    WHERE CARTEIRA = @CART_CORTE AND payment_type = 'direct'
      AND (@CART_CORTE = 'LL' OR @CORTE_SO_DOWNPAYMENT = 0 OR new_old = 'downpayment')
);
DECLARE @DATA_CORTE DATE =
    CASE WHEN @DATA_CORTE_MODO = 'PAGAMENTO' AND @MAX_PAGAMENTO IS NOT NULL
         THEN CASE WHEN @MAX_PAGAMENTO > @D_MENOS_1 THEN @D_MENOS_1 ELSE @MAX_PAGAMENTO END
         ELSE @D_MENOS_1 END;
IF @DATA_CORTE_MODO = 'PAGAMENTO' AND @MAX_PAGAMENTO IS NULL
    RAISERROR('AVISO: nenhum pagamento direct válido no arquivo de pagamentos; usando D-1 como corte.', 0, 1) WITH NOWAIT;

DECLARE @INI_ATUAL DATE = DATEFROMPARTS(YEAR(@DATA_CORTE), MONTH(@DATA_CORTE), 1);
DECLARE @INI_ANT   DATE = DATEADD(MONTH, -1, @INI_ATUAL);

-- Fim de semana pelo resto da divisão por 7 a partir de 01/01/1900 (segunda): 5 = sáb, 6 = dom.
-- Não depende de SET DATEFIRST.
DECLARE @CALENDARIO TABLE (DATA DATE PRIMARY KEY, UTIL BIT NOT NULL, N_UTIL INT NULL);
;WITH N AS (SELECT TOP (62) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) - 1 AS I FROM sys.all_objects)
INSERT INTO @CALENDARIO (DATA, UTIL)
SELECT D.DATA,
       CASE WHEN DATEDIFF(DAY, '19000101', D.DATA) % 7 IN (5, 6) THEN 0
            WHEN @CONSIDERAR_FERIADOS = 1 AND EXISTS (SELECT 1 FROM @FERIADOS F WHERE F.DATA = D.DATA) THEN 0
            ELSE 1 END
FROM N
CROSS APPLY (SELECT DATEADD(DAY, N.I, @INI_ANT) AS DATA) D
WHERE D.DATA <= EOMONTH(@INI_ATUAL);

UPDATE C SET N_UTIL = X.RN               -- posição do dia útil dentro do seu mês (1º, 2º, ...)
FROM @CALENDARIO C
JOIN (SELECT DATA, ROW_NUMBER() OVER (PARTITION BY YEAR(DATA), MONTH(DATA) ORDER BY DATA) AS RN
      FROM @CALENDARIO WHERE UTIL = 1) X ON X.DATA = C.DATA;

DECLARE @N_DIAS_UTEIS INT = ISNULL(@DIAS_UTEIS_FORCADO,
        (SELECT COUNT(*) FROM @CALENDARIO WHERE UTIL = 1 AND DATA BETWEEN @INI_ATUAL AND @DATA_CORTE));
IF @N_DIAS_UTEIS IS NULL OR @N_DIAS_UTEIS < 1
BEGIN
    SET @MSG = CONCAT(N'Nenhum dia útil até a data de corte (', CONVERT(VARCHAR(10), @DATA_CORTE, 103),
                      N'). Use @DIAS_UTEIS_FORCADO, @DATA_CORTE_MODO = ''D1'' ou @DATA_BASE.');
    RAISERROR(@MSG, 16, 1);
    RETURN;
END;

DECLARE @FIM_ATUAL DATE = CASE WHEN @DIAS_UTEIS_FORCADO IS NULL THEN @DATA_CORTE
                               ELSE (SELECT DATA FROM @CALENDARIO WHERE N_UTIL = @N_DIAS_UTEIS
                                     AND DATA BETWEEN @INI_ATUAL AND EOMONTH(@INI_ATUAL)) END;
DECLARE @FIM_ANT   DATE = (SELECT DATA FROM @CALENDARIO WHERE N_UTIL = @N_DIAS_UTEIS
                           AND DATA BETWEEN @INI_ANT AND EOMONTH(@INI_ANT));
IF @FIM_ATUAL IS NULL
BEGIN
    SET @FIM_ATUAL = EOMONTH(@INI_ATUAL);
    RAISERROR('AVISO: o mês atual tem menos dias úteis que o pedido; período vai até o fim do mês.', 0, 1) WITH NOWAIT;
END;
IF @FIM_ANT IS NULL
BEGIN
    SET @FIM_ANT = EOMONTH(@INI_ANT);
    RAISERROR('AVISO: o mês anterior tem menos dias úteis que o pedido; período vai até o fim do mês.', 0, 1) WITH NOWAIT;
END;
IF @DIAS_UTEIS_FORCADO IS NOT NULL AND @FIM_ATUAL > @DATA_CORTE
BEGIN
    SET @MSG = CONCAT(N'AVISO: com ', @N_DIAS_UTEIS, N' d.u. forçados o período atual vai até ',
                      CONVERT(VARCHAR(10), @FIM_ATUAL, 103), N', depois do corte (dias ainda sem pagamento).');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;
END;
IF @ESTENDER_FIM_NAO_UTIL = 1 AND (SELECT UTIL FROM @CALENDARIO WHERE DATA = @FIM_ATUAL) = 0
    SELECT @FIM_ANT = ISNULL(DATEADD(DAY, -1, MIN(DATA)), EOMONTH(@INI_ANT))
    FROM @CALENDARIO WHERE UTIL = 1 AND DATA > @FIM_ANT AND DATA <= EOMONTH(@INI_ANT);

DECLARE @PERIODOS TABLE (ORDEM INT PRIMARY KEY, REFERENCIA VARCHAR(10), INICIO DATE, FIM DATE, DIAS_UTEIS INT);
INSERT INTO @PERIODOS (ORDEM, REFERENCIA, INICIO, FIM, DIAS_UTEIS)
SELECT P.ORDEM, P.REFERENCIA, P.INICIO, P.FIM,
       (SELECT COUNT(*) FROM @CALENDARIO C WHERE C.UTIL = 1 AND C.DATA BETWEEN P.INICIO AND P.FIM)
FROM (VALUES (1, 'ANTERIOR', @INI_ANT, @FIM_ANT), (2, 'ATUAL', @INI_ATUAL, @FIM_ATUAL)) P (ORDEM, REFERENCIA, INICIO, FIM)
WHERE @PERIODOS_RODAR IN ('AMBOS', P.REFERENCIA);

SET @MSG = CONCAT(
    N'=========================================================', NCHAR(10),
    N'Execução em ', CONVERT(VARCHAR(10), @HOJE, 103), N' | ', @N_DIAS_UTEIS, N' dias úteis',
    CASE WHEN @DIAS_UTEIS_FORCADO IS NULL THEN N' (automático)' ELSE N' (forçado)' END,
    CASE WHEN @CONSIDERAR_FERIADOS = 1 THEN N', descontando feriados' ELSE N', só fins de semana' END, NCHAR(10),
    N'Carteiras: ', @CARTEIRAS_RODAR, N' | corte por ', @DATA_CORTE_MODO, N' (arquivo ', @CART_CORTE, N') = ', CONVERT(VARCHAR(10), @DATA_CORTE, 103),
    N' | último pagamento no arquivo: ', ISNULL(CONVERT(VARCHAR(10), @MAX_PAGAMENTO, 103), N'(não encontrado)'),
    N' | D-1: ', CONVERT(VARCHAR(10), @D_MENOS_1, 103), NCHAR(10),
    N'ANTERIOR: ', CONVERT(VARCHAR(10), @INI_ANT, 103), N' a ', CONVERT(VARCHAR(10), @FIM_ANT, 103),
    CASE WHEN @PERIODOS_RODAR IN ('AMBOS', 'ANTERIOR') THEN N'  -> será processado' ELSE N'  (não será processado)' END, NCHAR(10),
    N'ATUAL   : ', CONVERT(VARCHAR(10), @INI_ATUAL, 103), N' a ', CONVERT(VARCHAR(10), @FIM_ATUAL, 103),
    CASE WHEN @PERIODOS_RODAR IN ('AMBOS', 'ATUAL') THEN N'  -> será processado' ELSE N'  (não será processado)' END, NCHAR(10),
    N'=========================================================');
RAISERROR(@MSG, 0, 1) WITH NOWAIT;

IF @APENAS_MOSTRAR_PERIODOS = 1
BEGIN
    SELECT P.REFERENCIA, P.INICIO, P.FIM, P.DIAS_UTEIS,
           M.NOME + N'/' + CAST(YEAR(P.INICIO) AS NVARCHAR(4))                                          AS MES,
           N'FOTO_MATCHING_' + M.NOME + N'_' + CAST(YEAR(P.INICIO) AS NVARCHAR(4))                      AS TABELA_MATCHING,
           N'FOTO_MATCHING_EMPRESTIMO_' + M.NOME + N'_' + CAST(YEAR(P.INICIO) AS NVARCHAR(4))           AS TABELA_EMPRESTIMO,
           N'FOTO_MATCHING_' + MA.NOME + N'_' + CAST(YEAR(DATEADD(MONTH, -1, P.INICIO)) AS NVARCHAR(4)) AS TABELA_BUCKET_FALLBACK
    FROM @PERIODOS P
    JOIN @MESES M  ON M.NUM  = MONTH(P.INICIO)
    JOIN @MESES MA ON MA.NUM = MONTH(DATEADD(MONTH, -1, P.INICIO))
    ORDER BY P.ORDEM;
    RETURN;
END;

/* =====================================================================================
   3) CARGA ÚNICA — o que não depende do período (marcadores)
      Uma linha por cliente, com índice único: viram LEFT JOIN simples (sem EXISTS por linha).
   ===================================================================================== */
SET @T_SECAO = SYSDATETIME();

IF OBJECT_ID('tempdb..#DESENROLA') IS NOT NULL DROP TABLE #DESENROLA;
SELECT DISTINCT B.MCI
INTO #DESENROLA
FROM [COBwebNUBANK].[dbo].[MARCADOR_MASTER] A WITH (NOLOCK)
LEFT JOIN [COBwebNUBANK].[dbo].[DIV_DEVEDORES] B WITH (NOLOCK) ON B.ID_DEVEDORES = A.ID_DEVEDORES
WHERE A.ID_MARCADOR = @MARCADOR_DESENROLA
  AND CAST(A.DATA AS DATE) >= @MARCADORES_DESDE
  AND B.MCI IS NOT NULL
OPTION (RECOMPILE);
CREATE UNIQUE CLUSTERED INDEX IX ON #DESENROLA (MCI);

IF OBJECT_ID('tempdb..#CAMPANHA_INTERNA') IS NOT NULL DROP TABLE #CAMPANHA_INTERNA;
SELECT DISTINCT B.MCI
INTO #CAMPANHA_INTERNA
FROM [COBwebNUBANK].[dbo].[MARCADOR_MASTER] A WITH (NOLOCK)
LEFT JOIN [COBwebNUBANK].[dbo].[DIV_DEVEDORES] B WITH (NOLOCK) ON B.ID_DEVEDORES = A.ID_DEVEDORES
WHERE A.ID_MARCADOR = @MARCADOR_CAMPANHA_INTERNA
  AND CAST(A.DATA AS DATE) >= @MARCADORES_DESDE
  AND B.MCI IS NOT NULL
OPTION (RECOMPILE);
CREATE UNIQUE CLUSTERED INDEX IX ON #CAMPANHA_INTERNA (MCI);

SET @MSG = CONCAT(N'Carga única (marcadores): ', DATEDIFF(SECOND, @T_SECAO, SYSDATETIME()), N' s');
RAISERROR(@MSG, 0, 1) WITH NOWAIT;

/* =====================================================================================
   4) FUNIL POR PERÍODO (ANTERIOR e/ou ATUAL, com os mesmos dias úteis) E POR CARTEIRA (CC e/ou LL)
   ===================================================================================== */
DECLARE @ORDEM_PERIODO       INT = 0;
DECLARE @REFERENCIA_PERIODO  VARCHAR(10);
DECLARE @DIAS_UTEIS_PERIODO  INT;
DECLARE @DATA_INICIO         DATE;
DECLARE @DATA_FIM            DATE;
DECLARE @MES_NOME            NVARCHAR(20);
DECLARE @MES_NOME_ANT        NVARCHAR(20);
DECLARE @ANO                 NVARCHAR(4);
DECLARE @ANO_ANT             NVARCHAR(4);
DECLARE @MES_REFERENCIA      NVARCHAR(30);
DECLARE @TABELA_MATCHING     NVARCHAR(200);
DECLARE @TABELA_EMPRESTIMO   NVARCHAR(200);
DECLARE @TABELA_MATCHING_ANT NVARCHAR(200);
DECLARE @T_PERIODO           DATETIME2;
DECLARE @DUPLICADOS          BIGINT;
DECLARE @EXCLUIDOS           BIGINT;
-- carteiras
DECLARE @CART_ORD            INT;
DECLARE @CART                VARCHAR(2);
DECLARE @CART_ID             INT;
DECLARE @REF_CACHE           VARCHAR(10);
DECLARE @TABELA_FONTE        NVARCHAR(200);
DECLARE @T_CART              DATETIME2;
DECLARE @FORA_PADRAO         NVARCHAR(1000);
DECLARE @LINHAS_INS          BIGINT;
DECLARE @ESC_LOANS           BIGINT;
DECLARE @ESC_MULTI           BIGINT;
-- discagens
DECLARE @DATA_ARQ            DATE;
DECLARE @NOME_ARQ            NVARCHAR(80);
DECLARE @MES_PASTA           NVARCHAR(20);
DECLARE @CAMINHO             NVARCHAR(600);
DECLARE @TENT                INT;
DECLARE @CARREGOU            INT;
DECLARE @ERRO_NUM            INT;
DECLARE @ERRO_MSG            NVARCHAR(400);
DECLARE @QTD_PRINCIPAL       INT;
DECLARE @QTD_ARQUIVADO       INT;
DECLARE @QTD_AUSENTE         INT;

IF OBJECT_ID('tempdb..#FUNIL_LONGO_TOTAL') IS NOT NULL DROP TABLE #FUNIL_LONGO_TOTAL;
CREATE TABLE #FUNIL_LONGO_TOTAL (
    REFERENCIA        VARCHAR(10),
    PERIODO_INICIO    DATE,
    PERIODO_FIM       DATE,
    DIAS_UTEIS        INT,
    MES               NVARCHAR(30),
    FAIXA_DE_VALOR    VARCHAR(20),
    BUCKET_ATUALIZADO VARCHAR(50),
    TIPO              NVARCHAR(10),
    QTD               BIGINT,
    BALANCE           DECIMAL(38,9),
    CASH              DECIMAL(38,2),
    VERSAO_BASE       VARCHAR(20),
    CARTEIRA          VARCHAR(2),
    CAMPANHA_INTERNA  INT,
    DESENROLA         INT,
    MDR               INT,
    FUNIL             VARCHAR(2)
);

-- ---------------------------------------------------------------------------------------------
-- CACHE INCREMENTAL (@USAR_CACHE = 1)
--   Guarda o ESTADO ACUMULADO do mês de cada período (ATUAL e ANTERIOR) e carteira — não um registro por dia.
--   Chaves: CC = ATUAL / ANTERIOR (as mesmas da V6: o cache já montado continua valendo); LL = L_ATUAL / L_ANTERIOR.
--   O tamanho é FIXO (no máximo 2 meses) e não cresce com o tempo: na virada do mês o estado do
--   período é substituído. Cada rodada lê só o(s) dia(s) novo(s) e soma ao que já existe.
--   Tabelas (criadas sozinhas no banco onde a consulta roda):
--     dbo.FUNIL_CACHE_CTRL, dbo.FUNIL_CACHE_MATCHING, dbo.FUNIL_CACHE_DISC, dbo.FUNIL_CACHE_DISC_ARQ
--   Para apagar tudo: DROP TABLE nas quatro (a próxima rodada recria e remonta).
-- ---------------------------------------------------------------------------------------------
DECLARE @CACHE_ATIVO   BIT = 0;
DECLARE @LOCK_RES      INT;
DECLARE @COLLATE       NVARCHAR(128) = CONVERT(NVARCHAR(128), DATABASEPROPERTYEX('tempdb', 'Collation'));
DECLARE @MODO_CACHE    VARCHAR(12);
DECLARE @C_MES         DATE;
DECLARE @C_ATE         DATE;
DECLARE @DE_LEITURA    DATE;
DECLARE @SQL_FOTO      NVARCHAR(MAX);
DECLARE @JAN_ANT       DATE;
DECLARE @JA            DATE;
DECLARE @PULAR         BIT;
DECLARE @QTD_PULADO    INT;

IF @USAR_CACHE = 1
BEGIN
    BEGIN TRY
        -- as colunas de texto usam a MESMA collation do tempdb: assim o cache se junta às tabelas temporárias sem conflito
        SET @SQL = N'
        IF OBJECT_ID(''dbo.FUNIL_CACHE_CTRL'') IS NULL
            CREATE TABLE dbo.FUNIL_CACHE_CTRL (
                CACHE VARCHAR(12) COLLATE ' + @COLLATE + N' NOT NULL,
                REFERENCIA VARCHAR(10) COLLATE ' + @COLLATE + N' NOT NULL,
                MES_INI DATE NOT NULL, ATE_DATA DATE NOT NULL, ATUALIZADO_EM DATETIME2(0) NOT NULL,
                CONSTRAINT PK_FUNIL_CACHE_CTRL PRIMARY KEY (CACHE, REFERENCIA));
        IF OBJECT_ID(''dbo.FUNIL_CACHE_MATCHING'') IS NULL
            CREATE TABLE dbo.FUNIL_CACHE_MATCHING (
                REFERENCIA VARCHAR(10) COLLATE ' + @COLLATE + N' NOT NULL,
                CUSTOMER__ID VARCHAR(36) COLLATE ' + @COLLATE + N' NOT NULL,
                collection__id VARCHAR(36) COLLATE ' + @COLLATE + N' NOT NULL,
                EM_ATRASO TINYINT NOT NULL,
                balance DECIMAL(18,9) NULL,
                CHAVE_BUCKET VARCHAR(100) COLLATE ' + @COLLATE + N' NULL,
                CONSTRAINT PK_FUNIL_CACHE_MATCHING PRIMARY KEY (REFERENCIA, CUSTOMER__ID, collection__id));
        IF OBJECT_ID(''dbo.FUNIL_CACHE_DISC'') IS NULL
            CREATE TABLE dbo.FUNIL_CACHE_DISC (
                REFERENCIA VARCHAR(10) COLLATE ' + @COLLATE + N' NOT NULL,
                customer__id VARCHAR(36) COLLATE ' + @COLLATE + N' NOT NULL,
                QTD_DISCADO INT NOT NULL, DISCADO INT NOT NULL, ATENDIDO INT NOT NULL, ALO INT NOT NULL, CPC INT NOT NULL,
                CONSTRAINT PK_FUNIL_CACHE_DISC PRIMARY KEY (REFERENCIA, customer__id));
        IF OBJECT_ID(''dbo.FUNIL_CACHE_DISC_ARQ'') IS NULL
            CREATE TABLE dbo.FUNIL_CACHE_DISC_ARQ (
                REFERENCIA VARCHAR(10) COLLATE ' + @COLLATE + N' NOT NULL,
                DATA_ARQUIVO DATE NOT NULL, JANELA_ATE DATE NOT NULL,
                CONSTRAINT PK_FUNIL_CACHE_DISC_ARQ PRIMARY KEY (REFERENCIA, DATA_ARQUIVO));';
        EXEC sp_executesql @SQL;

        -- trava: duas execuções ao mesmo tempo somariam o mesmo dia duas vezes
        EXEC @LOCK_RES = sp_getapplock @Resource = N'FUNIL_NUBANK_CACHE', @LockMode = 'Exclusive',
                                       @LockOwner = 'Session', @LockTimeout = 0;
        IF @LOCK_RES >= 0
            SET @CACHE_ATIVO = 1;
        ELSE
            RAISERROR(N'AVISO: o cache está em uso por outra execução; esta rodada lê tudo direto (sem cache).', 0, 1) WITH NOWAIT;
    END TRY
    BEGIN CATCH
        SET @MSG = REPLACE(CONCAT(N'AVISO: cache indisponível (', ERROR_MESSAGE(), N'). Esta rodada lê tudo direto.'), N'%', N'%%');
        RAISERROR(@MSG, 0, 1) WITH NOWAIT;
    END CATCH;
END;
SET @MSG = CASE WHEN @CACHE_ATIVO = 1 THEN N'Cache incremental: ATIVO' ELSE N'Cache incremental: desligado (leitura direta, igual à V3)' END;
RAISERROR(@MSG, 0, 1) WITH NOWAIT;

WHILE 1 = 1
BEGIN
    SELECT TOP (1) @ORDEM_PERIODO = ORDEM, @REFERENCIA_PERIODO = REFERENCIA,
                   @DATA_INICIO = INICIO, @DATA_FIM = FIM, @DIAS_UTEIS_PERIODO = DIAS_UTEIS
    FROM @PERIODOS WHERE ORDEM > @ORDEM_PERIODO ORDER BY ORDEM;
    IF @@ROWCOUNT = 0 BREAK;

    SET @MES_NOME       = (SELECT NOME FROM @MESES WHERE NUM = MONTH(@DATA_INICIO));
    SET @ANO            = CAST(YEAR(@DATA_INICIO) AS NVARCHAR(4));
    SET @MES_NOME_ANT   = (SELECT NOME FROM @MESES WHERE NUM = MONTH(DATEADD(MONTH, -1, @DATA_INICIO)));
    SET @ANO_ANT        = CAST(YEAR(DATEADD(MONTH, -1, @DATA_INICIO)) AS NVARCHAR(4));
    SET @MES_REFERENCIA = @MES_NOME + N'/' + @ANO;   -- igual à aba "Dados": Setembro/2026

    -- Nomes das tabelas (Março: tenta com cedilha e, se não existir, sem)
    SET @TABELA_MATCHING = N'FOTO_MATCHING_' + @MES_NOME + N'_' + @ANO;
    IF OBJECT_ID(@TABELA_MATCHING) IS NULL
        SET @TABELA_MATCHING = N'FOTO_MATCHING_' + REPLACE(@MES_NOME, N'ç', N'c') + N'_' + @ANO;
    SET @TABELA_EMPRESTIMO = N'FOTO_MATCHING_EMPRESTIMO_' + @MES_NOME + N'_' + @ANO;
    IF OBJECT_ID(@TABELA_EMPRESTIMO) IS NULL
        SET @TABELA_EMPRESTIMO = N'FOTO_MATCHING_EMPRESTIMO_' + REPLACE(@MES_NOME, N'ç', N'c') + N'_' + @ANO;
    SET @TABELA_MATCHING_ANT = N'FOTO_MATCHING_' + @MES_NOME_ANT + N'_' + @ANO_ANT;
    IF OBJECT_ID(@TABELA_MATCHING_ANT) IS NULL
        SET @TABELA_MATCHING_ANT = N'FOTO_MATCHING_' + REPLACE(@MES_NOME_ANT, N'ç', N'c') + N'_' + @ANO_ANT;

    SET @T_PERIODO = SYSDATETIME();
    SET @T_SECAO   = SYSDATETIME();
    SET @MSG = CONCAT(NCHAR(10), N'##### PERÍODO ', @REFERENCIA_PERIODO, N' (', @MES_REFERENCIA, N'): ',
                      CONVERT(VARCHAR(10), @DATA_INICIO, 103), N' a ', CONVERT(VARCHAR(10), @DATA_FIM, 103),
                      N' | ', @DIAS_UTEIS_PERIODO, N' d.u. #####');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;

    IF OBJECT_ID(@TABELA_EMPRESTIMO) IS NULL
       OR (OBJECT_ID(@TABELA_MATCHING) IS NULL AND EXISTS (SELECT 1 FROM @CARTEIRAS WHERE CARTEIRA = 'CC'))
    BEGIN
        SET @MSG = CONCAT(N'ERRO: não existe ', @TABELA_MATCHING, N' e/ou ', @TABELA_EMPRESTIMO, N' — período PULADO.');
        RAISERROR(@MSG, 0, 1) WITH NOWAIT;
        CONTINUE;
    END;

    /* =================================================================================
       PARTE COMUM DO PERÍODO (lida UMA vez e usada pelas duas carteiras)
       ================================================================================= */
    -- Clientes que aparecem na tabela de empréstimo no período (base do MDR nas duas consultas)
    IF OBJECT_ID('tempdb..#EMP_CLI') IS NOT NULL DROP TABLE #EMP_CLI;
    CREATE TABLE #EMP_CLI (customer__id VARCHAR(36) NOT NULL);
    SET @SQL = N'
    INSERT INTO #EMP_CLI WITH (TABLOCK) (customer__id)
    SELECT DISTINCT CAST(A.customer__id AS VARCHAR(36))
    FROM ' + QUOTENAME(@TABELA_EMPRESTIMO) + N' A WITH (NOLOCK)
    WHERE A.date BETWEEN @P_INICIO AND @P_FIM
      AND A.customer__id IS NOT NULL
    OPTION (RECOMPILE);';
    EXEC sp_executesql @SQL, N'@P_INICIO DATE, @P_FIM DATE', @P_INICIO = @DATA_INICIO, @P_FIM = @DATA_FIM;
    CREATE UNIQUE CLUSTERED INDEX IX ON #EMP_CLI (customer__id);

    -- Acordos com vencimento no período (REC_MASTER filtrada uma vez; as carteiras 5 e 6 estão juntas, como nas originais)
    IF OBJECT_ID('tempdb..#REC_PERIODO') IS NOT NULL DROP TABLE #REC_PERIODO;
    SELECT A.NR_RECEBIMENTO, A.ID_DEV, A.ID_USUARIO, A.ID_CARTEIRA, A.DT_VENCIMENTO, A.DT_EMISSAO, A.valor
    INTO #REC_PERIODO
    FROM [COBwebNUBANK].[dbo].[REC_MASTER] A WITH (NOLOCK)
    WHERE A.DT_VENCIMENTO BETWEEN @DATA_INICIO AND @DATA_FIM
      AND A.PARCELA = 0
      AND A.ID_USUARIO NOT IN (SELECT ID_USUARIO FROM #USUARIOS_EXCLUIDOS)
      AND A.ID_CARTEIRA IN (@CARTEIRA_CC, @CARTEIRA_LL)
    OPTION (RECOMPILE);
    CREATE CLUSTERED INDEX IX ON #REC_PERIODO (NR_RECEBIMENTO);

    SET @MSG = CONCAT(N'  Empréstimo (MDR) e acordos do período ', DATEDIFF(SECOND, @T_SECAO, SYSDATETIME()), N' s');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;  SET @T_SECAO = SYSDATETIME();

    /* ---------------------------------------------------------------------------------
       4.6) DISCAGENS — um CSV por dia. Cada arquivo é resumido por cliente ASSIM que carrega
            (só os dias do período), então o tempdb nunca guarda as linhas cruas de todos os dias.
            Procura: 1) pasta principal  2) <pasta>\<ANO>\<Mês>\  3) idem sem cedilha (Março)
       --------------------------------------------------------------------------------- */
    -- Modo do cache de discagens (mesma regra do matching). O estado guarda, por cliente, a soma das
    -- ligações e os indicadores já apurados; uma tabela de controle guarda, POR ARQUIVO, até que data de
    -- ligação ele já foi consumido. Arquivo que ainda não chegou não é registrado, então é lido na próxima rodada.
    SELECT @MODO_CACHE = 'DIRETO', @C_MES = NULL, @C_ATE = NULL;
    IF @CACHE_ATIVO = 1
    BEGIN
        SELECT @C_MES = MES_INI, @C_ATE = ATE_DATA
        FROM dbo.FUNIL_CACHE_CTRL WHERE CACHE = 'DISCAGENS' AND REFERENCIA = @REFERENCIA_PERIODO;
        SET @MODO_CACHE = CASE WHEN @RECRIAR_CACHE = 1 OR @C_MES IS NULL OR @DATA_INICIO > @C_MES THEN 'RECRIAR'
                               WHEN @DATA_INICIO = @C_MES AND @DATA_FIM >= @C_ATE                 THEN 'INCREMENTAL'
                               ELSE 'DIRETO' END;
    END;

    IF OBJECT_ID('tempdb..#DISC_STG') IS NOT NULL DROP TABLE #DISC_STG;
    CREATE TABLE #DISC_STG (
        [date] DATE, [time] TIME, customer__id VARCHAR(36), telefone VARCHAR(20),
        discado INT, atendido INT, alo INT, cpc INT, author VARCHAR(50),
        call_duration INT, post_call_duration INT, tab_reason VARCHAR(200)
    );
    IF OBJECT_ID('tempdb..#DISC_DIA') IS NOT NULL DROP TABLE #DISC_DIA;
    CREATE TABLE #DISC_DIA (customer__id VARCHAR(36), QTD_DISCADO INT, DISCADO INT, ATENDIDO INT, ALO INT, CPC INT);
    IF OBJECT_ID('tempdb..#ARQ_NOVOS') IS NOT NULL DROP TABLE #ARQ_NOVOS;
    CREATE TABLE #ARQ_NOVOS (DATA_ARQUIVO DATE NOT NULL);

    SELECT @DATA_ARQ = @DATA_INICIO, @QTD_PRINCIPAL = 0, @QTD_ARQUIVADO = 0, @QTD_AUSENTE = 0, @QTD_PULADO = 0;
    WHILE @DATA_ARQ <= DATEADD(DAY, @DIAS_FOLGA_DISCAGENS, @DATA_FIM)
    BEGIN
        -- janela de ligações a somar deste arquivo: (@JAN_ANT, @DATA_FIM]
        SELECT @PULAR = 0, @JAN_ANT = DATEADD(DAY, -1, @DATA_INICIO);
        IF @MODO_CACHE = 'INCREMENTAL'
        BEGIN
            SET @JA = NULL;
            SELECT @JA = JANELA_ATE FROM dbo.FUNIL_CACHE_DISC_ARQ
            WHERE REFERENCIA = @REFERENCIA_PERIODO AND DATA_ARQUIVO = @DATA_ARQ;
            IF @JA IS NOT NULL
            BEGIN
                -- um arquivo só traz ligações até a sua própria data: se já foi consumido até lá (ou até o fim), pula
                IF @JA >= CASE WHEN @DATA_ARQ < @DATA_FIM THEN @DATA_ARQ ELSE @DATA_FIM END
                    SET @PULAR = 1;
                ELSE
                    SET @JAN_ANT = @JA;   -- já foi lido antes: soma só as ligações NOVAS (depois da janela anterior)
            END;
        END;

        IF @PULAR = 1
            SET @QTD_PULADO = @QTD_PULADO + 1;
        ELSE
        BEGIN
            SET @NOME_ARQ  = @PREFIXO_DISCAGENS + FORMAT(@DATA_ARQ, 'yyyyMMdd') + N'.csv';
            SET @MES_PASTA = (SELECT NOME FROM @MESES WHERE NUM = MONTH(@DATA_ARQ));
            SELECT @CARREGOU = 0, @ERRO_NUM = NULL, @ERRO_MSG = NULL, @TENT = 1;
            WHILE @TENT <= 3 AND @CARREGOU = 0
            BEGIN
                SET @CAMINHO = CASE @TENT
                    WHEN 1 THEN @PASTA_DISCAGENS + @NOME_ARQ
                    WHEN 2 THEN @PASTA_DISCAGENS + CAST(YEAR(@DATA_ARQ) AS NVARCHAR(4)) + N'\' + @MES_PASTA + N'\' + @NOME_ARQ
                    WHEN 3 THEN CASE WHEN @MES_PASTA LIKE N'%ç%'
                                     THEN @PASTA_DISCAGENS + CAST(YEAR(@DATA_ARQ) AS NVARCHAR(4)) + N'\'
                                          + REPLACE(@MES_PASTA, N'ç', N'c') + N'\' + @NOME_ARQ END
                END;
                IF @CAMINHO IS NOT NULL
                BEGIN
                    TRUNCATE TABLE #DISC_STG;
                    SET @SQL = N'BULK INSERT #DISC_STG FROM ''' + REPLACE(@CAMINHO, N'''', N'''''') + N'''
                        WITH (FIRSTROW = 2, FIELDTERMINATOR = '';'', ROWTERMINATOR = ''0x0a'', CODEPAGE = ''65001'', TABLOCK);';
                    BEGIN TRY
                        EXEC sp_executesql @SQL;
                        SET @CARREGOU = @TENT;
                    END TRY
                    BEGIN CATCH
                        SELECT @ERRO_NUM = ERROR_NUMBER(), @ERRO_MSG = LEFT(ERROR_MESSAGE(), 300);
                    END CATCH;
                END;
                SET @TENT = @TENT + 1;
            END;

            IF @CARREGOU > 0
            BEGIN
                -- resume o arquivo (só as ligações da janela) e descarta as linhas cruas
                INSERT INTO #DISC_DIA (customer__id, QTD_DISCADO, DISCADO, ATENDIDO, ALO, CPC)
                SELECT customer__id, COUNT(*), MAX(discado), MAX(atendido), MAX(alo), MAX(cpc)
                FROM #DISC_STG
                WHERE [date] > @JAN_ANT AND [date] <= @DATA_FIM
                GROUP BY customer__id;
                INSERT INTO #ARQ_NOVOS (DATA_ARQUIVO) VALUES (@DATA_ARQ);
                IF @CARREGOU = 1 SET @QTD_PRINCIPAL = @QTD_PRINCIPAL + 1; ELSE SET @QTD_ARQUIVADO = @QTD_ARQUIVADO + 1;
            END
            ELSE
            BEGIN
                SET @QTD_AUSENTE = @QTD_AUSENTE + 1;
                -- 4860 = arquivo não existe (normal em fim de semana). Outro erro é mostrado.
                IF @ERRO_NUM IS NOT NULL AND @ERRO_NUM <> 4860
                BEGIN
                    SET @MSG = REPLACE(CONCAT(N'  AVISO discagem ', @NOME_ARQ, N': erro ', @ERRO_NUM, N' - ', @ERRO_MSG), N'%', N'%%');
                    RAISERROR(@MSG, 0, 1) WITH NOWAIT;
                END;
            END;
        END;
        SET @DATA_ARQ = DATEADD(DAY, 1, @DATA_ARQ);
    END;
    DROP TABLE #DISC_STG;

    -- cliente nulo nunca se junta a nada na base; fora do resultado e do cache (nulo no flag = 0, como antes)
    IF OBJECT_ID('tempdb..#DISC_NOVO') IS NOT NULL DROP TABLE #DISC_NOVO;
    SELECT customer__id,
           SUM(QTD_DISCADO)         AS QTD_DISCADO,
           ISNULL(MAX(DISCADO), 0)  AS DISCADO,
           ISNULL(MAX(ATENDIDO), 0) AS ATENDIDO,
           ISNULL(MAX(ALO), 0)      AS ALO,
           ISNULL(MAX(CPC), 0)      AS CPC
    INTO #DISC_NOVO
    FROM #DISC_DIA
    WHERE customer__id IS NOT NULL
    GROUP BY customer__id;
    DROP TABLE #DISC_DIA;

    IF OBJECT_ID('tempdb..#DISCAGENS') IS NOT NULL DROP TABLE #DISCAGENS;
    CREATE TABLE #DISCAGENS (customer__id VARCHAR(36) NOT NULL, QTD_DISCADO INT, DISCADO INT, ATENDIDO INT, ALO INT, CPC INT);

    IF @MODO_CACHE IN ('RECRIAR', 'INCREMENTAL')
    BEGIN
        BEGIN TRY
            BEGIN TRAN;
                IF @MODO_CACHE = 'RECRIAR'
                BEGIN
                    DELETE FROM dbo.FUNIL_CACHE_DISC     WHERE REFERENCIA = @REFERENCIA_PERIODO;
                    DELETE FROM dbo.FUNIL_CACHE_DISC_ARQ WHERE REFERENCIA = @REFERENCIA_PERIODO;
                    INSERT INTO dbo.FUNIL_CACHE_DISC (REFERENCIA, customer__id, QTD_DISCADO, DISCADO, ATENDIDO, ALO, CPC)
                    SELECT @REFERENCIA_PERIODO, customer__id, QTD_DISCADO, DISCADO, ATENDIDO, ALO, CPC FROM #DISC_NOVO;
                END
                ELSE
                BEGIN
                    -- ligações novas: quantidade SOMA; discado/atendido/alô/CPC = MAX
                    UPDATE C SET
                        QTD_DISCADO = C.QTD_DISCADO + N.QTD_DISCADO,
                        DISCADO     = CASE WHEN N.DISCADO  > C.DISCADO  THEN N.DISCADO  ELSE C.DISCADO  END,
                        ATENDIDO    = CASE WHEN N.ATENDIDO > C.ATENDIDO THEN N.ATENDIDO ELSE C.ATENDIDO END,
                        ALO         = CASE WHEN N.ALO      > C.ALO      THEN N.ALO      ELSE C.ALO      END,
                        CPC         = CASE WHEN N.CPC      > C.CPC      THEN N.CPC      ELSE C.CPC      END
                    FROM dbo.FUNIL_CACHE_DISC C
                    INNER JOIN #DISC_NOVO N ON N.customer__id = C.customer__id
                    WHERE C.REFERENCIA = @REFERENCIA_PERIODO;

                    INSERT INTO dbo.FUNIL_CACHE_DISC (REFERENCIA, customer__id, QTD_DISCADO, DISCADO, ATENDIDO, ALO, CPC)
                    SELECT @REFERENCIA_PERIODO, N.customer__id, N.QTD_DISCADO, N.DISCADO, N.ATENDIDO, N.ALO, N.CPC
                    FROM #DISC_NOVO N
                    WHERE NOT EXISTS (SELECT 1 FROM dbo.FUNIL_CACHE_DISC C
                                      WHERE C.REFERENCIA = @REFERENCIA_PERIODO AND C.customer__id = N.customer__id);
                END;

                DELETE A FROM dbo.FUNIL_CACHE_DISC_ARQ A
                INNER JOIN #ARQ_NOVOS N ON N.DATA_ARQUIVO = A.DATA_ARQUIVO
                WHERE A.REFERENCIA = @REFERENCIA_PERIODO;
                INSERT INTO dbo.FUNIL_CACHE_DISC_ARQ (REFERENCIA, DATA_ARQUIVO, JANELA_ATE)
                SELECT @REFERENCIA_PERIODO, DATA_ARQUIVO, @DATA_FIM FROM #ARQ_NOVOS;

                UPDATE dbo.FUNIL_CACHE_CTRL
                SET MES_INI = @DATA_INICIO, ATE_DATA = @DATA_FIM, ATUALIZADO_EM = SYSDATETIME()
                WHERE CACHE = 'DISCAGENS' AND REFERENCIA = @REFERENCIA_PERIODO;
                IF @@ROWCOUNT = 0
                    INSERT INTO dbo.FUNIL_CACHE_CTRL (CACHE, REFERENCIA, MES_INI, ATE_DATA, ATUALIZADO_EM)
                    VALUES ('DISCAGENS', @REFERENCIA_PERIODO, @DATA_INICIO, @DATA_FIM, SYSDATETIME());
            COMMIT;

            INSERT INTO #DISCAGENS (customer__id, QTD_DISCADO, DISCADO, ATENDIDO, ALO, CPC)
            SELECT customer__id, QTD_DISCADO, DISCADO, ATENDIDO, ALO, CPC
            FROM dbo.FUNIL_CACHE_DISC WHERE REFERENCIA = @REFERENCIA_PERIODO;
            SET @LINHAS = @@ROWCOUNT;
        END TRY
        BEGIN CATCH
            IF @@TRANCOUNT > 0 ROLLBACK;
            SET @MSG = REPLACE(CONCAT(N'  ERRO no cache de discagens: ', ERROR_MESSAGE(),
                                      N' | Corrija ou rode com @USAR_CACHE = 0.'), N'%', N'%%');
            RAISERROR(@MSG, 0, 1) WITH NOWAIT;
            THROW;
        END CATCH;
    END
    ELSE
    BEGIN
        INSERT INTO #DISCAGENS (customer__id, QTD_DISCADO, DISCADO, ATENDIDO, ALO, CPC)
        SELECT customer__id, QTD_DISCADO, DISCADO, ATENDIDO, ALO, CPC FROM #DISC_NOVO;
        SET @LINHAS = @@ROWCOUNT;
    END;
    CREATE UNIQUE CLUSTERED INDEX IX ON #DISCAGENS (customer__id);
    DROP TABLE #DISC_NOVO;
    DROP TABLE #ARQ_NOVOS;

    SET @MSG = CONCAT(N'  4.6 Discagens (CC e LL) .......... ', DATEDIFF(SECOND, @T_SECAO, SYSDATETIME()), N' s | cache: ', @MODO_CACHE,
                      N' | arquivos: ', @QTD_PRINCIPAL, N' lidos na pasta principal, ', @QTD_ARQUIVADO, N' na do mês arquivado, ',
                      @QTD_PULADO, N' já no cache, ', @QTD_AUSENTE, N' não encontrados | ', @LINHAS, N' clientes discados');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;  SET @T_SECAO = SYSDATETIME();


    /* =================================================================================
       FUNIL POR CARTEIRA
       ================================================================================= */
    SET @CART_ORD = 0;
    WHILE 1 = 1
    BEGIN
    SELECT TOP (1) @CART_ORD = ORDEM, @CART = CARTEIRA, @CART_ID = ID
    FROM @CARTEIRAS WHERE ORDEM > @CART_ORD ORDER BY ORDEM;
    IF @@ROWCOUNT = 0 BREAK;

    SET @REF_CACHE    = CASE WHEN @CART = 'CC' THEN @REFERENCIA_PERIODO ELSE 'L_' + @REFERENCIA_PERIODO END;
    SET @TABELA_FONTE = CASE WHEN @CART = 'CC' THEN @TABELA_MATCHING ELSE @TABELA_EMPRESTIMO END;
    SET @T_CART = SYSDATETIME();
    SET @T_SECAO = SYSDATETIME();
    SET @MSG = CONCAT(N'  ------ ', @CART, N' (carteira ', @CART_ID, N') — foto ', @TABELA_FONTE, N' ------');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;

    /* ---------------------------------------------------------------------------------
       4.1) MATCHING — UMA leitura da foto (a parte pesada: dezenas de milhões de linhas).
            Resultado por cliente+contrato:
              • em atraso = teve ao menos um dia com FINANCE_DAYS_LATE <> 0;
              • balance   = maior saldo entre os dias em atraso;
              • bucket    = o da foto mais antiga do CLIENTE que tem bucket preenchido
                            (considera todas as linhas do cliente, como na original).
            O bucket vem de um MIN("data|bucket") em texto ISO, que ordena igual à data.
       --------------------------------------------------------------------------------- */
    -- Modo do cache neste período:
    --   RECRIAR     = não há estado deste mês (1ª vez, virada do mês ou @RECRIAR_CACHE = 1): lê o mês todo e grava
    --   INCREMENTAL = há estado do MESMO mês até @C_ATE: lê só de @C_ATE em diante e soma ao estado
    --   DIRETO      = sem cache (desligado, ou pedido anterior ao que já está guardado: o cache NÃO é alterado)
    SELECT @MODO_CACHE = 'DIRETO', @DE_LEITURA = @DATA_INICIO, @C_MES = NULL, @C_ATE = NULL;
    IF @CACHE_ATIVO = 1
    BEGIN
        SELECT @C_MES = MES_INI, @C_ATE = ATE_DATA
        FROM dbo.FUNIL_CACHE_CTRL WHERE CACHE = 'MATCHING' AND REFERENCIA = @REF_CACHE;
        SET @MODO_CACHE = CASE WHEN @RECRIAR_CACHE = 1 OR @C_MES IS NULL OR @DATA_INICIO > @C_MES THEN 'RECRIAR'
                               WHEN @DATA_INICIO = @C_MES AND @DATA_FIM >= @C_ATE                 THEN 'INCREMENTAL'
                               ELSE 'DIRETO' END;
        -- recomeça no último dia guardado: re-ler é seguro porque o estado só usa MAX/MIN (idempotentes)
        IF @MODO_CACHE = 'INCREMENTAL' SET @DE_LEITURA = @C_ATE;
    END;

    IF OBJECT_ID('tempdb..#GRUPO') IS NOT NULL DROP TABLE #GRUPO;
    CREATE TABLE #GRUPO (
        CUSTOMER__ID   VARCHAR(36),
        collection__id VARCHAR(36),
        EM_ATRASO      TINYINT,
        balance        DECIMAL(18,9),
        CHAVE_BUCKET   VARCHAR(100)
    );
    IF @CART = 'CC'
    SET @SQL_FOTO = N'
        INSERT INTO #GRUPO WITH (TABLOCK) (CUSTOMER__ID, collection__id, EM_ATRASO, balance, CHAVE_BUCKET)
        SELECT CAST(CUSTOMER__ID   AS VARCHAR(36)),
               CAST(collection__id AS VARCHAR(36)),
               MAX(CASE WHEN FINANCE_DAYS_LATE <> 0 THEN 1 ELSE 0 END),
               MAX(CASE WHEN FINANCE_DAYS_LATE <> 0 THEN TRY_CAST(OUTSTANDING_BALANCE AS DECIMAL(18,9)) END),
               MIN(CASE WHEN bucket_atualizado IS NOT NULL
                        THEN CONVERT(VARCHAR(30), DATA_MATCHING, 126) + ''|'' + CAST(bucket_atualizado AS VARCHAR(50)) END)
        FROM ' + QUOTENAME(@TABELA_FONTE) + N' WITH (NOLOCK)
        WHERE reference_date BETWEEN @P_INICIO AND @P_FIM
        GROUP BY CAST(CUSTOMER__ID AS VARCHAR(36)), CAST(collection__id AS VARCHAR(36))
        OPTION (RECOMPILE);';
    ELSE
        -- LL (regra da CONSULTA_LL): filtra por DATA_MATCHING; saldo = soma dos loans do contrato no dia (só os em
        -- atraso) e depois o pico do período; bucket = o primeiro preenchido por CLIENTE + CONTRATO, na ordem
        -- DATA_MATCHING crescente e loan_days_late decrescente. Tudo numa leitura só (chave ordenável no MIN).
        SET @SQL_FOTO = N'
    INSERT INTO #GRUPO WITH (TABLOCK) (CUSTOMER__ID, collection__id, EM_ATRASO, balance, CHAVE_BUCKET)
    SELECT CUSTOMER__ID, collection__id, MAX(EM_ATRASO_DIA), MAX(BAL_DIA), MIN(CHAVE_DIA)
    FROM (
        SELECT CAST(customer__id   AS VARCHAR(36)) AS CUSTOMER__ID,
               CAST(collection__id AS VARCHAR(36)) AS collection__id,
               TRY_CAST(DATA_MATCHING AS DATE)     AS DM,
               MAX(CASE WHEN TRY_CAST(loan_days_late AS INT) <> 0 THEN 1 ELSE 0 END) AS EM_ATRASO_DIA,
               SUM(CASE WHEN TRY_CAST(loan_days_late AS INT) <> 0
                        THEN TRY_CAST(outstanding_balance AS DECIMAL(18,9)) END)   AS BAL_DIA,
               MIN(CASE WHEN bucket_atualizado IS NOT NULL
                        THEN CONVERT(VARCHAR(10), TRY_CAST(DATA_MATCHING AS DATE), 23) + ''|''
                             + RIGHT(''00000000000'' + CAST(CAST(1000000000 AS BIGINT)
                                     - ISNULL(TRY_CAST(loan_days_late AS BIGINT), CAST(-9000000000 AS BIGINT)) AS VARCHAR(20)), 11)
                             + ''|'' + CAST(bucket_atualizado AS VARCHAR(20)) END)     AS CHAVE_DIA
        FROM ' + QUOTENAME(@TABELA_FONTE) + N' WITH (NOLOCK)
        WHERE TRY_CAST(DATA_MATCHING AS DATE) BETWEEN @P_INICIO AND @P_FIM
        GROUP BY CAST(customer__id AS VARCHAR(36)), CAST(collection__id AS VARCHAR(36)), TRY_CAST(DATA_MATCHING AS DATE)
    ) D
    GROUP BY CUSTOMER__ID, collection__id
    OPTION (RECOMPILE);';
    EXEC sp_executesql @SQL_FOTO, N'@P_INICIO DATE, @P_FIM DATE', @P_INICIO = @DE_LEITURA, @P_FIM = @DATA_FIM;
    SET @LINHAS = @@ROWCOUNT;

    IF @MODO_CACHE IN ('RECRIAR', 'INCREMENTAL')
    BEGIN
        BEGIN TRY
            IF EXISTS (SELECT 1 FROM #GRUPO WHERE CUSTOMER__ID IS NULL OR collection__id IS NULL)
                RAISERROR('A foto tem cliente/contrato sem identificador; o cache nao e seguro.', 16, 1);

            BEGIN TRAN;
                IF @MODO_CACHE = 'RECRIAR'
                BEGIN
                    DELETE FROM dbo.FUNIL_CACHE_MATCHING WHERE REFERENCIA = @REF_CACHE;
                    INSERT INTO dbo.FUNIL_CACHE_MATCHING (REFERENCIA, CUSTOMER__ID, collection__id, EM_ATRASO, balance, CHAVE_BUCKET)
                    SELECT @REF_CACHE, CUSTOMER__ID, collection__id, EM_ATRASO, balance, CHAVE_BUCKET FROM #GRUPO;
                END
                ELSE
                BEGIN
                    -- soma o(s) dia(s) novo(s) ao estado: flag = MAX, saldo = MAX, bucket = MIN("data|bucket")
                    UPDATE C SET
                        EM_ATRASO    = CASE WHEN N.EM_ATRASO > C.EM_ATRASO THEN N.EM_ATRASO ELSE C.EM_ATRASO END,
                        balance      = CASE WHEN C.balance IS NULL THEN N.balance
                                            WHEN N.balance > C.balance THEN N.balance ELSE C.balance END,
                        CHAVE_BUCKET = CASE WHEN C.CHAVE_BUCKET IS NULL THEN N.CHAVE_BUCKET
                                            WHEN N.CHAVE_BUCKET < C.CHAVE_BUCKET THEN N.CHAVE_BUCKET ELSE C.CHAVE_BUCKET END
                    FROM dbo.FUNIL_CACHE_MATCHING C
                    INNER JOIN #GRUPO N ON N.CUSTOMER__ID = C.CUSTOMER__ID AND N.collection__id = C.collection__id
                    WHERE C.REFERENCIA = @REF_CACHE
                      AND (N.EM_ATRASO > C.EM_ATRASO
                           OR (N.balance IS NOT NULL AND (C.balance IS NULL OR N.balance > C.balance))
                           OR (N.CHAVE_BUCKET IS NOT NULL AND (C.CHAVE_BUCKET IS NULL OR N.CHAVE_BUCKET < C.CHAVE_BUCKET)));

                    INSERT INTO dbo.FUNIL_CACHE_MATCHING (REFERENCIA, CUSTOMER__ID, collection__id, EM_ATRASO, balance, CHAVE_BUCKET)
                    SELECT @REF_CACHE, N.CUSTOMER__ID, N.collection__id, N.EM_ATRASO, N.balance, N.CHAVE_BUCKET
                    FROM #GRUPO N
                    WHERE NOT EXISTS (SELECT 1 FROM dbo.FUNIL_CACHE_MATCHING C
                                      WHERE C.REFERENCIA = @REF_CACHE
                                        AND C.CUSTOMER__ID = N.CUSTOMER__ID AND C.collection__id = N.collection__id);
                END;

                UPDATE dbo.FUNIL_CACHE_CTRL
                SET MES_INI = @DATA_INICIO, ATE_DATA = @DATA_FIM, ATUALIZADO_EM = SYSDATETIME()
                WHERE CACHE = 'MATCHING' AND REFERENCIA = @REF_CACHE;
                IF @@ROWCOUNT = 0
                    INSERT INTO dbo.FUNIL_CACHE_CTRL (CACHE, REFERENCIA, MES_INI, ATE_DATA, ATUALIZADO_EM)
                    VALUES ('MATCHING', @REF_CACHE, @DATA_INICIO, @DATA_FIM, SYSDATETIME());
            COMMIT;

            -- daqui para frente o resto da consulta usa o estado completo do período, igual à leitura direta
            TRUNCATE TABLE #GRUPO;
            INSERT INTO #GRUPO WITH (TABLOCK) (CUSTOMER__ID, collection__id, EM_ATRASO, balance, CHAVE_BUCKET)
            SELECT CUSTOMER__ID, collection__id, EM_ATRASO, balance, CHAVE_BUCKET
            FROM dbo.FUNIL_CACHE_MATCHING WHERE REFERENCIA = @REF_CACHE;
            SET @LINHAS = @@ROWCOUNT;
        END TRY
        BEGIN CATCH
            IF @@TRANCOUNT > 0 ROLLBACK;
            SET @MSG = REPLACE(CONCAT(N'  ERRO no cache de matching: ', ERROR_MESSAGE(),
                                      N' | Corrija ou rode com @USAR_CACHE = 0.'), N'%', N'%%');
            RAISERROR(@MSG, 0, 1) WITH NOWAIT;
            THROW;
        END CATCH;
    END;
    SET @MSG = CONCAT(N'  ', @CART, N' 4.1a Leitura da foto ............. ', DATEDIFF(SECOND, @T_SECAO, SYSDATETIME()), N' s | cache: ', @MODO_CACHE,
                      CASE WHEN @MODO_CACHE = 'INCREMENTAL' THEN CONCAT(N' (lido de ', CONVERT(VARCHAR(10), @DE_LEITURA, 103), N' em diante)') ELSE N'' END,
                      N' | ', @LINHAS, N' grupos');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;  SET @T_SECAO = SYSDATETIME();

    CREATE CLUSTERED INDEX IX ON #GRUPO (CUSTOMER__ID);

    IF OBJECT_ID('tempdb..#MATCHING') IS NOT NULL DROP TABLE #MATCHING;
    CREATE TABLE #MATCHING (CUSTOMER__ID VARCHAR(36), collection__id VARCHAR(36), BUCKET_ATUALIZADO VARCHAR(50), balance DECIMAL(18,9));
    IF @CART = 'CC'
    BEGIN
        -- CC: bucket da foto mais antiga do CLIENTE (todas as linhas do cliente), como na consulta do CC
        IF OBJECT_ID('tempdb..#BUCKET_CLIENTE') IS NOT NULL DROP TABLE #BUCKET_CLIENTE;
        SELECT CUSTOMER__ID, MIN(CHAVE_BUCKET) AS CHAVE
        INTO #BUCKET_CLIENTE
        FROM #GRUPO
        WHERE CHAVE_BUCKET IS NOT NULL
        GROUP BY CUSTOMER__ID;
        CREATE UNIQUE CLUSTERED INDEX IX ON #BUCKET_CLIENTE (CUSTOMER__ID);
        INSERT INTO #MATCHING WITH (TABLOCK) (CUSTOMER__ID, collection__id, BUCKET_ATUALIZADO, balance)
        SELECT G.CUSTOMER__ID, G.collection__id,
               CAST(RIGHT(B.CHAVE, CHARINDEX('|', REVERSE(B.CHAVE)) - 1) AS VARCHAR(50)),
               G.balance
        FROM #GRUPO G
        LEFT JOIN #BUCKET_CLIENTE B ON B.CUSTOMER__ID = G.CUSTOMER__ID
        WHERE G.EM_ATRASO = 1;
        SET @LINHAS = @@ROWCOUNT;
        DROP TABLE #BUCKET_CLIENTE;
    END
    ELSE
    BEGIN
        -- LL: bucket do próprio CLIENTE + CONTRATO, como na consulta do LL
        INSERT INTO #MATCHING WITH (TABLOCK) (CUSTOMER__ID, collection__id, BUCKET_ATUALIZADO, balance)
        SELECT G.CUSTOMER__ID, G.collection__id,
               CAST(CASE WHEN G.CHAVE_BUCKET IS NULL THEN NULL
                         ELSE RIGHT(G.CHAVE_BUCKET, CHARINDEX('|', REVERSE(G.CHAVE_BUCKET)) - 1) END AS VARCHAR(50)),
               G.balance
        FROM #GRUPO G
        WHERE G.EM_ATRASO = 1;
        SET @LINHAS = @@ROWCOUNT;
    END;
    CREATE CLUSTERED INDEX IX ON #MATCHING (collection__id);
    CREATE NONCLUSTERED INDEX IX_CUSTOMER ON #MATCHING (CUSTOMER__ID);
    DROP TABLE #GRUPO;

    -- clientes distintos do matching (para filtrar a tabela de empréstimo)
    IF OBJECT_ID('tempdb..#CLI_MATCHING') IS NOT NULL DROP TABLE #CLI_MATCHING;
    SELECT DISTINCT CUSTOMER__ID INTO #CLI_MATCHING FROM #MATCHING;
    CREATE UNIQUE CLUSTERED INDEX IX ON #CLI_MATCHING (CUSTOMER__ID);

    SET @MSG = CONCAT(N'  ', @CART, N' 4.1b Matching .................... ', DATEDIFF(SECOND, @T_SECAO, SYSDATETIME()), N' s | ', @LINHAS, N' contratos');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;  SET @T_SECAO = SYSDATETIME();

    /* ---------------------------------------------------------------------------------
       4.2) MDR (cliente também tem empréstimo no período) e VERSÃO DA BASE (data de entrada)
       --------------------------------------------------------------------------------- */
    -- MDR: cliente desta base que aparece na tabela de empréstimo no período (mesma regra das duas consultas)
    IF OBJECT_ID('tempdb..#MDR_CLI') IS NOT NULL DROP TABLE #MDR_CLI;
    SELECT E.customer__id
    INTO #MDR_CLI
    FROM #EMP_CLI E
    WHERE EXISTS (SELECT 1 FROM #CLI_MATCHING M WHERE M.CUSTOMER__ID = E.customer__id);
    CREATE UNIQUE CLUSTERED INDEX IX ON #MDR_CLI (customer__id);

    -- Mesmo agrupamento da original (contrato + data de entrada + ativo).
    IF OBJECT_ID('tempdb..#ENTRADAS') IS NOT NULL DROP TABLE #ENTRADAS;
    SELECT CAST(A.CAMPO001 AS VARCHAR(36)) AS COLLECTION_ID,    -- 36, como o matching (evita índice de 1000 bytes)
           CAST(DT_ENTRADA AS DATE) AS DT_ENTRADA,
           B.ATIVO,
           CAST(CASE WHEN CAST(MAX(DT_ENTRADA) AS DATE) >= @DATA_INICIO THEN 'NOVOS CONTRATOS' ELSE 'ROLAGEM' END AS VARCHAR(20)) AS VERSAO_BASE
    INTO #ENTRADAS
    FROM [COBwebNUBANK].[dbo].[DIV_AUX] A WITH (NOLOCK)
    LEFT JOIN [COBwebNUBANK].[dbo].[DIV_MASTER] B WITH (NOLOCK) ON B.ID_DIVIDA = A.ID_DIVIDA
    INNER JOIN #MATCHING C ON C.collection__id = B.NR_OPERACAO
    WHERE CAST(DT_ENTRADA AS DATE) <= @DATA_FIM
    GROUP BY A.CAMPO013, CAST(DT_ENTRADA AS DATE), B.ATIVO, A.CAMPO001
    OPTION (RECOMPILE);
    CREATE CLUSTERED INDEX IX ON #ENTRADAS (COLLECTION_ID);

    SET @MSG = CONCAT(N'  ', @CART, N' 4.2 MDR e versão da base ......... ', DATEDIFF(SECOND, @T_SECAO, SYSDATETIME()), N' s');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;  SET @T_SECAO = SYSDATETIME();

    /* ---------------------------------------------------------------------------------
       4.3) BASE_FOTO — só as colunas que o funil usa. Uma linha de DIV_AUX por contrato.
       --------------------------------------------------------------------------------- */
    IF OBJECT_ID('tempdb..#AUX_CONTRATO') IS NOT NULL DROP TABLE #AUX_CONTRATO;
    SELECT CAMPO001, CAMPO013
    INTO #AUX_CONTRATO
    FROM (
        SELECT CAST(B.CAMPO001 AS VARCHAR(36)) AS CAMPO001, CAST(B.CAMPO013 AS VARCHAR(36)) AS CAMPO013,
               ROW_NUMBER() OVER (PARTITION BY B.CAMPO001 ORDER BY B.ID_DIVIDA) AS RN
        FROM [COBwebNUBANK].[dbo].[DIV_AUX] B WITH (NOLOCK)
        WHERE EXISTS (SELECT 1 FROM #MATCHING M WHERE M.collection__id = B.CAMPO001)
    ) X
    WHERE RN = 1;
    CREATE UNIQUE CLUSTERED INDEX IX ON #AUX_CONTRATO (CAMPO001);

    IF OBJECT_ID('tempdb..#BASE_FOTO') IS NOT NULL DROP TABLE #BASE_FOTO;
    SELECT A.CUSTOMER__ID, A.collection__id, A.BUCKET_ATUALIZADO, A.balance,
           CASE WHEN M.customer__id IS NOT NULL THEN 1 ELSE 0 END AS MDR,
           D.VERSAO_BASE,
           CASE WHEN E.MCI IS NOT NULL THEN 1 ELSE 0 END          AS DESENROLA,
           CASE WHEN F.MCI IS NOT NULL THEN 1 ELSE 0 END          AS CAMPANHA_INTERNA
    INTO #BASE_FOTO
    FROM #MATCHING A
    LEFT JOIN #AUX_CONTRATO     B ON B.CAMPO001      = A.collection__id
    LEFT JOIN #MDR_CLI          M ON M.customer__id  = B.CAMPO013
    LEFT JOIN #ENTRADAS         D ON D.COLLECTION_ID = A.collection__id
    LEFT JOIN #DESENROLA        E ON E.MCI           = A.CUSTOMER__ID
    LEFT JOIN #CAMPANHA_INTERNA F ON F.MCI           = A.CUSTOMER__ID;
    SET @LINHAS = @@ROWCOUNT;
    CREATE NONCLUSTERED INDEX IX_COLLECTION ON #BASE_FOTO (collection__id);
    CREATE NONCLUSTERED INDEX IX_CUSTOMER   ON #BASE_FOTO (CUSTOMER__ID);
    DROP TABLE #AUX_CONTRATO;
    DROP TABLE #MATCHING;
    DROP TABLE #CLI_MATCHING;

    SET @MSG = CONCAT(N'  ', @CART, N' 4.3 Base foto .................... ', DATEDIFF(SECOND, @T_SECAO, SYSDATETIME()), N' s | ', @LINHAS, N' linhas');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;  SET @T_SECAO = SYSDATETIME();

    /* ---------------------------------------------------------------------------------
       4.4) ACORDOS (vencimento no período) — regra da original.
            Novo: filtra a REC_MASTER ANTES dos joins (poucos milhares de linhas guiam o resto).
       --------------------------------------------------------------------------------- */
    -- (#REC_PERIODO já foi lido uma vez para o período, antes do laço das carteiras)

    IF OBJECT_ID('tempdb..#ACORDOS') IS NOT NULL DROP TABLE #ACORDOS;
    ;WITH ACORDOS_BRUTO AS (
        SELECT DISTINCT
            b.CPF_CNPJ,
            CC.NOME,
            D.NR_OPERACAO,
            e.MCI,
            A.ID_CARTEIRA,
            CAST(A.DT_VENCIMENTO AS DATE) AS DT_VENCIMENTO,
            CAST(A.DT_EMISSAO    AS DATE) AS DT_EMISSAO,
            1                             AS acordo,
            MAX(A.valor)                  AS cash,
            MAX(A.NR_RECEBIMENTO)         AS NR_RECEBIMENTO,
            FA.campo045                   AS BUCKET,
            ISNULL(FB.balance, CAMPO007)  AS BALANCE    -- CAMPO007 sem prefixo, como na original
        FROM #REC_PERIODO A
        LEFT JOIN [COBwebNUBANK].[dbo].[DEV_MASTER]    B  WITH (NOLOCK) ON A.ID_DEV = B.ID_DEV
        LEFT JOIN [COBwebNUBANK].[dbo].[REC_DIVIDAS]   C  WITH (NOLOCK) ON C.NR_RECEBIMENTO = A.NR_RECEBIMENTO
        LEFT JOIN [COBwebNUBANK].[dbo].[DIV_MASTER]    D  WITH (NOLOCK) ON D.ID_DIVIDA = C.ID_DIVIDA
        LEFT JOIN [COBwebNUBANK].[dbo].[DIV_DEVEDORES] E  WITH (NOLOCK) ON E.ID_DEV = A.ID_DEV
        LEFT JOIN #BASE_FOTO                           FB               ON FB.collection__id = CAST(D.NR_OPERACAO AS VARCHAR(36))
        LEFT JOIN [COBwebNUBANK].[dbo].[DIV_AUX]       FA WITH (NOLOCK) ON FA.ID_DIVIDA = D.ID_DIVIDA
        LEFT JOIN [COBwebNUBANK].[dbo].[USU_MASTER]    CC WITH (NOLOCK) ON CC.ID_USUARIO = A.ID_USUARIO
        GROUP BY b.CPF_CNPJ, A.DT_VENCIMENTO, A.ID_USUARIO, A.DT_EMISSAO, D.NR_OPERACAO, e.MCI, A.ID_CARTEIRA, CAMPO007,
                 FB.BUCKET_ATUALIZADO, FB.balance, FA.campo045, NOME
    ),
    ACORDOS_DEDUP AS (   -- acordo repetido (mesmo bucket e valor): fica o último
        SELECT *, ROW_NUMBER() OVER (PARTITION BY MCI, NR_OPERACAO, ID_CARTEIRA, BUCKET, cash
                                     ORDER BY NR_RECEBIMENTO DESC, DT_EMISSAO DESC, DT_VENCIMENTO DESC) AS RN
        FROM ACORDOS_BRUTO
    )
    SELECT CPF_CNPJ, NOME, NR_OPERACAO, MCI, ID_CARTEIRA, DT_VENCIMENTO, DT_EMISSAO, acordo, cash,
           CAST(BUCKET AS VARCHAR(50)) AS BUCKET, BALANCE
    INTO #ACORDOS
    FROM ACORDOS_DEDUP
    WHERE RN = 1
    OPTION (RECOMPILE);

    -- Acordo sem bucket: busca bucket/saldo na foto do mês anterior (só desses contratos)
    IF EXISTS (SELECT 1 FROM #ACORDOS WHERE BUCKET IS NULL)
    BEGIN
        IF OBJECT_ID(@TABELA_MATCHING_ANT) IS NULL
        BEGIN
            SET @MSG = CONCAT(N'  AVISO: ', @TABELA_MATCHING_ANT, N' não existe; acordos sem bucket ficam sem bucket.');
            RAISERROR(@MSG, 0, 1) WITH NOWAIT;
        END
        ELSE
        BEGIN
            IF OBJECT_ID('tempdb..#ACORDOS_SEM_BUCKET') IS NOT NULL DROP TABLE #ACORDOS_SEM_BUCKET;
            SELECT DISTINCT CAST(NR_OPERACAO AS VARCHAR(36)) AS collection__id
            INTO #ACORDOS_SEM_BUCKET FROM #ACORDOS WHERE BUCKET IS NULL;
            CREATE UNIQUE CLUSTERED INDEX IX ON #ACORDOS_SEM_BUCKET (collection__id);

            IF OBJECT_ID('tempdb..#FOTO_ANTERIOR') IS NOT NULL DROP TABLE #FOTO_ANTERIOR;
            CREATE TABLE #FOTO_ANTERIOR (collection__id VARCHAR(36), bucket_end_of_month VARCHAR(50), outstanding_balance DECIMAL(18,9));
            SET @SQL = N'
            INSERT INTO #FOTO_ANTERIOR (collection__id, bucket_end_of_month, outstanding_balance)
            SELECT CAST(FM.collection__id AS VARCHAR(36)), FM.bucket_end_of_month,
                   TRY_CAST(FM.outstanding_balance AS DECIMAL(18,9))
            FROM ' + QUOTENAME(@TABELA_MATCHING_ANT) + N' FM WITH (NOLOCK)
            WHERE EXISTS (SELECT 1 FROM #ACORDOS_SEM_BUCKET S WHERE S.collection__id = CAST(FM.collection__id AS VARCHAR(36)));';
            EXEC sp_executesql @SQL;
            CREATE CLUSTERED INDEX IX ON #FOTO_ANTERIOR (collection__id);

            UPDATE A SET BUCKET = FM.bucket_end_of_month, BALANCE = FM.outstanding_balance
            FROM #ACORDOS A
            INNER JOIN #FOTO_ANTERIOR FM ON FM.collection__id = A.NR_OPERACAO
            WHERE A.BUCKET IS NULL;
            DROP TABLE #FOTO_ANTERIOR;
            DROP TABLE #ACORDOS_SEM_BUCKET;
        END;
    END;

    -- Um acordo por cliente+contrato: o de maior vencimento (como na original).
    IF OBJECT_ID('tempdb..#ACORDOS_AGG') IS NOT NULL DROP TABLE #ACORDOS_AGG;
    ;WITH ACORDOS_RANK AS (
        SELECT *, ROW_NUMBER() OVER (PARTITION BY MCI, NR_OPERACAO
                                     ORDER BY CAST(acordo AS BIGINT) DESC, CAST(DT_VENCIMENTO AS DATE) DESC) AS RN
        FROM #ACORDOS
    )
    SELECT MCI                         AS customer__id,
           NR_OPERACAO                 AS collection__id,
           ID_CARTEIRA,
           CAST(acordo AS BIGINT)      AS ACORDO,
           COALESCE(TRY_CONVERT(DECIMAL(18,2), cash),
                    TRY_CONVERT(DECIMAL(18,2), REPLACE(CAST(cash AS VARCHAR(50)), ',', '.'))) AS VALOR_ACORDO,
           BUCKET,
           BALANCE,
           CAST(DT_VENCIMENTO AS DATE) AS DT_VENCIMENTO
    INTO #ACORDOS_AGG
    FROM ACORDOS_RANK
    WHERE RN = 1;
    CREATE UNIQUE NONCLUSTERED INDEX IX ON #ACORDOS_AGG (customer__id, collection__id) INCLUDE (ID_CARTEIRA, DT_VENCIMENTO);
    DROP TABLE #ACORDOS;

    -- Bucket do acordo substitui o da foto (só quando o acordo trouxe bucket)
    UPDATE B SET B.BUCKET_ATUALIZADO = A.BUCKET
    FROM #BASE_FOTO B
    INNER JOIN #ACORDOS_AGG A ON B.collection__id = A.collection__id
    WHERE A.BUCKET IS NOT NULL;

    -- Acordo CC de contrato fora da base: entram os contratos do cliente (como na original)
    IF OBJECT_ID('tempdb..#CLIENTES_FORA_ACORDO') IS NOT NULL DROP TABLE #CLIENTES_FORA_ACORDO;
    SELECT A.customer__id, A.collection__id, A.BUCKET, A.BALANCE
    INTO #CLIENTES_FORA_ACORDO
    FROM #ACORDOS_AGG A
    WHERE A.ID_CARTEIRA = @CART_ID
      AND NOT EXISTS (SELECT 1 FROM #BASE_FOTO B WHERE B.collection__id = A.collection__id);

    INSERT INTO #BASE_FOTO (CUSTOMER__ID, collection__id, BUCKET_ATUALIZADO, balance, MDR, VERSAO_BASE, DESENROLA, CAMPANHA_INTERNA)
    SELECT B.CAMPO013, B.CAMPO001, A.BUCKET, ISNULL(A.BALANCE, B.CAMPO007),
           CASE WHEN F.customer__id IS NOT NULL THEN 1 ELSE 0 END,
           CASE WHEN CAST(MAX(DT_ENTRADA) AS DATE) >= @DATA_INICIO THEN 'NOVOS CONTRATOS' ELSE 'ROLAGEM' END,
           CASE WHEN E.MCI IS NOT NULL THEN 1 ELSE 0 END,
           CASE WHEN D.MCI IS NOT NULL THEN 1 ELSE 0 END
    FROM #CLIENTES_FORA_ACORDO A
    INNER JOIN COBwebNUBANK.dbo.DIV_AUX    B WITH (NOLOCK) ON B.CAMPO013 = A.customer__id
    INNER JOIN COBwebNUBANK.dbo.DIV_MASTER C WITH (NOLOCK) ON C.ID_DIVIDA = B.ID_DIVIDA
    LEFT JOIN #CAMPANHA_INTERNA D ON D.MCI = B.CAMPO013
    LEFT JOIN #DESENROLA        E ON E.MCI = B.CAMPO013
    LEFT JOIN #MDR_CLI          F ON F.customer__id = A.customer__id
    GROUP BY B.CAMPO013, B.CAMPO001, A.BUCKET, A.BALANCE, B.CAMPO007, F.customer__id, E.MCI, D.MCI;
    SET @LINHAS = @@ROWCOUNT;
    DROP TABLE #CLIENTES_FORA_ACORDO;

    SET @MSG = CONCAT(N'  ', @CART, N' 4.4 Acordos ...................... ', DATEDIFF(SECOND, @T_SECAO, SYSDATETIME()), N' s | ',
                      (SELECT COUNT(*) FROM #ACORDOS_AGG), N' contratos com acordo | ', @LINHAS, N' linhas novas na base');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;  SET @T_SECAO = SYSDATETIME();

    /* ---------------------------------------------------------------------------------
       4.5) PAGAMENTOS CC — regra da original:
              DIRECT           -> payment_type = 'direct' (não exige janela)
              INDIRECT_*       -> só entra se cair entre D-2 e D+2 do vencimento do acordo CC do contrato
            PAGO = 1 para DIRECT ou INDIRECT válido; o VALOR do funil é só o DIRECT.
       --------------------------------------------------------------------------------- */
    -- LL: a que contrato vai o pagamento de cada loan NESTE período (@LL_VINCULO_PAGAMENTO).
    --   Um loan pode estar em mais de um contrato na DIV_AUX (histórico). "Na base" = o contrato existe em #BASE_FOTO.
    --   B: fica só com os contratos da base (se o loan não tem nenhum na base, mantém todos, como na original).
    --   C: fica com UM contrato por loan: o da base; havendo mais de um na base (ou nenhum), o mais recente (maior ID_DIVIDA).
    IF OBJECT_ID('tempdb..#LOAN_ESCOLHA') IS NOT NULL DROP TABLE #LOAN_ESCOLHA;
    CREATE TABLE #LOAN_ESCOLHA (loan__id VARCHAR(36) NOT NULL, collection__id VARCHAR(36) NOT NULL);
    IF @CART = 'LL' AND @LL_VINCULO_PAGAMENTO IN ('B', 'C')
    BEGIN
        ;WITH CAND AS (
            SELECT M.loan__id, M.collection__id, M.ID_DIVIDA_REC,
                   CASE WHEN EXISTS (SELECT 1 FROM #BASE_FOTO B WHERE B.collection__id = M.collection__id) THEN 1 ELSE 0 END AS NA_BASE
            FROM #LOAN_COLL M
        ),
        TEM AS (
            SELECT loan__id, MAX(NA_BASE) AS ALGUM_NA_BASE FROM CAND GROUP BY loan__id
        ),
        ORD AS (
            SELECT C.loan__id, C.collection__id, C.NA_BASE, T.ALGUM_NA_BASE,
                   ROW_NUMBER() OVER (PARTITION BY C.loan__id
                                      ORDER BY C.NA_BASE DESC, ISNULL(C.ID_DIVIDA_REC, 0) DESC, C.collection__id) AS RN
            FROM CAND C
            INNER JOIN TEM T ON T.loan__id = C.loan__id
        )
        INSERT INTO #LOAN_ESCOLHA (loan__id, collection__id)
        SELECT loan__id, collection__id
        FROM ORD
        WHERE (@LL_VINCULO_PAGAMENTO = 'C' AND RN = 1)
           OR (@LL_VINCULO_PAGAMENTO = 'B' AND (NA_BASE = 1 OR ALGUM_NA_BASE = 0));
        CREATE UNIQUE CLUSTERED INDEX IX ON #LOAN_ESCOLHA (loan__id, collection__id);
        IF @MOSTRAR_DIAGNOSTICO = 1
        BEGIN
            SELECT @ESC_LOANS = COUNT(DISTINCT loan__id) FROM #LOAN_COLL;
            SELECT @ESC_MULTI = COUNT(*) FROM (
                SELECT M.loan__id FROM #LOAN_COLL M
                WHERE EXISTS (SELECT 1 FROM #BASE_FOTO B WHERE B.collection__id = M.collection__id)
                GROUP BY M.loan__id HAVING COUNT(*) > 1) X;
            SET @MSG = CONCAT(N'  LL vínculo do pagamento: regra ', @LL_VINCULO_PAGAMENTO, N' | ', @ESC_LOANS, N' loans no arquivo | ',
                              @ESC_MULTI, N' com 2 ou mais contratos na base (',
                              CASE WHEN @LL_VINCULO_PAGAMENTO = 'C' THEN N'o pagamento vai a 1 só)' ELSE N'o pagamento conta em cada um)' END);
            RAISERROR(@MSG, 0, 1) WITH NOWAIT;
        END;
    END;

    IF OBJECT_ID('tempdb..#PAGAMENTOS_AGG') IS NOT NULL DROP TABLE #PAGAMENTOS_AGG;
    ;WITH PAGAMENTOS_BASE AS (
        SELECT customer__id, collection__id, DT_HORA, DT_PAGAMENTO, payment_type, finance_days_late_bucket, VALOR
        FROM #PAG P
        WHERE CARTEIRA = @CART
          AND DT_PAGAMENTO BETWEEN @DATA_INICIO AND @DATA_FIM
          AND VALOR IS NOT NULL
          AND (@CART = 'CC' OR @LL_VINCULO_PAGAMENTO = 'A'
               OR EXISTS (SELECT 1 FROM #LOAN_ESCOLHA E WHERE E.loan__id = P.loan__id AND E.collection__id = P.collection__id))
    ),
    PAGAMENTOS_VALIDOS AS (
        SELECT b.*,
               CASE WHEN b.payment_type = 'direct' THEN 1 ELSE 0 END AS E_DIRECT
        FROM PAGAMENTOS_BASE b
        WHERE b.payment_type = 'direct'
           OR EXISTS (SELECT 1
                      FROM #ACORDOS_AGG a
                      WHERE a.ID_CARTEIRA    = @CART_ID
                        AND a.customer__id   = b.customer__id
                        AND a.collection__id = b.collection__id
                        AND b.DT_PAGAMENTO BETWEEN DATEADD(DAY, -@DIAS_ANTES_VENC, a.DT_VENCIMENTO)
                                               AND DATEADD(DAY,  @DIAS_APOS_VENC,  a.DT_VENCIMENTO))
    ),
    PAGAMENTOS_REGRA AS (
        SELECT v.*, ROW_NUMBER() OVER (PARTITION BY v.customer__id, v.collection__id ORDER BY v.DT_HORA DESC) AS RN_ULTIMO
        FROM PAGAMENTOS_VALIDOS v
    )
    SELECT customer__id,
           collection__id,
           CAST(1 AS INT)                                                     AS PAGO,
           SUM(CASE WHEN E_DIRECT = 1 THEN VALOR ELSE 0 END)                  AS VALOR_PAGO_DIRECT,
           MAX(CASE WHEN RN_ULTIMO = 1 THEN finance_days_late_bucket END)     AS BUCKET_NO_PAGAMENTO
    INTO #PAGAMENTOS_AGG
    FROM PAGAMENTOS_REGRA
    GROUP BY customer__id, collection__id;
    SET @LINHAS = @@ROWCOUNT;
    CREATE UNIQUE NONCLUSTERED INDEX IX ON #PAGAMENTOS_AGG (customer__id, collection__id);

    -- Bucket do pagamento substitui o da base; contratos pagos fora da base entram na base
    UPDATE B SET B.BUCKET_ATUALIZADO = A.BUCKET_NO_PAGAMENTO
    FROM #BASE_FOTO B
    INNER JOIN #PAGAMENTOS_AGG A ON B.collection__id = A.collection__id;

    IF OBJECT_ID('tempdb..#CLIENTES_FORA') IS NOT NULL DROP TABLE #CLIENTES_FORA;
    SELECT A.customer__id, A.collection__id, A.BUCKET_NO_PAGAMENTO
    INTO #CLIENTES_FORA
    FROM #PAGAMENTOS_AGG A
    WHERE NOT EXISTS (SELECT 1 FROM #BASE_FOTO B WHERE B.collection__id = A.collection__id);

    INSERT INTO #BASE_FOTO (CUSTOMER__ID, collection__id, BUCKET_ATUALIZADO, balance, MDR, VERSAO_BASE, DESENROLA, CAMPANHA_INTERNA)
    SELECT B.CAMPO013, B.CAMPO001, A.BUCKET_NO_PAGAMENTO, B.CAMPO007,
           CASE WHEN F.customer__id IS NOT NULL THEN 1 ELSE 0 END,
           CASE WHEN CAST(MAX(DT_ENTRADA) AS DATE) >= @DATA_INICIO THEN 'NOVOS CONTRATOS' ELSE 'ROLAGEM' END,
           CASE WHEN E.MCI IS NOT NULL THEN 1 ELSE 0 END,
           CASE WHEN D.MCI IS NOT NULL THEN 1 ELSE 0 END
    FROM #CLIENTES_FORA A
    INNER JOIN COBwebNUBANK.dbo.DIV_AUX    B WITH (NOLOCK) ON B.CAMPO013 = A.customer__id
    INNER JOIN COBwebNUBANK.dbo.DIV_MASTER C WITH (NOLOCK) ON C.ID_DIVIDA = B.ID_DIVIDA
    LEFT JOIN #CAMPANHA_INTERNA D ON D.MCI = B.CAMPO013
    LEFT JOIN #DESENROLA        E ON E.MCI = B.CAMPO013
    LEFT JOIN #MDR_CLI          F ON F.customer__id = A.customer__id
    GROUP BY B.CAMPO013, B.CAMPO001, A.BUCKET_NO_PAGAMENTO, B.CAMPO007, F.customer__id, E.MCI, D.MCI;
    SET @LINHAS_INS = @@ROWCOUNT;

    SET @MSG = CONCAT(N'  ', @CART, N' 4.5 Pagamentos ................... ', DATEDIFF(SECOND, @T_SECAO, SYSDATETIME()), N' s | ',
                      @LINHAS, N' contratos pagos | ', (SELECT COUNT(*) FROM #CLIENTES_FORA), N' fora da base (',
                      @LINHAS_INS, N' linhas inseridas na base por essa regra: todos os contratos do cliente na DIV_AUX)');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;  SET @T_SECAO = SYSDATETIME();
    DROP TABLE #CLIENTES_FORA;
    DROP TABLE #LOAN_ESCOLHA;

    IF @MOSTRAR_DIAGNOSTICO = 1
    BEGIN
        SELECT @DUPLICADOS = COUNT(*) - COUNT(DISTINCT collection__id) FROM #BASE_FOTO;
        IF @DUPLICADOS > 0
        BEGIN
            SET @MSG = CONCAT(N'  ', @CART, N' DIAGNÓSTICO: ', @DUPLICADOS, N' linhas repetem um contrato na base (contrato com mais de uma ',
                              N'entrada/ativo, ou contratos do cliente inseridos como "fora da base"). É o comportamento ',
                              N'da consulta original e entra na contagem da BASE; mantido de propósito.');
            RAISERROR(@MSG, 0, 1) WITH NOWAIT;
        END;
    END;

    /* ---------------------------------------------------------------------------------
       4.7) BASE_FINAL — base + discagem + acordo CC + pagamento, JÁ com os ajustes do funil:
              • bucket padronizado pelo de/para;
              • cascata: pago => acordo => CPC => alô => atendido => discado;
              • valor do acordo nunca menor que o valor pago.
            Tudo no próprio SELECT (a original fazia ~8 UPDATEs sobre a tabela inteira).
       --------------------------------------------------------------------------------- */
    IF OBJECT_ID('tempdb..#BASE_FINAL') IS NOT NULL DROP TABLE #BASE_FINAL;
    SELECT A.CUSTOMER__ID, A.collection__id,
           ISNULL(M.DESTINO, A.BUCKET_ATUALIZADO)   AS BUCKET_ATUALIZADO,
           A.balance, A.MDR, A.VERSAO_BASE, A.DESENROLA, A.CAMPANHA_INTERNA,
           ISNULL(D.QTD_DISCADO, 0)                 AS QTD_DISCADO,
           C5.DISCADO, C4.ATENDIDO, C3.ALO, C2.CPC, C1.ACORDO,
           CASE WHEN ISNULL(P.VALOR_PAGO_DIRECT, 0) > ISNULL(AC.VALOR_ACORDO, 0)
                THEN ISNULL(P.VALOR_PAGO_DIRECT, 0) ELSE ISNULL(AC.VALOR_ACORDO, 0) END AS VALOR_ACORDO,
           ISNULL(P.PAGO, 0)                        AS PAGO,
           ISNULL(P.VALOR_PAGO_DIRECT, 0)           AS VALOR_PAGO_DIRECT
    INTO #BASE_FINAL
    FROM #BASE_FOTO A
    LEFT JOIN #DISCAGENS     D  ON D.customer__id  = A.CUSTOMER__ID
    LEFT JOIN #ACORDOS_AGG   AC ON AC.customer__id = A.CUSTOMER__ID AND AC.collection__id = A.collection__id
                               AND AC.ID_CARTEIRA  = @CART_ID              -- só o acordo da própria carteira
    LEFT JOIN #PAGAMENTOS_AGG P ON P.customer__id  = A.CUSTOMER__ID AND P.collection__id  = A.collection__id
    LEFT JOIN #DE_PARA_BUCKET M ON M.ORIGEM COLLATE DATABASE_DEFAULT = REPLACE(A.BUCKET_ATUALIZADO, ':', '.') COLLATE DATABASE_DEFAULT
    CROSS APPLY (SELECT CASE WHEN ISNULL(P.PAGO, 0) = 1 THEN 1 ELSE ISNULL(AC.ACORDO, 0) END         AS ACORDO)   C1
    CROSS APPLY (SELECT CASE WHEN C1.ACORDO   = 1 THEN 1 ELSE ISNULL(D.CPC, 0)      END              AS CPC)      C2
    CROSS APPLY (SELECT CASE WHEN C2.CPC      = 1 THEN 1 ELSE ISNULL(D.ALO, 0)      END              AS ALO)      C3
    CROSS APPLY (SELECT CASE WHEN C3.ALO      = 1 THEN 1 ELSE ISNULL(D.ATENDIDO, 0) END              AS ATENDIDO) C4
    CROSS APPLY (SELECT CASE WHEN C4.ATENDIDO = 1 THEN 1 ELSE ISNULL(D.DISCADO, 0)  END              AS DISCADO)  C5;
    SET @LINHAS = @@ROWCOUNT;
    DROP TABLE #BASE_FOTO;
    DROP TABLE #ACORDOS_AGG;
    DROP TABLE #PAGAMENTOS_AGG;
    DROP TABLE #MDR_CLI;
    DROP TABLE #ENTRADAS;

    SET @MSG = CONCAT(N'  ', @CART, N' 4.7 Base final ................... ', DATEDIFF(SECOND, @T_SECAO, SYSDATETIME()), N' s | ', @LINHAS, N' linhas');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;  SET @T_SECAO = SYSDATETIME();

    -- Aviso: bucket que sobrou fora do padrão (não estava no de/para) — fica como veio
    SET @FORA_PADRAO = STUFF((
        SELECT TOP (10) N', ' + BUCKET_ATUALIZADO + N'=' + CAST(COUNT(*) AS NVARCHAR(20))
        FROM #BASE_FINAL
        WHERE BUCKET_ATUALIZADO IS NOT NULL
          AND BUCKET_ATUALIZADO NOT IN ('0.0-60','1.61-90','2.91-120','3.121-180','4.181-360','5.361-720','6.721-1080','7.1081-1440','8.1441+')
          AND REPLACE(BUCKET_ATUALIZADO, ':', '.') NOT LIKE '00.%'
        GROUP BY BUCKET_ATUALIZADO
        ORDER BY COUNT(*) DESC
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(1000)'), 1, 2, N'');
    IF @FORA_PADRAO IS NOT NULL
    BEGIN
        SET @MSG = REPLACE(CONCAT(N'  ', @CART, N' AVISO: buckets fora do padrão do CC (ficam como vieram; incluir no de/para): ', @FORA_PADRAO), N'%', N'%%');
        RAISERROR(@MSG, 0, 1) WITH NOWAIT;
    END;

    /* ---------------------------------------------------------------------------------
       4.8) FUNIL LONGO — 8 linhas (TIPO) por segmento, como na original.
            Uma leitura da base + "despivotar" (a original lê a base 8 vezes).
       --------------------------------------------------------------------------------- */
    IF @EXCLUIR_BUCKETS_INVALIDOS = 1
    BEGIN
        SELECT @EXCLUIDOS = COUNT(*) FROM #BASE_FINAL
        WHERE BUCKET_ATUALIZADO IS NULL OR REPLACE(BUCKET_ATUALIZADO, ':', '.') LIKE '00.%';
        SET @MSG = CONCAT(N'  ', @CART, N' Buckets inválidos ("(vazio)" e "00.xx") fora do funil: ', @EXCLUIDOS, N' contratos');
        RAISERROR(@MSG, 0, 1) WITH NOWAIT;
    END;

    ;WITH BF AS (
        SELECT *,
               ISNULL(BUCKET_ATUALIZADO, '(vazio)') AS BUCKET_FINAL,
               CASE WHEN ISNULL(balance, 0) < 200   THEN '0.Até 200'
                    WHEN ISNULL(balance, 0) < 500   THEN '1.200-500'
                    WHEN ISNULL(balance, 0) < 1000  THEN '2.500-1.000'
                    WHEN ISNULL(balance, 0) < 2500  THEN '3.1.000-2.500'
                    WHEN ISNULL(balance, 0) < 5000  THEN '4.2.500-5.000'
                    WHEN ISNULL(balance, 0) < 10000 THEN '5.5.000-10.000'
                    ELSE '6.10.000+' END AS FAIXA_DE_VALOR
        FROM #BASE_FINAL
    ),
    SEGMENTO AS (   -- agrupa por VERSAO_BASE "crua" (NULL separado de ROLAGEM), igual à original
        SELECT FAIXA_DE_VALOR, BUCKET_FINAL, VERSAO_BASE, CAMPANHA_INTERNA, DESENROLA, MDR,
               COUNT(*)                                                                 AS Q_BASE,
               ISNULL(SUM(balance), 0)                                                  AS B_BASE,
               SUM(DISCADO)                                                             AS Q_DISCADO,
               ISNULL(SUM(CASE WHEN DISCADO  = 1 THEN balance ELSE 0 END), 0)           AS B_DISCADO,
               SUM(ATENDIDO)                                                            AS Q_ATENDIDO,
               ISNULL(SUM(CASE WHEN ATENDIDO = 1 THEN balance ELSE 0 END), 0)           AS B_ATENDIDO,
               SUM(ALO)                                                                 AS Q_ALO,
               ISNULL(SUM(CASE WHEN ALO      = 1 THEN balance ELSE 0 END), 0)           AS B_ALO,
               SUM(CPC)                                                                 AS Q_CPC,
               ISNULL(SUM(CASE WHEN CPC      = 1 THEN balance ELSE 0 END), 0)           AS B_CPC,
               SUM(ACORDO)                                                              AS Q_ACORDO,
               ISNULL(SUM(CASE WHEN ACORDO   = 1 THEN balance ELSE 0 END), 0)           AS B_ACORDO,
               ISNULL(SUM(CASE WHEN ACORDO   = 1 THEN VALOR_ACORDO ELSE 0 END), 0)      AS C_ACORDO,   -- valor do acordo
               SUM(PAGO)                                                                AS Q_PAGO,
               ISNULL(SUM(CASE WHEN PAGO     = 1 THEN balance ELSE 0 END), 0)           AS B_PAGO,
               ISNULL(SUM(CASE WHEN PAGO     = 1 THEN VALOR_PAGO_DIRECT ELSE 0 END), 0) AS C_PAGO,     -- só pagamento direct
               SUM(QTD_DISCADO)                                                         AS Q_LIGACOES, -- tentativas (base do "spin")
               ISNULL(SUM(CASE WHEN QTD_DISCADO > 0 THEN balance ELSE 0 END), 0)        AS B_LIGACOES
        FROM BF
        WHERE @EXCLUIR_BUCKETS_INVALIDOS = 0
           OR (BUCKET_FINAL <> '(vazio)' AND REPLACE(BUCKET_FINAL, ':', '.') NOT LIKE '00.%')   -- vale para todas as etapas do funil
        GROUP BY FAIXA_DE_VALOR, BUCKET_FINAL, VERSAO_BASE, CAMPANHA_INTERNA, DESENROLA, MDR
    )
    INSERT INTO #FUNIL_LONGO_TOTAL (REFERENCIA, PERIODO_INICIO, PERIODO_FIM, DIAS_UTEIS,
                                    MES, FAIXA_DE_VALOR, BUCKET_ATUALIZADO, TIPO, QTD, BALANCE, CASH,
                                    VERSAO_BASE, CARTEIRA, CAMPANHA_INTERNA, DESENROLA, MDR, FUNIL)
    SELECT @REFERENCIA_PERIODO, @DATA_INICIO, @DATA_FIM, @DIAS_UTEIS_PERIODO,
           @MES_REFERENCIA, S.FAIXA_DE_VALOR, S.BUCKET_FINAL, T.TIPO, T.QTD, T.BALANCE, T.CASH,
           ISNULL(S.VERSAO_BASE, 'ROLAGEM'), 'CC', S.CAMPANHA_INTERNA, S.DESENROLA, S.MDR, @CART
    FROM SEGMENTO S
    CROSS APPLY (VALUES
        (N'BASE',     S.Q_BASE,     S.B_BASE,     CAST(0 AS DECIMAL(38,2))),
        (N'DISCADO',  S.Q_DISCADO,  S.B_DISCADO,  0),
        (N'ATENDIDO', S.Q_ATENDIDO, S.B_ATENDIDO, 0),
        (N'ALÔ',      S.Q_ALO,      S.B_ALO,      0),
        (N'CPC',      S.Q_CPC,      S.B_CPC,      0),
        (N'ACORDO',   S.Q_ACORDO,   S.B_ACORDO,   S.C_ACORDO),
        (N'PAGO',     S.Q_PAGO,     S.B_PAGO,     S.C_PAGO),
        (N'LIGACOES', S.Q_LIGACOES, S.B_LIGACOES, 0)
    ) T (TIPO, QTD, BALANCE, CASH);
    DROP TABLE #BASE_FINAL;

    SET @MSG = CONCAT(N'  ', @CART, N' 4.8 Funil longo ............... ', DATEDIFF(SECOND, @T_SECAO, SYSDATETIME()), N' s', NCHAR(10),
                      N'  ', @CART, N' concluído em ', DATEDIFF(SECOND, @T_CART, SYSDATETIME()), N' s');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;  SET @T_SECAO = SYSDATETIME();
    END;   -- fim do laço de carteiras

    DROP TABLE #DISCAGENS;
    DROP TABLE #EMP_CLI;
    DROP TABLE #REC_PERIODO;
    SET @MSG = CONCAT(N'  Período ', @REFERENCIA_PERIODO, N' concluído em ', DATEDIFF(SECOND, @T_PERIODO, SYSDATETIME()), N' s');
    RAISERROR(@MSG, 0, 1) WITH NOWAIT;
END;   -- fim do laço de períodos

IF @CACHE_ATIVO = 1
    EXEC sp_releaseapplock @Resource = N'FUNIL_NUBANK_CACHE', @LockOwner = 'Session';

/* =====================================================================================
   5) RESULTADO
   ===================================================================================== */
SET @MSG = CONCAT(NCHAR(10), N'=== TEMPO TOTAL: ', DATEDIFF(SECOND, @T_GERAL, SYSDATETIME()), N' s ===');
RAISERROR(@MSG, 0, 1) WITH NOWAIT;

IF NOT EXISTS (SELECT 1 FROM #FUNIL_LONGO_TOTAL)
BEGIN
    RAISERROR('Nenhum período foi processado (ver mensagens acima).', 0, 1) WITH NOWAIT;
    RETURN;
END;

-- BALANCE e CASH saem com 2 casas (o Excel só guarda 15 dígitos). @DECIMAL_VIRGULA = 1: vírgula, cola como número no Excel pt-BR.
-- Grade 1: as 12 colunas de sempre da aba "Dados" + FUNIL no final ('CC' ou 'LL').
--          A coluna CARTEIRA continua como nas consultas originais ('CC' nas duas).
IF @DECIMAL_VIRGULA = 1
    SELECT MES, FAIXA_DE_VALOR, BUCKET_ATUALIZADO, TIPO, QTD,
           REPLACE(CONVERT(VARCHAR(50), CAST(BALANCE AS DECIMAL(38,2))), '.', ',') AS BALANCE,
           REPLACE(CONVERT(VARCHAR(50), CAST(CASH    AS DECIMAL(38,2))), '.', ',') AS CASH,
           VERSAO_BASE, CARTEIRA, CAMPANHA_INTERNA, DESENROLA, MDR, FUNIL
    FROM #FUNIL_LONGO_TOTAL
    ORDER BY FUNIL, PERIODO_INICIO, TIPO, BUCKET_ATUALIZADO, FAIXA_DE_VALOR;
ELSE
    SELECT MES, FAIXA_DE_VALOR, BUCKET_ATUALIZADO, TIPO, QTD,
           CAST(BALANCE AS DECIMAL(38,2)) AS BALANCE,
           CAST(CASH    AS DECIMAL(38,2)) AS CASH,
           VERSAO_BASE, CARTEIRA, CAMPANHA_INTERNA, DESENROLA, MDR, FUNIL
    FROM #FUNIL_LONGO_TOTAL
    ORDER BY FUNIL, PERIODO_INICIO, TIPO, BUCKET_ATUALIZADO, FAIXA_DE_VALOR;

-- Grade 2: total por etapa de cada período e carteira (o mesmo que a aba "Base Dinâmica" mostra)
IF @MOSTRAR_RESUMO = 1
BEGIN
    IF OBJECT_ID('tempdb..#RESUMO') IS NOT NULL DROP TABLE #RESUMO;
    SELECT FUNIL, MES, REFERENCIA, PERIODO_INICIO, PERIODO_FIM, DIAS_UTEIS, TIPO,
           SUM(QTD) AS QTD, SUM(BALANCE) AS BALANCE, SUM(CASH) AS CASH
    INTO #RESUMO
    FROM #FUNIL_LONGO_TOTAL
    GROUP BY FUNIL, MES, REFERENCIA, PERIODO_INICIO, PERIODO_FIM, DIAS_UTEIS, TIPO;

    IF @DECIMAL_VIRGULA = 1
        SELECT FUNIL, MES, REFERENCIA, PERIODO_INICIO, PERIODO_FIM, DIAS_UTEIS, TIPO, QTD,
               REPLACE(CONVERT(VARCHAR(50), CAST(BALANCE AS DECIMAL(38,2))), '.', ',') AS BALANCE,
               REPLACE(CONVERT(VARCHAR(50), CAST(CASH    AS DECIMAL(38,2))), '.', ',') AS CASH
        FROM #RESUMO
        ORDER BY FUNIL, PERIODO_INICIO,
                 CASE TIPO WHEN N'BASE' THEN 1 WHEN N'DISCADO' THEN 2 WHEN N'ATENDIDO' THEN 3 WHEN N'ALÔ' THEN 4
                           WHEN N'CPC' THEN 5 WHEN N'ACORDO' THEN 6 WHEN N'PAGO' THEN 7 ELSE 8 END;
    ELSE
        SELECT FUNIL, MES, REFERENCIA, PERIODO_INICIO, PERIODO_FIM, DIAS_UTEIS, TIPO, QTD,
               CAST(BALANCE AS DECIMAL(38,2)) AS BALANCE,
               CAST(CASH    AS DECIMAL(38,2)) AS CASH
        FROM #RESUMO
        ORDER BY FUNIL, PERIODO_INICIO,
                 CASE TIPO WHEN N'BASE' THEN 1 WHEN N'DISCADO' THEN 2 WHEN N'ATENDIDO' THEN 3 WHEN N'ALÔ' THEN 4
                           WHEN N'CPC' THEN 5 WHEN N'ACORDO' THEN 6 WHEN N'PAGO' THEN 7 ELSE 8 END;
END;