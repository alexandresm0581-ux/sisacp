-- =====================================================================
-- SISACP - Sistema de Acompanhamento e Controle de Peso
-- Esquema do banco de dados (SQLite)
--
-- Convenções:
--   * Datas   -> texto no formato ISO  'AAAA-MM-DD'   (ex.: '2026-11-13')
--   * Horas   -> texto 'HH:MM'                         (ex.: '12:30')
--   * Nada que possa ser CALCULADO é gravado (IMC, totais, % da meta...).
--     Esses valores saem de VIEWS (consultas salvas) ou do código Python.
-- =====================================================================

PRAGMA foreign_keys = ON;

-- ---------------------------------------------------------------------
-- PERFIL: uma única linha (id = 1) com seus dados de referência.
-- Equivale ao bloco "DADOS DE REFERÊNCIA" do painel.
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS perfil (
    id                  INTEGER PRIMARY KEY CHECK (id = 1),
    nome                TEXT    NOT NULL,
    data_nascimento     TEXT    NOT NULL,
    altura_m            REAL    NOT NULL CHECK (altura_m > 0),
    peso_meta_kg        REAL    NOT NULL CHECK (peso_meta_kg > 0),
    data_inicio_acomp   TEXT    NOT NULL,
    meta_agua_ml        INTEGER NOT NULL DEFAULT 3000
);

-- ---------------------------------------------------------------------
-- PESAGEM  (aba "Controle de Peso")
-- IMC e status NÃO são gravados: ver vw_pesagem mais abaixo.
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS pesagem (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    data      TEXT NOT NULL,
    peso_kg   REAL NOT NULL CHECK (peso_kg > 0)
);
CREATE INDEX IF NOT EXISTS ix_pesagem_data ON pesagem(data);

-- ---------------------------------------------------------------------
-- HIDRATAÇÃO  (aba "Controle de Hidratação")
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS hidratacao (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    data       TEXT    NOT NULL,
    hora       TEXT    NOT NULL,
    volume_ml  INTEGER NOT NULL CHECK (volume_ml > 0)
);
CREATE INDEX IF NOT EXISTS ix_hidratacao_data ON hidratacao(data);

-- ---------------------------------------------------------------------
-- ATIVIDADE FÍSICA  (aba "Controle Atividades Fisicas")
-- kcal_gastas e distancia_km existiam no formulário de coleta do Excel
-- (ciclismo/vôlei), mas nunca foram guardados. Já ficam previstos aqui.
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS atividade_fisica (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    data          TEXT    NOT NULL,
    atividade     TEXT    NOT NULL,           -- Musculação, Ciclismo, Vôlei...
    hora_inicio   TEXT,
    duracao_min   INTEGER NOT NULL CHECK (duracao_min > 0),
    kcal_gastas   REAL,
    distancia_km  REAL
);
CREATE INDEX IF NOT EXISTS ix_atividade_data ON atividade_fisica(data);

-- ---------------------------------------------------------------------
-- METAS POR REFEIÇÃO  (Apoio > Tabela6)
-- 'ordem' diz a sequência do dia; é usada em "meta até a última refeição".
-- A meta do dia = soma das 5 linhas (nada é gravado como total).
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS meta_refeicao (
    refeicao         TEXT    PRIMARY KEY,
    ordem            INTEGER NOT NULL UNIQUE,
    horario          TEXT,
    kcal             REAL NOT NULL,
    proteinas_g      REAL NOT NULL,
    carboidratos_g   REAL NOT NULL,
    gorduras_g       REAL NOT NULL
);

-- ---------------------------------------------------------------------
-- ITEM DO PLANO  (Apoio > Tabela5)
-- Cada linha é uma PORÇÃO PRESCRITA (ex.: "Ovos mexidos, 150 g = 221 kcal").
-- No Excel a chave era "Refeição | Cod Item"; aqui é (refeicao, cod_item).
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS item_plano (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    refeicao         TEXT    NOT NULL REFERENCES meta_refeicao(refeicao),
    cod_item         INTEGER NOT NULL,
    descricao        TEXT    NOT NULL,
    peso_g           REAL    NOT NULL,
    kcal             REAL    NOT NULL,
    proteinas_g      REAL    NOT NULL,
    carboidratos_g   REAL    NOT NULL,
    gorduras_g       REAL    NOT NULL,
    UNIQUE (refeicao, cod_item)
);

-- ---------------------------------------------------------------------
-- CONSUMO  (aba "Controle Alimentação")
-- Guarda só o fato: "comi ESTE item, NESTA refeição, NESTE dia/hora".
-- Os nutrientes vêm do item_plano (era o PROCV/VLOOKUP do Excel).
--   refeicao   = refeição em que você comeu (importa p/ o gráfico por refeição)
--   substituto = 1 quando comeu um item de outra refeição como substituição
--   porcoes    = 1 porção inteira; 0.5 = metade; 2 = o dobro
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS consumo (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    data           TEXT    NOT NULL,
    hora           TEXT    NOT NULL,
    refeicao       TEXT    NOT NULL REFERENCES meta_refeicao(refeicao),
    item_plano_id  INTEGER NOT NULL REFERENCES item_plano(id),
    substituto     INTEGER NOT NULL DEFAULT 0 CHECK (substituto IN (0, 1)),
    porcoes        REAL    NOT NULL DEFAULT 1 CHECK (porcoes > 0)
);
CREATE INDEX IF NOT EXISTS ix_consumo_data ON consumo(data);

-- ---------------------------------------------------------------------
-- MEDICAMENTOS  (aba "Medicamentos")
-- uso: 'Contínuo', 'Pontual' ou 'Com prazo'. Custo = preco * qtde (calculado).
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS medicamento (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    nome            TEXT NOT NULL,
    posologia       TEXT,
    preco           REAL,
    qtde            INTEGER DEFAULT 1,
    uso             TEXT NOT NULL DEFAULT 'Contínuo'
                    CHECK (uso IN ('Contínuo', 'Pontual', 'Com prazo')),
    data_inicio     TEXT,
    data_alteracao  TEXT
);

-- ---------------------------------------------------------------------
-- CARDÁPIO  (aba "Cardápio Semanal": a dieta prescrita)
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS cardapio_item (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    dieta           TEXT    NOT NULL,
    refeicao        TEXT    NOT NULL,
    ordem_item      INTEGER NOT NULL,
    descricao       TEXT    NOT NULL,
    peso_g          REAL,
    kcal            REAL,
    proteinas_g     REAL,
    carboidratos_g  REAL,
    gorduras_g      REAL
);

-- =====================================================================
-- VIEWS = consultas salvas com nome. Funcionam como "colunas calculadas".
-- =====================================================================

-- Consumo já com os nutrientes calculados (substitui o PROCV do Excel).
DROP VIEW IF EXISTS vw_consumo;
CREATE VIEW vw_consumo AS
SELECT  c.id,
        c.data,
        c.hora,
        c.refeicao,
        c.substituto,
        c.porcoes,
        i.descricao,
        i.peso_g         * c.porcoes AS peso_g,
        i.kcal           * c.porcoes AS kcal,
        i.proteinas_g    * c.porcoes AS proteinas_g,
        i.carboidratos_g * c.porcoes AS carboidratos_g,
        i.gorduras_g     * c.porcoes AS gorduras_g
FROM consumo c
JOIN item_plano i ON i.id = c.item_plano_id;

-- Pesagens com IMC calculado a partir da altura do perfil.
DROP VIEW IF EXISTS vw_pesagem;
CREATE VIEW vw_pesagem AS
SELECT  p.id,
        p.data,
        p.peso_kg,
        ROUND(p.peso_kg / (f.altura_m * f.altura_m), 2) AS imc
FROM pesagem p
CROSS JOIN perfil f;
