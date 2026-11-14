"""
database.py - Tudo que fala diretamente com o SQLite fica aqui.

Ideia: o resto do sistema nunca abre o banco "na mão". Ele chama
consulta() para LER e executa() para GRAVAR. Se um dia você trocar o
SQLite por PostgreSQL, só este arquivo muda.
"""
import sqlite3
from pathlib import Path

import pandas as pd

PASTA = Path(__file__).resolve().parent
ARQUIVO_DB = PASTA / "controle_peso.db"
ARQUIVO_SCHEMA = PASTA / "schema.sql"


def conectar() -> sqlite3.Connection:
    con = sqlite3.connect(ARQUIVO_DB)
    con.execute("PRAGMA foreign_keys = ON")  # o SQLite exige ligar isso a cada conexão
    return con


def criar_estrutura() -> None:
    """Cria as tabelas e views (lendo o schema.sql). Pode rodar várias vezes."""
    with conectar() as con:
        con.executescript(ARQUIVO_SCHEMA.read_text(encoding="utf-8"))


def banco_existe() -> bool:
    return ARQUIVO_DB.exists()


def consulta(sql: str, params: tuple = ()) -> pd.DataFrame:
    """Executa um SELECT e devolve um DataFrame do pandas."""
    con = conectar()
    try:
        return pd.read_sql_query(sql, con, params=params)
    finally:
        con.close()


def executa(sql: str, params: tuple = ()) -> int:
    """Executa INSERT/UPDATE/DELETE. Devolve o id da última linha inserida."""
    con = conectar()
    try:
        with con:  # 'with' faz COMMIT se deu certo e ROLLBACK se deu erro
            cur = con.execute(sql, params)
            return cur.lastrowid
    finally:
        con.close()
