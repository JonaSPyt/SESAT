"""
Módulo de banco de dados SQLite para persistência dos registros de equipamentos.
"""
import sqlite3
import os
import sys
import hashlib
import configparser
from datetime import datetime, timedelta


def _parse_data_para_ordenacao(valor: str | None) -> datetime | None:
    """Converte datas DD/MM/AA ou DD/MM/AAAA para datetime para ordenação."""
    if not valor:
        return None

    texto = valor.strip()
    for formato in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto, formato)
        except ValueError:
            continue
    return None


def _ordenar_registros(registros: list[dict], ordenacao: str) -> list[dict]:
    """Ordena registros em memória para tratar datas armazenadas como texto."""
    if ordenacao == "mais_antigos":
        return sorted(
            registros,
            key=lambda reg: (
                _parse_data_para_ordenacao(reg.get("data_entrada")) is None,
                _parse_data_para_ordenacao(reg.get("data_entrada")) or datetime.max,
                reg.get("id", 0),
            ),
        )

    if ordenacao == "saida_mais_recentes":
        return sorted(
            registros,
            key=lambda reg: (
                _parse_data_para_ordenacao(reg.get("data_saida")) or datetime.min,
                reg.get("id", 0),
            ),
            reverse=True,
        )

    if ordenacao == "saida_mais_antigos":
        return sorted(
            registros,
            key=lambda reg: (
                _parse_data_para_ordenacao(reg.get("data_saida")) is None,
                _parse_data_para_ordenacao(reg.get("data_saida")) or datetime.max,
                reg.get("id", 0),
            ),
        )

    if ordenacao == "tombamento_az":
        return sorted(
            registros,
            key=lambda reg: (
                str(reg.get("tombamento", "")).strip().lower(),
                reg.get("id", 0),
            ),
        )

    if ordenacao == "tombamento_za":
        return sorted(
            registros,
            key=lambda reg: (
                str(reg.get("tombamento", "")).strip().lower(),
                reg.get("id", 0),
            ),
            reverse=True,
        )

    return sorted(
        registros,
        key=lambda reg: (
            _parse_data_para_ordenacao(reg.get("data_entrada")) or datetime.min,
            reg.get("id", 0),
        ),
        reverse=True,
    )


def _get_app_dir():
    """Retorna o diretório real do executável (não a pasta temp do PyInstaller)."""
    if getattr(sys, 'frozen', False):
        # Executando como .exe empacotado pelo PyInstaller
        return os.path.dirname(sys.executable)
    # Executando como script .py normal
    return os.path.dirname(os.path.abspath(__file__))


def _get_db_path():
    """
    Determina o caminho do banco de dados.
    Se existir um config.ini com [database] caminho_rede, usa a pasta da rede.
    Caso contrário, usa o diretório local do executável.
    """
    app_dir = _get_app_dir()
    config_path = os.path.join(app_dir, "config.ini")

    if os.path.exists(config_path):
        config = configparser.ConfigParser()
        config.read(config_path, encoding="utf-8")
        caminho_rede = config.get(
            "database", "caminho_rede", fallback="").strip()
        if caminho_rede:
            # Criar a pasta da rede se não existir
            try:
                os.makedirs(caminho_rede, exist_ok=True)
            except OSError:
                pass
            db = os.path.join(caminho_rede, "sesat.db")
            # Verificar se é acessível
            try:
                if os.path.isdir(caminho_rede) or os.access(
                        os.path.dirname(caminho_rede), os.W_OK):
                    return db
            except OSError:
                pass

    return os.path.join(app_dir, "sesat.db")


DB_PATH = _get_db_path()


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    """Cria as tabelas caso não existam."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS equipamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_entrada TEXT NOT NULL,
            secao_zona TEXT,
            nome_setor TEXT,
            local_setor TEXT,
            tombamento TEXT NOT NULL,
            tipo TEXT,
            descricao_completa TEXT,
            nome_responsavel TEXT,
            valor_unitario TEXT,
            num_chamado TEXT,
            tecnico_opr_entrada TEXT,
            tecnico_sesat TEXT,
            tecnico_opr_devolucao TEXT,
            data_saida TEXT,
            observacoes TEXT,
            anotacoes TEXT,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_hora TEXT NOT NULL,
            usuario TEXT NOT NULL,
            acao TEXT NOT NULL,
            tombamento TEXT,
            data_entrada TEXT,
            data_saida TEXT,
            detalhes TEXT,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL UNIQUE,
            senha_hash TEXT NOT NULL,
            is_super INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        )
    """)
    # Criar superusuário padrão se não existir
    cursor.execute(
        "SELECT COUNT(*) FROM usuarios WHERE usuario = 'Supervisor'")
    if cursor.fetchone()[0] == 0:
        senha_hash = _hash_senha("Sesat2026")
        cursor.execute(
            "INSERT INTO usuarios (usuario, senha_hash, is_super) VALUES (?, ?, 1)",
            ("Supervisor", senha_hash)
        )
    # Criar usuário Consultor (somente leitura) se não existir
    cursor.execute(
        "SELECT COUNT(*) FROM usuarios WHERE usuario = 'Consultor'")
    if cursor.fetchone()[0] == 0:
        senha_hash = _hash_senha("Tre2026")
        cursor.execute(
            "INSERT INTO usuarios (usuario, senha_hash, is_super) VALUES (?, ?, 0)",
            ("Consultor", senha_hash)
        )
    conn.commit()
    conn.close()


def inserir_equipamento(dados: dict) -> int:
    """Insere um novo registro e retorna o ID."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO equipamentos (
            data_entrada, secao_zona, nome_setor, local_setor, tombamento,
            tipo, descricao_completa, nome_responsavel, valor_unitario,
            num_chamado, tecnico_opr_entrada, tecnico_sesat,
            tecnico_opr_devolucao, data_saida, observacoes, anotacoes
        ) VALUES (
            :data_entrada, :secao_zona, :nome_setor, :local_setor, :tombamento,
            :tipo, :descricao_completa, :nome_responsavel, :valor_unitario,
            :num_chamado, :tecnico_opr_entrada, :tecnico_sesat,
            :tecnico_opr_devolucao, :data_saida, :observacoes, :anotacoes
        )
    """, dados)
    conn.commit()
    row_id = cursor.lastrowid
    conn.close()
    return row_id


def inserir_equipamentos_batch(lista_dados: list) -> int:
    """Insere vários registros em uma única transação. Retorna o total inserido."""
    if not lista_dados:
        return 0
    conn = get_connection()
    cursor = conn.cursor()
    sql = """
        INSERT INTO equipamentos (
            data_entrada, secao_zona, nome_setor, local_setor, tombamento,
            tipo, descricao_completa, nome_responsavel, valor_unitario,
            num_chamado, tecnico_opr_entrada, tecnico_sesat,
            tecnico_opr_devolucao, data_saida, observacoes, anotacoes
        ) VALUES (
            :data_entrada, :secao_zona, :nome_setor, :local_setor, :tombamento,
            :tipo, :descricao_completa, :nome_responsavel, :valor_unitario,
            :num_chamado, :tecnico_opr_entrada, :tecnico_sesat,
            :tecnico_opr_devolucao, :data_saida, :observacoes, :anotacoes
        )
    """
    cursor.executemany(sql, lista_dados)
    conn.commit()
    total = cursor.rowcount
    conn.close()
    return total


def buscar_chaves_existentes() -> set:
    """Retorna set de (tombamento, data_entrada) de todos os registros."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT tombamento, data_entrada FROM equipamentos")
    chaves = {(row[0] or "", row[1] or "") for row in cursor.fetchall()}
    conn.close()
    return chaves


def atualizar_equipamento(equip_id: int, dados: dict):
    """Atualiza um registro existente pelo ID."""
    conn = get_connection()
    cursor = conn.cursor()
    dados["id"] = equip_id
    dados["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
        UPDATE equipamentos SET
            data_entrada = :data_entrada,
            secao_zona = :secao_zona,
            nome_setor = :nome_setor,
            local_setor = :local_setor,
            tombamento = :tombamento,
            tipo = :tipo,
            descricao_completa = :descricao_completa,
            nome_responsavel = :nome_responsavel,
            valor_unitario = :valor_unitario,
            num_chamado = :num_chamado,
            tecnico_opr_entrada = :tecnico_opr_entrada,
            tecnico_sesat = :tecnico_sesat,
            tecnico_opr_devolucao = :tecnico_opr_devolucao,
            data_saida = :data_saida,
            observacoes = :observacoes,
            anotacoes = :anotacoes,
            updated_at = :updated_at
        WHERE id = :id
    """, dados)
    conn.commit()
    conn.close()


def deletar_equipamento(equip_id: int):
    """Remove um registro pelo ID."""
    conn = get_connection()
    conn.execute("DELETE FROM equipamentos WHERE id = ?", (equip_id,))
    conn.commit()
    conn.close()


def buscar_todos(filtro_texto: str = "", filtro_campo: str = "todos",
                 ordenacao: str = "mais_recentes") -> list:
    """Busca registros com filtro opcional."""
    conn = get_connection()
    cursor = conn.cursor()

    if filtro_texto.strip():
        like = f"%{filtro_texto.strip()}%"
        if filtro_campo == "todos":
            cursor.execute("""
                SELECT * FROM equipamentos
                WHERE tombamento LIKE ? OR secao_zona LIKE ? OR tipo LIKE ?
                    OR num_chamado LIKE ? OR tecnico_opr_entrada LIKE ?
                    OR tecnico_sesat LIKE ? OR tecnico_opr_devolucao LIKE ?
                    OR observacoes LIKE ? OR anotacoes LIKE ?
                    OR nome_setor LIKE ? OR descricao_completa LIKE ?
            """, (like,) * 11)
        else:
            campo_map = {
                "tombamento": "tombamento",
                "zona": "secao_zona",
                "tipo": "tipo",
                "chamado": "num_chamado",
                "opr_entrada": "tecnico_opr_entrada",
                "sesat": "tecnico_sesat",
                "opr_devolucao": "tecnico_opr_devolucao",
                "observacoes": "observacoes",
            }
            campo = campo_map.get(filtro_campo, "tombamento")
            cursor.execute(f"""
                SELECT * FROM equipamentos
                WHERE {campo} LIKE ?
            """, (like,))
    else:
        cursor.execute("""
            SELECT * FROM equipamentos
        """)

    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return _ordenar_registros(rows, ordenacao)


def buscar_por_id(equip_id: int) -> dict | None:
    """Busca um registro pelo ID."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM equipamentos WHERE id = ?", (equip_id,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


# ═══════════════════════════════════════════════════════════════════
#  Funções de log
# ═══════════════════════════════════════════════════════════════════

def registrar_log(usuario: str, acao: str, tombamento: str = "",
                  data_entrada: str = "", data_saida: str = "",
                  detalhes: str = ""):
    """Insere um registro de log e executa a limpeza rolling de 6 meses."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO logs (data_hora, usuario, acao, tombamento,
                          data_entrada, data_saida, detalhes)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        usuario, acao, tombamento,
        data_entrada, data_saida, detalhes
    ))
    conn.commit()
    conn.close()
    # Limpeza rolling após cada inserção
    limpar_logs_antigos()


def limpar_logs_antigos():
    """Remove logs com mais de 6 meses (180 dias) — sistema de fila rolling."""
    conn = get_connection()
    limite = (datetime.now() - timedelta(days=180)
              ).strftime("%Y-%m-%d %H:%M:%S")
    # A coluna created_at armazena no formato ISO — usamos ela para comparar
    conn.execute("DELETE FROM logs WHERE created_at < ?", (limite,))
    conn.commit()
    conn.close()


def buscar_logs(filtro_texto: str = "") -> list:
    """Busca logs com filtro opcional."""
    conn = get_connection()
    cursor = conn.cursor()
    if filtro_texto.strip():
        like = f"%{filtro_texto.strip()}%"
        cursor.execute("""
            SELECT * FROM logs
            WHERE usuario LIKE ? OR acao LIKE ? OR tombamento LIKE ?
                  OR detalhes LIKE ? OR data_hora LIKE ?
            ORDER BY id DESC
        """, (like, like, like, like, like))
    else:
        cursor.execute("SELECT * FROM logs ORDER BY id DESC")
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


# ═══════════════════════════════════════════════════════════════════
#  Funções de usuários
# ═══════════════════════════════════════════════════════════════════

def _hash_senha(senha: str) -> str:
    """Retorna o hash SHA-256 da senha."""
    return hashlib.sha256(senha.encode("utf-8")).hexdigest()


def autenticar_usuario(usuario: str, senha: str) -> dict | None:
    """Autentica um usuário. Retorna dict com dados ou None se falhar."""
    conn = get_connection()
    cursor = conn.cursor()
    senha_hash = _hash_senha(senha)
    cursor.execute(
        "SELECT * FROM usuarios WHERE usuario = ? AND senha_hash = ?",
        (usuario, senha_hash)
    )
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def criar_usuario(usuario: str, senha: str) -> bool:
    """Cria um novo usuário comum. Retorna True se criado com sucesso."""
    conn = get_connection()
    cursor = conn.cursor()
    senha_hash = _hash_senha(senha)
    try:
        cursor.execute(
            "INSERT INTO usuarios (usuario, senha_hash, is_super) VALUES (?, ?, 0)",
            (usuario, senha_hash)
        )
        conn.commit()
        conn.close()
        return True
    except sqlite3.IntegrityError:
        conn.close()
        return False


def listar_usuarios() -> list:
    """Lista todos os usuários cadastrados."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, usuario, is_super, created_at FROM usuarios ORDER BY usuario")
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def excluir_usuario(user_id: int) -> bool:
    """Exclui um usuário pelo ID (não permite excluir superusuário)."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT is_super FROM usuarios WHERE id = ?", (user_id,))
    row = cursor.fetchone()
    if not row or row["is_super"] == 1:
        conn.close()
        return False
    cursor.execute("DELETE FROM usuarios WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()
    return True


def alterar_senha_usuario(user_id: int, nova_senha: str):
    """Altera a senha de um usuário."""
    conn = get_connection()
    senha_hash = _hash_senha(nova_senha)
    conn.execute(
        "UPDATE usuarios SET senha_hash = ? WHERE id = ?",
        (senha_hash, user_id)
    )
    conn.commit()
    conn.close()


# Inicializar banco ao importar
init_db()
