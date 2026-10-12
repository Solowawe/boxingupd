"""
База данных для приложения «Соревнования по боксу».

Два бэкенда:
- PostgreSQL — если задана переменная окружения DATABASE_URL (продакшен, Layero)
- SQLite     — fallback для локальной разработки (файл boxing.db)

Оба бэкенда доступны через единый sqlite3-подобный интерфейс:
    conn = get_db()
    cur = conn.execute("SELECT * FROM athletes WHERE id = ?", (1,))
    row = cur.fetchone()          # dict-строка: row['name'] и row[0]
    conn.commit()
"""

import os
import re
from pathlib import Path

USE_POSTGRES = bool(os.environ.get('DATABASE_URL'))

DB_PATH = Path(__file__).parent / 'boxing.db'

# Таблицы с serial-колонкой id — для них INSERT автоматически дополнится RETURNING id
# (нужно для эмуляции cursor.lastrowid в PostgreSQL)
SERIAL_ID_TABLES = {'athletes', 'tournaments', 'fights', 'lottery'}


# ======================== СТРОКИ-РЕЗУЛЬТАТЫ ========================

class DictRow(dict):
    """dict с дополнительным доступом по индексу (как sqlite3.Row)."""

    def __init__(self, keys, values):
        super().__init__(zip(keys, values))
        self._values = tuple(values)

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._values[key]
        return super().__getitem__(key)


# ======================== ПРЕОБРАЗОВАНИЕ SQL ДЛЯ POSTGRESQL ========================

def _convert_sql_for_pg(sql):
    """Конвертирует SQLite-диалект в PostgreSQL."""
    s = sql

    # INSERT OR IGNORE -> INSERT ... ON CONFLICT DO NOTHING
    if re.search(r'INSERT\s+OR\s+IGNORE\s+INTO', s, re.I):
        s = re.sub(r'INSERT\s+OR\s+IGNORE\s+INTO', 'INSERT INTO', s, flags=re.I)
        s = s.rstrip().rstrip(';') + ' ON CONFLICT DO NOTHING'

    # LIKE -> ILIKE (регистронезависимый поиск, как ожидает фронтенд)
    s = re.sub(r'\bLIKE\s', 'ILIKE ', s, flags=re.I)

    # Плейсхолдеры ? -> %s
    s = s.replace('?', '%s')

    return s


def _append_returning_id(pg_sql):
    """Добавляет RETURNING id к INSERT в таблицу с serial id."""
    if 'RETURNING' in pg_sql.upper():
        return pg_sql
    m = re.match(r'\s*INSERT\s+INTO\s+([a-zA-Z_][a-zA-Z0-9_]*)', pg_sql, re.I)
    if m and m.group(1).lower() in SERIAL_ID_TABLES:
        return pg_sql.rstrip().rstrip(';') + ' RETURNING id'
    return pg_sql


# ======================== ОБЁРТКИ КУРСОРА И СОЕДИНЕНИЯ ========================

class CursorWrapper:
    """Курсор с sqlite3-подобным поведением (fetchone/fetchall/lastrowid)."""

    def __init__(self, cur, use_pg):
        self._cur = cur
        self._use_pg = use_pg
        self._rows = None
        self._returned_id = None

    def _load(self):
        if self._rows is None:
            if self._cur.description is None:
                self._rows = []
            else:
                cols = [d[0] for d in self._cur.description]
                self._rows = [DictRow(cols, r) for r in self._cur.fetchall()]
        return self._rows

    def fetchone(self):
        rows = self._load()
        return rows[0] if rows else None

    def fetchall(self):
        return list(self._load())

    @property
    def lastrowid(self):
        if self._use_pg:
            return self._returned_id
        return self._cur.lastrowid


class ConnectionWrapper:
    """Соединение с sqlite3-подобным API: conn.execute(sql, params)."""

    def __init__(self, raw_conn, use_pg, release_fn=None):
        self._conn = raw_conn
        self._use_pg = use_pg
        self._release_fn = release_fn  # для возврата соединения в пул
        self._closed = False

    def execute(self, sql, params=()):
        if self._closed:
            raise RuntimeError('Соединение уже закрыто')

        use_sql = sql
        is_insert_returning = False
        if self._use_pg:
            use_sql = _convert_sql_for_pg(sql)
            if use_sql.strip().upper().startswith('INSERT INTO'):
                new_sql = _append_returning_id(use_sql)
                if new_sql != use_sql:
                    is_insert_returning = True
                use_sql = new_sql

        cur = self._conn.cursor()
        try:
            if params:
                cur.execute(use_sql, params)
            else:
                cur.execute(use_sql)
        except Exception:
            # В PostgreSQL ошибка прерывает транзакцию — откатываем,
            # иначе все последующие запросы упадут с InFailedSqlTransaction
            if self._use_pg:
                try:
                    self._conn.rollback()
                except Exception:
                    pass
            raise

        wrapper = CursorWrapper(cur, self._use_pg)

        # Для PG: забираем id из RETURNING (только для INSERT)
        if is_insert_returning and cur.description is not None and cur.description[0][0] == 'id':
            row = cur.fetchone()
            if row is not None:
                wrapper._returned_id = row[0]
                wrapper._rows = [DictRow(['id'], [row[0]])]

        return wrapper

    def commit(self):
        if not self._closed:
            self._conn.commit()

    def rollback(self):
        if not self._closed:
            self._conn.rollback()

    def close(self):
        """Идемпотентно: возвращает соединение в пул (PG) или закрывает (SQLite)."""
        if self._closed:
            return
        self._closed = True
        try:
            self._conn.rollback()  # сбросить незавершённую транзакцию
        except Exception:
            pass
        if self._release_fn is not None:
            self._release_fn(self._conn)
        else:
            try:
                self._conn.close()
            except Exception:
                pass


# ======================== ПУЛ СОЕДИНЕНИЙ POSTGRESQL ========================

_pg_pool = None


def _get_pg_pool():
    global _pg_pool
    if _pg_pool is None:
        import psycopg2.pool

        dsn = os.environ['DATABASE_URL']
        # Managed-провайдеры (YC, Neon, Supabase) обычно требуют SSL
        if 'sslmode=' not in dsn and 'localhost' not in dsn and '127.0.0.1' not in dsn:
            sep = '&' if '?' in dsn else '?'
            dsn = f'{dsn}{sep}sslmode=require'

        _pg_pool = psycopg2.pool.SimpleConnectionPool(1, 10, dsn)
    return _pg_pool


# ======================== ПУБЛИЧНЫЙ ИНТЕРФЕЙС ========================

def get_db():
    """Возвращает соединение с БД (sqlite3-подобный интерфейс).

    Внутри Flask-запроса соединение переиспользуется через g,
    вне контекста создаётся отдельное соединение.
    """
    from flask import has_app_context, g

    if USE_POSTGRES:
        pool = _get_pg_pool()
        if has_app_context():
            if 'db' not in g:
                g.db = ConnectionWrapper(pool.getconn(), use_pg=True,
                                         release_fn=lambda c: pool.putconn(c))
            return g.db
        return ConnectionWrapper(pool.getconn(), use_pg=True,
                                 release_fn=lambda c: pool.putconn(c))

    # SQLite fallback (локальная разработка)
    import sqlite3
    conn = sqlite3.connect(str(DB_PATH))
    # row_factory НЕ ставим: DictRow строит CursorWrapper._load единообразно
    conn.execute("PRAGMA foreign_keys=ON")
    return ConnectionWrapper(conn, use_pg=False)


def init_db():
    """Создаёт таблицы, если они ещё не существуют."""
    conn = get_db()
    cursor = conn._conn.cursor()

    ddl = '''
        CREATE TABLE IF NOT EXISTS athletes (
            id          SERIAL PRIMARY KEY,
            name        TEXT    NOT NULL,
            birth_year  INTEGER NOT NULL,
            weight      REAL    NOT NULL,
            gender      TEXT    NOT NULL CHECK(gender IN ('мужской', 'женский')),
            rank        TEXT    DEFAULT '',
            region      TEXT    DEFAULT '',
            coach       TEXT    DEFAULT '',
            created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS tournaments (
            id          SERIAL PRIMARY KEY,
            name        TEXT    NOT NULL,
            date_from   TEXT    NOT NULL,
            date_to     TEXT    NOT NULL,
            weight_step REAL    DEFAULT 5.0,
            days        INTEGER DEFAULT 1,
            completed   INTEGER DEFAULT 0,
            created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS fights (
            id            SERIAL PRIMARY KEY,
            tournament_id INTEGER NOT NULL,
            round         INTEGER NOT NULL,
            athlete1_id   INTEGER,
            athlete2_id   INTEGER,
            weight_class  TEXT    NOT NULL,
            winner_id     INTEGER,
            method        TEXT,
            completed     INTEGER DEFAULT 0,
            FOREIGN KEY (tournament_id) REFERENCES tournaments(id) ON DELETE CASCADE,
            FOREIGN KEY (athlete1_id)   REFERENCES athletes(id)    ON DELETE SET NULL,
            FOREIGN KEY (athlete2_id)   REFERENCES athletes(id)    ON DELETE SET NULL,
            FOREIGN KEY (winner_id)     REFERENCES athletes(id)    ON DELETE SET NULL
        );

        CREATE INDEX IF NOT EXISTS idx_fights_tournament ON fights(tournament_id);
        CREATE INDEX IF NOT EXISTS idx_fights_winner     ON fights(winner_id);
        CREATE INDEX IF NOT EXISTS idx_fights_athlete1   ON fights(athlete1_id);
        CREATE INDEX IF NOT EXISTS idx_fights_athlete2   ON fights(athlete2_id);

        CREATE TABLE IF NOT EXISTS tournament_athletes (
            tournament_id INTEGER NOT NULL,
            athlete_id    INTEGER NOT NULL,
            PRIMARY KEY (tournament_id, athlete_id),
            FOREIGN KEY (tournament_id) REFERENCES tournaments(id) ON DELETE CASCADE,
            FOREIGN KEY (athlete_id)    REFERENCES athletes(id)    ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS lottery (
            id            SERIAL PRIMARY KEY,
            tournament_id INTEGER NOT NULL,
            athlete_id    INTEGER NOT NULL,
            slot          INTEGER NOT NULL,
            weight_class  TEXT    NOT NULL,
            day           INTEGER DEFAULT 1,
            FOREIGN KEY (tournament_id) REFERENCES tournaments(id) ON DELETE CASCADE,
            FOREIGN KEY (athlete_id)    REFERENCES athletes(id)    ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS idx_lottery_tournament ON lottery(tournament_id);

        CREATE TABLE IF NOT EXISTS bracket_results (
            tournament_id INTEGER NOT NULL,
            weight_class  TEXT    NOT NULL,
            pair_index    INTEGER NOT NULL,
            winner_id     INTEGER NOT NULL,
            method        TEXT    DEFAULT '',
            day           INTEGER DEFAULT 1,
            PRIMARY KEY (tournament_id, weight_class, pair_index),
            FOREIGN KEY (tournament_id) REFERENCES tournaments(id) ON DELETE CASCADE,
            FOREIGN KEY (winner_id)     REFERENCES athletes(id)    ON DELETE CASCADE
        );
    '''

    if USE_POSTGRES:
        # psycopg2 умеет выполнять несколько выражений одним execute
        cursor.execute(ddl)

        # Миграция: колонка method в старых БД
        cursor.execute("""
            SELECT column_name FROM information_schema.columns
            WHERE table_name = 'bracket_results' AND column_name = 'method'
        """)
        if cursor.fetchone() is None:
            cursor.execute("ALTER TABLE bracket_results ADD COLUMN method TEXT DEFAULT ''")

        # Миграция: колонка day в старых БД (PostgreSQL)
        cursor.execute("""
            SELECT column_name FROM information_schema.columns
            WHERE table_name = 'bracket_results' AND column_name = 'day'
        """)
        if cursor.fetchone() is None:
            cursor.execute("ALTER TABLE bracket_results ADD COLUMN day INTEGER DEFAULT 1")

        # Миграция: колонка days в таблице tournaments (PostgreSQL)
        cursor.execute("""
            SELECT column_name FROM information_schema.columns
            WHERE table_name = 'tournaments' AND column_name = 'days'
        """)
        if cursor.fetchone() is None:
            cursor.execute("ALTER TABLE tournaments ADD COLUMN days INTEGER DEFAULT 1")

        # Миграция: колонка day в таблице lottery (PostgreSQL)
        cursor.execute("""
            SELECT column_name FROM information_schema.columns
            WHERE table_name = 'lottery' AND column_name = 'day'
        """)
        if cursor.fetchone() is None:
            cursor.execute("ALTER TABLE lottery ADD COLUMN day INTEGER DEFAULT 1")
    else:
        # SQLite: AUTOINCREMENT вместо SERIAL
        conn._conn.executescript(ddl.replace('SERIAL PRIMARY KEY', 'INTEGER PRIMARY KEY AUTOINCREMENT'))

        # Миграция: колонка method в старых БД
        cols = [r[1] for r in conn._conn.execute("PRAGMA table_info(bracket_results)").fetchall()]
        if 'method' not in cols:
            conn._conn.execute("ALTER TABLE bracket_results ADD COLUMN method TEXT DEFAULT ''")

        # Миграция: колонка day в старых БД
        if 'day' not in cols:
            conn._conn.execute("ALTER TABLE bracket_results ADD COLUMN day INTEGER DEFAULT 1")

        # Миграция: колонка days в таблице tournaments
        tcols = [r[1] for r in conn._conn.execute("PRAGMA table_info(tournaments)").fetchall()]
        if 'days' not in tcols:
            conn._conn.execute("ALTER TABLE tournaments ADD COLUMN days INTEGER DEFAULT 1")

        # Миграция: колонка day в таблице lottery (SQLite)
        lcols = [r[1] for r in conn._conn.execute("PRAGMA table_info(lottery)").fetchall()]
        if 'day' not in lcols:
            conn._conn.execute("ALTER TABLE lottery ADD COLUMN day INTEGER DEFAULT 1")

    conn.commit()
    conn.close()


if __name__ == '__main__':
    init_db()
    backend = 'PostgreSQL' if USE_POSTGRES else f'SQLite ({DB_PATH})'
    print(f"База данных инициализирована: {backend}")
