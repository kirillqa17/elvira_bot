import sqlite3
import uuid
from datetime import datetime

# --- Настройка базы данных ---
DB_NAME = "bot.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    # Пользователи
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER UNIQUE,
            username TEXT,
            first_name TEXT,
            date_registered TEXT,
            ref_marker TEXT,
            current_stage TEXT
        )
    """)
    # Маркеры
    cur.execute("""
        CREATE TABLE IF NOT EXISTS markers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            marker TEXT UNIQUE,
            created_at TEXT,
            users_total INTEGER DEFAULT 0
        )
    """)
    conn.commit()
    conn.close()



def add_or_update_user(user, marker=None):
    """
    Добавляет нового пользователя или обновляет данные существующего.
    Если пользователь новый и пришёл с маркером — увеличиваем users_total для этого маркера.
    Если пользователь был, но ref_marker пустой — заполняем и увеличиваем счётчик.
    Если у пользователя уже есть ref_marker — не трогаем (чтобы не портить статистику).
    """
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()

    cur.execute("SELECT id, ref_marker FROM users WHERE telegram_id = ?", (user.id,))
    row = cur.fetchone()

    username = getattr(user, 'username', None)
    first_name = getattr(user, 'first_name', None)

    if row:
        # пользователь существует
        user_id, existing_ref = row
        cur.execute("UPDATE users SET username = ?, first_name = ? WHERE telegram_id = ?", (username, first_name, user.id))
        # если ранее не было маркера, и сейчас пришёл маркер — ставим его и обновляем счётчик
        if (existing_ref is None or existing_ref == '') and marker:
            cur.execute("UPDATE users SET ref_marker = ? WHERE telegram_id = ?", (marker, user.id))
            increment_marker_count(marker, conn=conn)
    else:
        initial_stage = "start"
        # новый пользователь
        cur.execute("INSERT INTO users (telegram_id, username, first_name, date_registered, ref_marker, current_stage) VALUES (?,?,?,?,?,?)",
                    (user.id, username, first_name, datetime.now().isoformat(), marker, initial_stage))
        if marker:
            increment_marker_count(marker, conn=conn)

    conn.commit()
    conn.close()

# Новая функция для обновления этапа пользователя
def update_user_stage(telegram_id, stage):
    """Обновляет текущий этап пользователя в воронке."""
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("UPDATE users SET current_stage = ? WHERE telegram_id = ?", (stage, telegram_id))
    conn.commit()
    conn.close()

def get_user_stage(telegram_id):
    """Получает текущий этап пользователя в воронке."""
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT current_stage from users WHERE telegram_id = ?", (telegram_id))
    cur_stage = cur.fetchall()
    conn.close()
    return cur_stage

def get_username(telegram_id):
    """Получает username пользователя в воронке."""
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT username from users WHERE telegram_id = ?", (telegram_id))
    username = cur.fetchall()
    conn.close()
    return username

def create_marker(name):
    marker = uuid.uuid4().hex[:8]
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("INSERT INTO markers (name, marker, created_at, users_total) VALUES (?,?,?,0)",
                (name, marker, datetime.now().isoformat()))
    conn.commit()
    conn.close()
    return marker

def get_markers():
    """
    Возвращает список маркеров: (id, name, marker, created_at, users_total)
    """
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT id, name, marker, created_at, COALESCE(users_total, 0) FROM markers ORDER BY id DESC")
    rows = cur.fetchall()
    conn.close()
    return rows

def get_all_users():
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT id, telegram_id, username, first_name, date_registered, ref_marker, current_stage FROM users ORDER BY id DESC")
    rows = cur.fetchall()
    conn.close()
    return rows

def get_users_count_by_marker(marker):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM users WHERE ref_marker=?", (marker,))
    count = cur.fetchone()[0]
    conn.close()
    return count

def increment_marker_count(marker, conn=None):
    """
    Увеличивает users_total для маркера на 1.
    Если маркера нет — ничего не делает.
    """
    need_close = False
    if conn is None:
        conn = sqlite3.connect(DB_NAME)
        need_close = True
    cur = conn.cursor()
    cur.execute("UPDATE markers SET users_total = COALESCE(users_total, 0) + 1 WHERE marker = ?", (marker,))
    if need_close:
        conn.commit()
        conn.close()
    else:
        conn.commit()

def update_all_markers_counts():
    """
    Пересчитывает users_total для всех маркеров по фактическим данным в users.
    Удобно вызывать после миграции/изменений.
    """
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT marker FROM markers")
    markers = [r[0] for r in cur.fetchall()]
    for m in markers:
        cur.execute("SELECT COUNT(*) FROM users WHERE ref_marker = ?", (m,))
        cnt = cur.fetchone()[0]
        cur.execute("UPDATE markers SET users_total = ? WHERE marker = ?", (cnt, m))
    conn.commit()
    conn.close()

