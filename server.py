# -*- coding: utf-8 -*-
"""
Flask API server for Boxing Competition Tracker
"""

import sys
import io

# Безопасная настройка UTF-8 для вывода (может не сработать под gunicorn — не критично)
try:
    if hasattr(sys.stdout, 'buffer'):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    if hasattr(sys.stderr, 'buffer'):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
except Exception:
    pass

import collections
import json
import logging
import math
import random
import secrets
import shutil
import threading
import time
import urllib.request
import webbrowser
import socket
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, send_from_directory, g, session, redirect, url_for
from flask_cors import CORS

from database import get_db, init_db

app = Flask(__name__, static_folder='.')
CORS(app)

# ======================== АВТОРИЗАЦИЯ ========================
# Безопасность: секретный ключ и админ-доступ задаются переменными окружения.
# Если SECRET_KEY не задан — генерируем случайный (сессии сбрасываются при рестарте).
import os
DEFAULT_SECRET = 'boxing-tracker-secret-key-change-me'
RANDOM_SECRET = secrets.token_hex(32)
app.secret_key = os.environ.get('SECRET_KEY', RANDOM_SECRET)

ADMIN_LOGIN = os.environ.get('ADMIN_LOGIN', 'admin')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD')

# Cookie-флаги сессии (HttpOnly, SameSite, Secure-только при HTTPS)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=os.environ.get('SECURE_COOKIES', '') == '1',
)

# Безопасность cookie по умолчанию выключена, чтобы работало на обычном HTTP локально.
if ADMIN_PASSWORD in (None, '') or ADMIN_PASSWORD == 'admin':
    ADMIN_PASSWORD = ADMIN_PASSWORD or 'admin'
    log_warning = logging.getLogger('boxing')
    _s = 'Установите переменные окружения ADMIN_PASSWORD и SECRET_KEY!'
    print(f'[WARN] Пароль администратора по умолчанию ("admin"). {_s}')
    logging.getLogger('boxing').warning('Используется пароль администратора по умолчанию. %s', _s)

# Брутфорс-защита: ограничение попыток входа с одного IP.
_login_attempts = collections.defaultdict(list)
LOGIN_MAX_ATTEMPTS = int(os.environ.get('LOGIN_MAX_ATTEMPTS', '10'))
LOGIN_WINDOW_SECONDS = 300  # 5 минут

# ======================== ОБНОВЛЕНИЯ ИЗ GIT (GITHUB) ========================
# Обновляются только файлы кода. Данные (boxing.db) НЕ трогаются.
# Задайте переменные окружения (в server.py или bat):
#   GITHUB_OWNER, GITHUB_REPO, GITHUB_BRANCH (по умолчанию 'main')
# В репозитории должен лежать файл update-manifest.json вида:
#   { "version": "1.0.3", "files": ["app.js","index.html","style.css","server.py","database.py","app.py","login.html"] }
UPDATE_OWNER = 'Solowawe'  
UPDATE_REPO = 'boxingupd'
UPDATE_BRANCH = os.environ.get('GITHUB_BRANCH', 'main').strip() or 'main'
UPDATE_MANIFEST = 'update-manifest.json'
LOCAL_VERSION_FILE = 'local_version.txt'
# Только эти файлы разрешено обновлять (защита от произвольной записи)
UPDATE_ALLOWED = {
    'app.js', 'index.html', 'style.css', 'login.html',
    'server.py', 'database.py', 'app.py', 'requirements.txt'
}


def _fetch_url(url, timeout=25):
    req = urllib.request.Request(url, headers={'User-Agent': 'boxing-updater/1.0'})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _read_local_version():
    try:
        with open(LOCAL_VERSION_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except Exception:
        return ''


def _write_local_version(v):
    try:
        with open(LOCAL_VERSION_FILE, 'w', encoding='utf-8') as f:
            f.write(str(v).strip())
    except Exception:
        pass


def _login_throttled(client_ip):
    """Возвращает True, если с этого IP превышен лимит попыток входа."""
    now = time.time()
    attempts = [t for t in _login_attempts.get(client_ip, []) if now - t < LOGIN_WINDOW_SECONDS]
    _login_attempts[client_ip] = attempts
    return len(attempts) >= LOGIN_MAX_ATTEMPTS


def _record_login_failure(client_ip):
    _login_attempts[client_ip].append(time.time())
    _login_attempts[client_ip] = [t for t in _login_attempts[client_ip]
                                  if time.time() - t < LOGIN_WINDOW_SECONDS]

# Пути, доступные без авторизации
PUBLIC_PATHS = ('/login', '/logout', '/style.css', '/login.html')


@app.route('/login', methods=['GET'])
def login_page():
    """Страница входа."""
    if session.get('authenticated'):
        return redirect('/')
    return send_from_directory('.', 'login.html')


@app.route('/login', methods=['POST'])
def login_submit():
    """Проверка логина/пароля."""
    client_ip = request.remote_addr or 'unknown'
    if _login_throttled(client_ip):
        return jsonify({'success': False, 'message': 'Слишком много попыток входа. Подождите и попробуйте снова.'}), 429

    data = request.get_json(silent=True) or request.form
    login = (data.get('login') or '').strip()
    password = (data.get('password') or '').strip()

    if login == ADMIN_LOGIN and password == ADMIN_PASSWORD:
        session['authenticated'] = True
        session['user'] = login
        return jsonify({'success': True, 'message': 'Вход выполнен'})

    _record_login_failure(client_ip)
    return jsonify({'success': False, 'message': 'Неверный логин или пароль'}), 401


@app.route('/logout', methods=['POST', 'GET'])
def logout():
    """Выход."""
    session.clear()
    return redirect('/login')

# ======================== ЛОГИРОВАНИЕ ========================
logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')
log = logging.getLogger('boxing')

# ======================== БАЗА ДАННЫХ ========================

@app.before_request
def before_request():
    """Проверка авторизации + соединение с БД перед каждым запросом."""
    # Проверяем авторизацию (кроме публичных путей)
    path = request.path
    is_public = (
        path in PUBLIC_PATHS
        or path.startswith('/login')
        or path.startswith('/logout')
        or path == '/favicon.ico'
    )
    if not is_public and not session.get('authenticated'):
        # Для API возвращаем 401, для страниц — редирект на форму входа
        if path.startswith('/api/'):
            return jsonify({'error': 'Требуется авторизация'}), 401
        return redirect('/login')

    g.db = get_db()


@app.teardown_appcontext
def teardown_db(exception):
    """Закрываем соединение после каждого запроса."""
    db = getattr(g, 'db', None)
    if db is not None:
        db.close()


# ======================== РЕЗЕРВНОЕ КОПИРОВАНИЕ ========================

DB_FILE = 'boxing.db'
BACKUP_DIR = 'backups'
_last_backup_time = 0.0
BACKUP_INTERVAL_SECONDS = 86400  # раз в сутки


def backup_db(force=False):
    """Копирует boxing.db в папку backups с отметкой времени.

    Вызывается при старте и раз в сутки. Безопасен, если файла нет.
    """
    global _last_backup_time
    now = time.time()
    if not force and (now - _last_backup_time < BACKUP_INTERVAL_SECONDS):
        return None

    db_path = Path(DB_FILE)
    if not db_path.exists():
        return None

    try:
        backup_path = Path(BACKUP_DIR)
        backup_path.mkdir(exist_ok=True)
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        target = backup_path / f'boxing_{stamp}.db'
        shutil.copy2(db_path, target)
        _last_backup_time = now
        log.info('Резервная копия БД создана: %s', target)
        return str(target)
    except Exception as e:
        log.error('Не удалось создать резервную копию БД: %s', e)
        return None

# ======================== ВСПОМОГАТЕЛЬНР«Р• ========================

def row_to_dict(row):
    """Преобразует sqlite3.Row РІ dict."""
    if row is None:
        return None
    return dict(row)


def rows_to_list(rows):
    """Преобразует список sqlite3.Row в список dict."""
    if rows is None:
        return []
    return [dict(r) for r in rows]


def format_weight_class(gender, weight, step):
    """Форматирует строку весовой категории (целые числа в границах)."""
    lower = int(weight / step) * step
    upper = int(lower + step - 1)
    lower_str = f"{int(lower)}" if lower == int(lower) else f"{lower}"
    return f"{gender} {lower_str}-{upper} кг"


def rows_to_list(rows):
    """Преобразует список sqlite3.Row в список dict."""
    return [dict(r) for r in rows]


# ======================== API: СПОРТСМЕНР« ========================

@app.route('/api/athletes', methods=['GET'])
def get_athletes():
    """Получить всех спортсменов. Поддерживает фильтрацию: ?gender=&search=&weight_min=&weight_max="""
    conn = get_db()
    query = "SELECT * FROM athletes WHERE 1=1"
    params = []

    gender = request.args.get('gender')
    if gender:
        query += " AND gender = ?"
        params.append(gender)

    search = request.args.get('search')
    if search:
        log.info(f"DEBUG search={search!r}")
        query += " AND name LIKE ?"
        params.append(f'%{search}%')

    weight_min = request.args.get('weight_min')
    if weight_min:
        query += " AND weight >= ?"
        params.append(float(weight_min))

    weight_max = request.args.get('weight_max')
    if weight_max:
        query += " AND weight <= ?"
        params.append(float(weight_max))

    query += " ORDER BY name"
    athletes = rows_to_list(conn.execute(query, params).fetchall())
    conn.close()
    return jsonify(athletes)


@app.route('/api/athletes/<int:athlete_id>', methods=['GET'])
def get_athlete(athlete_id):
    """Получить одного спортсмена."""
    conn = get_db()
    athlete = row_to_dict(conn.execute("SELECT * FROM athletes WHERE id = ?", (athlete_id,)).fetchone())
    conn.close()
    if athlete is None:
        return jsonify({'error': 'Спортсмен не нР°Р№РґРµРЅ'}), 404
    return jsonify(athlete)


@app.route('/api/athletes', methods=['POST'])
def add_athlete():
    """Добавить нового спортсмена."""
    data = request.get_json()
    required = ['name', 'birth_year', 'weight', 'gender']
    for field in required:
        if field not in data:
            return jsonify({'error': f'ПоР»е "{field}" обязательно'}), 400

    if data['gender'] not in ('мужской', 'женский'):
        return jsonify({'error': 'Пол должен быть "мужской" или "женский"'}), 400

    conn = get_db()
    cursor = conn.execute(
        '''INSERT INTO athletes (name, birth_year, weight, gender, rank, region, coach)
           VALUES (?, ?, ?, ?, ?, ?, ?)''',
        (
            data['name'].strip(),
            int(data['birth_year']),
            float(data['weight']),
            data['gender'],
            data.get('rank', '').strip(),
            data.get('region', '').strip(),
            data.get('coach', '').strip()
        )
    )
    athlete_id = cursor.lastrowid
    conn.commit()
    athlete = row_to_dict(conn.execute("SELECT * FROM athletes WHERE id = ?", (athlete_id,)).fetchone())
    conn.close()
    return jsonify(athlete), 201


@app.route('/api/athletes/<int:athlete_id>', methods=['PUT'])
def update_athlete(athlete_id):
    """Обновить данные спортсмена."""
    data = request.get_json()
    conn = get_db()

    # Проверяем, что спортсмен существует
    existing = conn.execute("SELECT * FROM athletes WHERE id = ?", (athlete_id,)).fetchone()
    if existing is None:
        conn.close()
        return jsonify({'error': 'Спортсмен не нР°Р№РґРµРЅ'}), 404

    fields = []
    params = []
    for field in ['name', 'birth_year', 'weight', 'gender', 'rank', 'region', 'coach']:
        if field in data:
            fields.append(f"{field} = ?")
            params.append(data[field])

    if not fields:
        conn.close()
        return jsonify({'error': 'Нет данных для обновления'}), 400

    params.append(athlete_id)
    conn.execute(f"UPDATE athletes SET {', '.join(fields)} WHERE id = ?", params)
    conn.commit()
    athlete = row_to_dict(conn.execute("SELECT * FROM athletes WHERE id = ?", (athlete_id,)).fetchone())
    conn.close()
    return jsonify(athlete)


@app.route('/api/athletes/<int:athlete_id>', methods=['DELETE'])
def delete_athlete(athlete_id):
    """Удалить спортсмена."""
    conn = get_db()
    conn.execute("DELETE FROM athletes WHERE id = ?", (athlete_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})


# ======================== API: СОРЕВНОВАНРРЇ ========================

@app.route('/api/tournaments', methods=['GET'])
def get_tournaments():
    """Получить все соревнования."""
    conn = get_db()
    tournaments = rows_to_list(conn.execute("SELECT * FROM tournaments ORDER BY date_from DESC").fetchall())
    # Добавляем информацию о количестве боёв
    for t in tournaments:
        row = conn.execute(
            "SELECT COUNT(*) as total, SUM(CASE WHEN completed THEN 1 ELSE 0 END) as completed FROM fights WHERE tournament_id = ?",
            (t['id'],)
        ).fetchone()
        t['total_fights'] = row['total']
        t['completed_fights'] = row['completed'] or 0
        t['athlete_count'] = conn.execute(
            "SELECT COUNT(*) FROM tournament_athletes WHERE tournament_id = ?",
            (t['id'],)
        ).fetchone()[0]
    conn.close()
    return jsonify(tournaments)


@app.route('/api/tournaments', methods=['POST'])
def add_tournament():
    """Создать новое соревнование."""
    data = request.get_json()
    if not data.get('name') or not data.get('date_from'):
        return jsonify({'error': 'Поля "name", "date_from" обязательны'}), 400

    conn = get_db()
    cursor = conn.execute(
        "INSERT INTO tournaments (name, date_from, date_to, weight_step, days) VALUES (?, ?, ?, ?, ?)",
        (
            data['name'].strip(),
            data['date_from'],
            data.get('date_to', data['date_from']),
            float(data.get('weight_step', 5.0)),
            int(data.get('days', 1) or 1)
        )
    )
    tournament_id = cursor.lastrowid

    # Привязываем участников
    athlete_ids = data.get('athlete_ids') or []
    for aid in athlete_ids:
        conn.execute(
            "INSERT OR IGNORE INTO tournament_athletes (tournament_id, athlete_id) VALUES (?, ?)",
            (tournament_id, int(aid))
        )

    conn.commit()
    tournament = row_to_dict(
        conn.execute("SELECT * FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    )
    # Добавляем участников в ответ
    tournament['athlete_ids'] = [r['athlete_id'] for r in conn.execute(
        "SELECT athlete_id FROM tournament_athletes WHERE tournament_id = ?",
        (tournament_id,)
    ).fetchall()]
    conn.close()
    return jsonify(tournament), 201


@app.route('/api/tournaments/<int:tournament_id>', methods=['GET'])
def get_tournament(tournament_id):
    """Получить соревнование с его боями."""
    conn = get_db()
    tournament = row_to_dict(
        conn.execute("SELECT * FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    )
    if tournament is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404

    # Получаем список участников соревнования
    athlete_ids = [r['athlete_id'] for r in conn.execute(
        "SELECT athlete_id FROM tournament_athletes WHERE tournament_id = ?",
        (tournament_id,)
    ).fetchall()]
    tournament['athlete_ids'] = athlete_ids

    # Получаем бои соревнования
    fights = rows_to_list(
        conn.execute("SELECT * FROM fights WHERE tournament_id = ?", (tournament_id,)).fetchall()
    )
    tournament['fights'] = fights
    conn.close()
    return jsonify(tournament)


@app.route('/api/tournaments/<int:tournament_id>/athlete-count', methods=['GET'])
def get_tournament_athlete_count(tournament_id):
    """Получить количество участников соревнования."""
    conn = get_db()
    tournament = conn.execute("SELECT id FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    if tournament is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404

    count = conn.execute(
        "SELECT COUNT(*) as count FROM tournament_athletes WHERE tournament_id = ?",
        (tournament_id,)
    ).fetchone()['count']
    conn.close()
    return jsonify({'count': count})


@app.route('/api/tournaments/<int:tournament_id>/athletes', methods=['PUT'])
def update_tournament_athletes(tournament_id):
    """Обновить список участников соревнования."""
    data = request.get_json()
    athlete_ids = data.get('athlete_ids', [])
    conn = get_db()

    # Проверяем, что соревнование существует
    tournament = conn.execute("SELECT id FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    if tournament is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404

    # Удаляем старых участников
    conn.execute("DELETE FROM tournament_athletes WHERE tournament_id = ?", (tournament_id,))
    # Добавляем новых
    for aid in athlete_ids:
        conn.execute(
            "INSERT OR IGNORE INTO tournament_athletes (tournament_id, athlete_id) VALUES (?, ?)",
            (tournament_id, int(aid))
        )
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'athlete_ids': athlete_ids})


@app.route('/api/tournaments/<int:tournament_id>/athletes', methods=['POST'])
def add_tournament_athlete(tournament_id):
    """Добавить спортсмена в участники соревнования."""
    data = request.get_json()
    athlete_id = data.get('athlete_id')
    
    if not athlete_id:
        return jsonify({'error': 'Не укР°Р·Р°РЅ athlete_id'}), 400

    conn = get_db()
    tournament = conn.execute("SELECT id FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    if tournament is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404

    athlete = conn.execute("SELECT id FROM athletes WHERE id = ?", (athlete_id,)).fetchone()
    if athlete is None:
        conn.close()
        return jsonify({'error': 'Спортсмен не нР°Р№РґРµРЅ'}), 404

    # Добавляем участника
    conn.execute(
        "INSERT OR IGNORE INTO tournament_athletes (tournament_id, athlete_id) VALUES (?, ?)",
        (tournament_id, int(athlete_id))
    )
    conn.commit()
    conn.close()
    return jsonify({'success': True})


@app.route('/api/tournaments/<int:tournament_id>/athletes/<int:athlete_id>', methods=['DELETE'])
def remove_tournament_athlete(tournament_id, athlete_id):
    """Удалить спортсмена из участников соревнования."""
    conn = get_db()
    tournament = conn.execute("SELECT id FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    if tournament is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404

    athlete = conn.execute("SELECT id FROM athletes WHERE id = ?", (athlete_id,)).fetchone()
    if athlete is None:
        conn.close()
        return jsonify({'error': 'Спортсмен не нР°Р№РґРµРЅ'}), 404

    # Удаляем участника
    conn.execute(
        "DELETE FROM tournament_athletes WHERE tournament_id = ? AND athlete_id = ?",
        (tournament_id, athlete_id)
    )
    conn.commit()
    conn.close()
    return jsonify({'success': True})


@app.route('/api/tournaments/<int:tournament_id>', methods=['DELETE'])
def delete_tournament(tournament_id):
    """Удалить соревнование со всеми боями (каскадное удаление через FOREIGN KEY)."""
    conn = get_db()
    conn.execute("DELETE FROM tournaments WHERE id = ?", (tournament_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})


@app.route('/api/tournaments/<int:tournament_id>', methods=['PUT'])
def update_tournament(tournament_id):
    """Обновить настройки соревнования (имя, даты, число дней, шаг веса)."""
    data = request.get_json(silent=True) or {}
    conn = get_db()
    existing = conn.execute("SELECT * FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    if existing is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404

    fields = []
    params = []
    if 'name' in data and data['name']:
        fields.append("name = ?")
        params.append(str(data['name']).strip())
    if 'date_from' in data:
        fields.append("date_from = ?")
        params.append(str(data['date_from']))
    if 'date_to' in data:
        fields.append("date_to = ?")
        params.append(str(data['date_to']))
    if 'days' in data:
        try:
            days = int(data['days'])
        except (TypeError, ValueError):
            days = 1
        fields.append("days = ?")
        params.append(max(1, min(days, 14)))
    if 'weight_step' in data:
        try:
            fields.append("weight_step = ?")
            params.append(float(data['weight_step']))
        except (TypeError, ValueError):
            pass

    if fields:
        params.append(tournament_id)
        conn.execute(f"UPDATE tournaments SET {', '.join(fields)} WHERE id = ?", params)
        conn.commit()

    tournament = row_to_dict(
        conn.execute("SELECT * FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    )
    conn.close()
    return jsonify(tournament)


# ======================== API: Р‘РћР ========================

@app.route('/api/tournaments/<int:tournament_id>/fights', methods=['GET'])
def get_fights(tournament_id):
    """Получить бои соревнования."""
    conn = get_db()
    fights = rows_to_list(
        conn.execute("SELECT * FROM fights WHERE tournament_id = ? ORDER BY round, id", (tournament_id,)).fetchall()
    )
    conn.close()
    return jsonify(fights)


@app.route('/api/tournaments/<int:tournament_id>/fights', methods=['POST'])
def add_fight(tournament_id):
    """Добавить бой в соревнование."""
    data = request.get_json()
    conn = get_db()

    # Проверяем существование соревнования
    tournament = conn.execute("SELECT * FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    if tournament is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404

    cursor = conn.execute(
        "INSERT INTO fights (tournament_id, round, athlete1_id, athlete2_id, weight_class) VALUES (?, ?, ?, ?, ?)",
        (
            tournament_id,
            int(data['round']),
            data.get('athlete1_id'),
            data.get('athlete2_id'),
            data['weight_class']
        )
    )
    fight_id = cursor.lastrowid
    conn.commit()
    fight = row_to_dict(
        conn.execute("SELECT * FROM fights WHERE id = ?", (fight_id,)).fetchone()
    )
    conn.close()
    return jsonify(fight), 201


@app.route('/api/fights/<int:fight_id>', methods=['PUT'])
def update_fight(fight_id):
    """Обновить результат боя."""
    data = request.get_json()
    conn = get_db()

    fight = conn.execute("SELECT * FROM fights WHERE id = ?", (fight_id,)).fetchone()
    if fight is None:
        conn.close()
        return jsonify({'error': 'Бой не нР°Р№РґРµРЅ'}), 404

    # Обновляем поля
    fields = []
    params = []
    for field in ['winner_id', 'method', 'completed', 'athlete1_id', 'athlete2_id', 'round']:
        if field in data:
            fields.append(f"{field} = ?")
            params.append(data[field])

    if fields:
        params.append(fight_id)
        conn.execute(f"UPDATE fights SET {', '.join(fields)} WHERE id = ?", params)
        conn.commit()

    fight = row_to_dict(conn.execute("SELECT * FROM fights WHERE id = ?", (fight_id,)).fetchone())
    conn.close()
    return jsonify(fight)


@app.route('/api/fights/<int:fight_id>', methods=['DELETE'])
def delete_fight(fight_id):
    """Удалить бой."""
    conn = get_db()
    conn.execute("DELETE FROM fights WHERE id = ?", (fight_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})


# ======================== API: ГЕНЕРАР¦РЯ СЕТКР ========================

def compute_weight_classes(weight_step, athlete_ids=None):
    """Возвращает отсортированный список весовых категорий (по полу и весу)
    только для указанных спортсменов. athlete_ids обязателен."""
    conn = get_db()
    if not athlete_ids:
        conn.close()
        return []

    placeholders = ','.join('?' * len(athlete_ids))
    athletes = rows_to_list(
        conn.execute(f"SELECT * FROM athletes WHERE id IN ({placeholders})", athlete_ids).fetchall()
    )
    conn.close()

    step = weight_step or 5.0
    categories = set()
    for a in athletes:
        if a['weight'] is None:
            continue
        lower = int(a['weight'] / step) * step
        upper = int(lower + step - 1)
        categories.add(f"{a['gender']} {int(lower)}-{upper} кг")

    def sort_key(cat):
        gender_part, weight_part = cat.split(' ', 1)
        lower = float(weight_part.split('-')[0])
        return (gender_part, lower)

    return sorted(categories, key=sort_key)


@app.route('/api/tournaments/<int:tournament_id>/weight-classes', methods=['GET'])
def tournament_weight_classes(tournament_id):
    """Вернуть доступные весовые категории для соревнования."""
    conn = get_db()
    tournament = row_to_dict(
        conn.execute("SELECT * FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    )
    if tournament is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404

    # Берём участников соревнования
    athlete_ids = [r['athlete_id'] for r in conn.execute(
        "SELECT athlete_id FROM tournament_athletes WHERE tournament_id = ?",
        (tournament_id,)
    ).fetchall()]
    conn.close()

    return jsonify({'weight_classes': compute_weight_classes(tournament['weight_step'], athlete_ids)})


@app.route('/api/tournaments/<int:tournament_id>/generate', methods=['POST'])
def generate_bracket(tournament_id):
    """Сгенерировать полную турнирную сетку (все раунды сразу).

    Строится полноценное дерево: победители каждого боя попадают
    в строго определённый слот следующего раунда. BYE-бои сразу
    завершаются, а их победители автоматически продвигаются дальше
    (в том числе каскадно через несколько раундов).

    Тело запроса (JSON, необязательно):
      {"weight_classes": ["мужской 55-59.9 кг", ...]}
    Если указано — сетка строится только для выбранных категорий.
    """
    import random

    conn = get_db()

    tournament = row_to_dict(
        conn.execute("SELECT * FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    )
    if tournament is None:
        return jsonify({'error': 'Соревнование не найдено'}), 404

    # Берём только участников соревнования (из таблицы tournament_athletes)
    athlete_ids = [r['athlete_id'] for r in conn.execute(
        "SELECT athlete_id FROM tournament_athletes WHERE tournament_id = ?",
        (tournament_id,)
    ).fetchall()]

    if not athlete_ids:
        return jsonify({'error': 'В соревновании нет участников. Добавьте спортсменов через кнопку "Участники"'}), 400

    placeholders = ','.join('?' * len(athlete_ids))
    athletes = rows_to_list(
        conn.execute(f"SELECT * FROM athletes WHERE id IN ({placeholders})", athlete_ids).fetchall()
    )

    if len(athletes) < 2:
        return jsonify({'error': 'Нужно минимум 2 спортсмена'}), 400

    # Какие категории строим
    data = request.get_json(silent=True) or {}
    selected = data.get('weight_classes') or None
    if selected is not None:
        selected = set(selected)

    # Удаляем старые бои перед созданием новой сетки
    log.info(f"generate_bracket: deleting old fights for tournament {tournament_id}")
    conn.execute("DELETE FROM fights WHERE tournament_id = ?", (tournament_id,))

    # Группируем по весу и полу
    weight_step = tournament['weight_step'] or 5.0
    weight_groups = {}

    for a in athletes:
        lower = int(a['weight'] / weight_step) * weight_step
        upper = lower + weight_step - 0.1
        key = format_weight_class(a['gender'], a['weight'], weight_step)
        if selected is not None and key not in selected:
            continue
        weight_groups.setdefault(key, []).append(a)

    log.info(f"generate_bracket: weight_groups = {list(weight_groups.keys())}")

    def advance_byes(rounds):
        """Каскадно продвигает победителей BYE-боев по дереву."""
        max_round = max(rounds.keys())
        changed = True
        while changed:
            changed = False
            for r in range(1, max_round):
                for i, pair in enumerate(rounds[r]):
                    a1, a2 = pair
                    winner = None
                    if a1 is not None and a2 is None:
                        winner = a1
                    elif a1 is None and a2 is not None:
                        winner = a2
                    if winner is None:
                        continue
                    parent = rounds[r + 1][i // 2]
                    slot = 0 if i % 2 == 0 else 1
                    if parent[slot] is None:
                        parent[slot] = winner
                        changed = True
        return rounds

    # Строим дерево для каждой категории
    for weight_class, group in weight_groups.items():
        n = len(group)
        if n < 2:
            continue

        import math
        total_slots = 2 ** math.ceil(math.log2(n))
        num_rounds = int(math.log2(total_slots))

        # Генерируем слоты для каждого раунда
        rounds = {}
        for r in range(1, num_rounds + 1):
            slots = total_slots // (2 ** (r - 1))
            rounds[r] = [(None, None) for _ in range(slots // 2)]

        # Заполняем первый раунд - парами спортсменов
        for i, a in enumerate(group):
            if i < len(rounds[1]):
                # Чётные индексы идут в athlete1, нечётнС‹Рµ - РІ athlete2
                pair_idx = i // 2
                slot = i % 2  # 0 = athlete1, 1 = athlete2
                if slot == 0:
                    rounds[1][pair_idx] = (a['id'], None)
                else:
                    # Заполняем athlete2 текущего боя
                    current = list(rounds[1][pair_idx])
                    current[1] = a['id']
                    rounds[1][pair_idx] = tuple(current)

        # Продвигаем BYE
        rounds = advance_byes(rounds)

        # Создаём бои
        for r in range(1, num_rounds + 1):
            for i, (a1_id, a2_id) in enumerate(rounds[r]):
                cur = conn.execute(
                    "INSERT INTO fights (tournament_id, round, athlete1_id, athlete2_id, weight_class) VALUES (?, ?, ?, ?, ?)",
                    (tournament_id, r, a1_id, a2_id, weight_class)
                )
                fight_id = cur.lastrowid

                # Если один из спортсменов None — это BYE
                if a1_id and not a2_id:
                    conn.execute(
                        "UPDATE fights SET winner_id = ?, completed = 1, method = 'авто' WHERE id = ?",
                        (a1_id, fight_id)
                    )
                elif not a1_id and a2_id:
                    conn.execute(
                        "UPDATE fights SET winner_id = ?, completed = 1, method = 'авто' WHERE id = ?",
                        (a2_id, fight_id)
                    )

    conn.commit()
    log.info(f"generate_bracket: committed for tournament {tournament_id}")
    return jsonify({
        'success': True,
        'message': 'Сетка успешно сформирована!'
    }), 201


@app.route('/api/tournaments/<int:tournament_id>/lottery', methods=['GET'])
def get_lottery(tournament_id):
    """Получить сохранённую жеребьёвку."""
    conn = get_db()
    tournament = row_to_dict(
        conn.execute("SELECT * FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    )
    if tournament is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404

    # Получаем сохранённые слоты из таблицы lottery
    lottery_rows = rows_to_list(
        conn.execute("SELECT * FROM lottery WHERE tournament_id = ? ORDER BY weight_class, slot", (tournament_id,)).fetchall()
    )
    
    # Получаем спортсменов
    athlete_ids = [r['athlete_id'] for r in conn.execute(
        "SELECT athlete_id FROM tournament_athletes WHERE tournament_id = ?", (tournament_id,)
    ).fetchall()]
    
    placeholders = ','.join('?' * len(athlete_ids)) if athlete_ids else "''"
    athletes = rows_to_list(
        conn.execute(f"SELECT * FROM athletes WHERE id IN ({placeholders})", athlete_ids).fetchall()
    )
    
    # Группируем по весовым категориям
    result = {}
    for row in lottery_rows:
        wc = row['weight_class']
        if wc not in result:
            result[wc] = []
        
        athlete = next((a for a in athletes if a['id'] == row['athlete_id']), None)
        if athlete:
            result[wc].append({
                'athlete_id': row['athlete_id'],
                'slot': row['slot'],
                'day': row['day'] if 'day' in row.keys() else 1,
                'name': athlete['name'],
                'weight': athlete['weight'],
                'gender': athlete['gender'],
                'birth_year': athlete['birth_year'],
                'region': athlete['region'],
                'coach': athlete['coach']
            })
    
    conn.close()
    return jsonify(result)


@app.route('/api/tournaments/<int:tournament_id>/lottery', methods=['POST'])
def save_lottery(tournament_id):
    """Сохранить жеребьёвку (слоты спортсменов по весовым категориям)."""
    data = request.get_json()
    if not data:
        return jsonify({'error': 'Нет даннС‹С…'}), 400
    
    conn = get_db()
    
    # Проверяем существование турнира
    tournament = conn.execute("SELECT id FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    if not tournament:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404
    
    # Удаляем старые слоты
    conn.execute("DELETE FROM lottery WHERE tournament_id = ?", (tournament_id,))
    
# Сохраняем новые слоты
    for weight_class, slots in data.items():
        for slot_data in slots:
            try:
                conn.execute(
                    "INSERT INTO lottery (tournament_id, athlete_id, slot, weight_class, day) VALUES (?, ?, ?, ?, ?)",
                    (tournament_id, slot_data['athlete_id'], slot_data['slot'], weight_class, slot_data.get('day', 1) or 1)
                )
            except Exception as e:
                conn.rollback()
                conn.close()
                return jsonify({'error': f'Ошибка сохранения: {str(e)}'}), 500
    
    conn.commit()
    conn.close()
    return jsonify({'message': 'Жеребьевка сохраненР°'})


@app.route('/api/tournaments/<int:tournament_id>/fights/<int:fight_id>/athlete/<int:slot>', methods=['PUT'])
def update_fight_athlete(tournament_id, fight_id, slot):
    """Обновить спортсмена в бою (slot=1 или slot=2)."""
    data = request.get_json()
    athlete_id = data.get('athlete_id')
    
    conn = get_db()
    
    # Проверяем существование боя
    fight = conn.execute("SELECT * FROM fights WHERE id = ? AND tournament_id = ?", (fight_id, tournament_id)).fetchone()
    if not fight:
        conn.close()
        return jsonify({'error': 'Бой не нР°Р№РґРµРЅ'}), 404
    
    # Обновляем спортсмена
    slot_field = 'athlete1_id' if slot == 1 else 'athlete2_id'
    conn.execute(
        f"UPDATE fights SET {slot_field} = ? WHERE id = ?",
        (athlete_id, fight_id)
    )
    
    conn.commit()
    conn.close()
    return jsonify({'message': 'Спортсмен обновлён'})


@app.route('/api/tournaments/<int:tournament_id>/clear', methods=['POST'])
def clear_bracket(tournament_id):
    """Очистить все бои турнира."""
    # Получаем тело запроса (может быть пустым)
    request.get_json(silent=True)
    
    conn = get_db()
    
    # Проверяем существование турнира
    tournament = conn.execute("SELECT id FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    if not tournament:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404
    
    # Удаляем все бои
    conn.execute("DELETE FROM fights WHERE tournament_id = ?", (tournament_id,))
    conn.execute("DELETE FROM bracket_results WHERE tournament_id = ?", (tournament_id,))
    conn.commit()
    conn.close()
    return jsonify({'message': 'Сетка очищена'})


@app.route('/api/tournaments/<int:tournament_id>/bracket-results', methods=['GET'])
def get_bracket_results(tournament_id):
    """Получить сохранённые победители пар сетки."""
    conn = get_db()
    rows = rows_to_list(
        conn.execute(
            "SELECT weight_class, pair_index, winner_id, method FROM bracket_results WHERE tournament_id = ?",
            (tournament_id,)
        ).fetchall()
    )
    conn.close()
    return jsonify(rows)


@app.route('/api/tournaments/<int:tournament_id>/bracket-results', methods=['DELETE'])
def clear_bracket_results(tournament_id):
    """Полностью очистить сохранённые результаты пар сетки турнира."""
    conn = get_db()
    conn.execute("DELETE FROM bracket_results WHERE tournament_id = ?", (tournament_id,))
    conn.commit()
    conn.close()
    return jsonify({'message': 'Результаты сетки очищены'})


@app.route('/api/tournaments/<int:tournament_id>/weight-classes/<path:weight_class>/results', methods=['DELETE'])
def clear_weight_class_results(tournament_id, weight_class):
    """Очистить результаты пар сетки только для одной весовой категории."""
    conn = get_db()
    conn.execute(
        "DELETE FROM bracket_results WHERE tournament_id = ? AND weight_class = ?",
        (tournament_id, weight_class)
    )
    conn.commit()
    conn.close()
    return jsonify({'message': 'Результаты категории очищены'})


@app.route('/api/tournaments/<int:tournament_id>/bracket-results', methods=['POST'])
def save_bracket_result(tournament_id):
    """Сохранить победителя пары сетки.

    Тело запроса: {"weight_class": "...", "pair_index": 5, "winner_id": 5, "method": "...", "day": 1}
    Ключ pair_index — сквозной индекс пары (для раундов 2+ он раунд-кодирован,
    например 1000+idx — раунд 2, 2000+idx — раунд 3). Повторный вызов с той же парой
    перезаписывает результат.
    """
    data = request.get_json()
    weight_class = data.get('weight_class')
    pair_index = data.get('pair_index')
    winner_id = data.get('winner_id')
    method = data.get('method') or ''
    day = data.get('day') or 1

    if weight_class is None or pair_index is None or not winner_id:
        return jsonify({'error': 'Нужны поля weight_class, pair_index, winner_id'}), 400

    try:
        pair_index = int(pair_index)
        day = int(day)
    except (TypeError, ValueError):
        return jsonify({'error': 'pair_index и day должны быть числами'}), 400

    conn = get_db()

    # Проверяем существование турнира
    tournament = conn.execute("SELECT id FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    if not tournament:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404

    # Проверяем, что победитель — участник этого турнира
    participant = conn.execute(
        "SELECT 1 FROM tournament_athletes WHERE tournament_id = ? AND athlete_id = ?",
        (tournament_id, winner_id)
    ).fetchone()
    if not participant:
        conn.close()
        return jsonify({'error': 'Победитель не является участником соревнования'}), 400

    conn.execute('''
        INSERT INTO bracket_results (tournament_id, weight_class, pair_index, winner_id, method, day)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(tournament_id, weight_class, pair_index)
        DO UPDATE SET winner_id = excluded.winner_id, method = excluded.method, day = excluded.day
    ''', (tournament_id, weight_class, pair_index, winner_id, method, day))

    conn.commit()
    conn.close()
    return jsonify({'message': 'Победитель сохранён'}), 201


@app.route('/api/tournaments/<int:tournament_id>/fights/<int:fight_id>/result', methods=['PUT', 'DELETE'])
def update_fight_result(tournament_id, fight_id):
    """Обновить результат боя (выбрать победителя) либо сбросить его (DELETE)."""
    conn = get_db()

    fight = conn.execute("SELECT * FROM fights WHERE id = ? AND tournament_id = ?", (fight_id, tournament_id)).fetchone()
    if not fight:
        conn.close()
        return jsonify({'error': 'Бой не найден'}), 404

    # DELETE — отменить/сбросить результат боя
    if request.method == 'DELETE':
        conn.execute(
            "UPDATE fights SET winner_id = NULL, completed = 0, method = NULL WHERE id = ?",
            (fight_id,)
        )
        conn.commit()
        conn.close()
        return jsonify({'message': 'Результат боя сброшен'})

    data = request.get_json()
    winner_id = data.get('winner_id')

    # Проверяем что победитель один из участников
    if fight['athlete1_id'] != winner_id and fight['athlete2_id'] != winner_id:
        return jsonify({'error': 'Победитель должен быть одним из участников боя'}), 400

    # Обновляем результат
    conn.execute(
        "UPDATE fights SET winner_id = ?, completed = 1, method = 'ручной' WHERE id = ?",
        (winner_id, fight_id)
    )
    conn.commit()
    conn.close()
    return jsonify({'message': 'Результат сохранён'})




# ======================== API: ОБНОВЛЕНИЯ ========================

@app.route('/api/update/status', methods=['GET'])
def update_status():
    """Проверка доступных обновлений из GitHub."""
    configured = bool(UPDATE_OWNER and UPDATE_REPO)
    if not configured:
        return jsonify({
            'configured': False,
            'message': 'Обновления не настроены. Укажите GITHUB_OWNER и GITHUB_REPO.',
            'local_version': _read_local_version()
        })

    try:
        manifest_url = f'https://raw.githubusercontent.com/{UPDATE_OWNER}/{UPDATE_REPO}/{UPDATE_BRANCH}/{UPDATE_MANIFEST}'
        raw = _fetch_url(manifest_url, timeout=20)
        manifest = json.loads(raw.decode('utf-8'))
    except Exception as e:
        return jsonify({
            'configured': True, 'owner': UPDATE_OWNER, 'repo': UPDATE_REPO, 'branch': UPDATE_BRANCH,
            'error': f'Не удалось получить манифест: {e}',
            'local_version': _read_local_version()
        })

    version = str(manifest.get('version', ''))
    files = [f for f in manifest.get('files', []) if f in UPDATE_ALLOWED]
    local_version = _read_local_version()
    up_to_date = bool(local_version and local_version == version)
    return jsonify({
        'configured': True,
        'owner': UPDATE_OWNER, 'repo': UPDATE_REPO, 'branch': UPDATE_BRANCH,
        'local_version': local_version,
        'latest_version': version,
        'up_to_date': up_to_date,
        'files': files
    })


@app.route('/api/update/apply', methods=['POST'])
def update_apply():
    """Скачать и применить обновления файлов кода из GitHub."""
    if not (UPDATE_OWNER and UPDATE_REPO):
        return jsonify({'error': 'Обновления не настроены. Укажите GITHUB_OWNER и GITHUB_REPO.'}), 400

    base = f'https://raw.githubusercontent.com/{UPDATE_OWNER}/{UPDATE_REPO}/{UPDATE_BRANCH}'
    app_dir = Path(__file__).resolve().parent

    try:
        manifest = json.loads(_fetch_url(f'{base}/{UPDATE_MANIFEST}', timeout=20).decode('utf-8'))
    except Exception as e:
        return jsonify({'error': f'Не удалось получить манифест: {e}'}), 500

    version = str(manifest.get('version', ''))
    requested = manifest.get('files', [])
    updated = []
    skipped = []
    for rel in requested:
        name = os.path.basename(str(rel))
        if name not in UPDATE_ALLOWED:
            skipped.append(name)
            continue
        if name in ('boxing.db',) or name == LOCAL_VERSION_FILE:
            skipped.append(name)
            continue
        try:
            content = _fetch_url(f'{base}/{rel}', timeout=30)
        except Exception as e:
            return jsonify({'error': f'Не удалось скачать {name}: {e}'}), 500
        try:
            target = app_dir / name
            with open(target, 'wb') as f:
                f.write(content)
            updated.append(name)
        except Exception as e:
            return jsonify({'error': f'Не удалось сохранить {name}: {e}'}), 500

    _write_local_version(version)
    return jsonify({
        'ok': True,
        'version': version,
        'updated': updated,
        'skipped': skipped
    })



# ======================== СТАТИЧЕСКИЕ ФАЙЛЫ ========================

@app.route('/')
def index():
    """Главная страница."""
    return send_from_directory('.', 'index.html')

@app.route('/<path:path>')
def static_files(path):
    """Отдача статических файлов (JS, CSS Рё С‚.Рґ.)."""
    if path.endswith('.js'):
        return send_from_directory('.', path, mimetype='application/javascript')
    return send_from_directory('.', path)


# ======================== ЗАПУСК ========================

DEFAULT_PORT = 8080


def find_free_port():
    """Рщет свободный порт, начиная с DEFAULT_PORT."""
    for port in range(DEFAULT_PORT, DEFAULT_PORT + 100):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('', port))
                return port
        except OSError:
            continue
    # Если ничего не нашли — берём случайный
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]


def open_browser(port):
    time.sleep(1.5)
    webbrowser.open(f'http://localhost:{port}')


if __name__ == '__main__':
    # Инициализация БД
    init_db()

    # Резервная копия БД при старте
    backup_db(force=True)

    port = find_free_port()

    print('=' * 60)
    print('  >> СОРЕВНОВАНРЯ ПО БОКСУ (SQLite + Flask)')
    print('=' * 60)
    print(f'\n  Сервер запущен: http://localhost:{port}')
    print('  Нажмите Ctrl+C для остановки\n')

    # Открываем браузер
    t = threading.Thread(target=open_browser, args=(port,), daemon=True)
    t.start()

    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
