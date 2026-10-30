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
import os
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

from flask import Flask, request, jsonify, send_file, send_from_directory, g, session, redirect, url_for
from flask_cors import CORS

from database import get_db, init_db

# Генерация отчётов в Word (.docx) — опционально
try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    _HAS_DOCX = True
except Exception:
    Document = None
    _HAS_DOCX = False

# ======================== ПУТИ (обычный запуск / заморозка PyInstaller) ========================
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))


def _get_data_dir():
    if getattr(sys, 'frozen', False):
        # Портативный режим: если рядом с exe (или внутри bundle) есть файл portable.flag —
        # данные храним в папке exe. Иначе — в %APPDATA%\BoxingApp.
        exe_dir = os.path.dirname(sys.executable)
        bundled = getattr(sys, '_MEIPASS', exe_dir)
        if os.path.exists(os.path.join(bundled, 'portable.flag')) or os.path.exists(os.path.join(exe_dir, 'portable.flag')):
            return exe_dir
        base = os.environ.get('BOXING_DATA_DIR') or \
               os.path.join(os.environ.get('APPDATA') or os.path.expanduser('~'), 'BoxingApp')
        os.makedirs(base, exist_ok=True)
        return base
    return PROJECT_DIR


DATA_DIR = _get_data_dir()


def _read_json_version(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return str(json.load(f).get('version', '')).strip()
    except Exception:
        return ''


def _ver_tuple(v):
    parts = (str(v or '').split('.') + ['0', '0', '0'])[:3]
    try:
        return tuple(int(p) for p in parts)
    except Exception:
        return (0, 0, 0)


# Ресурсы (index.html, app.js и др.) отдаём из каталога данных. В заморозке (exe)
# копируем их туда из _MEIPASS, если файла нет ИЛИ сохранённая версия СТАРШЕ встроенной —
# так новая сборка обновляет старый каталог, но не затирает пользовательские обновления.
if getattr(sys, 'frozen', False):
    _bundled = getattr(sys, '_MEIPASS', PROJECT_DIR)
    _bm = _read_json_version(os.path.join(_bundled, 'update-manifest.json'))
    for _f in ('index.html', 'login.html', 'app.js', 'style.css', 'update-manifest.json'):
        _dst = os.path.join(DATA_DIR, _f)
        _src = os.path.join(_bundled, _f)
        if not os.path.exists(_src):
            continue
        _need_copy = False
        if not os.path.exists(_dst):
            _need_copy = True
        elif _f == 'update-manifest.json':
            _dm = _read_json_version(_dst)
            _need_copy = _ver_tuple(_bm) > _ver_tuple(_dm)
        else:
            # статический файл: обновляем, если встроенная версия не ниже сохранённой
            _dm = _read_json_version(os.path.join(DATA_DIR, 'update-manifest.json'))
            _need_copy = _ver_tuple(_bm) > _ver_tuple(_dm)
        if _need_copy:
            try:
                shutil.copy2(_src, _dst)
            except Exception:
                pass
    # Шаблоны отчётов тоже копируем в каталог данных (для переносной версии —
    # в папку рядом с exe), чтобы их можно было редактировать/заменять.
    # При обновлении версии (встроенная новее сохранённой) — перезаписываем шаблоны,
    # чтобы старые/повреждённые копии не оставались после апдейта.
    for _f in ('ogl_tmp.docx', 'sp_tmp.docx', 'sp_tmp_new.docx', 'reglament_tmp.docx'):
        _src = os.path.join(_bundled, _f)
        _dst = os.path.join(DATA_DIR, _f)
        if not os.path.exists(_src):
            continue
        _need_tpl = False
        if not os.path.exists(_dst):
            _need_tpl = True
        else:
            # перезапишем, если встроенная версия новее сохранённой (для .docx по update-manifest)
            _dm = _read_json_version(os.path.join(DATA_DIR, 'update-manifest.json'))
            _need_tpl = _ver_tuple(_bm) > _ver_tuple(_dm)
        if _need_tpl:
            try:
                shutil.copy2(_src, _dst)
            except Exception:
                pass
    ASSETS_DIR = DATA_DIR
else:
    ASSETS_DIR = PROJECT_DIR


def _db_path():
    return os.path.join(DATA_DIR, 'boxing.db')


def _local_version_path():
    return os.path.join(DATA_DIR, LOCAL_VERSION_FILE)


def _backup_dir():
    return os.path.join(DATA_DIR, 'backups')


app = Flask(__name__, static_folder=ASSETS_DIR)
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
    'server.py', 'database.py', 'app.py', 'requirements.txt',
    'update-manifest.json', 'ogl_tmp.docx', 'sp_tmp.docx', 'sp_tmp_new.docx',
    'reglament_tmp.docx'
}


def _fetch_url(url, timeout=25):
    req = urllib.request.Request(url, headers={'User-Agent': 'boxing-updater/1.0'})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _read_local_version():
    try:
        with open(_local_version_path(), 'r', encoding='utf-8') as f:
            return f.read().strip()
    except Exception:
        return ''


def _write_local_version(v):
    try:
        with open(_local_version_path(), 'w', encoding='utf-8') as f:
            f.write(str(v).strip())
    except Exception:
        pass


def _shipped_version():
    """Фактическая версия установленного приложения — из встроенного/локального
    update-manifest.json (та самая, что показывается в «О приложении»).
    Это надёжнее, чем local_version.txt (файл отслеживания обновлений), который
    может расходиться с реальной версией после пересборки/копирования.
    """
    try:
        mpath = os.path.join(ASSETS_DIR, 'update-manifest.json')
        with open(mpath, 'r', encoding='utf-8') as f:
            m = json.load(f)
        v = str(m.get('version', ''))
        if v:
            return v
    except Exception:
        pass
    return _read_local_version()


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
    return send_from_directory(ASSETS_DIR, 'login.html')


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

    db_path = Path(_db_path())
    if not db_path.exists():
        return None

    try:
        backup_path = Path(_backup_dir())
        backup_path.mkdir(parents=True, exist_ok=True)
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
    """Форматирует строку весовой категории по точному весу спортсмена."""
    w = float(weight)
    w_str = f"{int(w)}" if w == int(w) else f"{w:g}"
    return f"{gender} {w_str} кг"


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


@app.route('/api/athletes/clear', methods=['DELETE'])
def clear_athletes():
    """Полностью очистить список спортсменов и связанные данные."""
    conn = get_db()
    try:
        conn.execute("DELETE FROM tournament_athletes")
        conn.execute("DELETE FROM bracket_results")
        conn.execute("DELETE FROM lottery")
        conn.execute("DELETE FROM fights")
        conn.execute("DELETE FROM athletes")
        conn.commit()
    finally:
        conn.close()
    return jsonify({'success': True, 'deleted': True})


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
        "INSERT INTO tournaments (name, date_from, date_to, weight_step, days, chief_referee, chief_secretary, venue) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (
            data['name'].strip(),
            data['date_from'],
            data.get('date_to', data['date_from']),
            float(data.get('weight_step', 5.0)),
            int(data.get('days', 1) or 1),
            (data.get('chief_referee') or '').strip(),
            (data.get('chief_secretary') or '').strip(),
            (data.get('venue') or '').strip()
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


@app.route('/api/tournaments/<int:tournament_id>/judges', methods=['POST'])
def update_tournament_judges(tournament_id):
    """Обновить судей и место проведения текущего соревнования."""
    data = request.get_json(silent=True) or {}
    conn = get_db()
    t = conn.execute("SELECT id FROM tournaments WHERE id = ?", (tournament_id,)).fetchone()
    if t is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404
    conn.execute(
        "UPDATE tournaments SET chief_referee = ?, chief_secretary = ?, venue = ? WHERE id = ?",
        ((data.get('chief_referee') or '').strip(),
         (data.get('chief_secretary') or '').strip(),
         (data.get('venue') or '').strip(),
         tournament_id)
    )
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
        categories.add(format_weight_class(a['gender'], a['weight'], step))

    def sort_key(cat):
        gender_part, weight_part = cat.split(' ', 1)
        lower = float(weight_part.replace(' кг', ''))
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




# ======================== API: НАСТРОЙКИ ========================

SETTINGS_FILE = os.path.join(DATA_DIR, 'settings.json')


def _load_settings():
    try:
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


@app.route('/api/settings', methods=['GET'])
def get_settings():
    return jsonify(_load_settings())


@app.route('/api/settings', methods=['POST'])
def save_settings():
    data = request.get_json(silent=True) or {}
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ======================== API: ОБНОВЛЕНИЯ ========================

@app.route('/api/update/status', methods=['GET'])
def update_status():
    """Проверка доступных обновлений из GitHub."""
    configured = bool(UPDATE_OWNER and UPDATE_REPO)
    local_version = _shipped_version()
    if not configured:
        return jsonify({
            'configured': False,
            'message': 'Обновления не настроены. Укажите GITHUB_OWNER и GITHUB_REPO.',
            'local_version': local_version
        })

    try:
        manifest_url = f'https://raw.githubusercontent.com/{UPDATE_OWNER}/{UPDATE_REPO}/{UPDATE_BRANCH}/{UPDATE_MANIFEST}'
        raw = _fetch_url(manifest_url, timeout=20)
        manifest = json.loads(raw.decode('utf-8'))
    except Exception as e:
        return jsonify({
            'configured': True, 'owner': UPDATE_OWNER, 'repo': UPDATE_REPO, 'branch': UPDATE_BRANCH,
            'error': f'Не удалось получить манифест: {e}',
            'local_version': local_version
        })

    version = str(manifest.get('version', ''))
    files = [f for f in manifest.get('files', []) if f in UPDATE_ALLOWED]
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
    # Пишем обновлённые файлы в каталог данных (там же, откуда они отдаются),
    # чтобы они сохранялись после перезапуска в портативной/собранной версии.
    app_dir = Path(ASSETS_DIR)

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



# ======================== API: ИМПОРТ ИЗ WORD (.docx) ========================

def _docx_cell_text(tc, W):
    """Собирает текст ячейки таблицы Word."""
    parts = []
    for t in tc.iter(W + 't'):
        parts.append(t.text or '')
    return ''.join(parts).strip()


def _docx_cell_yellow(tc, W):
    """True, если ячейка подсвечена жёлтым (shading FFFFxx или highlight yellow)."""
    try:
        shd = tc.find('.//' + W + 'shd')
        if shd is not None:
            fill = (shd.get(W + 'fill') or '').upper()
            if fill.startswith('FFFF') or fill in ('FFF2CC', 'FFE699', 'FFD966'):
                return True
        for hl in tc.iter(W + 'highlight'):
            if (hl.get(W + 'val') or '').lower() in ('yellow',):
                return True
    except Exception:
        pass
    return False


def parse_docx_tables(file_bytes):
    """Возвращает список таблиц из .docx.

    Каждая таблица — список строк; каждая строка — список ячеек-словарей:
        {'t': текст, 'y': подсвечена ли жёлтым}.
    """
    from io import BytesIO
    from zipfile import ZipFile
    from xml.etree import ElementTree as ET
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    z = ZipFile(BytesIO(file_bytes))
    xml = z.read('word/document.xml').decode('utf-8', 'replace')
    root = ET.fromstring(xml)
    tables = []
    for tbl in root.iter(W + 'tbl'):
        rows = []
        for tr in tbl.iter(W + 'tr'):
            cells = [{'t': _docx_cell_text(tc, W), 'y': _docx_cell_yellow(tc, W)}
                     for tc in tr.findall(W + 'tc')]
            rows.append(cells)
        if rows:
            tables.append(rows)
    return tables


def _norm_header(s):
    """Нормализуем заголовок: нижний регистр, убираем подчёркивания."""
    return (s or '').strip().lower().replace('_', ' ')


def _find_col(colmap, *keys):
    for k in keys:
        if k in colmap:
            return colmap[k]
    return None


def _extract_weight(s):
    """Из 'Весовая категория' берём первое число: '-36' -> 36, '55-59.9' -> 55."""
    import re
    m = re.search(r'(\d+(?:[.,]\d+)?)', str(s or ''))
    return float(m.group(1).replace(',', '.')) if m else None


def _extract_year(s):
    """Из даты/строки '02.07.2015' берём год (2015)."""
    import re
    m = re.search(r'(19|20)\d{2}', str(s or ''))
    return int(m.group(0)) if m else 0


def _pick_col(header, *keys):
    """Выбирает индекс столбца по ключевым словам в заголовке.

    Приоритет — столбцы, подсвеченные жёлтым (header[i]['y']); если их нет —
    любой подходящий по тексту заголовка.
    """
    n = len(header)
    def htext(i):
        return _norm_header(header[i]['t'])
    cands = [i for i in range(n) if any(k in htext(i) for k in keys)]
    if not cands:
        return None
    yellow = [i for i in cands if header[i].get('y')]
    return (yellow or cands)[0]


@app.route('/api/import/docx', methods=['POST'])
def import_docx():
    """Импорт бойцов из таблицы Word-файла.

    Сопоставление по названию столбцов: ФИО, Весовая категория, Тренер.
    Остальное (пол, год, разряд, регион) — необязательно / по умолчанию.
    """
    upload = request.files.get('file')
    if upload is None or not upload.filename:
        return jsonify({'error': 'Файл не передан'}), 400

    default_gender = (request.form.get('gender') or 'мужской').strip()
    if default_gender not in ('мужской', 'женский'):
        default_gender = 'мужской'

    # Если установлен флаг replace — полностью очищаем текущий список спортсменов
    # (в т.ч. «родных»), чтобы остался только импортируемый состав.
    if request.form.get('replace', '0') == '1':
        conn = get_db()
        try:
            conn.execute("DELETE FROM tournament_athletes")
            conn.execute("DELETE FROM bracket_results")
            conn.execute("DELETE FROM lottery")
            conn.execute("DELETE FROM fights")
            conn.execute("DELETE FROM athletes")
            conn.commit()
        finally:
            conn.close()

    try:
        tables = parse_docx_tables(upload.read())
    except Exception as e:
        return jsonify({'error': f'Не удалось прочитать Word-файл: {e}'}), 400

    # Имена уже присутствующих спортсменов (name.lower -> id) — чтобы не плодить дубли,
    # но при этом привязать к турниру и уже существующих бойцов из файла.
    try:
        conn0 = get_db()
        existing_names = {r[0].lower(): r[1] for r in conn0.execute("SELECT name, id FROM athletes").fetchall()}
        conn0.close()
    except Exception:
        existing_names = {}

    imported = []
    errors = []
    new_ids = []
    skipped = 0  # уже были в базе, но привязаны к турниру
    for table in tables:
        if not table:
            continue
        # Ищем строку-заголовок (содержит ключевые слова или №)
        header_row = None
        for idx, row in enumerate(table):
            joined = ' '.join(_norm_header(c['t']) for c in row)
            if any(k in joined for k in ('фио', 'фамилия', 'имя', 'спортсмен')) or 'вес' in joined or '№' in joined:
                header_row = idx
                break
        if header_row is None:
            continue
        header = table[header_row]
        if not header:
            continue

        # Умное сопоставление полей по заголовкам (с приоритетом жёлтых столбцов)
        ci_name = _pick_col(header, 'фио', 'ф.и.о', 'имя', 'фамилия', 'спортсмен')
        ci_birth = _pick_col(header, 'год рождения', 'год рожд', 'дата рождения')
        ci_weight = _pick_col(header, 'вес', 'весовая категория', 'категория', 'вес, кг')
        ci_rank = _pick_col(header, 'разряд', 'квалификация')
        ci_coach = _pick_col(header, 'тренер', 'физрук')
        ci_region = _pick_col(header, 'субъект', 'город', 'регион')
        if ci_name is None:
            continue

        def cell(row, ci):
            return row[ci].get('t', '') if ci is not None and ci < len(row) else ''

        for row in table[header_row + 1:]:
            if ci_name >= len(row):
                continue
            name = cell(row, ci_name).strip()
            if not name:
                continue
            weight = _extract_weight(cell(row, ci_weight)) if ci_weight is not None else None
            birth_year = _extract_year(cell(row, ci_birth)) if ci_birth is not None else 0
            rank = cell(row, ci_rank).strip()
            region = cell(row, ci_region).strip()
            coach = cell(row, ci_coach).strip()

            # Если спортсмен с таким именем уже есть — не плодим дубль,
            # но привязываем его к турниру (по существующему id).
            if name.lower() in existing_names:
                new_ids.append(existing_names[name.lower()])
                skipped += 1  # уже есть — привязываем существующего, не создавая дубль
                continue

            conn = get_db()
            try:
                cur = conn.execute(
                    '''INSERT INTO athletes (name, birth_year, weight, gender, rank, region, coach)
                       VALUES (?, ?, ?, ?, ?, ?, ?)''',
                    (name, birth_year, weight if weight is not None else 0.0, default_gender, rank, region, coach)
                )
                conn.commit()
                existing_names[name.lower()] = cur.lastrowid
                new_ids.append(cur.lastrowid)
                imported.append({'name': name, 'weight': weight})
            except Exception as e:
                errors.append(f'{name}: {e}')
            finally:
                conn.close()

    if not imported and not errors and not new_ids:
        return jsonify({'error': 'Не найдено строк для импорта (проверьте, что в таблице есть столбцы ФИО/имя и заголовки)'}), 400

    return jsonify({'imported': len(imported), 'errors': len(errors), 'ids': new_ids, 'skipped': skipped, 'errorlist': errors[:20]})


    # ======================== API: ОТЧЁТ ГЛАВНОГО СУДЬИ (PDF) ========================

@app.route('/api/report/<int:tournament_id>/preview', methods=['GET'])
def report_preview(tournament_id):
    """Сформировать отчёт на основе шаблона ogl_tmp.docx, сконвертировать в PDF
    и вернуть HTML-предпросмотр (embed) для отображения в браузере без скачивания."""
    import io
    import base64
    from collections import defaultdict as _dd
    import re as _re

    if not _HAS_DOCX:
        return jsonify({'error': 'Библиотека Word не установлена (python-docx)'}), 500

    # Найти шаблон ogl_tmp.docx
    tmpl = None
    for cand in (os.path.join(PROJECT_DIR, 'ogl_tmp.docx'),
                 os.path.join(ASSETS_DIR, 'ogl_tmp.docx'),
                 os.path.join(_get_data_dir(), 'ogl_tmp.docx')):
        if os.path.exists(cand):
            tmpl = cand
            break
    if not tmpl:
        return jsonify({'error': 'Шаблон отчёта (ogl_tmp.docx) не найден'}), 500

    doc = _build_report_document(tournament_id, tmpl)
    if isinstance(doc, tuple):  # (json, status)
        return doc

    # === КОНВЕРТАЦИЯ .docx -> PDF ===
    try:
        # Сначала пробуем LibreOffice CLI (лучшее качество/точность сохранения шаблона)
        import subprocess, tempfile, os as _os
        _tmp_docx = tempfile.NamedTemporaryFile(delete=False, suffix='.docx', prefix='report_').name
        doc.save(_tmp_docx)

        soffice = None
        for cand in (os.environ.get('SOFFICE_PATH'),
                     'soffice', 'libreoffice', 'libreoffice.exe',
                     r'C:\Program Files\LibreOffice\program\soffice.exe',
                     r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'):
            if cand:
                p = cand
                try:
                    import shutil as _sh
                    if not _sh.which(p) and not _os.path.exists(p):
                        continue
                except Exception:
                    continue
                soffice = p
                break

        if soffice:
            _out_dir = tempfile.mkdtemp(prefix='report_pdf_')
            try:
                subprocess.run([soffice, '--headless', '--convert-to', 'pdf', '--outdir',
                                _out_dir, _tmp_docx],
                               check=True, capture_output=True, timeout=120)
                _pdf_path = _os.path.join(_out_dir,
                    _os.path.splitext(_os.path.basename(_tmp_docx))[0] + '.pdf')
                if _os.path.exists(_pdf_path):
                    with open(_pdf_path, 'rb') as f:
                        pdf_bytes = f.read()
                    _os.unlink(_tmp_docx)
                    import shutil as _shd
                    _shd.rmtree(_out_dir, ignore_errors=True)
                    return _pdf_preview_response(pdf_bytes, tournament_id)
            except Exception as e:
                import traceback; traceback.print_exc()
                # fallback ниже
                try:
                    _os.unlink(_tmp_docx)
                except Exception:
                    pass
    except Exception:
        import traceback; traceback.print_exc()

    # Важная уборка temp-файла и сообщение об ошибке, если конвертация не удалась
    try:
        _os.unlink(_tmp_docx)
    except Exception:
        pass
    return jsonify({'error': 'Не удалось конвертировать отчёт в PDF. Установите LibreOffice (soffice.exe) и добавьте его в PATH или задайте переменную окружения SOFFICE_PATH.'}), 500


def _pdf_preview_response(pdf_bytes, tournament_id):
    """Оборачивает байты PDF в полноценную HTML-страницу с embed
    (предпросмотр без скачивания). Возвращается напрямую как text/html,
    чтобы открываться во вкладке браузера, а не в окне приложения."""
    import base64
    from flask import Response
    pdf_base64 = base64.b64encode(pdf_bytes).decode('utf-8')
    preview_html = f'''
    <!DOCTYPE html>
    <html lang="ru">
    <head><meta charset="utf-8"><title>Предпросмотр отчёта</title></head>
    <body style="margin:0;padding:20px;background:#f5f5f5;font-family:Arial,sans-serif;">
        <div style="text-align:center;margin-bottom:15px;">
            <h2 style="margin:0 0 10px 0;">Предпросмотр отчёта</h2>
            <button onclick="window.print()" style="padding:8px 16px;cursor:pointer;">Печать PDF</button>
            <button onclick="downloadPdf()" style="padding:8px 16px;margin-left:10px;cursor:pointer;">Скачать PDF</button>
            <button onclick="window.close()" style="padding:8px 16px;margin-left:10px;cursor:pointer;">Закрыть</button>
        </div>
        <embed src="data:application/pdf;base64,{pdf_base64}"
               type="application/pdf"
               width="100%"
               height="800px"
               style="border:1px solid #ccc;background:#fff;">
    </body>
    <script>
        function downloadPdf() {{
            const b64 = '{pdf_base64}';
            const byteChars = atob(b64);
            const byteNums = new Array(byteChars.length);
            for (let i = 0; i < byteChars.length; i++) {{
                byteNums[i] = byteChars.charCodeAt(i);
            }}
            const arr = new Uint8Array(byteNums);
            const blob = new Blob([arr], {{type: 'application/pdf'}});
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = 'otchet_glavnogo_sudi_{tournament_id}.pdf';
            a.click();
        }}
    </script>
    '''
    return Response(preview_html, mimetype='text/html')


def _fid_act(v):
    """ID из объекта слота или числа."""
    return v['athlete_id'] if isinstance(v, dict) and 'athlete_id' in v else v


def _build_weight_rounds(slots, results, step=5.0):
    """Строит раунды и определяет призёров для одной весовой категории.

    Возвращает (rounds, champion_id, silver_id, [bronze_ids], names).
    rounds = список раундов; each раунд = список (a1_id, a2_id, winner_id, method).
    """
    # результаты по индексу пары
    results_by_key = {}
    for r in (results or []):
        results_by_key[(r['weight_class'], r['pair_index'])] = r

    rounds = []
    fr = []
    pidx = 0
    for i in range(0, len(slots) - 1, 2):
        a1 = slots[i]; a2 = slots[i + 1]
        r = results_by_key.get((a1['weight_class'], pidx))
        fr.append((_fid_act(a1), _fid_act(a2), r['winner_id'] if r else None, (r['method'] if r else '')))
        pidx += 1
    if len(slots) % 2 != 0:
        a1 = slots[-1]
        fr.append((_fid_act(a1), None, None, ''))  # ожидает соперника — не авто-победитель
    rounds.append(fr)

    round_num = 1
    while True:
        prev = rounds[-1]
        winners = [f[2] for f in prev]
        if any(w is None for w in winners) or len(winners) <= 1:
            break
        round_num += 1
        offset = (round_num - 1) * 1000
        nxt = []
        for i in range(0, len(winners) - 1, 2):
            a1 = winners[i]; a2 = winners[i + 1]
            r = results_by_key.get((slots[0]['weight_class'], offset + i))
            nxt.append((a1, a2, r['winner_id'] if r else None, (r['method'] if r else '')))
        if not nxt:
            break
        rounds.append(nxt)

    wc = slots[0]['weight_class']
    champion = None; silver = None; bronzes = []
    top = rounds[-1]
    any_unresolved = any((f[1] is not None and f[2] is None) for rd in rounds for f in rd)
    top_winner = None
    if not any_unresolved and len(top) == 1 and top[0][2] is not None:
        top_winner = top[0][2]
    champion = top_winner
    if champion is not None and len(top) == 1 and top[0][1] is not None:
        silver = top[0][0] if top[0][2] == top[0][1] else top[0][1]
    if len(rounds) >= 2:
        semi = rounds[-2]
        for f in semi:
            if f[1] is None or f[2] is None:
                continue
            loser = f[1] if f[2] == f[0] else f[0]
            if loser is not None and loser != champion and len(bronzes) < 2:
                bronzes.append(loser)
    return rounds, champion, silver, bronzes


def _build_report_document(tournament_id, tmpl):
    """Собирает отчёт из шаблона ogl_tmp.docx и данных приложения.

    Возвращает объект Document (python-docx) или кортеж (jsonify-ответ, status).
    Заполняет плейсхолдеры ({TITLE}, {DATE}, {DATES}, {CITY}, {VENUE},
    {REFEREE}, {SECRETARY}, _145_ = число участников) по всему телу,
    таблицам и колонтитулам; затем синхронизирует таблицы протоколов:
      - Регламент соревнований (кол-во участников и боёв по категориям);
      - Протоколы победителей и призёров по категориям;
      - Заголовки-подзаголовки протоколов (название турнира).
    """
    import re as _re
    import copy as _copy
    from collections import defaultdict as _dd
    from docx.oxml import parse_xml as _parse_xml

    conn = get_db()
    t = row_to_dict(conn.execute("SELECT * FROM tournaments WHERE id=?", (tournament_id,)).fetchone())
    if t is None:
        conn.close()
        return (jsonify({'error': 'Соревнование не найдено'}), 404)

    n = conn.execute("SELECT COUNT(*) c FROM tournament_athletes WHERE tournament_id=?", (tournament_id,)).fetchone()['c'] or 0
    ath_ids = [r['athlete_id'] for r in conn.execute(
        "SELECT athlete_id FROM tournament_athletes WHERE tournament_id=?", (tournament_id,)).fetchall()]
    athletes = []
    if ath_ids:
        _ph = ','.join('?' * len(ath_ids))
        athletes = rows_to_list(conn.execute(f"SELECT * FROM athletes WHERE id IN ({_ph})", ath_ids).fetchall())
    lottery = rows_to_list(conn.execute(
        "SELECT * FROM lottery WHERE tournament_id=? ORDER BY weight_class, slot", (tournament_id,)).fetchall())
    results = rows_to_list(conn.execute(
        "SELECT * FROM bracket_results WHERE tournament_id=?", (tournament_id,)).fetchall())
    conn.close()

    doc = Document(tmpl)

    groups = _dd(list); order = []
    for s in lottery:
        if s['weight_class'] not in groups:
            order.append(s['weight_class'])
        groups[s['weight_class']].append(s)
    for wc in groups:
        groups[wc].sort(key=lambda x: x['slot'])
    order.sort(key=lambda wc: float((_re.search(r'(\d+(?:\.\d+)?)', wc or '') or _re.search(r'0', ''))[1] or 0))

    def _fmt_date(s):
        p = (s or '').split('-')
        return f'{p[2]}.{p[1]}.{p[0]}' if len(p) == 3 else (s or '').replace('-', '.')
    d1 = _fmt_date(t.get('date_from') or '')
    d2 = _fmt_date(t.get('date_to') or t.get('date_from') or '')
    dates = f'{d1} — {d2}' if d2 and d2 != d1 else d1
    title = t.get('name') or ''
    city = t.get('venue') or ''
    ref = t.get('chief_referee') or ''
    sec = t.get('chief_secretary') or ''

    # Карта спортсменов: id -> athlete
    by_id = {a['id']: a for a in athletes}

    def para_replace(para, old, new):
        text = ''.join(r.text for r in para.runs)
        if old not in text:
            return False
        merged = text.replace(old, new)
        if para.runs:
            para.runs[0].text = merged
            for r in para.runs[1:]:
                r.text = ''
        return True

    def fill_para(para):
        para_replace(para, '_145_', str(n))
        para_replace(para, '{TITLE}', title or '____')
        para_replace(para, '{DATES}', dates)
        para_replace(para, '{DATE}', d1)
        para_replace(para, 'г. {CITY}', ('г. ' + city) if city else 'г. ____')
        para_replace(para, '{CITY}', city or '____')
        para_replace(para, '{VENUE}', city or '')
        para_replace(para, '{REFEREE}', ref or '________')
        para_replace(para, '{SECRETARY}', sec or '________')

    def fill_cell(cell):
        for p in cell.paragraphs:
            fill_para(p)

    for para in doc.paragraphs:
        fill_para(para)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                fill_cell(cell)

    # Колонтитулы
    for section in doc.sections:
        for hf in (section.header, section.footer, section.first_page_header, section.first_page_footer):
            try:
                for p in hf.paragraphs:
                    fill_para(p)
                for tbl in hf.tables:
                    for row in tbl.rows:
                        for cell in row.cells:
                            fill_cell(cell)
            except Exception:
                pass

    # Главный судья/секретарь в первой таблице (судейская коллегия)
    try:
        _t0 = doc.tables[0]
        for _r in _t0.rows:
            _j = ' '.join(x.text for x in _r.cells)
            if 'Главный судья' in _j and len(_r.cells) > 1:
                _r.cells[1].text = ref or '________'
            elif 'Главный секретарь' in _j and len(_r.cells) > 1:
                _r.cells[1].text = sec or '________'
    except Exception:
        pass

    # --- Заголовки протоколов: подзаголовки возрастных групп -> название турнира ---
    # Строки вида «ДЕВОЧКИ 2014-2015гг.р.» / «ЮНОШИ 2016г.р.» обновляются названием турнира,
    # чтобы в шапках протоколов фигурировало актуальное название соревнования.
    for para in doc.paragraphs:
        tx = para.text.strip()
        if _re.match(r'^[А-ЯЁ\s]+(?:\d{4}[-–]\d{4}|\d{4})г{0,1}\.{0,1}р\.?\s*$', tx) and 'ПРОТОКОЛ' not in tx and title:
            para_replace(para, tx, title)

    # --- Регламент соревнований ---
    # Шаблон содержит: строка-заголовок (r0), строка-группа «Юноши {год рождения}»,
    # строка-данных с маркерами {весизжеребьевки} {количествоизжеребьевки} {всегобоев}.
    # Строим группы по полу (мужской→Юноши, женский→Девушки) и возрастному диапазону
    # (непрерывные диапазоны лет); внутри каждой группы — по строке на весовую категорию.
    try:
        from docx.oxml.ns import qn as _qn
        _reg = None
        for _tbl in doc.tables:
            if len(_tbl.rows) > 0:
                _hdr = ' '.join(c.text for c in _tbl.rows[0].cells)
                if 'Кол-во уч-ков' in _hdr or 'Количество участников' in _hdr:
                    if 'Кол-во боёв' in _hdr or 'боёв' in _hdr:
                        _reg = _tbl
                        break
        if _reg is not None and len(_reg.rows) >= 3:
            _group_tr = _copy.deepcopy(_reg.rows[1]._tr)   # «Юноши {год рождения}»
            _data_tr = _copy.deepcopy(_reg.rows[2]._tr)    # {весизжеребьевки}...
            def _reg_replace_para(para_el, mapping):
                """Объединяет текст абзаца, заменяет маркеры целиком (маркер может быть
                разбит на несколько w:t), результат кладёт в первый run."""
                _runs = para_el.findall(_qn('w:r'))
                if not _runs:
                    return
                _full = ''.join(_rt.text or '' for _r in _runs for _rt in _r.findall(_qn('w:t')))
                _changed = False
                for _k, _v in mapping.items():
                    if _k in _full:
                        _full = _full.replace(_k, _v)
                        _changed = True
                if _changed:
                    _runs[0].findall(_qn('w:t'))[0].text = _full if _runs[0].findall(_qn('w:t')) else None
                    if _runs[0].findall(_qn('w:t')):
                        _runs[0].findall(_qn('w:t'))[0].text = _full
                        for _r in _runs[1:]:
                            for _rt in _r.findall(_qn('w:t')):
                                _rt.text = ''

            def _reg_markers_in_el(el, mapping):
                for _p in el.iter(_qn('w:p')):
                    _reg_replace_para(_p, mapping)
            def _reg_cell_text(tr_el, idx):
                tcs = tr_el.findall(_qn('w:tc'))
                if idx < len(tcs):
                    return ''.join(_n.text or '' for _n in tcs[idx].iter(_qn('w:t')))
                return ''
            # Читаем пол каждого спортсмена лота и его категории
            def _slot_gender(slot):
                _a = by_id.get(slot['athlete_id'])
                if not _a:
                    return None
                _g = (_a.get('gender') or '').strip()
                # нормализуем: мужской/муж/м → 'мужской', женский/жен/ж → 'женский'
                _gl = _g.lower()
                if 'муж' in _gl or _gl in ('м', 'm', 'male'):
                    return 'мужской'
                if 'жен' in _gl or _gl in ('ж', 'f', 'female'):
                    return 'женский'
                return _g
            def _slot_year(slot):
                _a = by_id.get(slot['athlete_id'])
                y = (_a.get('birth_year') if _a else None)
                try:
                    return int(str(y).strip())
                except (TypeError, ValueError):
                    return None

            # --- Собираем данные по категориям ---
            # cat = order; для каждой категории — пол (по первому лоту) и годы участников
            _cat_gender = {}
            _cat_years = {}
            for _wc in order:
                _slots = groups[_wc]
                if not _slots:
                    continue
                _g = _slot_gender(_slots[0])
                _cat_gender[_wc] = _g or ''
                _ys = []
                for _s in _slots:
                    _y = _slot_year(_s)
                    if _y is not None:
                        _ys.append(_y)
                _cat_years[_wc] = _ys

            # --- Группы: (слово_пола, диапазон_лет) -> [категории] ---
            # Слово пола: мужской → «Юноши», женский → «Девушки»
            def _gender_word(g):
                return 'Юноши' if g == 'мужской' else ('Девушки' if g == 'женский' else (g or ''))
            def _range_label(vals):
                if not vals:
                    return ''
                _ms = min(vals); _mx = max(vals)
                return str(_ms) if _ms == _mx else f'{_ms}-{_mx}'
            # Сначала: по каждому полу строим НЕПРЕРЫВНЫЕ диапазоны лет из всех участников.
            # Диапазон = последовательность годов без пропусков.
            def _continuous_ranges(years_set):
                s = sorted(set(years_set))
                if not s:
                    return []
                ranges = []
                start = prev = s[0]
                for y in s[1:]:
                    if y == prev + 1:
                        prev = y
                    else:
                        ranges.append((start, prev))
                        start = prev = y
                ranges.append((start, prev))
                return ranges
            # Соберём все годы по полу
            _all_years_by_g = {}
            for _wc in order:
                _g = _cat_gender.get(_wc) or ''
                for _y in (_cat_years.get(_wc) or []):
                    _all_years_by_g.setdefault(_g, set()).add(_y)
            # диапазоны по полу (слово пола)
            _gender_ranges = {}   # gword -> [(y0,y1), ...]
            for _g, _yset in _all_years_by_g.items():
                _gender_ranges[_gender_word(_g)] = _continuous_ranges(_yset)
            def _cat_in_range(wc, gword):
                """Возвращает диапазон (y0,y1), в который попадают годы категории."""
                _ys = set(_cat_years.get(wc) or [])
                if not _ys:
                    return (None, None)
                for (a, b) in _gender_ranges.get(gword, []):
                    if all(a <= y <= b for y in _ys):
                        return (a, b)
                # если годы не из одного диапазона — берём (min,max)
                return (min(_ys), max(_ys))
            _groups_map = {}   # (gword, (y0,y1)) -> [wc]
            for _wc in order:
                _g = _gender_word(_cat_gender.get(_wc) or '')
                _r = _cat_in_range(_wc, _g)
                _groups_map.setdefault((_g, _r), []).append(_wc)
            # Сортируем группы: пол (Юноши первыми), затем по возрасту, затем по минимальному весу
            def _gkey(item):
                (gw, yrange), wcs = item
                _grank = 0 if gw == 'Юноши' else (1 if gw == 'Девушки' else 2)
                _y0 = yrange[0] if yrange and yrange[0] is not None else 10e9
                _minw = None
                for _wc in wcs:
                    _m = _re.search(r'([\d.]+)', _wc or '')
                    if _m:
                        _v = float(_m.group(1))
                        if _minw is None or _v < _minw:
                            _minw = _v
                return (_grank, _y0, _minw if _minw is not None else 10e9)
            _groups_sorted = sorted(_groups_map.items(), key=_gkey)

            # Удаляем все строки кроме заголовка
            for _r in list(_reg.rows[1:]):
                _reg._tbl.remove(_r._tr)

            # --- Вставляем группы ---
            for (gword, yrange), wcs in _groups_sorted:
                _g = _copy.deepcopy(_group_tr)
                _rlabel = ''
                if yrange and yrange[0] is not None:
                    _rlabel = _range_label(list(yrange))
                _reg_markers_in_el(_g, {'{год рождения}': _rlabel})
                # заменяем слово «Юноши» на актуальное (Юноши/Девушки)
                _g_txt = ''.join(_n.text or '' for _n in _g.iter(_qn('w:t')))
                if gword and gword != 'Юноши':
                    # заменим первое вхождение «Юноши» на gword
                    for _t in _g.iter(_qn('w:t')):
                        if _t.text and 'Юноши' in _t.text:
                            _t.text = _t.text.replace('Юноши', gword, 1)
                            break
                _reg._tbl.append(_g)
                for _wc in wcs:
                    _cnt = len(groups[_wc])
                    _fights = 0
                    if groups[_wc]:
                        _rounds = _build_weight_rounds(groups[_wc], results)[0]
                        _fights = sum(len(rd) for rd in _rounds)
                    _v = _re.search(r'мужской\s+([\d.,+-]+)|\s+([\d.,+-]+)\s*кг', _wc or '')
                    _weight = _v.group(1) if _v and _v.group(1) else (_v.group(2) if _v and _v.group(2) else '')
                    _d = _copy.deepcopy(_data_tr)
                    _reg_markers_in_el(_d, {
                        '{весизжеребьевки}': _weight or _wc,
                        '{количествоизжеребьевки}': str(_cnt),
                        '{всегобоев}': str(_fights),
                    })
                    _reg._tbl.append(_d)
    except Exception as e:
        import traceback; traceback.print_exc()

    # --- Протоколы победителей и призёров (таблицы с первой колонкой «N место») ---
    # В шаблоне каждая таблица призёров начинается со строки-заголовка вида «44 кг».
    # Для каждой такой таблицы определяем весовую категорию, ищем её в данных приложения
    # и заполняем реальными призёрами (1/2/3 место), если чемпион категории определён.
    def _wc_find(weight_float):
        for wc in order:
            m = _re.search(r'(\d+(?:\.\d+)?)', wc or '')
            if m and abs(float(m[1]) - weight_float) < 0.01:
                return wc
        return None

    def _prizers_for(wc_key):
        """Возвращает [(place_str, name, region), ...] для категории."""
        g = groups.get(wc_key)
        if not g:
            return []
        rounds, champion, silver, bronzes = _build_weight_rounds(g, results)
        if champion is None:
            return []
        out = []
        c = by_id.get(champion)
        if c:
            out.append(('1', c['name'], c.get('region') or ''))
        if silver is not None:
            s = by_id.get(silver)
            if s:
                out.append(('2', s['name'], s.get('region') or ''))
        for b in bronzes:
            bobj = by_id.get(b)
            if bobj:
                out.append(('3', bobj['name'], bobj.get('region') or ''))
        return out

    try:
        for tbl in doc.tables:
            cells0 = [c.text.strip() for c in tbl.rows[0].cells] if tbl.rows else []
            if not cells0:
                continue
            # Таблицей призёров считаем таблицу, где в строке-заголовке есть «N кг»
            # (и строка дублирует одну и ту же категорию по всем объединённым ячейкам).
            wmatch = None
            for c0 in cells0:
                mm = _re.match(r'^\s*(\d+(?:\.\d+)?)\s*кг\s*$', c0)
                if mm:
                    wmatch = float(mm.group(1))
                    break
            if wmatch is None:
                continue
            # Все строки данных (кроме строки заголовка категории) удаляем всегда,
            # чтобы не просачивались данные из шаблона по категориям, которых нет в турнире.
            rows_to_del = list(tbl.rows[1:])
            for r in rows_to_del:
                tbl._tbl.remove(r._tr)
            wc_key = _wc_find(wmatch)
            if wc_key is None:
                continue
            prizes = _prizers_for(wc_key)
            for place, name, region in prizes:
                cells = tbl.add_row().cells
                cells[0].text = place
                if len(cells) > 1:
                    cells[1].text = 'место'
                if len(cells) > 2:
                    cells[2].text = name
                if len(cells) > 3:
                    cells[3].text = ''
                if len(cells) > 4:
                    cells[4].text = ''
                if len(cells) > 5:
                    cells[5].text = region
    except Exception as e:
        import traceback; traceback.print_exc()

    # --- Протоколы графиков: динамическая турнирная сетка по весовым категориям ---
    # В шаблоне есть ОДНА таблица-прототип (бракет на N слотов с плейсхолдерами SLOT*).
    # Под каждую весовую категорию с участниками строим свою таблицу-дерево:
    # участники по слотам из жеребьёвки, победители пар и чемпион — из bracket_results.
    from docx.oxml.ns import qn as _qn

    try:
        _WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        def _qw(tag):
            return '{%s}%s' % (_WNS, tag)

        # Находим таблицу-прототип графиков (содержит плейсхолдеры SLOT/SLOTS)
        proto_tbl_el = None
        for _tb in doc.element.body.iter(_qw('tbl')):
            _txt_all = ''.join(_n.text or '' for _n in _tb.iter(_qw('t')))
            if '{SLOT' in _txt_all or '{SLOTS' in _txt_all:
                proto_tbl_el = _tb
                break
        if proto_tbl_el is not None:
            _rows = proto_tbl_el.findall(_qw('tr'))
            def _tc(r, i):
                return _cells_at(r)[i]
            def _cells_at(tr_el):
                return tr_el.findall(_qw('tc'))
            # Стилевые атомы из прототипа (N=4, строки 0..7, как T12)
            _r0 = _rows[0]; _r1 = _rows[1]; _r3 = _rows[3]
            _atoms = {
                'num': _copy.deepcopy(_cells_at(_r0)[0]),
                'name': _copy.deepcopy(_cells_at(_r0)[1]),
                'club': _copy.deepcopy(_cells_at(_r1)[1]),
                'spacer': _copy.deepcopy(_cells_at(_r1)[2]),
                'year': _copy.deepcopy(_cells_at(_r1)[3]),
                'winner': _copy.deepcopy(_cells_at(_r1)[4]),
                'champ': _copy.deepcopy(_cells_at(_r3)[5]),
            }
            # champ должен быть жёлтым (FFFF00); champ_plain — без заливки
            def _ensure_shd(tc, fill):
                tcPr = tc.find(_qw('tcPr'))
                if tcPr is None:
                    tcPr = _parse_xml('<w:tcPr xmlns:w="%s"/>' % _WNS)
                    tc.insert(0, tcPr)
                shd = tcPr.find(_qw('shd'))
                if shd is None:
                    shd = _parse_xml('<w:shd xmlns:w="%s"/>' % _WNS)
                    tcPr.append(shd)
                shd.set(_qw('val'), 'clear'); shd.set(_qw('color'), 'auto'); shd.set(_qw('fill'), fill)
            def _clear_shd(tc):
                tcPr = tc.find(_qw('tcPr'))
                if tcPr is not None and tcPr.find(_qw('shd')) is not None:
                    tcPr.remove(tcPr.find(_qw('shd')))
            _ensure_shd(_atoms['champ'], 'FFFF00')
            _atoms['champ_plain'] = _copy.deepcopy(_atoms['champ'])
            _clear_shd(_atoms['champ_plain'])

            def _set_tc_text(tc, text):
                ps = tc.findall(_qw('p'))
                if not ps:
                    p = _parse_xml('<w:p xmlns:w="%s"/>' % _WNS)
                    tc.append(p); ps = [p]
                p = ps[0]
                runs = p.findall(_qw('r'))
                if runs:
                    r0 = runs[0]
                    t = r0.find(_qw('t'))
                    if t is None:
                        t = _parse_xml('<w:t xmlns:w="%s"/>' % _WNS)
                        r0.append(t)
                    t.text = text
                    for r in runs[1:]:
                        p.remove(r)
                else:
                    _r = _parse_xml('<w:r xmlns:w="%s"><w:t>%s</w:t></w:r>' % (_WNS, text))
                    p.append(_r)

            def _clear_tc(tc):
                for p in tc.findall(_qw('p')):
                    for r in p.findall(_qw('r')):
                        p.remove(r)

            def _set_borders(tc, sides, val='single', sz='4'):
                """sides: iterable {top,bottom,left,right} — какие границы проставить."""
                tcPr = tc.find(_qw('tcPr'))
                if tcPr is None:
                    tcPr = _parse_xml('<w:tcPr xmlns:w="%s"/>' % _WNS)
                    tc.insert(0, tcPr)
                borders = tcPr.find(_qw('tcBorders'))
                if borders is None:
                    borders = _parse_xml('<w:tcBorders xmlns:w="%s"/>' % _WNS)
                    _tn = tcPr.find(_qw('tcW'))
                    if _tn is not None:
                        _tn.addnext(borders)
                    else:
                        tcPr.append(borders)
                for sd in sides:
                    el = borders.find(_qw(sd))
                    if el is None:
                        el = _parse_xml('<w:%s xmlns:w="%s"/>' % (sd, _WNS))
                        borders.append(el)
                    el.set(_qw('val'), val); el.set(_qw('sz'), sz)
                    el.set(_qw('space'), '0'); el.set(_qw('color'), 'auto')
                for el in list(borders):
                    sd = el.tag.split('}')[-1]
                    if sd not in sides:
                        borders.remove(el)

            def _build_bracket(members, results_map):
                """members: [(name, club, year), ...]; results_map: pair_index -> name."""
                n = len(members)
                total_slots = 2 ** math.ceil(math.log2(n)) if n > 1 else 1
                num_rounds = int(math.log2(total_slots)) if total_slots > 1 else 1
                champ_col = num_rounds - 1
                m_final = 2 ** num_rounds
                top_ws = m_final // 4 - 1 if num_rounds >= 1 else 0
                bot_ws = m_final // 4 * 3 - 1 if num_rounds >= 1 else 0
                grid_widths = [540, 2340, 900, 1323]
                round1_idx = len(grid_widths)
                for _r in range(num_rounds):
                    grid_widths.append(1701 if _r < num_rounds - 1 else 1843)
                total_w = sum(grid_widths)
                tblPr = _parse_xml(
                    '<w:tblPr xmlns:w="%s"><w:tblW w:w="%s" w:type="dxa"/>'
                    '<w:tblLayout w:type="fixed"/>'
                    '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" '
                    'w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>' % (_WNS, total_w))
                grid_el = _parse_xml('<w:tblGrid xmlns:w="%s"/>' % _WNS)
                for _wd in grid_widths:
                    grid_el.append(_parse_xml('<w:gridCol xmlns:w="%s" w:w="%s"/>' % (_WNS, _wd)))
                tbl = _parse_xml('<w:tbl xmlns:w="%s"/>' % _WNS)
                tbl.append(tblPr); tbl.append(grid_el)
                # победители по слотам
                winner_cells = {}
                for _r in range(num_rounds):
                    n_groups = total_slots // (2 ** (_r + 1))
                    for _g in range(n_groups):
                        _s = (_g * (2 ** (_r + 1)) + (2 ** _r - 1)) if _r > 0 else (_g * 2)
                        _row = 2 * _s + 1
                        _col = round1_idx + _r
                        _pi = _g if _r == 0 else (_r * 1000 + _g)
                        winner_cells[(_row, _col)] = results_map.get(_pi, '')
                def _outcome_borders(_c, slot_idx, is_name_row):
                    """Границы ячейки в outcome-колонке (победители/чемпион).
                    Точное воспроизведение вилки прототипа T12."""
                    if _c == champ_col:
                        # чемпионская колонка: ствол [left] + жёлтая ячейка [bottom,left]
                        if not is_name_row and is_champ_slot(slot_idx):
                            return ['bottom', 'left']
                        if num_rounds >= 2 and (top_ws < slot_idx <= bot_ws):
                            return ['left']
                        return []
                    # --- колонки победителей ---
                    span = 2 ** (_c + 1)          # слотов на группу
                    _g = slot_idx // span
                    s_node = _g * span + (span // 2 - 1)   # слот победителя группы (<- win slot parity)
                    # пары рисуются только для реально занятых слотов (BYE-напарник не создаёт вилку)
                    # пара round1: (s_node, s_node+1) оба должны существовать
                    pair_complete = (s_node + 1 < n) if _c == 0 else True
                    if not pair_complete:
                        return []
                    if not is_name_row and slot_idx == s_node:
                        return ['left', 'bottom']          # бокс победителя (клуб-строка)
                    if is_name_row and slot_idx == s_node + 1:
                        return ['left', 'top']             # угловая скобка (имя-строка след. слота)
                    if not is_name_row and slot_idx == total_slots - 1:
                        return ['left']                    # продолжение ствола вниз
                    return []

                def is_champ_slot(sidx):
                    return (sidx * 2 + 1 == total_slots - 1)

                def _make_row(is_name_row, slot_idx, sl):
                    tr = _parse_xml('<w:tr xmlns:w="%s"/>' % _WNS)
                    out = []
                    if is_name_row:
                        num_tc = _copy.deepcopy(_atoms['num']); _set_tc_text(num_tc, str(slot_idx + 1) if sl else '')
                        name_tc = _copy.deepcopy(_atoms['name']); _set_tc_text(name_tc, sl[0] if sl else '')
                        yr_tc = _copy.deepcopy(_atoms['year']); _clear_tc(yr_tc)
                        # right-граница год-колонки = там, где у round1 winner есть LEFT (прерывистая вилка)
                        if 'left' in _outcome_borders(0, slot_idx, True):
                            _set_borders(yr_tc, ['right'])
                        else:
                            _set_borders(yr_tc, [])
                        out = [num_tc, name_tc, yr_tc]
                        for _c in range(num_rounds):
                            base = _atoms['winner'] if _c < num_rounds - 1 else _atoms['champ_plain']
                            _oc = _copy.deepcopy(base); _clear_tc(_oc)
                            _set_borders(_oc, _outcome_borders(_c, slot_idx, True))
                            out.append(_oc)
                    else:
                        num_tc = _copy.deepcopy(_atoms['num']); _clear_tc(num_tc)
                        club_tc = _copy.deepcopy(_atoms['club']); _set_tc_text(club_tc, sl[1] if sl else '')
                        sp_tc = _copy.deepcopy(_atoms['spacer']); _clear_tc(sp_tc)
                        yr_tc = _copy.deepcopy(_atoms['year']); _set_tc_text(yr_tc, sl[2] if sl else '')
                        if 'left' in _outcome_borders(0, slot_idx, False):
                            _set_borders(yr_tc, ['right'])
                        else:
                            _set_borders(yr_tc, [])
                        out = [num_tc, club_tc, sp_tc, yr_tc]
                        champion_row = total_slots - 1
                        for _c in range(num_rounds):
                            _col = round1_idx + _c
                            is_champ_cell = (_c == num_rounds - 1) and (slot_idx * 2 + 1 == champion_row)
                            base = _atoms['winner'] if _c < num_rounds - 1 else (
                                _atoms['champ'] if is_champ_cell else _atoms['champ_plain'])
                            tc = _copy.deepcopy(base); _clear_tc(tc)
                            _rr = 2 * slot_idx + 1
                            _text = winner_cells.get((_rr, _col), '')
                            _set_borders(tc, _outcome_borders(_c, slot_idx, False))
                            _set_tc_text(tc, _text)
                            out.append(tc)
                    for tc in out:
                        tr.append(tc)
                    return tr
                for _i in range(total_slots):
                    sl = members[_i] if _i < n else None
                    tbl.append(_make_row(True, _i, sl))
                    tbl.append(_make_row(False, _i, sl))
                return tbl

            # Готовим данные по категориям из жеребьёвки
            # results: [(pair_index, winner_id)] по категории
            results_by_wc = {}
            for _r in results:
                results_by_wc.setdefault(_r['weight_class'], {})[_r['pair_index']] = _r['winner_id']

            # Строим новые элементы: (заголовок-абзац «Весовая категория N кг», таблица)
            new_gfx = []   # [(wc_para, tbl, weight_value)]
            for wc in order:
                slots = groups.get(wc) or []
                slots = sorted(slots, key=lambda x: x['slot'])
                if not slots:
                    continue
                members = []
                for _s in slots:
                    a = by_id.get(_s['athlete_id'])
                    name = a['name'] if a else ('#' + str(_s['athlete_id']))
                    club = ((a or {}).get('region') or '').replace('Калмыкия', '').strip() if a else ''
                    year = str(a.get('birth_year') or '') if a else ''
                    members.append((name, club, year))
                # результаты победителей по слотам: pair_index -> имя
                rm = {}
                for _pi, _wid in (results_by_wc.get(wc) or {}).items():
                    _a = by_id.get(_wid)
                    rm[_pi] = _a['name'] if _a else ''
                _mm = _re.search(r'(\d+(?:\.\d+)?)', wc or '')
                _wv = float(_mm.group(1)) if _mm else None
                # wc вида «мужской 44 кг» уже грамматически корректен как заголовок
                wc_para = _parse_xml(
                    '<w:p xmlns:w="%s"><w:r><w:t>Весовая категория %s</w:t></w:r></w:p>' % (_WNS, wc))
                new_gfx.append((wc_para, _build_bracket(members, rm), _wv))

        # Заменяем прототипную таблицу сгенерированными парами (абзац «Весовая категория» + таблица).
        # Абзац {WEIGHT_CLASS} перед прототипом в шаблоне не нужен — удаляем его.
        _gfx_ready = ('new_gfx' in dir() and new_gfx and proto_tbl_el is not None)
        if _gfx_ready:
            _body = doc.element.body
            _proto_parent = proto_tbl_el.getparent()
            # позиция вставки = индекс прототипа в родителе (ДО удаления)
            _pos = list(_proto_parent).index(proto_tbl_el)
            # удаляем соседний абзац {WEIGHT_CLASS} прямо перед прототипом
            _prev_sib = proto_tbl_el.getprevious()
            _proto_parent.remove(proto_tbl_el)
            if _prev_sib is not None and _prev_sib.tag == _qw('p'):
                _ptxt = ''.join(_n.text or '' for _n in _prev_sib.iter(_qw('t')))
                if '{WEIGHT_CLASS}' in _ptxt:
                    _proto_parent.remove(_prev_sib)
                    _pos -= 1
            # вставляем пары (абзац + таблица) на место прототипа
            _children_now = list(_proto_parent)
            for _wp, _tb, _w in new_gfx:
                _children_now.insert(_pos, _tb)
                _children_now.insert(_pos, _wp)
                _pos += 2
            # пересобираем родителя в новом порядке
            for _el in list(_proto_parent):
                _proto_parent.remove(_el)
            for _el in _children_now:
                _proto_parent.append(_el)
    except Exception as e:
        import traceback; traceback.print_exc()

    # === Сортировка весовых категорий по возрастанию в протоколах ===
    # Призёры: категории идут парами (заголовок «ПРОТОКОЛ ПОБЕДИТЕЛЕЙ» + таблица «N кг»)
    # и в шаблоне расположены хаотично — сортируем пары по весу.
    # Графики: таблицы с привязкой к абзацу «Весовая категория N кг» сортируем по весу.
    try:
        from docx.oxml.ns import qn as _sqn
        _body = doc.element.body
        _children = list(_body)

        def _txt(el):
            return ''.join(t.text or '' for t in el.iter(_sqn('w:t')))

        def _is_win_hdr(el):
            return el.tag == _sqn('w:p') and 'ПРОТОКОЛ ПОБЕДИТЕЛЕЙ' in _txt(el)

        def _prz_w(el):
            if el.tag != _sqn('w:tbl'):
                return None
            _tr = el.find(_sqn('w:tr'))
            if _tr is None:
                return None
            _tc = _tr.find(_sqn('w:tc'))
            if _tc is None:
                return None
            _m = _re.match(r'^\s*(\d+(?:\.\d+)?)\s*кг\s*$', _txt(_tc).strip())
            return float(_m.group(1)) if _m else None

        # --- Сортировка призёров ---
        _win_units = []
        _i = 0
        while _i < len(_children):
            if _is_win_hdr(_children[_i]) and _i + 1 < len(_children):
                _w = _prz_w(_children[_i + 1])
                if _w is not None:
                    _win_units.append([_i, _w, _children[_i], _children[_i + 1]])
                    _i += 2
                    continue
            _i += 1
        if _win_units:
            _sorted_units = sorted(_win_units, key=lambda u: u[1])
            _positions = [u[0] for u in _win_units]
            _elems = []
            for u in _sorted_units:
                _elems.append(u[2])
                _elems.append(u[3])
            for _k, _p in enumerate(_positions):
                _children[_p] = _elems[2 * _k]
                _children[_p + 1] = _elems[2 * _k + 1]

        # --- Сортировка графиков по возрастанию ВНУТРИ каждой секции ---
        # Секция начинается с абзаца «ПРОТОКОЛ ГРАФИКОВ СОРЕВНОВАНИЙ»; внутри секции
        # идут пары (абзац «Весовая категория N кг» + таблица-сетка). Сортируем пары каждой
        # секции отдельно, не перемешивая секции между собой.
        def _wc_w(el):
            if el.tag != _sqn('w:p'):
                return None
            _m = _re.match(r'^Весовая\s+категория\s+(\d+(?:\.\d+)?)\s*кг', _txt(el).strip(), _re.IGNORECASE)
            return float(_m.group(1)) if _m else None

        def _is_gfx_hdr(el):
            return el.tag == _sqn('w:p') and 'ПРОТОКОЛ ГРАФИКОВ' in _txt(el)

        # Находим границы секций (по заголовкам GHDR) и обрабатываем каждую
        _gfx_sections = []
        _i = 0
        while _i < len(_children):
            if _is_gfx_hdr(_children[_i]):
                _start = _i + 1
                _j = _start
                while _j < len(_children) and not _is_gfx_hdr(_children[_j]):
                    _j += 1
                _gfx_sections.append((_start, _j))
                _i = _j
            else:
                _i += 1

        for _start, _end in _gfx_sections:
            # собираем пары (wc_para, tbl) в пределах секции; между абзацем категории
            # и таблицей могут стоять пустые абзацы — пропускаем их.
            _pairs = []
            _i = _start
            while _i < _end:
                _w = _wc_w(_children[_i])
                if _w is not None:
                    # ищем следующую таблицу в секции
                    _k = _i + 1
                    while _k < _end and _children[_k].tag != _sqn('w:tbl'):
                        _k += 1
                    if _k < _end:
                        _pairs.append([_i, _w, _children[_i], _children[_k]])
                        _i = _k + 1
                        continue
                _i += 1
            if not _pairs:
                continue
            _sorted_pairs = sorted(_pairs, key=lambda u: u[1])
            _positions = [u[0] for u in _pairs]
            _elems = []
            for u in _sorted_pairs:
                _elems.append(u[2])
                _elems.append(u[3])
            for _k, _p in enumerate(_positions):
                _children[_p] = _elems[2 * _k]
                _children[_p + 1] = _elems[2 * _k + 1]

        # Записать обратно в тело
        for _el in list(_body):
            _body.remove(_el)
        for _el in _children:
            _body.append(_el)
    except Exception as e:
        import traceback; traceback.print_exc()

    return doc


@app.route('/api/report/<int:tournament_id>', methods=['GET'])
def report_docx(tournament_id):
    """Сформировать отчёт главного судьи на основе ogl_tmp.docx.

    Колонтитулы, заголовки, стили и таблицы шаблона сохраняются как есть;
    заполняются только места с плейсхолдерами данными из приложения:
    {TITLE}, {DATE}, {CITY}, {VENUE}, {REFEREE}, {SECRETARY}, число участников,
    а также синхронизируются протоколы (регламент, призёры, графики).
    """
    if not _HAS_DOCX:
        return jsonify({'error': 'Библиотека Word не установлена (python-docx)'}), 500
    from io import BytesIO

    # Найти шаблон (единый ogl_tmp.docx)
    tmpl = None
    for cand in (os.path.join(PROJECT_DIR, 'ogl_tmp.docx'),
                 os.path.join(ASSETS_DIR, 'ogl_tmp.docx'),
                 os.path.join(_get_data_dir(), 'ogl_tmp.docx')):
        if os.path.exists(cand):
            tmpl = cand
            break
    if not tmpl:
        return jsonify({'error': 'Шаблон отчёта (ogl_tmp.docx) не найден'}), 500

    doc = _build_report_document(tournament_id, tmpl)
    if isinstance(doc, tuple):  # (json, status)
        return doc

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"otchet_glavnogo_sudi_{tournament_id}.docx",
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/api/report/<int:tournament_id>/sostav-par', methods=['GET'])
def report_sostav_par(tournament_id):
    """Сформировать «Состав пар» на основе sp_tmp.docx.

    Заполняет плейсхолдеры {TOURNZMENTNAME}, {TOURNZMENTDATE}, {FIGTHINGYEARS},
    строит таблицу пар всех раундов из жеребьёвки/сетки:
    № п/п | в/к | боец1 (красный угол) | Спорт/организация | Рез.боя | боец2 (синий угол) | Спорт/организация.
    Группировка как в шаблоне: подзаголовок «Юноши {годы}» / «Девушки {годы}» + строки пар.
    """
    if not _HAS_DOCX:
        return jsonify({'error': 'Библиотека Word не установлена (python-docx)'}), 500
    from io import BytesIO
    import re as _re
    import copy as _copy
    from collections import defaultdict as _dd
    from docx.oxml.ns import qn as _qn
    _W_MAIN = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # Найти шаблон sp_tmp.docx (или заготовку sp_tmp_new.docx)
    tmpl = None
    for _dir in (DATA_DIR, ASSETS_DIR, PROJECT_DIR, _get_data_dir()):
        for cand in ('sp_tmp_new.docx', 'sp_tmp.docx'):
            c = os.path.join(_dir, cand)
            if os.path.exists(c):
                tmpl = c
                break
        if tmpl:
            break
    if not tmpl:
        return jsonify({'error': 'Шаблон состава пар (sp_tmp.docx) не найден'}), 500

    # Данные турнира и жеребьёвки (аналогично _build_report_document)
    conn = get_db()
    t = row_to_dict(conn.execute("SELECT * FROM tournaments WHERE id=?", (tournament_id,)).fetchone())
    if t is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404
    lottery = rows_to_list(conn.execute(
        "SELECT * FROM lottery WHERE tournament_id=? ORDER BY weight_class, slot", (tournament_id,)).fetchall())
    results = rows_to_list(conn.execute(
        "SELECT * FROM bracket_results WHERE tournament_id=?", (tournament_id,)).fetchall())
    ath_ids = [r['athlete_id'] for r in conn.execute(
        "SELECT athlete_id FROM lottery WHERE tournament_id=?", (tournament_id,)).fetchall()]
    conn.close()

    athletes = []
    if ath_ids:
        _ph = ','.join('?' * len(ath_ids))
        connb = get_db()
        athletes = rows_to_list(connb.execute(f"SELECT * FROM athletes WHERE id IN ({_ph})", ath_ids).fetchall())
        connb.close()
    by_id = {a['id']: a for a in athletes}

    def _fmt_date(s):
        p = (s or '').split('-')
        return f'{p[2]}.{p[1]}.{p[0]}' if len(p) == 3 else (s or '').replace('-', '.')
    d1 = _fmt_date(t.get('date_from') or '')
    d2 = _fmt_date(t.get('date_to') or t.get('date_from') or '')
    dates = f'{d1} — {d2}' if d2 and d2 != d1 else d1

    # --- Собираем раунды по категориям (пары) ---
    groups = _dd(list); order = []
    for s in lottery:
        if s['weight_class'] not in groups:
            order.append(s['weight_class'])
        groups[s['weight_class']].append(s)
    for wc in groups:
        groups[wc].sort(key=lambda x: x['slot'])
    order.sort(key=lambda wc: float((_re.search(r'(\d+(?:\.\d+)?)', wc or '') or _re.search(r'0', '0')).group(1) or 0))

    # Коллекция пар: (wc, round_no, a1_id, a2_id, winner_id) — красный = a1 (меньший слот), синий = a2
    from collections import OrderedDict
    pairs_by_wc = OrderedDict()  # wc -> список пар
    for wc in order:
        slots = groups[wc]
        if not slots:
            continue
        # пол категории — по первому лоту
        g = by_id.get(slots[0]['athlete_id']).get('gender') or '' if by_id.get(slots[0]['athlete_id']) else ''
        gl = g.lower()
        gender = 'мужской' if 'муж' in gl else ('женский' if 'жен' in gl else '')
        # годы участников
        ys = []
        for s in slots:
            a = by_id.get(s['athlete_id'])
            if a and (a.get('birth_year') or ''):
                try: ys.append(int(str(a['birth_year']).strip()))
                except: pass
        wc_key = (wc, gender, ys)
        _rounds, champion, silver, bronzes = _build_weight_rounds(slots, results)
        plist = []
        rno = 1
        for rd in _rounds:
            for (a1, a2, win, method) in rd:
                # пропускаем пары без двух реальных бойцов (BYE — одиночные слоты не идут в состав)
                if a1 is None or a2 is None:
                    continue
                plist.append((wc, gender, ys, rno, a1, a2, win))
            rno += 1
        pairs_by_wc[wc] = plist

    # Группировка по полу (Юноши/Девушки) и годам
    def gender_word(g):
        return 'Юноши' if g == 'мужской' else ('Девушки' if g == 'женский' else '')
    def year_label(ys):
        if not ys: return ''
        m, M = min(ys), max(ys)
        return str(m) if m == M else f'{m}-{M}'

    doc = Document(tmpl)

    # --- Заполняем заголовок (header1/header2/header3) ---
    def _replace_in_part(part_el, mapping):
        for p in part_el.iter(_qn('w:p')):
            runs = p.findall(_qn('w:r'))
            if not runs: continue
            full = ''.join(rt.text or '' for r in runs for rt in r.findall(_qn('w:t')))
            changed = False
            for k, v in mapping.items():
                if k in full:
                    full = full.replace(k, v)
                    changed = True
            if changed:
                t0s = runs[0].findall(_qn('w:t'))
                if t0s:
                    t0s[0].text = full
                    for r in runs[1:]:
                        for rt in r.findall(_qn('w:t')):
                            rt.text = ''
                # если название турнира не в первом run (разбито), поставим в любом случае в первый w:t
                for r in runs:
                    for rt in r.findall(_qn('w:t')):
                        if rt.text and '{' in rt.text:
                            rt.text = full
                            break

    _year_all = []
    for s in lottery:
        a = by_id.get(s['athlete_id'])
        if a and (a.get('birth_year') or ''):
            try: _year_all.append(int(str(a['birth_year']).strip()))
            except: pass
    _years_label = year_label(_year_all) if _year_all else ''
    for sec_i in range(len(doc.sections)):
        sec = doc.sections[sec_i]
        for part_name in ('header', 'footer'):
            try:
                part = getattr(sec, part_name)
                hdr_el = part._element
                _replace_in_part(hdr_el, {
                    '{TOURNZMENTNAME}': t.get('name') or '',
                    '{TOURNZMENTDATE}': dates,
                    '{FIGTHINGYEARS}': _years_label,
                })
            except Exception:
                pass

    # --- Строим таблицу ---
    t0 = doc.tables[0]
    trs = t0._tbl.findall(_qn('w:tr'))
    # Прототипы: подзаголовок (строка 1) и данных (строка 2)
    proto_sub = None; proto_data = None
    if len(trs) >= 1:
        proto_sub = _copy.deepcopy(trs[1]) if len(trs) > 1 and len(trs[1].findall(_qn('w:tc'))) == 1 else None
    if len(trs) > 2:
        proto_data = _copy.deepcopy(trs[2])
    # Сбрасываем таблицу до заголовка
    for tr in trs[1:]:
        t0._tbl.remove(tr)

    def set_cell_text(tc, text):
        ps = tc.findall(_qn('w:p'))
        if not ps: return
        p = ps[0]
        runs = p.findall(_qn('w:r'))
        if runs:
            ts = runs[0].findall(_qn('w:t'))
            if ts:
                ts[0].text = text
                for r in runs[1:]:
                    for rt in r.findall(_qn('w:t')):
                        rt.text = ''
        else:
            from docx.oxml import parse_xml
            _safe = (text or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
            p.append(parse_xml('<w:r xmlns:w="%s"><w:t>%s</w:t></w:r>' % (_W_MAIN, _safe)))
        # если текст содержит {}, ставим во второй run? — не нужно, ставили в 0

    def replace_markers_in_cells(tr_el):
        # заменить {FIGTHINGYEARS} и прочие в подзаголовке
        for tc in tr_el.findall(_qn('w:tc')):
            for p in tc.iter(_qn('w:p')):
                runs = p.findall(_qn('w:r'))
                if not runs: continue
                full = ''.join(rt.text or '' for r in runs for rt in r.findall(_qn('w:t')))
                if '{' in full:
                    full = full.replace('{FIGTHINGYEARS}', _years_label)
                    ts = runs[0].findall(_qn('w:t'))
                    if ts:
                        ts[0].text = full
                        for r in runs[1:]:
                            for rt in r.findall(_qn('w:t')):
                                rt.text = ''

    # Собираем группы (пол, годы) -> [wcs], сортируем
    from collections import OrderedDict
    group_map = OrderedDict()
    for wc, plist in pairs_by_wc.items():
        if not plist: continue
        gender = plist[0][1]
        ys = plist[0][2]
        gkey = (gender_word(gender), year_label(set(ys)))
        group_map.setdefault(gkey, []).append(wc)
    # сортировка групп: Юноши сначала, затем Девушки; внутри — по весу
    def gsort(item):
        (gw, yl), wcs = item
        rank = 0 if gw == 'Юноши' else (1 if gw == 'Девушки' else 2)
        return (rank, yl)
    group_sorted = sorted(group_map.items(), key=gsort)

    seq = 0
    for (gword, ylabel), wcs in group_sorted:
        # подзаголовок
        if proto_sub is not None:
            sub = _copy.deepcopy(proto_sub)
            replace_markers_in_cells(sub)
            # подставим годы/пол
            for tc in sub.findall(_qn('w:tc')):
                for p in tc.iter(_qn('w:p')):
                    runs = p.findall(_qn('w:r'))
                    if not runs: continue
                    full = ''.join(rt.text or '' for r in runs for rt in r.findall(_qn('w:t')))
                    if '{' in full or 'Юноши' in full or 'Девуш' in full:
                        full = (gword + ' ' + ylabel)
                        ts = runs[0].findall(_qn('w:t'))
                        if ts:
                            ts[0].text = full
                            for r in runs[1:]:
                                for rt in r.findall(_qn('w:t')):
                                    rt.text = ''
            t0._tbl.append(sub)
        # пары
        for wc in wcs:
            for (_, gender, ys, rno, a1, a2, win) in pairs_by_wc[wc]:
                seq += 1
                # вес категории
                mm = _re.search(r'([\d.,+-]+)', wc or '')
                weight = mm.group(1) if mm else wc
                # боец1 = a1 (меньший слот), боец2 = a2
                def name_of(aid):
                    a = by_id.get(aid)
                    return a['name'] if a else ('#' + str(aid))
                def club_of(aid):
                    a = by_id.get(aid)
                    return ((a.get('region') or '').replace('Калмыкия','').strip()) if a else ''
                # результат
                res = ''
                if win is not None:
                    res = 'в'
                    # попробуем метод
                    for r in results:
                        if r.get('weight_class') == wc and r.get('winner_id') == win:
                            mth = (r.get('method') or '')
                            if mth: res = mth if mth != 'авто' else 'в'
                            break
                row = _copy.deepcopy(proto_data) if proto_data is not None else None
                if row is None:
                    continue
                cells = row.findall(_qn('w:tc'))
                # колонки: 0 №, 1 в/к, 2 боец1, 3 спорт1, 4 рез, 5 боец2, 6 спорт2
                if len(cells) >= 7:
                    set_cell_text(cells[0], str(seq))
                    set_cell_text(cells[1], weight)
                    set_cell_text(cells[2], name_of(a1))
                    set_cell_text(cells[3], club_of(a1))
                    set_cell_text(cells[4], res)
                    set_cell_text(cells[5], name_of(a2))
                    set_cell_text(cells[6], club_of(a2))
                t0._tbl.append(row)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"sostav_par_{tournament_id}.docx",
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/api/weighin/export', methods=['POST'])
def export_weighin_word():
    """Сохранить таблицу взвешивания в Word (.docx) с текущим списком из UI."""
    if not _HAS_DOCX:
        return jsonify({'error': 'Библиотека Word не установлена (python-docx)'}), 500
    data = request.get_json(silent=True) or {}
    rows = data.get('rows') or []
    if not rows:
        return jsonify({'error': 'Нет данных для сохранения'}), 400

    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from io import BytesIO

    doc = Document()
    title = doc.add_heading('Таблица взвешивания', level=1)
    title.alignment = 1  # центр

    headers = ['№', 'Ф.И.О.', 'Год рожд.', 'Пол', 'Регион', 'Вес (кг)']
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr = table.rows[0].cells
    for j, h in enumerate(headers):
        hdr[j].text = h
        for par in hdr[j].paragraphs[0].runs:
            par.font.bold = True

    for r in rows:
        cells = table.add_row().cells
        cells[0].text = str(r.get('n', ''))
        cells[1].text = str(r.get('name', ''))
        cells[2].text = str(r.get('birthYear', ''))
        cells[3].text = str(r.get('gender', ''))
        cells[4].text = str(r.get('region', '-'))
        cells[5].text = str(r.get('weight', ''))

    # Ширина колонок
    widths = [Cm(1.2), Cm(6.0), Cm(2.2), Cm(2.2), Cm(4.0), Cm(2.5)]
    for row in table.rows:
        for j, w in enumerate(widths):
            if j < len(row.cells):
                row.cells[j].width = w

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"vzveshivanie_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/api/weighin/reglament', methods=['GET'])
def export_weighin_reglament():
    """Сформировать отдельный регламент соревнований (reglament_tmp.docx).

    Заполняет шапку ({TOURNIRNAME}, {FIGTHERYO}, даты), группирует весовые
    категории по полу (мужской → «Юноши», женский → «Девушки») и годам
    рождения участников (диапазон от большего к меньшему, e.g. «Юноши 2016-2015
    г.р.»), выводит по строке на категорию: вес | кол-во участников | бои по
    дням | кол-во боёв. Юноши — в Табл1, Девушки — в Табл2.
    """
    if not _HAS_DOCX:
        return jsonify({'error': 'Библиотека Word не установлена (python-docx)'}), 500
    from io import BytesIO
    import re as _re
    import copy as _copy
    from collections import defaultdict as _dd
    from docx.oxml.ns import qn as _qn
    _W_MAIN = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    tournament_id = request.args.get('tournament_id', type=int)
    if not tournament_id:
        return jsonify({'error': 'Не передан tournament_id'}), 400

    tmpl = None
    for _dir in (DATA_DIR, ASSETS_DIR, PROJECT_DIR, _get_data_dir()):
        c = os.path.join(_dir, 'reglament_tmp.docx')
        if os.path.exists(c):
            tmpl = c
            break
    if not tmpl:
        return jsonify({'error': 'Шаблон регламента (reglament_tmp.docx) не найден'}), 500

    conn = get_db()
    t = row_to_dict(conn.execute("SELECT * FROM tournaments WHERE id=?", (tournament_id,)).fetchone())
    if t is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404
    lottery = rows_to_list(conn.execute(
        "SELECT * FROM lottery WHERE tournament_id=? ORDER BY weight_class, slot", (tournament_id,)).fetchall())
    results = rows_to_list(conn.execute(
        "SELECT * FROM bracket_results WHERE tournament_id=?", (tournament_id,)).fetchall())
    ath_ids = [r['athlete_id'] for r in lottery]
    conn.close()

    athletes = []
    if ath_ids:
        _ph = ','.join('?' * len(ath_ids))
        connb = get_db()
        athletes = rows_to_list(connb.execute(f"SELECT * FROM athletes WHERE id IN ({_ph})", ath_ids).fetchall())
        connb.close()
    by_id = {a['id']: a for a in athletes}

    def _fmt_date(s):
        p = (s or '').split('-')
        return f'{p[2]}.{p[1]}.{p[0]}' if len(p) == 3 else (s or '').replace('-', '.')
    d1 = _fmt_date(t.get('date_from') or '')
    d2 = _fmt_date(t.get('date_to') or t.get('date_from') or '')
    dates = f'{d1} — {d2}' if d2 and d2 != d1 else d1
    # число дней: max(дни, интервал дат)
    try:
        _days_from_dates = (datetime.strptime(str(t['date_to']), '%Y-%m-%d') -
                            datetime.strptime(str(t['date_from']), '%Y-%m-%d')).days + 1
        _days_from_dates = max(1, _days_from_dates)
    except Exception:
        _days_from_dates = 1
    total_days = max(1, int(t.get('days') or 1), _days_from_dates)

    def _gender_athlete(a):
        g = (a.get('gender') or '').strip().lower() if a else ''
        if 'муж' in g or g in ('м', 'm', 'male'):
            return 'мужской'
        if 'жен' in g or g in ('ж', 'f', 'female'):
            return 'женский'
        return ''
    def _gender_word(g):
        return 'Юноши' if g == 'мужской' else ('Девушки' if g == 'женский' else (g or ''))

    # --- Собираем категории (порядок по возрастанию веса) ---
    groups = _dd(list)
    for s in lottery:
        groups[s['weight_class']].append(s)
    for wc in groups:
        groups[wc].sort(key=lambda x: x['slot'])
    order = sorted(groups.keys(), key=lambda wc: (
        float((_re.search(r'(\d+(?:\.\d+)?)', wc or '') or _re.search(r'(\d)', '0')).group(1) or 0)))

    # пол и годы категории (по лотам)
    def _slot_gender(s):
        return _gender_athlete(by_id.get(s['athlete_id']))
    def _slot_year(s):
        a = by_id.get(s['athlete_id'])
        try:
            return int(str((a or {}).get('birth_year') or '').strip())
        except (TypeError, ValueError):
            return None

    _cat_gender = {}
    _cat_years = {}
    for wc in order:
        _slots = groups[wc]
        _g = _slot_gender(_slots[0]) if _slots else ''
        _cat_gender[wc] = _g
        _ys = [_slot_year(s) for s in _slots]
        _ys = [y for y in _ys if y is not None]
        _cat_years[wc] = _ys

    # --- Раскидываем категории: Юноши -> Табл1, Девушки -> Табл2 ---
    by_gender = {'Юноши': [], 'Девушки': [], '': []}
    for wc in order:
        gw = _gender_word(_cat_gender.get(wc) or '')
        by_gender.setdefault(gw, []).append(wc)

    # --- Помощь: распределение боёв по дням (аналог roundToDay) ---
    def _round_to_day(round_num, max_rounds, days):
        days = max(1, int(days) or 1)
        if not max_rounds or max_rounds <= 0:
            return 1
        if days >= max_rounds:
            return round_num
        return min(days, 1 + int((round_num - 1) * days / max_rounds))

    def _fights_by_day(wc):
        """Возвращает список [кол-во реальных боёв в 1-й день, ..., total_days].

        Строит раунды как в сетке (app.js): 1-й раунд — пары из жеребьёвки,
        нечётный участник — авто-проход (НЕ бой); последующие раунды — пары
        победителей. Учитываются только реальные бои (оба соперника), которые
        распределяются по дням через roundToDay.
        """
        results_map = {}
        for r in (results or []):
            results_map[(r['weight_class'], r['pair_index'])] = r
        slots = sorted(groups[wc], key=lambda x: x['slot'])

        # 1-й раунд
        rounds = []
        fr = []
        pidx = 0
        for i in range(0, len(slots) - 1, 2):
            a1, a2 = slots[i], slots[i + 1]
            r = results_map.get((wc, pidx))
            fr.append((a1['athlete_id'], a2['athlete_id'],
                       (r['winner_id'] if r else None), (r['method'] if r else '')))
            pidx += 1
        # нечётный хвост — авто-проход (не бой): пара-заглушка (a1, None)
        if len(slots) % 2 != 0:
            fr.append((slots[-1]['athlete_id'], None, None, 'авто'))
        rounds.append(fr)

        # последующие раунды
        round_num = 1
        while True:
            prev = rounds[-1]
            winners = [f[2] for f in prev]
            if any(w is None for w in winners) or len(winners) <= 1:
                break
            round_num += 1
            offset = (round_num - 1) * 1000
            nxt = []
            for i in range(0, len(winners) - 1, 2):
                a1, a2 = winners[i], winners[i + 1]
                r = results_map.get((wc, offset + i))
                nxt.append((a1, a2, (r['winner_id'] if r else None), (r['method'] if r else '')))
            if not nxt:
                break
            # если последний элемент нечётный в паре — не бывает (доп. обработаем ниже)
            rounds.append(nxt)

        max_rounds = len(rounds)
        per_day = [0] * total_days
        for round_num, fights in enumerate(rounds, start=1):
            day = _round_to_day(round_num, max_rounds, total_days)
            # только реальные бои (у обоих участников есть id)
            real = [f for f in fights if f[0] is not None and f[1] is not None]
            per_day[day - 1] += len(real)
        return per_day

    def _find_wc_by_weight(weight):
        """Находит исходную весовую категорию по цифре веса (для расчёта боёв)."""
        if weight is None:
            return None
        for wc in order:
            mm = _re.search(r'(\d+(?:\.\d+)?)', wc or '')
            if mm and mm.group(1) == str(weight):
                return wc
        return None

    def _set_para_text(para_el, mapping):
        """Заменяет маркеры в абзаце, сохраняя стиль первого run."""
        runs = para_el.findall(_qn('w:r'))
        if not runs:
            return
        wts = runs[0].findall(_qn('w:t'))
        _full = ''.join(_t.text or '' for _r in runs for _t in _r.findall(_qn('w:t')))
        _changed = False
        for _k, _v in mapping.items():
            if _k in _full:
                _full = _full.replace(_k, _v)
                _changed = True
        if _changed and wts:
            wts[0].text = _full
            for _r in runs[1:]:
                for _t in _r.findall(_qn('w:t')):
                    _t.text = ''

    def _set_cell_text(cell_el, text):
        """Записывает текст в ячейку (в первый w:t первого run)."""
        for p in cell_el.findall(f'{{{_W_MAIN}}}p'):
            runs = p.findall(f'{{{_W_MAIN}}}r')
            if runs:
                wts = runs[0].findall(f'{{{_W_MAIN}}}t')
                if wts:
                    wts[0].text = str(text)
                    for _t in wts[1:]:
                        _t.text = ''
                    for _r in runs[1:]:
                        for _t in _r.findall(f'{{{_W_MAIN}}}t'):
                            _t.text = ''
                    return
        # нет run — добавим
        p = cell_el.findall(f'{{{_W_MAIN}}}p')[0]
        from docx.oxml import OxmlElement
        r = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.text = str(text)
        r.append(t)
        p.append(r)

    def _clone_row(tr_el):
        return _copy.deepcopy(tr_el)

    def _head_word(para_el, mapping):
        for p in para_el.iter(_qn('w:p')):
            _set_para_text(p, mapping)

    doc = Document(tmpl)
    body = doc.element.body

    # --- Заполняем шапку (Табл0) ---
    for tbl in doc.tables[:1]:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _set_para_text(p._p, {'{TOURNIRNAME}': t.get('name') or '',
                                          '{FIGTHERYO}': dates})

    # --- Находим Табл1, Табл2 и параграф-группу ---
    _tabs = list(body.iter(_qn('w:tbl')))
    tbl1 = _tabs[1] if len(_tabs) > 1 else None
    tbl2 = _tabs[2] if len(_tabs) > 2 else None

    # маркеры {VK}/{KU} в шапках таблиц данных заменяем на подписи
    for _tbl in (tbl1, tbl2):
        if _tbl is not None:
            _head_word(_tbl, {'{VK}': 'Весовая категория', '{KU}': 'Количество участников'})

    # параграф-группа «Юноши {YO2-YO1}» (первый p с этим маркером)
    group_p = None
    for p in body.iter(_qn('w:p')):
        txt = ''.join(_t.text or '' for _t in p.iter(_qn('w:t')))
        if '{YO2-YO1}' in txt:
            group_p = p
            break

    # --- Создание отдельной таблицы данных из шаблона Табл1 ---
    def _new_data_table():
        """Клонирует Табл1, оставляет только шапку, возвращает объект таблицы."""
        _np = _copy.deepcopy(tbl1)
        _trs = _np.findall(_qn('w:tr'))
        for _tr in _trs[1:]:          # удаляем строку-образец данных, оставляем шапку
            _np.remove(_tr)
        return _np

    def _fill_row(tbl_el, weight, slots, wc):
        # берём строку-образец из шапки таблицы (клонируем последнюю строку исходной Табл1)
        _sample = tbl_sample_row
        row = _clone_row(_sample)
        _slots = sorted(slots, key=lambda x: x['slot'])
        cnt = len(_slots)
        # бои по дням — из всей категории веса (wc)
        _fights = _fights_by_day(wc)
        total_fights = sum(_fights)
        cells = row.findall(_qn('w:tc'))
        if len(cells) >= 6:
            _set_cell_text(cells[0], weight)
            _set_cell_text(cells[1], str(cnt))
            for di in range(min(total_days, 3)):
                if di < len(cells) - 2:
                    _set_cell_text(cells[2 + di], str(_fights[di]) if di < len(_fights) else '')
            _set_cell_text(cells[5], str(total_fights))
        tbl_el.append(row)

    # строка-образец данных (из исходной Табл1, последняя строка-данных)
    tbl_sample_row = None
    if tbl1 is not None:
        _trs = tbl1.findall(_qn('w:tr'))
        if len(_trs) >= 2:
            tbl_sample_row = _clone_row(_trs[-1])

    # --- Группировка по (пол, пара лет). Важно: участники распределяются по парам
    # по ИХ годём рождения, а внутри пары группируются по весу. Один и тот же вес
    # может встретиться в разных парах (если в нём бойцы разных лет), поэтому вес
    # повторяется, но количество — только участники с годом в этой паре.
    def _pair_groups(wcs):
        """Возвращает список (label_years, {weight: [slots]}).

        Годы всех участников этих категорий разбиваются на пары от большего
        к меньшему (2016-2015, 2014-2013, ...). Каждый слот относится к той паре,
        в которую попадает год рождения его спортсмена. Одиночный год — своя пара.
        """
        if not wcs:
            return []
        # собрать все годы (по слотам, а не по категории)
        all_years = set()
        for wc in wcs:
            for s in groups[wc]:
                _y = _slot_year(s)
                if _y is not None:
                    all_years.add(_y)
        if not all_years:
            # нет годов — одна «пустая» группа со всеми
            agg = {}
            for wc in wcs:
                agg.setdefault(wc, []).extend(groups[wc])
            return [('', agg)]
        years_sorted = sorted(all_years, reverse=True)
        pairs = []
        i = 0
        while i < len(years_sorted):
            if i + 1 < len(years_sorted):
                pairs.append((years_sorted[i], years_sorted[i + 1]))
                i += 2
            else:
                pairs.append((years_sorted[i], None))
                i += 1
        # под каждую пару собираем слоты с годом в диапазоне пары
        bucket_by_wc = {i: {} for i in range(len(pairs))}   # i -> {weight: [slots]}
        for wc in wcs:
            _w = _re.search(r'(\d+(?:\.\d+)?)', wc or '')
            _weight = _w.group(1) if _w else wc
            for s in groups[wc]:
                _y = _slot_year(s)
                if _y is None:
                    continue
                # ищем пару, в которую входит _y
                for idx, (yhi, ylo) in enumerate(pairs):
                    low = ylo if ylo is not None else yhi
                    if low <= _y <= yhi:
                        bucket_by_wc[idx].setdefault(_weight, []).append(s)
                        break
        result = []
        for idx, (yhi, ylo) in enumerate(pairs):
            label = str(yhi) if ylo is None else f'{yhi}-{ylo}'
            if bucket_by_wc[idx]:
                result.append((label, bucket_by_wc[idx]))
        return result

    # --- Собираем готовый перечень групп: (gender, label, [wc]) в нужном порядке ---
    _gender_order = [('Юноши', by_gender.get('Юноши', [])),
                     ('Девушки', by_gender.get('Девушки', [])),
                     ('', by_gender.get('', []))]
    group_blocks = []   # (gender, label, [wc])
    for gender, wcs in _gender_order:
        for (label, members) in _pair_groups(wcs):
            group_blocks.append((gender, label, members))

    # --- Удаляем исходные Табл1, Табл2 и параграф-группу (они станут прототипами) ---
    # Не удаляем Tабл0 (шапку) и заголовок «РЕГЛАМЕНТ СОРЕВНОВАНИЙ» (P1).
    if tbl1 is not None and tbl1.getparent() is not None:
        tbl1.getparent().remove(tbl1)
    if tbl2 is not None and tbl2.getparent() is not None:
        tbl2.getparent().remove(tbl2)
    if group_p is not None and group_p.getparent() is not None:
        group_p.getparent().remove(group_p)

    # --- Вставляем группы: параграф-подзаголовок + отдельная таблица данных ---
    # Вставку ведём ПОСЛЕ параграфа «РЕГЛАМЕНТ СОРЕВНОВАНИЙ» (или в конец body).
    # Используем курсор: каждый новый элемент вставляется ПОСЛЕ предыдущего.
    anchor_before = None
    for _p in body.iter(_qn('w:p')):
        if 'РЕГЛАМЕНТ СОРЕВНОВАНИЙ' in ''.join(_t.text or '' for _t in _p.iter(_qn('w:t'))):
            anchor_before = _p
            break

    def _make_group_para(gender, label):
        _np = _copy.deepcopy(group_p) if group_p is not None else None
        if _np is None:
            return None
        _label_txt = label if label else ''
        _gender_txt = gender + ' ' if gender else ''
        _set_para_text(_np, {'{YO2-YO1}': _label_txt, 'Юноши ': _gender_txt})
        return _np

    if anchor_before is None:
        # fallback — в конец тела (перед sectPr)
        _append_here = None
        _body_children = list(body)
        for _c in _body_children:
            if _c.tag == _qn('w:sectPr'):
                _append_here = _c
                break
        for gender, label, weight_map in group_blocks:
            np_el = _make_group_para(gender, label)
            nt_el = _new_data_table()
            # weight_map: {вес: [слоты этой пары]} — сортируем по весу
            for weight, slots in sorted(weight_map.items(), key=lambda kv: float(kv[0] or 0)):
                _wc = _find_wc_by_weight(weight)
                _fill_row(nt_el, weight, slots, _wc)
            if _append_here is not None:
                _append_here.addprevious(np_el) if np_el is not None else None
                if np_el is not None:
                    np_el.addnext(nt_el)
            else:
                body.append(np_el) if np_el is not None else None
                body.append(nt_el)
    else:
        cursor = anchor_before
        for gender, label, weight_map in group_blocks:
            np_el = _make_group_para(gender, label)
            nt_el = _new_data_table()
            for weight, slots in sorted(weight_map.items(), key=lambda kv: float(kv[0] or 0)):
                _wc = _find_wc_by_weight(weight)
                _fill_row(nt_el, weight, slots, _wc)
            if np_el is not None:
                cursor.addnext(np_el)   # параграф после курсора
                cursor = np_el
            cursor.addnext(nt_el)       # таблица после параграфа/курсора
            cursor = nt_el

    # --- Общий итог: строка в самом конце (после всех групп) ---
    # Считаем по всем весовым категориям: всего участников, бои по каждому дню,
    # общее число боёв. Вставляем итоговую строку в последнюю созданную таблицу данных.
    if group_blocks:
        _tot_ath = 0
        _tot_by_day = [0] * total_days
        for wc in order:
            _slots = groups[wc]
            _tot_ath += len(_slots)
            _fd = _fights_by_day(wc)
            for _di in range(len(_fd)):
                _tot_by_day[_di] += _fd[_di]
        _tot_fights = sum(_tot_by_day)
        # последняя таблица в body (это последняя вставленная таблица данных)
        _last_tbl_el = None
        for _t in reversed(list(body.iter(_qn('w:tbl')))):
            _last_tbl_el = _t
            break
        if _last_tbl_el is not None:
            _row = _clone_row(tbl_sample_row) if tbl_sample_row is not None else None
            if _row is not None:
                _cells = _row.findall(_qn('w:tc'))
                if len(_cells) >= 6:
                    _set_cell_text(_cells[0], 'ИТОГО')
                    _set_cell_text(_cells[1], str(_tot_ath))
                    for _di in range(min(total_days, 3)):
                        if _di < len(_cells) - 2:
                            _set_cell_text(_cells[2 + _di], str(_tot_by_day[_di]) if _di < len(_tot_by_day) else '')
                    _set_cell_text(_cells[5], str(_tot_fights))
                _last_tbl_el.append(_row)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"reglament_{tournament_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


# ======================== СТАТИЧЕСКИЕ ФАЙЛЫ ========================

@app.route('/')
def index():
    """Главная страница."""
    return send_from_directory(ASSETS_DIR, 'index.html')

@app.route('/<path:path>')
def static_files(path):
    """Отдача статических файлов (JS, CSS Рё С‚.Рґ.)."""
    if path.endswith('.js'):
        return send_from_directory(ASSETS_DIR, path, mimetype='application/javascript')
    return send_from_directory(ASSETS_DIR, path)


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


def run_server(host='0.0.0.0', open_browser_on_start=True):
    """Инициализирует БД, готовит каталог данных и запускает Flask.

    Возвращает порт, на котором поднят сервер. Используется и в консольном
    варианте, и в GUI-оболочке (собранный .exe).
    """
    import database
    database.DB_PATH = Path(_db_path())
    init_db()
    backup_db(force=True)
    port = find_free_port()

    def _serve():
        app.run(host=host, port=port, debug=False, threaded=True, use_reloader=False)

    t = threading.Thread(target=_serve, daemon=True)
    t.start()
    if open_browser_on_start:
        threading.Thread(target=open_browser, args=(port,), daemon=True).start()
    return port


if __name__ == '__main__':
    port = run_server()
    print('=' * 60)
    print('  >> СОРЕВНОВАНИЯ ПО БОКСУ (SQLite + Flask)')
    print('=' * 60)
    print(f'\n  Сервер запущен: http://localhost:{port}')
    print('  Нажмите Ctrl+C для остановки\n')
    # ждём завершения (поток сервера — daemon)
    try:
        while True:
            time.sleep(3600)
    except KeyboardInterrupt:
        pass
