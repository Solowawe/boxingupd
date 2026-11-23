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
import xml.etree.ElementTree as ET

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
    # При обновлении версии — перезаписываем шаблоны, чтобы старые/повреждённые
    # копии не оставались после апдейта.

    def _has_token(path, token):
        """Проверяет, есть ли в .docx-файле искомый текстовый маркер (в теле/колонтитулах)."""
        try:
            import zipfile as _zipfile
            with _zipfile.ZipFile(path) as _zf:
                for _n in _zf.namelist():
                    if _n.endswith('.xml'):
                        _c = _zf.read(_n).decode('utf-8', 'ignore')
                        if token in _c:
                            return True
            return False
        except Exception:
            return False

    _TEMPLATE_FILES = ('pkdrk_template.docx', 'ppp2_template.docx', 'ogl_tmp.docx',
                       'sp_tmp.docx', 'sp_tmp_new.docx', 'grso_tmp.docx', 'reglament_tmp.docx')
    for _f in _TEMPLATE_FILES:
        _src = os.path.join(_bundled, 'reports', _f)
        _dst_dir = os.path.join(DATA_DIR, 'reports')
        os.makedirs(_dst_dir, exist_ok=True)
        _dst = os.path.join(_dst_dir, _f)
        if not os.path.exists(_src):
            continue
        _need_tpl = False
        if not os.path.exists(_dst):
            _need_tpl = True
        else:
            _dm = _read_json_version(os.path.join(DATA_DIR, 'update-manifest.json'))
            _need_tpl = _ver_tuple(_bm) > _ver_tuple(_dm)
            # sp_tmp — официальные шаблоны «Состава пар». Если в сохранённой копии
            # нет маркера {TOURNAMENT_NAME} (признак старого шаблона с зашитым текстом),
            # принудительно обновляем её из сборки, чтобы убрать устаревшую шапку.
            if _f in ('sp_tmp.docx', 'sp_tmp_new.docx') and not _need_tpl:
                if not _has_token(_dst, '{TOURNAMENT_NAME}') or not _has_token(_dst, '{DATE1-DATE2_YEAR}'):
                    _need_tpl = True
        if _need_tpl:
            try:
                shutil.copy2(_src, _dst)
            except Exception:
                pass
    ASSETS_DIR = DATA_DIR
else:
    ASSETS_DIR = PROJECT_DIR


# ======================== ШАБЛОНЫ ОТЧЁТОВ (папка reports) ========================
# Все шаблоны отчётов (.docx) лежат в подпапке «reports» каталога данных / проекта.
REPORT_DIR = os.path.join(DATA_DIR, 'reports')


def _report_template(name):
    """Вернуть полный путь к шаблону отчёта <name> в папке reports.

    Ищем в каталоге данных (портативный режим), каталоге ресурсов и каталоге
    проекта — чтобы поддержать и замороженный exe, и запуск из исходников.
    """
    import os as _os
    for _d in (DATA_DIR, ASSETS_DIR, PROJECT_DIR, _get_data_dir()):
        _c = _os.path.join(_d, 'reports', name)
        if _os.path.exists(_c):
            return _c
    return None


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
    'update-manifest.json', 'ogl_tmp.docx', 'grso_tmp.docx', 'sp_tmp.docx', 'sp_tmp_new.docx',
    'reglament_tmp.docx', 'boxingapp.exe',
    'updater.bat', 'restart.bat'
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
    # Место проведения и судьи (как при создании турнира)
    if 'venue' in data:
        fields.append("venue = ?")
        params.append(str(data['venue']).strip())
    if 'chief_referee' in data:
        fields.append("chief_referee = ?")
        params.append(str(data['chief_referee']).strip())
    if 'chief_secretary' in data:
        fields.append("chief_secretary = ?")
        params.append(str(data['chief_secretary']).strip())

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


def _simulate_calc_day(tournament_id, known=None):
    """Строит схватки всех раундов по жеребьёвке и возвращает список пар с их днём.

    known — dict {(weight_class, pair_index): winner_id} уже назначенных победителей.
    Каждая пара: (weight_class, pair_index, a1_id, a2_id, day, winner_when_done).
    День пары = max(дней двух соперников), день участника берётся из lottery.day.
    Используется тем же алгоритмом, что и построение сетки на клиенте.
    Возвращает pairs.
    """
    from collections import defaultdict
    known = known or {}
    conn = get_db()
    lottery = rows_to_list(conn.execute(
        "SELECT * FROM lottery WHERE tournament_id=? ORDER BY weight_class, slot", (tournament_id,)).fetchall())
    _trow = conn.execute("SELECT date_from, date_to, days FROM tournaments WHERE id=?", (tournament_id,)).fetchone()
    conn.close()

    # Турнирные дни (как на клиенте): max(дней из интервала дат, поле days)
    _tdays = 1
    try:
        if _trow and _trow['date_from'] and _trow['date_to']:
            _from = _trow['date_from'][:10]
            _to = _trow['date_to'][:10]
            _diff = (__import__('datetime').datetime.strptime(_to, '%Y-%m-%d')
                     - __import__('datetime').datetime.strptime(_from, '%Y-%m-%d')).days
            if _diff >= 0:
                _tdays = _diff + 1
    except Exception:
        pass
    try:
        _tdays = max(_tdays, max(1, int(_trow['days'] or 1)))
    except Exception:
        pass

    # День схватки — по номеру раунда (как «ёлка» на клиенте roundToDay),
    # чтобы финалы попадали на правильные дни турнира, а не на день спортсмена.
    def _round_to_day(round_num, max_rounds, total_days):
        days = max(1, int(total_days or 1))
        if not max_rounds or max_rounds <= 0:
            return 1
        if days >= max_rounds:
            return round_num
        return min(days, 1 + (round_num - 1) * days // max_rounds)

    # Группируем по категориям, сортируем по slot
    groups = defaultdict(list)
    order = []
    for s in lottery:
        if s['weight_class'] not in groups:
            order.append(s['weight_class'])
        groups[s['weight_class']].append(s)
    for wc in groups:
        groups[wc].sort(key=lambda x: x['slot'])

    day_by_ath = {}
    for s in lottery:
        try:
            day_by_ath[s['athlete_id']] = int(s.get('day') or 1)
        except Exception:
            day_by_ath[s['athlete_id']] = 1

    def _day(aid):
        return day_by_ath.get(aid, 1)

    def _known_winner(wc, p):
        return known.get((wc, p))

    # Строим дерево каждой категории ПО КЛИЕНТСКОЙ СХЕМЕ (_build_tree_rounds),
    # чтобы pair_index совпадали с тем, что ожидает рендер результатов/сетка.
    # Из клиентских объектов слота берём только athlete_id (как числа).
    pairs = []
    for wc in order:
        slots = groups[wc]
        rounds = _build_tree_rounds(slots, wc, _fid_act, known)
        _max_rounds = len(rounds)
        for rnd_idx, rnd in enumerate(rounds, start=1):
            _rd_day = _round_to_day(rnd_idx, _max_rounds, _tdays)
            for f in rnd:
                if f['a1'] is None:
                    # участники пары ещё не определены (победитель прошлой пары неизвестен)
                    continue
                if f['a2'] is None:
                    pairs.append((wc, f['p'], f['a1'], None, _rd_day, True))
                else:
                    pairs.append((wc, f['p'], f['a1'], f['a2'], _rd_day, False))
    return pairs


@app.route('/api/tournaments/<int:tournament_id>/simulate-day/<int:day>', methods=['POST'])
def simulate_day(tournament_id, day):
    """Автоматически «провести» день N демо-турнира: сгенерировать схватки дня и
    назначить победителей пары (детерминированно — побеждает боец с меньшим id)."""
    if not data_dir_is_demo(tournament_id):
        return jsonify({'error': 'Доступно только для демо-турнира'}), 400
    conn = get_db()
    # Уже сохранённые результаты (для каскада победителей из прошлых дней)
    saved = {}
    for r in rows_to_list(conn.execute(
            "SELECT weight_class, pair_index, winner_id FROM bracket_results WHERE tournament_id=?",
            (tournament_id,)).fetchall()):
        saved[(r['weight_class'], r['pair_index'])] = r['winner_id']
    conn.close()
    conn = get_db()

    written = []
    # Каскад: после записи победителей дня дерево может «открыть» следующие
    # раунды этого же дня (их победители зависят от только что определённых).
    # Повторяем, пока день даёт новые решённые пары; saved внутри увеличивается.
    while True:
        pairs = _simulate_calc_day(tournament_id, known=saved)
        newly_resolved = []
        for (wc, p, a1, a2, d, auto) in pairs:
            if p is None:
                # Пара без слота-индекса (бай/авто-выход) не сохраняется в bracket_results
                continue
            if d != day:
                continue
            if (wc, p) in saved:
                continue
            if a2 is None:
                # BYE — авто-победитель (a1 проходит дальше)
                new_winner = a1
                conn.execute("""
                    INSERT INTO bracket_results (tournament_id, weight_class, pair_index, winner_id, method, day)
                    VALUES (?,?,?,?, 'авто', ?)
                    ON CONFLICT(tournament_id, weight_class, pair_index)
                    DO UPDATE SET winner_id=excluded.winner_id, method='авто', day=excluded.day
                """, (tournament_id, wc, p, new_winner, day))
            else:
                winner = a1 if a1 < a2 else a2
                conn.execute("""
                    INSERT INTO bracket_results (tournament_id, weight_class, pair_index, winner_id, method, day)
                    VALUES (?,?,?,?, 'единогласное', ?)
                    ON CONFLICT(tournament_id, weight_class, pair_index)
                    DO UPDATE SET winner_id=excluded.winner_id, method='единогласное', day=excluded.day
                """, (tournament_id, wc, p, winner, day))
                new_winner = winner
            saved[(wc, p)] = new_winner
            written.append(f"{wc}|{p}")
            newly_resolved.append((wc, p))
        if not newly_resolved:
            break
    conn.commit()
    conn.close()
    return jsonify({'ok': True, 'day': day, 'written': len(written), 'pairs': written})


@app.route('/api/tournaments/<int:tournament_id>/reset-day/<int:day>', methods=['POST'])
def reset_day(tournament_id, day):
    """Отменить «проведение» дня N демо-турнира: удаляет результаты пар этого дня,
    включая авто-победителей, оставляя остальные дни нетронутыми."""
    if not data_dir_is_demo(tournament_id):
        return jsonify({'error': 'Доступно только для демо-турнира'}), 400
    conn = get_db()
    conn.execute(
        "DELETE FROM bracket_results WHERE tournament_id=? AND day=?",
        (tournament_id, day))
    conn.commit()
    conn.close()
    return jsonify({'ok': True, 'day': day, 'reset': True})


def data_dir_is_demo(tournament_id):
    """Признак «демо»-турнира для включения кнопок авто-проведения."""
    conn = get_db()
    t = conn.execute("SELECT id, name FROM tournaments WHERE id=?", (tournament_id,)).fetchone()
    conn.close()
    if not t:
        return False
    return ('демо' in (t['name'] or '').lower()) or ('Демо' in (t['name'] or ''))


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


# ======================== ЭКСПОРТ / ИМПОРТ ТУРНИРА (XML) ========================

def _q(tag, val):
    """Вспомогательная функция для создания элемента с числовым/текстовым значением."""
    el = ET.Element(tag)
    if isinstance(val, bool):
        val = '1' if val else '0'
    el.text = '' if val is None else str(val)
    return el


def _clean(s):
    """Экранирование необязательно — ET сам экранирует текст и атрибуты."""
    return '' if s is None else str(s)


@app.route('/api/tournaments/<int:tournament_id>/export', methods=['GET'])
def export_tournament_xml(tournament_id):
    """Экспортировать соревнование (реквизиты + спортсмены + жеребьёвка + результаты пар) в XML.

    Формат файла (корневой элемент <boxing_tournament>):
      <tournament>  — реквизиты соревнования
      <athletes>    — полные профили спортсменов-участников
      <lottery>     — жеребьёвка (слоты по весовым категориям)
      <results>     — сохранённые победители пар сетки
    Спортсмены внутри <athletes> несут локальный id (myid), на который ссылаются
    <lottery> и <results> — чтобы импорт сам сгенерировал новые id в базе получателя.
    """
    conn = get_db()
    tournament = conn.execute(
        "SELECT id, name, date_from, date_to, weight_step, days, completed, "
        "chief_referee, chief_secretary, venue FROM tournaments WHERE id = ?",
        (tournament_id,)).fetchone()
    if tournament is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404

    athlete_ids = [r['athlete_id'] for r in conn.execute(
        "SELECT athlete_id FROM tournament_athletes WHERE tournament_id = ?",
        (tournament_id,)).fetchall()]

    # Профили спортсменов
    athletes = []
    if athlete_ids:
        ph = ','.join('?' * len(athlete_ids))
        athletes = rows_to_list(conn.execute(
            "SELECT * FROM athletes WHERE id IN (" + ph + ")", athlete_ids).fetchall())

    lottery = rows_to_list(conn.execute(
        "SELECT athlete_id, slot, weight_class, day FROM lottery WHERE tournament_id = ? "
        "ORDER BY weight_class, slot", (tournament_id,)).fetchall())

    # Результаты пар с именем победителя (для читабельности) — но ссылаемся по myid
    results = rows_to_list(conn.execute(
        "SELECT weight_class, pair_index, winner_id, method, day FROM bracket_results "
        "WHERE tournament_id = ? ORDER BY weight_class, pair_index", (tournament_id,)).fetchall())
    conn.close()

    root = ET.Element('boxing_tournament')
    root.set('version', '1.0')

    t = ET.SubElement(root, 'tournament')
    for key in ('name', 'date_from', 'date_to', 'weight_step', 'days', 'completed',
                'chief_referee', 'chief_secretary', 'venue'):
        t.append(_q(key, tournament[key]))

    ath = ET.SubElement(root, 'athletes')
    aid_to_myid = {}
    for a in athletes:
        ae = ET.SubElement(ath, 'athlete')
        ae.set('myid', str(a['id']))
        for key in ('name', 'birth_year', 'weight', 'gender', 'rank', 'region', 'coach'):
            ae.append(_q(key, a[key]))
        aid_to_myid[a['id']] = str(a['id'])

    lot = ET.SubElement(root, 'lottery')
    for r in lottery:
        le = ET.SubElement(lot, 'entry')
        le.set('athlete_id', aid_to_myid.get(r['athlete_id'], str(r['athlete_id'])))
        le.append(_q('slot', r['slot']))
        le.append(_q('weight_class', r['weight_class']))
        le.append(_q('day', r['day']))

    res = ET.SubElement(root, 'results')
    for r in results:
        re = ET.SubElement(res, 'result')
        re.set('weight_class', str(r['weight_class']))
        re.set('pair_index', str(r['pair_index']))
        re.set('winner_id', aid_to_myid.get(r['winner_id'], str(r['winner_id'])))
        re.append(_q('method', r['method']))
        re.append(_q('day', r['day']))

    # Формируем XML-строку с декларацией UTF-8
    body = ET.tostring(root, encoding='utf-8', xml_declaration=True)

    safe_name = _clean(tournament['name']).replace('"', '').replace('/', '_').replace('\\', '_')
    fname = f"tournament_{safe_name or tournament_id}_export.xml"

    import urllib.parse
    from flask import Response
    resp = Response(body, mimetype='application/xml')
    resp.headers['Content-Disposition'] = "attachment; filename*=UTF-8''" + urllib.parse.quote(fname)
    return resp


@app.route('/api/tournaments/import', methods=['POST'])
def import_tournament_xml():
    """Импортировать соревнование из XML (файл или сырой текст в поле 'xml').

    Безопасно для основной базы:
      * всегда создаётся НОВОЕ соревнование (существующие не меняются);
      * спортсмены создаются заново (при совпадении ФИО+год+пол — переиспользуется
        уже существующий спортсмен из текущей базы, чтобы не плодить дубликаты);
      * жеребьёвка и результаты пар переносятся со ссылкой на новых спортсменов.
    """
    data = None
    xml_text = None

    # Вариант: multipart/form-data (file)
    if request.files:
        f = request.files.get('file')
        if f:
            try:
                xml_text = f.read().decode('utf-8')
            except Exception:
                return jsonify({'error': 'Не удалось прочитать файл (ожидалась кодировка UTF-8)'}), 400

    # Вариант: application/json {"xml": "..."}
    if xml_text is None:
        raw = request.get_data(as_text=True)
        try:
            data = json.loads(raw)
        except Exception:
            data = None
        if data and data.get('xml'):
            xml_text = data['xml']

    if not xml_text:
        return jsonify({'error': 'Отсутствуют данные XML (передайте файл или поле "xml")'}), 400

    # Парсинг
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as e:
        return jsonify({'error': 'Ошибка парсинга XML: ' + str(e)}), 400

    tour_el = root.find('tournament')
    if tour_el is None:
        return jsonify({'error': 'Некорректный XML: отсутствует элемент <tournament>'}), 400

    def _el(el, key):
        node = el.find(key)
        if node is None or node.text is None:
            return ''
        return node.text

    name = (_el(tour_el, 'name') or '').strip()
    date_from = (_el(tour_el, 'date_from') or '').strip()
    if not name or not date_from:
        return jsonify({'error': 'Поля name и date_from обязательны'}), 400

    conn = get_db()

    # 1) Создаём новое соревнование
    cur = conn.execute(
        "INSERT INTO tournaments (name, date_from, date_to, weight_step, days, completed, "
        "chief_referee, chief_secretary, venue) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (
            name,
            date_from,
            (_el(tour_el, 'date_to') or date_from).strip() or date_from,
            _safe_float(_el(tour_el, 'weight_step'), 5.0),
            _safe_int(_el(tour_el, 'days'), 1),
            _safe_int(_el(tour_el, 'completed'), 0),
            (_el(tour_el, 'chief_referee') or '').strip(),
            (_el(tour_el, 'chief_secretary') or '').strip(),
            (_el(tour_el, 'venue') or '').strip(),
        )
    )
    new_tid = cur.lastrowid

    # 2) Импортируем спортсменов: myid (из файла) -> новый id в базе
    #    Дедупликация по (name, birth_year, gender) внутри JVV импорта и против существующих.
    new_athete_id_by_myid = {}
    created_ids = []

    # Загружаем существующих спортсменов для дедупликации
    existing = {}
    for r in conn.execute("SELECT * FROM athletes").fetchall():
        key = (str(r['name']), int(r['birth_year']), str(r['gender']))
        existing.setdefault(key, r['id'])

    # Множество созданных за этот импорт, чтобы один и тот же спортсмен не дублировался
    dup_within = {}

    for ae in root.findall('.//athletes/athlete'):
        myid = ae.get('myid')
        aname = (_el(ae, 'name') or '').strip()
        abirth = _safe_int(_el(ae, 'birth_year'), 0)
        agender = (_el(ae, 'gender') or 'мужской').strip()
        aweight = _safe_float(_el(ae, 'weight'), 0.0)

        if not aname or not abirth:
            continue  # пропускаем некорректные записи

        key = (aname, abirth, agender)
        new_id = None
        if key in dup_within:
            new_id = dup_within[key]
        elif key in existing:
            new_id = existing[key]
        else:
            c2 = conn.execute(
                "INSERT INTO athletes (name, birth_year, weight, gender, rank, region, coach) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                (
                    aname, abirth, aweight, agender,
                    (_el(ae, 'rank') or '').strip(),
                    (_el(ae, 'region') or '').strip(),
                    (_el(ae, 'coach') or '').strip(),
                )
            )
            new_id = c2.lastrowid
            existing[key] = new_id
            dup_within[key] = new_id
            created_ids.append(new_id)

        if myid is not None:
            new_athete_id_by_myid[myid] = new_id

    # 3) Привязываем спортсменов к турниру
    for myid, real_id in new_athete_id_by_myid.items():
        conn.execute(
            "INSERT OR IGNORE INTO tournament_athletes (tournament_id, athlete_id) VALUES (?, ?)",
            (new_tid, real_id)
        )

    # 4) Жеребьёвка
    for le in root.findall('.//lottery/entry'):
        src = le.get('athlete_id')
        real_id = new_athete_id_by_myid.get(src)
        if real_id is None:
            continue
        conn.execute(
            "INSERT INTO lottery (tournament_id, athlete_id, slot, weight_class, day) "
            "VALUES (?, ?, ?, ?, ?)",
            (
                new_tid,
                real_id,
                _safe_int(_el(le, 'slot'), 0),
                (_el(le, 'weight_class') or '').strip(),
                _safe_int(_el(le, 'day'), 1),
            )
        )

    # 5) Результаты пар сетки
    for re in root.findall('.//results/result'):
        src = re.get('winner_id')
        real_winner = new_athete_id_by_myid.get(src)
        if real_winner is None:
            continue
        wc = re.get('weight_class', '')
        pi = _safe_int(re.get('pair_index'), -1)
        if not wc or pi < 0:
            continue
        conn.execute(
            "INSERT OR REPLACE INTO bracket_results "
            "(tournament_id, weight_class, pair_index, winner_id, method, day) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (
                new_tid, wc, pi, real_winner,
                (_el(re, 'method') or '').strip(),
                _safe_int(_el(re, 'day'), 1),
            )
        )

    conn.commit()

    # Собираем статистику
    result = {
        'success': True,
        'tournament_id': new_tid,
        'name': name,
        'athletes_created': len(created_ids),
        'athletes_linked': len(new_athete_id_by_myid),
        'lottery_entries': len(root.findall('.//lottery/entry')),
        'results_entries': len(root.findall('.//results/result')),
    }
    conn.close()
    return jsonify(result), 201


def _safe_int(v, default=0):
    try:
        return int(v)
    except Exception:
        return default


def _safe_float(v, default=0.0):
    try:
        return float(v)
    except Exception:
        return default


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
        # utf-8-sig пропускает BOM в начале файла, если он есть
        manifest = json.loads(raw.decode('utf-8-sig'))
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
        # utf-8-sig пропускает BOM в начале файла, если он есть
        manifest = json.loads(_fetch_url(f'{base}/{UPDATE_MANIFEST}', timeout=20).decode('utf-8-sig'))
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

        # Обновление самого исполняемого файла (exe) — только СКАЧИВАЕМ вручную:
        # новый exe с версией в имени сохраняется в папку update/ рядом с приложением,
        # НО НЕ применяется автоматически — пользователь ставит его вручную.
        if name == 'boxingapp.exe':
            need_restart = False
            try:
                content = _fetch_url(f'{base}/{rel}', timeout=60)
                exe_dir = os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, 'frozen', False) else str(app_dir)
                upd_dir = os.path.join(exe_dir, 'update')
                os.makedirs(upd_dir, exist_ok=True)
                _ver = str(manifest.get('version', ''))
                new_exe = os.path.join(upd_dir, f'boxingapp_{_ver}.exe' if _ver else 'boxingapp_new.exe')
                with open(new_exe, 'wb') as f:
                    f.write(content)
                updated.append(name)
                # НЕ запускаем скрипт-обновлятель и НЕ запрашиваем перезапуск — exe вручную
                continue
            except Exception as e:
                return jsonify({'error': f'Не удалось скачать обновление exe: {e}'}), 500

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


@app.route('/api/update/launch-updater', methods=['POST'])
def update_launch_updater():
    """Запускает скрипт-обновлятель (apply_update.bat) в отдельном процессе.

    Скрипт ждёт завершения основного процесса, перезаписывает boxingapp.exe
    новым файлом из временной папки и перезапускает приложение.
    """
    if not (UPDATE_OWNER and UPDATE_REPO):
        return jsonify({'error': 'Обновления не настроены.'}), 400
    try:
        data = request.get_json(silent=True) or {}
        script = data.get('script') or ''
        # Безопасность: разрешаем только путь внутри папки __exe_update__ рядом с exe/app_dir
        exe_dir = os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, 'frozen', False) else str(Path(ASSETS_DIR))
        allowed_root = os.path.join(exe_dir, '__exe_update__')
        abs_script = os.path.abspath(script)
        if not abs_script.lower().startswith(allowed_root.lower()) or not os.path.exists(abs_script):
            return jsonify({'error': 'Недопустимый путь скрипта обновления.'}), 400
        # Запускаем через cmd /c старт так, чтобы процесс жил после завершения сервера
        import subprocess
        subprocess.Popen(
            ['cmd', '/c', 'start', '', abs_script],
            creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
            close_fds=True,
            cwd=exe_dir
        )
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': f'Не удалось запустить обновлятель: {e}'}), 500



# ======================== ЛОКАЛЬНОЕ ОБНОВЛЕНИЕ ИЗ update-package.zip ========================

@app.route('/api/update/local', methods=['POST'])
def update_local():
    """Применить локальный пакет обновления (update-package.zip), выбранный пользователем.

    Метод: multipart/form-data, поле file = .zip.
    Внутри архива должен лежать update-manifest.json с полями version и files.
    Файлы распаковываются в каталог данных/ассетов (портативная версия -> рядом с exe),
    список разрешённых файлов — UPDATE_ALLOWED. Если в пакете есть boxingapp.exe —
    новый exe с версией в имени сохраняется в папку update/ рядом с приложением и
    НЕ применяется автоматически: пользователь ставит его вручную.
    """
    if 'file' not in request.files:
        return jsonify({'error': 'Не передан файл обновления'}), 400
    f = request.files['file']
    if not f or not f.filename:
        return jsonify({'error': 'Файл пуст'}), 400

    try:
        file_bytes = f.read()
    except Exception as e:
        return jsonify({'error': f'Не удалось прочитать файл: {e}'}), 400

    import zipfile
    from io import BytesIO

    # 1) Читаем манифест из архива (без извлечения содержимого сразу, чтобы проверить версию)
    manifest = None
    try:
        with zipfile.ZipFile(BytesIO(file_bytes)) as z:
            names = z.namelist()
            # Нормализуем имена: ищем update-manifest.json в корне
            mn = next((n for n in names if n.replace('\\', '/').endswith('update-manifest.json')), None)
            if not mn:
                return jsonify({'error': 'В архиве нет update-manifest.json'}), 400
            manifest = json.loads(z.read(mn).decode('utf-8-sig'))
    except Exception as e:
        return jsonify({'error': f'Не удалось прочитать манифест из архива: {e}'}), 400

    if not isinstance(manifest, dict) or not manifest.get('version'):
        return jsonify({'error': 'Некорректный манифест: отсутствует version'}), 400

    version = str(manifest.get('version', ''))
    requested = manifest.get('files', []) or []

    app_dir = Path(ASSETS_DIR)
    updated = []
    skipped = []
    need_restart = False
    pending_exe = None

    # 2) Распаковка файлов по манифесту
    try:
        with zipfile.ZipFile(BytesIO(file_bytes)) as z:
            # Находим соответствие имён в манифесте и записей в архиве (по имени файла в конце пути)
            for rel in requested:
                name = os.path.basename(str(rel).replace('\\', '/'))
                if name not in UPDATE_ALLOWED:
                    skipped.append(name)
                    continue
                if name in ('boxing.db',) or name == LOCAL_VERSION_FILE:
                    skipped.append(name)
                    continue

                # Ищем запись в архиве по имени файла
                entry = None
                for n in z.namelist():
                    if os.path.basename(n.replace('\\', '/')) == name:
                        entry = n
                        break
                if entry is None:
                    skipped.append(name)
                    continue

                content = z.read(entry)

                # Обновление exe — скачиваем в папку update/ рядом с приложением,
                # НО не применяем автоматически: пользователь ставит новый exe вручную.
                if name == 'boxingapp.exe':
                    exe_dir = os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, 'frozen', False) else str(app_dir)
                    upd_dir = os.path.join(exe_dir, 'update')
                    os.makedirs(upd_dir, exist_ok=True)
                    _ver = str(manifest.get('version', ''))
                    new_exe = os.path.join(upd_dir, f'boxingapp_{_ver}.exe' if _ver else 'boxingapp_new.exe')
                    with open(new_exe, 'wb') as fw:
                        fw.write(content)
                    updated.append(name)
                    # НЕ ставим need_restart — exe применяется вручную пользователем
                    continue

                # Обычные файлы — пишем напрямую в каталог данных/ассетов
                target = app_dir / name
                with open(target, 'wb') as fw:
                    fw.write(content)
                updated.append(name)
    except Exception as e:
        return jsonify({'error': f'Ошибка применения локального обновления: {e}'}), 500

    _write_local_version(version)

    return jsonify({
        'ok': True,
        'version': version,
        'updated': updated,
        'skipped': skipped,
        'need_restart': need_restart,
        'restart_script': pending_exe
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
    tmpl = _report_template('ogl_tmp.docx')
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


def _build_tree_rounds(slots, wc, fid, results_map):
    """Строит полное дерево раундов категории ТОЧНО по клиентскому алгоритму
    (app.js `_buildTreeRounds`), чтобы pair_index совпадали с тем, что ожидает
    рендер результатов и сетка.

    slots — список слотов (объекты с athlete_id, slot), отсортированные по slot.
    fid — функция ID из слота/числа.
    results_map — dict {(weight_class, pair_index): winner_id} уже решённых пар
        (для каскада победителей в последующих раундах).

    Возвращает список rounds: каждый round — список диктов {a1,a2,p,r},
    где a1/a2 — ID (или None для бая), p — pair_index (или None), r — winner_id.
    Соответствует клиентской схеме индексов: раунд 1 = 0.., раунд 2 = 1000..,
    раунд 3 = 2000.. и т.д. Баи не имеют p.
    Прерванные раунды (a1/a2 ещё не определены) опускаются.
    """
    N = len(slots)
    is_pow2 = (N & (N - 1)) == 0
    first = []
    pi = 0
    # --- Первый раунд ---
    if N % 2 == 0 and not is_pow2:
        # Чётное не-степень-двойки: часть бойцов в 1-м раунде, остальные — баи.
        rc = 0
        x = N
        while x > 1:
            x = (x + 1) // 2
            rc += 1
        half_pow = 2 ** (rc - 1)
        play_in = 2 * (N - half_pow)
        i = 0
        while i < play_in - 1:
            a1 = fid(slots[i]); a2 = fid(slots[i + 1])
            first.append({'a1': a1, 'a2': a2, 'p': pi, 'r': results_map.get((wc, pi))})
            pi += 1
            i += 2
        for i in range(play_in, N):
            aid = fid(slots[i])
            first.append({'a1': aid, 'a2': None, 'p': None, 'r': aid})
    else:
        for i in range(0, N - 1, 2):
            a1 = fid(slots[i]); a2 = fid(slots[i + 1])
            first.append({'a1': a1, 'a2': a2, 'p': pi, 'r': results_map.get((wc, pi))})
            pi += 1
        if N % 2 != 0:
            aid = fid(slots[N - 1])
            first.append({'a1': aid, 'a2': None, 'p': None, 'r': aid})
    rounds = [first]
    bye_slot = slots[N - 1] if N % 2 != 0 else None
    total_rounds = max(1, int(round(__import__('math').log2(N))))

    def winner_of(f):
        if not f or f.get('r') in (None, 0):
            return None
        w = f['r']
        if f.get('a1') is not None and f['a1'] == w:
            return w
        if f.get('a2') is not None and f['a2'] == w:
            return w
        return None

    if bye_slot is not None:
        # Нечётное число бойцов: схема X-Б7 (как на клиенте).
        bye_id = fid(bye_slot)
        day2 = []
        w_first = winner_of(first[0])
        w_first_no_bye = (w_first if (w_first is not None and w_first != bye_id) else None)
        day2.append({'a1': w_first_no_bye, 'a2': bye_id, 'p': 1000, 'r': results_map.get((wc, 1000))})
        rest = []
        for i in range(1, len(first)):
            w = winner_of(first[i])
            if w is not None and w != bye_id:
                rest.append(w)
        si = 0
        pidx = 1001
        while si + 1 < len(rest):
            day2.append({'a1': rest[si], 'a2': rest[si + 1], 'p': pidx, 'r': results_map.get((wc, pidx))})
            pidx += 1
            si += 2
        if len(rest) % 2 != 0:
            lone = rest[-1]
            x_obj = winner_of(day2[0])
            day2.append({'a1': x_obj, 'a2': lone, 'p': pidx, 'r': results_map.get((wc, pidx))})
        rounds.append(day2)
        rn = 3
        while True:
            prev = rounds[-1]
            cnt = len(prev)
            if rn == 3 and len(rest) % 2 != 0:
                cnt -= 1
            if cnt <= 1:
                break
            off = (rn - 1) * 1000
            nxt = []
            start_idx = 0
            if rn == 3 and len(rest) % 2 != 0:
                start_idx = 1
            i = start_idx
            while i + 1 < len(prev):
                pp = off + ((i - start_idx) // 2)
                nxt.append({'a1': winner_of(prev[i]), 'a2': winner_of(prev[i + 1]),
                            'p': pp, 'r': results_map.get((wc, pp))})
                i += 2
            rounds.append(nxt)
            rn += 1
        return rounds

    # Чётное число бойцов: обычные сохраняющие раунды.
    # Добавляем раунды, пока в последнем раунде больше одной пары — так
    # любой N (в т.ч. не степень двойки, например 10) сходится к финалу.
    rn = 2
    while True:
        prev = rounds[-1]
        if len(prev) <= 1:
            break
        off = (rn - 1) * 1000
        nxt = []
        i = 0
        while i + 1 < len(prev):
            pp = off + (i // 2)
            nxt.append({'a1': winner_of(prev[i]), 'a2': winner_of(prev[i + 1]),
                        'p': pp, 'r': results_map.get((wc, pp))})
            i += 2
        # если в предпоследнем раунде была нечётная пара бая — то в nxt уже
        # может остаться «хвост»-бай (без пары). Добавляем его как пару с None.
        if len(prev) % 2 != 0:
            last = winner_of(prev[-1])
            if last is not None:
                nxt.append({'a1': last, 'a2': None, 'p': None, 'r': last})
        rounds.append(nxt)
        rn += 1
    return rounds


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
    tmpl = _report_template('ogl_tmp.docx')
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
    tmpl = _report_template('sp_tmp_new.docx') or _report_template('sp_tmp.docx')
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

    # --- Системная дата/время формирования документа ---
    import datetime as _dt
    _now = _dt.datetime.now()
    form_date = _now.strftime('%d.%m.%Y')
    form_time = _now.strftime('%H:%M')

    # Даты проведения прописью: "28 октября – 31 октября 2026"
    _rmonths = ['января','февраля','марта','апреля','мая','июня',
                'июля','августа','сентября','октября','ноября','декабря']
    def _date_words(s):
        try:
            yy, mm, dd = (int(x) for x in (s or '').split('-'))
            return f"{dd} {_rmonths[mm-1]} {yy}"
        except Exception:
            return s or ''
    _d1w = _date_words(t.get('date_from') or '')
    _d2w = _date_words(t.get('date_to') or t.get('date_from') or '')
    date_label = f"{_d1w} — {_d2w}" if _d2w and _d2w != _d1w else _d1w

    # Определяем «день турнира», соответствующий СЕГОДНЯШНЕЙ системной дате.
    # День N бойца (lottery.day) → календарная дата = date_from + (N-1) дней.
    _today = _dt.date.today()
    _target_day = None
    try:
        _df = _dt.datetime.strptime((t.get('date_from') or ''), '%Y-%m-%d').date()
        _dt_ = None
        if (t.get('date_to') or ''):
            try: _dt_ = _dt.datetime.strptime(t.get('date_to'), '%Y-%m-%d').date()
            except Exception: _dt_ = None
        # фильтруем только если сегодня в диапазоне дат турнира [date_from, date_to]
        if _df <= _today and (_dt_ is None or _today <= _dt_):
            _off = (_today - _df).days + 1
            if 1 <= _off <= 31:
                _target_day = _off
    except Exception:
        _target_day = None
    # день бойца (из жеребьёвки) по его athlete_id
    _lot_day = {}
    for _s in lottery:
        _lot_day[_s['athlete_id']] = int(_s.get('day') or 1)

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
                # Одиночный боец (BYE) попадает в состав, если в категории больше одного
                # участника (он ждёт соперника/продвигается по сетке). Если спортсмен вообще
                # один в категории — он автоматически победитель, ни с кем не боксирует,
                # поэтому в «Состав пар» его НЕ вносим, чтобы не путать судейскую коллегию.
                if len(slots) <= 1:
                    continue
                if a1 is None and a2 is not None:
                    a1, a2 = a2, None
                # Только реальные пары: если второго бойца нет (a2 is None), это
                # одиночный бай/ожидание соперника — в «Состав пар» не выводим.
                if a2 is None:
                    continue
                # день пары = максимальный день из дней соперников (боец не может выступать раньше своего дня)
                _pday = max(_lot_day.get(a1, 1), _lot_day.get(a2, 1))
                plist.append((wc, gender, ys, rno, a1, a2, win, _pday))
            rno += 1
        pairs_by_wc[wc] = plist

    # День, для которого строим «Состав пар»:
    #   1) явный параметр ?day=N — приоритет;
    #   2) иначе «текущий» день турнира (сегодня в диапазоне дат), если он определён;
    #   3) иначе день не задан — печатаем ВСЕ пары.
    _req_day = None
    try:
        _req_day = int(request.args.get('day') or '') if request.args.get('day') else None
    except Exception:
        _req_day = None
    _selected_day = _req_day if _req_day is not None else _target_day
    if _selected_day is not None:
        _filtered = OrderedDict()
        for _wc, _plist in pairs_by_wc.items():
            _fp = [pp for pp in _plist if pp[7] == _selected_day]
            if _fp:
                _filtered[_wc] = _fp
        pairs_by_wc = _filtered

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
                    '{TOURNAMENT_NAME}': t.get('name') or '',
                    '{DATE1-DATE2_YEAR}': date_label,
                })
                if part_name == 'footer':
                    _replace_in_part(hdr_el, {
                        '{ГлавныйСудьяФИО}': (t.get('chief_referee') or ''),
                        '{ГлавныйСудьяРегион}': '',
                        '{ГлавныйСекретарьФИО}': (t.get('chief_secretary') or ''),
                        '{ГлавныйСекретарьРегион}': '',
                    })
            except Exception:
                pass

    # --- Заполняем дату/время формирования в теле документа ---
    _replace_in_part(doc.element.body, {
        '{ДД.ММ.ГГГГ}': form_date,
        '{время}': form_time,
        '{TOURNAMENT_NAME}': t.get('name') or '',
        '{DATE1-DATE2_YEAR}': date_label,
    })

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
        # Убираем автонумерацию списка (numPr) — иначе Word сам подставляет
        # номера с плавающей запятой (1.1, 1.2 и т.п.) поверх вставляемого текста.
        try:
            _np = p.find(_qn('w:pPr'))
            if _np is not None:
                _num = _np.find(_qn('w:numPr'))
                if _num is not None:
                    _np.remove(_num)
        except Exception:
            pass
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

    # Возрастные блоки — из названия турнира: ищем пары годов «2016-2015» и т.п.,
    # нормализуем каждую пару как (min, max) и сортируем по младшему году.
    def gender_word(g):
        return 'Юноши' if g == 'мужской' else ('Девушки' if g == 'женский' else '')

    # Возрастные блоки — всегда строим по ФАКТИЧЕСКИМ годам участников жеребьёвки
    # (по 2 года, начиная с чётного года от младшего). Название турнира НЕ используем
    # как источник, чтобы блоки покрывали всех реальных спортсменов (например 2010-2016).
    def _build_age_ranges(years):
        yrs = sorted(set(years))
        if not yrs:
            return []
        _lo = yrs[0] - (yrs[0] % 2)
        _hi = max(yrs)
        out = []
        _r = _lo
        while _r <= _hi:
            out.append((_r, _r + 1))
            _r += 2
        return out

    _age_ranges = _build_age_ranges(_year_all)

    def _block_for_year(y):
        # найти блок, в который входит год y (по приоритету ближайший/наименьшая граница)
        if not y:
            return None, None
        for (_lo, _hi) in _age_ranges:
            if _lo <= y <= _hi:
                return _lo, _hi
        # не вошёл никуда — ближайший по разнице к младшей границе
        best = None; bestd = 10**9
        for (_lo, _hi) in _age_ranges:
            d = abs(y - _lo)
            if d < bestd:
                bestd = d; best = (_lo, _hi)
        return best

    # Перестраиваем: каждая ПАРА относится к возрастному блоку по году рождения участника.
    # Группа = (пол, блок_лет). Внутри — весовые категории (от меньшего веса к большему).
    group_map = OrderedDict()
    pair_to_block = {}   # (wc, rno, a1, a2) -> (lo, hi) для согласованной сортировки
    for wc, plist in pairs_by_wc.items():
        for pp in plist:
            _g = pp[1]
            _ys = [by_id.get(pp[4], {}).get('birth_year'), by_id.get(pp[5], {}).get('birth_year')]
            _yint = None
            for _yv in _ys:
                try:
                    if _yv: _yint = int(_yv); break
                except Exception:
                    continue
            _blk = _block_for_year(_yint) if _yint is not None else None
            if _blk is None:
                continue
            _gl = gender_word(_g)
            _gkey = (_gl, _blk)
            group_map.setdefault(_gkey, []).append(wc)
    # убираем дубли категорий и сортируем группы: Юноши/Девушки, возрастные блоки по младшему году
    for _gk in group_map:
        group_map[_gk] = list(dict.fromkeys(group_map[_gk]))
    def _min_year(yl):
        return yl[0] if isinstance(yl, tuple) else 9999
    def gsort(item):
        (gw, blk), wcs = item
        rank = 0 if gw == 'Юноши' else (1 if gw == 'Девушки' else 2)
        return (rank, _min_year(blk))
    group_sorted = sorted(group_map.items(), key=gsort)
    # внутри каждой возрастной группы веса идут от малого к большему (численно)
    for _gi, ((_gw, _blk), _wcs) in enumerate(group_sorted):
        group_sorted[_gi] = ((_gw, _blk), sorted(_wcs, key=lambda w: float(
            (_re.search(r'(\d+(?:\.\d+)?)', w or '') or _re.search(r'0', '0')).group(1) or 0)))

    # для подзаголовка
    def block_label(blk):
        return f'{blk[0]}-{blk[1]}' if blk else ''

    seq = 0
    for (gword, ylabel_in), wcs in group_sorted:
        ylabel = block_label(ylabel_in)
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
                    if 'YO1-YO2' in full or '{' in full or 'Юноши' in full or 'Девуш' in full:
                        # заменяем маркер {YO1-YO2} и подставляем слово пола (Юноши/Девушки),
                        # чтобы заголовок группы соответствовал фактическому полу участников
                        full = full.replace('{YO1-YO2}', ylabel)
                        full = full.replace('Юноши', gword or 'Юноши')
                        full = full.replace('Девушки', gword or 'Девушки')
                        if full.strip() == ylabel or (gword and (full.strip() == gword)):
                            # ячейка без маркера/пола — вставляем "Юноши {годы}"
                            full = (gword + ' ' + ylabel).strip()
                        ts = runs[0].findall(_qn('w:t'))
                        if ts:
                            ts[0].text = full
                            for r in runs[1:]:
                                for rt in r.findall(_qn('w:t')):
                                    rt.text = ''
            t0._tbl.append(sub)
        # пары
        for wc in wcs:
            for (_, gender, ys, rno, a1, a2, win, _pday) in pairs_by_wc[wc]:
                # BYE-боец (одиночный, без соперника) печатается ТОЛЬКО в свой день боя,
                # т.е. когда выбран конкретный день (?day=N или активен «текущий» день).
                # В полном списке (без фильтра по дню) одиночные бойцы не выводятся.
                if (a2 is None) and _selected_day is None:
                    continue
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


@app.route('/api/report/<int:tournament_id>/ppp', methods=['GET'])
def report_ppp(tournament_id):
    """Сформировать «Протокол победителей и призёров» на основе ppp2_template.docx.

    Шаблон (ppp2_template.docx с гербами) содержит: две готовые таблицы с 7 колонками
    (юниорки/юниоры), в каждой — строка-шапка и БЛОК-прототип:
      - строка-вес «Весовая категория N КГ»,
      - строки призёров с маркерами {ФИО}, {ДатаРождения}, {Разряд}, {Регион},
        {СпортОрг}, {Код}.

    Под каждую весовую категорию блок клонируется и заполняется призёрами
    из жеребьёвки и bracket_results. Девушки идут в таблицу «юниорок», юноши —
    в таблицу «юниоров». Возрастные подписи (17-18 ЛЕТ 2009-2008) заменяются
    реальным диапазоном года рождения соответствующего пола.
    """
    if not _HAS_DOCX:
        return jsonify({'error': 'Библиотека Word не установлена (python-docx)'}), 500
    from io import BytesIO
    import re as _re
    import copy as _copy
    from collections import defaultdict as _dd
    from docx.oxml.ns import qn as _qn
    from docx.oxml import parse_xml as _px

    # Найти шаблон ppp2_template.docx
    tmpl = _report_template('ppp2_template.docx')
    if not tmpl:
        return jsonify({'error': 'Шаблон протокола победителей (ppp2_template.docx) не найден'}), 500

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

    # Группируем по весовым категориям
    groups = _dd(list); order = []
    for s in lottery:
        if s['weight_class'] not in groups:
            order.append(s['weight_class'])
        groups[s['weight_class']].append(s)
    for wc in groups:
        groups[wc].sort(key=lambda x: x['slot'])
    order.sort(key=lambda wc: (
        0 if ('жен' in wc) else 1,
        float((_re.search(r'(\d+(?:\.\d+)?)', wc or '') or _re.search(r'0', '0')).group(1) or 0)
    ))

    def _gender_word(g):
        return 'Юноши' if g == 'мужской' else ('Девушки' if g == 'женский' else (g or ''))
    def _range_label(vals):
        if not vals:
            return ''
        _ms = min(vals); _mx = max(vals)
        return str(_ms) if _ms == _mx else f'{_ms}-{_mx}'
    def _short_rank(r):
        r = (r or '').strip()
        m = _re.match(r'^(\d+)', r)
        if not m:
            return r
        d = m.group(1)
        return f'{d} ю.р.' if 'юн' in r.lower() else f'{d} р.'

    # Пол и годы категории
    _cat_gender = {}; _cat_years = {}
    for wc in order:
        g = ''
        if groups[wc]:
            ga = by_id.get(groups[wc][0]['athlete_id'])
            if ga:
                gl = (ga.get('gender') or '').lower()
                g = 'мужской' if 'муж' in gl else ('женский' if 'жен' in gl else '')
        _cat_gender[wc] = g
        ys = []
        for s in groups[wc]:
            a = by_id.get(s['athlete_id'])
            if a and (a.get('birth_year') or ''):
                try: ys.append(int(str(a['birth_year']).strip()))
                except Exception: pass
        _cat_years[wc] = ys

    # Непрерывные возрастные диапазоны по полу
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
    _all_years_by_g = {}
    for wc in order:
        _g = _gender_word(_cat_gender.get(wc) or '')
        for _y in (_cat_years.get(wc) or []):
            _all_years_by_g.setdefault(_g, set()).add(_y)
    _gender_ranges = {g: _continuous_ranges(ys) for g, ys in _all_years_by_g.items()}
    def _cat_in_range(wc, gword):
        _ys = set(_cat_years.get(wc) or [])
        if not _ys:
            return (None, None)
        for (a, b) in _gender_ranges.get(gword, []):
            if all(a <= y <= b for y in _ys):
                return (a, b)
        return (min(_ys), max(_ys))
    _groups_map = {}
    for wc in order:
        _g = _gender_word(_cat_gender.get(wc) or '')
        _r = _cat_in_range(wc, _g)
        _groups_map.setdefault((_g, _r), []).append(wc)
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

    # Призёры категории
    def _medallists(wc, slots):
        _rounds, champ, silver, bronzes = _build_weight_rounds(slots, results)
        put = []
        if champ is not None:
            put.append((1, champ))
        if silver is not None and silver != champ:
            put.append((2, silver))
        bset = [b for b in (bronzes or []) if b is not None and b != champ and b != silver]
        for b in bset[:2]:
            put.append((3, b))
        # НЕ выводим 5-е места (по запросу: только 1, 2, 3-3)
        seen = set(); uniq = []
        for (place, _id) in put:
            if _id not in seen:
                seen.add(_id); uniq.append((place, _id))
        return uniq

    # Годы всех участников конкретного пола (для подписи «… — годы»)
    def _gender_years(gword):
        ys = set()
        for _wc in order:
            if _gender_word(_cat_gender.get(_wc) or '') == gword:
                for _y in (_cat_years.get(_wc) or []):
                    ys.add(_y)
        return sorted(ys)

    doc = __import__('docx').Document(tmpl)
    _WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    def _replace_in_el(el, mapping):
        for p in el.iter(_qn('w:p')):
            runs = p.findall(_qn('w:r'))
            if not runs:
                continue
            full = ''.join(rt.text or '' for r in runs for rt in r.findall(_qn('w:t')))
            changed = False
            for k, v in mapping.items():
                if k in full:
                    full = full.replace(k, str(v)); changed = True
            if changed:
                ts0 = runs[0].findall(_qn('w:t'))
                if ts0:
                    ts0[0].text = full
                    for r in runs[1:]:
                        for rt in r.findall(_qn('w:t')):
                            rt.text = ''
    def set_cell_text(tc, text):
        ps = tc.findall(_qn('w:p'))
        if not ps: return
        p = ps[0]
        try:
            _np = p.find(_qn('w:pPr'))
            if _np is not None:
                _num = _np.find(_qn('w:numPr'))
                if _num is not None:
                    _np.remove(_num)
        except Exception:
            pass
        runs = p.findall(_qn('w:r'))
        if runs:
            ts = runs[0].findall(_qn('w:t'))
            if ts:
                ts[0].text = text
                for r in runs[1:]:
                    for rt in r.findall(_qn('w:t')):
                        rt.text = ''
        else:
            _safe = (text or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
            p.append(_px('<w:r xmlns:w="%s"><w:t>%s</w:t></w:r>' % (_WNS, _safe)))

    # ===== Заполнение шапки, подписей и двух готовых таблиц =====
    # Данные контекста
    _ttl = (t.get('name') or '')
    _venue = (t.get('venue') or '')
    _fmt2 = lambda s: (f"{s[8:10]}.{s[5:7]}.{s[0:4]}" if len(s or '') == 10 else (s or ''))
    _d1 = _fmt2(t.get('date_from') or '')
    _d2 = _fmt2(t.get('date_to') or t.get('date_from') or '')
    _dates = f'{_d1} — {_d2}' if _d2 and _d2 != _d1 else _d1

    # Заполняем маркеры по всему телу, а также в колонтитулах (header/footer),
    # т.к. шапка турнира и блок подписей теперь размещены в колонтитулах.
    _ctx_mapping = {
        '{НАЗВАНИЕ_ТУРНИРА}': _ttl,
        '{МестоПроведения}': _venue,
        '{ДатыПроведения}': _dates,
        '{ТехническийДелегатФИО}': (t.get('chief_referee') or ''),
        '{ТехническийДелегатРегион}': '',
        '{ГлавныйСудьяФИО}': (t.get('chief_referee') or ''),
        '{ГлавныйСудьяРегион}': '',
        '{ГлавныйСекретарьФИО}': (t.get('chief_secretary') or ''),
        '{ГлавныйСекретарьРегион}': '',
    }
    _replace_in_el(doc.element.body, _ctx_mapping)
    # колонтитулы: шапка турнира (header) и подписи (footer)
    for _sec in doc.sections:
        try:
            _replace_in_el(_sec.header._element, _ctx_mapping)
        except Exception:
            pass
        try:
            _replace_in_el(_sec.footer._element, _ctx_mapping)
        except Exception:
            pass

    def _fill_table(tbl, gword, wcs):
        """Заполняет одну таблицу призёрами выбранного пола (девушки/юноши)."""
        trs = tbl._tbl.findall(_qn('w:tr'))
        if len(trs) < 7:
            return
        proto_weight = _copy.deepcopy(trs[1])
        proto_rows = [_copy.deepcopy(trs[i]) for i in range(2, 6)]  # 1,2,3,3 место (без 5-го)
        # удаляем строки кроме заголовка (row0)
        for tr in trs[1:]:
            tbl._tbl.remove(tr)
        for wc in wcs:
            _mm = _re.search(r'(\d+(?:\.\d+)?)', wc or '')
            weight = _mm.group(1) if _mm else ''
            med = _medallists(wc, groups[wc])
            # вес-строка (ячейка объединена на всю ширину)
            wrow = _copy.deepcopy(proto_weight)
            wcells = wrow.findall(_qn('w:tc'))
            if wcells:
                set_cell_text(wcells[0], f'Весовая категория {weight} КГ')
            else:
                _replace_in_el(wrow, {'{WEIGHT_CLASS}': f'Весовая категория {weight} КГ'})
            tbl._tbl.append(wrow)
            # строки призёров — до 4 (в шаблоне 1,2,3,3; без 5-го места)
            for idx in range(len(proto_rows)):
                prow = _copy.deepcopy(proto_rows[idx])
                cells = prow.findall(_qn('w:tc'))
                place = ''; name = ''; dob = ''; rank = ''; region = ''; org = ''; code = ''
                if idx < len(med):
                    plaza, aid = med[idx]
                    a = by_id.get(aid)
                    place = f'{plaza} место'
                    name = (a.get('name') or ('#' + str(aid))) if a else ('#' + str(aid))
                    b = (a.get('birth_year') or '') if a else ''
                    dob = str(b)  # только год рождения (данных о дне/месяце нет)
                    rank = _short_rank(a.get('rank') or '') if a else ''
                    region = (a.get('region') or '') if a else ''
                    org = ''  # спорт организация не хранится
                    code = ''  # спорт общ. не хранится
                else:
                    continue
                if len(cells) >= 7:
                    set_cell_text(cells[0], place)
                    set_cell_text(cells[1], name)
                    set_cell_text(cells[2], dob)
                    set_cell_text(cells[3], rank)
                    set_cell_text(cells[4], region)
                    set_cell_text(cells[5], org)
                    set_cell_text(cells[6], code)
                else:
                    _replace_in_el(prow, {
                        '{ФИО}': name, '{ДатаРождения}': dob, '{Разряд}': rank,
                        '{Регион}': region, '{СпортОрг}': org, '{Код}': code,
                    })
                tbl._tbl.append(prow)

    # Девушки / Юноши: категории по полу
    gender_cats = {}
    for (gword, yrange), wcs in _groups_sorted:
        gender_cats.setdefault(gword, []).extend(wcs)

    # Две готовые таблицы: индекс 0 = юниорки(девушки), 1 = юниоры(юноши)
    tables = doc.tables
    if len(tables) >= 1 and gender_cats.get('Девушки'):
        _fill_table(tables[0], 'Девушки', gender_cats['Девушки'])
    if len(tables) >= 2 and gender_cats.get('Юноши'):
        _fill_table(tables[1], 'Юноши', gender_cats['Юноши'])

    # Возрастные подписи «среди {ПОЛ} {в1}-{в2} ЛЕТ»:
    #   {ПОЛ} — Девушки (1-я таблица) / Юноши (2-я таблица);
    #   {в1} — наименьший возраст участника, {в2} — наибольший;
    #   возраст = текущий год − год рождения участника.
    import datetime as _dt
    _cur_year = _dt.datetime.now().year

    # Возрастной диапазон по категории 2017-2010 (общий для обоих полов):
    # все годы рождения участников турнира
    _all_birth = set()
    for _wc in order:
        for _y in (_cat_years.get(_wc) or []):
            _all_birth.add(_y)
    _age_min = None; _age_max = None
    if _all_birth:
        _age_max = _cur_year - min(_all_birth)
        _age_min = _cur_year - max(_all_birth)

    def _fill_para_markers(p, mapping):
        # заменяет маркеры в абзаце (объединяя текст)
        runs = p.findall(_qn('w:r'))
        if not runs:
            return
        full = ''.join(rt.text or '' for r in runs for rt in r.findall(_qn('w:t')))
        changed = False
        for k, v in mapping.items():
            if k in full:
                full = full.replace(k, str(v)); changed = True
        if changed:
            ts0 = runs[0].findall(_qn('w:t'))
            if ts0:
                ts0[0].text = full
                ts0[0].set(_qn('xml:space'), 'preserve')
                for r in runs[1:]:
                    for rt in r.findall(_qn('w:t')):
                        rt.text = ''

    for pol, tab in (('Девушки', tables[0] if len(tables) >= 1 else None),
                     ('Юноши', tables[1] if len(tables) >= 2 else None)):
        if tab is None:
            continue
        # ищем абзац-подпись с {ПОЛ}/{в1}/{в2}, расположенный перед этой таблицей
        # в последовательности body (игнорируя другие таблицы)
        tbl_el = tab._tbl
        body_els = list(doc.element.body)
        ti = body_els.index(tbl_el)
        # ищем ближайший w:p выше по порядку (до предыдущей таблицы), содержащий маркеры
        found = False
        for j in range(ti - 1, -1, -1):
            el = body_els[j]
            if el.tag == _qn('w:tbl'):
                break
            if el.tag == _qn('w:p'):
                full = ''.join((t.text or '') for t in el.iter(_qn('w:t')))
                if '{ПОЛ}' in full or '{в1}' in full or '{в2}' in full:
                    _fill_para_markers(el, {
                        '{ПОЛ}': pol,
                        '{в1}': _age_min if _age_min is not None else '',
                        '{в2}': _age_max if _age_max is not None else '',
                    })
                    found = True
                    break
        # если над таблицей подписи не нашлось — запасной вариант: первая подпись в документе
        if not found:
            for p in doc.element.body.iter(_qn('w:p')):
                full = ''.join(_t.text or '' for _t in p.iter(_qn('w:t')))
                if '{ПОЛ}' in full or '{в1}' in full or '{в2}' in full:
                    _fill_para_markers(p, {
                        '{ПОЛ}': pol,
                        '{в1}': _age_min if _age_min is not None else '',
                        '{в2}': _age_max if _age_max is not None else '',
                    })
                    break

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"protokol_prizery_{tournament_id}.docx",
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/api/report/<int:tournament_id>/participants', methods=['GET'])
def report_participants(tournament_id):
    """Сформировать «Протокол участников» (.docx): таблица участников выбранного соревнования
    с колонками №, Ф.И.О., Вес, Пол, Год рождения, Разряд, Регион, Тренер."""
    if not _HAS_DOCX:
        return jsonify({'error': 'Библиотека Word не установлена (python-docx)'}), 500
    from io import BytesIO
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn as _qn

    conn = get_db()
    t = row_to_dict(conn.execute("SELECT * FROM tournaments WHERE id=?", (tournament_id,)).fetchone())
    if t is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404
    a_ids = [r['athlete_id'] for r in conn.execute(
        "SELECT athlete_id FROM tournament_athletes WHERE tournament_id=?", (tournament_id,)).fetchall()]
    athletes = []
    if a_ids:
        _ph = ','.join('?' * len(a_ids))
        rows = conn.execute(f"SELECT * FROM athletes WHERE id IN ({_ph})", a_ids).fetchall()
        athletes = rows_to_list(rows)
    conn.close()
    by_id = {a['id']: a for a in athletes}
    # Порядок — как в списке участников турнира (athlete_ids)
    ordered = [by_id[i] for i in a_ids if i in by_id]

    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = Cm(1.5); sec.right_margin = Cm(1.5)
    sec.top_margin = Cm(1.5); sec.bottom_margin = Cm(1.5)

    # Шаблон с гербом и подвалом (подписи судей), если есть; иначе программно
    _pt_tmpl = _report_template('participants_tmp.docx')
    if _pt_tmpl:
        doc = Document(_pt_tmpl)
        # удалить абзацы-заголовки из тела шаблона (перезаполним ниже)
        for _p in list(doc.paragraphs):
            _p._p.getparent().remove(_p._p)
    else:
        doc = Document()
        sec = doc.sections[0]
        sec.left_margin = Cm(1.5); sec.right_margin = Cm(1.5)
        sec.top_margin = Cm(1.5); sec.bottom_margin = Cm(1.5)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('ПРОТОКОЛ УЧАСТНИКОВ')
    run.bold = True; run.font.size = Pt(14)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def _fmt_date(s):
        return f"{s[8:10]}.{s[5:7]}.{s[0:4]}" if len(s or '') == 10 else (s or '')
    d1 = _fmt_date(t.get('date_from') or '')
    d2 = _fmt_date(t.get('date_to') or t.get('date_from') or '')
    dates = f'{d1} — {d2}' if d2 and d2 != d1 else d1
    sub_run = sub.add_run(f"{t.get('name') or ''}".strip() + (f"\n{dates}" if dates else ''))
    sub_run.font.size = Pt(12)

    # Заполняем маркеры в верхнем и нижнем колонтитулах шаблона (если он используется)
    if _pt_tmpl:
        for _sec in doc.sections:
            for _pn in ('header', 'footer', 'first_page_header', 'first_page_footer'):
                try:
                    _part_el = getattr(_sec, _pn)._element
                    for _p in _part_el.iter(_qn('w:p')):
                        _full = ''.join(_rt.text or '' for _r in _p.findall(_qn('w:r')) for _rt in _r.findall(_qn('w:t')))
                        _changed = False
                        for _k, _v in {
                            '{НАЗВАНИЕ_ТУРНИРА}': t.get('name') or '',
                            '{ДатыПроведения}': dates,
                            '{ГлавныйСудьяФИО}': (t.get('chief_referee') or ''),
                            '{ГлавныйСудьяРегион}': '',
                            '{ГлавныйСекретарьФИО}': (t.get('chief_secretary') or ''),
                            '{ГлавныйСекретарьРегион}': '',
                        }.items():
                            if _k in _full:
                                _full = _full.replace(_k, str(_v)); _changed = True
                        if _changed:
                            _runs = _p.findall(_qn('w:r'))
                            if _runs:
                                _ts0 = _runs[0].findall(_qn('w:t'))
                                if _ts0:
                                    _ts0[0].text = _full
                                    for _r in _runs[1:]:
                                        for _rt in _r.findall(_qn('w:t')):
                                            _rt.text = ''
                except Exception:
                    pass

    headers = ['№', 'Ф.И.О.', 'Вес (кг)', 'Пол', 'Год рождения', 'Разряд', 'Регион', 'Тренер']
    widths = [Cm(1.0), Cm(4.2), Cm(1.6), Cm(1.9), Cm(2.0), Cm(2.2), Cm(2.6), Cm(2.8)]

    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    try:
        table.style = 'Table Grid'
    except Exception:
        _W2 = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        _tw = table._tbl.find('{%s}tblPr' % _W2) if table._tbl is not None else None
        _bx = ('<w:tblBorders xmlns:w="%s">'
               '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '</w:tblBorders>') % _W2
        from docx.oxml import parse_xml
        if _tw is not None:
            _tw.append(parse_xml(_bx))
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ''
        p = hdr[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9)

    for idx, a in enumerate(ordered, start=1):
        cells = table.add_row().cells
        vals = [
            str(idx), a.get('name') or '', str(a.get('weight') or ''), a.get('gender') or '',
            str(a.get('birth_year') or ''), a.get('rank') or '', a.get('region') or '', a.get('coach') or ''
        ]
        for i, v in enumerate(vals):
            cells[i].text = ''
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 1 else WD_ALIGN_PARAGRAPH.CENTER
            rr = p.add_run(v)
            rr.font.size = Pt(9)

    # фиксируем ширину колонок
    for i, w in enumerate(widths):
        for cell in table.columns[i].cells:
            cell.width = w

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"protokol_uchastnikov_{tournament_id}.docx",
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/api/report/<int:tournament_id>/pkdrk', methods=['GET'])
def report_pkdrk(tournament_id):
    """«Протокол комиссии по допуску» на основе pkdrk_template.docx.

    Таблица «по районам» (Республика Калмыкия) пересобирается динамически:
      * строки районов — по фактическим населённым пунктам участников турнира,
        г. Элиста всегда первой, остальные по алфавиту;
      * колонки весов юношей/девушек — по фактическим весовым категориям
        турнира (по возрастанию), сколько их есть;
      * строка ВСЕГО — итоги по весам и общие.
    """
    if not _HAS_DOCX:
        return jsonify({'error': 'Библиотека Word не установлена (python-docx)'}), 500
    from io import BytesIO
    import re as _re
    from collections import defaultdict as _dd
    from docx.oxml.ns import qn as _qn
    from docx.oxml import parse_xml as _px

    tmpl = _report_template('pkdrk_template.docx')
    if not tmpl:
        return jsonify({'error': 'Шаблон протокола комиссии (pkdrk_template.docx) не найден'}), 500

    conn = get_db()
    t = row_to_dict(conn.execute("SELECT * FROM tournaments WHERE id=?", (tournament_id,)).fetchone())
    if t is None:
        conn.close()
        return jsonify({'error': 'Соревнование не найдено'}), 404
    lottery = rows_to_list(conn.execute(
        "SELECT * FROM lottery WHERE tournament_id=? ORDER BY weight_class, slot", (tournament_id,)).fetchall())
    ath_ids = [r['athlete_id'] for r in conn.execute(
        "SELECT DISTINCT athlete_id FROM lottery WHERE tournament_id=?", (tournament_id,)).fetchall()]
    conn.close()

    athletes = []
    if ath_ids:
        _ph = ','.join('?' * len(ath_ids))
        connb = get_db()
        athletes = rows_to_list(connb.execute(f"SELECT * FROM athletes WHERE id IN ({_ph})", ath_ids).fetchall())
        connb.close()
    by_id = {a['id']: a for a in athletes}

    def _wc_gender(wc):
        wc = wc or ''
        return 'мужской' if wc.startswith('муж') else ('женский' if wc.startswith('жен') else '')
    def _wc_weight(wc):
        m = _re.search(r'(\d+(?:\.\d+)?)', wc or '')
        return float(m.group(1)) if m else 0.0
    def _fmt_w(v):
        v = float(v)
        return f"{int(v)}" if v == int(v) else f"{v:g}"

    # Все весовые категории турнира по полу, возрастая
    male_weights = sorted(set(_wc_weight(s['weight_class']) for s in lottery if _wc_gender(s['weight_class']) == 'мужской'))
    female_weights = sorted(set(_wc_weight(s['weight_class']) for s in lottery if _wc_gender(s['weight_class']) == 'женский'))

    # ---- Районы: нормализация значения region ----
    _ELISTA_CLUBS = {'багатур', 'гром 08', 'рсшор', 'ск боец', 'спк-1',
                     'ойрат-арена', 'ойрат - арена', 'ойрат арена'}
    _REGION_KEYS = {  # канон -> ключ сортировки
        'Элиста': 'элиста', 'Городовиковск': 'городовиковск', 'Лагань': 'лагань',
        'Большой Царын': 'большой царын', 'Приютное': 'приютное', 'Артезиан': 'артезиан',
        'Садовое': 'садовое', 'Комсомольский': 'комсомольский', 'Цаган Аман': 'цаган аман',
        'Ики-Бурул': 'ики-бурул', 'Яшкуль': 'яшкуль',
    }
    def _canon_region(raw):
        raw = (raw or '').strip()
        if not raw:
            return None
        low = raw.lower()
        for pre in ('калмыкия г.', 'калмыкия п.', 'калмыкия с.', 'калмыкия г',
                    'калмыкия', 'республика калмыкия'):
            if low.startswith(pre):
                low = low[len(pre):].strip()
                break
        low = ' '.join(low.split())
        if low in _ELISTA_CLUBS:
            return 'Элиста'
        for canon in _REGION_KEYS:
            if low == canon.lower():
                return canon
        return None

    # grouped[район] = True для участников турнира
    rayon_set = set()
    for s in lottery:
        a = by_id.get(s['athlete_id'])
        if a is None:
            continue
        r = _canon_region(a.get('region'))
        if r:
            rayon_set.add(r)
    def _rk(item):
        return (0 if item[0] == 'Элиста' else 1, item[0])
    rayon_order = sorted(rayon_set, key=lambda r: (0 if r == 'Элиста' else 1, r))

    # ---- Нормализация весовой категории к «мужской 43 кг» (без «, xx кг» в шаблоне юл.) ----
    # В шаблоне таблица «по районам» — первая таблица документа
    doc = __import__('docx').Document(tmpl)
    _WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    # итоговая ширина таблицы (dxa); поля документа лево/право 300 -> текст ~11306
    _total_w = 11282

    def _replace_in_el(el, mapping):
        for p in el.iter(_qn('w:p')):
            runs = p.findall(_qn('w:r'))
            if not runs:
                continue
            full = ''.join(rt.text or '' for r in runs for rt in r.findall(_qn('w:t')))
            changed = False
            for k, v in mapping.items():
                if k in full:
                    full = full.replace(k, str(v)); changed = True
            if changed:
                ts0 = runs[0].findall(_qn('w:t'))
                if ts0:
                    ts0[0].text = full
                    for r in runs[1:]:
                        for rt in r.findall(_qn('w:t')):
                            rt.text = ''
    def set_cell_text(tc, text):
        ps = tc.findall(_qn('w:p'))
        if not ps:
            return
        p = ps[0]
        # Если ячейка — продолжение вертикального слияния (vMerge=cont из шаблона
        # таблицы по возрасту), снимаем слияние, чтобы каждая строка была отдельной.
        _tcp0 = tc.find(_qn('w:tcPr'))
        if _tcp0 is not None:
            _vm = _tcp0.find(_qn('w:vMerge'))
            if _vm is not None:
                _tcp0.remove(_vm)
        try:
            _np = p.find(_qn('w:pPr'))
            if _np is not None:
                _num = _np.find(_qn('w:numPr'))
                if _num is not None:
                    _np.remove(_num)
        except Exception:
            pass
        runs = p.findall(_qn('w:r'))
        if runs:
            ts = runs[0].findall(_qn('w:t'))
            if ts:
                ts[0].text = text
                for r in runs[1:]:
                    for rt in r.findall(_qn('w:t')):
                        rt.text = ''
        else:
            _safe = (text or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
            p.append(_px('<w:r xmlns:w="%s"><w:t>%s</w:t></w:r>' % (_WNS, _safe)))

    # Шапка/подписи
    _fmt2 = lambda s: (f"{s[8:10]}.{s[5:7]}.{s[0:4]}" if len(s or '') == 10 else (s or ''))
    _d1 = _fmt2(t.get('date_from') or '')
    _d2 = _fmt2(t.get('date_to') or t.get('date_from') or '')
    _dates = f'{_d1} — {_d2}' if _d2 and _d2 != _d1 else _d1

    # ---- Заголовок протокола: маркеры в фразе «проводились с … года в … по адресу: …» ----
    _rmonths = ['января','февраля','марта','апреля','мая','июня',
                'июля','августа','сентября','октября','ноября','декабря']
    def _date_words(s):
        try:
            yy, mm, dd = (int(x) for x in (s or '').split('-'))
            return f"{dd} {_rmonths[mm-1]} {yy}"
        except Exception:
            return s or ''
    _d1w = _date_words(t.get('date_from') or '')
    _d2w = _date_words(t.get('date_to') or t.get('date_from') or '')
    _dates_propis = f"{_d1w} по {_d2w}" if _d2w and _d2w != _d1w else _d1w
    # количество допущенных боксёров (все участники лотереи, муж + жен)
    _cnt_all = sum(1 for s in lottery if by_id.get(s['athlete_id']) is not None)
    # количество районов и его запись прописью для «из … районов Республики»
    _n_rayons = len(rayon_order)
    def _num_words_rod(n):   # числительное в родительном падеже: пять->пяти, два->двух и т.д.
        n = int(n)
        _map = {1:'одного', 2:'двух', 3:'трёх', 4:'четырёх', 5:'пяти',
                6:'шести', 7:'семи', 8:'восьми', 9:'девяти', 10:'десяти',
                11:'одиннадцати', 12:'двенадцати', 13:'тринадцати',
                14:'четырнадцати', 15:'пятнадцати', 16:'шестнадцати',
                17:'семнадцати', 18:'восемнадцати', 19:'девятнадцати',
                20:'двадцати', 30:'тридцати', 40:'сорока', 50:'пятидесяти',
                60:'шестидесяти', 70:'семидесяти', 80:'восьмидесяти', 90:'девяноста'}
        if n in _map:
            return _map[n]
        if n < 100:
            return _map[(n//10)*10] + ' ' + _map.get(n % 10, '')
        return str(n)
    def _rayons_words_rod(n):
        n = int(n)
        word = _num_words_rod(n)
        _last2 = n % 100
        _last = n % 10
        if 11 <= _last2 <= 14:
            unit = 'районов'
        elif _last == 1:
            unit = 'района'
        else:
            unit = 'районов'
        return f'{word} {unit}'
    _rayons_words = _rayons_words_rod(_n_rayons)
    # возраста и годы рождения участников
    _b_years = []
    for s in lottery:
        a = by_id.get(s['athlete_id'])
        if a:
            try:
                _y = int(str(a.get('birth_year') or '').strip())
                if _y > 0:
                    _b_years.append(_y)
            except Exception:
                pass
    _by_min = min(_b_years) if _b_years else 0
    _by_max = max(_b_years) if _b_years else 0
    _year_now = datetime.now().year
    # возраст = текущий год - год рождения
    _age_min = (_year_now - _by_max) if _by_max else 0   # младший (наименьший возраст)
    _age_max = (_year_now - _by_min) if _by_min else 0   # старший (наибольший возраст)
    _years_range = f"{_by_max}-{_by_min}" if (_by_min and _by_max) else ''

    _ctx = {
        '{НАЗВАНИЕ_ТУРНИРА}': (t.get('name') or ''),
        '{МестоПроведения}': (t.get('venue') or ''),
        '{ДатыПроведения}': _dates,
        '{ГлавныйСудьяФИО}': (t.get('chief_referee') or ''),
        '{ГлавныйСудьяРегион}': '',
        '{ГлавныйСекретарьФИО}': (t.get('chief_secretary') or ''),
        '{ГлавныйСекретарьРегион}': '',
        '{ТехническийДелегатФИО}': (t.get('chief_referee') or ''),
        '{ТехническийДелегатРегион}': '',
        '{ПОЛ}': '', '{в1}': '', '{в2}': '',
        # маркеры заголовка протокола
        '{Даты проведения}': _dates_propis,
        '{Место проведения}': 'Республике Калмыкия, г. Элиста',
        '{Адрес проведения}': '',
        '{количество}': str(_cnt_all),
        '{количество районов прописью}': _rayons_words,
        '{в1}': str(_age_min) if _by_max else '',
        '{в2}': str(_age_max) if _by_min else '',
        '{гггг2-гггг1}': _years_range,
    }
    _replace_in_el(doc.element.body, _ctx)
    for _sec in doc.sections:
        try: _replace_in_el(_sec.header._element, _ctx)
        except Exception: pass
        try: _replace_in_el(_sec.footer._element, _ctx)
        except Exception: pass

    # ================= Перестройка таблицы «по районам» =================
    if not doc.tables:
        buf = BytesIO(); doc.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name=f"protokol_priyoma_{tournament_id}.docx",
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    tbl = doc.tables[0]
    tbl_el = tbl._tbl
    trs = tbl_el.findall(_qn('w:tr'))
    if len(trs) < 16:
        buf = BytesIO(); doc.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name=f"protokol_priyoma_{tournament_id}.docx",
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    header0 = _copy.deepcopy if False else None
    proto_row = __import__('copy').deepcopy(trs[2])   # строка-прототип района
    total_row = __import__('copy').deepcopy(trs[15])  # строка итога

    # Удаляем строки районов и итога (оставляем два заголовка: стр0 и стр1)
    for tr in trs[2:]:
        tbl_el.remove(tr)

    # ---- Шапка таблицы (строка 0) ---- пересобираем заголовки-группы
    # Оставляем заголовок как есть (Весовые категории кг ЮНОШИ/ДЕВУШКИ),
    # но изменяем gridSpan блока юношей/девушек фактическим числом весов.
    htr = tbl_el.findall(_qn('w:tr'))[0]
    for tc in htr.findall(_qn('w:tc')):
        txt = ''.join(x.text or '' for x in tc.iter(_qn('w:t'))).strip()
        if 'ЮНОШИ' in txt and 'категории' in txt:
            _tcp = tc.find(_qn('w:tcPr')); gs = _tcp.find(_qn('w:gridSpan')) if _tcp is not None else None
            if gs is not None:
                gs.set(_qn('w:val'), str(max(1, len(male_weights))))
        if 'ДЕВУШКИ' in txt and 'категории' in txt:
            _tcp = tc.find(_qn('w:tcPr')); gs = _tcp.find(_qn('w:gridSpan')) if _tcp is not None else None
            if gs is not None:
                gs.set(_qn('w:val'), str(max(1, len(female_weights))))

    # ---- Шапка строки 1: перечень весов ----
    # Берём строку 1 (весовые колонки) и заменяем: {vm}->первый вес, {vd}->первый вес
    # но фактически нужно разложить веса по колонкам. Учитывая, что колонок в базовом
    # шаблоне 12/6, а фактических весов может быть 25 — пересобираем строку заново.
    str1 = tbl_el.findall(_qn('w:tr'))[1]
    # Сохраним прототип одной весовой ячейки из шаблона (например {vm} или первую цифру)
    base_tcs = str1.findall(_qn('w:tc'))
    weight_proto = None
    for tc in base_tcs:
        txt = ''.join(x.text or '' for x in tc.iter(_qn('w:t'))).strip()
        if txt and _re.match(r'^\d+$', txt):
            weight_proto = __import__('copy').deepcopy(tc)
            break
    if weight_proto is None and base_tcs:
        # последняя прибежище — использовать {vm} ячейку
        for tc in base_tcs:
            if ''.join(x.text or '' for x in tc.iter(_qn('w:t'))).strip() == '{vm}':
                weight_proto = __import__('copy').deepcopy(tc)
                break
    if weight_proto is None:
        weight_proto = __import__('copy').deepcopy(base_tcs[0])

    # Удаляем строку 1 шапки и строим новую
    tbl_el.remove(str1)
    # Прототипы разделителей: пустая ячейка КОЛ-ВО уч-ниц
    sep_proto = None
    for tc in base_tcs:
        txt = ''.join(x.text or '' for x in tc.iter(_qn('w:t'))).strip()
        if txt == '' and sep_proto is None:
            sep_proto = __import__('copy').deepcopy(tc)
    if sep_proto is None:
        sep_proto = __import__('copy').deepcopy(base_tcs[4])

    new_row = __import__('copy').deepcopy(str1)
    new_row.clear()
    # 3 пустые ячейки (№, название, КОЛ-ВО уч-ков м.п.)
    for _ in range(3):
        c = __import__('copy').deepcopy(sep_proto)
        set_cell_text(c, '')
        new_row.append(c)
    # веса юношей
    for w in male_weights:
        c = __import__('copy').deepcopy(weight_proto)
        set_cell_text(c, _fmt_w(w))
        new_row.append(c)
    # разделитель (КОЛ-ВО уч-ниц)
    c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, ''); new_row.append(c)
    # веса девушек
    for w in female_weights:
        c = __import__('copy').deepcopy(weight_proto)
        set_cell_text(c, _fmt_w(w))
        new_row.append(c)

    # вставляем строку 1 сразу после строки 0
    _row0 = tbl_el.findall(_qn('w:tr'))[0]
    _row0.addnext(new_row)

    # ---- Строки районов ----
    from collections import defaultdict as _cc
    def _row_vals(rname):
        mcount = _cc(int); dcount = _cc(int)
        for s in lottery:
            a = by_id.get(s['athlete_id'])
            if a is None: continue
            if _canon_region(a.get('region')) != rname: continue
            g = _wc_gender(s['weight_class']); w = _wc_weight(s['weight_class'])
            if g == 'мужской': mcount[w] += 1
            elif g == 'женский': dcount[w] += 1
        return mcount, dcount

    for idx, rname in enumerate(rayon_order, start=1):
        mcount, dcount = _row_vals(rname)
        nm = sum(mcount.values()); nd = sum(dcount.values())
        nrow = __import__('copy').deepcopy(proto_row)
        nrow.clear()
        cells = []
        c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, str(idx)); cells.append(c)
        c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, rname); cells.append(c)
        c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, str(nm)); cells.append(c)  # КОЛ-ВО уч-ков (м.п.)
        for w in male_weights:
            v = mcount.get(w, 0)
            c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, str(v) if v else ''); cells.append(c)
        c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, str(nd)); cells.append(c)  # КОЛ-ВО уч-ниц
        for w in female_weights:
            v = dcount.get(w, 0)
            c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, str(v) if v else ''); cells.append(c)
        for cell in cells:
            nrow.append(cell)
        _last = tbl_el[-1]
        tbl_el.append(nrow)

    # ---- Итоговая строка ----
    all_m = 0; all_d = 0
    for s in lottery:
        a = by_id.get(s['athlete_id'])
        if a is None: continue
        if _canon_region(a.get('region')) is None: continue
        g = _wc_gender(s['weight_class'])
        if g == 'мужской': all_m += 1
        elif g == 'женский': all_d += 1
    ntotal = __import__('copy').deepcopy(total_row)
    ntotal.clear()
    cells = []
    c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, ''); cells.append(c)
    c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, 'ВСЕГО:'); cells.append(c)
    c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, str(all_m)); cells.append(c)  # itum — всего участников мужчин
    for w in male_weights:
        v = sum(1 for s in lottery
                if _wc_gender(s['weight_class']) == 'мужской'
                and abs(_wc_weight(s['weight_class']) - w) < 1e-6
                and (by_id.get(s['athlete_id']) and _canon_region(by_id.get(s['athlete_id']).get('region'))))
        c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, str(v) if v else ''); cells.append(c)
    c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, str(all_d)); cells.append(c)
    for w in female_weights:
        v = sum(1 for s in lottery
                if _wc_gender(s['weight_class']) == 'женский'
                and abs(_wc_weight(s['weight_class']) - w) < 1e-6
                and (by_id.get(s['athlete_id']) and _canon_region(by_id.get(s['athlete_id']).get('region'))))
        c = __import__('copy').deepcopy(sep_proto); set_cell_text(c, str(v) if v else ''); cells.append(c)
    for cell in cells:
        ntotal.append(cell)
    tbl_el.append(ntotal)

    # ================= Таблица «по возрасту» =================
    # Строки — уникальные года рождения участников (от наименьшего).
    # Группировка та же, что и в таблице по районам, но по полю birth_year.
    def _ath_year(a):
        try:
            y = int(str((a or {}).get('birth_year') or '').strip())
            return y if y > 0 else None
        except Exception:
            return None

    if len(doc.tables) > 1:
        t_age = doc.tables[1]
        age_el = t_age._tbl
        atrs = age_el.findall(_qn('w:tr'))
        if len(atrs) >= 6:
            age_proto = __import__('copy').deepcopy(atrs[2])   # строка-год {yro}
            age_total = __import__('copy').deepcopy(atrs[5])   # строка итога
            for tr in atrs[2:]:
                age_el.remove(tr)

            # ---- Шапка строки 1: перечень весов ----
            ah = age_el.findall(_qn('w:tr'))[1]
            base_age_tcs = ah.findall(_qn('w:tc'))
            weight_proto_a = None
            for tc in base_age_tcs:
                txt = ''.join(x.text or '' for x in tc.iter(_qn('w:t'))).strip()
                if (_re.match(r'^\d+$', txt) or txt == '{vm}') and weight_proto_a is None:
                    weight_proto_a = __import__('copy').deepcopy(tc)
            if weight_proto_a is None and base_age_tcs:
                weight_proto_a = __import__('copy').deepcopy(base_age_tcs[0])
            sep_proto_a = None
            for tc in base_age_tcs:
                txt = ''.join(x.text or '' for x in tc.iter(_qn('w:t'))).strip()
                if txt == '' and sep_proto_a is None:
                    sep_proto_a = __import__('copy').deepcopy(tc)
            if sep_proto_a is None and base_age_tcs:
                sep_proto_a = __import__('copy').deepcopy(base_age_tcs[0])
            age_el.remove(ah)

            # Шапка таблицы (строка 0): обновляем gridSpan блоков ЮНОШИ/ДЕВУШКИ
            # под фактическое число весов (как для таблицы по районам).
            ahtr = age_el.findall(_qn('w:tr'))[0]
            for tc in ahtr.findall(_qn('w:tc')):
                txt = ''.join(x.text or '' for x in tc.iter(_qn('w:t'))).strip()
                if 'ЮНОШИ' in txt and 'категории' in txt:
                    _tcp = tc.find(_qn('w:tcPr')); gs = _tcp.find(_qn('w:gridSpan')) if _tcp is not None else None
                    if gs is not None:
                        gs.set(_qn('w:val'), str(max(1, len(male_weights))))
                if 'ДЕВУШКИ' in txt and 'категории' in txt:
                    _tcp = tc.find(_qn('w:tcPr')); gs = _tcp.find(_qn('w:gridSpan')) if _tcp is not None else None
                    if gs is not None:
                        gs.set(_qn('w:val'), str(max(1, len(female_weights))))

            new_h0 = __import__('copy').deepcopy(ah); new_h0.clear()
            for _ in range(3):
                c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, ''); new_h0.append(c)
            for w in male_weights:
                c = __import__('copy').deepcopy(weight_proto_a); set_cell_text(c, _fmt_w(w)); new_h0.append(c)
            c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, ''); new_h0.append(c)
            for w in female_weights:
                c = __import__('copy').deepcopy(weight_proto_a); set_cell_text(c, _fmt_w(w)); new_h0.append(c)
            age_el.findall(_qn('w:tr'))[0].addnext(new_h0)

            # ---- Группировка участников по году рождения ----
            years = set()
            for s in lottery:
                a = by_id.get(s['athlete_id'])
                y = _ath_year(a)
                if y is not None:
                    years.add(y)
            for y_idx, y in enumerate(sorted(years), start=1):
                mm = _cc(int); dd = _cc(int)
                nm = 0; nd = 0
                for s in lottery:
                    a = by_id.get(s['athlete_id'])
                    if _ath_year(a) != y:
                        continue
                    g = _wc_gender(s['weight_class']); w = _wc_weight(s['weight_class'])
                    if g == 'мужской':
                        mm[w] += 1; nm += 1
                    elif g == 'женский':
                        dd[w] += 1; nd += 1
                nrow = __import__('copy').deepcopy(age_proto); nrow.clear()
                cells = []
                c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, str(y_idx)); cells.append(c)
                c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, str(y)); cells.append(c)
                c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, str(nm)); cells.append(c)  # КОЛ-ВО (допущено мужчин)
                for w in male_weights:
                    v = mm.get(w, 0)
                    c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, str(v) if v else ''); cells.append(c)
                c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, str(nd)); cells.append(c)  # КОЛ-ВО уч-ниц (допущено девушек)
                for w in female_weights:
                    v = dd.get(w, 0)
                    c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, str(v) if v else ''); cells.append(c)
                for cell in cells:
                    nrow.append(cell)
                age_el.append(nrow)

            # ---- Итоговая строка таблицы по возрасту ----
            tm_all = 0; td_all = 0
            for s in lottery:
                a = by_id.get(s['athlete_id'])
                if _ath_year(a) is None:
                    continue
                g = _wc_gender(s['weight_class'])
                if g == 'мужской': tm_all += 1
                elif g == 'женский': td_all += 1
            nt2 = __import__('copy').deepcopy(age_total); nt2.clear()
            cells = []
            c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, ''); cells.append(c)
            c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, 'ВСЕГО:'); cells.append(c)
            c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, str(tm_all)); cells.append(c)  # itumv — итого участников мужчин по возрасту
            # (было: отдельный столбец «ВСЕГО уч-ков» всех — убран из шаблона)
            for w in male_weights:
                v = sum(1 for s in lottery
                        if _wc_gender(s['weight_class']) == 'мужской'
                        and abs(_wc_weight(s['weight_class']) - w) < 1e-6
                        and _ath_year(by_id.get(s['athlete_id'])) is not None)
                c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, str(v) if v else ''); cells.append(c)  # itkmvy
            c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, str(td_all)); cells.append(c)  # itddv — допущено девушек
            for w in female_weights:
                v = sum(1 for s in lottery
                        if _wc_gender(s['weight_class']) == 'женский'
                        and abs(_wc_weight(s['weight_class']) - w) < 1e-6
                        and _ath_year(by_id.get(s['athlete_id'])) is not None)
                c = __import__('copy').deepcopy(sep_proto_a); set_cell_text(c, str(v) if v else ''); cells.append(c)  # itkdvy
            for cell in cells:
                nt2.append(cell)
            age_el.append(nt2)

            # ---- Пересобираем сетку колонок таблицы по возрасту ----
            _ncols_a = 3 + len(male_weights) + 1 + len(female_weights)
            age_grid = age_el.find(_qn('w:tblGrid'))
            if age_grid is not None:
                age_grid.clear()
                _geom_a = [500, 900, 1300]   # №, год рождения, КОЛ-ВО уч-ков (м.п.)
                _rest_a = _total_w - sum(_geom_a)
                _n_rest_a = max(1, _ncols_a - 3)
                _w_rest_a = int(_rest_a / _n_rest_a)
                _col_w_a = list(_geom_a)
                for _ in range(_ncols_a - 3):
                    _col_w_a.append(_w_rest_a)
                for _w in _col_w_a:
                    el = _px('<w:gridCol xmlns:w="%s" w:w="%d"/>' % (_WNS, _w))
                    age_grid.append(el)
                _atpr = age_el.find(_qn('w:tblPr'))
                if _atpr is not None:
                    _atw = _atpr.find(_qn('w:tblW'))
                    if _atw is None:
                        _atw = _px('<w:tblW xmlns:w="%s"/>' % _WNS)
                        _atpr.append(_atw)
                    _atw.set(_qn('w:w'), str(_total_w)); _atw.set(_qn('w:type'), 'dxa')
                    _atl = _atpr.find(_qn('w:tblLayout'))
                    if _atl is None:
                        _atl = _px('<w:tblLayout xmlns:w="%s"/>' % _WNS)
                        _atpr.append(_atl)
                    _atl.set(_qn('w:type'), 'fixed')
                for tr in age_el.findall(_qn('w:tr')):
                    _ci = 0
                    for tc in tr.findall(_qn('w:tc')):
                        _tcp = tc.find(_qn('w:tcPr')); gs = _tcp.find(_qn('w:gridSpan')) if _tcp is not None else None
                        span = int(gs.get(_qn('w:val'))) if gs is not None else 1
                        _sum = sum(_col_w_a[_ci:_ci + span])
                        _tcp2 = tc.find(_qn('w:tcPr')); _twc = _tcp2.find(_qn('w:tcW')) if _tcp2 is not None else None
                        if _twc is None:
                            _twc = _px('<w:tcW xmlns:w="%s"/>' % _WNS)
                            tc.find(_qn('w:tcPr')).append(_twc)
                        _twc.set(_qn('w:w'), str(_sum)); _twc.set(_qn('w:type'), 'dxa')
                        _ci += span

    # ================= Таблица «по спортивной квалификации» =================
    # Строки — уникальные спортивные разряды участников (из поля rank),
    # отсортированные: сначала кмс, затем юношеские по возрастанию номера.
    def _rank_key(rank):
        r = (rank or '').strip().lower()
        if 'кмс' in r:
            return (0, 0)
        m = _re.search(r'(\d+)', r)
        if m:
            return (1, int(m.group(1)))
        return (2, 0)

    def _short_sk(rank):
        r = (rank or '').strip()
        if not r:
            return r
        # 'кмс' -> 'кмс', '1 юн. разряд' -> '1 юн', '2 разряд' -> '2 р'
        base = r.lower()
        if base.startswith('кмс'):
            return 'кмс'
        m = _re.search(r'(\d+)', r)
        if not m:
            return r
        d = m.group(1)
        if 'юн' in base:
            return f'{d} юн'
        if 'нов' in base or 'н' in base:
            return 'н/р'
        return f'{d} р'

    if len(doc.tables) > 2:
        t_sk = doc.tables[2]
        sk_el = t_sk._tbl
        strs = sk_el.findall(_qn('w:tr'))
        if len(strs) >= 7:
            sk_proto = __import__('copy').deepcopy(strs[2])   # строка-разряд {sk},{u},{dm}
            sk_total = __import__('copy').deepcopy(strs[6])   # строка итога
            for tr in strs[2:]:
                sk_el.remove(tr)

            # ---- Шапка таблицы: обновляем gridSpan блоков ЮНОШИ/ДЕВУШКИ ----
            skhtr = sk_el.findall(_qn('w:tr'))[0]
            for tc in skhtr.findall(_qn('w:tc')):
                txt = ''.join(x.text or '' for x in tc.iter(_qn('w:t'))).strip()
                if 'ЮНОШИ' in txt and 'категории' in txt:
                    _tcp = tc.find(_qn('w:tcPr')); gs = _tcp.find(_qn('w:gridSpan')) if _tcp is not None else None
                    if gs is not None:
                        gs.set(_qn('w:val'), str(max(1, len(male_weights))))
                if 'ДЕВУШКИ' in txt and 'категории' in txt:
                    _tcp = tc.find(_qn('w:tcPr')); gs = _tcp.find(_qn('w:gridSpan')) if _tcp is not None else None
                    if gs is not None:
                        gs.set(_qn('w:val'), str(max(1, len(female_weights))))

            # ---- Шапка весов (строка 1) ----
            sh = sk_el.findall(_qn('w:tr'))[1]
            base_sk_tcs = sh.findall(_qn('w:tc'))
            weight_proto_s = None
            for tc in base_sk_tcs:
                txt = ''.join(x.text or '' for x in tc.iter(_qn('w:t'))).strip()
                if (_re.match(r'^\d+$', txt) or txt == '{vm}') and weight_proto_s is None:
                    weight_proto_s = __import__('copy').deepcopy(tc)
            if weight_proto_s is None and base_sk_tcs:
                weight_proto_s = __import__('copy').deepcopy(base_sk_tcs[0])
            sep_proto_s = None
            for tc in base_sk_tcs:
                txt = ''.join(x.text or '' for x in tc.iter(_qn('w:t'))).strip()
                if txt == '' and sep_proto_s is None:
                    sep_proto_s = __import__('copy').deepcopy(tc)
            if sep_proto_s is None and base_sk_tcs:
                sep_proto_s = __import__('copy').deepcopy(base_sk_tcs[0])
            sk_el.remove(sh)

            new_h1 = __import__('copy').deepcopy(sh); new_h1.clear()
            for _ in range(3):
                c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, ''); new_h1.append(c)
            for w in male_weights:
                c = __import__('copy').deepcopy(weight_proto_s); set_cell_text(c, _fmt_w(w)); new_h1.append(c)
            c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, ''); new_h1.append(c)
            for w in female_weights:
                c = __import__('copy').deepcopy(weight_proto_s); set_cell_text(c, _fmt_w(w)); new_h1.append(c)
            sk_el.findall(_qn('w:tr'))[0].addnext(new_h1)

            # ---- Группировка участников по разряду ----
            ranks = set()
            for s in lottery:
                a = by_id.get(s['athlete_id'])
                rk = (a.get('rank') or '').strip() if a else ''
                if rk:
                    ranks.add(rk)
            ordered_ranks = sorted(ranks, key=_rank_key)

            _mmW = _cc(int); _ddW = _cc(int)
            for rk_idx, rk_name in enumerate(ordered_ranks, start=1):
                mc = _cc(int); dc = _cc(int)
                nm = 0; nd = 0
                for s in lottery:
                    a = by_id.get(s['athlete_id'])
                    if a is None or (a.get('rank') or '').strip() != rk_name:
                        continue
                    g = _wc_gender(s['weight_class']); w = _wc_weight(s['weight_class'])
                    if g == 'мужской':
                        mc[w] += 1; nm += 1; _mmW[w] += 1
                    elif g == 'женский':
                        dc[w] += 1; nd += 1; _ddW[w] += 1
                nrow = __import__('copy').deepcopy(sk_proto); nrow.clear()
                cells = []
                c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, str(rk_idx)); cells.append(c)
                c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, _short_sk(rk_name)); cells.append(c)
                c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, str(nm)); cells.append(c)  # КОЛ-ВО уч-ков (м.п.)
                for w in male_weights:
                    v = mc.get(w, 0)
                    c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, str(v) if v else ''); cells.append(c)
                c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, str(nd)); cells.append(c)
                for w in female_weights:
                    v = dc.get(w, 0)
                    c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, str(v) if v else ''); cells.append(c)
                for cell in cells:
                    nrow.append(cell)
                sk_el.append(nrow)

            # ---- Итоговая строка таблицы по квалификации ----
            it_all = 0; it_m = 0; it_d = 0
            for s in lottery:
                a = by_id.get(s['athlete_id'])
                if a is None or not (a.get('rank') or '').strip():
                    continue
                g = _wc_gender(s['weight_class'])
                if g == 'мужской': it_m += 1
                elif g == 'женский': it_d += 1
            it_all = it_m + it_d
            nt3 = __import__('copy').deepcopy(sk_total); nt3.clear()
            cells = []
            c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, ''); cells.append(c)
            c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, 'ВСЕГО:'); cells.append(c)
            c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, str(it_m)); cells.append(c)    # itkum — итого участников мужчин по квалификации
            for w in male_weights:
                v = _mmW.get(w, 0)
                c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, str(v) if v else ''); cells.append(c)  # ikuvm
            c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, str(it_d)); cells.append(c)   # ikdd
            for w in female_weights:
                v = _ddW.get(w, 0)
                c = __import__('copy').deepcopy(sep_proto_s); set_cell_text(c, str(v) if v else ''); cells.append(c)  # ikuvd
            for cell in cells:
                nt3.append(cell)
            sk_el.append(nt3)

            # ---- Пересобираем сетку колонок таблицы по квалификации ----
            _ncols_s = 3 + len(male_weights) + 1 + len(female_weights)
            sk_grid = sk_el.find(_qn('w:tblGrid'))
            if sk_grid is not None:
                sk_grid.clear()
                _geom_s = [500, 1300, 1300]   # №, спорт. квалификация, КОЛ-ВО уч-ков (м.п.)
                _rest_s = _total_w - sum(_geom_s)
                _n_rest_s = max(1, _ncols_s - 3)
                _w_rest_s = int(_rest_s / _n_rest_s)
                _col_w_s = list(_geom_s)
                for _ in range(_ncols_s - 3):
                    _col_w_s.append(_w_rest_s)
                for _w in _col_w_s:
                    el = _px('<w:gridCol xmlns:w="%s" w:w="%d"/>' % (_WNS, _w))
                    sk_grid.append(el)
                _spr = sk_el.find(_qn('w:tblPr'))
                if _spr is not None:
                    _stw = _spr.find(_qn('w:tblW'))
                    if _stw is None:
                        _stw = _px('<w:tblW xmlns:w="%s"/>' % _WNS)
                        _spr.append(_stw)
                    _stw.set(_qn('w:w'), str(_total_w)); _stw.set(_qn('w:type'), 'dxa')
                    _stl = _spr.find(_qn('w:tblLayout'))
                    if _stl is None:
                        _stl = _px('<w:tblLayout xmlns:w="%s"/>' % _WNS)
                        _spr.append(_stl)
                    _stl.set(_qn('w:type'), 'fixed')
                for tr in sk_el.findall(_qn('w:tr')):
                    _ci = 0
                    for tc in tr.findall(_qn('w:tc')):
                        _tcp = tc.find(_qn('w:tcPr')); gs = _tcp.find(_qn('w:gridSpan')) if _tcp is not None else None
                        span = int(gs.get(_qn('w:val'))) if gs is not None else 1
                        _sum = sum(_col_w_s[_ci:_ci + span])
                        _tcp2 = tc.find(_qn('w:tcPr')); _twc = _tcp2.find(_qn('w:tcW')) if _tcp2 is not None else None
                        if _twc is None:
                            _twc = _px('<w:tcW xmlns:w="%s"/>' % _WNS)
                            tc.find(_qn('w:tcPr')).append(_twc)
                        _twc.set(_qn('w:w'), str(_sum)); _twc.set(_qn('w:type'), 'dxa')
                        _ci += span

    # ---- Пересобираем сетку колонок (tblGrid) под фактическое число колонок ----
    # Колонки: 3 (№, район, КОЛ-ВО уч-ков м.п.) + мужские веса + 1 (разделитель) + женские веса
    _ncols = 3 + len(male_weights) + 1 + len(female_weights)
    _grid = tbl_el.find(_qn('w:tblGrid'))
    if _grid is not None:
        _grid.clear()
        # итоговая ширина таблицы (dxa); поля документа лево/право 300 -> текст ~11306
        _total_w = 11282
        # геометрия фиксированных колонок: №, район, КОЛ-ВО; остальное равными долями
        _geom = [500, 2100, 1300]
        _rest = _total_w - sum(_geom)
        _n_rest = max(1, _ncols - 3)
        _w_rest = int(_rest / _n_rest)
        _col_w = list(_geom)
        for _ in range(_ncols - 3):
            _col_w.append(_w_rest)
        for _w in _col_w:
            el = _px('<w:gridCol xmlns:w="%s" w:w="%d"/>' % (_WNS, _w))
            _grid.append(el)
        # fixed layout + ширина таблицы
        _tpr = tbl_el.find(_qn('w:tblPr'))
        if _tpr is not None:
            # tblW
            _tw = _tpr.find(_qn('w:tblW'))
            if _tw is None:
                _tw = _px('<w:tblW xmlns:w="%s"/>' % _WNS)
                _tpr.append(_tw)
            _tw.set(_qn('w:w'), str(_total_w))
            _tw.set(_qn('w:type'), 'dxa')
            # tblLayout fixed
            _tl = _tpr.find(_qn('w:tblLayout'))
            if _tl is None:
                _tl = _px('<w:tblLayout xmlns:w="%s"/>' % _WNS)
                _tpr.append(_tl)
            _tl.set(_qn('w:type'), 'fixed')
        # нормализуем ширины всех ячеек всех строк по сетке
        for tr in tbl_el.findall(_qn('w:tr')):
            _ci = 0
            for tc in tr.findall(_qn('w:tc')):
                _tcp = tc.find(_qn('w:tcPr')); gs = _tcp.find(_qn('w:gridSpan')) if _tcp is not None else None
                span = int(gs.get(_qn('w:val'))) if gs is not None else 1
                # суммарная ширина по сетке для этого span
                _tcw_sum = sum(_col_w[_ci:_ci + span])
                _tcp2 = tc.find(_qn('w:tcPr')); _tccw = _tcp2.find(_qn('w:tcW')) if _tcp2 is not None else None
                if _tccw is None:
                    _tccw = _px('<w:tcW xmlns:w="%s"/>' % _WNS)
                    tc.find(_qn('w:tcPr')).append(_tccw)
                _tccw.set(_qn('w:w'), str(_tcw_sum))
                _tccw.set(_qn('w:type'), 'dxa')
                _ci += span

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"protokol_priyoma_{tournament_id}.docx",
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/api/report/<int:tournament_id>/grso', methods=['GET'])
def report_grso(tournament_id):
    """Сформировать «Протокол графиков соревнований» на основе grso_tmp.docx.

    Шаблон содержит: заголовок ({TOURNAMENT_NAME}, {DATE1-DATE2_YEAR}),
    абзац-прототип с {WEIGHT_CLASS} и ОДНУ таблицу-прототип на 4 слота
    ({SLOT1_NAME},{SLOT1_CLUB},{SLOT1_YEAR}, {SLOTS-1-2_COACH}, ... {SLOTS_CHAMPION}).

    Под каждую весовую категорию дублируется блок «абзац + таблица» и заполняется
    данными из жеребьёвки и bracket_results. Победители пар — из результатов боя.
    """
    if not _HAS_DOCX:
        return jsonify({'error': 'Библиотека Word не установлена (python-docx)'}), 500
    from io import BytesIO
    import re as _re
    import copy as _copy
    from collections import defaultdict as _dd
    from docx import Document as _Doc

    # Найти шаблон grso_new.docx (компактная «ёлка», ширины синхронны с кодом)
    tmpl = _report_template('grso_new.docx')
    if not tmpl:
        return jsonify({'error': 'Шаблон протокола графиков (grso_new.docx) не найден'}), 500

    # Данные турнира и жеребьёвки
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

    # Группируем по весовым категориям
    groups = _dd(list); order = []
    for s in lottery:
        if s['weight_class'] not in groups:
            order.append(s['weight_class'])
        groups[s['weight_class']].append(s)
    for wc in groups:
        groups[wc].sort(key=lambda x: x['slot'])
    order.sort(key=lambda wc: (
        0 if ('жен' in wc) else 1,  # женские категории идут первыми
        float((_re.search(r'(\d+(?:\.\d+)?)', wc or '') or _re.search(r'0', '0')).group(1) or 0)
    ))

    def _fmt_date(s):
        p = (s or '').split('-')
        return f'{p[2]}.{p[1]}.{p[0]}' if len(p) == 3 else (s or '').replace('-', '.')
    d1 = _fmt_date(t.get('date_from') or '')
    d2 = _fmt_date(t.get('date_to') or t.get('date_from') or '')
    dates = f'{d1} — {d2}' if d2 and d2 != d1 else d1

    from docx.oxml.ns import qn as _qn
    from docx.oxml import parse_xml as _px
    _WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    def _all_t(el):
        return ''.join(_n.text or '' for _n in el.iter(_qn('w:t')))

    def _replace_in_el(el, mapping):
        # объединяем текст и заменяем маркеры (маркер может быть разбит на w:t)
        for p in el.iter(_qn('w:p')):
            runs = p.findall(_qn('w:r'))
            if not runs:
                continue
            full = ''.join(rt.text or '' for r in runs for rt in r.findall(_qn('w:t')))
            changed = False
            for k, v in mapping.items():
                if k in full:
                    full = full.replace(k, str(v))
                    changed = True
            if changed:
                ts0 = runs[0].findall(_qn('w:t'))
                if ts0:
                    ts0[0].text = full
                    for r in runs[1:]:
                        for rt in r.findall(_qn('w:t')):
                            rt.text = ''

    doc = _Doc(tmpl)

    # Заголовок (колонтитулы и тело)
    for sec in doc.sections:
        for pn in ('header', 'footer', 'first_page_header', 'first_page_footer'):
            try:
                _replace_in_el(getattr(sec, pn)._element, {
                    '{TOURNAMENT_NAME}': t.get('name') or '',
                    '{DATE1-DATE2_YEAR}': dates,
                    '{НАЗВАНИЕ_ТУРНИРА}': t.get('name') or '',
                    '{ДатыПроведения}': dates,
                    '{МестоПроведения}': (t.get('venue') or ''),
                    '{ГлавныйСудьяФИО}': (t.get('chief_referee') or ''),
                    '{ГлавныйСекретарьФИО}': (t.get('chief_secretary') or ''),
                })
            except Exception:
                pass
    _replace_in_el(doc.element.body, {
        '{TOURNAMENT_NAME}': t.get('name') or '',
        '{DATE1-DATE2_YEAR}': dates,
    })

    # Заменяем маркеры в footer'ах (подписи судей) отдельно надёжно
    for sec in doc.sections:
        for pn in ('footer', 'first_page_footer'):
            try:
                _replace_in_el(getattr(sec, pn)._element, {
                    '{ГлавныйСудьяФИО}': (t.get('chief_referee') or ''),
                    '{ГлавныйСудьяРегион}': '',
                    '{ГлавныйСекретарьФИО}': (t.get('chief_secretary') or ''),
                    '{ГлавныйСекретарьРегион}': '',
                })
            except Exception:
                pass

    # Находим прототип: абзац с {WEIGHT_CLASS} + следующую за ним таблицу
    body = doc.element.body
    _W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    proto_para = None
    proto_tbl = None
    children = list(body)
    for i, ch in enumerate(children):
        if ch.tag == _W + 'p' and '{WEIGHT_CLASS}' in _all_t(ch):
            proto_para = ch
            # следующая за ним таблица
            for j in range(i + 1, len(children)):
                if children[j].tag == _W + 'tbl':
                    proto_tbl = children[j]
                    break
            break
    if proto_para is None or proto_tbl is None:
        return jsonify({'error': 'В шаблоне grso_tmp.docx не найден блок-прототип'}), 500

    def _name(slot_id):
        a = by_id.get(slot_id)
        return a['name'] if a else ''
    def _club(slot_id):
        a = by_id.get(slot_id)
        return ((a.get('region') or '').replace('Калмыкия', '').strip()) if a else ''
    def _year(slot_id):
        a = by_id.get(slot_id)
        try:
            return int(str(a['birth_year'] or '').strip()) if a and a.get('birth_year') else ''
        except Exception:
            return ''
    def _winner_of(pair):
        # pair = (a1, a2, winner_id, method)
        return pair[2] if pair and len(pair) >= 3 and pair[2] is not None else None

    # --- Динамическое построение турнирной сетки на ВСЕХ бойцов категории ---
    # (перенесено из _build_report_document: генерирует таблицу-дерево на N слотов)
    _copy2 = _copy
    _atoms = {}
    if proto_tbl is not None:
        _rows2 = proto_tbl.findall(_qn('w:tr'))
        def _cells_at2(tr_el):
            return tr_el.findall(_qn('w:tc'))
        if len(_rows2) >= 4:
            _r0 = _rows2[0]; _r1 = _rows2[1]; _r3 = _rows2[3]
            _atoms = {
                'num': _copy2.deepcopy(_cells_at2(_r0)[0]),
                'name': _copy2.deepcopy(_cells_at2(_r0)[1]),
                'club': _copy2.deepcopy(_cells_at2(_r1)[1]),
                'spacer': _copy2.deepcopy(_cells_at2(_r1)[2]),
                'year': _copy2.deepcopy(_cells_at2(_r0)[2]),
                'winner': _copy2.deepcopy(_cells_at2(_r1)[4]),
                'champ': _copy2.deepcopy(_cells_at2(_r1)[6]),
            }
            def _ensure_shd2(tc, fill):
                tcPr = tc.find(_qn('w:tcPr'))
                if tcPr is None:
                    tcPr = _px('<w:tcPr xmlns:w="%s"/>' % _WNS)
                    tc.insert(0, tcPr)
                shd = tcPr.find(_qn('w:shd'))
                if shd is None:
                    shd = _px('<w:shd xmlns:w="%s"/>' % _WNS)
                    tcPr.append(shd)
                shd.set(_qn('w:val'), 'clear'); shd.set(_qn('w:color'), 'auto'); shd.set(_qn('w:fill'), fill)
            def _clear_shd2(tc):
                tcPr = tc.find(_qn('w:tcPr'))
                if tcPr is not None and tcPr.find(_qn('w:shd')) is not None:
                    tcPr.remove(tcPr.find(_qn('w:shd')))
            # без жёлтой заливки чемпиона (подгон под содержимое)
            _atoms['champ_plain'] = _copy2.deepcopy(_atoms['champ'])
            _clear_shd2(_atoms['champ_plain'])
            # champ-атом тоже без фон, чтобы финал не был жёлтым
            _atoms['champ'] = _copy2.deepcopy(_atoms['champ_plain'])

            def _set_tc_text2(tc, text):
                ps = tc.findall(_qn('w:p'))
                if not ps:
                    p = _px('<w:p xmlns:w="%s"/>' % _WNS)
                    tc.append(p); ps = [p]
                p = ps[0]
                runs = p.findall(_qn('w:r'))
                if runs:
                    r0 = runs[0]
                    tt = r0.find(_qn('w:t'))
                    if tt is None:
                        tt = _px('<w:t xmlns:w="%s"/>' % _WNS)
                        r0.append(tt)
                    tt.text = text
                    for rr in runs[1:]:
                        p.remove(rr)
                else:
                    _r = _px('<w:r xmlns:w="%s"><w:t>%s</w:t></w:r>' % (_WNS, text))
                    p.append(_r)
                # уменьшаем шрифт до 8pt (16 half-points) во всех runs, чтобы данные
                # помещались на одной строке в тесных ячейках, не ломая структуры
                try:
                    for _r in p.findall(_qn('w:r')):
                        _rpr = _r.find(_qn('w:rPr'))
                        if _rpr is None:
                            _rpr = _px('<w:rPr xmlns:w="%s"/>' % _WNS)
                            _r.insert(0, _rpr)
                        _sz = _rpr.find(_qn('w:sz'))
                        if _sz is None:
                            _sz = _px('<w:sz xmlns:w="%s"/>' % _WNS)
                            _rpr.append(_sz)
                        _sz.set(_qn('w:val'), '16')  # 8pt
                        _szcs = _rpr.find(_qn('w:szCs'))
                        if _szcs is None:
                            _szcs = _px('<w:szCs xmlns:w="%s"/>' % _WNS)
                            _rpr.append(_szcs)
                        _szcs.set(_qn('w:val'), '16')
                except Exception:
                    pass

            def _clear_tc2(tc):
                for p in tc.findall(_qn('w:p')):
                    for r in p.findall(_qn('w:r')):
                        p.remove(r)

            def _set_borders2(tc, sides, val='single', sz='4'):
                tcPr = tc.find(_qn('w:tcPr'))
                if tcPr is None:
                    tcPr = _px('<w:tcPr xmlns:w="%s"/>' % _WNS)
                    tc.insert(0, tcPr)
                borders = tcPr.find(_qn('w:tcBorders'))
                if borders is None:
                    borders = _px('<w:tcBorders xmlns:w="%s"/>' % _WNS)
                    _tn = tcPr.find(_qn('w:tcW'))
                    if _tn is not None:
                        _tn.addnext(borders)
                    else:
                        tcPr.append(borders)
                for sd in sides:
                    el = borders.find(_qn('w:' + sd))
                    if el is None:
                        el = _px('<w:%s xmlns:w="%s"/>' % (sd, _WNS))
                        borders.append(el)
                    el.set(_qn('w:val'), val); el.set(_qn('w:sz'), sz)
                    el.set(_qn('w:space'), '0'); el.set(_qn('w:color'), 'auto')
                for el in list(borders):
                    sd = el.tag.split('}')[-1]
                    if sd not in sides:
                        borders.remove(el)

            def _build_bracket(members, results_map):
                """members: [(name, club, year), ...]; results_map: pair_index -> name.

                Строит универсальную «ёлку» (single-elimination): n участников ->
                total_slots = 2^ceil(log2(n)) слотов; пары -> полуфиналы -> финал.
                Лесенка рисуется вертикальными объединениями (vMerge) + верхними
                перекладинами-развилками (top=single на середине блока), как в шаблоне.
                """
                n = len(members)
                total_slots = 2 ** math.ceil(math.log2(n)) if n > 1 else 1
                num_rounds = int(math.log2(total_slots)) if total_slots > 1 else 1
                # Компактные ширины (порядок сохранён: № | Имя | Клуб | Год | Лесенка).
                # Суммарная ширина укладывается в страницу (~9000 twips).
                grid_widths = [380, 1600, 1300, 640]
                round1_idx = len(grid_widths)
                _base_budget = 5000
                _rw = max(700, _base_budget // max(1, num_rounds))
                for _r in range(num_rounds):
                    grid_widths.append(_rw)
                total_w = sum(grid_widths)
                tblPr = _px(
                    '<w:tblPr xmlns:w="%s"><w:tblW w:w="%s" w:type="dxa"/>'
                    '<w:tblLayout w:type="fixed"/>'
                    '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" '
                    'w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>' % (_WNS, total_w))
                grid_el = _px('<w:tblGrid xmlns:w="%s"/>' % _WNS)
                for _wd in grid_widths:
                    grid_el.append(_px('<w:gridCol xmlns:w="%s" w:w="%s"/>' % (_WNS, _wd)))
                tbl = _px('<w:tbl xmlns:w="%s"/>' % _WNS)
                tbl.append(tblPr); tbl.append(grid_el)

                nrows = 2 * total_slots
                ncols = len(grid_widths)
                # ячейки левой части (слоты) и раундов — раздельно, чтобы не путать индексы
                # из-за gridSpan. cells[row][колонка сетки 0..], round_cells[row][round]
                cells = [[None]*ncols for _ in range(nrows)]
                round_cells = [[None]*num_rounds for _ in range(nrows)]

                def _gridspan(tc, val):
                    tcPr = tc.find(_qn('w:tcPr'))
                    if tcPr is None:
                        tcPr = _px('<w:tcPr xmlns:w="%s"/>' % _WNS)
                        tc.insert(0, tcPr)
                    gs = tcPr.find(_qn('w:gridSpan'))
                    if gs is None:
                        gs = _px('<w:gridSpan xmlns:w="%s"/>' % _WNS)
                        tcPr.append(gs)
                    gs.set(_qn('w:val'), str(val))

                def _nowrap(tc):
                    # данные на одной строке: запрет переноса внутри ячейки
                    tcPr = tc.find(_qn('w:tcPr'))
                    if tcPr is None:
                        tcPr = _px('<w:tcPr xmlns:w="%s"/>' % _WNS)
                        tc.insert(0, tcPr)
                    if tcPr.find(_qn('w:noWrap')) is None:
                        tcPr.append(_px('<w:noWrap xmlns:w="%s"/>' % _WNS))

                def _set_tcw(tc, width):
                    # перезаписать ширину ячейки (tcW) компактным значением из grid_widths,
                    # иначе Word берёт старый tcW атома из шаблона (540/1587/1701/...)
                    tcPr = tc.find(_qn('w:tcPr'))
                    if tcPr is None:
                        tcPr = _px('<w:tcPr xmlns:w="%s"/>' % _WNS)
                        tc.insert(0, tcPr)
                    tcW = tcPr.find(_qn('w:tcW'))
                    if tcW is None:
                        tcW = _px('<w:tcW xmlns:w="%s"/>' % _WNS)
                        tcPr.insert(0, tcW)
                    tcW.set(_qn('w:type'), 'dxa')
                    tcW.set(_qn('w:w'), str(width))

                def _fill_left(tr, is_name_row, slot_idx, sl):
                    # name: [№][имя(gridSpan=2)] [год]     -> занимает сеточные колонки 0,1-2,3
                    # club: [№][клуб] [пусто] [пусто]       -> занимает сеточные колонки 0,1,2,3
                    if is_name_row:
                        num_tc = _copy2.deepcopy(_atoms['num']); _set_tc_text2(num_tc, str(slot_idx+1) if sl else '')
                        name_tc = _copy2.deepcopy(_atoms['name']); _set_tc_text2(name_tc, sl[0] if sl else '')
                        _gridspan(name_tc, 2)
                        yr_tc = _copy2.deepcopy(_atoms['year']); _set_tc_text2(yr_tc, sl[2] if sl else '')
                        _set_borders2(yr_tc, ['right'])
                        _nowrap(num_tc); _nowrap(name_tc); _nowrap(yr_tc)
                        # компактные ширины: №=380, имя(span2)=1600+1300, год=640
                        _set_tcw(num_tc, 380)
                        _set_tcw(name_tc, 1600 + 1300)
                        _set_tcw(yr_tc, 640)
                        tcs = [num_tc, name_tc, yr_tc]
                    else:
                        num_tc = _copy2.deepcopy(_atoms['num']); _clear_tc2(num_tc)
                        club_tc = _copy2.deepcopy(_atoms['club']); _set_tc_text2(club_tc, sl[1] if sl else '')
                        sp1_tc = _copy2.deepcopy(_atoms['spacer']); _clear_tc2(sp1_tc)
                        sp2_tc = _copy2.deepcopy(_atoms['spacer']); _clear_tc2(sp2_tc)
                        _set_borders2(sp2_tc, ['right'])
                        _nowrap(num_tc); _nowrap(club_tc); _nowrap(sp1_tc); _nowrap(sp2_tc)
                        # №=380, клуб=1600, спейсер1=1300, спейсер2(=год зона)=640
                        _set_tcw(num_tc, 380)
                        _set_tcw(club_tc, 1600)
                        _set_tcw(sp1_tc, 1300)
                        _set_tcw(sp2_tc, 640)
                        tcs = [num_tc, club_tc, sp1_tc, sp2_tc]
                    for tc in tcs:
                        tr.append(tc)
                    return tcs

                # 1) Левая часть (слоты) и ПУСТЫЕ ячейки раундов (по одному tc в колонке)
                for _i in range(total_slots):
                    sl = members[_i] if _i < n else None
                    nr = 2*_i; cr = nr+1
                    trN = _px('<w:tr xmlns:w="%s"/>' % _WNS)
                    leftN = _fill_left(trN, True, _i, sl)
                    trC = _px('<w:tr xmlns:w="%s"/>' % _WNS)
                    leftC = _fill_left(trC, False, _i, sl)
                    # раунд-колонки: отдельные tc в каждой строке; индексы — абсолютные
                    for _rr_idx in (nr, cr):
                        _tr = trN if _rr_idx == nr else trC
                        left = leftN if _rr_idx == nr else leftC
                        # записываем левые ячейки в cells по их абс. колонкам
                        for _k, _tc in enumerate(left):
                            cells[_rr_idx][_k] = _tc
                        for _c in range(num_rounds):
                            _col = round1_idx + _c
                            base = _atoms['champ'] if (_c == num_rounds-1) else _atoms['winner']
                            tc = _copy2.deepcopy(base); _clear_tc2(tc)
                            _set_tcw(tc, grid_widths[_col])  # компактная ширина раунда
                            _nowrap(tc)
                            _tr.append(tc)
                            cells[_rr_idx][_col] = tc
                            round_cells[_rr_idx][_c] = tc
                    tbl.append(trN); tbl.append(trC)

                # 2) vMerge-блоки раундов + развилки
                def _add_borders2(tc, sides):
                    # добавить (не заменяя) границы tcBorders для списка сторон
                    _set_borders2(tc, sides + list(_existing_sides(tc)))

                def _existing_sides(tc):
                    tcPr = tc.find(_qn('w:tcPr'))
                    res = []
                    if tcPr is not None:
                        b = tcPr.find(_qn('w:tcBorders'))
                        if b is not None:
                            for el in b:
                                sd = el.tag.split('}')[-1]
                                res.append(sd)
                    return res

                for _r in range(num_rounds):
                    gsize = 2 ** (_r + 1)
                    n_groups = total_slots // gsize
                    for _g in range(n_groups):
                        base = _g * gsize
                        top_row = 2 * base
                        bot_row = 2 * (base + gsize) - 1
                        _pi = _g if _r == 0 else (_r * 1000 + _g)
                        winner = results_map.get(_pi, '')
                        block_rows = list(range(top_row, bot_row+1))
                        # верхняя ячейка блока: имя победителя (без жёлтой заливки финала)
                        top_tc = round_cells[top_row][_r]
                        _set_tc_text2(top_tc, winner)
                        _nowrap(top_tc)
                        # границы блока: left на все строки, top-перекладина на середине
                        mid_row = top_row + gsize - 1
                        for ri in block_rows:
                            tc = round_cells[ri][_r]
                            _add_borders2(tc, ['left'])
                            if ri in (top_row, bot_row):
                                _add_borders2(tc, ['right'])
                            if ri == mid_row:
                                _add_borders2(tc, ['top'])
                        # vMerge
                        for idx, ri in enumerate(block_rows):
                            tc = round_cells[ri][_r]
                            tcPr = tc.find(_qn('w:tcPr'))
                            if tcPr is None:
                                tcPr = _px('<w:tcPr xmlns:w="%s"/>' % _WNS)
                                tc.insert(0, tcPr)
                            # удалить существующий vMerge (атом из шаблона его уже содержит)
                            old = tcPr.find(_qn('w:vMerge'))
                            if old is not None:
                                tcPr.remove(old)
                            vm = _px('<w:vMerge xmlns:w="%s"/>' % _WNS)
                            if idx == 0:
                                vm.set(_qn('w:val'), 'restart')
                            else:
                                vm.set(_qn('w:val'), 'continue')
                            tcPr.append(vm)
                return tbl

    # Для каждой категории клонируем блок прототипа
    anchor = proto_para  # вставляем после последнего добавленного
    try:
        for wc in order:
            slots = groups[wc]
            if not slots:
                continue

            # Динамическая турнирная сетка на ВСЕХ бойцов категории
            members = []
            for _s in slots:
                _a = by_id.get(_s['athlete_id'])
                members.append((
                    (_a['name'] if _a else ('#' + str(_s['athlete_id']))),
                    ((_a.get('region') or '').replace('Калмыкия', '').strip()) if _a else '',
                    str(_a.get('birth_year') or '') if _a else ''
                ))
            # results_map: pair_index -> имя победителя (для подстановки в сетку)
            rm = {}
            for _r in results:
                if _r['weight_class'] == wc and _r.get('winner_id') is not None:
                    _a = by_id.get(_r['winner_id'])
                    rm[_r['pair_index']] = _a['name'] if _a else ''

            if _atoms and _build_bracket:
                # Строим сетку динамически
                new_para = _copy.deepcopy(proto_para)
                _replace_in_el(new_para, {'{WEIGHT_CLASS}': wc, '{SLOT1_NAME}': '', '{SLOT1_CLUB}': '', '{SLOT1_YEAR}': '', '{SLOT2_NAME}': '', '{SLOT2_CLUB}': '', '{SLOT2_YEAR}': '', '{SLOT3_NAME}': '', '{SLOT3_CLUB}': '', '{SLOT3_YEAR}': '', '{SLOT4_NAME}': '', '{SLOT4_CLUB}': '', '{SLOT4_YEAR}': '', '{SLOTS-1-2_COACH}': '', '{SLOTS-3-4_COACH}': '', '{SLOTS_CHAMPION}': ''})
                new_tbl = _build_bracket(members, rm)
            else:
                # Фолбэк: клонируем прототип (4 слота) — старый путь
                slot_ids = beams
                s1 = slot_ids[0] if len(slot_ids) > 0 else None
                s2 = slot_ids[1] if len(slot_ids) > 1 else None
                s3 = slot_ids[2] if len(slot_ids) > 2 else None
                s4 = slot_ids[3] if len(slot_ids) > 3 else None
                mapping = {
                    '{WEIGHT_CLASS}': wc,
                    '{SLOT1_NAME}': _name(s1), '{SLOT1_CLUB}': _club(s1), '{SLOT1_YEAR}': _year(s1),
                    '{SLOT2_NAME}': _name(s2), '{SLOT2_CLUB}': _club(s2), '{SLOT2_YEAR}': _year(s2),
                    '{SLOT3_NAME}': _name(s3), '{SLOT3_CLUB}': _club(s3), '{SLOT3_YEAR}': _year(s3),
                    '{SLOT4_NAME}': _name(s4), '{SLOT4_CLUB}': _club(s4), '{SLOT4_YEAR}': _year(s4),
                    '{SLOTS-1-2_COACH}': '', '{SLOTS-3-4_COACH}': '',
                    '{SLOTS_CHAMPION}': '',
                }
                new_para = _copy.deepcopy(proto_para)
                new_tbl = _copy.deepcopy(proto_tbl)
                _replace_in_el(new_para, mapping)
                _replace_in_el(new_tbl, mapping)
                n_use_slots = max(1, min(len(slot_ids), 4))
                _trs = new_tbl.findall(_qn('w:tr'))
                for _slot_k in range(4, n_use_slots, -1):
                    _r_i = (_slot_k - 1) * 2
                    for _ri in range(_slot_k * 2 - 1, _r_i - 1, -1):
                        if _ri < len(_trs):
                            _tr = _trs[_ri]
                            if _tr.getparent() is not None:
                                _tr.getparent().remove(_tr)
            # «cantSplit» на каждой строке: Word не разрывает категорию между страницами
            for _tr in new_tbl.findall(_qn('w:tr')):
                _trPr = _tr.find(_qn('w:trPr'))
                if _trPr is None:
                    _trPr = _px('<w:trPr xmlns:w="%s"/>' % _WNS)
                    _tr.insert(0, _trPr)
                if _trPr.find(_qn('w:cantSplit')) is None:
                    _trPr.append(_px('<w:cantSplit xmlns:w="%s"/>' % _WNS))
            # Вставляем: [anchor, new_para, new_tbl]
            if anchor is not None:
                anchor.addnext(new_tbl)
                anchor.addnext(new_para)
            else:
                body.insert(list(body).index(proto_tbl) if proto_tbl in list(body) else len(list(body)), new_tbl)
                body.insert(list(body).index(proto_para) if proto_para in list(body) else len(list(body)), new_para)
            anchor = new_tbl
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': f'Ошибка заполнения: {e}'}), 500

    # Удаляем прототип
    if proto_para.getparent() is not None:
        body.remove(proto_para)
    if proto_tbl.getparent() is not None:
        body.remove(proto_tbl)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"protokol_grafikov_{tournament_id}.docx",
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/api/weighin/export', methods=['POST'])
def export_weighin_word():
    """Сохранить таблицу взвешивания в Word (.docx) с текущим списком из UI."""
    if not _HAS_DOCX:
        return jsonify({'error': 'Библиотека Word не установлена (python-docx)'}), 500
    data = request.get_json(silent=True) or {}
    rows = data.get('rows')
    if rows is None and request.form:
        # Поддержка отправки через скрытую HTML-форму (прямое скачивание файла)
        _rv = request.form.get('rows') or ''
        try:
            rows = json.loads(_rv)
        except Exception:
            rows = []
    rows = rows or []
    if not rows:
        return jsonify({'error': 'Нет данных для сохранения'}), 400

    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn as _qn
    from io import BytesIO

    # tournament_id — из payload (JSON или form-data), для шапки/подписей
    tournament_id = data.get('tournament_id')
    if tournament_id is None and request.form:
        try:
            tournament_id = int(request.form.get('tournament_id') or 0) or None
        except Exception:
            tournament_id = None

    _t_dates = ''
    _t_name = ''
    _t_chief_referee = ''
    _t_chief_secretary = ''
    if tournament_id:
        _conn = get_db()
        _tr = row_to_dict(_conn.execute(
            "SELECT * FROM tournaments WHERE id=?", (tournament_id,)).fetchone())
        _conn.close()
        if _tr:
            _t_name = _tr.get('name') or ''
            _t_chief_referee = _tr.get('chief_referee') or ''
            _t_chief_secretary = _tr.get('chief_secretary') or ''
            def _fmtd(s):
                p = (s or '').split('-')
                return f'{p[2]}.{p[1]}.{p[0]}' if len(p) == 3 else (s or '').replace('-', '.')
            _d1 = _fmtd(_tr.get('date_from') or '')
            _d2 = _fmtd(_tr.get('date_to') or _tr.get('date_from') or '')
            _t_dates = f'{_d1} — {_d2}' if _d2 and _d2 != _d1 else _d1

    # Шаблон с гербом и нижним колонтитулом (подписи судей)
    tmpl = _report_template('vzveshivanie_tmp.docx') or _report_template('pkdrk_template.docx')
    doc = Document(tmpl) if tmpl else Document()

    # Заполняем маркеры в колонтитулах шаблона (название, даты, ФИО и регионы судей)
    if tmpl:
        _w_main = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        for _sec in doc.sections:
            for _pn in ('header', 'footer', 'first_page_header', 'first_page_footer'):
                try:
                    _part_el = getattr(_sec, _pn)._element
                    if _part_el is None:
                        continue
                    for _p in _part_el.iter('{%s}p' % _w_main):
                        _full = ''.join(_rt.text or '' for _r in _p.findall('{%s}r' % _w_main)
                                        for _rt in _r.findall('{%s}t' % _w_main))
                        _changed = False
                        for _k, _v in {
                            '{НАЗВАНИЕ_ТУРНИРА}': _t_name,
                            '{ДатыПроведения}': _t_dates,
                            '{ГлавныйСудьяФИО}': _t_chief_referee,
                            '{ГлавныйСудьяРегион}': '',
                            '{ГлавныйСекретарьФИО}': _t_chief_secretary,
                            '{ГлавныйСекретарьРегион}': '',
                        }.items():
                            if _k in _full:
                                _full = _full.replace(_k, str(_v)); _changed = True
                        if _changed:
                            _runs = _p.findall('{%s}r' % _w_main)
                            if _runs:
                                _ts0 = _runs[0].findall('{%s}t' % _w_main)
                                if _ts0:
                                    _ts0[0].text = _full
                                    for _r in _runs[1:]:
                                        for _rt in _r.findall('{%s}t' % _w_main):
                                            _rt.text = ''
                except Exception:
                    pass

        # Маркеры в теле документа (например «СОРЕВНОВАНИЯ {НАЗВАНИЕ_ТУРНИРА} {ДатыПроведения}»)
        try:
            for _p in doc.paragraphs:
                _full = ''.join(_rt.text or '' for _r in _p._p.findall('{%s}r' % _w_main)
                                for _rt in _r.findall('{%s}t' % _w_main))
                _changed = False
                for _k, _v in {
                    '{НАЗВАНИЕ_ТУРНИРА}': _t_name,
                    '{ДатыПроведения}': _t_dates,
                }.items():
                    if _k in _full:
                        _full = _full.replace(_k, str(_v)); _changed = True
                if _changed:
                    _runs = _p._p.findall('{%s}r' % _w_main)
                    if _runs:
                        _ts0 = _runs[0].findall('{%s}t' % _w_main)
                        if _ts0:
                            _ts0[0].text = _full
                            for _r in _runs[1:]:
                                for _rt in _r.findall('{%s}t' % _w_main):
                                    _rt.text = ''
        except Exception:
            pass

    headers = ['№', 'Ф.И.О.', 'Год рожд.', 'Пол', 'Регион', 'Вес (кг)']
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Стиль «Table Grid» есть не во всех шаблонах — пробуем, иначе границы вручную
    try:
        table.style = 'Table Grid'
    except Exception:
        _W2 = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        _tw = table._tbl.find('{%s}tblPr' % _W2) if table._tbl is not None else None
        _border_xml = ('<w:tblBorders xmlns:w="%s">'
                       '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                       '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                       '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                       '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                       '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                       '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                       '</w:tblBorders>') % 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        from docx.oxml import parse_xml
        if _tw is not None:
            _tw.append(parse_xml(_border_xml))

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

    tmpl = _report_template('reglament_tmp.docx')
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

    # --- Заполняем нижний колонтитул (подписи судей) в export_weighin_reglament ---
    try:
        for sec in doc.sections:
            for pn in ('footer', 'first_page_footer'):
                _fpart = getattr(sec, pn)
                for _fp in _fpart._element.iter(_qn('w:p')):
                    _set_para_text(_fp, {
                        '{ГлавныйСудьяФИО}': (t.get('chief_referee') or ''),
                        '{ГлавныйСудьяРегион}': '',
                        '{ГлавныйСекретарьФИО}': (t.get('chief_secretary') or ''),
                        '{ГлавныйСекретарьРегион}': '',
                    })
    except Exception:
        pass

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
