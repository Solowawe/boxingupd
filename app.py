"""
Точка входа для Layero (python_web preset).
Импортирует Flask-приложение из server.py.
"""

from database import init_db
from server import app

# Инициализация БД при старте контейнера
init_db()

# gunicorn ищет переменную `app`
application = app

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=False, threaded=True)
