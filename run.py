#!/usr/bin/env python3
"""
🥊 Запуск приложения «Соревнования по боксу»

Запускает локальный веб-сервер и открывает приложение в браузере.
Работает полностью локально, без интернета (кроме загрузки XLSX библиотеки).
"""

import subprocess
import sys
import os

def main():
    # Устанавливаем зависимости если нужно
    req_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'requirements.txt')
    try:
        import flask
    except ImportError:
        print('Устанавливаю зависимости...')
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', req_file])

    # Запускаем сервер Python
    server_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'server.py')

    if not os.path.exists(server_script):
        print(f'Ошибка: не найден {server_script}')
        input('Нажмите Enter для выхода...')
        sys.exit(1)

    # Запускаем сервер в отдельном окне
    if sys.platform == 'win32':
        cmd = f'start "Boxing App" /MIN python "{server_script}"'
        os.system(cmd)
    else:
        subprocess.Popen(
            [sys.executable, server_script],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )

    print('=' * 60)
    print('  🥊 СОРЕВНОВАНИЯ ПО БОКСУ')
    print('=' * 60)
    print('\n  ✅ Сервер запущен!')
    print('  Сервер работает в отдельном окне.')
    print('  Закройте окно сервера или нажмите Ctrl+C для остановки.\n')

    try:
        input('  Нажмите Enter для выхода...')
    except KeyboardInterrupt:
        pass

    print('\n  Программа завершена.')

if __name__ == '__main__':
    main()