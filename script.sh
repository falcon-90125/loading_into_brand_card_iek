# Запустить в терминале командой ./script.sh

# Создание виртуального окружения
python -m venv env

# Активация виртуального окружения
env\Scripts\activate

# Установка необходимых зависимостей
pip install -r requirements.txt

# Сохранить все пакеты из виртуального окружения в файл requirements.txt:
# pip freeze > requirements.txt

# Создать новый .ipynb файл
# Нажмите Ctrl+Shift+P → введите Create: New Jupyter Notebook → Enter.
# VS Code создаст файл с расширением .ipynb.
# Если при запуске появляется ошибка (Jupyter not found), установите его:
# pip install jupyter