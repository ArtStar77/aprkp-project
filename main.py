import sys
from pathlib import Path

# Добавляем путь к исходникам в PYTHONPATH
project_root = Path(__file__).parent
sys.path.append(str(project_root))

import tkinter as tk
from datetime import datetime
from decimal import Decimal
import locale

from src.ui.main_window import MainWindow
from src.config import settings

def main():
    try:
        # Устанавливаем русскую локаль для форматирования чисел
        try:
            locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
        except:
            try:
                locale.setlocale(locale.LC_ALL, 'Russian_Russia.1251')
            except:
                print("Предупреждение: Не удалось установить русскую локаль")
        
        # Создаем главное окно
        root = tk.Tk()
        root.title(settings.APP_NAME)
        root.geometry(settings.DEFAULT_WINDOW_SIZE)
        
        # Создаем основной интерфейс
        app = MainWindow(root)
        
        # Запускаем главный цикл
        root.mainloop()
        
    except Exception as e:
        print(f"Критическая ошибка: {str(e)}")
        # Показываем диалог с ошибкой
        try:
            import tkinter.messagebox as messagebox
            messagebox.showerror("Критическая ошибка", str(e))
        except:
            pass

if __name__ == "__main__":
    main()