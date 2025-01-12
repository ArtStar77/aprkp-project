# src/ui/header_dialog.py

import tkinter as tk
from tkinter import ttk

class HeaderDialog:
    def __init__(self, parent, title="Добавить заголовок", initial_value=""):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.transient(parent)
        self.top.grab_set()  # Делаем окно модальным
        
        # Центрируем окно
        window_width = 400
        window_height = 150
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.top.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        self.result = None
        self._create_widgets(initial_value)
        
    def _create_widgets(self, initial_value):
        # Создаем основной фрейм с отступами
        main_frame = ttk.Frame(self.top, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Метка
        ttk.Label(main_frame, text="Название группы:").pack(fill=tk.X, pady=(0, 5))
        
        # Поле ввода
        self.entry = ttk.Entry(main_frame)
        self.entry.pack(fill=tk.X, pady=(0, 20))
        self.entry.insert(0, initial_value)
        self.entry.focus_set()
        
        # Кнопки
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(button_frame, text="OK", command=self._on_ok).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Отмена", command=self._on_cancel).pack(side=tk.RIGHT)
        
        # Привязываем клавиши
        self.top.bind("<Return>", lambda e: self._on_ok())
        self.top.bind("<Escape>", lambda e: self._on_cancel())
        
    def _on_ok(self):
        name = self.entry.get().strip()
        if name:
            self.result = name
            self.top.destroy()
        
    def _on_cancel(self):
        self.top.destroy()