# src/ui/product_dialog.py
import tkinter as tk
from tkinter import ttk, messagebox
from decimal import Decimal, InvalidOperation
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ProductDialog:
    def __init__(self, parent, edit_mode=False, initial_values=None):
        self.top = tk.Toplevel(parent)
        self.top.title("Редактировать продукт" if edit_mode else "Добавить продукт")
        self.top.grab_set()
        
        # Центрируем окно
        window_width = 400
        window_height = 300
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.top.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.result = None
        self.edit_mode = edit_mode
        self.initial_values = initial_values or {}
        self._create_widgets()
        
    def _create_widgets(self):
        """Создание элементов интерфейса"""
        main_frame = ttk.Frame(self.top, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        fields = [
            ("Артикул:", "article"),
            ("Наименование:", "name"),
            ("Количество:", "quantity"),
            ("Ед.изм.:", "unit"),
            ("Цена поставщика (руб.):", "price_supplier"),
            ("Наценка (%):", "markup")
        ]
        self.entries = {}
        
        for label_text, field_name in fields:
            frame = ttk.Frame(main_frame)
            frame.pack(fill=tk.X, pady=5)
            
            label = ttk.Label(frame, text=label_text, width=20)
            label.pack(side=tk.LEFT)
            
            entry = ttk.Entry(frame)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.entries[field_name] = entry
            # Заполняем значения, если это режим редактирования
            if self.edit_mode and field_name in self.initial_values:
                value = self.initial_values.get(field_name, "")
                entry.insert(0, str(value))
        # Устанавливаем значения по умолчанию для полей
        if not self.edit_mode:
            self.entries["quantity"].insert(0, "1")
            self.entries["unit"].insert(0, "шт.")
        # Кнопки
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        save_text = "Сохранить" if self.edit_mode else "Добавить"
        ttk.Button(button_frame, text=save_text, command=self.on_save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Отмена", command=self.on_cancel).pack(side=tk.LEFT, padx=5)
        # Привязка клавиш
        self.top.bind("<Return>", lambda e: self.on_save())
        self.top.bind("<Escape>", lambda e: self.on_cancel())
        # Устанавливаем фокус на первое поле
        self.entries["article"].focus_set()

    def on_save(self):
        """Обработка сохранения"""
        try:
            # Проверка наименования
            name = self.entries["name"].get().strip()
            if not name:
                messagebox.showwarning("Внимание", "Наименование продукта не может быть пустым")
                self.entries["name"].focus_set()
                return
            # Проверка количества
            try:
                quantity = int(self.entries["quantity"].get().strip())
                if quantity <= 0:
                    raise ValueError("Количество должно быть больше нуля")
            except ValueError:
                messagebox.showwarning("Внимание", "Количество должно быть целым положительным числом")
                self.entries["quantity"].focus_set()
                return
            # Проверка цены поставщика
            price_str = self.entries["price_supplier"].get().strip().replace(' ', '').replace(',', '.')
            try:
                price = Decimal(price_str)
                if price < 0:
                    raise ValueError("Цена не может быть отрицательной")
            except (InvalidOperation, ValueError):
                messagebox.showwarning("Внимание", "Неверный формат цены")
                self.entries["price_supplier"].focus_set()
                return
            # Проверка наценки
            markup_str = self.entries["markup"].get().strip().replace(' ', '').replace(',', '.')
            try:
                markup = Decimal(markup_str) if markup_str else None
                if markup is not None and markup < 0:
                    raise ValueError("Наценка не может быть отрицательной")
            except (InvalidOperation, ValueError):
                messagebox.showwarning("Внимание", "Неверный формат наценки")
                self.entries["markup"].focus_set()
                return
            article = self.entries["article"].get().strip()
            unit = self.entries["unit"].get().strip() or 'шт.'
            
            # Формируем результат
            self.result = {
                'article': article,
                'name': name,
                'quantity': quantity,
                'unit': unit,
                'supplier_price': price,
                'markup': markup  # Добавляем наценку
            }
            
            self.top.destroy()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить изменения: {str(e)}")
            logger.error(f"Ошибка при сохранении продукта: {str(e)}")

    def on_cancel(self):
        """Обработка отмены"""
        self.top.destroy()