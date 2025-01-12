# src/ui/column_mapping_dialog.py

import tkinter as tk
from tkinter import ttk
from pathlib import Path
from typing import List
from src.models.product import Product

class ColumnMappingDialog:
    def __init__(self, parent, columns):
        self.window = tk.Toplevel(parent)
        self.window.title("Выбор столбцов для импорта")
        self.window.geometry("400x350")
        self.window.resizable(False, False)
        self.window.grab_set()  # Делаем окно модальным
        
        # Центрируем окно
        self.window.transient(parent)
        self.window.geometry("+%d+%d" % (
            parent.winfo_rootx() + parent.winfo_width()/2 - 200,
            parent.winfo_rooty() + parent.winfo_height()/2 - 175
        ))

        self.columns = columns
        self.result = None

        self._create_widgets()
        
    def _create_widgets(self):
        """Создание элементов интерфейса"""
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Описание
        ttk.Label(
            main_frame, 
            text="Укажите соответствие столбцов:",
            wraplength=380
        ).pack(pady=(0, 10))

        # Создаем фрейм для полей выбора
        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill=tk.BOTH, expand=True)

        # Словарь для хранения комбобоксов
        self.field_vars = {}

        # Создаем поля выбора
        fields = {
            'name': 'Наименование *',
            'article': 'Артикул',
            'quantity': 'Количество',
            'unit': 'Единица измерения',
            'price': 'Цена'
        }

        for i, (field, label) in enumerate(fields.items()):
            # Создаем фрейм для строки
            row_frame = ttk.Frame(fields_frame)
            row_frame.pack(fill=tk.X, pady=5)

            # Метка
            ttk.Label(row_frame, text=label, width=20).pack(side=tk.LEFT)

            # Комбобокс
            combobox = ttk.Combobox(row_frame, values=[''] + self.columns, width=30)
            combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            # Автоматический поиск подходящего столбца
            for col in self.columns:
                col_lower = col.lower()
                if field == 'name' and any(x in col_lower for x in ['наименование', 'название', 'товар']):
                    combobox.set(col)
                elif field == 'article' and any(x in col_lower for x in ['артикул', '№', 'номер']):
                    combobox.set(col)
                elif field == 'quantity' and any(x in col_lower for x in ['количество', 'кол-во']):
                    combobox.set(col)
                elif field == 'unit' and any(x in col_lower for x in ['ед.изм', 'единица']):
                    combobox.set(col)
                elif field == 'price' and any(x in col_lower for x in ['цена']):
                    combobox.set(col)

            self.field_vars[field] = combobox

        # Примечание о обязательных полях
        ttk.Label(
            main_frame,
            text="* - обязательное поле",
            font=('', 8)
        ).pack(pady=(5, 10))

        # Кнопки
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(
            buttons_frame,
            text="Импортировать",
            command=self._on_submit
        ).pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            buttons_frame,
            text="Отмена",
            command=self._on_cancel
        ).pack(side=tk.RIGHT, padx=5)

    def _on_submit(self):
        """Обработка подтверждения"""
        # Проверяем наличие обязательных полей
        if not self.field_vars['name'].get():
            tk.messagebox.showerror(
                "Ошибка",
                "Необходимо выбрать столбец с наименованием"
            )
            return

        # Формируем результат
        self.result = {
            field: var.get()
            for field, var in self.field_vars.items()
            if var.get()  # Включаем только выбранные поля
        }
        
        self.window.destroy()

    def _on_cancel(self):
        """Обработка отмены"""
        self.result = None
        self.window.destroy()
    
    def import_with_gui(file_path: Path) -> List[Product]:
        """
        Импорт продуктов из Excel с использованием графического интерфейса для выбора столбцов.

        Args:
            file_path: Путь к Excel файлу.

        Returns:
            Список продуктов.
        """
        try:
            # Автоматическое обнаружение столбцов
            suggested_mapping = ImportService.auto_detect_columns(file_path)
            # Показываем диалог выбора столбцов
            dialog = ColumnMappingDialog(file_path, suggested_mapping)
            dialog.run()
            if dialog.result:
                # Импортируем с выбранным маппингом
                products = ImportService.import_excel(file_path, column_mapping=dialog.result)
                return products
            else:
                logger.info("Пользователь отменил импорт.")
                return []
        except Exception as e:
            logger.error(f"Ошибка при импорте с GUI: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка при импорте с GUI: {str(e)}")
            raise ValueError(f"Ошибка при импорте с GUI: {str(e)}")

