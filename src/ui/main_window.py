# src/ui/main_window.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from pathlib import Path
from typing import Callable, Optional, List
from decimal import Decimal, InvalidOperation as DecimalException
from datetime import datetime, date
import locale
import logging
import pyperclip
import os
import platform
import subprocess
from tkcalendar import DateEntry  # Убедитесь, что tkcalendar установлен
from src.ui.column_mapping_dialog import ColumnMappingDialog
from src.ui.header_dialog import HeaderDialog  # Добавьте этот импорт к остальным импортам
from src.ui.product_dialog import ProductDialog
from src.ui.product_tree import ProductTreeView
from src.services.price_calculator import PriceCalculator
from src.services.file_service import FileService
from src.services.import_service import ImportService
from src.services.export_service import ExportService
from src.models.commercial_offer import CommercialOffer
from src.models.product import Product
from src.utils.formatters import Formatters
from src.config import settings

logger = logging.getLogger(__name__)

def open_file(path: Path):
    """Открытие файла с использованием стандартного приложения"""
    try:
        if platform.system() == 'Windows':
            os.startfile(path)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.call(['open', path])
        else:  # Linux и другие
            subprocess.call(['xdg-open', path])
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть файл: {str(e)}")

class MainMenu:
    def __init__(self, parent, callbacks):
        """
        Инициализация главного меню
        Args:
            parent: Родительское окно
            callbacks: Словарь с функциями обратного вызова
        """
        self.parent = parent
        self.callbacks = callbacks
        # Создаем главное меню
        self.menubar = tk.Menu(self.parent)
        self.parent.config(menu=self.menubar)
        # Создаем подменю
        self._create_file_menu()
        self._create_edit_menu()
        self._create_import_export_menu()
        self._create_tools_menu()
        self._create_help_menu()
        # Привязываем глобальные горячие клавиши
        self._bind_global_hotkeys()

    def _create_file_menu(self):
        """Создание меню Файл"""
        self.file_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Файл", menu=self.file_menu)
        self.file_menu.add_command(
            label="Новый", 
            command=self.callbacks['new_offer'], 
            accelerator="Ctrl+N"
        )
        self.file_menu.add_command(
            label="Открыть", 
            command=self.callbacks['load_offer'],
            accelerator="Ctrl+O"
        )
        self.file_menu.add_command(
            label="Сохранить как", 
            command=self.callbacks['save_offer_as'], 
            accelerator="Ctrl+Shift+S"
        )
        self.file_menu.add_command(
            label="Сохранить", 
            command=self.callbacks['save_offer'],
            accelerator="Ctrl+S"
        )
        self.file_menu.add_separator()
        self.file_menu.add_command(
            label="Выход", 
            command=self.parent.quit,
            accelerator="Alt+F4"
        )

    def _create_edit_menu(self):
        """Создание меню Правка"""
        self.edit_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Правка", menu=self.edit_menu)
        self.edit_menu.add_command(
            label="Добавить продукт", 
            command=self.callbacks['add_product'],
            accelerator="Ctrl+A"
        )
        self.edit_menu.add_command(
            label="Редактировать", 
            command=self.callbacks['edit_product'],
            accelerator="F2"
        )
        self.edit_menu.add_command(
            label="Удалить", 
            command=self.callbacks['delete_product'],
            accelerator="Delete"
        )
        self.edit_menu.add_separator()
        self.edit_menu.add_command(
            label="Копировать", 
            command=lambda: self.callbacks['copy_selected'](),
            accelerator="Ctrl+C"
        )
        self.edit_menu.add_command(
            label="Вставить", 
            command=lambda: self.callbacks['paste_selected'](),
            accelerator="Ctrl+V"
        )
        self.edit_menu.add_separator()
        self.edit_menu.add_command(
            label="Добавить заголовок", 
            command=self.callbacks['add_group_header'],
            accelerator="Ctrl+H"
        )
        self.edit_menu.add_separator()
        self.edit_menu.add_command(
            label="Переместить вверх", 
            command=lambda: self.callbacks['move_selected']('up'),
            accelerator="Alt+↑"
        )
        self.edit_menu.add_command(
            label="Переместить вниз", 
            command=lambda: self.callbacks['move_selected']('down'),
            accelerator="Alt+↓"
        )

    def _create_import_export_menu(self):
        """Создание меню Импорт/Экспорт"""
        self.export_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Импорт/Экспорт", menu=self.export_menu)
        self.export_menu.add_command(
            label="Импорт", 
            command=self.callbacks['import_data'],
            accelerator="Ctrl+I"
        )
        self.export_menu.add_separator()
        self.export_menu.add_command(
            label="Экспорт Word", 
            command=self.callbacks['export_word'],
            accelerator="Ctrl+W"
        )
        self.export_menu.add_command(
            label="Экспорт Excel", 
            command=self.callbacks['export_excel'],
            accelerator="Ctrl+X"
        )
        self.export_menu.add_command(
            label="Экспорт PDF", 
            command=self.callbacks['export_pdf'],
            accelerator="Ctrl+P"
        )

    def _create_tools_menu(self):
        """Создание меню Инструменты"""
        self.tools_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Инструменты", menu=self.tools_menu)
        self.tools_menu.add_command(
            label="Очистить все", 
            command=self.callbacks['clear_all_data']
        )
        self.tools_menu.add_command(
            label="Обновить все", 
            command=self.callbacks['update_all_data'],
            accelerator="F5"
        )
        self.tools_menu.add_command(
            label="Сбросить цены", 
            command=self.callbacks['reset_prices'],
            accelerator="Ctrl+R"
        )
        # Добавляем разделитель и пункт для работы с шаблоном
        self.tools_menu.add_separator()
        self.tools_menu.add_command(
            label="Открыть шаблон Word",
            command=self._open_word_template
        )

    def _open_word_template(self):
        """Открытие шаблона Word для редактирования"""
        template_path = settings.TEMPLATES_DIR / "KP_shablon.docx"
        if template_path.exists():
            open_file(template_path)
        else:
            messagebox.showerror(
                "Ошибка", 
                f"Шаблон не найден: {template_path}\n"
                f"Пожалуйста, убедитесь, что файл 'KP_shablon.docx' находится в папке templates"
            )

    def _create_help_menu(self):
        """Создание меню Справка"""
        self.help_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Справка", menu=self.help_menu)
        self.help_menu.add_command(
            label="Горячие клавиши", 
            command=self._show_hotkeys
        )
        self.help_menu.add_separator()
        self.help_menu.add_command(
            label="О программе", 
            command=self._show_about
        )

    def _bind_global_hotkeys(self):
        """Привязка глобальных горячих клавиш"""
        # Файл
        self.parent.bind('<Control-n>', lambda e: self.callbacks['new_offer']())
        self.parent.bind('<Control-o>', lambda e: self.callbacks['load_offer']())
        self.parent.bind('<Control-s>', lambda e: self.callbacks['save_offer']())
        self.parent.bind('<Control-Shift-s>', lambda e: self.callbacks['save_offer_as']())
        # Правка
        self.parent.bind('<Control-a>', lambda e: self.callbacks['add_product']())
        self.parent.bind('<F2>', lambda e: self.callbacks['edit_product']())
        self.parent.bind('<Delete>', lambda e: self.callbacks['delete_product']())
        self.parent.bind('<Control-h>', lambda e: self.callbacks['add_group_header']())
        # Копирование/Вставка
        self.parent.bind('<Control-c>', lambda e: self.callbacks['copy_selected']())
        self.parent.bind('<Control-v>', lambda e: self.callbacks['paste_selected']())
        # Перемещение
        self.parent.bind('<Alt-Up>', lambda e: self.callbacks['move_selected']('up'))
        self.parent.bind('<Alt-Down>', lambda e: self.callbacks['move_selected']('down'))
        # Импорт/Экспорт
        self.parent.bind('<Control-i>', lambda e: self.callbacks['import_data']())
        self.parent.bind('<Control-p>', lambda e: self.callbacks['export_pdf']())
        self.parent.bind('<Control-x>', lambda e: self.callbacks['export_excel']())
        self.parent.bind('<Control-w>', lambda e: self.callbacks['export_word']())
        # Инструменты
        self.parent.bind('<F5>', lambda e: self.callbacks['update_all_data']())
        self.parent.bind('<Control-r>', lambda e: self.callbacks['reset_prices']())

    def _show_hotkeys(self):
        """Показ окна с горячими клавишами"""
        hotkeys_window = tk.Toplevel(self.parent)
        hotkeys_window.title("Горячие клавиши")
        hotkeys_window.geometry("400x500")
        # Создаем текст с описанием горячих клавиш
        text = ttk.Label(hotkeys_window, text="""
Файл:
Ctrl+N - Новое КП
Ctrl+O - Открыть
Ctrl+S - Сохранить
Alt+F4 - Выход
Правка:
F2 - Редактировать
Ctrl+A - Добавить продукт
Delete - Удалить
Ctrl+H - Добавить заголовок
Ctrl+C - Копировать
Ctrl+V - Вставить
Alt+↑ - Переместить вверх
Alt+↓ - Переместить вниз
Импорт/Экспорт:
Ctrl+I - Импорт
Ctrl+W - Экспорт Word
Ctrl+X - Экспорт Excel
Ctrl+P - Экспорт PDF
Инструменты:
F5 - Обновить все
Ctrl+R - Сбросить цены
        """, justify=tk.LEFT)
        text.pack(padx=20, pady=20)

    def _show_about(self):
        """Показ окна о программе"""
        about_window = tk.Toplevel(self.parent)
        about_window.title("О программе")
        about_window.geometry("300x200")
        ttk.Label(about_window, text="""
Аэропро
Создание коммерческих предложений
Версия: 2.0.0
        """).pack(padx=20, pady=20)

class MainWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Аэропро. Создание КП")
        self.root.geometry("1200x800")
        # Инициализация сервисов
        self.calculator = PriceCalculator()
        self.file_service = FileService(Path("data"))
        # Убедитесь, что директории существуют
        self.file_service.base_path.mkdir(parents=True, exist_ok=True)
        (self.file_service.base_path / "offers").mkdir(parents=True, exist_ok=True)
        (self.file_service.base_path / "templates").mkdir(parents=True, exist_ok=True)
        (self.file_service.base_path / "exports").mkdir(parents=True, exist_ok=True)
        # Состояние приложения
        self.current_offer: Optional[CommercialOffer] = None
        # Создаем основные фреймы
        self.toolbar = ttk.Frame(self.root)
        self.toolbar.pack(fill=tk.X, padx=5, pady=5)
        self._init_ui()
        self._bind_events()
        self.reset_bindings()

    def _init_ui(self):
        """Инициализация пользовательского интерфейса"""
        # Создаем главное меню
        callbacks = {
            'new_offer': self._new_offer,
            'load_offer': self._load_offer,
            'save_offer': self._save_offer,
            'save_offer_as': self._save_offer_as,
            'import_data': self._import_data,
            'add_product': self._add_product,
            'edit_product': self._edit_product,
            'delete_product': self._delete_product,
            'export_word': self._export_word,
            'export_excel': self._export_excel,
            'export_pdf': self._export_pdf,
            'clear_all_data': self._clear_all_data,
            'update_all_data': self._update_all_data,
            'reset_prices': self._reset_prices,
            'add_group_header': self._add_group_header,
            'copy_selected': self._handle_copy,
            'paste_selected': self._handle_paste,
            'move_selected': self._move_selected
        }
        # Создаем основные фреймы
        self.params_frame = ttk.LabelFrame(self.root, text="Параметры КП")
        self.params_frame.pack(fill=tk.X, padx=5, pady=5)
        self.content = ttk.Frame(self.root)
        self.content.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.status_bar = ttk.Frame(self.root)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=5)
        # Создаем меню и компоненты
        self.main_menu = MainMenu(self.root, callbacks)
        self._create_params()
        self._create_product_tree()
        self._create_status_bar()

    def _handle_copy(self, event=None):
        """Обработчик копирования"""
        try:
            self.product_tree.copy_selected()
            self.status_label.config(text="Скопировано")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при копировании: {str(e)}")

    def _handle_paste(self, event=None):
        """Обработчик вставки"""
        try:
            self.paste_selected()
            self.status_label.config(text="Вставлено")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при вставке: {str(e)}")

    def _edit_product(self, event=None):
        """Редактирование выбранного продукта"""
        if not self.current_offer:
            return
        # Получаем выбранный элемент
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите продукт для редактирования")
            return
        item_id = selected[0]
        index = self.product_tree.get_item_index(item_id)
        if index is None:
            return
        product = self.current_offer.products[index]
        # Не редактируем заголовки
        if product.is_header:
            return
        try:
            # Создаем диалог редактирования с текущими значениями
            dialog = ProductDialog(
                self.root, 
                edit_mode=True,
                initial_values={
                    'article': product.article,
                    'name': product.name,
                    'quantity': product.quantity,
                    'unit': product.unit,
                    'supplier_price': product.supplier_price,
                    'markup': product.markup  # Добавляем наценку
                }
            )
            self.root.wait_window(dialog.top)
            if dialog.result:
                # Обновляем данные продукта
                product.article = dialog.result['article']
                product.name = dialog.result['name']
                product.quantity = dialog.result['quantity']
                product.unit = dialog.result['unit']
                product.supplier_price = dialog.result['supplier_price']
                product.markup = dialog.result['markup']  # Обновляем наценку
                # Пересчитываем цены
                self.calculator.calculate_client_price(
                    product,
                    self.current_offer.discount_from_supplier,
                    self.current_offer.markup_for_client
                )
                # Обновляем отображение
                self.product_tree.update_product(item_id, product)
                self._update_totals()
                self.status_label.config(text=f"Продукт обновлен: {product.name}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось отредактировать продукт: {str(e)}")

    def _create_params(self):
        """Создание панели параметров"""
        # Параметры в две строки
        params_frame1 = ttk.Frame(self.params_frame)
        params_frame1.pack(fill=tk.X, padx=5, pady=5)
        params_frame2 = ttk.Frame(self.params_frame)
        params_frame2.pack(fill=tk.X, padx=5, pady=5)
        # Первая строка
        ttk.Label(params_frame1, text="Номер КП:").pack(side=tk.LEFT, padx=5)
        self.number_entry = ttk.Entry(params_frame1, width=15)
        self.number_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(params_frame1, text="Дата КП:").pack(side=tk.LEFT, padx=5)
        self.date_entry = DateEntry(
            params_frame1,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd.mm.yyyy'
        )
        self.date_entry.pack(side=tk.LEFT, padx=5)
        # Вторая строка
        ttk.Label(params_frame2, text="Скидка от поставщика (%):").pack(side=tk.LEFT, padx=5)
        self.discount_entry = ttk.Entry(params_frame2, width=10)
        self.discount_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(params_frame2, text="Наценка для клиента (%):").pack(side=tk.LEFT, padx=5)
        self.markup_entry = ttk.Entry(params_frame2, width=10)
        self.markup_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(params_frame2, text="НДС (%):").pack(side=tk.LEFT, padx=5)
        self.vat_entry = ttk.Entry(params_frame2, width=10)
        self.vat_entry.pack(side=tk.LEFT, padx=5)
        # Создаем скрытые переменные для хранения значений
        self.delivery_terms_entry = ttk.Entry(self.root)
        self.self_pickup_entry = ttk.Entry(self.root)
        self.warranty_entry = ttk.Entry(self.root)
        self.delivery_time_entry = ttk.Entry(self.root)

    def _create_product_tree(self):
        """Создание таблицы продуктов"""
        tree_frame = ttk.Frame(self.content)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        # Создаем скроллбары
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        # Создаем дерево с привязкой к скроллбарам
        self.product_tree = ProductTreeView(
            tree_frame,
            on_change=self._on_product_change,
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set
        )
        # Настраиваем скроллбары
        vsb.config(command=self.product_tree.yview)
        hsb.config(command=self.product_tree.xview)
        # Размещаем элементы
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.product_tree.pack(fill=tk.BOTH, expand=True)

    def _on_product_change(self, item_id: str, column: str, value: str):
        """Обработчик изменения значения в дереве"""
        try:
            # Получаем индекс элемента в списке продуктов
            index = self.product_tree.get_item_index(item_id)
            if index is None:
                return
            # Получаем продукт
            product = self.current_offer.products[index]
            # Обновляем значение в зависимости от колонки
            if column == 'article':
                product.article = value
            elif column == 'name':
                product.name = value
            elif column == 'quantity':
                product.quantity = int(value)
                product.calculate_total()
            elif column == 'unit':
                product.unit = value
            elif column == 'price_supplier':
                product.supplier_price = Decimal(value)
                # Пересчитываем цены
                self.calculator.calculate_client_price(
                    product,
                    self.current_offer.discount_from_supplier,
                    self.current_offer.markup_for_client
                )
            elif column == 'markup':
                product.markup = Decimal(value)
                # Пересчитываем цены
                self.calculator.calculate_client_price(
                    product,
                    self.current_offer.discount_from_supplier,
                    self.current_offer.markup_for_client
                )
            # Обновляем отображение продукта в дереве
            self.product_tree.update_product(item_id, product)
            # Обновляем итоги
            self._update_totals()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить значение: {str(e)}")

    def _create_status_bar(self):
        """Создание строки состояния"""
        self.status_label = ttk.Label(
            self.status_bar,
            text="Готов к работе"
        )
        self.status_label.pack(side=tk.LEFT, padx=5)
        self.total_label = ttk.Label(self.status_bar, text="")
        self.total_label.pack(side=tk.RIGHT, padx=5)

    def _bind_events(self):
        """Привязка обработчиков событий"""
        # Горячие клавиши для основных операций
        self.root.bind('<Control-n>', self._new_offer)      # Ctrl+N - Новое КП
        self.root.bind('<Control-o>', self._load_offer)     # Ctrl+O - Открыть
        self.root.bind('<Control-s>', self._save_offer)     # Ctrl+S - Сохранить
        self.root.bind('<Control-Shift-s>', self._save_offer_as) # Ctrl+Shift+S - Сохранить как
        self.root.bind('<Control-i>', self._import_data)    # Ctrl+I - Импорт
        self.root.bind('<F2>', self._edit_product)          # F2 - Редактировать
        self.root.bind('<Control-p>', self._export_pdf)     # Ctrl+P - Экспорт PDF
        self.root.bind('<Control-x>', self._export_excel)   # Ctrl+X - Экспорт Excel
        self.root.bind('<Control-w>', self._export_word)    # Ctrl+W - Экспорт Word
        # Работа с продуктами
        self.root.bind('<Control-a>', self._add_product)    # Ctrl+A - Добавить
        self.root.bind('<Delete>', self._delete_product)    # Delete - Удалить
        self.root.bind('<Control-h>', self._add_group_header)  # Ctrl+H - Добавить заголовок
        # Буфер обмена
        self.root.bind('<Control-c>', lambda e: self.product_tree.copy_selected())
        self.root.bind('<Control-v>', lambda e: self.product_tree.paste_selected())
        # Обновление данных
        self.root.bind('<F5>', self._update_all_data)       # F5 - Обновить все
        self.root.bind('<Control-r>', self._reset_prices)   # Ctrl+R - Сбросить цены
        # Добавляем новые привязки для работы с заголовками
        self.root.bind('<F2>', self._edit_selected)
        # Навигация и перемещение
        self.root.bind('<Alt-Up>', lambda e: self._move_selected('up'))      # Alt+↑ - Переместить вверх
        self.root.bind('<Alt-Down>', lambda e: self._move_selected('down'))  # Alt+↓ - Переместить вниз
        # Поля ввода
        self.discount_entry.bind('<KeyRelease>', self._on_discount_change)
        self.markup_entry.bind('<KeyRelease>', self._on_markup_change)
        self.vat_entry.bind('<KeyRelease>', self._on_vat_change)
        # Дерево продуктов
        self.product_tree.bind('<Double-1>', self._on_double_click)
        self.product_tree.bind('<Button-3>', self._show_context_menu)
        # Добавляем новые привязки для событий дерева
        self.product_tree.bind('<Insert>', lambda e: self._add_product())
        self.product_tree.bind('<Control-h>', lambda e: self._add_group_header())
        self.product_tree.bind('<F2>', lambda e: self._edit_product())
        self.product_tree.bind('<Delete>', lambda e: self._delete_product())
        self.product_tree.bind('<F5>', lambda e: self._update_all_data())
        # Привязываем новые события от дерева
        self.product_tree.bind('<F2>', lambda e: self._edit_header())
        self.product_tree.bind('<Delete>', lambda e: self._delete_selected())

    def _on_double_click(self, event):
        """Обработка двойного клика по продукту"""
        item = self.product_tree.identify_row(event.y)
        if not item:
            return
        # Не редактируем заголовки
        if 'header' in self.product_tree.item(item, 'tags'):
            return
        # Запускаем редактирование
        self._edit_product()
        return 'break'  # Предотвращаем дальнейшую обработку события

    def _show_context_menu(self, event):
        """Показ контекстного меню"""
        menu = tk.Menu(self.root, tearoff=0)
        # Добавление пунктов меню
        menu.add_command(label="Добавить продукт (Ctrl+A)", command=self._add_product)
        menu.add_command(label="Добавить заголовок (Ctrl+H)", command=self._add_group_header)
        menu.add_separator()
        # Пункты для работы с выделенными строками
        if self.product_tree.selection():
            menu.add_command(label="Копировать (Ctrl+C)", command=self.product_tree.copy_selected)
            menu.add_command(label="Вставить (Ctrl+V)", command=self.product_tree.paste_selected)
            menu.add_command(label="Удалить (Delete)", command=self._delete_product)
            menu.add_separator()
            menu.add_command(label="Редактировать", command=lambda: self._edit_selected())
            menu.add_command(label="Переместить вверх (Alt+↑)", 
                             command=lambda: self._move_selected('up'))
            menu.add_command(label="Переместить вниз (Alt+↓)", 
                             command=lambda: self._move_selected('down'))
        menu.post(event.x_root, event.y_root)

    def _edit_selected(self, event=None):
        """Редактирование выбранного элемента (продукта или заголовка)"""
        selected = self.product_tree.selection()
        if not selected:
            return
        item_id = selected[0]
        if 'header' in self.product_tree.item(item_id, 'tags'):
            self._edit_header()
        else:
            self._edit_product() 

    def _edit_header(self):
        """Редактирование заголовка группы"""
        selected = self.product_tree.selection()
        if not selected:
            return
        item_id = selected[0]
        if not 'header' in self.product_tree.item(item_id, 'tags'):
            return
        # Получаем текущий текст заголовка
        current_name = self.product_tree.item(item_id, 'values')[1]
        # Показываем диалог редактирования
        dialog = HeaderDialog(
            self.root,
            title="Редактировать заголовок",
            initial_value=current_name
        )
        self.root.wait_window(dialog.top)
        if dialog.result:
            # Получаем индекс элемента
            index = self.product_tree.get_item_index(item_id)
            if index is not None:
                # Обновляем заголовок в модели данных
                self.current_offer.products[index].name = dialog.result
                # Обновляем отображение
                self.product_tree.set(item_id, 'name', dialog.result)
                self.status_label.config(text=f"Заголовок изменен: {dialog.result}")

    def _on_drag_start(self, event):
        """Начало перетаскивания"""
        self._drag_data = {'item': None, 'x': 0, 'y': 0}
        # Определяем элемент под курсором
        item = self.product_tree.identify_row(event.y)
        if item:
            self._drag_data['item'] = item
            self._drag_data['x'] = event.x
            self._drag_data['y'] = event.y

    def _on_drag_motion(self, event):
        """Обработка перемещения при перетаскивании"""
        if self._drag_data['item']:
            # Определяем новую позицию
            target = self.product_tree.identify_row(event.y)
            if target and target != self._drag_data['item']:
                # Перемещаем элемент
                self.product_tree.move(self._drag_data['item'], 
                                       self.product_tree.parent(target),
                                       target)
                # Обновляем порядок в списке продуктов
                self._update_products_order()

    def _on_drag_release(self, event):
        """Завершение перетаскивания"""
        self._drag_data = {'item': None, 'x': 0, 'y': 0}

    def _update_products_order(self):
        """Обновление порядка продуктов после перетаскивания"""
        if not self.current_offer:
            return
        # Получаем новый порядок из дерева
        new_order = []
        for item_id in self.product_tree.get_children():
            index = self.product_tree.get_item_index(item_id)
            if index is not None:
                new_order.append(self.current_offer.products[index])
        # Обновляем список продуктов
        self.current_offer.products = new_order
        self._update_totals()

    def _move_selected(self, direction: str):
        """Перемещение выбранных строк вверх/вниз"""
        selected = self.product_tree.selection()
        if not selected:
            return
        for item_id in selected:
            index = self.product_tree.get_item_index(item_id)
            if index is None:
                continue
            if direction == 'up' and index > 0:
                # Перемещаем вверх
                self.current_offer.products.insert(index - 1, 
                                                   self.current_offer.products.pop(index))
            elif direction == 'down' and index < len(self.current_offer.products) - 1:
                # Перемещаем вниз
                self.current_offer.products.insert(index + 1, 
                                                   self.current_offer.products.pop(index))
        # Обновляем отображение
        self._update_product_tree()
        # Восстанавливаем выделение
        for item_id in selected:
            self.product_tree.selection_add(item_id)

    def _update_ui(self):
        """Обновление интерфейса при изменении данных"""
        try:
            if not self.current_offer:
                return
            # Обновляем поля параметров
            self.number_entry.delete(0, tk.END)
            self.number_entry.insert(0, self.current_offer.number)
            self.date_entry.set_date(self.current_offer.date.date() if self.current_offer.date else date.today())
            self.discount_entry.delete(0, tk.END)
            self.discount_entry.insert(0, str(self.current_offer.discount_from_supplier))
            self.markup_entry.delete(0, tk.END)
            self.markup_entry.insert(0, str(self.current_offer.markup_for_client))
            self.vat_entry.delete(0, tk.END)
            self.vat_entry.insert(0, str(self.current_offer.vat))
            self.delivery_terms_entry.delete(0, tk.END)
            self.delivery_terms_entry.insert(0, self.current_offer.delivery_terms)
            self.self_pickup_entry.delete(0, tk.END)
            self.self_pickup_entry.insert(0, self.current_offer.self_pickup_warehouse)
            self.warranty_entry.delete(0, tk.END)
            self.warranty_entry.insert(0, self.current_offer.warranty)
            self.delivery_time_entry.delete(0, tk.END)
            self.delivery_time_entry.insert(0, self.current_offer.delivery_time)
            # Обновляем дерево продуктов
            self._update_product_tree()
            # Обновляем итоги
            self._update_totals()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при обновлении интерфейса: {str(e)}")

    def _update_product_tree(self):
        """Обновление отображения продуктов в дереве"""
        try:
            # Сохраняем текущее выделение
            selected = self.product_tree.selection()
            # Очищаем дерево
            for item in self.product_tree.get_children():
                self.product_tree.delete(item)
            # Добавляем продукты заново
            for product in self.current_offer.products:
                if product.is_header:
                    # Для заголовков групп
                    self.product_tree.insert(
                        '',
                        'end',
                        values=('', product.name, '', '', '', '', ''),
                        tags=('header',)
                    )
                else:
                    # Для обычных продуктов
                    self.product_tree.add_product(product)
            # Настраиваем стиль заголовков
            self.product_tree.tag_configure('header', font=('Helvetica', 10, 'bold'))
            # Восстанавливаем выделение
            if selected:
                for item in selected:
                    try:
                        self.product_tree.selection_add(item)
                    except:
                        pass
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при обновлении списка продуктов: {str(e)}")

    def _update_totals(self):
        """Обновление итоговых значений"""
        if not self.current_offer:
            return
        try:
            # Получаем итоги
            totals = self.calculator.calculate_totals(
                self.current_offer.products,
                self.current_offer.discount_from_supplier,
                self.current_offer.markup_for_client,
                self.current_offer.vat
            )
            # Форматируем строку итогов
            if totals['total_amount'] != Decimal('0'):
                margin_percentage = (totals['margin'] / totals['total_amount'] * Decimal('100')).quantize(Decimal('0.1'))
            else:
                margin_percentage = Decimal('0.0')
            total_text = (
                f"Итого: {Formatters.format_currency(totals['total_amount'])} | "
                f"НДС: {Formatters.format_currency(totals['vat_amount'])} | "
                f"Маржа: {Formatters.format_currency(totals['margin'])} "
                f"({margin_percentage}%) | "
                f"Цена клиента у Поставщика: {Formatters.format_currency(totals['client_savings'])} ({totals['client_savings_percentage']}%)"
            )
            # Обновляем метку с итогами
            self.total_label.config(text=total_text)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при расчете итогов: {str(e)}")

    def _new_offer(self, event=None):
        """Создание нового коммерческого предложения"""
        try:
            # Проверяем, существует ли уже коммерческое предложение
            if self.current_offer and not messagebox.askyesno(
                "Подтверждение",
                "Создать новое КП? Несохраненные изменения будут потеряны."
            ):
                return

            # Устанавливаем значения по умолчанию
            self.vat_entry.delete(0, tk.END)
            self.vat_entry.insert(0, "20")
            self.discount_entry.delete(0, tk.END)
            self.discount_entry.insert(0, "0")
            self.markup_entry.delete(0, tk.END)
            self.markup_entry.insert(0, "0")

            # Создаём новое коммерческое предложение
            self.current_offer = CommercialOffer(
                number=self._generate_offer_number(),
                date=datetime.now(),
                discount_from_supplier=Decimal('0'),
                markup_for_client=Decimal('0'),
                vat=Decimal('20'),
                delivery_terms="",
                self_pickup_warehouse="",
                warranty="",
                delivery_time="",
                include_delivery=False,
                delivery_cost=Decimal('0'),
                products=[]
            )

            # Обновляем интерфейс
            self._update_ui()
            self.status_label.config(text="Создано новое КП")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать новое КП: {str(e)}")
        
    def _generate_offer_number(self) -> str:
        """Генерация номера КП в формате ТКП00000"""
        # Создаем папку для хранения счетчика, если её нет
        counter_dir = self.file_service.base_path / "system"
        counter_dir.mkdir(parents=True, exist_ok=True)
        counter_file = counter_dir / "counter.txt"
        try:
            # Читаем текущий счетчик
            if counter_file.exists():
                with open(counter_file, 'r') as f:
                    counter = int(f.read().strip())
            else:
                # Если файла нет, начинаем с 60 (текущее значение)
                counter = 60
            # Увеличиваем счетчик
            counter += 1
            # Сохраняем новое значение счетчика
            with open(counter_file, 'w') as f:
                f.write(str(counter))
            # Форматируем номер: ТКП + пятизначное число с ведущими нулями
            return f"ТКП{counter:05d}"
        except Exception as e:
            # В случае ошибки возвращаем номер со временем как запасной вариант
            from datetime import datetime
            return f"ТКП{datetime.now().strftime('%Y%m%d-%H%M%S')}"

    def _save_offer(self, event=None):
        """Сохранение текущего КП"""
        if not self.current_offer:
            messagebox.showwarning("Предупреждение", "Нет активного КП для сохранения")
            return
        try:
            # Обновляем данные КП из полей ввода
            self._update_offer_from_ui()
            # Проверяем существование директории
            self.file_service.offers_path.mkdir(parents=True, exist_ok=True)
            # Сохраняем КП
            filename = f"КП_{self.current_offer.number}"
            self.file_service.save_offer(
                self.current_offer,
                filename
            )
            # Обновляем интерфейс после сохранения
            self._refresh_interface()
            self.status_label.config(text=f"КП {self.current_offer.number} сохранено")
            messagebox.showinfo("Успех", "Коммерческое предложение сохранено")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить КП: {str(e)}")

    def _save_offer_as(self, event=None):
        """Сохранение текущего КП в пользовательском месте"""
        if not self.current_offer:
            messagebox.showwarning("Предупреждение", "Нет активного КП для сохранения")
            return
        try:
            filename = filedialog.asksaveasfilename(
                title="Сохранить как",
                defaultextension=".json",
                filetypes=[("JSON Files", "*.json")],
                initialdir=self.file_service.offers_path,
                initialfile=f"КП_{self.current_offer.number}.json"
            )
            if not filename:
                return  # Пользователь отменил операцию
            self.file_service.save_offer(self.current_offer, Path(filename).stem, Path(filename))
            self.status_label.config(text=f"КП сохранено в {Path(filename).name}")
            messagebox.showinfo("Успех", "Коммерческое предложение успешно сохранено.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить КП: {str(e)}")

    def _load_offer(self, event=None):
        """Загрузка КП"""
        try:
            if self.current_offer and not messagebox.askyesno(
                "Подтверждение",
                "Загрузить другое КП? Несохраненные изменения будут потеряны."
            ):
                return
            filename = filedialog.askopenfilename(
                initialdir=self.file_service.offers_path,
                title="Выберите файл КП",
                filetypes=[("JSON files", "*.json")]
            )
            if not filename:
                return
            try:
                offer = self.file_service.load_offer(Path(filename).stem)
                if offer:
                    self.current_offer = offer
                    self._update_ui()
                    self._refresh_interface()
                    self.status_label.config(text=f"Загружено КП {offer.number}")
                else:
                    messagebox.showerror("Ошибка", "Не удалось загрузить КП")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при загрузке КП: {str(e)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при выборе файла: {str(e)}")

    def _add_product(self, event=None):
        """Добавление нового продукта"""
        if not self.current_offer:
            messagebox.showwarning("Предупреждение", "Сначала создайте новое КП")
            return
        try:
            dialog = ProductDialog(self.root)
            self.root.wait_window(dialog.top)
            if dialog.result:
                # Создаем продукт с явным указанием что это не заголовок
                product = Product(
                    article=dialog.result['article'],
                    name=dialog.result['name'],
                    quantity=dialog.result['quantity'],
                    unit=dialog.result['unit'],
                    supplier_price=dialog.result['supplier_price'],
                    markup=dialog.result['markup'],  # Добавляем наценку
                    is_group_header=False  # Явно указываем, что это не заголовок
                )
                # Рассчитываем цены
                self.calculator.calculate_client_price(
                    product,
                    self.current_offer.discount_from_supplier,
                    self.current_offer.markup_for_client
                )
                # Добавляем продукт
                self.current_offer.products.append(product)
                self.product_tree.add_product(product)
                self._update_totals()
                self.status_label.config(text=f"Добавлен продукт: {product.name}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось добавить продукт: {str(e)}")

    def _delete_product(self, event=None):
        """Удаление выбранного продукта"""
        if not self.current_offer:
            return
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите продукт для удаления")
            return
        if messagebox.askyesno("Подтверждение", "Удалить выбранные продукты?"):
            try:
                # Сортируем индексы в обратном порядке, чтобы избежать смещения
                item_indices = sorted(
                    [self.product_tree.get_item_index(item_id) for item_id in selected],
                    reverse=True
                )
                for index, item_id in zip(item_indices, selected):
                    if index is not None and 0 <= index < len(self.current_offer.products):
                        del self.current_offer.products[index]
                        self.product_tree.delete(item_id)
                        self._update_totals()
                self.status_label.config(text="Продукты удалены")
            except Exception as e:
                messagebox.showerror(
                    "Ошибка",
                    f"Не удалось удалить продукты: {str(e)}"
                )

    def _on_product_change(self, item_id: str, column: str, value: str):
        """Обработчик изменения значения в дереве"""
        try:
            # Получаем индекс элемента в списке продуктов
            index = self.product_tree.get_item_index(item_id)
            if index is None:
                return
            # Получаем продукт
            product = self.current_offer.products[index]
            # Обновляем значение в зависимости от колонки
            if column == 'article':
                product.article = value
            elif column == 'name':
                product.name = value
            elif column == 'quantity':
                product.quantity = int(value)
                product.calculate_total()
            elif column == 'unit':
                product.unit = value
            elif column == 'price_supplier':
                product.supplier_price = Decimal(value)
            elif column == 'markup':  # Добавляем обработку markup
                product.markup = Decimal(value)
            # Пересчитываем цены
            self.calculator.calculate_client_price(
                product,
                self.current_offer.discount_from_supplier,
                self.current_offer.markup_for_client
            )
            # Обновляем отображение продукта в дереве
            self.product_tree.update_product(item_id, product)
            # Обновляем итоги
            self._update_totals()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить значение: {str(e)}")

    def _create_status_bar(self):
        """Создание строки состояния"""
        self.status_label = ttk.Label(
            self.status_bar,
            text="Готов к работе"
        )
        self.status_label.pack(side=tk.LEFT, padx=5)
        self.total_label = ttk.Label(self.status_bar, text="")
        self.total_label.pack(side=tk.RIGHT, padx=5)

    def _bind_events(self):
        """Привязка обработчиков событий"""
        # Горячие клавиши для основных операций
        self.root.bind('<Control-n>', lambda e: self.callbacks['new_offer']())
        self.root.bind('<Control-o>', lambda e: self.callbacks['load_offer']())
        self.root.bind('<Control-s>', lambda e: self.callbacks['save_offer']())
        self.root.bind('<Control-Shift-s>', lambda e: self.callbacks['save_offer_as']())
        # Правка
        self.root.bind('<Control-a>', lambda e: self.callbacks['add_product']())
        self.root.bind('<F2>', lambda e: self.callbacks['edit_product']())
        self.root.bind('<Delete>', lambda e: self.callbacks['delete_product']())
        self.root.bind('<Control-h>', lambda e: self.callbacks['add_group_header']())
        # Копирование/Вставка
        self.root.bind('<Control-c>', lambda e: self.callbacks['copy_selected']())
        self.root.bind('<Control-v>', lambda e: self.callbacks['paste_selected']())
        # Перемещение
        self.root.bind('<Alt-Up>', lambda e: self.callbacks['move_selected']('up'))
        self.root.bind('<Alt-Down>', lambda e: self.callbacks['move_selected']('down'))
        # Импорт/Экспорт
        self.root.bind('<Control-i>', lambda e: self.callbacks['import_data']())
        self.root.bind('<Control-p>', lambda e: self.callbacks['export_pdf']())
        self.root.bind('<Control-x>', lambda e: self.callbacks['export_excel']())
        self.root.bind('<Control-w>', lambda e: self.callbacks['export_word']())
        # Инструменты
        self.root.bind('<F5>', lambda e: self.callbacks['update_all_data']())
        self.root.bind('<Control-r>', lambda e: self.callbacks['reset_prices']())
        # Добавляем новые привязки для работы с заголовками
        self.root.bind('<F2>', self._edit_selected)
        # Навигация и перемещение
        self.root.bind('<Insert>', lambda e: self._add_product())
        self.root.bind('<Control-h>', lambda e: self._add_group_header())
        self.root.bind('<F2>', lambda e: self._edit_product())
        self.root.bind('<Delete>', lambda e: self._delete_product())
        self.root.bind('<F5>', lambda e: self._update_all_data())
        # Привязываем новые события от дерева
        self.root.bind('<F2>', lambda e: self._edit_header())
        self.root.bind('<Delete>', lambda e: self._delete_selected())
        # Поля ввода
        self.discount_entry.bind('<KeyRelease>', self._on_discount_change)
        self.markup_entry.bind('<KeyRelease>', self._on_markup_change)  # Добавляем привязку для markup_entry
        self.vat_entry.bind('<KeyRelease>', self._on_vat_change)
        # Дерево продуктов
        self.product_tree.bind('<Double-1>', self._on_double_click)
        self.product_tree.bind('<Button-3>', self._show_context_menu)
        # Добавляем новые привязки для событий дерева
        self.product_tree.bind('<Insert>', lambda e: self._add_product())
        self.product_tree.bind('<Control-h>', lambda e: self._add_group_header())
        self.product_tree.bind('<F2>', lambda e: self._edit_product())
        self.product_tree.bind('<Delete>', lambda e: self._delete_product())
        self.product_tree.bind('<F5>', lambda e: self._update_all_data())
        # Привязываем новые события от дерева
        self.product_tree.bind('<F2>', lambda e: self._edit_header())
        self.product_tree.bind('<Delete>', lambda e: self._delete_selected())

    def _on_double_click(self, event):
        """Обработка двойного клика по продукту"""
        item = self.product_tree.identify_row(event.y)
        if not item:
            return
        # Не редактируем заголовки
        if 'header' in self.product_tree.item(item, 'tags'):
            return
        # Определяем колонку под курсором
        column = self.product_tree.identify_column(event.x)
        if not column:
            return
        # Запускаем редактирование
        self._edit_product()
        return 'break'  # Предотвращаем дальнейшую обработку события

    def _show_context_menu(self, event):
        """Показ контекстного меню"""
        menu = tk.Menu(self.root, tearoff=0)
        # Добавление пунктов меню
        menu.add_command(label="Добавить продукт (Ctrl+A)", command=self._add_product)
        menu.add_command(label="Добавить заголовок (Ctrl+H)", command=self._add_group_header)
        menu.add_separator()
        # Пункты для работы с выделенными строками
        if self.product_tree.selection():
            menu.add_command(label="Копировать (Ctrl+C)", command=self.product_tree.copy_selected)
            menu.add_command(label="Вставить (Ctrl+V)", command=self.product_tree.paste_selected)
            menu.add_command(label="Удалить (Delete)", command=self._delete_product)
            menu.add_separator()
            menu.add_command(label="Редактировать", command=lambda: self._edit_selected())
            menu.add_command(label="Переместить вверх (Alt+↑)", 
                            command=lambda: self._move_selected('up'))
            menu.add_command(label="Переместить вниз (Alt+↓)", 
                            command=lambda: self._move_selected('down'))
        menu.post(event.x_root, event.y_root)

    def _edit_selected(self, event=None):
        """Редактирование выбранного элемента (продукта или заголовка)"""
        selected = self.product_tree.selection()
        if not selected:
            return
        item_id = selected[0]
        if 'header' in self.product_tree.item(item_id, 'tags'):
            self._edit_header()
        else:
            self._edit_product() 

    def _edit_header(self):
        """Редактирование заголовка группы"""
        selected = self.product_tree.selection()
        if not selected:
            return
        item_id = selected[0]
        if not 'header' in self.product_tree.item(item_id, 'tags'):
            return
        # Получаем текущий текст заголовка
        current_name = self.product_tree.item(item_id, 'values')[1]
        # Показываем диалог редактирования
        dialog = HeaderDialog(
            self.root,
            title="Редактировать заголовок",
            initial_value=current_name
        )
        self.root.wait_window(dialog.top)
        if dialog.result:
            # Получаем индекс элемента
            index = self.product_tree.get_item_index(item_id)
            if index is not None:
                # Обновляем заголовок в модели данных
                self.current_offer.products[index].name = dialog.result
                # Обновляем отображение
                self.product_tree.set(item_id, 'name', dialog.result)
                self.status_label.config(text=f"Заголовок изменен: {dialog.result}")

    def _on_drag_start(self, event):
        """Начало перетаскивания"""
        self._drag_data = {'item': None, 'x': 0, 'y': 0}
        # Определяем элемент под курсором
        item = self.product_tree.identify_row(event.y)
        if item:
            self._drag_data['item'] = item
            self._drag_data['x'] = event.x
            self._drag_data['y'] = event.y

    def _on_drag_motion(self, event):
        """Обработка перемещения при перетаскивании"""
        if self._drag_data['item']:
            # Определяем новую позицию
            target = self.product_tree.identify_row(event.y)
            if target and target != self._drag_data['item']:
                # Перемещаем элемент
                self.product_tree.move(self._drag_data['item'], 
                                    self.product_tree.parent(target),
                                    target)
                # Обновляем порядок в списке продуктов
                self._update_products_order()

    def _on_drag_release(self, event):
        """Завершение перетаскивания"""
        self._drag_data = {'item': None, 'x': 0, 'y': 0}

    def _update_products_order(self):
        """Обновление порядка продуктов после перетаскивания"""
        if not self.current_offer:
            return
        # Получаем новый порядок из дерева
        new_order = []
        for item_id in self.product_tree.get_children():
            index = self.product_tree.get_item_index(item_id)
            if index is not None:
                new_order.append(self.current_offer.products[index])
        # Обновляем список продуктов
        self.current_offer.products = new_order
        self._update_totals()

    def _move_selected(self, direction: str):
        """Перемещение выбранных строк вверх/вниз"""
        selected = self.product_tree.selection()
        if not selected:
            return
        for item_id in selected:
            index = self.product_tree.get_item_index(item_id)
            if index is None:
                continue
            if direction == 'up' and index > 0:
                # Перемещаем вверх
                self.current_offer.products.insert(index - 1, 
                                                self.current_offer.products.pop(index))
            elif direction == 'down' and index < len(self.current_offer.products) - 1:
                # Перемещаем вниз
                self.current_offer.products.insert(index + 1, 
                                                self.current_offer.products.pop(index))
        # Обновляем отображение
        self._update_product_tree()
        # Восстанавливаем выделение
        for item_id in selected:
            self.product_tree.selection_add(item_id)

    def _update_ui(self):
        """Обновление интерфейса при изменении данных"""
        try:
            if not self.current_offer:
                return
            # Обновляем поля параметров
            self.number_entry.delete(0, tk.END)
            self.number_entry.insert(0, self.current_offer.number)
            self.date_entry.set_date(self.current_offer.date.date() if self.current_offer.date else date.today())
            self.discount_entry.delete(0, tk.END)
            self.discount_entry.insert(0, str(self.current_offer.discount_from_supplier))
            self.markup_entry.delete(0, tk.END)
            self.markup_entry.insert(0, str(self.current_offer.markup_for_client))
            self.vat_entry.delete(0, tk.END)
            self.vat_entry.insert(0, str(self.current_offer.vat))
            self.delivery_terms_entry.delete(0, tk.END)
            self.delivery_terms_entry.insert(0, self.current_offer.delivery_terms)
            self.self_pickup_entry.delete(0, tk.END)
            self.self_pickup_entry.insert(0, self.current_offer.self_pickup_warehouse)
            self.warranty_entry.delete(0, tk.END)
            self.warranty_entry.insert(0, self.current_offer.warranty)
            self.delivery_time_entry.delete(0, tk.END)
            self.delivery_time_entry.insert(0, self.current_offer.delivery_time)
            # Обновляем дерево продуктов
            self._update_product_tree()
            # Обновляем итоги
            self._update_totals()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при обновлении интерфейса: {str(e)}")

    def _update_product_tree(self):
        """Обновление отображения продуктов в дереве"""
        try:
            # Сохраняем текущее выделение
            selected = self.product_tree.selection()
            # Очищаем дерево
            for item in self.product_tree.get_children():
                self.product_tree.delete(item)
            # Добавляем продукты заново
            for product in self.current_offer.products:
                if product.is_header:
                    # Для заголовков групп
                    self.product_tree.insert(
                        '',
                        'end',
                        values=('', product.name, '', '', '', '', ''),
                        tags=('header',)
                    )
                else:
                    # Для обычных продуктов
                    self.product_tree.add_product(product)
            # Настраиваем стиль заголовков
            self.product_tree.tag_configure('header', font=('Helvetica', 10, 'bold'))
            # Восстанавливаем выделение
            if selected:
                for item in selected:
                    try:
                        self.product_tree.selection_add(item)
                    except:
                        pass
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при обновлении списка продуктов: {str(e)}")

    def _update_totals(self):
        """Обновление итоговых значений"""
        if not self.current_offer:
            return
        try:
            # Получаем итоги
            totals = self.calculator.calculate_totals(
                self.current_offer.products,
                self.current_offer.discount_from_supplier,
                self.current_offer.markup_for_client,
                self.current_offer.vat
            )
            # Форматируем строку итогов
            if totals['total_amount'] != Decimal('0'):
                margin_percentage = (totals['margin'] / totals['total_amount'] * Decimal('100')).quantize(Decimal('0.1'))
            else:
                margin_percentage = Decimal('0.0')
            total_text = (
                f"Итого: {Formatters.format_currency(totals['total_amount'])} | "
                f"НДС: {Formatters.format_currency(totals['vat_amount'])} | "
                f"Маржа: {Formatters.format_currency(totals['margin'])} "
                f"({margin_percentage}%) | "
                f"Цена клиента у Поставщика: {Formatters.format_currency(totals['client_savings'])} ({totals['client_savings_percentage']}%)"
            )
            # Обновляем метку с итогами
            self.total_label.config(text=total_text)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при расчете итогов: {str(e)}")

    def _new_offer(self, event=None):
        """Создание нового КП"""
        try:
            if self.current_offer and not messagebox.askyesno(
                "Подтверждение",
                "Создать новое КП? Несохраненные изменения будут потеряны."
            ):
                return
            # Устанавливаем значения по умолчанию
            self.vat_entry.delete(0, tk.END)
            self.vat_entry.insert(0, "20")
            self.discount_entry.delete(0, tk.END)
            self.discount_entry.insert(0, "0")
            self.markup_entry.delete(0, tk.END)
            self.markup_entry.insert(0, "0")
            # Создаем новое КП с дефолтными значениями, преобразуя числа в Decimal
            self.current_offer = CommercialOffer(
                number=self._generate_offer_number(),
                date=datetime.now(),
                discount_from_supplier=Decimal('0'),    # Преобразуем в Decimal
                markup_for_client=Decimal('0'),        # Преобразуем в Decimal
                vat=Decimal('20'),                    # Преобразуем в Decimal
                delivery_terms="",
                self_pickup_warehouse="",
                warranty="",
                delivery_time="",
                include_delivery=False,
                delivery_cost=Decimal('0'),            # Преобразуем в Decimal
                products=[]
            )
            # Обновляем интерфейс
            self._update_ui()
            self.status_label.config(text="Создано новое КП")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать новое КП: {str(e)}")

    def _handle_copy(self, event=None):
        """Обработчик копирования"""
        try:
            self.product_tree.copy_selected()
            self.status_label.config(text="Скопировано")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при копировании: {str(e)}")

    def _handle_paste(self, event=None):
        """Обработчик вставки"""
        try:
            self.paste_selected()
            self.status_label.config(text="Вставлено")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при вставке: {str(e)}")

    def _edit_product(self, event=None):
        """Редактирование выбранного продукта"""
        if not self.current_offer:
            return
        # Получаем выбранный элемент
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите продукт для редактирования")
            return
        item_id = selected[0]
        index = self.product_tree.get_item_index(item_id)
        if index is None:
            return
        product = self.current_offer.products[index]
        # Не редактируем заголовки
        if product.is_header:
            return
        try:
            # Создаем диалог редактирования с текущими значениями
            dialog = ProductDialog(
                self.root, 
                edit_mode=True,
                initial_values={
                    'article': product.article,
                    'name': product.name,
                    'quantity': product.quantity,
                    'unit': product.unit,
                    'supplier_price': product.supplier_price,
                    'markup': product.markup  # Добавляем наценку
                }
            )
            self.root.wait_window(dialog.top)
            if dialog.result:
                # Обновляем данные продукта
                product.article = dialog.result['article']
                product.name = dialog.result['name']
                product.quantity = dialog.result['quantity']
                product.unit = dialog.result['unit']
                product.supplier_price = dialog.result['supplier_price']
                product.markup = dialog.result['markup']  # Обновляем наценку
                # Пересчитываем цены
                self.calculator.calculate_client_price(
                    product,
                    self.current_offer.discount_from_supplier,
                    self.current_offer.markup_for_client
                )
                # Обновляем отображение
                self.product_tree.update_product(item_id, product)
                self._update_totals()
                self.status_label.config(text=f"Продукт обновлен: {product.name}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось отредактировать продукт: {str(e)}")

    def _on_discount_change(self, event=None):
        """Обработка изменения скидки от поставщика"""
        if not self.current_offer:
            return
        try:
            new_discount = Decimal(self.discount_entry.get() or '0')
            if new_discount != self.current_offer.discount_from_supplier:
                if not (Decimal('0') <= new_discount <= Decimal('100')):
                    raise ValueError("Скидка должна быть от 0 до 100%")
                self.current_offer.discount_from_supplier = new_discount
                self._recalculate_prices()
        except (ValueError, InvalidOperation) as e:
            messagebox.showerror("Ошибка", f"Неверный формат скидки от поставщика: {str(e)}")
            self.discount_entry.delete(0, tk.END)
            self.discount_entry.insert(0, str(self.current_offer.discount_from_supplier))

    def _on_markup_change(self, event=None):
        """Обработка изменения наценки для клиента"""
        if not self.current_offer:
            return
        try:
            new_markup = Decimal(self.markup_entry.get().strip() or '0')
            if new_markup != self.current_offer.markup_for_client:
                if not (Decimal('0') <= new_markup <= Decimal('100')):
                    raise ValueError("Наценка должна быть от 0 до 100%")
                self.current_offer.markup_for_client = new_markup
                self._recalculate_prices()
        except (ValueError, InvalidOperation) as e:
            messagebox.showerror("Ошибка", f"Неверный формат наценки для клиента: {str(e)}")
            self.markup_entry.delete(0, tk.END)
            self.markup_entry.insert(0, str(self.current_offer.markup_for_client))

    def _on_vat_change(self, event=None):
        """Обработка изменения НДС"""
        if not self.current_offer:
            return
        try:
            new_vat = Decimal(self.vat_entry.get().strip() or '0')
            if new_vat != self.current_offer.vat:
                if not (Decimal('0') <= new_vat <= Decimal('100')):
                    raise ValueError("НДС должен быть от 0 до 100%")
                self.current_offer.vat = new_vat
                self._update_totals()
        except (ValueError, InvalidOperation) as e:
            messagebox.showerror("Ошибка", f"Неверный формат НДС: {str(e)}")
            self.vat_entry.delete(0, tk.END)
            self.vat_entry.insert(0, str(self.current_offer.vat))

    def _recalculate_prices(self):
        """Пересчет всех цен"""
        if not self.current_offer:
            return
        try:
            self.calculator.update_prices(
                self.current_offer.products,
                self.current_offer.discount_from_supplier,
                self.current_offer.markup_for_client
            )
            self._update_product_tree()
            self._update_totals()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при пересчете цен: {str(e)}")

    def _import_data(self, event=None):
        """Импорт данных"""
        if not self.current_offer:
            messagebox.showwarning("Предупреждение", "Сначала создайте новое КП")
            return
        try:
            file_path = filedialog.askopenfilename(
                title="Выберите файл для импорта",
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
            if not file_path:
                return
            # Получаем список столбцов
            columns = ImportService.import_excel(Path(file_path))
            # Показываем диалог выбора столбцов
            dialog = ColumnMappingDialog(self.root, columns)
            dialog.run()
            if dialog.result:
                # Импортируем с выбранным маппингом
                products = ImportService.import_excel(
                    Path(file_path), 
                    column_mapping=dialog.result
                )
                if not products:
                    messagebox.showwarning(
                        "Предупреждение",
                        "Нет данных для импорта"
                    )
                    return
                # Добавляем импортированные продукты
                for product in products:
                    self.calculator.calculate_client_price(
                        product,
                        self.current_offer.discount_from_supplier,
                        self.current_offer.markup_for_client
                    )
                    self.current_offer.products.append(product)
                    self.product_tree.add_product(product)
                self._update_totals()
                self.status_label.config(text=f"Импортировано {len(products)} продуктов")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при импорте данных: {str(e)}")

    def _export_word(self, event=None):
        """Экспорт в Word с использованием предустановленного шаблона"""
        if not self.current_offer:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return
        try:
            # Определяем путь к шаблону
            template_path = settings.TEMPLATES_DIR / "KP_shablon.docx"
            # Проверяем существование шаблона
            if not template_path.exists():
                messagebox.showerror(
                    "Ошибка", 
                    f"Шаблон не найден: {template_path}\n"
                    f"Пожалуйста, убедитесь, что файл 'KP_shablon.docx' находится в папке templates"
                )
                return
            # Создаем директорию для экспорта, если её нет
            export_dir = self.file_service.base_path / "exports"
            export_dir.mkdir(parents=True, exist_ok=True)
            # Формируем путь для сохранения файла
            output_path = filedialog.asksaveasfilename(
                title="Сохранить КП как",
                defaultextension=".docx",
                filetypes=[("Word files", "*.docx")],
                initialdir=export_dir,
                initialfile=f"КП_{self.current_offer.number}.docx"
            )
            if not output_path:
                return
            ExportService.export_word(
                self.current_offer,
                template_path,
                Path(output_path)
            )
            self.status_label.config(text=f"КП экспортировано в Word: {Path(output_path).name}")
            # Открываем созданный файл
            open_file(Path(output_path))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при экспорте в Word: {str(e)}")

    def _export_excel(self, event=None):
        """Экспорт в Excel"""
        if not self.current_offer:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return
        try:
            output_path = filedialog.asksaveasfilename(
                title="Сохранить КП как",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialdir=self.file_service.base_path / "exports",
                initialfile=f"КП_{self.current_offer.number}.xlsx"
            )
            if not output_path:
                return
            ExportService.export_excel(
                self.current_offer,
                Path(output_path)
            )
            self.status_label.config(text=f"КП экспортировано в Excel: {Path(output_path).name}")
            open_file(Path(output_path))  # Используем функцию open_file для кроссплатформенного открытия файла
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при экспорте в Excel: {str(e)}")

    def _export_pdf(self, event=None):
        """Экспорт в PDF."""
        if not self.current_offer:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return
        try:
            output_path = filedialog.asksaveasfilename(
                title="Сохранить PDF как",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialdir=self.file_service.base_path / "exports",
                initialfile=f"КП_{self.current_offer.number}.pdf"
            )
            if not output_path:
                return
            ExportService.export_pdf(
                self.current_offer,
                Path(output_path)
            )
            self.status_label.config(text=f"КП экспортировано в PDF: {Path(output_path).name}")
            # Открываем созданный файл
            open_file(Path(output_path))
        except Exception as e:
            logger.error(f"Ошибка при экспорте в PDF: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка при экспорте в PDF: {str(e)}")
            raise ValueError(f"Ошибка при экспорте в PDF: {str(e)}")

        def _update_offer_from_ui(self):
            """Обновление данных КП из полей ввода"""
            if not self.current_offer:
                return
            try:
                self.current_offer.number = self.number_entry.get().strip()
                date_str = self.date_entry.get_date().strftime("%d.%m.%Y")
                if date_str:
                    try:
                        self.current_offer.date = datetime.strptime(date_str, "%d.%m.%Y")
                    except ValueError:
                        self.current_offer.date = datetime.now()
                else:
                    self.current_offer.date = None
                self.current_offer.discount_from_supplier = Decimal(self.discount_entry.get().strip() or '0')
                self.current_offer.markup_for_client = Decimal(self.markup_entry.get().strip() or '0')
                self.current_offer.vat = Decimal(self.vat_entry.get().strip() or '20')
                self.current_offer.delivery_terms = self.delivery_terms_entry.get().strip()
                self.current_offer.self_pickup_warehouse = self.self_pickup_entry.get().strip()
                self.current_offer.warranty = self.warranty_entry.get().strip()
                self.current_offer.delivery_time = self.delivery_time_entry.get().strip()
            except (ValueError, InvalidOperation) as e:
                messagebox.showerror("Ошибка", f"Ошибка при обновлении данных КП из интерфейса: {str(e)}")

        def _clear_all_data(self, event=None):
            """Очистка всех данных текущего коммерческого предложения"""
            try:
                if not self.current_offer:
                    messagebox.showwarning("Предупреждение", "Нет активного КП для очистки")
                    return
                if not messagebox.askyesno(
                    "Подтверждение",
                    "Вы действительно хотите очистить все данные? Несохраненные изменения будут потеряны."
                ):
                    return

                # Создаём новое КП с дефолтными значениями
                self.current_offer = CommercialOffer(
                    number=self._generate_offer_number(),
                    date=datetime.now(),
                    discount_from_supplier=Decimal('0'),
                    markup_for_client=Decimal('0'),
                    vat=Decimal('20'),
                    delivery_terms="",
                    self_pickup_warehouse="",
                    warranty="",
                    delivery_time="",
                    include_delivery=False,
                    delivery_cost=Decimal('0'),
                    products=[]
                )

                # Обновляем интерфейс
                self._update_ui()
                self.status_label.config(text="Все данные очищены")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось очистить данные: {str(e)}")

        def _update_all_data(self, event=None):
            """Обновить все данные (пересчитать цены)"""
            if not self.current_offer:
                messagebox.showwarning("Предупреждение", "Нет активного КП для обновления")
                return
            try:
                # Получаем актуальные значения скидки и наценки
                discount = Decimal(self.discount_entry.get().strip() or '0')
                markup = Decimal(self.markup_entry.get().strip() or '0')
                # Проверяем корректность значений
                if not (0 <= discount <= 100) or not (0 <= markup <= 100):
                    raise ValueError("Скидка и наценка должны быть от 0 до 100%")
                # Обновляем значения в текущем предложении
                self.current_offer.discount_from_supplier = discount
                self.current_offer.markup_for_client = markup
                # Пересчитываем цены для всех продуктов
                for product in self.current_offer.products:
                    if not product.is_header:  # Пропускаем заголовки
                        self.calculator.calculate_client_price(
                            product,
                            self.current_offer.discount_from_supplier,
                            self.current_offer.markup_for_client
                        )
                # Обновляем отображение в таблице
                self._update_product_tree()
                # Обновляем итоги
                self._update_totals()
                self.status_label.config(text="Цены обновлены и пересчитаны")
            except (ValueError, InvalidOperation) as e:
                messagebox.showerror("Ошибка", f"Не удалось обновить данные: {str(e)}")

        def _reset_prices(self, event=None):
            """Сбросить цены всех продуктов на основе текущих параметров"""
            if not self.current_offer:
                messagebox.showwarning("Предупреждение", "Нет активного КП для сброса цен")
                return
            try:
                # Пересчитываем цены всех продуктов
                self.calculator.update_prices(
                    self.current_offer.products,
                    self.current_offer.discount_from_supplier,
                    self.current_offer.markup_for_client
                )
                # Обновляем отображение в таблице продуктов
                self._update_product_tree()
                # Обновляем итоговые суммы
                self._update_totals()
                self.status_label.config(text="Цены всех продуктов сброшены и пересчитаны")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сбросить цены: {str(e)}")

        def _validate_quantity(self, value: str) -> int:
            """Валидация количества"""
            try:
                quantity = int(value)
                if quantity < 0:
                    raise ValueError("Количество не может быть отрицательным")
                return quantity
            except ValueError:
                raise ValueError("Количество должно быть целым числом")

        def _validate_price(self, value: str) -> Decimal:
            """Валидация цены"""
            try:
                price = Decimal(value.replace(',', '.'))
                if price < 0:
                    raise ValueError("Цена не может быть отрицательной")
                return price
            except (InvalidOperation, ValueError):
                raise ValueError("Цена должна быть числом")

        def _add_group_header(self, event=None):
            """Добавление заголовка группы"""
            if not self.current_offer:
                messagebox.showwarning("Предупреждение", "Сначала создайте новое КП")
                return
            dialog = HeaderDialog(self.root)
            self.root.wait_window(dialog.top)
            if dialog.result:
                # Создаем продукт-заголовок
                header_product = Product(
                    article='',
                    name=dialog.result,
                    quantity=0,
                    unit='',
                    supplier_price=Decimal('0'),
                    is_group_header=True
                )
                # Определяем позицию для вставки
                insert_index = len(self.current_offer.products)
                selected = self.product_tree.selection()
                if selected:
                    index = self.product_tree.get_item_index(selected[0])
                    if index is not None:
                        insert_index = index
                # Добавляем в список и в дерево
                self.current_offer.products.insert(insert_index, header_product)
                self._update_product_tree()
                self.status_label.config(text=f"Добавлен заголовок группы: {dialog.result}")

        def _delete_selected(self, event=None):
            """Удаление выбранных элементов (продуктов или заголовков)"""
            selected = self.product_tree.selection()
            if not selected:
                messagebox.showwarning("Предупреждение", "Выберите элементы для удаления")
                return
            # Определяем тип удаляемых элементов для сообщения
            is_header = 'header' in self.product_tree.item(selected[0], 'tags')
            message = "Удалить выбранные заголовки?" if is_header else "Удалить выбранные продукты?"
            if messagebox.askyesno("Подтверждение", message):
                # Сортируем индексы в обратном порядке
                indices = sorted(
                    [self.product_tree.get_item_index(item_id) for item_id in selected],
                    reverse=True
                )
                # Удаляем элементы
                for index in indices:
                    if index is not None:
                        del self.current_offer.products[index]
                # Обновляем отображение
                for item_id in selected:
                    self.product_tree.delete(item_id)
                # Обновляем итоги
                self._update_totals()
                # Обновляем статус
                elements_type = "заголовков" if is_header else "продуктов"
                self.status_label.config(text=f"Удалено {len(selected)} {elements_type}")

        def paste_selected(self):
            """Вставляет строки из буфера обмена"""
            if not self.current_offer:
                return
            try:
                clipboard_content = pyperclip.paste()
                if not clipboard_content:
                    return
                rows = clipboard_content.strip().split('\n')
                for row in rows:
                    values = row.split('\t')
                    if len(values) < 7:
                        continue
                    # Обрабатываем заголовок
                    if not values[0]:  # Если артикул пустой
                        header_product = Product(
                            article='',
                            name=values[1],
                            quantity=0,
                            unit='',
                            supplier_price=Decimal('0'),
                            is_group_header=True
                        )
                        self.current_offer.products.append(header_product)
                        continue
                    # Обрабатываем обычный продукт
                    try:
                        product = Product(
                            article=values[0],
                            name=values[1],
                            quantity=int(str(values[2]).replace(' ', '')),
                            unit=values[3],
                            supplier_price=Decimal(str(values[4]).replace(' ', '').replace('₽', '').replace(',', '.') or '0'),
                            markup=Decimal(str(values[5]).replace(' ', '').replace(',', '.') or '0'),  # Добавляем наценку
                            is_group_header=False
                        )
                        # Пересчитываем цены
                        self.calculator.calculate_client_price(
                            product,
                            self.current_offer.discount_from_supplier,
                            self.current_offer.markup_for_client
                        )
                        self.current_offer.products.append(product)
                    except (ValueError, IndexError, DecimalException) as e:
                        logger.warning(f"Ошибка при обработке строки: {str(e)}")
                        continue
                # Обновляем отображение
                self._update_product_tree()
                self._update_totals()
                self.status_label.config(text="Вставка выполнена")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при вставке: {str(e)}")

        def reset_clipboard_handlers(self):
            """Сброс обработчиков буфера обмена"""
            try:
                # Обновляем привязки для дерева продуктов
                self.product_tree.bind('<Control-c>', self.product_tree.copy_selected)
                self.product_tree.bind('<Control-v>', self.product_tree.paste_selected)
                # Обновляем привязки для главного окна
                self.root.bind('<Control-c>', lambda e: self.product_tree.copy_selected())
                self.root.bind('<Control-v>', lambda e: self.product_tree.paste_selected())
                logger.info("Clipboard handlers reset successfully")
            except Exception as e:
                logger.error(f"Error in reset_clipboard_handlers: {str(e)}")
                print(f"Error in reset_clipboard_handlers: {str(e)}")

        def reset_bindings(self):
            """Сброс и переустановка всех привязок"""
            if hasattr(self, 'product_tree'):
                self.product_tree._setup_bindings()
            # Добавляем глобальные привязки
            self.root.bind('<Control-n>', self._new_offer)      # Ctrl+N - Новое КП
            self.root.bind('<Control-o>', self._load_offer)     # Ctrl+O - Открыть
            self.root.bind('<Control-s>', self._save_offer)     # Ctrl+S - Сохранить
            self.root.bind('<Control-Shift-s>', self._save_offer_as) # Ctrl+Shift+S - Сохранить как
            self.root.bind('<Control-i>', self._import_data)    # Ctrl+I - Импорт
            self.root.bind('<F2>', self._edit_product)          # F2 - Редактировать
            self.root.bind('<Control-p>', self._export_pdf)     # Ctrl+P - Экспорт PDF
            self.root.bind('<Control-x>', self._export_excel)   # Ctrl+X - Экспорт Excel
            self.root.bind('<Control-w>', self._export_word)    # Ctrl+W - Экспорт Word
            # Работа с продуктами
            self.root.bind('<Control-a>', self._add_product)    # Ctrl+A - Добавить
            self.root.bind('<Delete>', self._delete_product)    # Delete - Удалить
            self.root.bind('<Control-h>', self._add_group_header)  # Ctrl+H - Добавить заголовок
            # Буфер обмена
            self.root.bind('<Control-c>', lambda e: self.product_tree.copy_selected())
            self.root.bind('<Control-v>', lambda e: self.product_tree.paste_selected())
            # Обновление данных
            self.root.bind('<F5>', self._update_all_data)       # F5 - Обновить все
            self.root.bind('<Control-r>', self._reset_prices)   # Ctrl+R - Сбросить цены
            # Добавляем новые привязки для работы с заголовками
            self.root.bind('<F2>', self._edit_selected)
            # Навигация и перемещение
            self.root.bind('<Alt-Up>', lambda e: self._move_selected('up'))      # Alt+↑ - Переместить вверх
            self.root.bind('<Alt-Down>', lambda e: self._move_selected('down'))  # Alt+↓ - Переместить вниз
            # Поля ввода
            self.discount_entry.bind('<KeyRelease>', self._on_discount_change)
            self.markup_entry.bind('<KeyRelease>', self._on_markup_change)  # Добавляем привязку для markup_entry
            self.vat_entry.bind('<KeyRelease>', self._on_vat_change)
            # Дерево продуктов
            self.product_tree.bind('<Double-1>', self._on_double_click)
            self.product_tree.bind('<Button-3>', self._show_context_menu)
            # Добавляем новые привязки для событий дерева
            self.product_tree.bind('<Insert>', lambda e: self._add_product())
            self.product_tree.bind('<Control-h>', lambda e: self._add_group_header())
            self.product_tree.bind('<F2>', lambda e: self._edit_product())
            self.product_tree.bind('<Delete>', lambda e: self._delete_product())
            self.product_tree.bind('<F5>', lambda e: self._update_all_data())
            # Привязываем новые события от дерева
            self.product_tree.bind('<F2>', lambda e: self._edit_header())
            self.product_tree.bind('<Delete>', lambda e: self._delete_selected())

        def _refresh_interface(self):
            """Полное обновление интерфейса"""
            self._update_ui()
            self._bind_events()
            self.product_tree._setup_bindings()
            self.reset_clipboard_handlers()

class ToolTip:
    """Класс для создания всплывающих подсказок"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind('<Enter>', self.show_tooltip)
        self.widget.bind('<Leave>', self.hide_tooltip)

    def show_tooltip(self, event=None):
        """Показать подсказку"""
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # Создаем окно подсказки
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(self.tooltip, text=self.text, 
                          background="#ffffe0", relief='solid', borderwidth=1)
        label.pack()

    def hide_tooltip(self, event=None):
        """Скрыть подсказку"""
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None