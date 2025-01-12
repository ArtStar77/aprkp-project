# src/ui/product_tree.py

import tkinter as tk
from tkinter import ttk, messagebox
from decimal import Decimal, InvalidOperation
from typing import Callable, Optional
import pyperclip

from src.models.product import Product
from src.utils.formatters import Formatters

class ProductTreeView(ttk.Treeview):
    def __init__(self, parent, on_change: Callable[[str, str, str], None], *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.on_change = on_change
        
        # Настройка выделения
        self.configure(selectmode='extended')
        
        # Стили для ячеек
        self.tag_configure('selected_cell', background='#0078D7', foreground='white')
        self.tag_configure('editable', background='#F0F0F0')
        self.tag_configure('header', font=('Helvetica', 10, 'bold'))
        self.tag_configure('readonly', background='#F8F8F8')
        
        # Текущая выбранная ячейка
        self.current_cell = None
        
        # Состояние перетаскивания
        self._drag_data = {'item': None, 'x': 0, 'y': 0}
        
        # Определяем колонки
        self['columns'] = ('article', 'name', 'quantity', 'unit', 'price_supplier', 'markup', 'price_client', 'total_price')
        self.column('#0', width=0, stretch=tk.NO)  # Скрываем первый столбец
        
        # Настройка ширины и заголовков столбцов
        column_configs = {
            'article': ('Артикул', 100),
            'name': ('Наименование', 300),
            'quantity': ('Количество', 80),
            'unit': ('Ед.изм.', 60),
            'price_supplier': ('Цена поставщика', 120),
            'price_client': ('Цена клиента', 120),
            'total_price': ('Сумма', 120)
        }
        
        for col, (text, width) in column_configs.items():
            self.heading(col, text=text)
            self.column(col, width=width, minwidth=width)
        
        # Определяем редактируемые и нередактируемые столбцы
        self.editable_columns = ['article', 'name', 'quantity', 'unit', 'price_supplier', 'markup']
        self.readonly_columns = ['price_client', 'total_price']
        
        # Добавляем поддержку сортировки
        for col in self['columns']:
            self.heading(col, text=column_configs[col][0],
                        command=lambda c=col: self._sort_by_column(c))
        
        # Привязываем обработчики событий
        self._setup_bindings()

    def _setup_bindings(self):
        """Настройка обработчиков событий"""
        # Основные события редактирования
        self.bind('<Double-1>', self._on_double_click)
        self.bind('<Return>', self._on_enter)
        self.bind('<F2>', lambda e: self.event_generate("<<EditProduct>>"))
        self.bind('<<EditProduct>>', lambda e: self._handle_edit())
        
        # Контекстное меню
        self.bind('<Button-3>', self._show_context_menu)
        
        # Буфер обмена
        self.bind('<Control-c>', lambda e: self.copy_selected())
        self.bind('<Control-v>', lambda e: self.paste_selected())
        
        # Навигация
        self.bind('<Tab>', self._on_tab)
        self.bind('<Shift-Tab>', self._on_shift_tab)
        self.bind('<Control-a>', self._select_all)

    def _handle_edit(self):
        """Обработка редактирования"""
        item = self.selection()[0] if self.selection() else None
        if item:
            column = '#1'  # Начинаем с первой колонки
            self._start_edit(item, column)
    
    def _on_tree_press(self, event):
        """Обработка нажатия на элемент дерева"""
        # Определяем элемент под курсором
        item = self.identify_row(event.y)
        if item:
            # Сохраняем данные о перетаскивании
            self._drag_data = {
                'item': item,
                'x': event.x,
                'y': event.y
            }
    
    def _on_tree_motion(self, event):
        """Обработка перемещения элемента дерева"""
        if self._drag_data.get('item'):
            # Определяем новую позицию
            target = self.identify_row(event.y)
            if target and target != self._drag_data['item']:
                # Получаем позиции элементов
                target_index = self.index(target)
                item_index = self.index(self._drag_data['item'])
                
                if target_index != item_index:
                    # Перемещаем элемент
                    self.move(self._drag_data['item'], '', target_index)
                    # Обновляем порядок в модели данных
                    if hasattr(self.master, '_update_products_order'):
                        self.master._update_products_order()
    
    def _on_tree_release(self, event):
        """Завершение перетаскивания"""
        self._drag_data = {'item': None, 'x': 0, 'y': 0}

    def _show_context_menu(self, event):
        """Показ контекстного меню"""
        menu = tk.Menu(self, tearoff=0)
        
        # Получаем выбранный элемент
        selected = self.selection()
        is_header = selected and 'header' in self.item(selected[0], 'tags')
        
        # Добавляем основные команды
        menu.add_command(
            label="Добавить продукт", 
            command=lambda: self.event_generate("<<AddProduct>>"),
            accelerator="Ctrl+A"
        )
        menu.add_command(
            label="Добавить заголовок", 
            command=lambda: self.event_generate("<<AddHeader>>"),
            accelerator="Ctrl+H"
        )
        menu.add_separator()
        
        if selected:
            if is_header:
                # Команды для заголовков
                menu.add_command(
                    label="Редактировать заголовок",
                    command=lambda: self.event_generate("<<EditHeader>>"),
                    accelerator="F2"
                )
                menu.add_command(
                    label="Удалить заголовок",
                    command=lambda: self.event_generate("<<DeleteHeader>>"),
                    accelerator="Delete"
                )
            else:
                # Команды для продуктов
                menu.add_command(
                    label="Редактировать", 
                    command=lambda: self.event_generate("<<EditProduct>>"),
                    accelerator="F2"
                )
                menu.add_command(
                    label="Удалить", 
                    command=lambda: self.event_generate("<<DeleteProduct>>"),
                    accelerator="Delete"
                )
            
            menu.add_separator()
            
            # Общие команды для выделенных элементов
            menu.add_command(
                label="Копировать",
                command=self.copy_selected,
                accelerator="Ctrl+C"
            )
            menu.add_command(
                label="Вставить",
                command=self.paste_selected,
                accelerator="Ctrl+V"
            )
            menu.add_separator()
            
            menu.add_command(
                label="Переместить вверх",
                command=lambda: self._move_selection('up'),
                accelerator="Alt+↑"
            )
            menu.add_command(
                label="Переместить вниз",
                command=lambda: self._move_selection('down'),
                accelerator="Alt+↓"
            )
        
        menu.add_separator()
        menu.add_command(
            label="Обновить все",
            command=lambda: self.event_generate("<<UpdateAll>>"),
            accelerator="F5"
        )
        
        menu.post(event.x_root, event.y_root)

    def _edit_selected(self, item):
        """Редактирование выбранного элемента"""
        # Проверяем, что элемент не является заголовком
        if 'header' in self.item(item, 'tags'):
            return
        
        # Начинаем редактирование первой доступной редактируемой ячейки
        for col in self.editable_columns:
            column = '#' + str(self['columns'].index(col) + 1)
            self._start_edit(item, column)
            break
    
    def _on_f2(self, event):
        """Обработка нажатия F2"""
        item = self.focus()
        if not item:
            return
        
        # Получаем текущую колонку или первую редактируемую
        column = self.identify_column(event.x)
        if not column:
            # Если колонка не определена, берем первую редактируемую
            for col in self.editable_columns:
                column = '#' + str(self['columns'].index(col) + 1)
                break
        
        # Запускаем редактирование
        self._start_edit(item, column)
        return 'break'

    def _on_key_press(self, event):
        """Обработка нажатия клавиш"""
        # Игнорируем специальные клавиши
        if event.char and event.char.isprintable():
            item = self.focus()
            if not item:
                return
                
            column = self.identify_column(self.winfo_pointerx() - self.winfo_rootx())
            if not column:
                # Берем первую редактируемую колонку
                for col in self.editable_columns:
                    column = '#' + str(self['columns'].index(col) + 1)
                    break
            
            # Начинаем редактирование и вставляем нажатую клавишу
            self._start_edit(item, column, initial_text=event.char)
            return 'break'

    def _start_edit(self, item: str, column: str, initial_text: str = None):
        """Начало редактирования ячейки"""
        if not item:
            return
                
        # Получаем название колонки
        try:
            col_num = int(column.replace('#', '')) - 1
            column_name = self['columns'][col_num]
        except (ValueError, IndexError):
            return
                    
        # Разрешаем редактирование заголовков только для колонки 'name'
        is_header = 'header' in self.item(item, 'tags')
        if is_header and column_name != 'name':
            return
                    
        # Проверяем, можно ли редактировать колонку
        if not is_header and column_name not in self.editable_columns:
            return
                    
        # Получаем геометрию ячейки
        bbox = self.bbox(item, column)
        if not bbox:
            return
        x, y, w, h = bbox
        
        # Создаем окно редактирования
        edit_window = tk.Toplevel(self)
        edit_window.withdraw()  # Скрываем окно до настройки
        edit_window.overrideredirect(True)
        edit_window.transient(self)
        
        # Получаем текущее значение и очищаем его от форматирования
        current_value = self.set(item, column_name)
        if column_name == 'price_supplier':
            current_value = current_value.replace('₽', '').replace(' ', '').replace(',', '.').strip()
        elif column_name == 'quantity':
            current_value = current_value.replace(' ', '')
        
        # Если передан начальный текст, используем его вместо текущего значения
        if initial_text is not None:
            current_value = initial_text
        
        # Создаем поле ввода
        edit_var = tk.StringVar(value=current_value)
        entry = ttk.Entry(edit_window, textvariable=edit_var, justify='left')
        entry.pack(fill=tk.BOTH, expand=True)
        
        # Устанавливаем размер и позицию окна редактирования
        edit_window.geometry(f'{w}x{h}+{self.winfo_rootx() + x}+{self.winfo_rooty() + y}')
        edit_window.deiconify()  # Показываем окно
        entry.focus_set()
        
        # Если есть начальный текст, помещаем курсор в конец
        if initial_text is not None:
            entry.icursor(tk.END)
        else:
            entry.select_range(0, tk.END)
        
        # Привязываем события
        entry.bind('<Return>', lambda e: self._validate_and_save(edit_window, entry, item, column_name, edit_var))
        entry.bind('<Escape>', lambda e: edit_window.destroy())
        edit_window.bind('<FocusOut>', lambda e: self._validate_and_save(edit_window, entry, item, column_name, edit_var))

    def _validate_and_save(self, edit_window, entry, item, column_name, edit_var):
        """Проверка и сохранение значения ячейки"""
        try:
            new_value = edit_var.get().strip()
            
            # Валидация в зависимости от типа столбца
            if column_name == 'quantity':
                if not new_value:
                    new_value = '0'
                value = int(new_value)
                if value < 0:
                    raise ValueError("Количество не может быть отрицательным")
                new_value = str(value)
                
            elif column_name == 'price_supplier':
                if not new_value:
                    new_value = '0.00'
                # Очищаем строку от всех символов кроме цифр, точки и запятой
                clean_value = ''.join(c for c in new_value if c.isdigit() or c in '.,')
                clean_value = clean_value.replace(',', '.')
                
                # Проверяем, что осталось что-то для преобразования
                if not clean_value:
                    clean_value = '0.00'
                    
                value = Decimal(clean_value)
                if value < 0:
                    raise ValueError("Цена не может быть отрицательной")
                new_value = str(value)
                
            elif column_name in ['article', 'name', 'unit']:
                # Для текстовых полей просто очищаем пробелы по краям
                new_value = new_value.strip()
            
            elif column_name == 'markup':
                product.markup = Decimal(value)
                self.calculator.calculate_client_price(
                    product,
                    self.master.current_offer.discount_from_supplier,
                    self.master.current_offer.markup_for_client
                )
                
            # Закрываем окно редактирования и обновляем значение
            edit_window.destroy()
            self.on_change(item, column_name, new_value)
            
        except (ValueError, InvalidOperation) as e:
            messagebox.showerror("Ошибка", f"Неверный формат данных: {str(e)}")
            entry.focus_set()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
            entry.focus_set()

    def _on_double_click(self, event):
        """Обработка двойного клика для редактирования ячеек"""
        # Определяем элемент и колонку под курсором
        item = self.identify('item', event.x, event.y)
        column = self.identify('column', event.x, event.y)
        
        if not item or not column:
            return 'break'

        # Запускаем редактирование
        self._start_edit(item, column)
        
        return 'break'

    def add_product(self, product: Product):
        """Добавляет продукт в дерево"""
        try:
            tags = []
            if product.is_header:
                tags.append('header')
                values = ('', product.name, '', '', '', '', '')
            else:
                # Добавляем теги для редактируемых и нередактируемых столбцов
                tags.extend(['editable' if col in self.editable_columns else 'readonly'
                           for col in self['columns']])
                values = (
                    product.article,
                    product.name,
                    Formatters.format_quantity(product.quantity),
                    product.unit,
                    Formatters.format_currency(product.supplier_price),
                    Formatters.format_currency(product.client_price),
                    Formatters.format_currency(product.total_price)
                )
            
            return self.insert('', 'end', values=values, tags=tags)
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении продукта: {str(e)}")

    def update_product(self, item_id: str, product: Product):
        """Обновляет продукт в дереве"""
        try:
            if product.is_header:
                values = ('', product.name, '', '', '', '', '')
                tags = ('header',)
            else:
                values = (
                    product.article,
                    product.name,
                    Formatters.format_quantity(product.quantity),
                    product.unit,
                    Formatters.format_currency(product.supplier_price),
                    Formatters.format_currency(product.client_price),
                    Formatters.format_currency(product.total_price)
                )
                tags = ['editable' if col in self.editable_columns else 'readonly'
                       for col in self['columns']]
            
            self.item(item_id, values=values, tags=tags)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при копировании: {str(e)}")
    
    def get_item_index(self, item_id: str) -> Optional[int]:
        """Возвращает индекс элемента в дереве, исключая заголовки"""
        children = self.get_children()
        try:
            index = children.index(item_id)
            if 'header' in self.item(item_id, 'tags'):
                return None
            return index
        except ValueError:
            return None

    def _on_enter(self, event):
        """Обработка нажатия Enter"""
        item = self.focus()
        if item:
            column = self.identify_column(self.winfo_pointerx() - self.winfo_rootx())
            if not column:
                # Если колонка не определена, берем первую редактируемую
                for col in self.editable_columns:
                    column = '#' + str(self['columns'].index(col) + 1)
                    break
            self._start_edit(item, column)
        return 'break'

    def _on_tab(self, event):
        """Обработка нажатия Tab"""
        item = self.focus()
        if item:
            next_item, next_col = self._get_next_cell(item)
            if next_item:
                self.focus(next_item)
                self.selection_set(next_item)
                return 'break'

    def _on_shift_tab(self, event):
        """Обработка нажатия Shift+Tab"""
        item = self.focus()
        if item:
            prev_item, prev_col = self._get_prev_cell(item)
            if prev_item:
                self.focus(prev_item)
                self.selection_set(prev_item)
                return 'break'

    def _get_next_cell(self, current_item):
        """Получение следующей ячейки для навигации"""
        cols = self['columns']
        current_col = self.identify_column(self.winfo_pointerx() - self.winfo_rootx())
        current_col = current_col.replace('#', '')
        
        try:
            col_idx = int(current_col) - 1
        except ValueError:
            col_idx = 0

        if col_idx < len(cols) - 1:
            return current_item, cols[col_idx + 1]
        else:
            items = self.get_children('')
            try:
                next_idx = items.index(current_item) + 1
                if next_idx < len(items):
                    return items[next_idx], cols[0]
            except ValueError:
                pass
        return None, None

    def _get_prev_cell(self, current_item):
        """Получение предыдущей ячейки для навигации"""
        cols = self['columns']
        current_col = self.identify_column(self.winfo_pointerx() - self.winfo_rootx())
        current_col = current_col.replace('#', '')
        
        try:
            col_idx = int(current_col) - 1
        except ValueError:
            col_idx = 0

        if col_idx > 0:
            return current_item, cols[col_idx - 1]
        else:
            items = self.get_children('')
            try:
                prev_idx = items.index(current_item) - 1
                if prev_idx >= 0:
                    return items[prev_idx], cols[-1]
            except ValueError:
                pass
        return None, None

    def _select_all(self, event):
        """Выделение всех строк"""
        for item in self.get_children():
            self.selection_add(item)
        return 'break'

    def _move_selection(self, direction: str):
        """Перемещение выбранных строк вверх/вниз"""
        selected_items = self.selection()
        if not selected_items:
            return 'break'
            
        all_items = self.get_children('')
        
        if direction == 'up':
            for item in selected_items:
                idx = all_items.index(item)
                if idx > 0:
                    self.move(item, '', idx - 1)
        else:  # down
            for item in reversed(selected_items):
                idx = all_items.index(item)
                if idx < len(all_items) - 1:
                    self.move(item, '', idx + 1)
        
        self.master._update_products_order()
        return 'break'

    def copy_selected(self):
        """Копирует выбранные строки в буфер обмена"""
        try:
            selected_items = self.selection()
            if not selected_items:
                return

            copied_data = []
            for item in selected_items:
                values = self.item(item, 'values')
                copied_data.append('\t'.join(str(v) for v in values))

            if copied_data:
                pyperclip.copy('\n'.join(copied_data))
                return True  # Успешное копирование
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при копировании: {str(e)}")
            return False

    def paste_selected(self):
        """Вставляет строки из буфера обмена"""
        try:
            clipboard_content = pyperclip.paste()
            if not clipboard_content:
                return

            rows = clipboard_content.strip().split('\n')
            for row in rows:
                values = row.split('\t')
                if len(values) < 7:
                    continue

                if not values[0]:  # Заголовок
                    product = Product(
                        article='',
                        name=values[1],
                        quantity=0,
                        unit='',
                        supplier_price=Decimal('0'),
                        is_group_header=True
                    )
                else:  # Обычный продукт
                    product = Product(
                        article=values[0],
                        name=values[1],
                        quantity=int(Formatters.parse_quantity(values[2]) or 0),
                        unit=values[3],
                        supplier_price=Formatters.parse_currency(values[4]) or Decimal('0'),
                        is_group_header=False
                    )

                self.master.calculator.calculate_client_price(
                    product,
                    self.master.current_offer.discount_from_supplier,
                    self.master.current_offer.markup_for_client
                )
                
                self.master.current_offer.products.append(product)
                self.add_product(product)

            self.master._update_totals()
            self.master.status_label.config(text="Вставка выполнена")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при вставке: {str(e)}")

    def _sort_by_column(self, col):
        """Сортировка по колонке"""
        data = []
        for item in self.get_children(''):
            value = self.set(item, col)
            if col in ['quantity']:
                try:
                    value = int(Formatters.parse_quantity(value) or 0)
                except:
                    value = 0
            elif col in ['price_supplier', 'price_client', 'total_price']:
                try:
                    value = Formatters.parse_currency(value) or Decimal('0')
                except:
                    value = Decimal('0')
            data.append((value, item))
        
        data.sort(reverse=getattr(self, '_sort_reverse', False))
        
        for idx, (_, item) in enumerate(data):
            self.move(item, '', idx)
        
        self._sort_reverse = not getattr(self, '_sort_reverse', False)
        
        self.master._update_products_order()
        
    def _highlight_cell(self, item, column):
        """Подсветка редактируемой ячейки"""
        for tag in self.tags():
            if tag.startswith('editing_'):
                self.tag_remove(tag, 'all')
        
        tag = f'editing_{column}'
        self.tag_configure(tag, background='#e0e0ff')
        self.item(item, tags=(tag,))
    
    def edit_cell(self, item, column):
        """Публичный метод для запуска редактирования ячейки"""
        self._start_edit(item, column)