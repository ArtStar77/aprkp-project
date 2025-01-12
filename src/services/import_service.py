# src/services/import_service.py
import pandas as pd
import pdfplumber
from pathlib import Path
from typing import List, Optional, Dict
from decimal import Decimal, InvalidOperation
import logging
from src.models.product import Product
from tkinter import messagebox

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ImportService:
    @staticmethod
    def import_excel(file_path: Path, column_mapping: Optional[Dict[str, str]] = None) -> List[Product]:
        """
        Импорт продуктов из Excel с пользовательским маппингом столбцов.
        
        Args:
            file_path: Путь к Excel файлу.
            column_mapping: Словарь соответствия полей и столбцов Excel.
        
        Returns:
            Список продуктов.
        """
        products = []
        try:
            # Читаем Excel файл
            df = pd.read_excel(file_path)
            # Получаем список столбцов для выбора, если маппинг не передан
            if not column_mapping:
                return list(df.columns)  # Возвращаем список столбцов для выбора
            # Проверяем наличие обязательного столбца наименования
            if not column_mapping.get('name'):
                raise ValueError("Не указан столбец наименования")
            current_group = None
            for _, row in df.iterrows():
                try:
                    # Получаем наименование (обязательное поле)
                    name = str(row[column_mapping['name']]).strip()
                    # Пропускаем пустые строки
                    if pd.isna(name) or not name:
                        continue
                    # Проверяем, является ли строка заголовком группы
                    if name.lower().startswith(('система', 'группа')):
                        current_group = name
                        products.append(Product(
                            article='',
                            name=name,
                            quantity=0,
                            unit='',
                            supplier_price=Decimal('0'),
                            is_group_header=True
                        ))
                        continue
                    # Получаем остальные поля с учетом маппинга
                    article = str(row[column_mapping.get('article', '')]).strip() if column_mapping.get('article') else ''
                    # Обработка количества
                    quantity = 1
                    if column_mapping.get('quantity'):
                        quantity_val = row[column_mapping['quantity']]
                        try:
                            quantity = int(float(quantity_val)) if not pd.isna(quantity_val) else 1
                        except (ValueError, TypeError):
                            logger.warning(f"Неверное количество в строке: {quantity_val}")
                            continue
                    # Обработка единиц измерения
                    unit = 'шт.'
                    if column_mapping.get('unit'):
                        unit_val = row[column_mapping['unit']]
                        unit = str(unit_val).strip() if not pd.isna(unit_val) else 'шт.'
                    # Обработка цены
                    price = Decimal('0')
                    if column_mapping.get('price'):
                        price_str = str(row[column_mapping['price']])
                        try:
                            price_str = ''.join(c for c in price_str if c.isdigit() or c in '.,')
                            price_str = price_str.replace(',', '.')
                            price = Decimal(price_str) if price_str else Decimal('0')
                        except (InvalidOperation, ValueError):
                            logger.warning(f"Неверная цена в строке: {price_str}")
                            continue
                    # Добавляем отступы для элементов группы
                    if current_group and not name.startswith(('    ', '\t')):
                        name = '    ' + name
                    # Создаем продукт
                    product = Product(
                        article=article,
                        name=name,
                        quantity=quantity,
                        unit=unit,
                        supplier_price=price,
                        is_group_header=False
                    )
                    products.append(product)
                    logger.info(f"Импортирован продукт: {product.name}")
                except Exception as e:
                    logger.error(f"Ошибка при обработке строки: {str(e)}")
                    continue
            return products
        except Exception as e:
            logger.error(f"Ошибка при импорте Excel: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка при импорте Excel: {str(e)}")
            raise ValueError(f"Ошибка при импорте Excel: {str(e)}")
    
    def import_excel_with_progress(file_path: Path, column_mapping: Dict[str, str]) -> List[Product]:
        """
        Импорт продуктов из Excel с отображением прогресса.
        
        Args:
            file_path: Путь к Excel файлу.
            column_mapping: Словарь соответствия полей и столбцов Excel.
        
        Returns:
            Список продуктов.
        """
        products = []
        try:
            # Читаем Excel файл
            df = pd.read_excel(file_path)
            total_rows = len(df)
            progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
            progress_bar.pack(pady=10)
            progress_bar["maximum"] = total_rows
            
            for index, row in df.iterrows():
                try:
                    # Получаем наименование (обязательное поле)
                    name = str(row[column_mapping['name']]).strip()
                    # Пропускаем пустые строки
                    if pd.isna(name) or not name:
                        continue
                    # Проверяем, является ли строка заголовком группы
                    if name.lower().startswith(('система', 'группа')):
                        current_group = name
                        products.append(Product(
                            article='',
                            name=name,
                            quantity=0,
                            unit='',
                            supplier_price=Decimal('0'),
                            is_group_header=True
                        ))
                        continue
                    # Получаем остальные поля с учетом маппинга
                    article = str(row[column_mapping.get('article', '')]).strip() if column_mapping.get('article') else ''
                    # Обработка количества
                    quantity = 1
                    if column_mapping.get('quantity'):
                        quantity_val = row[column_mapping['quantity']]
                        try:
                            quantity = int(float(quantity_val)) if not pd.isna(quantity_val) else 1
                        except (ValueError, TypeError):
                            logger.warning(f"Неверное количество в строке: {quantity_val}")
                            continue
                    # Обработка единиц измерения
                    unit = 'шт.'
                    if column_mapping.get('unit'):
                        unit_val = row[column_mapping['unit']]
                        unit = str(unit_val).strip() if not pd.isna(unit_val) else 'шт.'
                    # Обработка цены
                    price = Decimal('0')
                    if column_mapping.get('price'):
                        price_str = str(row[column_mapping['price']])
                        try:
                            price_str = ''.join(c for c in price_str if c.isdigit() or c in '.,')
                            price_str = price_str.replace(',', '.')
                            price = Decimal(price_str) if price_str else Decimal('0')
                        except (InvalidOperation, ValueError):
                            logger.warning(f"Неверная цена в строке: {price_str}")
                            continue
                    # Добавляем отступы для элементов группы
                    if current_group and not name.startswith(('    ', '\t')):
                        name = '    ' + name
                    # Создаем продукт
                    product = Product(
                        article=article,
                        name=name,
                        quantity=quantity,
                        unit=unit,
                        supplier_price=price,
                        is_group_header=False
                    )
                    products.append(product)
                    logger.info(f"Импортирован продукт: {product.name}")
                except Exception as e:
                    logger.error(f"Ошибка при обработке строки {index + 1}: {str(e)}")
                    continue
                progress_bar["value"] = index + 1
                root.update_idletasks()
            
            progress_bar.destroy()
            return products
        except Exception as e:
            logger.error(f"Ошибка при импорте Excel: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка при импорте Excel: {str(e)}")
            raise ValueError(f"Ошибка при импорте Excel: {str(e)}")

    @staticmethod
    def import_pdf(file_path: Path) -> List[Product]:
        """
        Импорт продуктов из PDF.
        
        Args:
            file_path: Путь к PDF файлу.
        
        Returns:
            Список импортированных продуктов.
        """
        products = []
        logger.info(f"Импорт из PDF: {file_path}")
        try:
            with pdfplumber.open(file_path) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    tables = page.extract_tables()
                    logger.debug(f"Страница {page_num}: найдено таблиц {len(tables)}")
                    for table_num, table in enumerate(tables, start=1):
                        logger.debug(f"Страница {page_num}, Таблица {table_num}: строки {len(table)}")
                        for row_num, row in enumerate(table, start=1):
                            try:
                                # Пропускаем заголовки и пустые строки
                                if not row or 'наименование' in str(row[0]).lower():
                                    logger.debug(f"Пропущена строка {row_num} в таблице {table_num} страницы {page_num}")
                                    continue
                                # Проверка наличия достаточного количества столбцов
                                if len(row) < 4:
                                    logger.warning(f"Недостаточно столбцов в строке {row_num} таблицы {table_num} страницы {page_num}")
                                    continue
                                article = str(row[0]).strip()
                                name = str(row[1]).strip()
                                if not name:
                                    logger.debug(f"Пропущена строка с пустым названием на строке {row_num} таблицы {table_num} страницы {page_num}")
                                    continue
                                quantity = ImportService._extract_quantity(row[2])
                                price = ImportService._extract_price(row[3])
                                if price > 0:
                                    product = Product(
                                        article=article,
                                        name=name,
                                        quantity=quantity,
                                        unit='шт.',  # Предполагается, что единица измерения - штуки
                                        supplier_price=price,
                                        is_group_header=False
                                    )
                                    products.append(product)
                                    logger.debug(f"Импортирован продукт из PDF: {product}")
                            except (ValueError, InvalidOperation) as e:
                                logger.warning(f"Ошибка при импорте строки {row_num} таблицы {table_num} страницы {page_num}: {e}")
                                continue
        except FileNotFoundError:
            logger.error(f"Файл не найден: {file_path}")
            messagebox.showerror("Ошибка", f"Файл не найден: {file_path}")
            raise
        except Exception as e:
            logger.error(f"Ошибка при импорте из PDF: {e}")
            messagebox.showerror("Ошибка", f"Ошибка при импорте из PDF: {e}")
            raise
        logger.info(f"Всего импортировано продуктов из PDF: {len(products)}")
        return products

    @staticmethod
    def _extract_quantity(value: str) -> int:
        """
        Извлечение количества из строки.
        
        Args:
            value: Строковое представление количества.
        
        Returns:
            Количество как целое число.
        """
        try:
            # Удаляем пробелы и символы разделителей тысяч
            qty_str = ''.join(filter(lambda x: x.isdigit() or x == '.', str(value)))
            quantity = int(float(qty_str)) if qty_str else 1
            logger.debug(f"Извлечено количество: {quantity} из значения: {value}")
            return quantity
        except Exception as e:
            logger.warning(f"Не удалось извлечь количество из значения '{value}': {e}")
            return 1

    @staticmethod
    def _extract_price(value: str) -> Decimal:
        """
        Извлечение цены из строки.
        
        Args:
            value: Строковое представление цены.
        
        Returns:
            Цена как Decimal.
        """
        try:
            # Удаляем пробелы и символы разделителей тысяч
            price_str = ''.join(filter(lambda x: x.isdigit() or x in '.,', str(value)))
            price_str = price_str.replace(',', '.')
            price = Decimal(price_str) if price_str else Decimal('0.00')
            logger.debug(f"Извлечена цена: {price} из значения: {value}")
            return price
        except Exception as e:
            logger.warning(f"Не удалось извлечь цену из значения '{value}': {e}")
            return Decimal('0.00')

    @staticmethod
    def _validate_quantity(value) -> int:
        """
        Валидация и преобразование количества.
        
        Args:
            value: Входное значение количества.
        
        Returns:
            Валидированное количество.
        """
        try:
            quantity = int(value)
            if quantity < 0:
                logger.warning(f"Отрицательное количество: {quantity}. Установлено значение 1.")
                return 1
            return quantity
        except (ValueError, TypeError) as e:
            logger.warning(f"Неверный формат количества '{value}': {e}. Установлено значение 1.")
            return 1

    @staticmethod
    def _validate_price(value) -> Decimal:
        """
        Валидация и преобразование цены.
        
        Args:
            value: Входное значение цены.
        
        Returns:
            Валидированная цена.
        """
        try:
            price = Decimal(value)
            if price < 0:
                logger.warning(f"Отрицательная цена: {price}. Установлено значение 0.")
                return Decimal('0.00')
            return price
        except (InvalidOperation, ValueError, TypeError) as e:
            logger.warning(f"Неверный формат цены '{value}': {e}. Установлено значение 0.")
            return Decimal('0.00')

    @staticmethod
    def auto_detect_columns(file_path: Path) -> Dict[str, str]:
        """
        Автоматическое обнаружение столбцов в Excel файле.
        
        Args:
            file_path: Путь к Excel файлу.
        
        Returns:
            Словарь с предложенными маппингами столбцов.
        """
        try:
            df = pd.read_excel(file_path)
            columns = list(df.columns)
            suggested_mapping = {}
            for col in columns:
                col_lower = col.lower()
                if 'наименование' in col_lower:
                    suggested_mapping['name'] = col
                elif 'арт.' in col_lower or 'артикул' in col_lower:
                    suggested_mapping['article'] = col
                elif 'кол.' in col_lower or 'количество' in col_lower:
                    suggested_mapping['quantity'] = col
                elif 'ед.' in col_lower or 'единица' in col_lower:
                    suggested_mapping['unit'] = col
                elif 'цена' in col_lower:
                    suggested_mapping['price'] = col
            return suggested_mapping
        except Exception as e:
            logger.error(f"Ошибка при автодетекции столбцов: {str(e)}")
            raise ValueError(f"Ошибка при автодетекции столбцов: {str(e)}")
    
    @staticmethod
    def preview_excel(file_path: Path) -> List[List[str]]:
        """
        Предпросмотр первых нескольких строк Excel файла.
        
        Args:
            file_path: Путь к Excel файлу.
        
        Returns:
            Список списков значений первых нескольких строк.
        """
        try:
            df = pd.read_excel(file_path)
            return df.head(10).values.tolist()
        except Exception as e:
            logger.error(f"Ошибка при предпросмотре Excel: {str(e)}")
            raise ValueError(f"Ошибка при предпросмотре Excel: {str(e)}")
    
    @staticmethod
    def _extract_quantity(value: str) -> int:
        """
        Извлечение количества из строки.
        
        Args:
            value: Строковое представление количества.
        
        Returns:
            Количество как целое число.
        """
        try:
            # Удаляем пробелы и символы разделителей тысяч
            qty_str = ''.join(filter(lambda x: x.isdigit() or x == '.', str(value)))
            quantity = int(float(qty_str)) if qty_str else 1
            logger.debug(f"Извлечено количество: {quantity} из значения: {value}")
            return quantity
        except Exception as e:
            logger.warning(f"Не удалось извлечь количество из значения '{value}': {e}")
            return 1

    @staticmethod
    def _extract_price(value: str) -> Decimal:
        """
        Извлечение цены из строки.
        
        Args:
            value: Строковое представление цены.
        
        Returns:
            Цена как Decimal.
        """
        try:
            # Удаляем пробелы и символы разделителей тысяч
            price_str = ''.join(filter(lambda x: x.isdigit() or x in '.,', str(value)))
            price_str = price_str.replace(',', '.')
            price = Decimal(price_str) if price_str else Decimal('0.00')
            logger.debug(f"Извлечена цена: {price} из значения: {value}")
            return price
        except Exception as e:
            logger.warning(f"Не удалось извлечь цену из значения '{value}': {e}")
            return Decimal('0.00')