from decimal import Decimal
import re
from typing import List, Optional
import pdfplumber
import logging

from src.models.product import Product

logger = logging.getLogger(__name__)

class VingsMPDFHandler:
    """Специализированный обработчик PDF файлов от ВИНГС-М"""
    
    @staticmethod
    def can_handle(text: str) -> bool:
        """Проверяет, является ли PDF файлом от ВИНГС-М"""
        return "ЗАО \"ВИНГС-М\"" in text or "ВИНГС-М" in text
    
    @staticmethod
    def extract_products(pdf_path) -> List[Product]:
        """Извлекает продукты из PDF файла ВИНГС-М"""
        products = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text:
                        continue
                    
                    # Извлекаем таблицу с продуктами
                    tables = page.extract_tables()
                    if not tables:
                        continue
                        
                    for table in tables:
                        header_found = False
                        for row in table:
                            # Пропускаем пустые строки
                            if not row or not any(row):
                                continue
                                
                            # Ищем заголовок таблицы
                            if "Наименование" in str(row[0]):
                                header_found = True
                                continue
                                
                            if not header_found:
                                continue
                                
                            try:
                                # Проверяем наличие номера позиции
                                if not row[0] or not str(row[0]).strip().isdigit():
                                    continue
                                    
                                # Извлекаем данные
                                name = str(row[1]).strip() if len(row) > 1 else ""
                                if not name or name.lower() in ["наименование", "итого", "всего"]:
                                    continue
                                    
                                # Извлекаем артикул из наименования
                                article = ""
                                article_match = re.search(r'\((.*?)\)', name)
                                if article_match:
                                    article = article_match.group(1)
                                
                                # Извлекаем количество и единицу измерения
                                quantity = 0
                                unit = "шт."
                                if len(row) > 2:
                                    qty_str = str(row[2]).strip()
                                    qty_match = re.match(r'(\d+)\s*(\w+)?', qty_str)
                                    if qty_match:
                                        quantity = int(qty_match.group(1))
                                        if qty_match.group(2):
                                            unit = qty_match.group(2)
                                
                                # Извлекаем цену
                                price = Decimal('0')
                                if len(row) > 5:  # Используем цену без НДС
                                    price_str = str(row[5]).strip()
                                    price_str = re.sub(r'[^\d.,]', '', price_str)
                                    price_str = price_str.replace(',', '.')
                                    if price_str:
                                        price = Decimal(price_str)
                                
                                # Создаем продукт
                                if name and quantity > 0:
                                    product = Product(
                                        article=article,
                                        name=name,
                                        quantity=quantity,
                                        unit=unit,
                                        supplier_price=price,
                                        is_group_header=False
                                    )
                                    products.append(product)
                                    logger.debug(f"Импортирован продукт: {product}")
                                
                            except Exception as e:
                                logger.warning(f"Ошибка при обработке строки: {str(e)}")
                                continue
        
        except Exception as e:
            logger.error(f"Ошибка при импорте PDF: {str(e)}")
            raise
        
        return products

class PDFImportService:
    """Сервис для импорта PDF файлов с поддержкой разных форматов"""
    
    @staticmethod
    def import_pdf(pdf_path) -> List[Product]:
        """
        Импортирует продукты из PDF файла с автоматическим определением формата
        """
        try:
            # Читаем первую страницу для определения формата
            with pdfplumber.open(pdf_path) as pdf:
                first_page_text = pdf.pages[0].extract_text() or ""
                
                # Определяем обработчик
                if VingsMPDFHandler.can_handle(first_page_text):
                    logger.info("Определен формат ВИНГС-М")
                    return VingsMPDFHandler.extract_products(pdf_path)
                else:
                    # Используем стандартный импорт для неизвестных форматов
                    logger.info("Используется стандартный импорт PDF")
                    from src.services.import_service import ImportService
                    return ImportService.import_pdf(pdf_path)
                    
        except Exception as e:
            logger.error(f"Ошибка при импорте PDF: {str(e)}")
            raise