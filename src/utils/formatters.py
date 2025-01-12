# src/utils/formatters.py

from decimal import Decimal
from typing import Union, Optional

class Formatters:
    @staticmethod
    def format_currency(value: Decimal, show_currency: bool = True) -> str:
        """
        Форматирование числа как валюты.
        
        Args:
            value: Сумма для форматирования
            show_currency: Показывать ли символ валюты
            
        Returns:
            Отформатированная строка с валютой
        """
        try:
            if value is None:
                return "0,00 ₽" if show_currency else "0,00"
                
            # Форматируем число с двумя десятичными знаками
            formatted = f"{value:,.2f}".replace(',', ' ').replace('.', ',')
            
            # Добавляем символ валюты если нужно
            return f"{formatted} ₽" if show_currency else formatted
        except Exception:
            return "0,00 ₽" if show_currency else "0,00"

    @staticmethod
    def format_quantity(value: int) -> str:
        """
        Форматирование количества.
        
        Args:
            value: Количество для форматирования
            
        Returns:
            Отформатированная строка с количеством
        """
        try:
            return f"{value:,}".replace(',', ' ')
        except Exception:
            return "0"

    @staticmethod
    def format_percentage(value: Decimal) -> str:
        """
        Форматирование процентов.
        
        Args:
            value: Процент для форматирования
            
        Returns:
            Отформатированная строка с процентами
        """
        try:
            if value is None:
                return "0,0%"
                
            return f"{value:.1f}%".replace('.', ',')
        except Exception:
            return "0,0%"

    @staticmethod
    def parse_currency(value: str) -> Optional[Decimal]:
        """
        Преобразование строки с валютой в Decimal.
        
        Args:
            value: Строка для преобразования
            
        Returns:
            Decimal или None в случае ошибки
        """
        try:
            # Убираем все пробелы, знак валюты и меняем разделители
            cleaned = value.replace(' ', '').replace('₽', '').replace(',', '.')
            return Decimal(cleaned) if cleaned else Decimal('0')
        except Exception:
            return None

    @staticmethod
    def parse_quantity(value: str) -> Optional[int]:
        """
        Преобразование строки с количеством в int.
        
        Args:
            value: Строка для преобразования
            
        Returns:
            int или None в случае ошибки
        """
        try:
            # Убираем все пробелы и преобразуем в число
            cleaned = value.replace(' ', '')
            return int(cleaned) if cleaned else 0
        except Exception:
            return None

    @staticmethod
    def parse_percentage(value: str) -> Optional[Decimal]:
        """
        Преобразование строки с процентами в Decimal.
        
        Args:
            value: Строка для преобразования
            
        Returns:
            Decimal или None в случае ошибки
        """
        try:
            # Убираем все пробелы, знак процента и меняем разделители
            cleaned = value.replace(' ', '').replace('%', '').replace(',', '.')
            return Decimal(cleaned) if cleaned else Decimal('0')
        except Exception:
            return None

    @staticmethod
    def strip_currency_symbol(value: str) -> str:
        """
        Удаление символа валюты из строки.
        
        Args:
            value: Строка с валютой
            
        Returns:
            Строка без символа валюты
        """
        return value.replace('₽', '').strip()