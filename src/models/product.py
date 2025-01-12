# src/models/product.py
from typing import Optional
from decimal import Decimal, InvalidOperation
import logging

# Настройка логирования
logger = logging.getLogger(__name__)

class Product:
    def __init__(self, article: str, name: str, quantity: int, unit: str, supplier_price: Decimal, markup: Optional[Decimal] = None, is_group_header: bool = False):
        self.article = article
        self.name = name
        self.quantity = quantity
        self.unit = unit
        self.supplier_price = supplier_price
        self.markup = markup if markup is not None else Decimal('0')  # Устанавливаем наценку по умолчанию 0
        self.is_group_header = is_group_header
        # Логгирование
        logger.info(f"Initialized Product: {self.article}, Markup: {self.markup}")

    @property
    def is_header(self) -> bool:
        """Проверяет, является ли продукт заголовком группы."""
        return self.is_group_header

    @property
    def client_price(self) -> Decimal:
        """Рассчитывает цену для клиента."""
        try:
            return self.supplier_price * (Decimal('1') + (self.markup / Decimal('100')))
        except InvalidOperation as e:
            logger.error(f"Ошибка при расчете client_price для продукта {self.article}: {str(e)}")
            raise ValueError(f"Ошибка при расчете client_price для продукта {self.article}: {str(e)}")

    @property
    def total_price(self) -> Decimal:
        """Рассчитывает общую стоимость позиции."""
        if self.is_header:
            return Decimal('0')  # Для заголовков групп возвращаем 0
        try:
            return self.client_price * Decimal(self.quantity)
        except InvalidOperation as e:
            logger.error(f"Ошибка при расчете total_price для продукта {self.article}: {str(e)}")
            raise ValueError(f"Ошибка при расчете total_price для продукта {self.article}: {str(e)}")

    def calculate_total(self) -> None:
        """Пересчитывает общую стоимость позиции (метод для вызова извне, если требуется)."""
        if not self.is_header:
            logger.info(f"Calculating total for Product: {self.article}")
            try:
                self.total_price = self.client_price * Decimal(self.quantity)
            except InvalidOperation as e:
                logger.error(f"Ошибка при пересчете total_price для продукта {self.article}: {str(e)}")
                raise ValueError(f"Ошибка при пересчете total_price для продукта {self.article}: {str(e)}")

    def __str__(self) -> str:
        if self.is_header:
            return f"== {self.name} =="  # Для заголовков групп
        return f"{self.name} ({self.article}) - {self.quantity} {self.unit}"