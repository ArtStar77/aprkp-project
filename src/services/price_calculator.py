# src/services/price_calculator.py
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from typing import List
from src.models.product import Product
import logging

logger = logging.getLogger(__name__)

class PriceCalculator:
    def calculate_client_price(self, product: Product, discount_from_supplier: Decimal, markup_for_client: Decimal) -> None:
        """
        Рассчитывает цену клиента на основе цены поставщика, скидки и наценки.
        Args:
            product: Продукт, для которого рассчитывается цена.
            discount_from_supplier: Скидка от поставщика в процентах.
            markup_for_client: Наценка для клиента в процентах.
        """
        try:
            if product.is_group_header:
                return
            # Логгирование
            logger.info(f"Calculating price for Product: {product.article}, Markup: {product.markup}, Global Markup: {markup_for_client}")
            # Переводим проценты в десятичные дроби
            discount_rate = discount_from_supplier / Decimal('100')
            markup_rate = product.markup / Decimal('100') if product.markup is not None else markup_for_client / Decimal('100')
            # Применяем скидку
            discounted_price = product.supplier_price * (Decimal('1') - discount_rate)
            # Применяем наценку
            product.client_price = discounted_price * (Decimal('1') + markup_rate)
            # Рассчитываем общую сумму
            product.total_price = product.client_price * Decimal(product.quantity)
            # Округляем
            product.client_price = product.client_price.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            product.total_price = product.total_price.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        except (InvalidOperation, TypeError, ValueError) as e:
            logger.error(f"Ошибка при расчете цены клиента: {str(e)}")
            raise ValueError(f"Ошибка при расчете цены клиента: {str(e)}")

    def calculate_totals(
        self,
        products: List[Product],
        discount_from_supplier: Decimal,
        markup_for_client: Decimal,
        vat: Decimal
    ) -> dict:
        """
        Рассчитывает итоговые суммы.
        Args:
            products: Список продуктов
            discount_from_supplier: Скидка от поставщика в процентах
            markup_for_client: Наценка для клиента в процентах
            vat: НДС в процентах
        Returns:
            Словарь с итоговыми значениями
        """
        try:
            # Начальные значения для сумм
            total_amount = Decimal('0')  # Общая сумма с учетом всех скидок и наценок
            supplier_total = Decimal('0')  # Сумма по ценам поставщика
            discounted_total = Decimal('0')  # Сумма со скидкой от поставщика

            # Переводим проценты в десятичные дроби
            discount_rate = discount_from_supplier / Decimal('100')
            markup_rate = markup_for_client / Decimal('100')
            vat_rate = vat / Decimal('100')

            # Проходим по всем продуктам
            for product in products:
                if product.is_group_header:
                    continue
                quantity = Decimal(product.quantity)
                # Считаем суммы
                supplier_amount = product.supplier_price * quantity
                discounted_amount = supplier_amount * (Decimal('1') - discount_rate)
                final_amount = discounted_amount * (Decimal('1') + markup_rate)
                supplier_total += supplier_amount
                discounted_total += discounted_amount
                total_amount += final_amount

            # Округляем все суммы
            supplier_total = supplier_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            discounted_total = discounted_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            total_amount = total_amount.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

            # Расчет НДС
            vat_amount = (total_amount * vat_rate / (Decimal('1') + vat_rate)).quantize(Decimal('0.01'))

            # Расчет маржи
            margin = total_amount - discounted_total
            margin_percentage = (margin / discounted_total * Decimal('100')).quantize(Decimal('0.1')) if discounted_total else Decimal('0')

            # Расчет экономии клиента
            savings = supplier_total - total_amount
            savings_percentage = (savings / supplier_total * Decimal('100')).quantize(Decimal('0.1')) if supplier_total else Decimal('0')

            return {
                'total_amount': total_amount,
                'vat_amount': vat_amount,
                'margin': margin,
                'margin_percentage': margin_percentage,
                'client_savings': savings,
                'client_savings_percentage': savings_percentage,
                'supplier_total': supplier_total,
                'discounted_total': discounted_total
            }
        except (InvalidOperation, TypeError) as e:
            logger.error(f"Ошибка при расчете итогов: {str(e)}")
            raise ValueError(f"Ошибка при расчете итогов: {str(e)}")

    def update_prices(
        self,
        products: List[Product],
        discount_from_supplier: Decimal,
        markup_for_client: Decimal
    ) -> None:
        """
        Обновляет цены всех продуктов в списке на основе скидки и наценки.
        Args:
            products: Список продуктов для обновления цен
            discount_from_supplier: Скидка от поставщика в процентах
            markup_for_client: Наценка для клиента в процентах
        """
        try:
            for product in products:
                if not product.is_group_header:
                    self.calculate_client_price(product, discount_from_supplier, markup_for_client)
        except Exception as e:
            logger.error(f"Ошибка при обновлении цен: {str(e)}")
            raise ValueError(f"Ошибка при обновлении цен: {str(e)}")