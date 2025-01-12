# src/models/commercial_offer.py

from decimal import Decimal
from typing import List, Optional
from datetime import datetime

from .product import Product

class CommercialOffer:
    def __init__(
        self,
        number: str,
        date: Optional[datetime],
        discount_from_supplier: Decimal,
        markup_for_client: Decimal,
        vat: Decimal,
        delivery_terms: str = "",
        self_pickup_warehouse: str = "",
        warranty: str = "",
        delivery_time: str = "",
        include_delivery: bool = False,
        delivery_cost: Decimal = Decimal('0'),
        products: Optional[List[Product]] = None
    ):
        self.number = number
        self.date = date
        self.discount_from_supplier = Decimal(str(discount_from_supplier))
        self.markup_for_client = Decimal(str(markup_for_client))
        self.vat = Decimal(str(vat))
        self.delivery_terms = delivery_terms
        self.self_pickup_warehouse = self_pickup_warehouse
        self.warranty = warranty
        self.delivery_time = delivery_time
        self.include_delivery = include_delivery
        self.delivery_cost = Decimal(str(delivery_cost))
        self.products = products or []

    @property
    def total_amount(self) -> Decimal:
        total = sum(p.total_price for p in self.products)
        return Decimal(str(total)).quantize(Decimal('0.01'))

    @property
    def vat_amount(self) -> Decimal:
        if self.vat:
            vat_rate = self.vat / Decimal('100')
            return (self.total_amount * vat_rate).quantize(Decimal('0.01'))
        return Decimal('0')
