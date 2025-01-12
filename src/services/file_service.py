# src/services/file_service.py
import json
from pathlib import Path
from typing import Optional
from tkinter import filedialog
from datetime import datetime
from decimal import Decimal, InvalidOperation
from src.models.commercial_offer import CommercialOffer
from src.models.product import Product
import logging

# Настройка логирования
logger = logging.getLogger(__name__)

class FileService:
    def __init__(self, base_path: Path):
        self.base_path = base_path
        self.offers_path = self.base_path / "offers"
        # Убедимся, что директория существует
        self.offers_path.mkdir(parents=True, exist_ok=True)

    def save_offer(self, offer: CommercialOffer, filename: str, custom_path: Optional[Path] = None) -> None:
        """
        Сохранение коммерческого предложения в JSON файл.
        
        Args:
            offer: Объект CommercialOffer для сохранения.
            filename: Имя файла без расширения.
            custom_path: Необязательный путь для сохранения файла.
        """
        try:
            data = {
                'number': offer.number,
                'date': offer.date.strftime("%d.%m.%Y") if offer.date else None,
                'discount_from_supplier': str(offer.discount_from_supplier),
                'markup_for_client': str(offer.markup_for_client),
                'vat': str(offer.vat),
                'delivery_terms': offer.delivery_terms,
                'self_pickup_warehouse': offer.self_pickup_warehouse,
                'warranty': offer.warranty,
                'delivery_time': offer.delivery_time,
                'include_delivery': offer.include_delivery,
                'delivery_cost': str(offer.delivery_cost),
                'products': [
                    {
                        'article': p.article,
                        'name': p.name,
                        'quantity': p.quantity,
                        'unit': p.unit,
                        'supplier_price': str(p.supplier_price),
                        'markup': str(p.markup) if p.markup is not None else None,  # Обработка markup
                        'client_price': str(p.client_price),
                        'total_price': str(p.total_price),
                        'is_group_header': p.is_group_header
                    } for p in offer.products
                ]
            }
            save_path = custom_path if custom_path else (self.offers_path / f"{filename}.json")
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            logger.info(f"Коммерческое предложение сохранено в {save_path}")
        except Exception as e:
            logger.error(f"Ошибка при сохранении коммерческого предложения: {str(e)}")
            raise

    def load_offer(self, filename: str) -> Optional[CommercialOffer]:
        """
        Загрузка коммерческого предложения из JSON файла.
        
        Args:
            filename: Имя файла без расширения или полный путь к файлу.
        
        Returns:
            Объект CommercialOffer или None в случае ошибки.
        """
        try:
            # Попытка загрузить из стандартного пути
            file_path = self.offers_path / f"{filename}.json"
            if not file_path.exists():
                # Если файл не найден, запрашиваем путь у пользователя
                logger.warning(f"Файл не найден по пути: {file_path}")
                file_path = filedialog.askopenfilename(
                    title="Выберите файл коммерческого предложения",
                    filetypes=[("JSON Files", "*.json")]
                )
                if not file_path:
                    logger.warning("Пользователь не выбрал файл.")
                    return None
                file_path = Path(file_path)
            
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Преобразование строки даты в datetime
            date = None
            if data.get('date'):
                try:
                    date = datetime.strptime(data['date'], "%d.%m.%Y")
                except ValueError:
                    logger.warning("Неверный формат даты, используем текущую дату.")
                    date = datetime.now()
            
            # Создаём список продуктов
            products = []
            for p_data in data.get('products', []):
                try:
                    product = Product(
                        article=p_data.get('article', ''),
                        name=p_data.get('name', ''),
                        quantity=int(p_data.get('quantity', 0)),
                        unit=p_data.get('unit', 'шт.'),
                        supplier_price=Decimal(p_data.get('supplier_price', '0')),
                        markup=Decimal(p_data.get('markup')) if p_data.get('markup') is not None else None,  # Обработка markup
                        is_group_header=p_data.get('is_group_header', False)
                    )
                    # Логгирование
                    logger.info(f"Загружен продукт: {product.article}, Markup: {product.markup}")
                    # Устанавливаем дополнительные поля
                    if 'client_price' in p_data:
                        product.client_price = Decimal(str(p_data['client_price']))
                    if 'total_price' in p_data:
                        product.total_price = Decimal(str(p_data['total_price']))
                    products.append(product)
                except (ValueError, InvalidOperation) as e:
                    logger.warning(f"Ошибка при обработке продукта: {str(e)}")
                    continue
            
            # Возвращаем объект КП
            return CommercialOffer(
                number=data.get('number', ''),
                date=date,
                discount_from_supplier=Decimal(str(data.get('discount_from_supplier', '0'))),
                markup_for_client=Decimal(str(data.get('markup_for_client', '0'))),
                vat=Decimal(str(data.get('vat', '20'))),
                delivery_terms=data.get('delivery_terms', ''),
                self_pickup_warehouse=data.get('self_pickup_warehouse', ''),
                warranty=data.get('warranty', ''),
                delivery_time=data.get('delivery_time', ''),
                include_delivery=data.get('include_delivery', False),
                delivery_cost=Decimal(str(data.get('delivery_cost', '0'))),
                products=products
            )
        except Exception as e:
            logger.error(f"Ошибка при загрузке коммерческого предложения: {str(e)}")
            raise ValueError(f"Не удалось загрузить коммерческое предложение: {str(e)}")