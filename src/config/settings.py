# src/config/settings.py

from pathlib import Path

# Базовые пути
BASE_DIR = Path(__file__).resolve().parent.parent.parent
ASSETS_DIR = BASE_DIR / "assets"
TEMPLATES_DIR = ASSETS_DIR / "templates"

# Настройки приложения
APP_NAME = "Аэропро. Создание КП"
APP_VERSION = "2.0.0"
DEFAULT_WINDOW_SIZE = "1200x800"

# Настройки по умолчанию
DEFAULT_VAT = 20  # НДС 20%
DEFAULT_MARKUP = 0  # Наценка 0%
DEFAULT_DISCOUNT = 0  # Скидка 0%
DEFAULT_UNIT = "шт."  # Единица измерения по умолчанию

# Форматирование
PRICE_DECIMAL_PLACES = 2
QUANTITY_DECIMAL_PLACES = 0

# Настройки файлов
SUPPORTED_IMPORT_FORMATS = [
    ("Excel Files", "*.xlsx *.xls"),
    ("PDF Files", "*.pdf"),
    ("All Files", "*.*")
]

SUPPORTED_EXPORT_FORMATS = [
    ("Word Files", "*.docx"),
    ("Excel Files", "*.xlsx"),
    ("All Files", "*.*")
]

# Языковые настройки
DEFAULT_LOCALE = 'ru_RU.UTF-8'
FALLBACK_LOCALE = 'Russian_Russia.1251'

# PDF Export Settings
PDF_FONT_DIR = BASE_DIR / "assets" / "fonts"
PDF_DEFAULT_FONT = "Arial"
PDF_HEADING_FONT = "Arial-Bold"
PDF_FONT_SIZE = 10
PDF_HEADING_SIZE = 12

# Цвета для PDF
PDF_HEADER_COLOR = "#283C5C"
PDF_ALTERNATE_ROW_COLOR = "#F8F8F8"
PDF_GROUP_HEADER_COLOR = "#F0F0F0"

# Отступы для PDF (в сантиметрах)
PDF_MARGIN_LEFT = 2
PDF_MARGIN_RIGHT = 2
PDF_MARGIN_TOP = 2
PDF_MARGIN_BOTTOM = 2