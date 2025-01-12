import logging
from pathlib import Path
from decimal import Decimal
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from src.models.commercial_offer import CommercialOffer
from src.utils.formatters import Formatters
from src.config import settings

logger = logging.getLogger(__name__)

class PDFExporter:
    def __init__(self):
        # Регистрируем шрифты
        self._register_fonts()
        
        # Создаем стили
        self.styles = getSampleStyleSheet()
        self.styles.add(ParagraphStyle(
            name='CustomNormal',
            fontName=settings.PDF_DEFAULT_FONT,
            fontSize=settings.PDF_FONT_SIZE,
            leading=settings.PDF_FONT_SIZE * 1.2,
        ))
        self.styles.add(ParagraphStyle(
            name='CustomHeading',
            fontName=settings.PDF_HEADING_FONT,
            fontSize=settings.PDF_HEADING_SIZE,
            leading=settings.PDF_HEADING_SIZE * 1.2,
            alignment=1,
        ))

    def _register_fonts(self):
        """Register required fonts for PDF generation."""
        try:
            font_dir = settings.PDF_FONT_DIR
            pdfmetrics.registerFont(TTFont('Arial', font_dir / 'arial.ttf'))
            pdfmetrics.registerFont(TTFont('Arial-Bold', font_dir / 'arialbd.ttf'))
        except Exception as e:
            logger.warning(f"Не удалось загрузить шрифты Arial: {e}")
            # Используем встроенные шрифты если не удалось загрузить Arial
            self.styles = getSampleStyleSheet()

    def export_pdf(self, offer: CommercialOffer, output_path: Path) -> None:
        """Export commercial offer to PDF."""
        try:
            # Create PDF document
            doc = SimpleDocTemplate(
                str(output_path),
                pagesize=A4,
                rightMargin=2*cm,
                leftMargin=2*cm,
                topMargin=2*cm,
                bottomMargin=2*cm
            )
            
            # Collect all elements
            elements = []
            
            # Add header
            elements.append(Paragraph(f"КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ № {offer.number}", self.styles['CustomHeading']))
            elements.append(Spacer(1, 0.5*cm))
            
            if offer.date:
                elements.append(Paragraph(f"Дата: {offer.date.strftime('%d.%m.%Y')}", self.styles['CustomNormal']))
            
            elements.append(Spacer(1, 1*cm))
            
            # Create table data
            data = [['№', 'Наименование продукции', 'Кол-во', 'Ед.изм.', 'Цена с НДС', 'Всего с НДС']]
            
            item_number = 1
            for product in offer.products:
                if product.is_header:
                    # Add group header
                    data.append(['', product.name, '', '', '', ''])
                else:
                    # Add product
                    data.append([
                        str(item_number),
                        product.name,
                        str(product.quantity),
                        product.unit,
                        Formatters.format_currency(product.client_price, False),
                        Formatters.format_currency(product.total_price, False)
                    ])
                    item_number += 1
            
            # Create table
            table = Table(data, colWidths=[1*cm, 8*cm, 2*cm, 2*cm, 3*cm, 3*cm])
            
            # Table style
            table_style = TableStyle([
                # Headers
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#283C5C')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Arial-Bold'),
                
                # Grid
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                
                # Alignment
                ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # №
                ('ALIGN', (2, 1), (3, -1), 'CENTER'),  # Количество, единицы
                ('ALIGN', (4, 1), (5, -1), 'RIGHT'),   # Цены
            ])
            
            # Add alternating row colors
            for i in range(1, len(data)):
                if i % 2 == 0:
                    table_style.add('BACKGROUND', (0, i), (-1, i), colors.HexColor('#F8F8F8'))
            
            table.setStyle(table_style)
            elements.append(table)
            
            # Add totals
            elements.append(Spacer(1, 1*cm))
            
            total_amount = sum(p.total_price for p in offer.products if not p.is_header)
            vat_amount = (total_amount * Decimal(str(offer.vat)) / (Decimal('100') + Decimal(str(offer.vat)))).quantize(Decimal('0.01'))
            
            elements.append(Paragraph(
                f"Итого: {Formatters.format_currency(total_amount)}",
                self.styles['CustomNormal']
            ))
            elements.append(Paragraph(
                f"В том числе НДС ({offer.vat}%): {Formatters.format_currency(vat_amount)}",
                self.styles['CustomNormal']
            ))
            
            # Add terms
            elements.append(Spacer(1, 1*cm))
            elements.append(Paragraph("Условия поставки:", self.styles['CustomHeading']))
            elements.append(Spacer(1, 0.5*cm))
            
            terms = [
                f"Срок поставки: {offer.delivery_time or '3-4 недели'}",
                f"Склад самовывоза: {offer.self_pickup_warehouse or 'МО, г. Люберцы, ул. Красная д 1, лит. С'}",
                f"Гарантия: {offer.warranty or '24 мес.'}",
                "Условия оплаты: 100% предоплата"
            ]
            
            for term in terms:
                elements.append(Paragraph(term, self.styles['CustomNormal']))
                elements.append(Spacer(1, 0.3*cm))
            
            # Build PDF
            doc.build(elements)
            logger.info(f"PDF файл успешно создан: {output_path}")
            
        except Exception as e:
            logger.error(f"Ошибка при создании PDF: {str(e)}")
            raise