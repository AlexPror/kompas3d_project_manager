#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для определения типа PDF (векторный или растровый)
"""

import sys
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    print("❌ Установите PyMuPDF: pip install PyMuPDF")
    sys.exit(1)

def check_pdf_type(pdf_path):
    """Проверяет, векторный PDF или растровый"""
    
    print("="*60)
    print("АНАЛИЗ PDF")
    print("="*60)
    print(f"Файл: {pdf_path}\n")
    
    try:
        doc = fitz.open(pdf_path)
        
        print(f"Страниц: {len(doc)}")
        print(f"Размер файла: {Path(pdf_path).stat().st_size / 1024:.1f} КБ\n")
        
        # Анализируем первую страницу
        page = doc[0]
        
        # Получаем векторные объекты
        paths = page.get_drawings()
        text_instances = page.get_text("dict")
        images = page.get_images()
        
        print("СОДЕРЖИМОЕ СТРАНИЦЫ 1:")
        print(f"  Векторных объектов (линии, дуги): {len(paths)}")
        print(f"  Текстовых блоков: {len(text_instances.get('blocks', []))}")
        print(f"  Изображений: {len(images)}\n")
        
        # Определяем тип
        is_vector = len(paths) > 10  # Если много векторных объектов
        is_raster = len(images) > 0 and len(paths) < 10  # Если есть изображения и мало векторов
        
        print("="*60)
        if is_vector:
            print("✅ ВЕКТОРНЫЙ PDF")
            print("   Содержит векторную графику (линии, дуги)")
            print("   Можно извлечь координаты отверстий и размеры")
        elif is_raster:
            print("⚠️ РАСТРОВЫЙ PDF")
            print("   Содержит изображения (сканы)")
            print("   Потребуется OCR и распознавание образов")
        else:
            print("❓ СМЕШАННЫЙ ИЛИ ПУСТОЙ PDF")
            print("   Проверьте содержимое вручную")
        print("="*60)
        
        # Дополнительная информация
        if len(paths) > 0:
            print(f"\nПример векторного объекта:")
            print(f"  Тип: {paths[0].get('type', 'unknown')}")
            print(f"  Точек: {len(paths[0].get('items', []))}")
        
        doc.close()
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")

if __name__ == "__main__":
    # Укажите путь к вашему PDF
    pdf_path = r"C:\Users\Vorob\Documents\ZVD GROUP\Заказы\Дарим Детям Добро А-191125-2\02_кд\ZVD.TURBO.80.230.2000\ZVD.TURBO.75.230.2000 А-191125-2 (ОЦИНКОВКА)\0010 - Корпус короба.pdf"
    
    if not Path(pdf_path).exists():
        print(f"❌ Файл не найден: {pdf_path}")
        print("\nУкажите путь к PDF файлу в скрипте (строка 50)")
    else:
        check_pdf_type(pdf_path)
