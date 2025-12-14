#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для чтения содержимого ячейки материала из чертежа КОМПАС-3D
"""

import sys
import pythoncom
from pathlib import Path

# Добавляем путь к компонентам
sys.path.insert(0, str(Path(__file__).parent))

from components.base_component import get_dynamic_dispatch

def read_material_cell(drawing_path):
    """Читает содержимое ячейки материала (ячейка 3) из чертежа"""
    
    pythoncom.CoInitialize()
    
    # Открываем файл для записи результатов
    output_file = Path(__file__).parent / "material_analysis.txt"
    
    def log(message):
        """Выводит сообщение в консоль и файл"""
        print(message)
        with open(output_file, 'a', encoding='utf-8') as f:
            f.write(message + '\n')
    
    # Очищаем файл
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('')
    
    try:
        log("="*60)
        log("ЧТЕНИЕ ЯЧЕЙКИ МАТЕРИАЛА ИЗ ЧЕРТЕЖА")
        log("="*60)
        log(f"Файл: {drawing_path}\n")
        
        # Подключаемся к КОМПАС-3D
        api7 = get_dynamic_dispatch("Kompas.Application.7")
        
        # Открываем чертеж
        log("Открытие чертежа...")
        doc = api7.Documents.Open(str(drawing_path), False, False)
        
        if not doc:
            log("❌ Не удалось открыть файл")
            return
        
        log("✓ Чертеж открыт\n")
        
        # Получаем активный документ
        kompas_document_2d = api7.ActiveDocument
        
        # Получаем штамп
        layout_sheets = kompas_document_2d.LayoutSheets
        sheet = layout_sheets.Item(0)
        stamp = sheet.Stamp
        
        if not stamp:
            log("❌ Штамп не найден")
            return
        
        log("✓ Штамп найден\n")
        
        # Читаем ячейку 3 (материал)
        log("Чтение ячейки 3 (Материал):")
        log("-"*60)
        
        try:
            text_item = stamp.Text(3)
            material_text = text_item.Str
            
            log(f"Содержимое: '{material_text}'")
            log(f"Длина: {len(material_text)} символов")
            log(f"\nПосимвольный разбор:")
            for i, char in enumerate(material_text):
                log(f"  [{i}] = '{char}' (код: {ord(char)})")
            
            # Проверяем на спецсимволы
            log(f"\nПроверка спецсимволов:")
            if '$' in material_text:
                log("  ✓ Найден символ '$'")
            if '\n' in material_text:
                log("  ✓ Найден символ переноса строки '\\n'")
            if '\r' in material_text:
                log("  ✓ Найден символ возврата каретки '\\r'")
            
            # Представление для Python кода
            log(f"\nДля использования в Python:")
            log(f'  "{repr(material_text)[1:-1]}"')
                
        except Exception as e:
            log(f"❌ Ошибка чтения ячейки: {e}")
        
        log("\n" + "="*60)
        log(f"Результаты сохранены в: {output_file}")
        log("="*60)
        
        # Закрываем документ
        api7.ActiveDocument.Close(False)
        log("Чертеж закрыт")
        
    except Exception as e:
        log(f"❌ Ошибка: {e}")
        import traceback
        log(traceback.format_exc())
    
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    drawing_path = r"C:\Users\Vorob\Documents\ZVD GROUP\Заказы\Дарим Детям Добро А-191125-2\02_кд\ZVD.TURBO.80.230.2000\0010 - Корпус короба.cdw"
    read_material_cell(drawing_path)
