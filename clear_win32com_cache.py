#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для очистки кэша win32com.

Используйте этот скрипт если у вас возникает ошибка:
  AttributeError: module 'win32com.gen_py.XXX' has no attribute 'CLSIDToClassMap'

Запуск:
  python clear_win32com_cache.py
"""

import os
import shutil
import sys


def clear_win32com_cache(full_clear=False):
    """
    Очистка кэша win32com
    
    Args:
        full_clear: Если True - удаляет ВЕСЬ кэш gen_py
                   Если False - удаляет только KOMPAS-связанные файлы
    """
    try:
        import win32com
        gen_py_path = os.path.join(os.path.dirname(win32com.__file__), 'gen_py')
        
        if not os.path.exists(gen_py_path):
            print(f"✅ Кэш не найден: {gen_py_path}")
            return True
        
        print(f"📁 Кэш найден: {gen_py_path}")
        
        deleted_count = 0
        
        if full_clear:
            # Полная очистка
            for item in os.listdir(gen_py_path):
                item_path = os.path.join(gen_py_path, item)
                if item == '__init__.py':
                    continue  # Не удаляем __init__.py
                try:
                    if os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                        deleted_count += 1
                        print(f"  🗑️ Удалена папка: {item}")
                    elif os.path.isfile(item_path):
                        os.remove(item_path)
                        deleted_count += 1
                        print(f"  🗑️ Удален файл: {item}")
                except Exception as e:
                    print(f"  ⚠️ Не удалось удалить {item}: {e}")
        else:
            # Выборочная очистка (только KOMPAS)
            kompas_guids = ['0422828C', '2CAF168C']  # Известные GUID КОМПАСа
            
            for item in os.listdir(gen_py_path):
                item_path = os.path.join(gen_py_path, item)
                
                # Проверяем, относится ли к KOMPAS
                is_kompas = any(guid in item.upper() for guid in kompas_guids)
                
                if is_kompas and os.path.isdir(item_path):
                    try:
                        shutil.rmtree(item_path)
                        deleted_count += 1
                        print(f"  🗑️ Удалена папка KOMPAS: {item}")
                    except Exception as e:
                        print(f"  ⚠️ Не удалось удалить {item}: {e}")
        
        if deleted_count > 0:
            print(f"\n✅ Удалено элементов: {deleted_count}")
        else:
            print(f"\n✅ Нечего удалять (кэш KOMPAS пуст)")
        
        return True
        
    except ImportError:
        print("❌ Библиотека win32com не установлена")
        print("   Установите: pip install pywin32")
        return False
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return False


if __name__ == '__main__':
    print("=" * 50)
    print("ОЧИСТКА КЭША WIN32COM")
    print("=" * 50)
    print()
    
    # Проверяем аргументы
    full_clear = '--full' in sys.argv or '-f' in sys.argv
    
    if full_clear:
        print("⚠️ Режим: ПОЛНАЯ очистка (все COM-объекты)")
        response = input("Продолжить? (y/n): ").strip().lower()
        if response != 'y':
            print("Отменено.")
            sys.exit(0)
    else:
        print("📌 Режим: Выборочная очистка (только KOMPAS)")
    
    print()
    
    success = clear_win32com_cache(full_clear)
    
    print()
    if success:
        print("✅ Готово! Теперь можно запускать программу.")
    else:
        print("❌ Произошла ошибка при очистке кэша.")
    
    sys.exit(0 if success else 1)

