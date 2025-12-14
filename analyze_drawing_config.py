#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для анализа конфигурации чертежа КОМПАС-3D
"""

import sys
import pythoncom
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from components.base_component import get_dynamic_dispatch

def analyze_drawing_configuration(drawing_path):
    """Анализирует конфигурацию чертежа и связанных моделей"""
    
    pythoncom.CoInitialize()
    
    output_file = Path(__file__).parent / "drawing_config_analysis.txt"
    
    def log(message):
        print(message)
        with open(output_file, 'a', encoding='utf-8') as f:
            f.write(message + '\n')
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('')
    
    try:
        log("="*60)
        log("АНАЛИЗ КОНФИГУРАЦИИ ЧЕРТЕЖА")
        log("="*60)
        log(f"Файл: {drawing_path}\n")
        
        api7 = get_dynamic_dispatch("Kompas.Application.7")
        
        log("Открытие чертежа...")
        doc = api7.Documents.Open(str(drawing_path), False, False)
        
        if not doc:
            log("❌ Не удалось открыть файл")
            return
        
        log("✓ Чертеж открыт\n")
        
        kompas_document_2d = api7.ActiveDocument
        
        # Получаем виды
        log("="*60)
        log("ВИДЫ НА ЧЕРТЕЖЕ")
        log("="*60)
        
        try:
            views_and_layers = kompas_document_2d.ViewsAndLayersManager
            views = views_and_layers.Views
            
            log(f"Всего видов: {views.Count}\n")
            
            for i in range(views.Count):
                view = views.Item(i)
                log(f"Вид {i+1}:")
                log(f"  Название: {view.Name}")
                
                # Проверяем ассоциативность
                try:
                    if hasattr(view, 'AssociatedDrawingView'):
                        assoc_view = view.AssociatedDrawingView
                        if assoc_view:
                            log(f"  Тип: Ассоциативный вид")
                            
                            # Пытаемся получить путь к модели
                            try:
                                if hasattr(assoc_view, 'SourceFileName'):
                                    source_file = assoc_view.SourceFileName
                                    log(f"  Исходная модель: {source_file}")
                                    
                                    # Пытаемся получить конфигурацию
                                    if hasattr(assoc_view, 'ConfigurationName'):
                                        config_name = assoc_view.ConfigurationName
                                        log(f"  Конфигурация: {config_name}")
                            except Exception as e:
                                log(f"  ⚠️ Не удалось получить детали: {e}")
                    else:
                        log(f"  Тип: Обычный вид")
                except Exception as e:
                    log(f"  ⚠️ Ошибка анализа вида: {e}")
                
                log("")
        
        except Exception as e:
            log(f"❌ Ошибка получения видов: {e}")
        
        # Получаем переменные
        log("="*60)
        log("ПЕРЕМЕННЫЕ ЧЕРТЕЖА")
        log("="*60)
        
        try:
            if hasattr(kompas_document_2d, 'VariableCollection'):
                variables = kompas_document_2d.VariableCollection
                log(f"Всего переменных: {variables.Count}\n")
                
                for i in range(variables.Count):
                    var = variables.Item(i)
                    log(f"Переменная {i+1}:")
                    log(f"  Имя: {var.Name}")
                    log(f"  Значение: {var.Value}")
                    log("")
            else:
                log("⚠️ Переменные недоступны в этом типе документа\n")
        
        except Exception as e:
            log(f"❌ Ошибка получения переменных: {e}\n")
        
        # Получаем размеры
        log("="*60)
        log("РАЗМЕРЫ НА ЧЕРТЕЖЕ")
        log("="*60)
        
        try:
            # Пытаемся получить размеры через DrawingContainer
            if hasattr(kompas_document_2d, 'DrawingContainer'):
                container = kompas_document_2d.DrawingContainer()
                if hasattr(container, 'Dimensions'):
                    dimensions = container.Dimensions
                    log(f"Всего размеров: {dimensions.Count}\n")
                else:
                    log("⚠️ Размеры недоступны через этот интерфейс\n")
            else:
                log("⚠️ DrawingContainer недоступен\n")
        
        except Exception as e:
            log(f"⚠️ Размеры: {e}\n")
        
        log("="*60)
        log(f"Результаты сохранены в: {output_file}")
        log("="*60)
        
        api7.ActiveDocument.Close(False)
        log("Чертеж закрыт")
        
    except Exception as e:
        log(f"❌ Ошибка: {e}")
        import traceback
        log(traceback.format_exc())
    
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    # Пример: чертеж детали
    drawing_path = r"C:\Users\Vorob\Documents\ZVD GROUP\Заказы\Дарим Детям Добро А-191125-2\02_кд\ZVD.TURBO.80.230.2000\0010 - Корпус короба.cdw"
    analyze_drawing_configuration(drawing_path)
