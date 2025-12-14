#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Автоматическое обновление всех чертежей после копирования проекта
"""

import time
import pythoncom
from pathlib import Path
from typing import Dict
from .base_component import BaseKompasComponent, get_dynamic_dispatch

class DrawingAutoUpdater(BaseKompasComponent):
    """Автоматическое открытие, пересборка и сохранение всех чертежей"""
    
    def __init__(self):
        super().__init__()
    
    def update_all_drawings(self, project_path: str, developer: str = None, checker: str = None, organization: str = None, material: str = None, 
                          tech_control: str = None, norm_control: str = None, approved: str = None, date: str = None,
                          order_number: str = None, check_cancel=None) -> Dict:
        """
        Автоматическое обновление всех чертежей в проекте
        
        Args:
            project_path: Путь к проекту
            developer: Имя разработчика (ячейка 110)
            checker: Имя проверяющего (ячейка 111)
            organization: Организация (ячейка 9)
            material: Материал (ячейка 3)
            tech_control: Т. контр. (ячейка 112)
            norm_control: Н. контр. (ячейка 114)
            approved: Утв. (ячейка 115)
            date: Дата (ячейки 130-135)
            order_number: Номер заказа (обновляется в наименовании детали)
            check_cancel: Функция проверки отмены (возвращает True если нужно прервать)
            
        Returns:
            Dict с результатами
        """
        result = {
            'success': False,
            'drawings_updated': 0,
            'drawings_failed': 0,
            'updated_files': [],
            'errors': []
        }
        
        pythoncom.CoInitialize()
        
        try:
            self.logger.info("="*60)
            self.logger.info("АВТОМАТИЧЕСКОЕ ОБНОВЛЕНИЕ ЧЕРТЕЖЕЙ (v3 - Full Fields)")
            self.logger.info("="*60)
            
            # Принудительное переподключение для стабильности
            if not self.connect_to_kompas(force_reconnect=True):
                result['errors'].append("Не удалось подключиться к КОМПАС-3D")
                return result
            
            # Закрываем все открытые документы
            self.close_all_documents()
            time.sleep(0.5)
            
            project_path_obj = Path(project_path)
            
            # Находим все чертежи (включая развертки)
            all_drawings = list(project_path_obj.glob("*.cdw"))
            
            self.logger.info(f"\nНайдено чертежей: {len(all_drawings)}\n")
            
            # Используем dynamic dispatch для обхода проблем с кэшем типов
            api7 = get_dynamic_dispatch("Kompas.Application.7")
            
            for drawing in all_drawings:
                # ПРОВЕРКА ОТМЕНЫ
                if check_cancel and check_cancel():
                    self.logger.warning("⚠️ ОПЕРАЦИЯ ПРЕРВАНА ПОЛЬЗОВАТЕЛЕМ")
                    break
                try:
                    self.logger.info(f"{'='*60}")
                    self.logger.info(f"{drawing.name}")
                    self.logger.info(f"{'='*60}")
                    
                    # Открываем документ (API7)
                    self.logger.info("  Открытие...")
                    doc7 = api7.Documents.Open(str(drawing), False, False)
                    if not doc7:
                        self.logger.error(f"  Не удалось открыть файл (API7)")
                        result['drawings_failed'] += 1
                        continue
                    
                    # ВАЖНО: Даем время на открытие документа
                    time.sleep(2)
                    
                    # Получаем интерфейс 2D документа
                    kompas_document_2d = api7.ActiveDocument
                    
                    # ПРОВЕРКА: Пропускаем развертки (они обновляют только геометрию)
                    is_unfolding = "развертка" in drawing.name.lower() or "razvertka" in drawing.name.lower()
                    
                    if is_unfolding:
                        self.logger.info(f"  ℹ️ Развертка - пропуск обновления штампа и наименования")
                        self.logger.info(f"  Обновление только геометрии...")
                        
                        # Только пересборка для обновления геометрии
                        try:
                            kompas_document_2d.RebuildDocument()
                        except:
                            pass
                        time.sleep(2)
                        
                        # Сохраняем
                        api7.ActiveDocument.Save()
                        time.sleep(1)
                        
                        # Закрываем
                        api7.ActiveDocument.Close(False)
                        time.sleep(0.5)
                        
                        result['drawings_updated'] += 1
                        result['updated_files'].append(drawing.name)
                        self.logger.info(f"  ✓ Геометрия обновлена\n")
                        continue  # Переходим к следующему чертежу
                    
                    # ОБНОВЛЕНИЕ ШТАМПА (API7)
                    if any([developer, checker, organization, material, tech_control, norm_control, approved, date, order_number]):
                        try:
                            self.logger.info(f"  Обновление штампа (API7)...")
                            
                            # Получаем коллекцию листов оформления
                            layout_sheets = kompas_document_2d.LayoutSheets
                            # Берем первый лист (обычно штамп там)
                            sheet = layout_sheets.Item(0)
                            # Получаем штамп
                            stamp = sheet.Stamp
                            
                            if stamp:
                                # Словарь полей: {номер_ячейки: значение}
                                fields_to_update = {}
                                
                                # Основные поля
                                if developer: fields_to_update[110] = developer
                                if checker: fields_to_update[111] = checker
                                if tech_control: fields_to_update[112] = tech_control
                                if norm_control: fields_to_update[114] = norm_control
                                if approved: fields_to_update[115] = approved
                                
                                if organization: fields_to_update[9] = organization
                                
                                # Материал НЕ для сборочных чертежей (СБ)
                                is_assembly = "СБ" in drawing.name or "сб" in drawing.name.lower()
                                if material and not is_assembly:
                                    fields_to_update[3] = material
                                    self.logger.info(f"    Материал: {material}")
                                elif material and is_assembly:
                                    self.logger.info(f"    Материал пропущен (сборочный чертеж)")
                                
                                # ЛОГИКА ДАТЫ:
                                # Заполняем только дату разработки (ячейка 130)
                                date_cells_updated = []
                                if date:
                                    fields_to_update[130] = date  # Дата разработки
                                    date_cells_updated = ["Разраб."]
                                
                                for cell_id, value in fields_to_update.items():
                                    try:
                                        # Получаем интерфейс текста ячейки
                                        text_item = stamp.Text(cell_id)
                                        # Записываем значение
                                        text_item.Str = str(value)
                                        # text_item.Update() # Убрали, так как вызывает ошибку, stamp.Update() достаточно
                                        self.logger.info(f"    Ячейка {cell_id}: {value}")
                                    except Exception as e:
                                        self.logger.warning(f"    ⚠️ Ошибка ячейки {cell_id}: {e}")
                                
                                # Обновляем сам штамп
                                stamp.Update()
                                
                                # КРИТИЧНО: Даем время на обновление штампа!
                                time.sleep(2)
                                
                                if date_cells_updated:
                                    self.logger.info(f"  📅 Дата '{date}' установлена для: {', '.join(date_cells_updated)}")
                                self.logger.info(f"  ✓ Штамп обработан")
                            else:
                                self.logger.warning(f"  ⚠️ Штамп не найден")
                                
                        except Exception as e:
                            self.logger.warning(f"  ⚠️ Общая ошибка штампа: {e}")
                    
                    # ОБНОВЛЕНИЕ НОМЕРА ЗАКАЗА В НАИМЕНОВАНИИ ДЕТАЛИ
                    if order_number:
                        try:
                            self.logger.info(f"  Обновление номера заказа в наименовании...")
                            
                            # Получаем спецификацию
                            specifications = kompas_document_2d.Specifications
                            if specifications and specifications.Count > 0:
                                spec = specifications.Item(0)
                                
                                # Получаем объекты спецификации
                                spec_objects = spec.Objects
                                
                                # Проходим по всем объектам
                                for i in range(spec_objects.Count):
                                    spec_obj = spec_objects.Item(i)
                                    
                                    # Получаем описание объекта
                                    obj_description = spec_obj.Description
                                    
                                    if obj_description:
                                        old_name = obj_description
                                        
                                        # Убираем старый номер заказа (в скобках в конце)
                                        import re
                                        clean_name = re.sub(r'\s*\([^)]*\)\s*$', '', old_name).strip()
                                        
                                        # Добавляем новый номер заказа
                                        new_name = f"{clean_name} ({order_number})"
                                        
                                        # Обновляем
                                        spec_obj.Description = new_name
                                        spec_obj.Update()
                                        
                                        self.logger.info(f"    Наименование: '{clean_name}' → '{new_name}'")
                                
                                # Обновляем спецификацию
                                spec.Update()
                                self.logger.info(f"  ✓ Номер заказа обновлен")
                            else:
                                self.logger.info(f"  ℹ️ Спецификация не найдена (пропуск обновления номера заказа)")
                        
                        except Exception as e:
                            # Спецификация отсутствует - это нормально для большинства чертежей
                            self.logger.info(f"  ℹ️ Спецификация не найдена (номер заказа обновляется в 3D-модели)")
                    
                    # Определяем, это сборочный чертеж или нет
                    is_assembly = "конвектор" in drawing.name.lower() or "сборочный" in drawing.name.lower()
                    
                    # Перестраиваем
                    self.logger.info("  Пересборка...")
                    try:
                        # Пробуем вызвать как метод
                        kompas_document_2d.RebuildDocument()
                    except TypeError:
                        # Если это свойство или не вызывается
                        pass
                    except Exception as e:
                        self.logger.warning(f"  ⚠️ Ошибка перестроения: {e}")
                    
                    # КРИТИЧНО: Для сборочного чертежа - больше времени!
                    if is_assembly:
                        self.logger.info("  (СБОРОЧНЫЙ ЧЕРТЕЖ - увеличенное время обновления)")
                        time.sleep(5)  # Даем 5 секунд на обновление!
                        
                        # ПОВТОРНАЯ ПЕРЕСБОРКА для надежности
                        self.logger.info("  Повторная пересборка...")
                        try:
                            kompas_document_2d.RebuildDocument()
                        except:
                            pass
                        time.sleep(3)
                    else:
                        time.sleep(2)
                    
                    # self.logger.info(f"  Результат пересборки: {rebuild_result}")
                    
                    # Сохраняем
                    self.logger.info("  Сохранение...")
                    api7.ActiveDocument.Save()
                    
                    # КРИТИЧНО: Даем больше времени на сохранение!
                    time.sleep(3)
                    
                    self.logger.info("  ✓ Готово")
                    
                    # Закрываем
                    api7.ActiveDocument.Close(False)
                    time.sleep(0.5)
                    
                    result['drawings_updated'] += 1
                    result['updated_files'].append(drawing.name)
                    self.logger.info("")
                
                except Exception as e:
                    error_msg = f"Ошибка {drawing.name}: {e}"
                    result['errors'].append(error_msg)
                    self.logger.error(f"  ✗ {error_msg}\n")
                    result['drawings_failed'] += 1
                    
                    # Пытаемся закрыть
                    try:
                        api7.ActiveDocument.Close(False)
                    except:
                        pass
            
            result['success'] = True
            
            self.logger.info("="*60)
            self.logger.info("ИТОГО:")
            self.logger.info(f"  Обновлено: {result['drawings_updated']}")
            self.logger.info(f"  Ошибок: {result['drawings_failed']}")
            self.logger.info("="*60)
            
        except Exception as e:
            error_msg = f"Общая ошибка: {e}"
            result['errors'].append(error_msg)
            self.logger.error(error_msg)
        
        finally:
            pythoncom.CoUninitialize()
        
        return result
    
    def _detect_old_project_path(self, new_project_path: Path) -> str:
        """Определение старого пути по имени сборки"""
        try:
            assembly_files = list(new_project_path.glob("*.a3d"))
            
            if assembly_files:
                assembly_name = assembly_files[0].stem
                project_name = new_project_path.name
                
                # Если имена не совпадают - старое имя в сборке
                if assembly_name != project_name:
                    old_path = new_project_path.parent / assembly_name
                    return str(old_path)
            
            return None
        except:
            return None

