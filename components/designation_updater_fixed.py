#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Компонент для обновления обозначений деталей (ИСПРАВЛЕННАЯ ВЕРСИЯ)
С правильной последовательностью: Update() + RebuildModel() + RebuildDocument() + Save() + Close(True)
"""

import logging
from pathlib import Path
from typing import Dict, List
from win32com.client import gencache
from .base_component import BaseKompasComponent, clear_kompas_cache, get_dynamic_dispatch
import time

class DesignationUpdaterFixed(BaseKompasComponent):
    """Компонент для обновления обозначений деталей с РАБОЧЕЙ логикой"""
    
    def __init__(self):
        super().__init__()
    
    def update_all_designations(self, project_path: str, H: int, B1: int, L1: int, order_number: str = None, project_type: str = None) -> Dict:
        """
        Обновление ВСЕХ обозначений (сборка + детали) с правильной логикой
        
        Логика обозначений:
        - Сборка: ZVD.LITE.H.B1.L1
        - Корпус короба (004): ZVD.LITE.H.B1.L1.XXX (с L1)
        - Остальные детали: ZVD.LITE.H.B1.XXX (без L1)
        
        Args:
            project_path: Путь к проекту
            H: Высота
            B1: Ширина
            L1: Длина
            order_number: Номер заказа
            project_type: Тип проекта (ZVD.LITE или ZVD.TURBO). Если None - определяется по папке.
            
        Returns:
            Dict с результатами
        """
        result = {
            'success': False,
            'assembly_renamed': False,
            'parts_renamed': 0,
            'updated_parts': [],
            'errors': [],
            'error': None
        }
        
        try:
            self.logger.info("ОБНОВЛЕНИЕ ВСЕХ ОБОЗНАЧЕНИЙ")
            self.logger.info("=" * 50)
            self.logger.info(f"H={H}, B1={B1}, L1={L1}")
            
            # Принудительное переподключение для стабильности
            if not self.connect_to_kompas(force_reconnect=True):
                result['error'] = "Не удалось подключиться к КОМПАС-3D"
                return result
            
            # Закрываем все открытые документы
            self.close_all_documents()
            time.sleep(0.5)
            
            project_path_obj = Path(project_path)
            
            # Определяем тип проекта
            project_folder_name = project_path_obj.name
            project_prefix = "ZVD.LITE"  # По умолчанию
            
            if project_type:
                # Если передан явно - используем его
                project_prefix = project_type
            else:
                # Иначе пытаемся определить из имени папки
                if "TURBO" in project_folder_name.upper():
                    project_prefix = "ZVD.TURBO"
                elif "LITE" in project_folder_name.upper():
                    project_prefix = "ZVD.LITE"
            
            self.logger.info(f"Определен тип проекта: {project_prefix}")
            
            # Формируем обозначения
            # Для marking сборки (с " СБ")
            full_name_with_sb = f"{project_prefix}.{H}.{B1}.{L1} СБ"
            # Для использования в обозначениях деталей (БЕЗ " СБ")
            full_name = f"{project_prefix}.{H}.{B1}.{L1}"
            short_name = f"{project_prefix}.{H}.{B1}"
            
            # Очистка кэша для устранения ошибок CLSIDToClassMap
            clear_kompas_cache()
            
            # Получаем API через dynamic dispatch
            api5_obj = get_dynamic_dispatch("Kompas.Application.5")
            api7_obj = get_dynamic_dispatch("Kompas.Application.7")
            
            # Загружаем константы (может потребовать кэша, обернем в try)
            try:
                kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants
            except Exception:
                # Если не удалось загрузить константы, создадим заглушку
                class Constants3D:
                    ksD3TextPropertyProduct = 1
                    ksD3TextPropertyName = 2
                kompas6_constants_3d = Constants3D()
            
            # ШАГ 1: ПЕРЕИМЕНОВАНИЕ СБОРКИ
            self.logger.info("\n--- ПЕРЕИМЕНОВАНИЕ СБОРКИ ---")
            
            assembly_files = list(project_path_obj.glob("*.a3d"))
            if not assembly_files:
                result['error'] = "Файл сборки не найден"
                return result
            
            assembly_file = assembly_files[0]
            self.logger.info(f"Файл сборки: {assembly_file.name}")
            
            # Переименовываем файл ПЕРЕД открытием (с СБ в названии)
            new_assembly_name = f"{project_prefix}.{H}.{B1}.{L1} СБ.a3d"
            new_assembly_path = project_path_obj / new_assembly_name
            
            if assembly_file.name != new_assembly_name:
                try:
                    self.logger.info(f"Переименование файла: {assembly_file.name} → {new_assembly_name}")
                    assembly_file.rename(new_assembly_path)
                    assembly_file = new_assembly_path  # Обновляем путь
                    self.logger.info(f"✓ Файл сборки переименован")
                except Exception as e:
                    error_msg = f"Ошибка переименования файла сборки: {e}"
                    result['errors'].append(error_msg)
                    self.logger.error(error_msg)
                    # Продолжаем работу с оригинальным именем
            
            try:
                doc3D_asm = api5_obj.Document3D
                doc3D_asm.Open(str(assembly_file), False)
                time.sleep(1)
                
                iDocument3D_asm = api5_obj.ActiveDocument3D
                iPart_asm = iDocument3D_asm.GetPart(kompas6_constants_3d.pTop_Part)
                
                old_asm = iPart_asm.marking
                # Используем вариант С " СБ" для marking сборки
                iPart_asm.marking = full_name_with_sb
                
                # REBUILD!
                iPart_asm.Update()
                iPart_asm.RebuildModel()
                iDocument3D_asm.RebuildDocument()
                time.sleep(1)
                
                # Сохраняем
                active_doc_asm = api7_obj.ActiveDocument
                active_doc_asm.Save()
                time.sleep(2)
                
                # Закрываем С СОХРАНЕНИЕМ
                active_doc_asm.Close(True)
                time.sleep(2)
                
                self.logger.info(f"✓ Сборка: '{old_asm}' → '{full_name_with_sb}'")
                result['assembly_renamed'] = True
                
            except Exception as e:
                error_msg = f"Ошибка переименования сборки: {e}"
                result['errors'].append(error_msg)
                self.logger.error(error_msg)
            
            # ШАГ 2: ОБНОВЛЕНИЕ MARKING В СБОРКЕ (СНАЧАЛА!)
            # ВАЖНО: Сначала обновляем marking в сборке, чтобы знать правильные номера деталей
            self.logger.info("\n--- ОБНОВЛЕНИЕ MARKING В СБОРКЕ ---")
            self.logger.info("(определение правильных номеров для всех деталей)")
            
            # Словарь: {имя_файла_детали: (marking_из_сборки, наименование)}
            file_to_marking = {}
            
            try:
                # Открываем сборку
                assembly_files_reopen = list(project_path_obj.glob("*.a3d"))
                if assembly_files_reopen:
                    assembly_file_reopen = assembly_files_reopen[0]
                    
                    doc3D_asm2 = api5_obj.Document3D
                    doc3D_asm2.Open(str(assembly_file_reopen), False)
                    time.sleep(2)
                    
                    iDocument3D_asm2 = api5_obj.ActiveDocument3D
                    iPart_asm2 = iDocument3D_asm2.GetPart(kompas6_constants_3d.pTop_Part)
                    
                    # Находим все детали в сборке
                    asm_parts_count = 0
                    for i in range(100):
                        try:
                            p = iPart_asm2.GetPart(i)
                            if p:
                                asm_parts_count = i + 1
                            else:
                                break
                        except:
                            break
                    
                    self.logger.info(f"Детал ей в сборке: {asm_parts_count}")
                    
                    # КРИТИЧНО: Группируем детали по имени!
                    # Детали с одинаковым именем получают один номер!
                    name_to_number = {}  # {имя_детали: номер}
                    next_number = 1
                    
                    # Первый проход - присваиваем номера уникальным именам
                    for i in range(asm_parts_count):
                        try:
                            p = iPart_asm2.GetPart(i)
                            if not p:
                                continue
                            
                            old_mark = p.marking.strip() if hasattr(p, 'marking') else ""
                            p_name = p.name.strip() if hasattr(p, 'name') else ""
                            
                            # Пропускаем служебные
                            if not old_mark or old_mark.startswith('-'):
                                continue
                            
                            # Присваиваем номер (если еще не присвоен)
                            if p_name not in name_to_number:
                                name_to_number[p_name] = next_number
                                next_number += 1
                        except:
                            pass
                    
                    self.logger.info(f"\n  Уникальных деталей: {len(name_to_number)}")
                    for name, num in sorted(name_to_number.items(), key=lambda x: x[1]):
                        self.logger.info(f"    {num:03d} - {name}")
                    
                    # Второй проход - обновляем marking И СОХРАНЯЕМ соответствие файл->marking
                    updated_instances = 0
                    
                    for i in range(asm_parts_count):
                        try:
                            p = iPart_asm2.GetPart(i)
                            if not p:
                                continue
                            
                            old_mark = p.marking.strip() if hasattr(p, 'marking') else ""
                            p_name = p.name.strip() if hasattr(p, 'name') else ""
                            p_filename = p.fileName.strip() if hasattr(p, 'fileName') else ""
                            
                            # Пропускаем служебные
                            if not old_mark or old_mark.startswith('-'):
                                continue
                            
                            # Получаем номер для этого имени
                            part_id = f"{name_to_number[p_name]:03d}"
                            
                            # СПЕЦИАЛЬНАЯ ОБРАБОТКА для покупных компонентов
                            if any(word in p_name.lower() for word in ["теплообменник", "вентилятор", "блок электропитания", "блок питания"]):
                                # Для покупных компонентов не меняем базовое обозначение
                                new_mark = old_mark  # Оставляем как есть
                            elif "короба" in p_name.lower():
                                # Корпус короба - используем full_name БЕЗ " СБ"
                                new_mark = f"{full_name}.{part_id}"
                            else:
                                new_mark = f"{short_name}.{part_id}"
                            
                            # СОХРАНЯЕМ соответствие fileName -> marking
                            if p_filename:
                                filename_normalized = Path(p_filename).name
                                if filename_normalized not in file_to_marking:
                                    file_to_marking[filename_normalized] = (new_mark, p_name, part_id)
                            
                            # Обновляем ТОЛЬКО если отличается
                            if old_mark != new_mark:
                                p.marking = new_mark
                                p.Update()
                                p.RebuildModel()
                                
                                self.logger.info(f"  [{i}] {p_name}: '{old_mark}' → '{new_mark}'")
                                updated_instances += 1
                            
                        except Exception as e:
                            self.logger.warning(f"Ошибка экземпляра {i}: {e}")
                    
                    if updated_instances > 0:
                        # Rebuild и сохранение сборки
                        iDocument3D_asm2.RebuildDocument()
                        time.sleep(1)
                        
                        active_doc2 = api7_obj.ActiveDocument
                        active_doc2.Save()
                        time.sleep(1)
                        
                        self.logger.info(f"✓ Обновлено экземпляров в сборке: {updated_instances}")
                    
                    # Закрываем
                    api7_obj.ActiveDocument.Close(True)
                    time.sleep(1)
                    
            except Exception as e:
                error_msg = f"Ошибка обновления экземпляров в сборке: {e}"
                result['errors'].append(error_msg)
                self.logger.error(error_msg)
            
            # ШАГ 3: ОБНОВЛЕНИЕ MARKING В ФАЙЛАХ ДЕТАЛЕЙ И ПЕРЕИМЕНОВАНИЕ ФАЙЛОВ
            self.logger.info("\n--- ОБНОВЛЕНИЕ MARKING В ФАЙЛАХ ДЕТАЛЕЙ ---")
            self.logger.info(f"Найдено соответствий файл->marking из сборки: {len(file_to_marking)}")
            
            # ВАЖНО: Также обрабатываем ВСЕ .m3d файлы в папке (даже если не в сборке)
            all_part_files = list(project_path_obj.glob("*.m3d"))
            self.logger.info(f"Всего .m3d файлов в папке: {len(all_part_files)}")
            
            # Добавляем файлы, которых нет в file_to_marking
            for part_file in all_part_files:
                if part_file.name not in file_to_marking:
                    # Пытаемся определить номер из имени файла
                    file_stem = part_file.stem
                    if " - " in file_stem:
                        parts = file_stem.split(" - ", 1)
                        part_number = parts[0].strip()
                        part_desc = parts[1].strip()
                    else:
                        part_number = "999"  # Неизвестный номер
                        part_desc = file_stem
                    
                    # Формируем marking (без L1 для обычных деталей)
                    new_marking = f"{short_name}.{part_number}"
                    
                    self.logger.info(f"  ⚠️ Деталь не в сборке: {part_file.name} → {new_marking}")
                    file_to_marking[part_file.name] = (new_marking, part_desc, part_number)
            
            self.logger.info(f"Всего файлов для обработки: {len(file_to_marking)}")
            
            renamed_count = 0
            files_to_rename = []  # [(старый_путь, новый_путь, marking)]
            
            for filename, (marking, part_name, part_id) in file_to_marking.items():
                try:
                    # Находим файл
                    part_files_found = list(project_path_obj.glob(filename))
                    if not part_files_found:
                        self.logger.warning(f"  ⚠ Файл не найден: {filename}")
                        continue
                    
                    part_file = part_files_found[0]
                    self.logger.info(f"\n  [{part_id}] {filename}")
                    
                    # Открываем файл детали
                    doc3D = api5_obj.Document3D
                    doc3D.Open(str(part_file), False)
                    time.sleep(0.5)
                    
                    iDocument3D = api5_obj.ActiveDocument3D
                    iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)
                    
                    old_marking = iPart.marking
                    
                    # СПЕЦИАЛЬНАЯ ОБРАБОТКА для покупных компонентов
                    if "теплообменник" in filename.lower():
                        # Для теплообменника обновляем длину (L1-300)
                        heat_length = L1 - 300
                        old_parts = old_marking.split()
                        if len(old_parts) >= 2 and '.' in old_parts[0]:
                            size_parts = old_parts[0].split('.')
                            if len(size_parts) == 3:
                                size_parts[2] = str(heat_length)
                                new_marking = f"{'.'.join(size_parts)} {' '.join(old_parts[1:])}"
                            else:
                                new_marking = old_marking
                        else:
                            new_marking = old_marking
                        self.logger.info(f"    Теплообменник: '{old_marking}' → '{new_marking}'")
                    elif any(word in filename.lower() for word in ["вентилятор", "блок электропитания", "блок питания"]):
                        # Для вентилятора и блока питания - не меняем обозначение
                        new_marking = old_marking
                        self.logger.info(f"    Покупной компонент (не изменяется): '{old_marking}'")
                    else:
                        # Обычная деталь - используем marking из сборки
                        new_marking = marking
                        self.logger.info(f"    Marking: '{old_marking}' → '{new_marking}'")
                    
                    # Обновляем marking
                    iPart.marking = new_marking
                    
                    # Обновляем наименование (добавляем номер заказа)
                    if order_number:
                        old_name = iPart.name if hasattr(iPart, 'name') else ""
                        
                        # Убираем старый номер заказа
                        import re
                        clean_name = re.sub(r'\s*\([^)]*\)\s*$', '', old_name).strip()
                        
                        # Добавляем новый
                        new_name = f"{clean_name} ({order_number})"
                        iPart.name = new_name
                        
                        self.logger.info(f"    Наименование: '{clean_name}' → '{new_name}'")
                    
                    # Rebuild и сохранение
                    iPart.Update()
                    iPart.RebuildModel()
                    iDocument3D.RebuildDocument()
                    time.sleep(0.5)
                    
                    active_doc = api7_obj.ActiveDocument
                    active_doc.Save()
                    time.sleep(0.5)
                    
                    active_doc.Close(True)
                    time.sleep(0.5)
                    
                    # Формируем новое имя файла по part_id
                    # Извлекаем оригинальное имя без номера
                    file_stem = part_file.stem
                    if " - " in file_stem:
                        file_description = file_stem.split(" - ", 1)[1]
                    else:
                        file_description = file_stem
                    
                    new_filename = f"{part_id} - {file_description}.m3d"
                    new_filepath = part_file.parent / new_filename
                    
                    # Сохраняем для последующего переименования
                    if part_file != new_filepath:
                        files_to_rename.append((part_file, new_filepath, marking))
                    
                    renamed_count += 1
                    
                except Exception as e:
                    error_msg = f"Ошибка обновления {filename}: {e}"
                    result['errors'].append(error_msg)
                    self.logger.error(f"    ✗ {error_msg}")
                    try:
                        api7_obj.ActiveDocument.Close(False)
                    except:
                        pass
            
            result['parts_renamed'] = renamed_count
            
            # Переименовываем файлы ПОСЛЕ закрытия всех документов
            self.logger.info(f"\n--- ПЕРЕИМЕНОВАНИЕ ФАЙЛОВ ДЕТАЛЕЙ ---")
            self.logger.info(f"Файлов к переименованию: {len(files_to_rename)}")
            
            for old_path, new_path, marking in files_to_rename:
                try:
                    self.logger.info(f"  '{old_path.name}' → '{new_path.name}'")
                    old_path.rename(new_path)
                except Exception as e:
                    self.logger.warning(f"    ⚠ Ошибка переименования: {e}")
            
            # ШАГ 4: ПЕРЕИМЕНОВАНИЕ ЧЕРТЕЖЕЙ
            self.logger.info("\n--- ПЕРЕИМЕНОВАНИЕ ЧЕРТЕЖЕЙ ---")
            
            # Обрабатываем ВСЕ чертежи в папке
            all_drawings = list(project_path_obj.glob("*.cdw"))
            self.logger.info(f"Найдено чертежей: {len(all_drawings)}")
            
            drawings_renamed = 0
            
            for drw_file in all_drawings:
                try:
                    old_name = drw_file.name
                    
                    # Определяем, сборочный это чертеж или чертеж детали
                    is_assembly_drawing = any(word in old_name.lower() for word in ["конвектор", "сборочный", "сборка"])
                    
                    if is_assembly_drawing:
                        # Сборочный чертеж - используем full_name_with_sb (с " СБ")
                        if " - " in old_name:
                            name_parts = old_name.split(" - ", 1)
                            description = name_parts[1]
                            new_drawing_name = f"{full_name_with_sb} - {description}"
                        else:
                            new_drawing_name = f"{full_name_with_sb} - {old_name}"
                    else:
                        # Чертеж детали - используем short_name (без L1)
                        # Пытаемся извлечь номер детали из имени файла
                        file_stem = drw_file.stem
                        if " - " in file_stem:
                            parts = file_stem.split(" - ", 1)
                            part_number = parts[0].strip()
                            description = parts[1].strip()
                            
                            # Формируем новое имя: номер - описание
                            new_drawing_name = f"{part_number} - {description}.cdw"
                        else:
                            # Если формат неизвестен, просто обновляем префикс
                            new_drawing_name = old_name
                    
                    new_drawing_path = drw_file.parent / new_drawing_name
                    
                    if drw_file != new_drawing_path:
                        self.logger.info(f"Чертеж: '{old_name}'")
                        self.logger.info(f"      → '{new_drawing_name}'")
                        
                        # Проверяем, существует ли целевой файл
                        if new_drawing_path.exists():
                            self.logger.warning(f"      ⚠️ Файл уже существует, удаляем")
                            new_drawing_path.unlink()
                        
                        drw_file.rename(new_drawing_path)
                        drawings_renamed += 1
                        self.logger.info(f"      ✓ Чертеж переименован")
                
                except Exception as e:
                    self.logger.warning(f"      ⚠️ Ошибка переименования чертежа {drw_file.name}: {e}")
            
            result['drawings_renamed'] = drawings_renamed
            result['success'] = result['assembly_renamed'] or renamed_count > 0 or drawings_renamed > 0
            
            self.logger.info(f"\n✓ Итого:")
            self.logger.info(f"  - Сборка переименована: {result['assembly_renamed']}")
            self.logger.info(f"  - Деталей переименовано: {renamed_count}")
            self.logger.info(f"  - Чертежей переименовано: {drawings_renamed}")
            
            return result
            
        except Exception as e:
            error_msg = f"Общая ошибка: {e}"
            result['error'] = error_msg
            self.logger.error(error_msg)
            import traceback
            traceback.print_exc()
            return result
        finally:
            self.disconnect_from_kompas()

