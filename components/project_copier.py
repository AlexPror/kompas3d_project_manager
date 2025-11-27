"""
Компонент для копирования проектов
"""
import shutil
import logging
from pathlib import Path
from typing import Dict
from .base_component import BaseKompasComponent

class ProjectCopier(BaseKompasComponent):
    """Компонент для копирования проектов КОМПАС-3D"""
    
    def __init__(self):
        super().__init__()
    
    def copy_project(self, source_path: str, target_folder: str, project_name: str) -> Dict:
        """
        Копирование проекта
        
        Args:
            source_path: Путь к исходному проекту
            target_folder: Папка назначения
            project_name: Имя нового проекта
            
        Returns:
            Dict с результатами копирования
        """
        result = {
            'success': False,
            'copied_path': None,
            'error': None
        }
        
        try:
            self.logger.info("НАЧАЛО КОПИРОВАНИЯ ПРОЕКТА")
            self.logger.info("=" * 40)
            
            source_path_obj = Path(source_path)
            target_path = Path(target_folder) / project_name
            
            # Проверяем исходный путь
            if not source_path_obj.exists():
                result['error'] = f"Исходная папка не найдена: {source_path}"
                return result
            
            if not source_path_obj.is_dir():
                result['error'] = f"Исходный путь не является папкой: {source_path}"
                return result
            
            # Проверяем наличие временных файлов в исходной папке (они будут проигнорированы при копировании)
            self.logger.info(f"Проверка исходной папки на временные файлы...")
            source_temp_files = list(source_path_obj.rglob("~$*")) + list(source_path_obj.rglob("*.cd~"))
            if source_temp_files:
                self.logger.info(f"📝 Обнаружено {len(source_temp_files)} временных файлов (будут проигнорированы при копировании)")
            else:
                self.logger.info(f"✅ Временных файлов не обнаружено")
            
            # Проверяем целевую папку
            target_folder_obj = Path(target_folder)
            if not target_folder_obj.exists():
                self.logger.info(f"Создание целевой папки: {target_folder}")
                target_folder_obj.mkdir(parents=True, exist_ok=True)
            
            # Удаляем существующий проект если есть
            if target_path.exists():
                self.logger.info(f"⚠️ Целевая папка уже существует: {target_path}")
                self.logger.info("   Удаление старого проекта...")
                try:
                    # Проверяем наличие открытых файлов
                    temp_files = list(target_path.rglob("~$*")) + list(target_path.rglob("*.cd~"))
                    if temp_files:
                        self.logger.warning(f"   ⚠️ Обнаружено {len(temp_files)} открытых файлов!")
                        self.logger.warning("   Закройте все файлы в КОМПАС-3D и других программах!")
                        result['error'] = f"В целевой папке есть открытые файлы. Закройте их и повторите попытку."
                        return result
                    
                    shutil.rmtree(target_path)
                    self.logger.info("   ✅ Старый проект удален")
                except PermissionError as e:
                    result['error'] = f"Нет доступа для удаления: {target_path}. Закройте все файлы и попробуйте снова."
                    self.logger.error(result['error'])
                    return result
                except Exception as e:
                    result['error'] = f"Ошибка удаления существующего проекта: {e}"
                    self.logger.error(result['error'])
                    return result
            
            # Копируем проект с фильтрацией
            self.logger.info(f"Копирование из {source_path} в {target_path}")
            
            # ОПТИМИЗАЦИЯ: Игнорируем временные и системные файлы!
            def ignore_files(directory, files):
                """Фильтр для исключения ненужных файлов"""
                ignore_list = []
                for f in files:
                    # Игнорируем:
                    if f.endswith('.bak'):  # Резервные копии КОМПАС
                        ignore_list.append(f)
                    elif f.endswith('~'):  # Временные файлы (Linux/Mac)
                        ignore_list.append(f)
                    elif f.startswith('~$'):  # Временные файлы КОМПАС-3D и MS Office
                        ignore_list.append(f)
                    elif f.startswith('~'):  # Другие временные файлы
                        ignore_list.append(f)
                    elif f.endswith('.tmp'):  # Временные файлы
                        ignore_list.append(f)
                    elif f.endswith('.temp'):  # Временные файлы
                        ignore_list.append(f)
                    elif f.endswith('.lock'):  # Файлы блокировки
                        ignore_list.append(f)
                    elif f.endswith('.cd~'):  # Временные файлы КОМПАС-3D
                        ignore_list.append(f)
                    elif f == 'Thumbs.db':  # Windows кеш
                        ignore_list.append(f)
                    elif f == '.DS_Store':  # macOS кеш
                        ignore_list.append(f)
                
                if ignore_list:
                    self.logger.info(f"  📝 Пропущено временных файлов: {len(ignore_list)}")
                
                return ignore_list
            
            try:
                shutil.copytree(source_path, target_path, ignore=ignore_files)
            except Exception as e:
                result['error'] = f"Ошибка при копировании файлов: {e}"
                self.logger.error(result['error'])
                # Удаляем частично скопированную папку
                if target_path.exists():
                    self.logger.info("Удаление частично скопированной папки...")
                    try:
                        shutil.rmtree(target_path)
                    except:
                        pass
                return result
            
            # Подсчет скопированных файлов
            copied_files = list(target_path.rglob("*"))
            copied_files_count = sum(1 for f in copied_files if f.is_file())
            
            result['success'] = True
            result['copied_path'] = str(target_path)
            result['copied_files'] = copied_files_count
            
            self.logger.info(f"✅ Проект успешно скопирован!")
            self.logger.info(f"   📁 Путь: {target_path}")
            self.logger.info(f"   📊 Скопировано файлов: {copied_files_count}")
            
            return result
            
        except Exception as e:
            error_msg = f"Ошибка копирования проекта: {e}"
            result['error'] = error_msg
            self.logger.error(error_msg)
            return result
    
    def rename_main_assembly(self, project_path: str, project_name: str) -> Dict:
        """
        Переименование главной сборки и чертежа сборки
        
        Args:
            project_path: Путь к проекту
            project_name: Имя проекта (например: ZVD.LITE.160.350.2600)
            
        Returns:
            Dict с результатами переименования
        """
        result = {
            'success': False,
            'renamed_files': [],
            'error': None
        }
        
        try:
            self.logger.info("ПЕРЕИМЕНОВАНИЕ ГЛАВНОЙ СБОРКИ И ЧЕРТЕЖА")
            self.logger.info("=" * 40)
            
            project_path_obj = Path(project_path)
            
            if not project_path_obj.exists():
                result['error'] = f"Папка проекта не найдена: {project_path}"
                return result
            
            renamed_count = 0
            
            # 1. Переименовываем главную сборку (.a3d)
            assembly_files = list(project_path_obj.glob("*.a3d"))
            if assembly_files:
                main_assembly = assembly_files[0]  # Берем первый найденный
                old_assembly_name = main_assembly.stem  # Имя без расширения
                new_assembly_name = f"{project_name}.a3d"
                new_assembly_path = main_assembly.parent / new_assembly_name
                
                self.logger.info(f"📦 Сборка: {main_assembly.name} → {new_assembly_name}")
                
                # Проверяем, существует ли целевой файл
                if new_assembly_path.exists() and new_assembly_path != main_assembly:
                    self.logger.warning(f"   ⚠️ Файл уже существует, удаляем: {new_assembly_name}")
                    new_assembly_path.unlink()
                
                if main_assembly != new_assembly_path:
                    main_assembly.rename(new_assembly_path)
                    result['renamed_files'].append(str(new_assembly_path))
                    renamed_count += 1
                
                # 2. Ищем и переименовываем чертеж сборки (.cdw) с таким же именем
                old_drawing_path = main_assembly.parent / f"{old_assembly_name}.cdw"
                if old_drawing_path.exists():
                    new_drawing_name = f"{project_name}.cdw"
                    new_drawing_path = old_drawing_path.parent / new_drawing_name
                    
                    self.logger.info(f"📐 Чертеж: {old_drawing_path.name} → {new_drawing_name}")
                    
                    # Проверяем, существует ли целевой файл
                    if new_drawing_path.exists() and new_drawing_path != old_drawing_path:
                        self.logger.warning(f"   ⚠️ Файл уже существует, удаляем: {new_drawing_name}")
                        new_drawing_path.unlink()
                    
                    if old_drawing_path != new_drawing_path:
                        old_drawing_path.rename(new_drawing_path)
                        result['renamed_files'].append(str(new_drawing_path))
                        renamed_count += 1
                else:
                    self.logger.warning(f"⚠️ Чертеж сборки не найден: {old_assembly_name}.cdw")
            else:
                result['error'] = "Файлы сборки (.a3d) не найдены"
                return result
            
            result['success'] = True
            result['renamed_count'] = renamed_count
            self.logger.info(f"✅ Переименовано файлов: {renamed_count}")
            
            return result
            
        except Exception as e:
            error_msg = f"Ошибка переименования: {e}"
            result['error'] = error_msg
            self.logger.error(error_msg)
            return result
    
    def get_project_info(self, project_path: str) -> Dict:
        """
        Получение информации о проекте
        
        Args:
            project_path: Путь к проекту
            
        Returns:
            Dict с информацией о проекте
        """
        info = {
            'project_path': project_path,
            'assembly_files': [],
            'drawing_files': [],
            'part_files': [],
            'other_files': [],
            'total_files': 0
        }
        
        try:
            project_path_obj = Path(project_path)
            
            if not project_path_obj.exists():
                return info
            
            # Подсчет файлов по типам
            for file_path in project_path_obj.rglob("*"):
                if file_path.is_file():
                    suffix = file_path.suffix.lower()
                    
                    if suffix == '.a3d':
                        info['assembly_files'].append(str(file_path))
                    elif suffix == '.cdw':
                        info['drawing_files'].append(str(file_path))
                    elif suffix == '.m3d':
                        info['part_files'].append(str(file_path))
                    else:
                        info['other_files'].append(str(file_path))
                    
                    info['total_files'] += 1
            
            self.logger.info(f"Информация о проекте:")
            self.logger.info(f"  Сборки: {len(info['assembly_files'])}")
            self.logger.info(f"  Чертежи: {len(info['drawing_files'])}")
            self.logger.info(f"  Детали: {len(info['part_files'])}")
            self.logger.info(f"  Всего файлов: {info['total_files']}")
            
            return info
            
        except Exception as e:
            self.logger.error(f"Ошибка получения информации о проекте: {e}")
            return info
