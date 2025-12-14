"""
Компонент для экспорта чертежей в изображения
"""
import logging
import time
from pathlib import Path
from typing import Dict, List
from win32com.client import gencache
from .base_component import BaseKompasComponent, clear_kompas_cache, get_dynamic_dispatch

class DrawingExporter(BaseKompasComponent):
    """Компонент для экспорта чертежей в изображения"""
    
    def __init__(self):
        super().__init__()
        self.supported_formats = {
            'BMP': 1,
            'PNG': 2,
            'JPEG': 3,
            'TIFF': 4
        }
    
    def find_drawing_files(self, project_path: str, exclude_unfoldings: bool = True) -> List[str]:
        """
        Поиск файлов чертежей в проекте
        
        Args:
            project_path: Путь к проекту
            exclude_unfoldings: Исключать развертки
            
        Returns:
            List путей к файлам чертежей
        """
        drawing_files = []
        
        try:
            self.logger.info("ПОИСК ФАЙЛОВ ЧЕРТЕЖЕЙ")
            self.logger.info("=" * 25)
            
            project_path_obj = Path(project_path)
            
            if not project_path_obj.exists():
                self.logger.error(f"Папка проекта не найдена: {project_path}")
                return drawing_files
            
            # Ищем все .cdw файлы
            cdw_files = list(project_path_obj.glob("*.cdw"))
            
            if exclude_unfoldings:
                # Исключаем развертки
                cdw_files = [f for f in cdw_files if not any(
                    word in f.name.lower() for word in ['развертка', 'развёртка']
                )]
            
            drawing_files = [str(f) for f in cdw_files]
            
            self.logger.info(f"Найдено чертежей: {len(drawing_files)}")
            for drawing_file in drawing_files:
                self.logger.info(f"  - {Path(drawing_file).name}")
            
            return drawing_files
            
        except Exception as e:
            self.logger.error(f"Ошибка поиска чертежей: {e}")
            return drawing_files
    
    def export_drawing_to_image(self, drawing_path: str, output_path: str, 
                               format_type: str = 'PNG', resolution: int = 300) -> Dict:
        """
        Экспорт одного чертежа в изображение
        
        Args:
            drawing_path: Путь к файлу чертежа
            output_path: Путь для сохранения изображения
            format_type: Тип формата (BMP, PNG, JPEG, TIFF)
            resolution: Разрешение в DPI
            
        Returns:
            Dict с результатами экспорта
        """
        result = {
            'success': False,
            'output_file': None,
            'file_size': 0,
            'error': None
        }
        
        try:
            self.logger.info(f"ЭКСПОРТ ЧЕРТЕЖА: {Path(drawing_path).name}")
            self.logger.info("=" * 40)
            
            # Принудительное переподключение для стабильности
            if not self.connect_to_kompas(force_reconnect=True):
                result['error'] = "Не удалось подключиться к КОМПАС-3D"
                return result
            
            if not Path(drawing_path).exists():
                result['error'] = f"Файл чертежа не найден: {drawing_path}"
                return result
            
            # Подготовка пути вывода
            output_path_obj = Path(output_path)
            output_path_obj.parent.mkdir(parents=True, exist_ok=True)
            
            # Удаляем существующий файл
            if output_path_obj.exists():
                output_path_obj.unlink()
            
            # Открываем чертеж
            if not self.open_document(drawing_path):
                result['error'] = "Не удалось открыть чертеж"
                return result
            
            # Принудительная перестройка чертежа
            if not self.force_rebuild_drawing():
                self.logger.warning("Не удалось перестроить чертеж, продолжаем...")
            else:
                # Удаляем watermark (надписи о некоммерческом использовании)
                self.remove_watermark()
                
                # ВАЖНО: Сохраняем чертеж после Rebuild!
                try:
                    self.logger.info("Сохранение чертежа после Rebuild...")
                    if self.current_document:
                        self.current_document.Save()
                        time.sleep(0.5)  # Даем время на сохранение
                        self.logger.info("Чертеж сохранен после Rebuild")
                except Exception as e:
                    self.logger.warning(f"Ошибка сохранения после Rebuild: {e}")
            
            # Получаем API 5 для работы с растровым форматом
            api5_obj = get_dynamic_dispatch("Kompas.Application.5")
            document_2d = api5_obj.ActiveDocument2D
            
            if not document_2d:
                result['error'] = "Не удалось получить интерфейс 2D документа"
                return result
            
            # Настройка параметров растрового формата
            raster_params = document_2d.RasterFormatParam()
            if not raster_params:
                result['error'] = "Не удалось получить параметры растрового формата"
                return result
            
            # Устанавливаем параметры
            try:
                if format_type.upper() in self.supported_formats:
                    raster_params.Format = self.supported_formats[format_type.upper()]
                    self.logger.info(f"Формат установлен: {format_type}")
                else:
                    self.logger.warning(f"Неподдерживаемый формат {format_type}, используем PNG")
                    raster_params.Format = self.supported_formats['PNG']
            except Exception as e:
                self.logger.warning(f"Не удалось установить формат: {e}")
            
            try:
                raster_params.colorBPP = 24  # 24-битный цвет
                self.logger.info("Цветность установлена: 24 бита")
            except Exception as e:
                self.logger.warning(f"Не удалось установить цветность: {e}")
            
            try:
                raster_params.extResolution = resolution
                self.logger.info(f"Разрешение установлено: {resolution} DPI")
            except Exception as e:
                self.logger.warning(f"Не удалось установить разрешение: {e}")
            
            try:
                raster_params.Init()
                self.logger.info("Параметры инициализированы")
            except Exception as e:
                self.logger.warning(f"Не удалось инициализировать параметры: {e}")
            
            # Сохранение в растровый формат
            try:
                format_code = self.supported_formats.get(format_type.upper(), self.supported_formats['PNG'])
                success = document_2d.SaveAsToRasterFormat(str(output_path_obj), raster_params)
                self.logger.info(f"SaveAsToRasterFormat выполнен, результат: {success}, формат: {format_code}")
                
                if success:
                    time.sleep(2)  # Даем время на создание файла
                    
                    # Проверяем результат
                    if output_path_obj.exists():
                        file_size = output_path_obj.stat().st_size
                        result['success'] = True
                        result['output_file'] = str(output_path_obj)
                        result['file_size'] = file_size
                        
                        self.logger.info(f"Изображение создано: {output_path_obj.name} ({file_size:,} байт)")
                    else:
                        result['error'] = "Файл изображения не создался"
                else:
                    result['error'] = "SaveAsToRasterFormat вернул False"
                    
            except Exception as e:
                result['error'] = f"Ошибка SaveAsToRasterFormat: {e}"
                self.logger.error(result['error'])
            
            # Закрываем документ
            self.close_document()
            
            return result
            
        except Exception as e:
            error_msg = f"Общая ошибка экспорта чертежа: {e}"
            result['error'] = error_msg
            self.logger.error(error_msg)
            return result
        finally:
            self.disconnect_from_kompas()
    
    def force_rebuild_drawing(self) -> bool:
        """Принудительная перестройка чертежа"""
        try:
            self.logger.info("Принудительная перестройка чертежа...")
            
            # Получаем API 5
            api5_obj = get_dynamic_dispatch("Kompas.Application.5")
            document_2d = api5_obj.ActiveDocument2D
            
            if document_2d and hasattr(document_2d, 'ksRebuildDocument'):
                document_2d.ksRebuildDocument()
                time.sleep(1)
                self.logger.info("ksRebuildDocument выполнен успешно")
                return True
            else:
                self.logger.warning("ksRebuildDocument недоступен")
                return False
                
        except Exception as e:
            self.logger.error(f"Ошибка перестройки чертежа: {e}")
            return False
    
    def remove_watermark(self) -> bool:
        """Удаление надписей о некоммерческом использовании"""
        try:
            self.logger.info("Удаление watermark...")
            
            # Получаем API 5
            api5_obj = get_dynamic_dispatch("Kompas.Application.5")
            document_2d = api5_obj.ActiveDocument2D
            
            if not document_2d:
                return False
            
            # Получаем все листы
            views_and_layers_manager = document_2d.ViewsAndLayersManager
            if not views_and_layers_manager:
                return False
            
            views = views_and_layers_manager.Views
            if not views:
                return False
            
            removed_count = 0
            watermark_keywords = [
                "не для коммерческого использования",
                "некоммерческ",
                "non-commercial",
                "educational"
            ]
            
            # Проходим по всем видам
            for view_index in range(views.Count):
                try:
                    view = views.Item(view_index)
                    if not view:
                        continue
                    
                    # Получаем объекты вида
                    drawing_container = view.DrawingContainer
                    if not drawing_container:
                        continue
                    
                    # Проходим по всем объектам
                    for obj_index in range(drawing_container.Count - 1, -1, -1):  # Обратный порядок для безопасного удаления
                        try:
                            obj = drawing_container.Item(obj_index)
                            if not obj:
                                continue
                            
                            # Проверяем, является ли объект текстом
                            if hasattr(obj, 'Text'):
                                text_content = str(obj.Text).lower()
                                
                                # Проверяем на наличие ключевых слов watermark
                                for keyword in watermark_keywords:
                                    if keyword in text_content:
                                        try:
                                            drawing_container.Delete(obj_index)
                                            removed_count += 1
                                            self.logger.info(f"  ✓ Удален watermark: {text_content[:50]}...")
                                            break
                                        except:
                                            pass
                        except:
                            continue
                            
                except:
                    continue
            
            if removed_count > 0:
                self.logger.info(f"Удалено watermark объектов: {removed_count}")
            else:
                self.logger.info("Watermark не найден")
            
            return True
                
        except Exception as e:
            self.logger.warning(f"Ошибка удаления watermark: {e}")
            return False
    
    def export_drawing_to_pdf(self, drawing_path: str, output_path: str) -> Dict:
        """
        Экспорт чертежа напрямую в PDF через КОМПАС API
        
        Args:
            drawing_path: Путь к файлу чертежа
            output_path: Путь для сохранения PDF
            
        Returns:
            Dict с результатами экспорта
        """
        result = {
            'success': False,
            'output_file': None,
            'file_size': 0,
            'error': None
        }
        
        try:
            self.logger.info(f"ЭКСПОРТ ЧЕРТЕЖА В PDF: {Path(drawing_path).name}")
            self.logger.info("=" * 40)
            
            # Подключение к КОМПАС
            if not self.connect_to_kompas(force_reconnect=True):
                result['error'] = "Не удалось подключиться к КОМПАС-3D"
                return result
            
            if not Path(drawing_path).exists():
                result['error'] = f"Файл чертежа не найден"
                return result
            
            # Подготовка пути вывода
            output_path_obj = Path(output_path)
            output_path_obj.parent.mkdir(parents=True, exist_ok=True)
            
            # Удаляем существующий файл
            if output_path_obj.exists():
                output_path_obj.unlink()
            
            # Открываем чертеж
            if not self.open_document(drawing_path):
                result['error'] = "Не удалось открыть чертеж"
                return result
            
            # Перестройка чертежа
            if not self.force_rebuild_drawing():
                self.logger.warning("Не удалось перестроить чертеж")
            
            # Получаем API 5 для SaveAs
            api5_obj = get_dynamic_dispatch("Kompas.Application.5")
            document_2d = api5_obj.ActiveDocument2D
            
            if not document_2d:
                result['error'] = "Не удалось получить интерфейс документа"
                return result
            
            # Сохраняем в PDF через Converter
            try:
                self.logger.info("Сохранение в PDF через Converter...")
                
                # Получаем API 7 для доступа к Converter (используем dynamic dispatch)
                kompas_api7 = get_dynamic_dispatch("Kompas.Application.7")
                
                # Путь к Pdf2d.dll (обычно в папке Bin КОМПАСа)
                # Пытаемся найти КОМПАС в стандартных местах
                import os
                pdf_dll_paths = [
                    r"C:\Program Files\ASCON\KOMPAS-3D v23\Bin\Pdf2d.dll",
                    r"C:\Program Files (x86)\ASCON\KOMPAS-3D v23\Bin\Pdf2d.dll",
                    r"C:\Program Files\ASCON\KOMPAS-3D v22\Bin\Pdf2d.dll",
                    r"C:\Program Files\ASCON\KOMPAS-3D v21\Bin\Pdf2d.dll",
                    r"C:\Program Files\ASCON\KOMPAS-3D v20\Bin\Pdf2d.dll",
                    r"C:\Program Files\ASCON\KOMPAS-3D v19\Bin\Pdf2d.dll",
                    r"C:\Program Files (x86)\ASCON\KOMPAS-3D v22\Bin\Pdf2d.dll",
                    r"C:\Program Files (x86)\ASCON\KOMPAS-3D v21\Bin\Pdf2d.dll",
                ]
                
                pdf_dll_path = None
                for path in pdf_dll_paths:
                    if os.path.exists(path):
                        pdf_dll_path = path
                        break
                
                if not pdf_dll_path:
                    # Пытаемся получить путь из реестра или переменных окружения
                    result['error'] = "Не найден Pdf2d.dll. Проверьте установку КОМПАС-3D"
                    self.logger.error(result['error'])
                    return result
                
                self.logger.info(f"Используется Pdf2d.dll: {pdf_dll_path}")
                
                # Создаем конвертер
                converter = kompas_api7.Converter(pdf_dll_path)
                if not converter:
                    result['error'] = "Не удалось создать конвертер PDF"
                    return result
                
                # Получаем параметры PDF
                pdf_params = converter.ConverterParameters(0, 0, 0)  # 0 = текущий документ
                if pdf_params:
                    # Настройка параметров (цветной PDF)
                    try:
                        pdf_params.ColorType = 3  # Цветной
                    except:
                        pass
                
                # Конвертируем
                convert_result = converter.Convert(str(output_path_obj), 0, 0)
                
                if convert_result:
                    time.sleep(2)  # Даем время на создание файла
                    
                    # Проверяем результат
                    if output_path_obj.exists():
                        file_size = output_path_obj.stat().st_size
                        result['success'] = True
                        result['output_file'] = str(output_path_obj)
                        result['file_size'] = file_size
                        
                        self.logger.info(f"PDF создан: {output_path_obj.name} ({file_size:,} байт)")
                    else:
                        result['error'] = "Файл PDF не создался"
                else:
                    result['error'] = "Converter.Convert вернул False"
                    
            except Exception as e:
                result['error'] = f"Ошибка сохранения PDF: {e}"
                self.logger.error(result['error'])
            
            # Закрываем документ
            self.close_document()
            
            return result
            
        except Exception as e:
            error_msg = f"Общая ошибка экспорта в PDF: {e}"
            result['error'] = error_msg
            self.logger.error(error_msg)
            return result
        finally:
            self.disconnect_from_kompas()
    
    def export_all_drawings(self, project_path: str, output_folder: str = None, 
                           format_type: str = 'PNG', resolution: int = 300) -> Dict:
        """
        Экспорт всех чертежей проекта
        
        Args:
            project_path: Путь к проекту
            output_folder: Папка для сохранения (если None, то рядом с чертежами)
            format_type: Тип формата
            resolution: Разрешение в DPI
            
        Returns:
            Dict с результатами экспорта
        """
        result = {
            'success': False,
            'total_drawings': 0,
            'exported_drawings': 0,
            'failed_drawings': 0,
            'exported_files': [],
            'errors': []
        }
        
        try:
            self.logger.info("ЭКСПОРТ ВСЕХ ЧЕРТЕЖЕЙ ПРОЕКТА")
            self.logger.info("=" * 35)
            
            # Находим чертежи
            drawing_files = self.find_drawing_files(project_path)
            
            if not drawing_files:
                result['errors'].append("Чертежи не найдены")
                return result
            
            result['total_drawings'] = len(drawing_files)
            
            # Подготавливаем папку вывода
            if output_folder:
                output_folder_obj = Path(output_folder)
                output_folder_obj.mkdir(parents=True, exist_ok=True)
            else:
                output_folder_obj = Path(project_path)
            
            # Экспортируем каждый чертеж
            for drawing_path in drawing_files:
                try:
                    drawing_file = Path(drawing_path)
                    
                    # Создаем имя выходного файла
                    output_filename = drawing_file.with_suffix(f'.{format_type.lower()}')
                    output_path = output_folder_obj / output_filename.name
                    
                    # Экспортируем чертеж
                    export_result = self.export_drawing_to_image(
                        drawing_path, str(output_path), format_type, resolution
                    )
                    
                    if export_result['success']:
                        result['exported_drawings'] += 1
                        result['exported_files'].append({
                            'source': drawing_path,
                            'output': export_result['output_file'],
                            'size': export_result['file_size']
                        })
                        self.logger.info(f"Чертеж экспортирован: {output_filename.name}")
                    else:
                        result['failed_drawings'] += 1
                        result['errors'].append(f"Ошибка экспорта {drawing_file.name}: {export_result['error']}")
                        
                except Exception as e:
                    result['failed_drawings'] += 1
                    error_msg = f"Ошибка обработки чертежа {drawing_path}: {e}"
                    result['errors'].append(error_msg)
                    self.logger.error(error_msg)
            
            result['success'] = result['exported_drawings'] > 0
            result['exported_count'] = result['exported_drawings']
            result['total_count'] = result['total_drawings']
            
            self.logger.info(f"Экспорт завершен:")
            self.logger.info(f"  Всего чертежей: {result['total_drawings']}")
            self.logger.info(f"  Экспортировано: {result['exported_drawings']}")
            self.logger.info(f"  Ошибок: {result['failed_drawings']}")
            
            return result
            
        except Exception as e:
            error_msg = f"Общая ошибка экспорта чертежей: {e}"
            result['error'] = error_msg
            self.logger.error(error_msg)
            return result
