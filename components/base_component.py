"""
Базовый компонент для работы с КОМПАС-3D
"""
import logging
import pythoncom
import sys
import os
import shutil
from typing import Dict, Optional

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')


def clear_kompas_cache():
    """Очистка кэша win32com для KOMPAS для устранения ошибок CLSIDToClassMap"""
    try:
        import win32com
        gen_py_path = os.path.join(os.path.dirname(win32com.__file__), 'gen_py')
        if os.path.exists(gen_py_path):
            for item in os.listdir(gen_py_path):
                item_path = os.path.join(gen_py_path, item)
                # Удаляем директории с KOMPAS GUID
                if os.path.isdir(item_path) and '0422828C' in item.upper():
                    try:
                        shutil.rmtree(item_path)
                    except:
                        pass
    except:
        pass


def get_dynamic_dispatch(prog_id):
    """Получение COM-объекта через dynamic dispatch (без использования кэша)"""
    from win32com.client import dynamic
    return dynamic.Dispatch(prog_id)

class BaseKompasComponent:
    """Базовый класс для всех компонентов КОМПАС-3D"""
    
    def __init__(self):
        self.logger = logging.getLogger(self.__class__.__name__)
        self.application = None
        self.connected = False
    
    def connect_to_kompas(self, force_reconnect: bool = False) -> bool:
        """Подключение к КОМПАС-3D
        
        Args:
            force_reconnect: Принудительное переподключение даже если уже подключен
        """
        try:
            # Если требуется переподключение или еще не подключались
            if force_reconnect or not self.connected:
                # Очищаем кэш win32com для устранения ошибок CLSIDToClassMap
                clear_kompas_cache()
                
                # Инициализируем COM (безопасно вызывать повторно)
                try:
                    pythoncom.CoInitialize()
                except:
                    pass  # Уже инициализирован
                
                self.logger.info("Подключение к КОМПАС-3D v7...")
                
                # Используем dynamic dispatch для обхода проблем с кэшем
                try:
                    self.application = get_dynamic_dispatch("Kompas.Application.7")
                except:
                    # Если не получилось, пробуем обычный Dispatch
                    from win32com.client import Dispatch
                    self.application = Dispatch("Kompas.Application.7")
                
                self.application.Visible = True
                self.logger.info("Подключение к КОМПАС-3D успешно")
                self.connected = True
                return True
            
            # Если уже подключены, проверяем что соединение живое
            if self.connected and self.application:
                try:
                    # Проверка - пытаемся получить доступ к объекту
                    _ = self.application.Visible
                    return True
                except:
                    # Соединение потеряно, переподключаемся
                    self.logger.warning("Соединение с КОМПАС-3D потеряно, переподключение...")
                    self.connected = False
                    return self.connect_to_kompas(force_reconnect=True)
            
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка подключения к КОМПАС-3D: {e}")
            self.connected = False
            return False
    
    def disconnect_from_kompas(self):
        """Отключение от КОМПАС-3D"""
        try:
            if self.application:
                self.application = None
            if self.connected:
                pythoncom.CoUninitialize()
                self.connected = False
                self.logger.info("Отключение от КОМПАС-3D")
        except Exception as e:
            self.logger.error(f"Ошибка отключения от КОМПАС-3D: {e}")
    
    def is_connected(self) -> bool:
        """Проверка подключения"""
        return self.connected and self.application is not None
    
    def get_active_document(self):
        """Получение активного документа"""
        if not self.is_connected():
            return None
        try:
            return self.application.ActiveDocument
        except Exception as e:
            self.logger.error(f"Ошибка получения активного документа: {e}")
            return None
    
    def close_all_documents(self):
        """Закрытие всех открытых документов"""
        try:
            if not self.is_connected():
                return False
                
            documents = self.application.Documents
            doc_count = documents.Count
            
            for i in range(doc_count - 1, -1, -1):
                try:
                    doc = documents.Item(i)
                    if doc:
                        self.logger.info(f"Закрытие документа: {doc.Name}")
                        doc.Close(False)
                except Exception as e:
                    self.logger.warning(f"Ошибка закрытия документа {i}: {e}")
            
            import time
            time.sleep(1)
            self.logger.info("Все документы закрыты")
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка закрытия документов: {e}")
            return False
    
    def open_document(self, file_path: str) -> bool:
        """Открытие документа"""
        try:
            if not self.is_connected():
                return False
                
            from pathlib import Path
            if not Path(file_path).exists():
                self.logger.error(f"Файл не найден: {file_path}")
                return False
            
            doc = self.application.Documents.Open(file_path, False)
            if not doc:
                self.logger.error("Не удалось открыть документ")
                return False
            
            self.application.ActiveDocument = doc
            self.logger.info(f"Документ открыт: {doc.Name}")
            
            import time
            time.sleep(1)
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка открытия документа: {e}")
            return False
    
    def close_document(self, save=False):
        """Закрытие активного документа
        
        Args:
            save: Сохранить изменения перед закрытием (по умолчанию False)
        """
        try:
            if not self.is_connected():
                return False
                
            active_doc = self.get_active_document()
            if active_doc:
                doc_name = active_doc.Name
                active_doc.Close(save)
                self.logger.info(f"Документ закрыт: {doc_name} (save={save})")
                return True
            return False
            
        except Exception as e:
            self.logger.error(f"Ошибка закрытия документа: {e}")
            return False
