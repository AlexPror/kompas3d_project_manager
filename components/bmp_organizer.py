#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
📁 Организатор BMP файлов
Создает папку BMP и перемещает туда все BMP файлы из корня проекта
"""

import logging
import shutil
from pathlib import Path
from typing import Dict


class BmpOrganizer:
    """Организатор BMP файлов в проекте"""
    
    def __init__(self):
        self.logger = logging.getLogger('BmpOrganizer')
    
    def organize_bmp_files(self, project_path: str) -> Dict:
        """
        Организует BMP файлы - создает папку BMP и перемещает туда файлы
        
        Args:
            project_path: Путь к проекту
            
        Returns:
            dict: Результат операции
        """
        try:
            project_path = Path(project_path)
            
            if not project_path.exists():
                return {
                    'success': False,
                    'error': 'Папка проекта не найдена'
                }
            
            self.logger.info(f"Организация BMP файлов в проекте: {project_path.name}")
            
            # Ищем все BMP файлы в корне проекта
            bmp_files = list(project_path.glob('*.bmp'))
            
            if not bmp_files:
                self.logger.info("BMP файлов в корне проекта не найдено")
                return {
                    'success': True,
                    'moved_count': 0,
                    'message': 'Нет BMP файлов для организации'
                }
            
            # Создаем папку BMP
            bmp_folder = project_path / 'BMP'
            bmp_folder.mkdir(exist_ok=True)
            self.logger.info(f"Создана/проверена папка: {bmp_folder}")
            
            # Перемещаем файлы
            moved_count = 0
            errors = []
            
            for bmp_file in bmp_files:
                try:
                    target_path = bmp_folder / bmp_file.name
                    
                    # Если файл уже существует в целевой папке, удаляем старый
                    if target_path.exists():
                        target_path.unlink()
                    
                    shutil.move(str(bmp_file), str(target_path))
                    moved_count += 1
                    self.logger.info(f"  ✓ {bmp_file.name} → BMP/")
                    
                except Exception as e:
                    error_msg = f"Ошибка перемещения {bmp_file.name}: {e}"
                    self.logger.error(f"  ✗ {error_msg}")
                    errors.append(error_msg)
            
            result = {
                'success': True,
                'moved_count': moved_count,
                'total_count': len(bmp_files),
                'bmp_folder': str(bmp_folder),
                'errors': errors
            }
            
            if moved_count > 0:
                self.logger.info(f"✅ Перемещено BMP файлов: {moved_count}/{len(bmp_files)}")
            
            if errors:
                self.logger.warning(f"⚠️ Ошибок при перемещении: {len(errors)}")
            
            return result
            
        except Exception as e:
            self.logger.error(f"Ошибка организации BMP файлов: {e}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def get_bmp_folder_path(self, project_path: str) -> str:
        """Возвращает путь к папке BMP (создает если нужно)"""
        bmp_folder = Path(project_path) / 'BMP'
        bmp_folder.mkdir(exist_ok=True)
        return str(bmp_folder)

