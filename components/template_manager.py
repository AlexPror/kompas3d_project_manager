#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
📚 Менеджер шаблонов проектов
Управление библиотекой шаблонов для разных типов конвекторов
"""

import json
import logging
import shutil
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional


class TemplateManager:
    """Менеджер шаблонов проектов"""
    
    def __init__(self, templates_dir: str = "templates"):
        self.logger = logging.getLogger('TemplateManager')
        self.templates_dir = Path(templates_dir)
        self.templates_dir.mkdir(exist_ok=True)
        
        # Файл базы данных шаблонов
        self.db_file = self.templates_dir / 'templates_db.json'
        
        # Загружаем базу данных
        self.db = self._load_database()
    
    def _load_database(self) -> Dict:
        """Загружает базу данных шаблонов"""
        if self.db_file.exists():
            try:
                with open(self.db_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                self.logger.error(f"Ошибка загрузки базы шаблонов: {e}")
                return {'templates': [], 'version': '1.0'}
        else:
            return {'templates': [], 'version': '1.0'}
    
    def _save_database(self):
        """Сохраняет базу данных шаблонов"""
        try:
            with open(self.db_file, 'w', encoding='utf-8') as f:
                json.dump(self.db, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            self.logger.error(f"Ошибка сохранения базы шаблонов: {e}")
            return False
    
    def list_templates(self) -> List[Dict]:
        """
        Возвращает список доступных шаблонов
        
        Returns:
            list: Список шаблонов с метаданными
        """
        return self.db.get('templates', [])
    
    def get_template(self, template_id: str) -> Optional[Dict]:
        """
        Получает информацию о шаблоне по ID
        
        Args:
            template_id: ID шаблона
            
        Returns:
            dict: Информация о шаблоне или None
        """
        for template in self.db['templates']:
            if template['id'] == template_id:
                return template
        return None
    
    def add_template_from_project(self, 
                                  project_path: str,
                                  template_name: str,
                                  description: str = "",
                                  parameters: Dict = None,
                                  tags: List[str] = None) -> Dict:
        """
        Создает шаблон из существующего проекта
        
        Args:
            project_path: Путь к проекту
            template_name: Название шаблона
            description: Описание шаблона
            parameters: Параметры (H, B1, L1 и т.д.)
            tags: Теги для поиска
            
        Returns:
            dict: Результат операции
        """
        try:
            self.logger.info(f"Создание шаблона '{template_name}' из проекта...")
            
            project_path = Path(project_path)
            
            if not project_path.exists():
                return {'success': False, 'error': 'Проект не найден'}
            
            # Генерируем уникальный ID
            template_id = f"tpl_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            
            # Создаем папку для шаблона
            template_folder = self.templates_dir / template_id
            template_folder.mkdir(exist_ok=True)
            
            # Копируем проект
            self.logger.info(f"Копирование файлов в {template_folder}...")
            copied_count = 0
            
            for item in project_path.iterdir():
                # Пропускаем временные файлы, архивы и служебные папки
                if item.name.startswith('~$'):
                    continue
                if item.suffix in ['.bak', '.tmp']:
                    continue
                if item.name.endswith('.rar') or item.name.endswith('.zip'):
                    continue
                if 'A-' in item.name and item.is_dir():  # Папки с номером заказа
                    continue
                if item.name in ['BMP', 'PDF', 'Исходники разверток с номерами']:
                    continue
                
                try:
                    target = template_folder / item.name
                    
                    if item.is_file():
                        shutil.copy2(item, target)
                        copied_count += 1
                    elif item.is_dir() and item.name == 'DXF':
                        # DXF папку тоже не копируем - она генерируется
                        pass
                    elif item.is_dir():
                        shutil.copytree(item, target)
                        copied_count += 1
                        
                except Exception as e:
                    self.logger.warning(f"Пропуск {item.name}: {e}")
            
            self.logger.info(f"Скопировано элементов: {copied_count}")
            
            # Создаем метаданные шаблона
            template_info = {
                'id': template_id,
                'name': template_name,
                'description': description,
                'created_at': datetime.now().isoformat(),
                'parameters': parameters or {},
                'tags': tags or [],
                'path': str(template_folder),
                'source_project': str(project_path),
                'file_count': copied_count,
                'usage_count': 0
            }
            
            # Добавляем в базу
            self.db['templates'].append(template_info)
            self._save_database()
            
            # Сохраняем README для шаблона
            self._create_template_readme(template_folder, template_info)
            
            self.logger.info(f"✅ Шаблон '{template_name}' создан успешно!")
            self.logger.info(f"   ID: {template_id}")
            self.logger.info(f"   Путь: {template_folder}")
            
            return {
                'success': True,
                'template_id': template_id,
                'template_info': template_info
            }
            
        except Exception as e:
            self.logger.error(f"Ошибка создания шаблона: {e}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def _create_template_readme(self, template_folder: Path, template_info: Dict):
        """Создает README файл для шаблона"""
        readme_content = f"""# {template_info['name']}

## Описание
{template_info.get('description', 'Нет описания')}

## Параметры
"""
        
        params = template_info.get('parameters', {})
        if params:
            for key, value in params.items():
                readme_content += f"- **{key}**: {value}\n"
        else:
            readme_content += "Параметры не указаны\n"
        
        readme_content += f"""
## Теги
{', '.join(template_info.get('tags', ['нет тегов']))}

## Информация
- **ID**: `{template_info['id']}`
- **Создан**: {template_info['created_at']}
- **Источник**: `{template_info.get('source_project', 'не указан')}`
- **Файлов**: {template_info.get('file_count', 0)}

---
*Шаблон создан автоматически КОМПАС-3D Project Manager*
"""
        
        readme_file = template_folder / 'README.md'
        with open(readme_file, 'w', encoding='utf-8') as f:
            f.write(readme_content)
    
    def delete_template(self, template_id: str) -> Dict:
        """
        Удаляет шаблон
        
        Args:
            template_id: ID шаблона
            
        Returns:
            dict: Результат операции
        """
        try:
            template = self.get_template(template_id)
            
            if not template:
                return {'success': False, 'error': 'Шаблон не найден'}
            
            # Удаляем папку
            template_folder = Path(template['path'])
            if template_folder.exists():
                shutil.rmtree(template_folder)
            
            # Удаляем из базы
            self.db['templates'] = [t for t in self.db['templates'] if t['id'] != template_id]
            self._save_database()
            
            self.logger.info(f"✅ Шаблон '{template['name']}' удален")
            
            return {'success': True}
            
        except Exception as e:
            self.logger.error(f"Ошибка удаления шаблона: {e}")
            return {'success': False, 'error': str(e)}
    
    def update_template_usage(self, template_id: str):
        """Обновляет счетчик использования шаблона"""
        for template in self.db['templates']:
            if template['id'] == template_id:
                template['usage_count'] = template.get('usage_count', 0) + 1
                template['last_used'] = datetime.now().isoformat()
                self._save_database()
                break
    
    def search_templates(self, query: str) -> List[Dict]:
        """
        Поиск шаблонов по названию, описанию или тегам
        
        Args:
            query: Поисковый запрос
            
        Returns:
            list: Найденные шаблоны
        """
        query_lower = query.lower()
        results = []
        
        for template in self.db['templates']:
            # Поиск в названии
            if query_lower in template['name'].lower():
                results.append(template)
                continue
            
            # Поиск в описании
            if query_lower in template.get('description', '').lower():
                results.append(template)
                continue
            
            # Поиск в тегах
            if any(query_lower in tag.lower() for tag in template.get('tags', [])):
                results.append(template)
                continue
        
        return results
    
    def get_template_statistics(self) -> Dict:
        """Возвращает статистику по шаблонам"""
        templates = self.db['templates']
        
        if not templates:
            return {
                'total': 0,
                'most_used': None,
                'newest': None
            }
        
        # Самый используемый
        most_used = max(templates, key=lambda t: t.get('usage_count', 0))
        
        # Самый новый
        newest = max(templates, key=lambda t: t['created_at'])
        
        return {
            'total': len(templates),
            'most_used': {
                'name': most_used['name'],
                'usage_count': most_used.get('usage_count', 0)
            },
            'newest': {
                'name': newest['name'],
                'created_at': newest['created_at']
            }
        }

