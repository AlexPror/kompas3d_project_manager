#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🔍 Анализатор проектов КОМПАС-3D для создания шаблонов
Изучает структуру проекта и извлекает переменные
"""

import sys
import json
import logging
from pathlib import Path
from datetime import datetime
from components.base_component import BaseKompasComponent

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)-8s | %(message)s',
    datefmt='%H:%M:%S',
    handlers=[
        logging.FileHandler('template_analysis.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)


class ProjectTemplateAnalyzer(BaseKompasComponent):
    """Анализатор проекта для создания шаблона"""
    
    def __init__(self):
        super().__init__()
        self.logger = logging.getLogger('TemplateAnalyzer')
    
    def analyze_project(self, project_path: str) -> dict:
        """
        Полный анализ проекта
        
        Args:
            project_path: Путь к папке проекта
            
        Returns:
            dict: Детальная информация о проекте
        """
        self.logger.info("="*60)
        self.logger.info("🔍 АНАЛИЗ ПРОЕКТА ДЛЯ СОЗДАНИЯ ШАБЛОНА")
        self.logger.info("="*60)
        
        project_path = Path(project_path)
        
        if not project_path.exists():
            return {'success': False, 'error': 'Папка проекта не найдена'}
        
        analysis = {
            'success': True,
            'project_path': str(project_path),
            'project_name': project_path.name,
            'analyzed_at': datetime.now().isoformat(),
            'files': {},
            'structure': {},
            'variables': {},
            'recommendations': []
        }
        
        # Анализ файловой структуры
        self.logger.info(f"\n📁 Проект: {project_path.name}")
        analysis['files'] = self._analyze_files(project_path)
        
        # Анализ структуры (сборка + детали)
        analysis['structure'] = self._analyze_structure(analysis['files'])
        
        # Извлечение переменных из сборки
        assembly_file = self._find_main_assembly(project_path)
        if assembly_file:
            analysis['variables'] = self._extract_variables(assembly_file)
        
        # Формирование рекомендаций
        analysis['recommendations'] = self._generate_recommendations(analysis)
        
        # Вывод итогов
        self._print_analysis_summary(analysis)
        
        return analysis
    
    def _analyze_files(self, project_path: Path) -> dict:
        """Анализ всех файлов в проекте"""
        files = {
            'assembly_files': [],
            'part_files': [],
            'drawing_files': [],
            'unfolding_files': [],
            'dxf_files': [],
            'pdf_files': [],
            'bmp_files': [],
            'other_files': [],
            'temp_files': []
        }
        
        for file in project_path.rglob('*'):
            if not file.is_file():
                continue
            
            suffix = file.suffix.lower()
            name = file.name
            relative = file.relative_to(project_path)
            
            file_info = {
                'name': name,
                'path': str(file),
                'relative_path': str(relative),
                'size': file.stat().st_size
            }
            
            # Временные файлы
            if name.startswith('~$') or suffix in ['.bak', '.tmp']:
                files['temp_files'].append(file_info)
            # Сборки
            elif suffix == '.a3d':
                files['assembly_files'].append(file_info)
            # Детали
            elif suffix == '.m3d':
                files['part_files'].append(file_info)
            # Чертежи
            elif suffix == '.cdw':
                # Различаем обычные чертежи и развертки
                if 'развертка' in name.lower() or 'razvertka' in name.lower():
                    files['unfolding_files'].append(file_info)
                else:
                    files['drawing_files'].append(file_info)
            # Экспортные файлы
            elif suffix == '.dxf':
                files['dxf_files'].append(file_info)
            elif suffix == '.pdf':
                files['pdf_files'].append(file_info)
            elif suffix == '.bmp':
                files['bmp_files'].append(file_info)
            else:
                files['other_files'].append(file_info)
        
        # Статистика
        self.logger.info(f"\n📊 Найдено файлов:")
        self.logger.info(f"  🔧 Сборок (.a3d):         {len(files['assembly_files'])}")
        self.logger.info(f"  ⚙️  Деталей (.m3d):        {len(files['part_files'])}")
        self.logger.info(f"  📐 Чертежей (.cdw):       {len(files['drawing_files'])}")
        self.logger.info(f"  📏 Разверток (.cdw):      {len(files['unfolding_files'])}")
        self.logger.info(f"  📄 DXF файлов:            {len(files['dxf_files'])}")
        self.logger.info(f"  📑 PDF файлов:            {len(files['pdf_files'])}")
        self.logger.info(f"  🖼️  BMP файлов:            {len(files['bmp_files'])}")
        self.logger.info(f"  ⚠️  Временных файлов:     {len(files['temp_files'])}")
        
        return files
    
    def _analyze_structure(self, files: dict) -> dict:
        """Анализ структуры проекта"""
        structure = {
            'main_assembly': None,
            'parts': [],
            'drawing_mapping': {}
        }
        
        # Находим главную сборку (в корне, не в подпапках)
        for assembly in files['assembly_files']:
            if '\\' not in assembly['relative_path'] and '/' not in assembly['relative_path']:
                structure['main_assembly'] = assembly
                break
        
        # Анализируем детали (ищем паттерн: NNN - Название)
        import re
        part_pattern = re.compile(r'^(\d{3})\s*-\s*(.+)\.m3d$', re.IGNORECASE)
        
        for part in files['part_files']:
            match = part_pattern.match(part['name'])
            if match:
                part_number = match.group(1)
                part_name = match.group(2)
                
                part_info = {
                    'number': part_number,
                    'name': part_name,
                    'file': part['name'],
                    'path': part['path'],
                    'drawings': [],
                    'unfoldings': []
                }
                
                # Ищем соответствующие чертежи
                drawing_pattern = f"{part_number}0"  # 001 -> 0010
                for drawing in files['drawing_files']:
                    if drawing['name'].startswith(drawing_pattern):
                        part_info['drawings'].append(drawing['name'])
                
                # Ищем развертки
                for unfolding in files['unfolding_files']:
                    if part_number in unfolding['name']:
                        part_info['unfoldings'].append(unfolding['name'])
                
                structure['parts'].append(part_info)
        
        # Сортируем детали по номеру
        structure['parts'].sort(key=lambda x: x['number'])
        
        self.logger.info(f"\n🔧 Структура проекта:")
        if structure['main_assembly']:
            self.logger.info(f"  📦 Главная сборка: {structure['main_assembly']['name']}")
        
        self.logger.info(f"\n  ⚙️  Детали ({len(structure['parts'])}):")
        for part in structure['parts']:
            self.logger.info(f"    {part['number']} - {part['name']}")
            if part['drawings']:
                self.logger.info(f"         └─ Чертежи: {', '.join(part['drawings'])}")
            if part['unfoldings']:
                self.logger.info(f"         └─ Развертки: {', '.join(part['unfoldings'])}")
        
        return structure
    
    def _find_main_assembly(self, project_path: Path) -> Path:
        """Находит главную сборку (.a3d) в корне проекта"""
        for file in project_path.glob('*.a3d'):
            if not file.name.startswith('~$'):
                return file
        return None
    
    def _extract_variables(self, assembly_file: Path) -> dict:
        """Извлекает переменные из сборки"""
        self.logger.info(f"\n🔍 Извлечение переменных из сборки...")
        
        variables = {
            'extracted': False,
            'base_variables': {},
            'dependent_variables': {},
            'all_variables': {}
        }
        
        try:
            if not self.connect_to_kompas():
                self.logger.error("Не удалось подключиться к КОМПАС-3D")
                return variables
            
            # Закрываем все документы
            self.close_all_documents()
            
            # Открываем сборку
            if not self.open_document(str(assembly_file)):
                self.logger.error("Не удалось открыть сборку")
                return variables
            
            doc = self.get_active_document()
            if not doc:
                return variables
            
            # Получаем переменные
            try:
                # Интерфейс переменных (константы КОМПАС API)
                vars_mng = doc.ParametersManager
                
                # Базовые переменные (задаются пользователем)
                base_vars = ['H', 'B1', 'L1', 'B2', 'A2']
                
                self.logger.info(f"\n📊 Переменные в сборке:")
                
                # Пытаемся получить все переменные
                for i in range(vars_mng.Count):
                    try:
                        param = vars_mng.Item(i)
                        var_name = param.Name
                        var_value = param.Value
                        
                        var_info = {
                            'name': var_name,
                            'value': var_value,
                            'type': 'base' if var_name in base_vars else 'dependent'
                        }
                        
                        variables['all_variables'][var_name] = var_info
                        
                        if var_name in base_vars:
                            variables['base_variables'][var_name] = var_value
                            self.logger.info(f"  ✓ {var_name} = {var_value} (базовая)")
                        else:
                            variables['dependent_variables'][var_name] = var_value
                            self.logger.info(f"  • {var_name} = {var_value}")
                        
                    except Exception as e:
                        self.logger.warning(f"Не удалось получить переменную {i}: {e}")
                
                variables['extracted'] = True
                
            except Exception as e:
                self.logger.error(f"Ошибка извлечения переменных: {e}")
            
            # Закрываем документ
            self.close_document(save=False)
            
        except Exception as e:
            self.logger.error(f"Ошибка при работе с КОМПАС: {e}")
        finally:
            self.disconnect_from_kompas()
        
        return variables
    
    def _generate_recommendations(self, analysis: dict) -> list:
        """Генерирует рекомендации по улучшению структуры"""
        recommendations = []
        
        files = analysis['files']
        structure = analysis['structure']
        
        # Проверка на BMP файлы в корне
        if files['bmp_files']:
            bmp_in_root = [f for f in files['bmp_files'] if '\\' not in f['relative_path']]
            if bmp_in_root:
                recommendations.append({
                    'type': 'file_organization',
                    'priority': 'high',
                    'message': f"Найдено {len(bmp_in_root)} BMP файлов в корне проекта. Рекомендуется создать папку 'BMP' и переместить их туда.",
                    'action': 'create_bmp_folder'
                })
        
        # Проверка на временные файлы
        if files['temp_files']:
            recommendations.append({
                'type': 'cleanup',
                'priority': 'medium',
                'message': f"Найдено {len(files['temp_files'])} временных файлов. Рекомендуется очистка при копировании.",
                'action': 'cleanup_temp_files'
            })
        
        # Проверка наличия всех необходимых файлов
        if not structure['main_assembly']:
            recommendations.append({
                'type': 'structure',
                'priority': 'critical',
                'message': "Главная сборка (.a3d) не найдена в корне проекта",
                'action': 'check_structure'
            })
        
        # Проверка соответствия деталей и чертежей
        parts_without_drawings = [p for p in structure['parts'] if not p['drawings']]
        if parts_without_drawings:
            recommendations.append({
                'type': 'completeness',
                'priority': 'medium',
                'message': f"У {len(parts_without_drawings)} деталей нет чертежей",
                'action': 'create_missing_drawings',
                'details': [p['name'] for p in parts_without_drawings]
            })
        
        # Проверка переменных
        if not analysis['variables'].get('extracted'):
            recommendations.append({
                'type': 'variables',
                'priority': 'high',
                'message': "Не удалось извлечь переменные из сборки",
                'action': 'check_variables'
            })
        elif len(analysis['variables'].get('base_variables', {})) < 3:
            recommendations.append({
                'type': 'variables',
                'priority': 'high',
                'message': "Не все базовые переменные (H, B1, L1) найдены в сборке",
                'action': 'check_base_variables'
            })
        
        return recommendations
    
    def _print_analysis_summary(self, analysis: dict):
        """Выводит итоговую сводку анализа"""
        self.logger.info("\n" + "="*60)
        self.logger.info("📋 ИТОГОВАЯ СВОДКА")
        self.logger.info("="*60)
        
        # Основная информация
        self.logger.info(f"\n📦 Проект: {analysis['project_name']}")
        self.logger.info(f"📁 Путь: {analysis['project_path']}")
        
        # Переменные
        if analysis['variables'].get('extracted'):
            base_vars = analysis['variables']['base_variables']
            self.logger.info(f"\n🔧 Базовые параметры:")
            for var, value in base_vars.items():
                self.logger.info(f"  {var} = {value}")
        
        # Структура
        structure = analysis['structure']
        self.logger.info(f"\n📊 Структура:")
        self.logger.info(f"  Деталей: {len(structure['parts'])}")
        self.logger.info(f"  Чертежей: {len(analysis['files']['drawing_files'])}")
        self.logger.info(f"  Разверток: {len(analysis['files']['unfolding_files'])}")
        self.logger.info(f"  DXF файлов: {len(analysis['files']['dxf_files'])}")
        
        # Рекомендации
        if analysis['recommendations']:
            self.logger.info(f"\n💡 Рекомендации ({len(analysis['recommendations'])}):")
            for i, rec in enumerate(analysis['recommendations'], 1):
                priority_icon = {
                    'critical': '🔴',
                    'high': '🟠',
                    'medium': '🟡',
                    'low': '🟢'
                }.get(rec['priority'], '⚪')
                
                self.logger.info(f"  {priority_icon} {i}. [{rec['type'].upper()}] {rec['message']}")
        
        self.logger.info("\n" + "="*60)
    
    def save_analysis_report(self, analysis: dict, output_file: str = 'template_analysis_report.json'):
        """Сохраняет отчет анализа в JSON"""
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(analysis, f, ensure_ascii=False, indent=2)
            
            self.logger.info(f"\n💾 Отчет сохранен: {output_file}")
            
            # Также создаем markdown версию
            md_file = output_file.replace('.json', '.md')
            self._create_markdown_report(analysis, md_file)
            
            return True
        except Exception as e:
            self.logger.error(f"Ошибка сохранения отчета: {e}")
            return False
    
    def _create_markdown_report(self, analysis: dict, output_file: str):
        """Создает отчет в формате Markdown"""
        md_content = f"""# 📋 Анализ проекта для создания шаблона

## 📦 Информация о проекте

- **Название**: {analysis['project_name']}
- **Путь**: `{analysis['project_path']}`
- **Дата анализа**: {analysis['analyzed_at']}

---

## 📊 Файловая структура

| Тип | Количество |
|-----|------------|
| 🔧 Сборки (.a3d) | {len(analysis['files']['assembly_files'])} |
| ⚙️ Детали (.m3d) | {len(analysis['files']['part_files'])} |
| 📐 Чертежи (.cdw) | {len(analysis['files']['drawing_files'])} |
| 📏 Развертки (.cdw) | {len(analysis['files']['unfolding_files'])} |
| 📄 DXF файлы | {len(analysis['files']['dxf_files'])} |
| 📑 PDF файлы | {len(analysis['files']['pdf_files'])} |
| 🖼️ BMP файлы | {len(analysis['files']['bmp_files'])} |
| ⚠️ Временные файлы | {len(analysis['files']['temp_files'])} |

---

## 🔧 Детали проекта

"""
        
        for part in analysis['structure']['parts']:
            md_content += f"""
### {part['number']} - {part['name']}

- **Файл**: `{part['file']}`
- **Чертежи**: {', '.join(part['drawings']) if part['drawings'] else '❌ Нет'}
- **Развертки**: {', '.join(part['unfoldings']) if part['unfoldings'] else '❌ Нет'}
"""
        
        # Переменные
        if analysis['variables'].get('extracted'):
            md_content += "\n---\n\n## 🔧 Переменные проекта\n\n### Базовые переменные\n\n"
            for var, value in analysis['variables']['base_variables'].items():
                md_content += f"- **{var}** = {value}\n"
            
            md_content += "\n### Зависимые переменные\n\n"
            for var, value in analysis['variables']['dependent_variables'].items():
                md_content += f"- {var} = {value}\n"
        
        # Рекомендации
        if analysis['recommendations']:
            md_content += "\n---\n\n## 💡 Рекомендации\n\n"
            for i, rec in enumerate(analysis['recommendations'], 1):
                priority_emoji = {
                    'critical': '🔴',
                    'high': '🟠',
                    'medium': '🟡',
                    'low': '🟢'
                }.get(rec['priority'], '⚪')
                
                md_content += f"{i}. {priority_emoji} **[{rec['type'].upper()}]** {rec['message']}\n"
        
        md_content += "\n---\n\n*Отчет создан автоматически анализатором проектов КОМПАС-3D*\n"
        
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(md_content)
            self.logger.info(f"💾 Markdown отчет сохранен: {output_file}")
        except Exception as e:
            self.logger.error(f"Ошибка создания Markdown отчета: {e}")


def main():
    """Главная функция"""
    import sys
    
    if len(sys.argv) < 2:
        print("Использование: python project_template_analyzer.py <путь_к_проекту>")
        print("\nПример:")
        print('  python project_template_analyzer.py "C:\\Projects\\ZVD.LITE.110.180.1330"')
        return
    
    project_path = sys.argv[1]
    
    analyzer = ProjectTemplateAnalyzer()
    analysis = analyzer.analyze_project(project_path)
    
    if analysis['success']:
        analyzer.save_analysis_report(analysis)
        print("\n✅ Анализ завершен успешно!")
        print(f"📄 Отчеты сохранены:")
        print(f"  - template_analysis_report.json")
        print(f"  - template_analysis_report.md")
    else:
        print(f"\n❌ Ошибка анализа: {analysis.get('error', 'Unknown')}")


if __name__ == "__main__":
    main()

