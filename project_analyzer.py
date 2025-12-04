"""
Универсальный анализатор проектов КОМПАС-3D и SolidWorks
Сканирует папку с проектами, собирает информацию о структуре и параметрах
Работает БЕЗ открытия файлов - только анализ имен и структуры
"""

import os
import sys
import re
import logging
from pathlib import Path
from collections import defaultdict
from typing import Dict, List, Tuple, Optional
import json
from datetime import datetime

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('project_analysis.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class UniversalProjectAnalyzer:
    """Универсальный анализатор проектов КОМПАС-3D и SolidWorks"""
    
    # Поддерживаемые форматы файлов
    KOMPAS_ASSEMBLY = ['.a3d']
    KOMPAS_PART = ['.m3d']
    KOMPAS_DRAWING = ['.cdw']
    SOLIDWORKS_ASSEMBLY = ['.sldasm']
    SOLIDWORKS_PART = ['.sldprt']
    SOLIDWORKS_DRAWING = ['.slddrw']
    DXF_FILES = ['.dxf']
    ARCHIVE_FILES = ['.rar', '.zip', '.7z']
    
    def __init__(self, projects_folder: str):
        """
        Инициализация анализатора
        
        Args:
            projects_folder: Путь к папке с проектами
        """
        self.projects_folder = Path(projects_folder)
        self.projects = []  # Список найденных проектов
        self.statistics = {
            'total_folders': 0,
            'total_files': 0,
            'kompas_assemblies': 0,
            'kompas_parts': 0,
            'kompas_drawings': 0,
            'solidworks_assemblies': 0,
            'solidworks_parts': 0,
            'solidworks_drawings': 0,
            'dxf_files': 0,
            'archives': 0,
            'projects_with_params': 0,
            'h_values': defaultdict(int),
            'b1_values': defaultdict(int),
            'l1_values': defaultdict(int),
            'configurations': defaultdict(int),
            'project_types': defaultdict(int)
        }
    
    def count_files_by_extension(self, folder: Path, extensions: List[str]) -> int:
        """Подсчет файлов с указанными расширениями"""
        count = 0
        for ext in extensions:
            count += len(list(folder.rglob(f"*{ext}")))
        return count
    
    def get_file_type(self, file_path: Path) -> str:
        """Определение типа файла по расширению"""
        ext = file_path.suffix.lower()
        
        if ext in self.KOMPAS_ASSEMBLY:
            return 'kompas_assembly'
        elif ext in self.KOMPAS_PART:
            return 'kompas_part'
        elif ext in self.KOMPAS_DRAWING:
            return 'kompas_drawing'
        elif ext in self.SOLIDWORKS_ASSEMBLY:
            return 'solidworks_assembly'
        elif ext in self.SOLIDWORKS_PART:
            return 'solidworks_part'
        elif ext in self.SOLIDWORKS_DRAWING:
            return 'solidworks_drawing'
        elif ext in self.DXF_FILES:
            return 'dxf'
        elif ext in self.ARCHIVE_FILES:
            return 'archive'
        else:
            return 'other'
    
    def extract_params_from_name(self, name_str: str) -> Optional[Dict]:
        """
        Извлечение параметров из имени файла/папки
        
        Ищет паттерны:
        - ZVD.LITE.H.B1.L1
        - ZVD.TURBO.H.B1.L1  
        - H.B1.L1
        - Любые три числа через точку/дефис
        """
        result = {}
        
        # Паттерны для полного формата
        full_patterns = [
            (r'ZVD\.LITE\.(\d+)\.(\d+)\.(\d+)', 'LITE'),
            (r'ZVD\.TURBO\.(\d+)\.(\d+)\.(\d+)', 'TURBO'),
            (r'LITE\.(\d+)\.(\d+)\.(\d+)', 'LITE'),
            (r'TURBO\.(\d+)\.(\d+)\.(\d+)', 'TURBO'),
        ]
        
        for pattern, product_type in full_patterns:
            match = re.search(pattern, name_str, re.IGNORECASE)
            if match:
                try:
                    result = {
                        'H': int(match.group(1)),
                        'B1': int(match.group(2)),
                        'L1': int(match.group(3)),
                        'product_type': product_type,
                        'confidence': 'high'
                    }
                    return result
                except (ValueError, IndexError):
                    continue
        
        # Паттерн для трех чисел (менее надежный)
        # Ищем числа 50-500 для H/B1, 200-8000 для L1
        numbers_pattern = r'(\d{2,3})\.(\d{2,3})\.(\d{3,4})'
        match = re.search(numbers_pattern, name_str)
        if match:
            try:
                n1, n2, n3 = int(match.group(1)), int(match.group(2)), int(match.group(3))
                # Проверяем разумность значений
                if 50 <= n1 <= 500 and 50 <= n2 <= 500 and 200 <= n3 <= 8000:
                    return {
                        'H': n1,
                        'B1': n2,
                        'L1': n3,
                        'product_type': 'unknown',
                        'confidence': 'medium'
                    }
            except (ValueError, IndexError):
                pass
        
        return None
    
    def analyze_project_folder(self, folder: Path) -> Optional[Dict]:
        """
        Анализ папки проекта
        
        Args:
            folder: Путь к папке проекта
            
        Returns:
            Словарь с информацией о проекте или None
        """
        try:
            folder_name = folder.name
            relative_path = str(folder.relative_to(self.projects_folder))
            
            # Извлекаем параметры из имени папки
            params = self.extract_params_from_name(folder_name)
            
            # Подсчитываем файлы разных типов
            file_counts = {
                'kompas_assemblies': 0,
                'kompas_parts': 0,
                'kompas_drawings': 0,
                'solidworks_assemblies': 0,
                'solidworks_parts': 0,
                'solidworks_drawings': 0,
                'dxf_files': 0,
                'archives': 0,
                'total_files': 0
            }
            
            # Ищем файлы в папке (не рекурсивно)
            for file_path in folder.iterdir():
                if file_path.is_file():
                    file_counts['total_files'] += 1
                    file_type = self.get_file_type(file_path)
                    
                    if file_type == 'kompas_assembly':
                        file_counts['kompas_assemblies'] += 1
                        # Пробуем извлечь параметры из имени файла, если еще не нашли
                        if not params:
                            params = self.extract_params_from_name(file_path.stem)
                    elif file_type == 'kompas_part':
                        file_counts['kompas_parts'] += 1
                    elif file_type == 'kompas_drawing':
                        file_counts['kompas_drawings'] += 1
                    elif file_type == 'solidworks_assembly':
                        file_counts['solidworks_assemblies'] += 1
                        if not params:
                            params = self.extract_params_from_name(file_path.stem)
                    elif file_type == 'solidworks_part':
                        file_counts['solidworks_parts'] += 1
                    elif file_type == 'solidworks_drawing':
                        file_counts['solidworks_drawings'] += 1
                    elif file_type == 'dxf':
                        file_counts['dxf_files'] += 1
                    elif file_type == 'archive':
                        file_counts['archives'] += 1
            
            # Определяем тип проекта
            project_type = 'unknown'
            if file_counts['kompas_assemblies'] > 0:
                project_type = 'KOMPAS'
            elif file_counts['solidworks_assemblies'] > 0:
                project_type = 'SolidWorks'
            elif file_counts['archives'] > 0:
                project_type = 'Archive'
            elif file_counts['dxf_files'] > 0:
                project_type = 'DXF_Only'
            
            result = {
                'folder_name': folder_name,
                'relative_path': relative_path,
                'project_type': project_type,
                **file_counts
            }
            
            # Добавляем параметры, если нашли
            if params:
                result.update(params)
                result['has_params'] = True
            else:
                result['has_params'] = False
            
            return result
            
        except Exception as e:
            logger.error(f"  ✗ Ошибка при анализе папки {folder.name}: {e}")
            return None
    
    def analyze_all_projects(self):
        """Анализ всех проектов в папке"""
        logger.info("=" * 80)
        logger.info("УНИВЕРСАЛЬНЫЙ АНАЛИЗ ПРОЕКТОВ (КОМПАС-3D + SolidWorks)")
        logger.info("=" * 80)
        logger.info(f"📁 Папка: {self.projects_folder}\n")
        
        # Получаем список всех подпапок первого уровня
        folders = [f for f in self.projects_folder.iterdir() if f.is_dir()]
        self.statistics['total_folders'] = len(folders)
        
        if not folders:
            logger.warning("⚠ Не найдено ни одной папки с проектами")
            return
        
        logger.info(f"📊 Найдено папок для анализа: {len(folders)}\n")
        
        # Анализируем каждую папку
        for i, folder in enumerate(folders, 1):
            logger.info(f"[{i}/{len(folders)}] {folder.name}")
            
            project_info = self.analyze_project_folder(folder)
            
            if project_info:
                self.projects.append(project_info)
                
                # Обновляем статистику
                self.statistics['kompas_assemblies'] += project_info['kompas_assemblies']
                self.statistics['kompas_parts'] += project_info['kompas_parts']
                self.statistics['kompas_drawings'] += project_info['kompas_drawings']
                self.statistics['solidworks_assemblies'] += project_info['solidworks_assemblies']
                self.statistics['solidworks_parts'] += project_info['solidworks_parts']
                self.statistics['solidworks_drawings'] += project_info['solidworks_drawings']
                self.statistics['dxf_files'] += project_info['dxf_files']
                self.statistics['archives'] += project_info['archives']
                self.statistics['total_files'] += project_info['total_files']
                
                # Считаем типы проектов
                self.statistics['project_types'][project_info['project_type']] += 1
                
                # Если нашли параметры H, B1, L1
                if project_info['has_params']:
                    self.statistics['projects_with_params'] += 1
                    h = project_info['H']
                    b1 = project_info['B1']
                    l1 = project_info['L1']
                    
                    self.statistics['h_values'][h] += 1
                    self.statistics['b1_values'][b1] += 1
                    self.statistics['l1_values'][l1] += 1
                    self.statistics['configurations'][f"{h}x{b1}x{l1}"] += 1
                    
                    logger.info(f"  ✓ {project_info['project_type']} | H={h}, B1={b1}, L1={l1} | Файлов: {project_info['total_files']}")
                else:
                    logger.info(f"  • {project_info['project_type']} | Параметры не найдены | Файлов: {project_info['total_files']}")
        
        logger.info("\n" + "=" * 80)
        logger.info("✅ АНАЛИЗ ЗАВЕРШЕН")
        logger.info("=" * 80)
    
    def generate_report(self, output_file: str = "project_analysis_report.md"):
        """Генерация отчета в формате Markdown"""
        logger.info(f"\nГенерация отчета: {output_file}")
        
        report = []
        report.append("# 📊 ОТЧЕТ ПО АНАЛИЗУ ПРОЕКТОВ (КОМПАС-3D + SolidWorks)\n")
        report.append(f"**Дата анализа:** {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
        report.append(f"**Папка проектов:** `{self.projects_folder}`\n")
        report.append("\n---\n")
        
        # Общая статистика
        report.append("## 📈 Общая статистика\n")
        report.append(f"- **Всего папок проанализировано:** {self.statistics['total_folders']}")
        report.append(f"- **Всего файлов:** {self.statistics['total_files']}")
        report.append(f"- **Проектов с параметрами (H, B1, L1):** {self.statistics['projects_with_params']} ✅")
        report.append(f"- **Проектов без параметров:** {self.statistics['total_folders'] - self.statistics['projects_with_params']} ⚠️")
        report.append(f"- **Уникальных конфигураций:** {len(self.statistics['configurations'])}\n")
        
        # Статистика по типам проектов
        report.append("## 🔧 Статистика по типам проектов\n")
        for project_type, count in sorted(self.statistics['project_types'].items(), key=lambda x: x[1], reverse=True):
            report.append(f"- **{project_type}:** {count} проектов")
        report.append("")
        
        # Статистика по типам файлов
        report.append("## 📄 Статистика по типам файлов\n")
        report.append("### КОМПАС-3D:\n")
        report.append(f"- **Сборки (.a3d):** {self.statistics['kompas_assemblies']}")
        report.append(f"- **Детали (.m3d):** {self.statistics['kompas_parts']}")
        report.append(f"- **Чертежи (.cdw):** {self.statistics['kompas_drawings']}")
        report.append("\n### SolidWorks:\n")
        report.append(f"- **Сборки (.sldasm):** {self.statistics['solidworks_assemblies']}")
        report.append(f"- **Детали (.sldprt):** {self.statistics['solidworks_parts']}")
        report.append(f"- **Чертежи (.slddrw):** {self.statistics['solidworks_drawings']}")
        report.append("\n### Другие:\n")
        report.append(f"- **DXF файлы:** {self.statistics['dxf_files']}")
        report.append(f"- **Архивы (RAR/ZIP/7Z):** {self.statistics['archives']}\n")
        
        # Статистика по параметрам
        report.append("## 🔢 Статистика по параметрам\n")
        
        # H (высота)
        report.append("### Параметр H (высота конвектора)\n")
        h_sorted = sorted(self.statistics['h_values'].items())
        if h_sorted:
            h_min, h_max = h_sorted[0][0], h_sorted[-1][0]
            report.append(f"- **Диапазон:** {h_min} - {h_max} мм")
            report.append(f"- **Уникальных значений:** {len(h_sorted)}")
            report.append("\n| H (мм) | Количество проектов |")
            report.append("|--------|---------------------|")
            for h, count in h_sorted:
                report.append(f"| {h} | {count} |")
        report.append("")
        
        # B1 (ширина)
        report.append("### Параметр B1 (ширина теплообменника)\n")
        b1_sorted = sorted(self.statistics['b1_values'].items())
        if b1_sorted:
            b1_min, b1_max = b1_sorted[0][0], b1_sorted[-1][0]
            report.append(f"- **Диапазон:** {b1_min} - {b1_max} мм")
            report.append(f"- **Уникальных значений:** {len(b1_sorted)}")
            report.append("\n| B1 (мм) | Количество проектов |")
            report.append("|---------|---------------------|")
            for b1, count in b1_sorted:
                report.append(f"| {b1} | {count} |")
        report.append("")
        
        # L1 (длина)
        report.append("### Параметр L1 (длина конвектора)\n")
        l1_sorted = sorted(self.statistics['l1_values'].items())
        if l1_sorted:
            l1_min, l1_max = l1_sorted[0][0], l1_sorted[-1][0]
            report.append(f"- **Диапазон:** {l1_min} - {l1_max} мм")
            report.append(f"- **Уникальных значений:** {len(l1_sorted)}")
            report.append("\n| L1 (мм) | Количество проектов |")
            report.append("|---------|---------------------|")
            for l1, count in l1_sorted:
                report.append(f"| {l1} | {count} |")
        report.append("")
        
        # Топ конфигураций
        report.append("## 🏆 Топ-20 самых используемых конфигураций\n")
        report.append("| № | H | B1 | L1 | Количество проектов |")
        report.append("|---|---|----|----|---------------------|")
        
        top_configs = sorted(
            self.statistics['configurations'].items(),
            key=lambda x: x[1],
            reverse=True
        )[:20]
        
        for i, (config, count) in enumerate(top_configs, 1):
            h, b1, l1 = config.split('x')
            report.append(f"| {i} | {h} | {b1} | {l1} | {count} |")
        report.append("")
        
        # Рекомендации по семействам
        report.append("## 💡 Рекомендации по созданию конфигурационных семейств\n")
        
        if h_sorted:
            # Анализ диапазонов H
            h_values = [h for h, _ in h_sorted]
            h_min, h_max = min(h_values), max(h_values)
            h_mid = (h_min + h_max) / 2
            
            report.append("### Предложение семейств по высоте H:\n")
            
            # LITE-SMALL
            small_h = [h for h in h_values if h < h_mid]
            if small_h:
                report.append(f"#### 1. **LITE-SMALL**")
                report.append(f"- **Диапазон H:** {min(small_h)} - {max(small_h)} мм")
                report.append(f"- **Проектов в этом диапазоне:** {len([p for p in self.projects if p.get('has_params') and p['H'] in small_h])}")
                report.append("")
            
            # LITE-STANDARD/LARGE
            large_h = [h for h in h_values if h >= h_mid]
            if large_h:
                report.append(f"#### 2. **LITE-STANDARD/LARGE**")
                report.append(f"- **Диапазон H:** {min(large_h)} - {max(large_h)} мм")
                report.append(f"- **Проектов в этом диапазоне:** {len([p for p in self.projects if p.get('has_params') and p['H'] in large_h])}")
                report.append("")
        
        report.append("### Стратегия:\n")
        report.append("1. **Создать 2-3 базовых шаблона** (семейства) по диапазонам H")
        report.append("2. **Каждое семейство** имеет свой набор допустимых диапазонов B1 и L1")
        report.append("3. **База проверенных комбинаций** - использовать топ-20 как основу")
        report.append("4. **Валидация в GUI** - проверять, что новые параметры в пределах семейства")
        report.append("5. **Предупреждения** - если комбинация новая (не из базы)")
        report.append("")
        
        # Все найденные проекты
        report.append("## 📁 Детальный список проанализированных проектов\n")
        
        # Проекты с параметрами
        projects_with_params = [p for p in self.projects if p['has_params']]
        if projects_with_params:
            report.append("### ✅ Проекты с параметрами (H, B1, L1):\n")
            report.append("| № | H | B1 | L1 | Тип | Папка | Файлов | Конфиденс |")
            report.append("|---|---|----|----|-----|-------|--------|-----------|")
            
            for i, project in enumerate(sorted(projects_with_params, key=lambda x: (x['H'], x['B1'], x['L1'])), 1):
                folder_short = project['folder_name'][:40] + "..." if len(project['folder_name']) > 40 else project['folder_name']
                confidence = project.get('confidence', 'N/A')
                report.append(f"| {i} | {project['H']} | {project['B1']} | {project['L1']} | {project['project_type']} | `{folder_short}` | {project['total_files']} | {confidence} |")
            report.append("")
        
        # Проекты без параметров
        projects_without_params = [p for p in self.projects if not p['has_params']]
        if projects_without_params:
            report.append("### ⚠️ Проекты без параметров:\n")
            report.append("| № | Папка | Тип проекта | Файлов | КОМПАС | SolidWorks | DXF | Архивов |")
            report.append("|---|-------|-------------|--------|--------|------------|-----|---------|")
            
            for i, project in enumerate(projects_without_params, 1):
                folder_short = project['folder_name'][:50] + "..." if len(project['folder_name']) > 50 else project['folder_name']
                k_count = project['kompas_assemblies'] + project['kompas_parts'] + project['kompas_drawings']
                sw_count = project['solidworks_assemblies'] + project['solidworks_parts'] + project['solidworks_drawings']
                report.append(f"| {i} | `{folder_short}` | {project['project_type']} | {project['total_files']} | {k_count} | {sw_count} | {project['dxf_files']} | {project['archives']} |")
            report.append("")
        
        report.append("\n---")
        report.append("\n**Легенда:**")
        report.append("- **Конфиденс:** high = точное совпадение шаблону, medium = приблизительное совпадение")
        report.append("- **Тип:** KOMPAS = КОМПАС-3D, SolidWorks = SolidWorks, Archive = только архивы, DXF_Only = только DXF")
        
        # Сохраняем отчет
        report_path = Path(output_file)
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(report))
        
        logger.info(f"✓ Отчет сохранен: {report_path.absolute()}")
        
        # Также сохраняем JSON для программной обработки
        json_file = report_path.with_suffix('.json')
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump({
                'statistics': {
                    k: dict(v) if isinstance(v, defaultdict) else v 
                    for k, v in self.statistics.items()
                },
                'projects': self.projects,
                'analysis_date': datetime.now().isoformat()
            }, f, ensure_ascii=False, indent=2)
        
        logger.info(f"✓ JSON данные сохранены: {json_file.absolute()}")
        
        return report_path


def main():
    """Основная функция"""
    print("=" * 80)
    print("УНИВЕРСАЛЬНЫЙ АНАЛИЗАТОР ПРОЕКТОВ (КОМПАС-3D + SolidWorks)")
    print("=" * 80)
    print()
    
    # Путь к папке с проектами (по умолчанию)
    default_path = r"C:\Users\Vorob\Downloads\Проекты (тут лежат Чертежи к проектам)-20251027T204111Z-1-001\Проекты (тут лежат Чертежи к проектам)"
    
    if len(sys.argv) > 1:
        projects_folder = sys.argv[1]
    else:
        projects_folder = default_path
    
    if not os.path.exists(projects_folder):
        print(f"❌ ОШИБКА: Папка не найдена: {projects_folder}")
        print()
        print("Использование:")
        print(f"  python project_analyzer.py [путь_к_папке_с_проектами]")
        print()
        print("Пример:")
        print(f'  python project_analyzer.py "C:\\Projects\\KOMPAS"')
        sys.exit(1)
    
    print(f"📁 Папка для анализа: {projects_folder}")
    print()
    print("ℹ️  Анализ проводится БЕЗ открытия файлов - только структура и имена")
    print("ℹ️  Поддержка: КОМПАС-3D (.a3d, .m3d, .cdw) и SolidWorks (.sldasm, .sldprt, .slddrw)")
    print()
    
    # Создаем анализатор
    analyzer = UniversalProjectAnalyzer(projects_folder)
    
    # Запускаем анализ
    try:
        analyzer.analyze_all_projects()
        
        # Генерируем отчет
        report_file = analyzer.generate_report()
        
        print()
        print("=" * 80)
        print("✅ АНАЛИЗ ЗАВЕРШЕН УСПЕШНО!")
        print("=" * 80)
        print(f"\n📄 Отчет (Markdown): {report_file.absolute()}")
        print(f"📊 JSON данные: {report_file.with_suffix('.json').absolute()}")
        print(f"📋 Лог анализа: {Path('project_analysis.log').absolute()}")
        print()
        print("🎯 Основные находки:")
        print(f"   • Проектов с параметрами: {analyzer.statistics['projects_with_params']}")
        print(f"   • Уникальных конфигураций: {len(analyzer.statistics['configurations'])}")
        print(f"   • КОМПАС-3D проектов: {analyzer.statistics['project_types'].get('KOMPAS', 0)}")
        print(f"   • SolidWorks проектов: {analyzer.statistics['project_types'].get('SolidWorks', 0)}")
        print()
        print("📖 Откройте отчет в Markdown редакторе для подробностей!")
        
    except KeyboardInterrupt:
        print("\n\n⚠️ Анализ прерван пользователем")
        sys.exit(1)
    except Exception as e:
        logger.error(f"❌ Критическая ошибка: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()

