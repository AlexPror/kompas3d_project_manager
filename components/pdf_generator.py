import os
import re
from pathlib import Path
from typing import List, Dict, Optional
from PIL import Image
from .base_component import BaseKompasComponent

class PdfGenerator(BaseKompasComponent):
    """Компонент для объединения BMP файлов в один PDF"""
    
    def __init__(self):
        super().__init__()

    def natural_sort_key(self, s: str):
        """Ключ для естественной сортировки (1, 2, 10 вместо 1, 10, 2)"""
        return [int(text) if text.isdigit() else text.lower()
                for text in re.split('([0-9]+)', s)]

    def generate_pdf(self, project_path: str, order_number: str, material: str) -> Dict:
        """
        Объединить BMP файлы в PDF
        
        Args:
            project_path: Путь к проекту
            order_number: Номер заказа
            material: Материал (для определения суффикса имени файла)
            
        Returns:
            Dict с результатом
        """
        result = {
            'success': False,
            'pdf_path': None,
            'error': None
        }
        
        try:
            self.logger.info("="*60)
            self.logger.info("ГЕНЕРАЦИЯ PDF ИЗ BMP")
            self.logger.info("="*60)
            
            project_dir = Path(project_path)
            bmp_dir = project_dir / "BMP"
            
            if not bmp_dir.exists():
                error_msg = f"Папка BMP не найдена: {bmp_dir}"
                self.logger.error(error_msg)
                result['error'] = error_msg
                return result
            
            # 1. Поиск BMP файлов
            bmp_files = list(bmp_dir.glob("*.bmp"))
            if not bmp_files:
                error_msg = "В папке BMP нет файлов .bmp"
                self.logger.error(error_msg)
                result['error'] = error_msg
                return result
            
            # 2. Сортировка
            bmp_files.sort(key=lambda x: self.natural_sort_key(x.name))
            self.logger.info(f"Найдено файлов: {len(bmp_files)}")
            
            # Определение имени файла
            # Имя проекта (берем из имени папки проекта)
            project_name = project_dir.name
            
            # Суффикс материала
            material_lower = material.lower() if material else ""
            
            # Проверяем на нержавейку (AISI 304)
            if "aisi 304" in material_lower or "08х18н10" in material_lower:
                material_suffix = "(НЕРЖАВЕЙКА)"
            # Проверяем на оцинковку
            elif "оцинкован" in material_lower or "08ps" in material_lower or "08пс" in material_lower:
                material_suffix = "(ОЦИНКОВКА)"
            # По умолчанию - оцинковка
            else:
                material_suffix = "(ОЦИНКОВКА)"
            
            # Формируем имя: ZVD... [Номер заказа] конструкторская документация (МАТЕРИАЛ).pdf
            # Если project_name уже содержит ZVD..., используем его.
            # Пример: ZVD.LITE.110.400.1900 А-151025-1235 конструкторская документация (ОЦИНКОВКА).pdf
            
            pdf_filename = f"{project_name} {order_number} конструкторская документация {material_suffix}.pdf"
            pdf_path = project_dir / pdf_filename
            
            self.logger.info(f"Имя файла: {pdf_filename}")
            
            # 4. Генерация PDF из BMP
            images = []
            for bmp_file in bmp_files:
                try:
                    img = Image.open(bmp_file)
                    # Конвертируем в RGB
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    images.append(img)
                except Exception as e:
                    self.logger.warning(f"  Не удалось загрузить {bmp_file.name}: {e}")
            
            if not images:
                error_msg = "Не удалось загрузить ни одного изображения"
                self.logger.error(error_msg)
                result['error'] = error_msg
                return result
            
            self.logger.info("Сохранение PDF...")
            # Сохраняем: первое изображение + остальные как append_images
            images[0].save(
                pdf_path, 
                "PDF", 
                resolution=100.0, 
                save_all=True, 
                append_images=images[1:]
            )
            
            self.logger.info(f"✅ PDF успешно создан: {pdf_path}")
            result['success'] = True
            result['pdf_path'] = str(pdf_path)
            
            # 5. Авто-открытие
            try:
                os.startfile(pdf_path)
                self.logger.info("  Файл открыт")
            except Exception as e:
                self.logger.warning(f"  Не удалось открыть файл автоматически: {e}")
                
        except Exception as e:
            self.logger.error(f"Ошибка генерации PDF: {e}")
            result['error'] = str(e)
            
        return result
