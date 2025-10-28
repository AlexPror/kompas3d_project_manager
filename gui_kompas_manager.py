#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🔧 КОМПАС-3D Project Manager
GUI на CustomTkinter с современным дизайном
"""

import sys
import logging
import threading
import time
from pathlib import Path
from tkinter import filedialog
from datetime import datetime

import customtkinter as ctk
from components.project_copier import ProjectCopier
from components.cascading_variables_updater import CascadingVariablesUpdater
from components.designation_updater_fixed import DesignationUpdaterFixed
from components.dxf_renamer import DxfRenamer
from components.drawing_auto_updater import DrawingAutoUpdater
from components.drawing_exporter import DrawingExporter
from components.unfolding_dxf_exporter import UnfoldingDxfExporter
from components.bmp_organizer import BmpOrganizer
from components.template_manager import TemplateManager

# Конфигурация кодировки
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# Настройка CustomTkinter
ctk.set_appearance_mode("dark")  # Темная тема
ctk.set_default_color_theme("blue")  # Синяя цветовая схема


class TextHandler(logging.Handler):
    """Обработчик логов для вывода в текстовое поле GUI"""
    
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    
    def emit(self, record):
        msg = self.format(record)
        
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert('end', msg + '\n')
            self.text_widget.see('end')
            self.text_widget.configure(state='disabled')
        
        # Безопасный вызов из другого потока
        self.text_widget.after(0, append)


class KompasManagerGUI(ctk.CTk):
    """Главное окно приложения"""
    
    def __init__(self):
        super().__init__()
        
        # Настройки окна
        self.title("🔧 КОМПАС-3D Project Manager")
        self.geometry("1200x800")
        
        # Компоненты
        self.copier = ProjectCopier()
        self.updater = CascadingVariablesUpdater()
        self.designation_updater = DesignationUpdaterFixed()
        self.dxf_exporter = UnfoldingDxfExporter()
        self.dxf_renamer = DxfRenamer()
        self.drawing_updater = DrawingAutoUpdater()
        self.drawing_exporter = DrawingExporter()
        self.bmp_organizer = BmpOrganizer()
        self.template_manager = TemplateManager()
        
        # Настройка логирования
        self.setup_logging()
        
        # Создание интерфейса
        self.create_widgets()
        
        # Переменные состояния
        self.current_project_path = None
        self.is_processing = False
        
        self.logger.info("="*60)
        self.logger.info("КОМПАС-3D Project Manager запущен!")
        self.logger.info("="*60)
        self.logger.info("Готов к работе. Выберите исходный проект для начала.\n")
    
    def setup_logging(self):
        """Настройка системы логирования"""
        # Основной логгер
        self.logger = logging.getLogger('KompasManager')
        self.logger.setLevel(logging.INFO)
        
        # Формат логов
        formatter = logging.Formatter(
            '%(asctime)s | %(levelname)-8s | %(message)s',
            datefmt='%H:%M:%S'
        )
        
        # Файловый обработчик
        file_handler = logging.FileHandler('kompas_manager.log', encoding='utf-8')
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        
        # Console handler (будет заменен на GUI после создания виджета)
        self.console_handler = logging.StreamHandler()
        self.console_handler.setFormatter(formatter)
        self.logger.addHandler(self.console_handler)
    
    def setup_gui_logging(self, text_widget):
        """Настройка вывода логов в GUI"""
        # Удаляем консольный обработчик
        self.logger.removeHandler(self.console_handler)
        
        # Добавляем GUI обработчик
        gui_handler = TextHandler(text_widget)
        formatter = logging.Formatter(
            '%(asctime)s | %(message)s',
            datefmt='%H:%M:%S'
        )
        gui_handler.setFormatter(formatter)
        self.logger.addHandler(gui_handler)
        
        # Настраиваем логгеры компонентов
        for component_logger in [
            logging.getLogger('ProjectCopier'),
            logging.getLogger('CascadingVariablesUpdater'),
            logging.getLogger('DesignationUpdaterFixed'),
            logging.getLogger('UnfoldingDxfExporter'),
            logging.getLogger('dxf_renamer'),
            logging.getLogger('DrawingAutoUpdater'),
            logging.getLogger('DrawingExporter'),
            logging.getLogger('BaseKompasComponent')
        ]:
            component_logger.setLevel(logging.INFO)
            component_logger.addHandler(gui_handler)
    
    def create_widgets(self):
        """Создание элементов интерфейса"""
        
        # Главный контейнер с отступами
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Левая панель (формы)
        left_panel = ctk.CTkFrame(main_container)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # Правая панель (логи)
        right_panel = ctk.CTkFrame(main_container)
        right_panel.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # =========================
        # ЛЕВАЯ ПАНЕЛЬ: ФОРМЫ
        # =========================
        
        # Заголовок
        title = ctk.CTkLabel(
            left_panel,
            text="🔧 Менеджер проектов КОМПАС-3D",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.pack(pady=(10, 20))
        
        # СЕКЦИЯ 1: Копирование проекта
        self.create_copy_section(left_panel)
        
        # СЕКЦИЯ 2: Переменные
        self.create_variables_section(left_panel)
        
        # СЕКЦИЯ 3: Быстрые действия
        self.create_quick_actions_section(left_panel)
        
        # Прогресс-бар
        self.progress_bar = ctk.CTkProgressBar(left_panel, mode="indeterminate")
        self.progress_bar.pack(fill="x", padx=20, pady=10)
        self.progress_bar.pack_forget()  # Скрываем по умолчанию
        
        # =========================
        # ПРАВАЯ ПАНЕЛЬ: ЛОГИ
        # =========================
        
        # Заголовок логов
        log_title = ctk.CTkLabel(
            right_panel,
            text="📋 Лог операций",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        log_title.pack(pady=(10, 10))
        
        # Текстовое поле для логов
        self.log_text = ctk.CTkTextbox(
            right_panel,
            font=ctk.CTkFont(family="Consolas", size=11),
            wrap="word"
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_text.configure(state='disabled')
        
        # Настройка GUI логирования
        self.setup_gui_logging(self.log_text)
        
        # Кнопки управления логом
        log_buttons_frame = ctk.CTkFrame(right_panel, fg_color="transparent")
        log_buttons_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        clear_log_btn = ctk.CTkButton(
            log_buttons_frame,
            text="🗑️ Очистить лог",
            command=self.clear_log,
            width=140,
            height=32
        )
        clear_log_btn.pack(side="left", padx=(0, 5))
        
        save_log_btn = ctk.CTkButton(
            log_buttons_frame,
            text="💾 Сохранить лог",
            command=self.save_log,
            width=140,
            height=32
        )
        save_log_btn.pack(side="left")
    
    def create_copy_section(self, parent):
        """Секция копирования проекта"""
        section = ctk.CTkFrame(parent)
        section.pack(fill="x", padx=20, pady=10)
        
        # Заголовок секции
        section_title = ctk.CTkLabel(
            section,
            text="📁 Копирование проекта",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        section_title.pack(pady=(15, 10), padx=15, anchor="w")
        
        # Исходный проект
        source_frame = ctk.CTkFrame(section, fg_color="transparent")
        source_frame.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(source_frame, text="Исходный проект:", width=120, anchor="w").pack(side="left")
        self.source_entry = ctk.CTkEntry(source_frame, placeholder_text="Выберите папку проекта...")
        self.source_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        source_btn = ctk.CTkButton(
            source_frame,
            text="📂 Обзор",
            command=self.select_source,
            width=100
        )
        source_btn.pack(side="left")
        
        # Целевая папка
        target_frame = ctk.CTkFrame(section, fg_color="transparent")
        target_frame.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(target_frame, text="Целевая папка:", width=120, anchor="w").pack(side="left")
        self.target_entry = ctk.CTkEntry(target_frame, placeholder_text="Куда копировать...")
        self.target_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        target_btn = ctk.CTkButton(
            target_frame,
            text="📂 Обзор",
            command=self.select_target,
            width=100
        )
        target_btn.pack(side="left")
        
        # Имя проекта
        name_frame = ctk.CTkFrame(section, fg_color="transparent")
        name_frame.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(name_frame, text="Имя проекта:", width=120, anchor="w").pack(side="left")
        self.project_name_entry = ctk.CTkEntry(name_frame, placeholder_text="ZVD.LITE.160.350.2600")
        self.project_name_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        # Кнопка копирования
        copy_btn = ctk.CTkButton(
            section,
            text="📁 Копировать проект",
            command=self.copy_project,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        copy_btn.pack(fill="x", padx=15, pady=(10, 15))
    
    def create_variables_section(self, parent):
        """Секция обновления переменных"""
        section = ctk.CTkFrame(parent)
        section.pack(fill="x", padx=20, pady=10)
        
        # Заголовок секции
        section_title = ctk.CTkLabel(
            section,
            text="🔧 Обновление переменных",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        section_title.pack(pady=(15, 10), padx=15, anchor="w")
        
        # Переменные в одной строке
        vars_frame = ctk.CTkFrame(section, fg_color="transparent")
        vars_frame.pack(fill="x", padx=15, pady=5)
        
        # H
        h_frame = ctk.CTkFrame(vars_frame, fg_color="transparent")
        h_frame.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkLabel(h_frame, text="H (высота):", anchor="w").pack()
        self.h_entry = ctk.CTkEntry(h_frame, placeholder_text="160", justify="center")
        self.h_entry.pack(fill="x")
        
        # B1
        b1_frame = ctk.CTkFrame(vars_frame, fg_color="transparent")
        b1_frame.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkLabel(b1_frame, text="B1 (ширина):", anchor="w").pack()
        self.b1_entry = ctk.CTkEntry(b1_frame, placeholder_text="350", justify="center")
        self.b1_entry.pack(fill="x")
        
        # L1
        l1_frame = ctk.CTkFrame(vars_frame, fg_color="transparent")
        l1_frame.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkLabel(l1_frame, text="L1 (длина):", anchor="w").pack()
        self.l1_entry = ctk.CTkEntry(l1_frame, placeholder_text="2600", justify="center")
        self.l1_entry.pack(fill="x")
        
        # Номер заказа
        order_frame = ctk.CTkFrame(section, fg_color="transparent")
        order_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        ctk.CTkLabel(order_frame, text="Номер заказа (опционально):", width=200, anchor="w").pack(side="left")
        self.order_entry = ctk.CTkEntry(order_frame, placeholder_text="А-180925-1801")
        self.order_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        # Кнопка обновления
        update_btn = ctk.CTkButton(
            section,
            text="🔧 Обновить переменные",
            command=self.update_variables,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        update_btn.pack(fill="x", padx=15, pady=(10, 15))
    
    def create_quick_actions_section(self, parent):
        """Секция быстрых действий"""
        section = ctk.CTkFrame(parent)
        section.pack(fill="x", padx=20, pady=10)
        
        # Заголовок секции
        section_title = ctk.CTkLabel(
            section,
            text="🚀 Быстрые действия",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        section_title.pack(pady=(15, 10), padx=15, anchor="w")
        
        # Большая кнопка "Всё сразу"
        all_btn = ctk.CTkButton(
            section,
            text="🚀 ПОЛНЫЙ ЦИКЛ (9 шагов): Копирование → Переменные → Обозначения → \nЭкспорт DXF → Переименование DXF → Чертежи → BMP → Организация → Готово!",
            command=self.do_everything,
            height=60,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2B7A0B",
            hover_color="#1F5808"
        )
        all_btn.pack(fill="x", padx=15, pady=(0, 10))
        
        # Дополнительные кнопки
        buttons_frame = ctk.CTkFrame(section, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=15, pady=(0, 15))
        
        info_btn = ctk.CTkButton(
            buttons_frame,
            text="ℹ️ Информация о проекте",
            command=self.show_project_info,
            height=35
        )
        info_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        template_btn = ctk.CTkButton(
            buttons_frame,
            text="💾 Сохранить как шаблон",
            command=self.save_as_template,
            height=35,
            fg_color="#9B59B6",
            hover_color="#7D3C98"
        )
        template_btn.pack(side="left", fill="x", expand=True, padx=(5, 0))
    
    # =========================
    # ОБРАБОТЧИКИ СОБЫТИЙ
    # =========================
    
    def select_source(self):
        """Выбор исходного проекта"""
        folder = filedialog.askdirectory(title="Выберите папку исходного проекта")
        if folder:
            self.source_entry.delete(0, 'end')
            self.source_entry.insert(0, folder)
            self.logger.info(f"Выбран исходный проект: {folder}")
    
    def select_target(self):
        """Выбор целевой папки"""
        folder = filedialog.askdirectory(title="Выберите целевую папку")
        if folder:
            self.target_entry.delete(0, 'end')
            self.target_entry.insert(0, folder)
            self.logger.info(f"Выбрана целевая папка: {folder}")
    
    def copy_project(self):
        """Копирование проекта"""
        if self.is_processing:
            self.logger.warning("⚠️ Дождитесь завершения текущей операции!")
            return
        
        source = self.source_entry.get().strip()
        target = self.target_entry.get().strip()
        project_name = self.project_name_entry.get().strip()
        
        if not source or not target or not project_name:
            self.logger.error("❌ Заполните все поля!")
            return
        
        def task():
            self.start_processing()
            try:
                result = self.copier.copy_project(source, target, project_name)
                
                if result['success']:
                    self.current_project_path = result['copied_path']
                    self.logger.info(f"✅ Проект скопирован: {self.current_project_path}\n")
                else:
                    self.logger.error(f"❌ {result['error']}\n")
            finally:
                self.stop_processing()
        
        threading.Thread(target=task, daemon=True).start()
    
    def rename_assembly(self):
        """Переименование главной сборки"""
        if self.is_processing:
            self.logger.warning("⚠️ Дождитесь завершения текущей операции!")
            return
        
        project_path = self.current_project_path or self.source_entry.get().strip()
        project_name = self.project_name_entry.get().strip()
        
        if not project_path or not project_name:
            self.logger.error("❌ Укажите путь к проекту и имя!")
            return
        
        def task():
            self.start_processing()
            try:
                result = self.copier.rename_main_assembly(project_path, project_name)
                
                if result['success']:
                    self.logger.info(f"✅ Сборка переименована!\n")
                else:
                    self.logger.error(f"❌ {result['error']}\n")
            finally:
                self.stop_processing()
        
        threading.Thread(target=task, daemon=True).start()
    
    def update_variables(self):
        """Обновление переменных"""
        if self.is_processing:
            self.logger.warning("⚠️ Дождитесь завершения текущей операции!")
            return
        
        project_path = self.current_project_path or self.source_entry.get().strip()
        
        try:
            h = int(self.h_entry.get().strip())
            b1 = int(self.b1_entry.get().strip())
            l1 = int(self.l1_entry.get().strip())
        except ValueError:
            self.logger.error("❌ Переменные должны быть числами!")
            return
        
        if not project_path:
            self.logger.error("❌ Укажите путь к проекту!")
            return
        
        def task():
            self.start_processing()
            try:
                result = self.updater.update_project_variables(project_path, h, b1, l1)
                
                if result['success']:
                    self.logger.info(f"\n✅ ПЕРЕМЕННЫЕ УСПЕШНО ОБНОВЛЕНЫ!")
                    self.logger.info(f"   Обновлено деталей: {result['parts_updated']}")
                    self.logger.info(f"   Переменных в деталях: {result['total_vars_in_parts']}\n")
                else:
                    self.logger.error(f"❌ Ошибки: {result['errors']}\n")
            finally:
                self.stop_processing()
        
        threading.Thread(target=task, daemon=True).start()
    
    def do_everything(self):
        """Выполнить все операции сразу"""
        if self.is_processing:
            self.logger.warning("⚠️ Дождитесь завершения текущей операции!")
            return
        
        source = self.source_entry.get().strip()
        target = self.target_entry.get().strip()
        project_name = self.project_name_entry.get().strip()
        
        try:
            h = int(self.h_entry.get().strip())
            b1 = int(self.b1_entry.get().strip())
            l1 = int(self.l1_entry.get().strip())
        except ValueError:
            self.logger.error("❌ Переменные должны быть числами!")
            return
        
        if not source or not target or not project_name:
            self.logger.error("❌ Заполните все поля!")
            return
        
        def task():
            self.start_processing()
            try:
                order_number = self.order_entry.get().strip() or None
                
                self.logger.info("\n" + "="*60)
                self.logger.info("🚀 ЗАПУСК ПОЛНОГО ЦИКЛА")
                self.logger.info("="*60)
                self.logger.info(f"Параметры: H={h}, B1={b1}, L1={l1}")
                if order_number:
                    self.logger.info(f"Номер заказа: {order_number}")
                self.logger.info("")
                
                # Шаг 1: Копирование
                self.logger.info("ШАГ 1/9: Копирование проекта...")
                result1 = self.copier.copy_project(source, target, project_name)
                
                if not result1['success']:
                    self.logger.error(f"❌ Ошибка копирования: {result1['error']}")
                    return
                
                self.current_project_path = result1['copied_path']
                self.logger.info(f"✅ Проект скопирован\n")
                time.sleep(1)
                
                # Шаг 2: Обновление переменных
                self.logger.info("ШАГ 2/9: Обновление переменных...")
                result2 = self.updater.update_project_variables(self.current_project_path, h, b1, l1)
                
                if not result2['success']:
                    self.logger.error(f"❌ Ошибки обновления переменных: {result2['errors']}")
                    return
                
                self.logger.info(f"✅ Переменные обновлены (деталей: {result2['parts_updated']})\n")
                time.sleep(1)
                
                # Шаг 3: Обновление обозначений (marking) + переименование файлов
                self.logger.info("ШАГ 3/9: Обновление обозначений и переименование...")
                result3 = self.designation_updater.update_all_designations(
                    self.current_project_path, h, b1, l1, order_number
                )
                
                if not result3['success']:
                    self.logger.error(f"❌ Ошибки обновления обозначений: {result3.get('error', 'Unknown')}")
                    self.logger.error(f"   Детали ошибок: {result3.get('errors', [])}")
                    # Продолжаем выполнение даже при ошибках
                
                self.logger.info(f"✅ Обозначения обновлены (деталей: {result3.get('parts_renamed', 0)})\n")
                time.sleep(1)
                
                # Шаг 4: Экспорт разверток в DXF
                self.logger.info("ШАГ 4/9: Экспорт разверток в DXF...")
                
                from pathlib import Path
                dxf_folder = Path(self.current_project_path) / "DXF"
                dxf_folder.mkdir(exist_ok=True)
                
                result4_export = self.dxf_exporter.export_all_unfoldings(
                    self.current_project_path,
                    str(dxf_folder)
                )
                
                if result4_export['success']:
                    self.logger.info(f"✅ DXF экспортированы (файлов: {result4_export['exported']}/{result4_export['total']})\n")
                else:
                    self.logger.warning(f"⚠️ DXF экспорт: {result4_export.get('errors', ['Нет разверток'])}\n")
                
                time.sleep(1)
                
                # Шаг 5: Переименование DXF
                self.logger.info("ШАГ 5/9: Переименование DXF файлов...")
                result5 = self.dxf_renamer.rename_dxf_files(self.current_project_path, order_number)
                
                if result5['success']:
                    self.logger.info(f"✅ DXF переименованы (файлов: {result5['renamed_count']})\n")
                else:
                    self.logger.warning(f"⚠️ DXF: {result5.get('errors', ['Нет DXF папки'])}\n")
                
                time.sleep(1)
                
                # Шаг 6: Обновление чертежей
                self.logger.info("ШАГ 6/9: Обновление чертежей...")
                result6 = self.drawing_updater.update_all_drawings(self.current_project_path)
                
                if result6['success']:
                    self.logger.info(f"✅ Чертежи обновлены (обновлено: {result6['drawings_updated']}, ошибок: {result6['drawings_failed']})\n")
                else:
                    self.logger.warning(f"⚠️ Чертежи: {result6.get('errors', ['Unknown'])}\n")
                
                time.sleep(1)
                
                # Шаг 7: Экспорт чертежей в BMP
                self.logger.info("ШАГ 7/9: Экспорт чертежей в BMP...")
                
                # Находим чертежи (исключая развертки)
                drawing_files = self.drawing_exporter.find_drawing_files(
                    self.current_project_path, 
                    exclude_unfoldings=True
                )
                
                if drawing_files:
                    self.logger.info(f"Найдено чертежей для экспорта: {len(drawing_files)}")
                    
                    # Экспортируем каждый чертеж
                    exported_count = 0
                    for drawing_file in drawing_files:
                        from pathlib import Path
                        drawing_path = Path(drawing_file)
                        output_path = str(drawing_path.with_suffix('.bmp'))
                        
                        export_result = self.drawing_exporter.export_drawing_to_image(
                            str(drawing_file),
                            output_path,
                            format_type='BMP',
                            resolution=300
                        )
                        
                        if export_result['success']:
                            exported_count += 1
                            self.logger.info(f"  ✓ {drawing_path.name} → BMP")
                        else:
                            self.logger.warning(f"  ✗ {drawing_path.name}: {export_result.get('error', 'Unknown')}")
                    
                    self.logger.info(f"✅ Экспортировано BMP: {exported_count}/{len(drawing_files)}\n")
                else:
                    self.logger.info("  Чертежей для экспорта не найдено\n")
                    exported_count = 0
                
                time.sleep(1)
                
                # Шаг 8: Организация BMP файлов
                self.logger.info("ШАГ 8/9: Организация BMP файлов...")
                result8 = self.bmp_organizer.organize_bmp_files(self.current_project_path)
                
                if result8['success'] and result8['moved_count'] > 0:
                    self.logger.info(f"✅ BMP файлы организованы: {result8['moved_count']} → папка BMP/\n")
                else:
                    self.logger.info(f"  Нет BMP файлов для организации\n")
                
                time.sleep(1)
                
                # Шаг 9: Итоговая информация
                self.logger.info("ШАГ 9/9: Формирование итогового отчета...")
                
                # Итоговое сообщение
                self.logger.info("\n" + "="*60)
                self.logger.info("🎉 ВСЕ ОПЕРАЦИИ ВЫПОЛНЕНЫ УСПЕШНО!")
                self.logger.info("="*60)
                self.logger.info(f"📁 Проект: {self.current_project_path}")
                self.logger.info(f"📊 Параметры: H={h}, B1={b1}, L1={l1}")
                if order_number:
                    self.logger.info(f"🏷️ Номер заказа: {order_number}")
                self.logger.info("")
                self.logger.info("📈 Статистика:")
                self.logger.info(f"  ✅ Переменных обновлено: {result2.get('total_vars_in_parts', 0)}")
                self.logger.info(f"  ✅ Деталей переименовано: {result3.get('parts_renamed', 0)}")
                self.logger.info(f"  ✅ DXF экспортировано: {result4_export.get('exported', 0)}")
                self.logger.info(f"  ✅ DXF переименовано: {result5.get('renamed_count', 0)}")
                self.logger.info(f"  ✅ Чертежей обновлено: {result6.get('drawings_updated', 0)}")
                self.logger.info(f"  ✅ BMP экспортировано: {exported_count if 'exported_count' in locals() else 0}")
                self.logger.info("="*60 + "\n")
                
            except Exception as e:
                self.logger.error(f"❌ Критическая ошибка: {e}")
                import traceback
                self.logger.error(traceback.format_exc())
            finally:
                self.stop_processing()
        
        threading.Thread(target=task, daemon=True).start()
    
    def show_project_info(self):
        """Показать информацию о проекте"""
        project_path = self.current_project_path or self.source_entry.get().strip()
        
        if not project_path:
            self.logger.error("❌ Укажите путь к проекту!")
            return
        
        info = self.copier.get_project_info(project_path)
        
        self.logger.info("\n" + "="*60)
        self.logger.info("ℹ️ ИНФОРМАЦИЯ О ПРОЕКТЕ")
        self.logger.info("="*60)
        self.logger.info(f"📁 Путь: {project_path}")
        self.logger.info(f"🔧 Сборок (.a3d): {len(info['assembly_files'])}")
        self.logger.info(f"📐 Чертежей (.cdw): {len(info['drawing_files'])}")
        self.logger.info(f"⚙️ Деталей (.m3d): {len(info['part_files'])}")
        self.logger.info(f"📄 Других файлов: {len(info['other_files'])}")
        self.logger.info(f"📊 Всего файлов: {info['total_files']}")
        self.logger.info("="*60 + "\n")
    
    def save_as_template(self):
        """Сохранить текущий проект как шаблон"""
        project_path = self.current_project_path or self.source_entry.get().strip()
        
        if not project_path:
            self.logger.error("❌ Укажите путь к проекту!")
            return
        
        # Диалог для ввода информации о шаблоне
        dialog = ctk.CTkInputDialog(
            text="Введите название шаблона:",
            title="Сохранить как шаблон"
        )
        template_name = dialog.get_input()
        
        if not template_name:
            self.logger.info("Отмена создания шаблона")
            return
        
        # Получаем параметры, если введены
        parameters = {}
        try:
            if self.h_entry.get().strip():
                parameters['H'] = int(self.h_entry.get().strip())
            if self.b1_entry.get().strip():
                parameters['B1'] = int(self.b1_entry.get().strip())
            if self.l1_entry.get().strip():
                parameters['L1'] = int(self.l1_entry.get().strip())
        except:
            pass
        
        # Описание
        description = f"Шаблон создан из проекта {Path(project_path).name}"
        
        # Создаем шаблон
        self.logger.info(f"\n💾 Создание шаблона '{template_name}'...")
        
        result = self.template_manager.add_template_from_project(
            project_path=project_path,
            template_name=template_name,
            description=description,
            parameters=parameters,
            tags=['ZVD', 'LITE']
        )
        
        if result['success']:
            self.logger.info(f"✅ Шаблон '{template_name}' создан успешно!")
            self.logger.info(f"   ID: {result['template_id']}")
            self.logger.info(f"   Всего шаблонов: {len(self.template_manager.list_templates())}\n")
        else:
            self.logger.error(f"❌ Ошибка создания шаблона: {result['error']}\n")
    
    def clear_log(self):
        """Очистить лог"""
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')
        self.logger.info("Лог очищен\n")
    
    def save_log(self):
        """Сохранить лог в файл"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"kompas_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get('1.0', 'end'))
                self.logger.info(f"✅ Лог сохранён: {filename}\n")
            except Exception as e:
                self.logger.error(f"❌ Ошибка сохранения: {e}\n")
    
    def start_processing(self):
        """Начало обработки (показать прогресс-бар)"""
        self.is_processing = True
        self.progress_bar.pack(fill="x", padx=20, pady=10)
        self.progress_bar.start()
    
    def stop_processing(self):
        """Конец обработки (скрыть прогресс-бар)"""
        self.is_processing = False
        self.progress_bar.stop()
        self.progress_bar.pack_forget()


def main():
    """Запуск приложения"""
    app = KompasManagerGUI()
    app.mainloop()


if __name__ == "__main__":
    main()

