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
import importlib
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

# Сохраняем ссылки на модули для перезагрузки
from components import (
    project_copier,
    cascading_variables_updater,
    designation_updater_fixed,
    dxf_renamer,
    drawing_auto_updater,
    drawing_exporter,
    unfolding_dxf_exporter,
    bmp_organizer,
    template_manager,
    base_component
)

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
        self.cancel_requested = False
        
        self.logger.info("="*60)
        self.logger.info("КОМПАС-3D Project Manager запущен!")
        self.logger.info("="*60)
        self.logger.info("Готов к работе. Выберите исходный проект для начала.")
        self.logger.info("")
        self.logger.info("💡 ПОДСКАЗКА:")
        self.logger.info("   1. Выберите тип проекта (ZVD.LITE или ZVD.TURBO)")
        self.logger.info("   2. Введите H, B1, L1 - имя проекта заполнится автоматически!")
        self.logger.info("   Например: ZVD.LITE.160.350.2600")
        self.logger.info("")
    
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
        
        reload_btn = ctk.CTkButton(
            log_buttons_frame,
            text="🔄 Обновить",
            command=self.reload_modules,
            width=120,
            height=32,
            fg_color="#2E7D32",
            hover_color="#1B5E20"
        )
        reload_btn.pack(side="left", padx=(0, 5))
        
        clear_log_btn = ctk.CTkButton(
            log_buttons_frame,
            text="🗑️ Очистить лог",
            command=self.clear_log,
            width=120,
            height=32
        )
        clear_log_btn.pack(side="left", padx=(0, 5))
        
        save_log_btn = ctk.CTkButton(
            log_buttons_frame,
            text="💾 Сохранить лог",
            command=self.save_log,
            width=120,
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
        
        ctk.CTkLabel(source_frame, text="Исходный проект: *", width=120, anchor="w").pack(side="left")
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
        
        ctk.CTkLabel(target_frame, text="Целевая папка: *", width=120, anchor="w").pack(side="left")
        self.target_entry = ctk.CTkEntry(target_frame, placeholder_text="Куда копировать...")
        self.target_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        target_btn = ctk.CTkButton(
            target_frame,
            text="📂 Обзор",
            command=self.select_target,
            width=100
        )
        target_btn.pack(side="left")
        
        # Тип проекта и имя проекта
        project_type_frame = ctk.CTkFrame(section, fg_color="transparent")
        project_type_frame.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(project_type_frame, text="Тип проекта: *", width=120, anchor="w").pack(side="left")
        self.project_type_menu = ctk.CTkOptionMenu(
            project_type_frame,
            values=["ZVD.LITE", "ZVD.TURBO"],
            command=self.on_project_type_changed,
            width=150
        )
        self.project_type_menu.set("ZVD.LITE")
        self.project_type_menu.pack(side="left", padx=5)
        
        # Имя проекта
        name_frame = ctk.CTkFrame(section, fg_color="transparent")
        name_frame.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(name_frame, text="Имя проекта: *", width=120, anchor="w").pack(side="left")
        self.project_name_entry = ctk.CTkEntry(name_frame, placeholder_text="Автозаполнение по H, B1, L1")
        self.project_name_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        # Кнопки управления
        buttons_frame = ctk.CTkFrame(section, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        # Кнопка копирования
        self.copy_btn = ctk.CTkButton(
            buttons_frame,
            text="📁 Копировать проект",
            command=self.copy_project,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.copy_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # Кнопка прервать
        self.cancel_btn = ctk.CTkButton(
            buttons_frame,
            text="⏹️ Прервать",
            command=self.cancel_operation,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#D32F2F",
            hover_color="#B71C1C",
            state="disabled"
        )
        self.cancel_btn.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # Кнопка очистить поля
        clear_btn = ctk.CTkButton(
            section,
            text="🗑️ Очистить все поля",
            command=self.clear_all_fields,
            height=35,
            fg_color="#757575",
            hover_color="#616161"
        )
        clear_btn.pack(fill="x", padx=15, pady=(5, 15))
    
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
        ctk.CTkLabel(h_frame, text="H (высота): *", anchor="w").pack()
        self.h_entry = ctk.CTkEntry(h_frame, placeholder_text="например: 160", justify="center")
        self.h_entry.pack(fill="x")
        self.h_entry.bind("<KeyRelease>", lambda e: self.update_project_name_from_variables())
        self.h_entry.bind("<FocusOut>", lambda e: self.update_project_name_from_variables())
        
        # B1
        b1_frame = ctk.CTkFrame(vars_frame, fg_color="transparent")
        b1_frame.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkLabel(b1_frame, text="B1 (ширина): *", anchor="w").pack()
        self.b1_entry = ctk.CTkEntry(b1_frame, placeholder_text="например: 350", justify="center")
        self.b1_entry.pack(fill="x")
        self.b1_entry.bind("<KeyRelease>", lambda e: self.update_project_name_from_variables())
        self.b1_entry.bind("<FocusOut>", lambda e: self.update_project_name_from_variables())
        
        # L1
        l1_frame = ctk.CTkFrame(vars_frame, fg_color="transparent")
        l1_frame.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkLabel(l1_frame, text="L1 (длина): *", anchor="w").pack()
        self.l1_entry = ctk.CTkEntry(l1_frame, placeholder_text="например: 2600", justify="center")
        self.l1_entry.pack(fill="x")
        self.l1_entry.bind("<KeyRelease>", lambda e: self.update_project_name_from_variables())
        self.l1_entry.bind("<FocusOut>", lambda e: self.update_project_name_from_variables())
        
        # Номер заказа
        order_frame = ctk.CTkFrame(section, fg_color="transparent")
        order_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        ctk.CTkLabel(order_frame, text="Номер заказа (опционально):", width=200, anchor="w").pack(side="left")
        self.order_entry = ctk.CTkEntry(order_frame, placeholder_text="А-180925-1801")
        self.order_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        # Разработчик
        developer_frame = ctk.CTkFrame(section, fg_color="transparent")
        developer_frame.pack(fill="x", padx=15, pady=(5, 5))
        
        ctk.CTkLabel(developer_frame, text="Разработчик (опционально):", width=200, anchor="w").pack(side="left")
        self.developer_entry = ctk.CTkEntry(developer_frame, placeholder_text="Иванов И.И.")
        self.developer_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        # Кнопки обновления
        buttons_frame = ctk.CTkFrame(section, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        # Кнопка обновления переменных
        update_vars_btn = ctk.CTkButton(
            buttons_frame,
            text="🔧 Обновить переменные",
            command=self.update_variables,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        update_vars_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # Кнопка обновления обозначений
        update_designations_btn = ctk.CTkButton(
            buttons_frame,
            text="📝 Обновить обозначения (LITE/TURBO)",
            command=self.update_designations_only,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#E67E22",
            hover_color="#D35400"
        )
        update_designations_btn.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # Кнопка обновления штампов
        buttons_frame2 = ctk.CTkFrame(section, fg_color="transparent")
        buttons_frame2.pack(fill="x", padx=15, pady=(5, 5))
        
        update_stamps_btn = ctk.CTkButton(
            buttons_frame2,
            text="📋 Обновить штампы чертежей (Разработчик)",
            command=self.update_drawing_stamps,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#16A085",
            hover_color="#138D75"
        )
        update_stamps_btn.pack(fill="x")
        
        # Подсказки
        hint_frame = ctk.CTkFrame(section, fg_color="transparent")
        hint_frame.pack(fill="x", padx=15, pady=(5, 5))
        
        hint1 = ctk.CTkLabel(
            hint_frame,
            text="💡 Для обновления существующего проекта:",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#E67E22",
            anchor="w"
        )
        hint1.pack(anchor="w")
        
        hint2 = ctk.CTkLabel(
            hint_frame,
            text="   1. Укажите путь к проекту в поле 'Исходный проект' ☝",
            font=ctk.CTkFont(size=10),
            text_color="gray",
            anchor="w"
        )
        hint2.pack(anchor="w")
        
        hint3 = ctk.CTkLabel(
            hint_frame,
            text="   2. Измените тип проекта (LITE/TURBO) и параметры H, B1, L1",
            font=ctk.CTkFont(size=10),
            text_color="gray",
            anchor="w"
        )
        hint3.pack(anchor="w")
        
        hint4 = ctk.CTkLabel(
            hint_frame,
            text="   3. Нажмите '📝 Обновить обозначения' →",
            font=ctk.CTkFont(size=10),
            text_color="gray",
            anchor="w"
        )
        hint4.pack(anchor="w", pady=(0, 10))
    
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
            text="🚀 ПОЛНЫЙ ЦИКЛ (10 шагов): Копирование → Переименование → Переменные → Обозначения → \nЭкспорт DXF → Переименование DXF → Чертежи → BMP → Организация → Готово!",
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
    
    def on_project_type_changed(self, choice):
        """Обработка изменения типа проекта - автозаполнение имени"""
        self.update_project_name_from_variables()
        self.logger.info(f"Тип проекта изменен на: {choice}")
    
    def update_project_name_from_variables(self):
        """Автообновление имени проекта из переменных H, B1, L1"""
        # Получаем текущие значения
        project_type = self.project_type_menu.get()
        h = self.h_entry.get().strip()
        b1 = self.b1_entry.get().strip()
        l1 = self.l1_entry.get().strip()
        
        # Формируем имя проекта только если все поля заполнены
        if h and b1 and l1:
            # Проверяем, что все значения - числа
            try:
                int(h)
                int(b1)
                int(l1)
                # Все значения корректны - формируем имя
                project_name = f"{project_type}.{h}.{b1}.{l1}"
                # Обновляем поле имени проекта
                self.project_name_entry.delete(0, 'end')
                self.project_name_entry.insert(0, project_name)
            except ValueError:
                # Если не все значения числа - не обновляем
                pass
    
    def clear_all_fields(self):
        """Очистить все поля ввода"""
        self.source_entry.delete(0, 'end')
        self.target_entry.delete(0, 'end')
        self.project_name_entry.delete(0, 'end')
        self.h_entry.delete(0, 'end')
        self.b1_entry.delete(0, 'end')
        self.l1_entry.delete(0, 'end')
        self.order_entry.delete(0, 'end')
        self.developer_entry.delete(0, 'end')
        self.current_project_path = None
        self.project_type_menu.set("ZVD.LITE")
        self.logger.info("🗑️ Все поля очищены\n")
    
    def cancel_operation(self):
        """Прервать текущую операцию"""
        if not self.is_processing:
            self.logger.info("Нет активных операций для прерывания")
            return
        
        self.cancel_requested = True
        self.logger.warning("\n" + "="*60)
        self.logger.warning("⚠️ ЗАПРОС НА ПРЕРЫВАНИЕ ОПЕРАЦИИ")
        self.logger.warning("Операция будет остановлена на ближайшей точке проверки...")
        self.logger.warning("="*60 + "\n")
    
    def validate_required_fields(self):
        """Проверка заполнения всех обязательных полей"""
        errors = []
        
        if not self.source_entry.get().strip():
            errors.append("Исходный проект")
        
        if not self.target_entry.get().strip():
            errors.append("Целевая папка")
        
        if not self.project_name_entry.get().strip():
            errors.append("Имя проекта")
        
        if not self.h_entry.get().strip():
            errors.append("H (высота)")
        
        if not self.b1_entry.get().strip():
            errors.append("B1 (ширина)")
        
        if not self.l1_entry.get().strip():
            errors.append("L1 (длина)")
        
        if errors:
            self.logger.error(f"❌ Заполните обязательные поля (*): {', '.join(errors)}")
            return False
        
        # Проверка, что переменные - числа
        try:
            int(self.h_entry.get().strip())
            int(self.b1_entry.get().strip())
            int(self.l1_entry.get().strip())
        except ValueError:
            self.logger.error("❌ H, B1, L1 должны быть числами!")
            return False
        
        return True
    
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
        
        # Валидация ВСЕХ обязательных полей перед копированием
        if not self.validate_required_fields():
            return
        
        source = self.source_entry.get().strip()
        target = self.target_entry.get().strip()
        project_name = self.project_name_entry.get().strip()
        
        def task():
            self.start_processing()
            try:
                # Проверка 1: Перед началом
                if self.cancel_requested:
                    self.logger.warning("❌ Операция отменена пользователем\n")
                    return
                
                # Шаг 1: Копирование
                self.logger.info("\n🚀 Начало копирования проекта...")
                result = self.copier.copy_project(source, target, project_name)
                
                # Проверка 2: После копирования
                if self.cancel_requested:
                    self.logger.warning("❌ Операция отменена пользователем после копирования\n")
                    return
                
                if result['success']:
                    self.current_project_path = result['copied_path']
                    files_count = result.get('copied_files', 0)
                    self.logger.info(f"✅ Проект скопирован: {self.current_project_path}")
                    if files_count > 0:
                        self.logger.info(f"   📊 Скопировано файлов: {files_count}\n")
                    else:
                        self.logger.info("")
                    
                    # Проверка 3: Перед переименованием
                    if self.cancel_requested:
                        self.logger.warning("❌ Операция отменена пользователем\n")
                        return
                    
                    # Шаг 2: Автоматическое переименование сборки и чертежа
                    self.logger.info("Переименование главной сборки и чертежа...")
                    result_rename = self.copier.rename_main_assembly(self.current_project_path, project_name)
                    
                    # Проверка 4: После переименования
                    if self.cancel_requested:
                        self.logger.warning("❌ Операция отменена пользователем\n")
                        return
                    
                    if result_rename['success']:
                        self.logger.info(f"✅ Переименовано файлов: {result_rename.get('renamed_count', 0)}")
                        for renamed_file in result_rename.get('renamed_files', []):
                            from pathlib import Path
                            self.logger.info(f"   • {Path(renamed_file).name}")
                        self.logger.info("")
                    else:
                        self.logger.warning(f"⚠️ Ошибка переименования: {result_rename.get('error', 'Unknown')}\n")
                else:
                    self.logger.error(f"❌ {result['error']}\n")
            except Exception as e:
                self.logger.error(f"❌ Критическая ошибка: {e}\n")
                import traceback
                self.logger.error(traceback.format_exc())
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
    
    def update_drawing_stamps(self):
        """Обновление штампов чертежей (разработчик)"""
        if self.is_processing:
            self.logger.warning("⚠️ Дождитесь завершения текущей операции!")
            return
        
        project_path = self.current_project_path or self.source_entry.get().strip()
        developer_name = self.developer_entry.get().strip()
        
        if not project_path:
            self.logger.error("❌ Укажите путь к проекту в поле 'Исходный проект' (📂 Обзор)")
            return
        
        if not developer_name:
            self.logger.error("❌ Введите имя разработчика!")
            return
        
        def task():
            self.start_processing()
            try:
                self.logger.info("\n" + "="*60)
                self.logger.info("📋 ОБНОВЛЕНИЕ ШТАМПОВ ЧЕРТЕЖЕЙ")
                self.logger.info("="*60)
                self.logger.info(f"Проект: {project_path}")
                self.logger.info(f"Разработчик: {developer_name}")
                self.logger.info("")
                
                # Проверка прерывания
                if self.cancel_requested:
                    self.logger.warning("❌ Операция отменена пользователем\n")
                    return
                
                # Обновление чертежей со штампами
                result = self.drawing_updater.update_all_drawings(
                    project_path, 
                    developer=developer_name
                )
                
                if self.cancel_requested:
                    self.logger.warning("❌ Операция отменена пользователем\n")
                    return
                
                if result['success']:
                    self.logger.info("\n" + "="*60)
                    self.logger.info("✅ ШТАМПЫ УСПЕШНО ОБНОВЛЕНЫ!")
                    self.logger.info("="*60)
                    self.logger.info(f"   Чертежей обновлено: {result.get('drawings_updated', 0)}")
                    self.logger.info(f"   Ошибок: {result.get('drawings_failed', 0)}")
                    self.logger.info("="*60 + "\n")
                else:
                    self.logger.error(f"❌ Ошибки обновления штампов: {result.get('errors', 'Unknown')}")
                    self.logger.error("")
            
            except Exception as e:
                self.logger.error(f"❌ Критическая ошибка: {e}\n")
                import traceback
                self.logger.error(traceback.format_exc())
            finally:
                self.stop_processing()
        
        threading.Thread(target=task, daemon=True).start()
    
    def update_designations_only(self):
        """Обновление обозначений в существующем проекте"""
        if self.is_processing:
            self.logger.warning("⚠️ Дождитесь завершения текущей операции!")
            return
        
        # Валидация полей
        if not self.validate_required_fields():
            return
        
        # Определяем путь к проекту (приоритет: исходный проект → целевая папка + имя → сохраненный путь)
        project_path = None
        
        # Вариант 1: Используем поле "Исходный проект"
        source_path = self.source_entry.get().strip()
        if source_path:
            from pathlib import Path
            if Path(source_path).exists():
                project_path = source_path
                self.logger.info(f"📂 Выбран проект из поля 'Исходный проект': {project_path}\n")
        
        # Вариант 2: Если был скопирован проект ранее
        if not project_path and self.current_project_path:
            project_path = self.current_project_path
            self.logger.info(f"📂 Используется ранее скопированный проект: {project_path}\n")
        
        if not project_path:
            self.logger.error("❌ Укажите путь к проекту в поле 'Исходный проект' (📂 Обзор)")
            self.logger.error("   Это должна быть папка с проектом, в котором нужно обновить обозначения\n")
            return
        
        try:
            h = int(self.h_entry.get().strip())
            b1 = int(self.b1_entry.get().strip())
            l1 = int(self.l1_entry.get().strip())
        except ValueError:
            self.logger.error("❌ H, B1, L1 должны быть числами!")
            return
        
        order_number = self.order_entry.get().strip() or None
        
        def task():
            self.start_processing()
            try:
                self.logger.info("\n" + "="*60)
                self.logger.info("📝 ОБНОВЛЕНИЕ ОБОЗНАЧЕНИЙ В СУЩЕСТВУЮЩЕМ ПРОЕКТЕ")
                self.logger.info("="*60)
                self.logger.info(f"Проект: {project_path}")
                self.logger.info(f"Параметры: H={h}, B1={b1}, L1={l1}")
                if order_number:
                    self.logger.info(f"Номер заказа: {order_number}")
                self.logger.info("")
                
                # Проверка прерывания
                if self.cancel_requested:
                    self.logger.warning("❌ Операция отменена пользователем\n")
                    return
                
                # Обновление обозначений
                result = self.designation_updater.update_all_designations(
                    project_path, h, b1, l1, order_number
                )
                
                if self.cancel_requested:
                    self.logger.warning("❌ Операция отменена пользователем\n")
                    return
                
                if result['success']:
                    self.logger.info("\n" + "="*60)
                    self.logger.info("✅ ОБОЗНАЧЕНИЯ УСПЕШНО ОБНОВЛЕНЫ!")
                    self.logger.info("="*60)
                    self.logger.info(f"   Деталей переименовано: {result.get('parts_renamed', 0)}")
                    self.logger.info("="*60 + "\n")
                else:
                    self.logger.error(f"❌ Ошибки обновления обозначений: {result.get('error', 'Unknown')}")
                    if result.get('errors'):
                        self.logger.error(f"   Детали ошибок: {result.get('errors', [])}")
                    self.logger.error("")
            
            except Exception as e:
                self.logger.error(f"❌ Критическая ошибка: {e}\n")
                import traceback
                self.logger.error(traceback.format_exc())
            finally:
                self.stop_processing()
        
        threading.Thread(target=task, daemon=True).start()
    
    def do_everything(self):
        """Выполнить все операции сразу"""
        if self.is_processing:
            self.logger.warning("⚠️ Дождитесь завершения текущей операции!")
            return
        
        # Валидация всех обязательных полей
        if not self.validate_required_fields():
            return
        
        source = self.source_entry.get().strip()
        target = self.target_entry.get().strip()
        project_name = self.project_name_entry.get().strip()
        
        try:
            h = int(self.h_entry.get().strip())
            b1 = int(self.b1_entry.get().strip())
            l1 = int(self.l1_entry.get().strip())
        except ValueError:
            self.logger.error("❌ H, B1, L1 должны быть числами!")
            return
        
        def task():
            self.start_processing()
            try:
                order_number = self.order_entry.get().strip() or None
                
                developer_name = self.developer_entry.get().strip() or None
                
                self.logger.info("\n" + "="*60)
                self.logger.info("🚀 ЗАПУСК ПОЛНОГО ЦИКЛА")
                self.logger.info("="*60)
                self.logger.info(f"Тип проекта: {project_name.split('.')[0]}.{project_name.split('.')[1]}")
                self.logger.info(f"Параметры: H={h}, B1={b1}, L1={l1}")
                if order_number:
                    self.logger.info(f"Номер заказа: {order_number}")
                if developer_name:
                    self.logger.info(f"Разработчик: {developer_name}")
                self.logger.info("")
                
                # Шаг 1: Копирование
                self.logger.info("ШАГ 1/10: Копирование проекта...")
                
                if self.cancel_requested:
                    self.logger.warning("❌ Операция прервана пользователем\n")
                    return
                
                result1 = self.copier.copy_project(source, target, project_name)
                
                if not result1['success']:
                    self.logger.error(f"❌ Ошибка копирования: {result1['error']}")
                    return
                
                self.current_project_path = result1['copied_path']
                files_count = result1.get('copied_files', 0)
                self.logger.info(f"✅ Проект скопирован (файлов: {files_count})\n")
                time.sleep(1)
                
                # Шаг 2: Переименование сборки и чертежа
                if self.cancel_requested:
                    self.logger.warning("❌ Операция прервана пользователем\n")
                    return
                
                self.logger.info("ШАГ 2/10: Переименование главной сборки и чертежа...")
                result_rename = self.copier.rename_main_assembly(self.current_project_path, project_name)
                
                if result_rename['success']:
                    self.logger.info(f"✅ Переименовано: {result_rename.get('renamed_count', 0)} файлов\n")
                else:
                    self.logger.warning(f"⚠️ Переименование: {result_rename.get('error', 'Unknown')}\n")
                
                time.sleep(1)
                
                # Шаг 3: Обновление переменных
                if self.cancel_requested:
                    self.logger.warning("❌ Операция прервана пользователем\n")
                    return
                
                self.logger.info("ШАГ 3/10: Обновление переменных...")
                result2 = self.updater.update_project_variables(self.current_project_path, h, b1, l1)
                
                if not result2['success']:
                    self.logger.error(f"❌ Ошибки обновления переменных: {result2['errors']}")
                    return
                
                self.logger.info(f"✅ Переменные обновлены (деталей: {result2['parts_updated']})\n")
                time.sleep(1)
                
                # Шаг 4: Обновление обозначений (marking) + переименование файлов
                if self.cancel_requested:
                    self.logger.warning("❌ Операция прервана пользователем\n")
                    return
                
                self.logger.info("ШАГ 4/10: Обновление обозначений и переименование...")
                result3 = self.designation_updater.update_all_designations(
                    self.current_project_path, h, b1, l1, order_number
                )
                
                if not result3['success']:
                    self.logger.error(f"❌ Ошибки обновления обозначений: {result3.get('error', 'Unknown')}")
                    self.logger.error(f"   Детали ошибок: {result3.get('errors', [])}")
                    # Продолжаем выполнение даже при ошибках
                
                self.logger.info(f"✅ Обозначения обновлены (деталей: {result3.get('parts_renamed', 0)})\n")
                time.sleep(1)
                
                # Проверка прерывания
                if self.cancel_requested:
                    self.logger.warning("❌ Операция прервана пользователем\n")
                    return
                
                # Шаг 5: Экспорт разверток в DXF
                self.logger.info("ШАГ 5/10: Экспорт разверток в DXF...")
                
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
                
                # Проверка прерывания
                if self.cancel_requested:
                    self.logger.warning("❌ Операция прервана пользователем\n")
                    return
                
                # Шаг 6: Переименование DXF
                self.logger.info("ШАГ 6/10: Переименование DXF файлов...")
                result5 = self.dxf_renamer.rename_dxf_files(self.current_project_path, order_number)
                
                if result5['success']:
                    self.logger.info(f"✅ DXF переименованы (файлов: {result5['renamed_count']})\n")
                else:
                    self.logger.warning(f"⚠️ DXF: {result5.get('errors', ['Нет DXF папки'])}\n")
                
                time.sleep(1)
                
                # Проверка прерывания
                if self.cancel_requested:
                    self.logger.warning("❌ Операция прервана пользователем\n")
                    return
                
                # Шаг 7: Обновление чертежей
                self.logger.info("ШАГ 7/10: Обновление чертежей...")
                result6 = self.drawing_updater.update_all_drawings(self.current_project_path, developer=developer_name)
                
                if result6['success']:
                    self.logger.info(f"✅ Чертежи обновлены (обновлено: {result6['drawings_updated']}, ошибок: {result6['drawings_failed']})\n")
                else:
                    self.logger.warning(f"⚠️ Чертежи: {result6.get('errors', ['Unknown'])}\n")
                
                time.sleep(1)
                
                # Проверка прерывания
                if self.cancel_requested:
                    self.logger.warning("❌ Операция прервана пользователем\n")
                    return
                
                # Шаг 8: Экспорт чертежей в BMP
                self.logger.info("ШАГ 8/10: Экспорт чертежей в BMP...")
                
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
                
                # Проверка прерывания
                if self.cancel_requested:
                    self.logger.warning("❌ Операция прервана пользователем\n")
                    return
                
                # Шаг 9: Организация BMP файлов
                self.logger.info("ШАГ 9/10: Организация BMP файлов...")
                result8 = self.bmp_organizer.organize_bmp_files(self.current_project_path)
                
                if result8['success'] and result8['moved_count'] > 0:
                    self.logger.info(f"✅ BMP файлы организованы: {result8['moved_count']} → папка BMP/\n")
                else:
                    self.logger.info(f"  Нет BMP файлов для организации\n")
                
                time.sleep(1)
                
                # Проверка прерывания
                if self.cancel_requested:
                    self.logger.warning("❌ Операция прервана пользователем\n")
                    return
                
                # Шаг 10: Итоговая информация
                self.logger.info("ШАГ 10/10: Формирование итогового отчета...")
                
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
    
    def reload_modules(self):
        """Перезагрузить все модули компонентов без перезапуска программы"""
        if self.is_processing:
            self.logger.warning("⚠️ Дождитесь завершения текущей операции!")
            return
        
        try:
            self.logger.info("\n" + "="*60)
            self.logger.info("🔄 ПЕРЕЗАГРУЗКА МОДУЛЕЙ")
            self.logger.info("="*60)
            
            # Перезагружаем все модули компонентов
            modules_to_reload = [
                ('base_component', base_component),
                ('project_copier', project_copier),
                ('cascading_variables_updater', cascading_variables_updater),
                ('designation_updater_fixed', designation_updater_fixed),
                ('dxf_renamer', dxf_renamer),
                ('drawing_auto_updater', drawing_auto_updater),
                ('drawing_exporter', drawing_exporter),
                ('unfolding_dxf_exporter', unfolding_dxf_exporter),
                ('bmp_organizer', bmp_organizer),
                ('template_manager', template_manager),
            ]
            
            reloaded_count = 0
            for module_name, module in modules_to_reload:
                try:
                    importlib.reload(module)
                    self.logger.info(f"  ✓ {module_name}")
                    reloaded_count += 1
                except Exception as e:
                    self.logger.error(f"  ✗ {module_name}: {e}")
            
            # Пересоздаем экземпляры компонентов
            self.copier = ProjectCopier()
            self.updater = CascadingVariablesUpdater()
            self.designation_updater = DesignationUpdaterFixed()
            self.dxf_exporter = UnfoldingDxfExporter()
            self.dxf_renamer = DxfRenamer()
            self.drawing_updater = DrawingAutoUpdater()
            self.drawing_exporter = DrawingExporter()
            self.bmp_organizer = BmpOrganizer()
            self.template_manager = TemplateManager()
            
            # Перенастраиваем логирование для новых компонентов
            gui_handler = None
            for handler in self.logger.handlers:
                if isinstance(handler, TextHandler):
                    gui_handler = handler
                    break
            
            if gui_handler:
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
                    # Удаляем старые обработчики
                    component_logger.handlers.clear()
                    # Добавляем GUI обработчик
                    component_logger.setLevel(logging.INFO)
                    component_logger.addHandler(gui_handler)
            
            self.logger.info("="*60)
            self.logger.info(f"✅ Перезагружено модулей: {reloaded_count}")
            self.logger.info("✅ Компоненты пересозданы")
            self.logger.info("="*60 + "\n")
            
        except Exception as e:
            self.logger.error(f"❌ Ошибка перезагрузки: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
    
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
        self.cancel_requested = False
        self.progress_bar.pack(fill="x", padx=20, pady=10)
        self.progress_bar.start()
        # Активируем кнопку "Прервать"
        self.cancel_btn.configure(state="normal")
        # Деактивируем кнопку копирования
        self.copy_btn.configure(state="disabled")
    
    def stop_processing(self):
        """Конец обработки (скрыть прогресс-бар)"""
        self.is_processing = False
        self.cancel_requested = False
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        # Деактивируем кнопку "Прервать"
        self.cancel_btn.configure(state="disabled")
        # Активируем кнопку копирования
        self.copy_btn.configure(state="normal")


def main():
    """Запуск приложения"""
    app = KompasManagerGUI()
    app.mainloop()


if __name__ == "__main__":
    main()

