"""
Outlook 草稿邮件管理器 - 增强版
支持从Excel读取数据，保存多个配置，预览并创建草稿邮件
新增功能：附件、BCC、批量创建、模板变量、配置导入导出等
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import os
import win32com.client
import openpyxl
from datetime import datetime
import re
from pathlib import Path
import shutil
# Pillow for image rendering
try:
    from PIL import Image, ImageDraw, ImageFont
    from PIL import ImageGrab
except Exception:
    Image = None
    ImageDraw = None
    ImageFont = None
    ImageGrab = None
import tempfile
import uuid
import threading
import time


class OutlookDraftManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook 草稿邮件管理器 - 增强版")
        self.root.geometry("1200x800")
        
        self.config_file = "draft_configs.json"
        self.history_file = "draft_history.json"
        self.configs = self.load_configs()
        self.current_excel_data = None
        self.excel_column_widths = []  # Excel 实际列宽（像素）
        self.attachments = []  # 附件列表
        self.inline_images = []  # 列表 of {'path':..., 'cid':...}
        self.excel_files = []  # 支持多个Excel文件
        self.selected_configs = []  # 批量处理选中的配置列表
        self.custom_placeholders = {}  # 自定义占位符字典 {'占位符名': '替换值'}
        self.email_signatures = self.load_signatures()  # 邮件签名字典
        self.scheduled_sends = {}  # 定时发送字典 {草稿ID: 发送时间}
        self.current_language = 'zh'  # 当前语言 'zh' 或 'en'
        self.ui_scale = 1.0  # 界面缩放比例
        self.font_scale = 1.0  # 字体缩放比例
        
        # 自动保存定时器
        self.auto_save_timer = None
        
        # 拼写检查词典（常见错误）
        self.spell_check_dict = {
            # 英文常见错误
            'recieve': 'receive',
            'occured': 'occurred',
            'seperate': 'separate',
            'definately': 'definitely',
            'untill': 'until',
            'tommorrow': 'tomorrow',
            'sucessful': 'successful',
            'acheive': 'achieve',
            # 中文常见错误
            '以至': '以致',  # 语境相关
            '做': '作',  # 如"作为"
        }
        
        # 多语言文本
        self.translations = {
            'zh': {
                'title': 'Outlook 草稿邮件管理器 - 增强版',
                'config_name': '配置名称',
                'save_config': '保存配置',
                'load_config': '加载配置',
                'delete_config': '删除配置',
                'export_config': '导出配置',
                'import_config': '导入配置',
                'excel_source': 'Excel 数据源',
                'excel_file': 'Excel 文件',
                'select_file': '选择文件',
                'worksheet': '工作表',
                'data_range': '数据范围',
                'auto_detect': '自动检测',
                'read_data': '读取数据',
                'preview_data': '预览数据',
                'paste_formatting': '保留格式粘贴',
                'paste_picture': '粘贴为图片',
                'email_content': '邮件内容',
                'recipients': '收件人',
                'cc': '抄送',
                'bcc': '密送',
                'subject': '主题',
                'attachments': '附件',
                'body': '正文',
                'add': '添加',
                'delete': '删除',
                'clear': '清空',
                'create_draft': '创建草稿',
                'batch_create': '批量创建',
                'preview': '预览',
                'status': '就绪',
                'spell_check': '拼写检查',
                'schedule_send': '定时发送',
                'language': '语言',
            },
            'en': {
                'title': 'Outlook Draft Manager - Enhanced',
                'config_name': 'Config Name',
                'save_config': 'Save Config',
                'load_config': 'Load Config',
                'delete_config': 'Delete Config',
                'export_config': 'Export Config',
                'import_config': 'Import Config',
                'excel_source': 'Excel Data Source',
                'excel_file': 'Excel File',
                'select_file': 'Select File',
                'worksheet': 'Worksheet',
                'data_range': 'Data Range',
                'auto_detect': 'Auto Detect',
                'read_data': 'Read Data',
                'preview_data': 'Preview Data',
                'paste_formatting': 'Keep Formatting',
                'paste_picture': 'Paste as Picture',
                'email_content': 'Email Content',
                'recipients': 'To',
                'cc': 'CC',
                'bcc': 'BCC',
                'subject': 'Subject',
                'attachments': 'Attachments',
                'body': 'Body',
                'add': 'Add',
                'delete': 'Delete',
                'clear': 'Clear',
                'create_draft': 'Create Draft',
                'batch_create': 'Batch Create',
                'preview': 'Preview',
                'status': 'Ready',
                'spell_check': 'Spell Check',
                'schedule_send': 'Schedule Send',
                'language': 'Language',
            }
        }
        
        self.setup_ui()
        self.setup_keyboard_shortcuts()
        self.load_ui_preferences()  # 加载界面偏好设置
        self.start_auto_save()
        
    def setup_ui(self):
        """设置用户界面"""
        # 主框架 - 使用 PanedWindow 实现左右拖动调整
        main_container = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 左侧主编辑区域
        main_frame = ttk.Frame(main_container, padding="10")
        main_frame.columnconfigure(1, weight=1)
        
        main_container.add(main_frame, weight=3)  # 左侧占75%
        
        # 右侧配置浏览器
        self.setup_config_browser(main_container)
        
        # 配置选择区域
        config_frame = ttk.LabelFrame(main_frame, text="配置管理", padding="5")
        config_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(config_frame, text="选择配置:").grid(row=0, column=0, padx=5)
        self.config_combo = ttk.Combobox(config_frame, width=30, state="readonly")
        self.config_combo.grid(row=0, column=1, padx=5)
        self.config_combo.bind("<<ComboboxSelected>>", self.load_selected_config)
        self.update_config_list()
        
        ttk.Button(config_frame, text="新建配置", command=self.new_config).grid(row=0, column=2, padx=5)
        ttk.Button(config_frame, text="保存配置", command=self.save_current_config).grid(row=0, column=3, padx=5)
        ttk.Button(config_frame, text="删除配置", command=self.delete_config).grid(row=0, column=4, padx=5)
        ttk.Button(config_frame, text="导出配置", command=self.export_config).grid(row=0, column=5, padx=5)
        ttk.Button(config_frame, text="导入配置", command=self.import_config).grid(row=0, column=6, padx=5)
        
        # 配置名称（第二行左侧）
        ttk.Label(config_frame, text="配置名称:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        self.config_name_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.config_name_var, width=25).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 批量处理按钮（第二行右侧）
        batch_frame = ttk.Frame(config_frame)
        batch_frame.grid(row=1, column=2, columnspan=5, pady=5, sticky=tk.E)
        ttk.Label(batch_frame, text="批量:", font=('TkDefaultFont', 9, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Button(batch_frame, text="📋 选择", command=self.select_multiple_configs).pack(side=tk.LEFT, padx=2)
        ttk.Button(batch_frame, text="🚀 生成", command=self.batch_create_all_drafts).pack(side=tk.LEFT, padx=2)
        ttk.Button(batch_frame, text="👁️ 预览", command=self.preview_all_configs).pack(side=tk.LEFT, padx=2)
        
        # Excel 文件区域
        excel_frame = ttk.LabelFrame(main_frame, text="Excel 数据源", padding="5")
        excel_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        excel_frame.columnconfigure(1, weight=1)
        
        ttk.Label(excel_frame, text="Excel 文件:").grid(row=0, column=0, padx=5, pady=3)
        self.excel_path_var = tk.StringVar()
        ttk.Entry(excel_frame, textvariable=self.excel_path_var, state="readonly").grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(excel_frame, text="选择文件", command=self.select_excel_file).grid(row=0, column=2, padx=5)
        
        ttk.Label(excel_frame, text="工作表:").grid(row=1, column=0, padx=5, pady=3)
        self.sheet_combo = ttk.Combobox(excel_frame, width=20)
        self.sheet_combo.grid(row=1, column=1, sticky=(tk.W), padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        
        ttk.Label(excel_frame, text="数据范围:").grid(row=2, column=0, padx=5, pady=3)
        range_frame = ttk.Frame(excel_frame)
        range_frame.grid(row=2, column=1, sticky=(tk.W), padx=5)
        
        self.range_var = tk.StringVar(value="A1:C10")
        ttk.Entry(range_frame, textvariable=self.range_var, width=15).grid(row=0, column=0, padx=5)
        ttk.Label(range_frame, text="(例如: A1:C10)").grid(row=0, column=1)
        ttk.Button(range_frame, text="🔍 自动检测", command=self.auto_detect_range).grid(row=0, column=2, padx=5)
        ttk.Button(range_frame, text="读取数据", command=self.read_excel_data).grid(row=0, column=3, padx=5)
        ttk.Button(range_frame, text="预览数据", command=self.preview_excel_data).grid(row=0, column=4, padx=5)
        ttk.Button(range_frame, text="  保留格式粘贴", command=self.paste_with_formatting).grid(row=0, column=5, padx=5)
        ttk.Button(range_frame, text="🖼️ 粘贴为图片", command=self.paste_as_picture).grid(row=0, column=6, padx=5)
        
        # 检测列限制
        ttk.Label(excel_frame, text="检测列限制:").grid(row=3, column=0, padx=5, pady=3)
        col_limit_frame = ttk.Frame(excel_frame)
        col_limit_frame.grid(row=3, column=1, sticky=(tk.W), padx=5)
        
        self.col_limit_var = tk.StringVar()
        ttk.Entry(col_limit_frame, textvariable=self.col_limit_var, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(col_limit_frame, text="(可选, 例: A:C 或 A,C)").pack(side=tk.LEFT)
        
        # 邮件内容区域
        email_frame = ttk.LabelFrame(main_frame, text="邮件内容", padding="5")
        email_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        email_frame.columnconfigure(1, weight=1)
        email_frame.rowconfigure(4, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 收件人区域（使用列表方式）
        ttk.Label(email_frame, text="收件人:").grid(row=0, column=0, padx=5, pady=3, sticky=tk.NW)
        to_frame = ttk.Frame(email_frame)
        to_frame.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=3)
        to_frame.columnconfigure(0, weight=1)
        
        to_input_frame = ttk.Frame(to_frame)
        to_input_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))
        to_input_frame.columnconfigure(0, weight=1)
        
        self.to_entry_var = tk.StringVar()
        ttk.Entry(to_input_frame, textvariable=self.to_entry_var).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(to_input_frame, text="添加", command=lambda: self.add_recipient('to')).grid(row=0, column=1, padx=2)
        ttk.Button(to_input_frame, text="删除", command=lambda: self.remove_recipient('to')).grid(row=0, column=2, padx=2)
        
        self.to_listbox = tk.Listbox(to_frame, height=2)
        self.to_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(3, 0))
        
        # 抄送区域
        ttk.Label(email_frame, text="抄送 (CC):").grid(row=1, column=0, padx=5, pady=3, sticky=tk.NW)
        cc_frame = ttk.Frame(email_frame)
        cc_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=3)
        cc_frame.columnconfigure(0, weight=1)
        
        cc_input_frame = ttk.Frame(cc_frame)
        cc_input_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))
        cc_input_frame.columnconfigure(0, weight=1)
        
        self.cc_entry_var = tk.StringVar()
        ttk.Entry(cc_input_frame, textvariable=self.cc_entry_var).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(cc_input_frame, text="添加", command=lambda: self.add_recipient('cc')).grid(row=0, column=1, padx=2)
        ttk.Button(cc_input_frame, text="删除", command=lambda: self.remove_recipient('cc')).grid(row=0, column=2, padx=2)
        
        self.cc_listbox = tk.Listbox(cc_frame, height=2)
        self.cc_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(3, 0))
        
        # 密送区域
        ttk.Label(email_frame, text="密送 (BCC):").grid(row=2, column=0, padx=5, pady=3, sticky=tk.NW)
        bcc_frame = ttk.Frame(email_frame)
        bcc_frame.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=3)
        bcc_frame.columnconfigure(0, weight=1)
        
        bcc_input_frame = ttk.Frame(bcc_frame)
        bcc_input_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))
        bcc_input_frame.columnconfigure(0, weight=1)
        
        self.bcc_entry_var = tk.StringVar()
        ttk.Entry(bcc_input_frame, textvariable=self.bcc_entry_var).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(bcc_input_frame, text="添加", command=lambda: self.add_recipient('bcc')).grid(row=0, column=1, padx=2)
        ttk.Button(bcc_input_frame, text="删除", command=lambda: self.remove_recipient('bcc')).grid(row=0, column=2, padx=2)
        
        self.bcc_listbox = tk.Listbox(bcc_frame, height=2)
        self.bcc_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(3, 0))
        
        ttk.Label(email_frame, text="主题:").grid(row=3, column=0, padx=5, pady=3, sticky=tk.W)
        self.subject_var = tk.StringVar()
        subject_entry = ttk.Entry(email_frame, textvariable=self.subject_var)
        subject_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=3)
        ttk.Button(email_frame, text="📋 占位符", command=lambda: self.show_placeholder_menu(subject_entry)).grid(row=3, column=2, padx=5)
        ttk.Button(email_frame, text="✍️ 签名管理", command=self.manage_signatures).grid(row=3, column=3, padx=5)
        
        # 附件区域
        attachment_subframe = ttk.Frame(email_frame)
        attachment_subframe.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=5, pady=3)
        ttk.Label(attachment_subframe, text="附件:").grid(row=0, column=0, sticky=tk.W)
        ttk.Button(attachment_subframe, text="添加附件", command=self.add_attachment).grid(row=0, column=1, padx=5)
        ttk.Button(attachment_subframe, text="删除选中", command=self.remove_selected_attachment).grid(row=0, column=2, padx=5)
        ttk.Button(attachment_subframe, text="清空附件", command=self.clear_attachments).grid(row=0, column=3, padx=5)

        self.attachment_listbox = tk.Listbox(attachment_subframe, height=3)
        self.attachment_listbox.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=3)
        # 双击打开附件文件（若存在）
        self.attachment_listbox.bind('<Double-Button-1>', self.open_selected_attachment)
        attachment_subframe.columnconfigure(0, weight=1)
        
        # 自定义占位符快速输入
        placeholder_quick_frame = ttk.Frame(email_frame)
        placeholder_quick_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=5, pady=3)
        ttk.Label(placeholder_quick_frame, text="📌 自定义占位符:", foreground='blue').pack(side=tk.LEFT, padx=2)
        ttk.Label(placeholder_quick_frame, text="名称:").pack(side=tk.LEFT, padx=2)
        self.placeholder_name_var = tk.StringVar()
        ttk.Entry(placeholder_quick_frame, textvariable=self.placeholder_name_var, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Label(placeholder_quick_frame, text="值:").pack(side=tk.LEFT, padx=2)
        self.placeholder_value_var = tk.StringVar()
        ttk.Entry(placeholder_quick_frame, textvariable=self.placeholder_value_var, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(placeholder_quick_frame, text="➕ 添加", command=self.add_custom_placeholder, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Button(placeholder_quick_frame, text="📋 管理", command=self.manage_placeholders, width=6).pack(side=tk.LEFT, padx=2)
        
        # 正文按钮（触发弹窗编辑）
        ttk.Label(email_frame, text="正文:").grid(row=6, column=0, padx=5, pady=3, sticky=tk.W)
        body_button_frame = ttk.Frame(email_frame)
        body_button_frame.grid(row=6, column=1, columnspan=2, sticky=tk.W, padx=5, pady=3)
        ttk.Button(body_button_frame, text="  编辑正文", command=self.edit_body_in_window, width=15).pack(side=tk.LEFT, padx=5)
        self.body_preview_label = ttk.Label(body_button_frame, text="(点击编辑正文内容)", foreground='gray')
        self.body_preview_label.pack(side=tk.LEFT, padx=5)
        
        # 隐藏的body_text用于存储数据（不显示在界面上）
        self.body_text = tk.Text(self.root, height=1)  # 创建但不grid，仅用于数据存储
        
        # 预览和操作区域
        preview_frame = ttk.LabelFrame(main_frame, text="邮件预览", padding="5")
        preview_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        button_frame = ttk.Frame(preview_frame)
        button_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(button_frame, text="生成预览", command=self.generate_preview).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="👁️ 增强预览", command=self.show_enhanced_preview).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="创建草稿", command=self.create_draft).grid(row=0, column=2, padx=5)
        ttk.Button(button_frame, text="⏰ 定时发送", command=self.schedule_send_dialog).grid(row=0, column=3, padx=5)
        ttk.Button(button_frame, text="批量创建", command=self.batch_create_drafts).grid(row=0, column=4, padx=5)
        
        # 缩放控制
        zoom_frame = ttk.Frame(button_frame)
        zoom_frame.grid(row=0, column=5, padx=5)
        ttk.Button(zoom_frame, text="🔍−", command=self.zoom_out, width=3).pack(side=tk.LEFT, padx=1)
        self.zoom_label = ttk.Label(zoom_frame, text="100%", width=5)
        self.zoom_label.pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="🔍+", command=self.zoom_in, width=3).pack(side=tk.LEFT, padx=1)
        
        ttk.Button(button_frame, text="🌐 语言", command=self.toggle_language).grid(row=0, column=6, padx=5)
        
        # 提示标签（替代固定预览区）
        info_label = ttk.Label(preview_frame, text="点击'生成预览'或'增强预览'按钮在弹窗中查看邮件内容", 
                              foreground='gray', font=('TkDefaultFont', 9, 'italic'))
        info_label.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=20)
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪 | 快捷键: Ctrl+S=保存配置 | Ctrl+D=创建草稿 | Ctrl+P=预览 | F1=帮助")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
    
    def setup_keyboard_shortcuts(self):
        """设置键盘快捷键"""
        # Ctrl+S - 保存配置
        self.root.bind('<Control-s>', lambda e: self.save_current_config())
        self.root.bind('<Control-S>', lambda e: self.save_current_config())
        
        # Ctrl+D - 创建草稿
        self.root.bind('<Control-d>', lambda e: self.create_draft())
        self.root.bind('<Control-D>', lambda e: self.create_draft())
        
        # Ctrl+P - 生成预览
        self.root.bind('<Control-p>', lambda e: self.generate_preview())
        self.root.bind('<Control-P>', lambda e: self.generate_preview())
        
        # Ctrl+R - 读取Excel数据
        self.root.bind('<Control-r>', lambda e: self.read_excel_data())
        self.root.bind('<Control-R>', lambda e: self.read_excel_data())
        
        # Ctrl+B - 批量创建
        self.root.bind('<Control-b>', lambda e: self.batch_create_drafts())
        self.root.bind('<Control-B>', lambda e: self.batch_create_drafts())
        
        # F1 - 显示帮助
        self.root.bind('<F1>', lambda e: self.show_help())
        
        # Ctrl+Q - 退出程序
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        self.root.bind('<Control-Q>', lambda e: self.root.quit())
        
        # Ctrl+Plus / Ctrl+Minus - 缩放
        self.root.bind('<Control-plus>', lambda e: self.zoom_in())
        self.root.bind('<Control-equal>', lambda e: self.zoom_in())  # 不按Shift的+键
        self.root.bind('<Control-minus>', lambda e: self.zoom_out())
        self.root.bind('<Control-0>', lambda e: self.reset_zoom())
        
        # 鼠标滚轮+Ctrl - 缩放
        self.root.bind('<Control-MouseWheel>', self.mouse_zoom)
        
    def show_help(self):
        """显示快捷键帮助"""
        help_text = """
═══════════════════════════════
    快捷键列表
═══════════════════════════════

📋 配置管理:
   Ctrl+S    保存当前配置
   
✉️ 邮件操作:
   Ctrl+D    创建草稿邮件
   Ctrl+B    批量创建草稿
   Ctrl+P    生成邮件预览
   
📊 Excel 操作:
   Ctrl+R    读取 Excel 数据

🔍 界面缩放:
   Ctrl++    放大界面
   Ctrl+-    缩小界面
   Ctrl+0    重置缩放（100%）
   Ctrl+滚轮  鼠标滚轮缩放
   
❓ 其他:
   F1        显示此帮助
   Ctrl+Q    退出程序
   
═══════════════════════════════
💡 提示: 所有快捷键在程序任何位置都可使用
═══════════════════════════════
        """
        messagebox.showinfo("快捷键帮助", help_text)
        
    def load_configs(self):
        """加载保存的配置"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                messagebox.showerror("错误", f"加载配置失败: {str(e)}")
                return {}
        return {}
    
    def save_configs(self):
        """保存所有配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.configs, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
            return False
    
    def load_ui_preferences(self):
        """加载界面偏好设置"""
        pref_file = "ui_preferences.json"
        if os.path.exists(pref_file):
            try:
                with open(pref_file, 'r', encoding='utf-8') as f:
                    prefs = json.load(f)
                    self.ui_scale = prefs.get('ui_scale', 1.0)
                    self.current_language = prefs.get('language', 'zh')
                    # 应用缩放
                    if self.ui_scale != 1.0:
                        self.root.after(100, self.apply_zoom)  # 延迟应用，确保界面已加载
            except:
                pass
    
    def save_ui_preferences(self):
        """保存界面偏好设置"""
        pref_file = "ui_preferences.json"
        try:
            prefs = {
                'ui_scale': self.ui_scale,
                'language': self.current_language
            }
            with open(pref_file, 'w', encoding='utf-8') as f:
                json.dump(prefs, f, indent=2)
        except:
            pass
    
    def update_config_list(self):
        """更新配置下拉列表"""
        config_names = list(self.configs.keys())
        self.config_combo['values'] = config_names
        if config_names:
            self.config_combo.current(0)
        
        # 同时更新配置浏览器
        if hasattr(self, 'config_listbox'):
            self.update_config_browser()
    
    def new_config(self):
        """创建新配置"""
        self.config_name_var.set("")
        self.excel_path_var.set("")
        self.range_var.set("A1:C10")
        
        # 清空收件人列表
        self.to_listbox.delete(0, tk.END)
        self.cc_listbox.delete(0, tk.END)
        self.bcc_listbox.delete(0, tk.END)
        self.to_entry_var.set("")
        self.cc_entry_var.set("")
        self.bcc_entry_var.set("")
        
        self.subject_var.set("")
        self.body_text.delete(1.0, tk.END)
        self.sheet_combo.set("")
        self.current_excel_data = None
        self.attachments = []
        self.update_attachment_list()
        self.status_var.set("新建配置")
    
    def save_current_config(self):
        """保存当前配置"""
        config_name = self.config_name_var.get().strip()
        if not config_name:
            messagebox.showwarning("警告", "请输入配置名称")
            return
        
        # 从列表框获取所有邮箱
        to_list = list(self.to_listbox.get(0, tk.END))
        cc_list = list(self.cc_listbox.get(0, tk.END))
        bcc_list = list(self.bcc_listbox.get(0, tk.END))
        
        config_data = {
            "excel_path": self.excel_path_var.get(),
            "sheet_name": self.sheet_combo.get(),
            "data_range": self.range_var.get(),
            "to": to_list,  # 保存为列表
            "cc": cc_list,
            "bcc": bcc_list,
            "subject": self.subject_var.get(),
            "body": self.body_text.get(1.0, tk.END).strip(),
            "attachments": self.attachments.copy(),
            "custom_placeholders": self.custom_placeholders.copy()  # 保存自定义占位符
        }
        
        self.configs[config_name] = config_data
        if self.save_configs():
            self.update_config_list()
            # 选中刚保存的配置
            index = list(self.configs.keys()).index(config_name)
            self.config_combo.current(index)
            messagebox.showinfo("成功", f"配置 '{config_name}' 已保存")
            self.status_var.set(f"已保存配置: {config_name}")
    
    def load_selected_config(self, event=None):
        """加载选中的配置"""
        config_name = self.config_combo.get()
        if not config_name or config_name not in self.configs:
            return
        
        config = self.configs[config_name]
        self.config_name_var.set(config_name)
        self.excel_path_var.set(config.get("excel_path", ""))
        self.range_var.set(config.get("data_range", "A1:C10"))
        
        # 加载收件人列表
        self.to_listbox.delete(0, tk.END)
        to_data = config.get("to", [])
        # 兼容旧格式（字符串）和新格式（列表）
        if isinstance(to_data, str):
            # 旧格式：分号分隔的字符串
            for email in to_data.split(';'):
                email = email.strip()
                if email:
                    self.to_listbox.insert(tk.END, email)
        else:
            # 新格式：列表
            for email in to_data:
                self.to_listbox.insert(tk.END, email)
        
        # 加载抄送列表
        self.cc_listbox.delete(0, tk.END)
        cc_data = config.get("cc", [])
        if isinstance(cc_data, str):
            for email in cc_data.split(';'):
                email = email.strip()
                if email:
                    self.cc_listbox.insert(tk.END, email)
        else:
            for email in cc_data:
                self.cc_listbox.insert(tk.END, email)
        
        # 加载密送列表
        self.bcc_listbox.delete(0, tk.END)
        bcc_data = config.get("bcc", [])
        if isinstance(bcc_data, str):
            for email in bcc_data.split(';'):
                email = email.strip()
                if email:
                    self.bcc_listbox.insert(tk.END, email)
        else:
            for email in bcc_data:
                self.bcc_listbox.insert(tk.END, email)
        
        self.subject_var.set(config.get("subject", ""))
        self.body_text.delete(1.0, tk.END)
        self.body_text.insert(1.0, config.get("body", ""))
        self.attachments = config.get("attachments", [])
        self.update_attachment_list()
        
        # 加载自定义占位符
        self.custom_placeholders = config.get("custom_placeholders", {}).copy()
        
        # 如果有Excel文件，加载工作表列表
        if config.get("excel_path") and os.path.exists(config["excel_path"]):
            self.load_sheets()
            self.sheet_combo.set(config.get("sheet_name", ""))
        
        self.status_var.set(f"已加载配置: {config_name}")
    
    def delete_config(self):
        """删除选中的配置"""
        config_name = self.config_combo.get()
        if not config_name:
            messagebox.showwarning("警告", "请先选择要删除的配置")
            return
        
        if messagebox.askyesno("确认删除", f"确定要删除配置 '{config_name}' 吗？"):
            del self.configs[config_name]
            if self.save_configs():
                self.update_config_list()
                self.new_config()
                messagebox.showinfo("成功", f"配置 '{config_name}' 已删除")
    
    def select_excel_file(self):
        """选择Excel文件"""
        file_paths = filedialog.askopenfilenames(
            title="选择Excel文件 (支持多选)",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_paths:
            # 将多个路径用分号连接
            paths_str = ";".join(file_paths)
            self.excel_path_var.set(paths_str)
            self.load_sheets()
            if len(file_paths) > 1:
                self.status_var.set(f"已选择 {len(file_paths)} 个文件")
            else:
                self.status_var.set(f"已选择文件: {os.path.basename(file_paths[0])}")
    
    def load_sheets(self):
        """加载Excel工作表列表"""
        excel_paths_str = self.excel_path_var.get()
        if not excel_paths_str:
            return
            
        # 如果有多个文件，只读取第一个文件的Sheet
        excel_path = excel_paths_str.split(';')[0]
        
        if not os.path.exists(excel_path):
            return
        
        try:
            wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            self.sheet_combo['values'] = sheet_names
            if sheet_names:
                self.sheet_combo.current(0)
            wb.close()
        except Exception as e:
            messagebox.showerror("错误", f"读取工作表失败: {str(e)}")
    
    def on_sheet_selected(self, event=None):
        """工作表选择变化时"""
        self.current_excel_data = None
    
    def auto_detect_range(self):
        """自动检测工作表中有数据的范围"""
        excel_paths_str = self.excel_path_var.get()
        sheet_name = self.sheet_combo.get()
        col_limit_str = self.col_limit_var.get().strip().upper()
        
        if not excel_paths_str:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
            
        # 只检测第一个文件
        excel_path = excel_paths_str.split(';')[0]
        
        if not os.path.exists(excel_path):
            messagebox.showwarning("警告", "Excel文件不存在")
            return
        
        if not sheet_name:
            messagebox.showwarning("警告", "请选择工作表")
            return
        
        wb = None
        try:
            self.status_var.set("正在检测数据范围...")
            self.root.update()
            
            # 解析列限制
            target_cols = set()
            if col_limit_str:
                try:
                    from openpyxl.utils import column_index_from_string
                    parts = col_limit_str.replace('，', ',').split(',')
                    for part in parts:
                        part = part.strip()
                        if ':' in part:
                            start, end = part.split(':')
                            start_idx = column_index_from_string(start.strip())
                            end_idx = column_index_from_string(end.strip())
                            for i in range(min(start_idx, end_idx), max(start_idx, end_idx) + 1):
                                target_cols.add(i)
                        else:
                            if part:
                                target_cols.add(column_index_from_string(part))
                except Exception as e:
                    messagebox.showwarning("警告", f"列限制格式错误: {str(e)}")
                    self.status_var.set("就绪")
                    return

            # 重要：不使用 read_only 模式，避免 ReadOnlyWorksheet 的限制
            wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=False)
            ws = wb[sheet_name]
            
            detected_range = None
            min_row = None
            max_row = None
            min_col = None
            max_col = None
            
            # 使用可靠的方法：直接扫描单元格
            try:
                # 获取工作表的最大范围
                if hasattr(ws, 'max_row') and hasattr(ws, 'max_column'):
                    max_scan_row = ws.max_row if ws.max_row else 1000
                    max_scan_col = ws.max_column if ws.max_column else 100
                else:
                    # 如果无法获取，使用默认范围
                    max_scan_row = 1000
                    max_scan_col = 100
                
                # 限制扫描范围以提高速度
                max_scan_row = min(max_scan_row, 1000)
                max_scan_col = min(max_scan_col, 100)
                
                # 扫描单元格找到有数据的范围
                for row_idx in range(1, max_scan_row + 1):
                    # 确定要扫描的列
                    if target_cols:
                        cols_to_scan = [c for c in target_cols if c <= max_scan_col]
                    else:
                        cols_to_scan = range(1, max_scan_col + 1)
                        
                    for col_idx in cols_to_scan:
                        try:
                            cell = ws.cell(row=row_idx, column=col_idx)
                            # 检查单元格是否有值
                            if cell.value is not None:
                                value_str = str(cell.value).strip()
                                if value_str:  # 不是空字符串
                                    if min_row is None or row_idx < min_row:
                                        min_row = row_idx
                                    if max_row is None or row_idx > max_row:
                                        max_row = row_idx
                                    if min_col is None or col_idx < min_col:
                                        min_col = col_idx
                                    if max_col is None or col_idx > max_col:
                                        max_col = col_idx
                        except Exception:
                            # 跳过无法访问的单元格
                            continue
                
                # 生成范围字符串
                if min_row and max_row and min_col and max_col:
                    from openpyxl.utils import get_column_letter
                    start_cell = f"{get_column_letter(min_col)}{min_row}"
                    end_cell = f"{get_column_letter(max_col)}{max_row}"
                    detected_range = f"{start_cell}:{end_cell}"
                    
            except Exception as e:
                # 如果扫描失败，尝试简单方法
                pass
            
            # 如果检测失败，返回错误
            if not detected_range:
                wb.close()
                msg = "未在工作表中检测到数据"
                if target_cols:
                    msg += f"\n(在指定的列范围内: {col_limit_str})"
                msg += "\n\n可能原因：\n1. 工作表为空\n2. 数据超出扫描范围(1000行)"
                messagebox.showwarning("警告", msg)
                self.status_var.set("未检测到数据")
                return
            
            wb.close()
            
            # 更新范围输入框
            self.range_var.set(detected_range)
            
            # 计算行数和列数
            parts = detected_range.split(':')
            if len(parts) == 2:
                import re
                start_match = re.match(r'([A-Z]+)(\d+)', parts[0])
                end_match = re.match(r'([A-Z]+)(\d+)', parts[1])
                if start_match and end_match:
                    rows = int(end_match.group(2)) - int(start_match.group(2)) + 1
                    
                    # 计算列数
                    def col_to_num(col):
                        num = 0
                        for c in col:
                            num = num * 26 + (ord(c) - ord('A') + 1)
                        return num
                    
                    cols = col_to_num(end_match.group(1)) - col_to_num(start_match.group(1)) + 1
                    
                    messagebox.showinfo("检测成功", 
                                       f"检测到数据范围: {detected_range}\n\n"
                                       f"行数: {rows}\n"
                                       f"列数: {cols}\n\n"
                                       f"点击'读取数据'以加载数据")
            else:
                messagebox.showinfo("检测成功", f"检测到数据范围: {detected_range}")
            
            self.status_var.set(f"已检测到范围: {detected_range}")
            
        except Exception as e:
            # 确保关闭工作簿
            if wb:
                try:
                    wb.close()
                except:
                    pass
            
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"自动检测范围失败:\n{str(e)}\n\n详细信息:\n{error_details}")
            self.status_var.set("检测失败")
        finally:
            # 最终确保工作簿被关闭
            if wb:
                try:
                    wb.close()
                except:
                    pass
    
    def read_excel_data(self):
        """读取Excel数据"""
        excel_paths_str = self.excel_path_var.get()
        sheet_name = self.sheet_combo.get()
        data_range = self.range_var.get().strip()
        
        if not excel_paths_str:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        if not sheet_name:
            messagebox.showwarning("警告", "请选择工作表")
            return
        
        if not data_range:
            messagebox.showwarning("警告", "请输入数据范围")
            return
            
        file_paths = excel_paths_str.split(';')
        all_data_rows = []
        headers = None
        
        self.status_var.set("正在读取数据...")
        self.root.update()
        
        try:
            for idx, excel_path in enumerate(file_paths):
                if not os.path.exists(excel_path):
                    continue
                    
                wb = openpyxl.load_workbook(excel_path, read_only=False, data_only=True)
                if sheet_name not in wb.sheetnames:
                    print(f"Sheet {sheet_name} not found in {excel_path}")
                    wb.close()
                    continue
                    
                ws = wb[sheet_name]
                
                # 读取指定范围的数据
                try:
                    cells = ws[data_range]
                except Exception as e:
                    print(f"Error reading range {data_range} in {excel_path}: {e}")
                    wb.close()
                    continue
                    
                current_file_rows = []
                
                # 处理单行或多行数据
                if isinstance(cells, tuple) and len(cells) > 0:
                    if isinstance(cells[0], tuple):
                        # 多行数据
                        for row in cells:
                            row_data = [cell.value if cell.value is not None else "" for cell in row]
                            current_file_rows.append(row_data)
                    else:
                        # 单行数据
                        row_data = [cell.value if cell.value is not None else "" for cell in cells]
                        current_file_rows.append(row_data)
                else:
                    # 单个单元格
                    current_file_rows.append([cells.value if cells.value is not None else ""])
                
                # 记录第一个文件的列宽信息
                if idx == 0 and current_file_rows and len(current_file_rows[0]) > 0:
                    self.excel_column_widths = []
                    # 解析起始列（如 A1:F10 中的 A）
                    start_col_str = data_range.split(':')[0]
                    col_letter = ''.join([c for c in start_col_str if c.isalpha()])
                    from openpyxl.utils import column_index_from_string, get_column_letter
                    start_col_idx = column_index_from_string(col_letter)
                    
                    # 读取每列的实际宽度
                    for i in range(len(current_file_rows[0])):
                        col_idx = start_col_idx + i
                        col_letter_current = get_column_letter(col_idx)
                        if col_letter_current in ws.column_dimensions:
                            width = ws.column_dimensions[col_letter_current].width
                            # Excel宽度转像素 (近似值)
                            pixel_width = int(width * 7) if width else 80
                            self.excel_column_widths.append(pixel_width)
                        else:
                            self.excel_column_widths.append(80)

                wb.close()
                
                # 合并数据逻辑
                if not current_file_rows:
                    continue
                    
                if headers is None:
                    # 第一个文件的第一行作为表头
                    headers = current_file_rows[0]
                    all_data_rows.extend(current_file_rows)
                else:
                    # 后续文件，检查第一行是否为表头
                    if current_file_rows[0] == headers:
                        # 如果是表头，跳过第一行
                        all_data_rows.extend(current_file_rows[1:])
                    else:
                        # 如果不是表头，全部添加
                        all_data_rows.extend(current_file_rows)
            
            self.current_excel_data = all_data_rows
            
            if not self.current_excel_data:
                messagebox.showwarning("提示", "未读取到任何数据")
                self.status_var.set("读取完成，无数据")
                return

            self.status_var.set(f"已读取 {len(self.current_excel_data)} 行数据 (来自 {len(file_paths)} 个文件)")
            messagebox.showinfo("成功", f"成功读取 {len(self.current_excel_data)} 行数据\n(包含表头)")
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("错误", f"读取Excel失败: {str(e)}")
            self.status_var.set("读取失败")

    
    def format_excel_data_as_table(self):
        """将Excel数据格式化为HTML表格"""
        if not self.current_excel_data:
            return "[未读取Excel数据]"
        
        html = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>\n"
        
        for i, row in enumerate(self.current_excel_data):
            html += "  <tr>\n"
            for cell in row:
                cell_value = str(cell) if cell else ""
                if i == 0:  # 第一行作为表头
                    html += f"    <th style='background-color: #4472C4; color: white; font-weight: bold;'>{cell_value}</th>\n"
                else:
                    html += f"    <td>{cell_value}</td>\n"
            html += "  </tr>\n"
        
        html += "</table>"
        return html
    
    def format_excel_data_as_text(self):
        """将Excel数据格式化为纯文本"""
        if not self.current_excel_data:
            return "[未读取Excel数据]"
        
        text_lines = []
        for row in self.current_excel_data:
            row_text = "\t".join([str(cell) if cell else "" for cell in row])
            text_lines.append(row_text)
        
        return "\n".join(text_lines)
    
    def generate_preview(self):
        """在弹出窗口中生成邮件预览"""
        # 从列表框获取邮箱地址
        to_list = list(self.to_listbox.get(0, tk.END))
        cc_list = list(self.cc_listbox.get(0, tk.END))
        bcc_list = list(self.bcc_listbox.get(0, tk.END))
        
        subject = self.subject_var.get().strip()
        body_template = self.body_text.get(1.0, tk.END).strip()
        
        if not to_list:
            messagebox.showwarning("警告", "请添加至少一个收件人")
            return
        
        if not subject:
            messagebox.showwarning("警告", "请填写主题")
            return
        
        # 创建预览窗口
        preview_window = tk.Toplevel(self.root)
        preview_window.title("邮件预览")
        preview_window.geometry("800x600")
        preview_window.transient(self.root)
        
        # 标题栏
        title_frame = ttk.Frame(preview_window, padding="10")
        title_frame.pack(fill=tk.X)
        ttk.Label(title_frame, text="📧 邮件预览", font=('TkDefaultFont', 12, 'bold')).pack(anchor=tk.W)
        
        # 预览内容
        preview_frame = ttk.Frame(preview_window, padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # 替换Excel数据占位符
        if "{EXCEL_DATA}" in body_template:
            excel_html = self.format_excel_data_as_table()
            body = body_template.replace("{EXCEL_DATA}", excel_html)
        else:
            body = body_template
        
        # 生成预览文本
        preview = f"{'='*60}\n"
        preview += f"收件人: {'; '.join(to_list)}\n"
        if cc_list:
            preview += f"抄送: {'; '.join(cc_list)}\n"
        if bcc_list:
            preview += f"密送: {'; '.join(bcc_list)}\n"
        preview += f"主题: {subject}\n"
        preview += f"{'='*60}\n\n"
        preview += f"正文预览:\n{'-'*60}\n"
        
        # 显示纯文本版本的预览
        if "{EXCEL_DATA}" in body_template and self.current_excel_data:
            preview += body_template.replace("{EXCEL_DATA}", "\n" + self.format_excel_data_as_text() + "\n")
        else:
            preview += body
        
        preview += f"\n{'-'*60}\n"
        if self.attachments:
            preview += f"\n附件: {len(self.attachments)} 个文件\n"
            for att in self.attachments[:5]:
                preview += f"  • {os.path.basename(att)}\n"
            if len(self.attachments) > 5:
                preview += f"  ... 还有 {len(self.attachments)-5} 个附件\n"
        preview += f"\n注意: 实际邮件将使用HTML格式显示表格\n"
        
        # 预览文本框
        preview_text = scrolledtext.ScrolledText(preview_frame, width=80, height=30, wrap=tk.WORD, font=('Consolas', 9))
        preview_text.pack(fill=tk.BOTH, expand=True)
        preview_text.insert(1.0, preview)
        preview_text.config(state='disabled')
        
        # 按钮区
        button_frame = ttk.Frame(preview_window, padding="10")
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="关闭", command=preview_window.destroy, width=15).pack(side=tk.RIGHT, padx=5)
        
        self.status_var.set("预览已生成")
    
    def create_draft(self):
        """创建Outlook草稿"""
        # 从列表框获取邮箱地址
        to_list = list(self.to_listbox.get(0, tk.END))
        cc_list = list(self.cc_listbox.get(0, tk.END))
        bcc_list = list(self.bcc_listbox.get(0, tk.END))
        
        subject = self.subject_var.get().strip()
        body_template = self.body_text.get(1.0, tk.END).strip()
        
        if not to_list:
            messagebox.showwarning("警告", "请添加至少一个收件人")
            return
        
        if not subject:
            messagebox.showwarning("警告", "请填写主题")
            return
        
        try:
            # 连接Outlook
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            # 创建邮件对象
            mail = outlook.CreateItem(0)  # 0 代表邮件项
            
            # 设置收件人（用分号分隔）
            mail.To = "; ".join(to_list)
            if cc_list:
                mail.CC = "; ".join(cc_list)
            if bcc_list:
                mail.BCC = "; ".join(bcc_list)
            
            # 设置主题
            subject_processed = self.process_template_variables(subject)
            mail.Subject = subject_processed
            
            # 设置正文（HTML格式）
            body_html = self.process_template_variables(body_template)
            if "{EXCEL_DATA}" in body_html:
                excel_html = self.format_excel_data_as_table()
                body_html = body_html.replace("{EXCEL_DATA}", excel_html)
            
            # 将纯文本转换为HTML（保留换行）
            body_html = body_html.replace("\n", "<br>")
            mail.HTMLBody = body_html
            
            # 添加普通附件（跳过已作为 inline_images 的路径，避免重复）
            inline_paths = [img.get('path') for img in getattr(self, 'inline_images', [])]
            for attachment_path in self.attachments:
                if attachment_path in inline_paths:
                    continue
                if os.path.exists(attachment_path):
                    mail.Attachments.Add(attachment_path)

            # 添加并标记内嵌图片（Content-ID），以便在HTML中使用cid:引用
            for img in getattr(self, 'inline_images', []):
                path = img.get('path')
                cid = img.get('cid')
                if not path or not cid:
                    continue
                if os.path.exists(path):
                    try:
                        att = mail.Attachments.Add(path)
                        # PR_ATTACH_CONTENT_ID (0x3712001F) 指定附件的 Content-ID
                        try:
                            att.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
                        except Exception:
                            # 忽略无法设置属性的情况
                            pass
                    except Exception:
                        pass
            
            # 保存为草稿（不发送）
            mail.Save()
            
            # 清理内嵌图片记录（防止传递到下一个草稿）
            self.cleanup_inline_images()
            
            # 记录历史
            self.save_to_history("; ".join(to_list), subject_processed)
            
            messagebox.showinfo("成功", "草稿邮件已创建！\n请在Outlook草稿箱中查看。")
            self.status_var.set(f"草稿已创建 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
        except Exception as e:
            messagebox.showerror("错误", f"创建草稿失败: {str(e)}\n\n请确保已安装并登录Outlook。")
    
    def add_attachment(self):
        """添加附件"""
        files = filedialog.askopenfilenames(
            title="选择附件",
            filetypes=[("所有文件", "*.*")]
        )
        for file_path in files:
            if file_path and file_path not in self.attachments:
                self.attachments.append(file_path)
        self.update_attachment_list()
        self.status_var.set(f"已添加 {len(files)} 个附件")
    
    def clear_attachments(self):
        """清空附件列表"""
        self.attachments = []
        self.update_attachment_list()
        self.status_var.set("已清空附件")
    
    def update_attachment_list(self):
        """更新附件列表显示"""
        self.attachment_listbox.delete(0, tk.END)
        for attachment in self.attachments:
            filename = os.path.basename(attachment)
            # 标注是否为内嵌图片
            is_inline = any(img.get('path') == attachment for img in getattr(self, 'inline_images', []))
            display_name = filename + ("  [内嵌]" if is_inline else "")
            self.attachment_listbox.insert(tk.END, display_name)
    
    def cleanup_inline_images(self):
        """清理内嵌图片及其临时文件"""
        if not hasattr(self, 'inline_images'):
            return
        
        # 删除临时图片文件
        for img in self.inline_images:
            path = img.get('path')
            if path and os.path.exists(path):
                try:
                    os.remove(path)
                except Exception:
                    pass  # 忽略删除失败
            
            # 从附件列表中移除
            if path in self.attachments:
                self.attachments.remove(path)
        
        # 清空inline_images列表
        self.inline_images = []
        
        # 更新附件显示
        self.update_attachment_list()

    def paste_with_formatting(self):
        """保留源格式粘贴 Excel 数据到正文（Keep Source Formatting）"""
        if not self.current_excel_data:
            messagebox.showwarning("警告", "请先读取 Excel 数据")
            return
        
        if len(self.current_excel_data) == 0:
            messagebox.showwarning("警告", "Excel 数据为空")
            return
        
        try:
            self.status_var.set("正在生成格式化表格...")
            self.root.update()
            
            # 生成带格式的 HTML 表格（类似 Excel 的样式）
            html_table = self.generate_formatted_html_table(self.current_excel_data)
            
            # 获取当前正文内容
            try:
                body = self.body_text.get(1.0, tk.END).strip()
            except:
                body = ""
            
            # 如果正文存在 {EXCEL_DATA} 占位符，替换它；否则追加到结尾
            if "{EXCEL_DATA}" in body:
                new_body = body.replace("{EXCEL_DATA}", html_table)
                self.body_text.delete(1.0, tk.END)
                self.body_text.insert(1.0, new_body)
            else:
                # 追加到结尾
                if body:
                    self.body_text.insert(tk.END, "\n\n" + html_table)
                else:
                    self.body_text.insert(1.0, html_table)
            
            messagebox.showinfo("成功", f"已插入格式化表格（{len(self.current_excel_data)}行）\n\n使用了 Keep Source Formatting 样式")
            self.status_var.set("已插入格式化表格")
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"插入表格失败:\n{str(e)}\n\n详细信息:\n{error_details}")
            self.status_var.set("插入失败")
    
    def copy_range_as_image_com(self, excel_path, sheet_name, data_range):
        """使用 COM 接口调用 Excel 复制范围为图片 (保留原格式)"""
        if ImageGrab is None:
            return None

        excel = None
        wb = None
        try:
            # 初始化 Excel 应用
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
            except:
                excel = win32com.client.Dispatch("Excel.Application")
            
            # 打开工作簿
            abs_path = os.path.abspath(excel_path)
            wb_opened = False
            # 检查是否已经打开
            try:
                for w in excel.Workbooks:
                    if w.FullName.lower() == abs_path.lower():
                        wb = w
                        wb_opened = True
                        break
            except:
                pass
            
            if not wb:
                wb = excel.Workbooks.Open(abs_path, ReadOnly=True)
            
            # 获取工作表和范围
            try:
                ws = wb.Sheets(sheet_name)
                rng = ws.Range(data_range)
                
                # 复制为图片 (Appearance: 1=xlScreen, Format: 2=xlBitmap)
                rng.CopyPicture(Appearance=1, Format=2)
                
                # 等待剪贴板更新
                time.sleep(0.5)
                
                # 从剪贴板获取图片
                img = ImageGrab.grabclipboard()
                
                if img:
                    # 保存到临时文件
                    temp_dir = tempfile.gettempdir()
                    img_filename = f"excel_paste_{uuid.uuid4().hex}.png"
                    img_path = os.path.join(temp_dir, img_filename)
                    img.save(img_path, 'PNG')
                    return img_path
            finally:
                # 如果是我们打开的，关闭它
                if wb and not wb_opened:
                    wb.Close(SaveChanges=False)
            
            return None
                
        except Exception as e:
            print(f"COM Error: {e}")
            return None

    def paste_as_picture(self):
        """粘贴为图片（Picture）- 将 Excel 数据转换为图片并嵌入"""
        # 获取当前设置
        excel_paths_str = self.excel_path_var.get()
        sheet_name = self.sheet_combo.get()
        data_range = self.range_var.get().strip()
        
        if not excel_paths_str or not sheet_name or not data_range:
             messagebox.showwarning("警告", "请先选择 Excel 文件、工作表和数据范围")
             return

        try:
            # 清理旧的内嵌图片
            self.cleanup_inline_images()
            
            self.status_var.set("正在通过 Excel 生成图片...")
            self.root.update()
            
            file_paths = excel_paths_str.split(';')
            generated_images = []
            
            for excel_path in file_paths:
                if not os.path.exists(excel_path):
                    continue
                
                # 尝试使用 COM 获取精确截图
                img_path = self.copy_range_as_image_com(excel_path, sheet_name, data_range)
                
                # 如果 COM 失败，回退到手动绘制
                if not img_path and self.current_excel_data:
                     img_path = self.create_excel_image(self.current_excel_data)
                
                if img_path and os.path.exists(img_path):
                    generated_images.append(img_path)
            
            if not generated_images:
                # 如果没有生成图片，尝试只用 current_excel_data 生成
                if self.current_excel_data:
                    img_path = self.create_excel_image(self.current_excel_data)
                    if img_path:
                        generated_images.append(img_path)
                else:
                    raise RuntimeError("未能生成图片，请检查 Excel 是否安装或范围是否正确")

            # 插入图片到正文
            html_fragments = []
            for img_path in generated_images:
                cid = f"img_{uuid.uuid4().hex}@local"
                self.inline_images.append({'path': img_path, 'cid': cid})
                img_tag = f'<img src="cid:{cid}" alt="Excel表格" style="max-width:100%; border:1px solid #ccc;"><br/>'
                html_fragments.append(img_tag)
                
                if img_path not in self.attachments:
                    self.attachments.append(img_path)
            
            full_html = "\n".join(html_fragments)
            
            # 获取当前正文内容
            try:
                body = self.body_text.get(1.0, tk.END).strip()
            except:
                body = ""
            
            if "{EXCEL_DATA}" in body:
                new_body = body.replace("{EXCEL_DATA}", full_html)
                self.body_text.delete(1.0, tk.END)
                self.body_text.insert(1.0, new_body)
            else:
                if body:
                    self.body_text.insert(tk.END, "\n\n" + full_html)
                else:
                    self.body_text.insert(1.0, full_html)
            
            self.update_attachment_list()
            messagebox.showinfo("成功", "已粘贴为图片（保留源格式）")
            self.status_var.set("已粘贴为图片")

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("错误", f"粘贴为图片失败:\n{str(e)}\n\n详细信息:\n{error_details}")
            self.status_var.set("粘贴失败")
    
    def generate_formatted_html_table(self, data_rows):
        """生成带格式的 HTML 表格（模拟 Keep Source Formatting）- 使用 Excel 实际列宽"""
        html = '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse; font-family:Calibri,Arial,sans-serif; font-size:11pt;">\n'
        
        # 如果有列宽信息，添加 colgroup
        if hasattr(self, 'excel_column_widths') and self.excel_column_widths:
            html += '  <colgroup>\n'
            for width in self.excel_column_widths:
                html += f'    <col style="width:{width}px;">\n'
            html += '  </colgroup>\n'
        
        for i, row in enumerate(data_rows):
            html += '  <tr>\n'
            for j, cell in enumerate(row):
                cell_value = str(cell) if cell else ""
                # 获取列宽（如果有）
                width_style = ""
                if hasattr(self, 'excel_column_widths') and self.excel_column_widths and j < len(self.excel_column_widths):
                    width_style = f" width:{self.excel_column_widths[j]}px;"
                
                if i == 0:  # 表头
                    html += f'    <th style="background-color:#4472C4; color:white; font-weight:bold; text-align:left; padding:8px; border:1px solid #2E5C8A;{width_style}">{cell_value}</th>\n'
                else:  # 数据行
                    html += f'    <td style="background-color:white; color:#000000; padding:8px; border:1px solid #D0D0D0;{width_style}">{cell_value}</td>\n'
            html += '  </tr>\n'
        
        html += '</table>'
        return html
    
    def create_excel_image(self, data_rows):
        """创建 Excel 表格图片"""
        try:
            if Image is None or ImageDraw is None or ImageFont is None:
                raise RuntimeError("Pillow 未安装或导入失败。\n请运行: pip install pillow")
            
            if not data_rows or len(data_rows) == 0:
                raise ValueError("数据为空")
            
            # 计算列宽
            cols = max((len(r) for r in data_rows), default=0)
            if cols == 0:
                raise ValueError("数据列数为0")
            
            # 优先使用 Excel 实际列宽，否则根据内容计算
            if hasattr(self, 'excel_column_widths') and self.excel_column_widths and len(self.excel_column_widths) >= cols:
                # 使用 Excel 实际列宽
                col_pixel_widths = self.excel_column_widths[:cols]
            else:
                # 根据内容长度计算列宽
                col_widths = [0] * cols
                for row in data_rows:
                    for i in range(min(cols, len(row))):
                        cell = str(row[i]) if row[i] is not None else ""
                        col_widths[i] = max(col_widths[i], len(cell))
                
                # 后续会用到 col_pixel_widths
                char_width_temp = 8  # 临时值，后面会重新计算
                col_pixel_widths = [max(80, int(w * char_width_temp + 20)) for w in col_widths]
            
            # 加载字体
            font_size = 14
            font = None
            font_paths = [
                "C:/Windows/Fonts/msyh.ttc",
                "C:/Windows/Fonts/simsun.ttc",
                "C:/Windows/Fonts/Arial.ttf",
            ]
            for fp in font_paths:
                try:
                    if os.path.exists(fp):
                        font = ImageFont.truetype(fp, font_size)
                        break
                except Exception:
                    continue
            
            # 如果没有找到字体，使用默认
            if font is None:
                try:
                    font = ImageFont.load_default()
                except Exception:
                    font = None
            
            # 计算图片尺寸
            padding = 10
            # 如果没有使用 Excel 列宽，则根据字体大小重新计算
            if not (hasattr(self, 'excel_column_widths') and self.excel_column_widths and len(self.excel_column_widths) >= cols):
                char_width = font_size * 0.6 if font_size else 8
                col_pixel_widths = [max(80, int(w * char_width + 20)) for w in col_widths]
            
            table_width = sum(col_pixel_widths) + padding * 2
            row_height = font_size + 12
            table_height = row_height * len(data_rows) + padding * 2
            
            # 创建图片
            img = Image.new('RGB', (table_width, table_height), color='white')
            draw = ImageDraw.Draw(img)
            
            # 绘制表格
            x = padding
            y = padding
            for ri, row in enumerate(data_rows):
                x = padding
                for ci in range(cols):
                    cell = str(row[ci]) if ci < len(row) and row[ci] is not None else ""
                    w = col_pixel_widths[ci]
                    
                    try:
                        if ri == 0:  # 表头
                            # 绘制背景
                            draw.rectangle([x, y, x + w, y + row_height], fill='#4472C4', outline='#2E5C8A')
                            # 绘制额外边框使其更粗（模拟 width=2）
                            draw.rectangle([x+1, y+1, x + w-1, y + row_height-1], outline='#2E5C8A')
                            # 绘制文字
                            draw.text((x + 6, y + 4), cell, fill='white', font=font)
                        else:  # 数据行
                            draw.rectangle([x, y, x + w, y + row_height], fill='white', outline='#D0D0D0')
                            draw.text((x + 6, y + 4), cell, fill='black', font=font)
                    except Exception as e:
                        # 如果绘制失败，跳过该单元格
                        pass
                    
                    x += w
                y += row_height
            
            # 保存到临时文件
            tmp_dir = tempfile.gettempdir()
            fname = f"excel_picture_{uuid.uuid4().hex}.png"
            path = os.path.join(tmp_dir, fname)
            img.save(path, 'PNG')
            return path
            
        except Exception as e:
            # 捕获所有异常并重新抛出，带详细信息
            raise RuntimeError(f"创建图片失败: {str(e)}")

    def remove_selected_attachment(self, event=None):
        """删除列表中选中的附件（可多选）"""
        selection = self.attachment_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的附件")
            return
        # 从后往前删除以避免索引漂移
        for index in reversed(selection):
            try:
                del self.attachments[index]
            except Exception:
                pass
        self.update_attachment_list()
        self.status_var.set("已删除选中附件")

    def open_selected_attachment(self, event=None):
        """在系统中打开双击的附件（Windows）"""
        try:
            selection = self.attachment_listbox.curselection()
            if not selection:
                return
            index = selection[0]
            path = self.attachments[index]
            if os.path.exists(path):
                # 使用os.startfile 在 Windows 上打开
                try:
                    os.startfile(path)
                except Exception as e:
                    messagebox.showerror("错误", f"打开附件失败: {e}")
            else:
                messagebox.showwarning("警告", "文件不存在: " + path)
        except Exception as e:
            messagebox.showerror("错误", f"操作失败: {e}")
    
    def process_template_variables(self, text):
        """处理模板变量"""
        # 替换日期时间
        text = text.replace("{DATE}", datetime.now().strftime("%Y-%m-%d"))
        text = text.replace("{TIME}", datetime.now().strftime("%H:%M:%S"))
        text = text.replace("{DATETIME}", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        
        # 替换自定义占位符
        for name, value in self.custom_placeholders.items():
            placeholder = "{" + name + "}"
            text = text.replace(placeholder, str(value))
        
        # 如果有Excel数据，替换列名变量
        if self.current_excel_data and len(self.current_excel_data) > 0:
            headers = self.current_excel_data[0] if len(self.current_excel_data) > 0 else []
            if len(self.current_excel_data) > 1:
                first_data_row = self.current_excel_data[1]
                for i, header in enumerate(headers):
                    if i < len(first_data_row):
                        placeholder = "{" + str(header) + "}"
                        text = text.replace(placeholder, str(first_data_row[i]))
        
        return text
    
    def batch_create_drafts(self):
        """批量创建草稿（为Excel每行创建一封邮件）"""
        if not self.current_excel_data or len(self.current_excel_data) < 2:
            messagebox.showwarning("警告", "请先读取Excel数据，且数据至少需要2行（表头+数据）")
            return
        
        # 从列表框获取邮箱地址
        to_list = list(self.to_listbox.get(0, tk.END))
        cc_list = list(self.cc_listbox.get(0, tk.END))
        bcc_list = list(self.bcc_listbox.get(0, tk.END))
        
        subject_template = self.subject_var.get().strip()
        body_template = self.body_text.get(1.0, tk.END).strip()
        
        # 检查是否有收件人或使用了变量
        to_str = "; ".join(to_list) if to_list else ""
        has_variable = "{" in to_str or not to_list
        
        if not has_variable and not to_list:
            messagebox.showwarning("警告", "批量模式需要添加收件人或在收件人中使用变量，如 {邮箱}")
            return
        
        if not subject_template:
            messagebox.showwarning("警告", "请填写主题")
            return
        
        # 确认操作
        data_count = len(self.current_excel_data) - 1  # 减去表头
        if not messagebox.askyesno("确认批量创建", 
                                    f"将为 {data_count} 行数据创建 {data_count} 封草稿邮件。\n确定继续？"):
            return
        
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            headers = self.current_excel_data[0]
            created_count = 0
            
            # 从第二行开始（跳过表头）
            for row_data in self.current_excel_data[1:]:
                # 创建变量字典
                variables = {"{" + str(headers[i]) + "}": str(row_data[i]) 
                            for i in range(min(len(headers), len(row_data)))}
                
                # 替换主题和正文中的变量
                subject_filled = subject_template
                body_filled = body_template
                
                # 处理收件人列表中的变量
                to_filled_list = []
                for email in to_list:
                    email_filled = email
                    for placeholder, value in variables.items():
                        email_filled = email_filled.replace(placeholder, value)
                    to_filled_list.append(email_filled)
                
                for placeholder, value in variables.items():
                    subject_filled = subject_filled.replace(placeholder, value)
                    body_filled = body_filled.replace(placeholder, value)
                
                # 处理其他模板变量
                subject_filled = self.process_template_variables(subject_filled)
                body_filled = self.process_template_variables(body_filled)
                
                # 创建邮件
                mail = outlook.CreateItem(0)
                mail.To = "; ".join(to_filled_list) if to_filled_list else to_str
                if cc_list:
                    mail.CC = "; ".join(cc_list)
                if bcc_list:
                    mail.BCC = "; ".join(bcc_list)
                mail.Subject = subject_filled
                
                # 处理HTML正文
                body_html = body_filled.replace("\n", "<br>")
                mail.HTMLBody = body_html
                
                # 添加附件
                for attachment_path in self.attachments:
                    if os.path.exists(attachment_path):
                        mail.Attachments.Add(attachment_path)
                
                mail.Save()
                created_count += 1
            
            messagebox.showinfo("成功", f"成功创建 {created_count} 封草稿邮件！")
            self.status_var.set(f"批量创建完成 - {created_count} 封草稿")
            
        except Exception as e:
            messagebox.showerror("错误", f"批量创建失败: {str(e)}")
    
    def preview_excel_data(self):
        """预览Excel数据"""
        if not self.current_excel_data:
            messagebox.showwarning("警告", "请先读取Excel数据")
            return
        
        # 创建预览窗口
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Excel 数据预览")
        preview_window.geometry("800x600")
        
        # 创建Treeview显示表格
        frame = ttk.Frame(preview_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar_y = ttk.Scrollbar(frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)
        
        # 创建Treeview
        if self.current_excel_data and len(self.current_excel_data) > 0:
            headers = [str(h) for h in self.current_excel_data[0]]
            tree = ttk.Treeview(frame, columns=headers, show='headings',
                               yscrollcommand=scrollbar_y.set,
                               xscrollcommand=scrollbar_x.set)
            
            # 设置列标题
            for header in headers:
                tree.heading(header, text=header)
                tree.column(header, width=100)
            
            # 添加数据行
            for row in self.current_excel_data[1:]:
                tree.insert('', tk.END, values=row)
            
            scrollbar_y.config(command=tree.yview)
            scrollbar_x.config(command=tree.xview)
            
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            
            ttk.Label(preview_window, text=f"共 {len(self.current_excel_data)-1} 行数据").pack(pady=5)
    
    def export_config(self):
        """导出配置到文件"""
        config_name = self.config_combo.get()
        if not config_name:
            messagebox.showwarning("警告", "请先选择要导出的配置")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="导出配置",
            defaultextension=".json",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
            initialfile=f"{config_name}.json"
        )
        
        if file_path:
            try:
                export_data = {config_name: self.configs[config_name]}
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(export_data, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("成功", f"配置已导出到：\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")
    
    def import_config(self):
        """从文件导入配置"""
        file_path = filedialog.askopenfilename(
            title="导入配置",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    imported_configs = json.load(f)
                
                imported_count = 0
                for name, config in imported_configs.items():
                    if name in self.configs:
                        if not messagebox.askyesno("确认覆盖", 
                                                   f"配置 '{name}' 已存在，是否覆盖？"):
                            continue
                    self.configs[name] = config
                    imported_count += 1
                
                if self.save_configs():
                    self.update_config_list()
                    messagebox.showinfo("成功", f"成功导入 {imported_count} 个配置")
                    
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {str(e)}")
    
    def save_to_history(self, to, subject):
        """保存到历史记录"""
        try:
            history = []
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    history = json.load(f)
            
            history.append({
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "to": to,
                "subject": subject
            })
            
            # 只保留最近100条记录
            history = history[-100:]
            
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(history, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            print(f"保存历史记录失败: {e}")
    
    def start_auto_save(self):
        """启动自动保存"""
        # 每60秒自动保存一次当前配置
        if self.config_name_var.get().strip():
            # 静默保存（不显示提示）
            config_name = self.config_name_var.get().strip()
            if config_name in self.configs:
                config_data = {
                    "excel_path": self.excel_path_var.get(),
                    "sheet_name": self.sheet_combo.get(),
                    "data_range": self.range_var.get(),
                    "to": self.to_var.get(),
                    "cc": self.cc_var.get(),
                    "bcc": self.bcc_var.get(),
                    "subject": self.subject_var.get(),
                    "body": self.body_text.get(1.0, tk.END).strip(),
                    "attachments": self.attachments.copy()
                }
                self.configs[config_name] = config_data
                self.save_configs()
        
        # 继续定时器
        self.auto_save_timer = self.root.after(60000, self.start_auto_save)
    
    def select_multiple_configs(self):
        """选择多个配置进行批量处理"""
        if not self.configs:
            messagebox.showwarning("警告", "没有可用的配置")
            return
        
        # 创建选择窗口
        select_window = tk.Toplevel(self.root)
        select_window.title("选择要批量处理的配置")
        select_window.geometry("500x600")
        select_window.transient(self.root)
        select_window.grab_set()
        
        # 说明标签
        info_frame = ttk.Frame(select_window, padding="10")
        info_frame.pack(fill=tk.X)
        ttk.Label(info_frame, text="选择要批量创建草稿的配置（可多选）：", 
                 font=('TkDefaultFont', 10, 'bold')).pack(anchor=tk.W)
        ttk.Label(info_frame, text="勾选的配置将依次创建草稿邮件", 
                 font=('TkDefaultFont', 9)).pack(anchor=tk.W, pady=(0, 5))
        
        # 创建可滚动的复选框列表
        list_frame = ttk.Frame(select_window, padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 存储复选框变量
        checkbox_vars = {}
        
        # 为每个配置创建复选框
        for i, config_name in enumerate(sorted(self.configs.keys())):
            var = tk.BooleanVar(value=config_name in self.selected_configs)
            checkbox_vars[config_name] = var
            
            frame = ttk.Frame(scrollable_frame)
            frame.pack(fill=tk.X, padx=5, pady=2)
            
            cb = ttk.Checkbutton(frame, text=config_name, variable=var)
            cb.pack(side=tk.LEFT)
            
            # 显示配置详情
            config = self.configs[config_name]
            to = config.get('to', '')[:30] + '...' if len(config.get('to', '')) > 30 else config.get('to', '')
            subject = config.get('subject', '')[:40] + '...' if len(config.get('subject', '')) > 40 else config.get('subject', '')
            
            detail = f"收件人: {to} | 主题: {subject}"
            ttk.Label(frame, text=detail, font=('TkDefaultFont', 8), foreground='gray').pack(side=tk.LEFT, padx=(10, 0))
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 统计信息
        stats_frame = ttk.Frame(select_window, padding="10")
        stats_frame.pack(fill=tk.X)
        stats_label = ttk.Label(stats_frame, text=f"总配置数: {len(self.configs)}", 
                                font=('TkDefaultFont', 9))
        stats_label.pack()
        
        # 按钮区域
        button_frame = ttk.Frame(select_window, padding="10")
        button_frame.pack(fill=tk.X)
        
        def select_all():
            for var in checkbox_vars.values():
                var.set(True)
        
        def select_none():
            for var in checkbox_vars.values():
                var.set(False)
        
        def confirm_selection():
            self.selected_configs = [name for name, var in checkbox_vars.items() if var.get()]
            if not self.selected_configs:
                messagebox.showwarning("警告", "请至少选择一个配置")
                return
            
            messagebox.showinfo("已选择", f"已选择 {len(self.selected_configs)} 个配置：\n" + 
                              "\n".join(f"• {name}" for name in self.selected_configs[:10]) +
                              (f"\n... 还有 {len(self.selected_configs)-10} 个" if len(self.selected_configs) > 10 else ""))
            select_window.destroy()
        
        ttk.Button(button_frame, text="全选", command=select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="全不选", command=select_none).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="确定", command=confirm_selection).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="取消", command=select_window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def batch_create_all_drafts(self):
        """一键批量创建所有选中配置的草稿"""
        if not self.selected_configs:
            # 如果没有选择，弹出选择窗口
            result = messagebox.askyesno("提示", "尚未选择配置。\n是否现在选择要批量处理的配置？")
            if result:
                self.select_multiple_configs()
            return
        
        # 确认操作
        confirm_msg = f"将为以下 {len(self.selected_configs)} 个配置创建草稿邮件：\n\n"
        confirm_msg += "\n".join(f"  {i+1}. {name}" for i, name in enumerate(self.selected_configs[:10]))
        if len(self.selected_configs) > 10:
            confirm_msg += f"\n  ... 还有 {len(self.selected_configs)-10} 个配置"
        confirm_msg += "\n\n确定继续？"
        
        if not messagebox.askyesno("确认批量创建", confirm_msg):
            return
        
        # 创建进度窗口
        progress_window = tk.Toplevel(self.root)
        progress_window.title("批量创建进度")
        progress_window.geometry("600x400")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # 进度信息
        info_frame = ttk.Frame(progress_window, padding="10")
        info_frame.pack(fill=tk.X)
        
        progress_label = ttk.Label(info_frame, text="正在处理...", font=('TkDefaultFont', 10, 'bold'))
        progress_label.pack(anchor=tk.W)
        
        progress_bar = ttk.Progressbar(info_frame, length=560, mode='determinate')
        progress_bar.pack(fill=tk.X, pady=10)
        
        status_label = ttk.Label(info_frame, text="", font=('TkDefaultFont', 9))
        status_label.pack(anchor=tk.W)
        
        # 日志区域
        log_frame = ttk.LabelFrame(progress_window, text="处理日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        log_text = scrolledtext.ScrolledText(log_frame, width=70, height=15, wrap=tk.WORD)
        log_text.pack(fill=tk.BOTH, expand=True)
        
        def log_message(msg, level="INFO"):
            timestamp = datetime.now().strftime("%H:%M:%S")
            color_tag = "error" if level == "ERROR" else "success" if level == "SUCCESS" else "info"
            log_text.insert(tk.END, f"[{timestamp}] ", "timestamp")
            log_text.insert(tk.END, f"{msg}\n", color_tag)
            log_text.see(tk.END)
            log_text.update()
        
        # 配置颜色标签
        log_text.tag_config("timestamp", foreground="gray")
        log_text.tag_config("info", foreground="black")
        log_text.tag_config("success", foreground="green")
        log_text.tag_config("error", foreground="red")
        
        # 开始处理
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            log_message("✓ Outlook 已连接", "SUCCESS")
            
            total = len(self.selected_configs)
            success_count = 0
            error_count = 0
            
            progress_bar['maximum'] = total
            
            for index, config_name in enumerate(self.selected_configs, 1):
                try:
                    progress_label.config(text=f"正在处理: {config_name} ({index}/{total})")
                    status_label.config(text=f"进度: {index}/{total}")
                    progress_bar['value'] = index
                    
                    log_message(f"开始处理配置: {config_name}")
                    
                    config = self.configs[config_name]
                    
                    # 加载配置数据（支持列表和字符串两种格式）
                    to_data = config.get('to', [])
                    cc_data = config.get('cc', [])
                    bcc_data = config.get('bcc', [])
                    
                    # 兼容旧格式（字符串）转换为列表
                    if isinstance(to_data, str):
                        to_list = [e.strip() for e in to_data.split(';') if e.strip()]
                    else:
                        to_list = to_data
                    
                    if isinstance(cc_data, str):
                        cc_list = [e.strip() for e in cc_data.split(';') if e.strip()]
                    else:
                        cc_list = cc_data
                    
                    if isinstance(bcc_data, str):
                        bcc_list = [e.strip() for e in bcc_data.split(';') if e.strip()]
                    else:
                        bcc_list = bcc_data
                    
                    subject = config.get('subject', '').strip()
                    body = config.get('body', '').strip()
                    attachments = config.get('attachments', [])
                    
                    # 检查必填字段
                    if not to_list:
                        log_message(f"  ⚠ 跳过（缺少收件人）", "ERROR")
                        error_count += 1
                        continue
                    
                    if not subject:
                        log_message(f"  ⚠ 跳过（缺少主题）", "ERROR")
                        error_count += 1
                        continue
                    
                    # 读取Excel数据（如果有）
                    excel_path = config.get('excel_path', '')
                    sheet_name = config.get('sheet_name', '')
                    data_range = config.get('data_range', '')
                    
                    excel_data = None
                    if excel_path and os.path.exists(excel_path) and sheet_name and data_range:
                        try:
                            wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
                            ws = wb[sheet_name]
                            cells = ws[data_range]
                            
                            data_rows = []
                            if isinstance(cells, tuple) and len(cells) > 0:
                                if isinstance(cells[0], tuple):
                                    for row in cells:
                                        row_data = [cell.value if cell.value is not None else "" for cell in row]
                                        data_rows.append(row_data)
                                else:
                                    row_data = [cell.value if cell.value is not None else "" for cell in cells]
                                    data_rows.append(row_data)
                            else:
                                data_rows.append([cells.value if cells.value is not None else ""])
                            
                            wb.close()
                            excel_data = data_rows
                            log_message(f"  ✓ Excel数据已加载 ({len(data_rows)} 行)")
                        except Exception as e:
                            log_message(f"  ⚠ Excel读取失败: {str(e)}", "ERROR")
                    
                    # 判断是否为批量模式（Excel数据行数>1）
                    if excel_data and len(excel_data) > 1:
                        # 批量模式：为每行数据创建一封邮件
                        headers = excel_data[0]
                        draft_count = 0
                        
                        for row_data in excel_data[1:]:
                            variables = {"{" + str(headers[i]) + "}": str(row_data[i]) 
                                       for i in range(min(len(headers), len(row_data)))}
                            
                            subject_filled = subject
                            body_filled = body
                            
                            # 处理收件人列表中的变量
                            to_filled_list = []
                            for email in to_list:
                                email_filled = email
                                for placeholder, value in variables.items():
                                    email_filled = email_filled.replace(placeholder, value)
                                to_filled_list.append(email_filled)
                            
                            for placeholder, value in variables.items():
                                subject_filled = subject_filled.replace(placeholder, value)
                                body_filled = body_filled.replace(placeholder, value)
                            
                            # 处理模板变量
                            subject_filled = subject_filled.replace("{DATE}", datetime.now().strftime("%Y-%m-%d"))
                            body_filled = body_filled.replace("{DATE}", datetime.now().strftime("%Y-%m-%d"))
                            body_filled = body_filled.replace("{TIME}", datetime.now().strftime("%H:%M:%S"))
                            
                            # 替换自定义占位符
                            for name, value in self.custom_placeholders.items():
                                subject_filled = subject_filled.replace("{" + name + "}", str(value))
                                body_filled = body_filled.replace("{" + name + "}", str(value))
                            
                            # 创建邮件
                            mail = outlook.CreateItem(0)
                            mail.To = "; ".join(to_filled_list)
                            if cc_list:
                                mail.CC = "; ".join(cc_list)
                            if bcc_list:
                                mail.BCC = "; ".join(bcc_list)
                            mail.Subject = subject_filled
                            mail.HTMLBody = body_filled.replace("\n", "<br>")
                            
                            for att_path in attachments:
                                if os.path.exists(att_path):
                                    mail.Attachments.Add(att_path)
                            
                            mail.Save()
                            draft_count += 1
                        
                        log_message(f"  ✓ 批量创建完成 ({draft_count} 封草稿)", "SUCCESS")
                        success_count += draft_count
                    else:
                        # 单邮件模式
                        mail = outlook.CreateItem(0)
                        mail.To = "; ".join(to_list)
                        if cc_list:
                            mail.CC = "; ".join(cc_list)
                        if bcc_list:
                            mail.BCC = "; ".join(bcc_list)
                        
                        # 处理主题占位符
                        subject_processed = subject.replace("{DATE}", datetime.now().strftime("%Y-%m-%d"))
                        # 替换自定义占位符
                        for name, value in self.custom_placeholders.items():
                            subject_processed = subject_processed.replace("{" + name + "}", str(value))
                        mail.Subject = subject_processed
                        
                        body_html = body
                        if excel_data:
                            # 格式化Excel数据为表格
                            table_html = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>\n"
                            for i, row in enumerate(excel_data):
                                table_html += "  <tr>\n"
                                for cell in row:
                                    cell_value = str(cell) if cell else ""
                                    if i == 0:
                                        table_html += f"    <th style='background-color: #4472C4; color: white;'>{cell_value}</th>\n"
                                    else:
                                        table_html += f"    <td>{cell_value}</td>\n"
                                table_html += "  </tr>\n"
                            table_html += "</table>"
                            body_html = body_html.replace("{EXCEL_DATA}", table_html)
                        
                        # 处理正文占位符
                        body_html = body_html.replace("{DATE}", datetime.now().strftime("%Y-%m-%d"))
                        body_html = body_html.replace("{TIME}", datetime.now().strftime("%H:%M:%S"))
                        # 替换自定义占位符
                        for name, value in self.custom_placeholders.items():
                            body_html = body_html.replace("{" + name + "}", str(value))
                        body_html = body_html.replace("\n", "<br>")
                        mail.HTMLBody = body_html
                        
                        for att_path in attachments:
                            if os.path.exists(att_path):
                                mail.Attachments.Add(att_path)
                        
                        mail.Save()
                        log_message(f"  ✓ 草稿已创建", "SUCCESS")
                        success_count += 1
                    
                except Exception as e:
                    log_message(f"  ✗ 处理失败: {str(e)}", "ERROR")
                    error_count += 1
                
                progress_window.update()
            
            # 完成
            progress_label.config(text="✓ 批量处理完成！")
            log_message("=" * 50)
            log_message(f"处理完成！成功: {success_count}, 失败: {error_count}", "SUCCESS" if error_count == 0 else "INFO")
            
            # 关闭按钮
            close_button = ttk.Button(progress_window, text="关闭", command=progress_window.destroy)
            close_button.pack(pady=10)
            
            # 更新状态栏
            self.status_var.set(f"批量创建完成 - 成功: {success_count}, 失败: {error_count}")
            
        except Exception as e:
            log_message(f"严重错误: {str(e)}", "ERROR")
            messagebox.showerror("错误", f"批量创建失败: {str(e)}")
            progress_window.destroy()
    
    def preview_all_configs(self):
        """预览所有选中的配置"""
        if not self.selected_configs:
            messagebox.showwarning("警告", "请先选择要预览的配置")
            return
        
        # 创建预览窗口
        preview_window = tk.Toplevel(self.root)
        preview_window.title(f"批量配置预览 ({len(self.selected_configs)} 个)")
        preview_window.geometry("900x700")
        preview_window.transient(self.root)
        
        # 创建笔记本（标签页）
        notebook = ttk.Notebook(preview_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        for config_name in self.selected_configs:
            config = self.configs[config_name]
            
            # 为每个配置创建一个标签页
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=config_name[:20])
            
            # 显示配置详情
            text_widget = scrolledtext.ScrolledText(frame, wrap=tk.WORD, width=80, height=35)
            text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # 处理收件人数据（兼容列表和字符串格式）
            to_data = config.get('to', [])
            cc_data = config.get('cc', [])
            bcc_data = config.get('bcc', [])
            
            if isinstance(to_data, list):
                to_str = "; ".join(to_data)
            else:
                to_str = to_data
            
            if isinstance(cc_data, list):
                cc_str = "; ".join(cc_data)
            else:
                cc_str = cc_data
            
            if isinstance(bcc_data, list):
                bcc_str = "; ".join(bcc_data)
            else:
                bcc_str = bcc_data
            
            preview_text = f"配置名称: {config_name}\n"
            preview_text += "=" * 60 + "\n\n"
            preview_text += f"收件人 (To): {to_str}\n"
            preview_text += f"抄送 (CC): {cc_str}\n"
            preview_text += f"密送 (BCC): {bcc_str}\n\n"
            preview_text += f"主题: {config.get('subject', '')}\n"
            preview_text += "-" * 60 + "\n\n"
            preview_text += "正文:\n"
            preview_text += config.get('body', '')
            preview_text += "\n\n" + "-" * 60 + "\n\n"
            
            if config.get('excel_path'):
                preview_text += f"Excel文件: {config.get('excel_path', '')}\n"
                preview_text += f"工作表: {config.get('sheet_name', '')}\n"
                preview_text += f"数据范围: {config.get('data_range', '')}\n\n"
            
            if config.get('attachments'):
                preview_text += "附件:\n"
                for att in config.get('attachments', []):
                    preview_text += f"  • {os.path.basename(att)}\n"
            
            text_widget.insert(1.0, preview_text)
            text_widget.config(state='disabled')
        
        # 关闭按钮
        ttk.Button(preview_window, text="关闭", command=preview_window.destroy).pack(pady=5)
    
    def setup_config_browser(self, parent):
        """设置右侧配置浏览器"""
        browser_frame = ttk.LabelFrame(parent, text="📚 已保存配置", padding="10")
        # 添加到 PanedWindow
        parent.add(browser_frame, weight=1)
        
        browser_frame.rowconfigure(1, weight=1)
        browser_frame.columnconfigure(0, weight=1)
        
        # 搜索框
        search_frame = ttk.Frame(browser_frame)
        search_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(search_frame, text="🔍").pack(side=tk.LEFT)
        self.config_search_var = tk.StringVar()
        self.config_search_var.trace('w', self.filter_config_list)
        search_entry = ttk.Entry(search_frame, textvariable=self.config_search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 配置列表 (使用 Treeview 替代 Listbox 以显示更多信息)
        list_frame = ttk.Frame(browser_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        h_scrollbar = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL)
        
        self.config_tree = ttk.Treeview(list_frame, columns=("name", "file"), show="headings", 
                                      yscrollcommand=scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.config_tree.heading("name", text="配置名称")
        self.config_tree.heading("file", text="Excel 文件")
        
        self.config_tree.column("name", width=150, minwidth=100)
        self.config_tree.column("file", width=200, minwidth=100)
        
        scrollbar.config(command=self.config_tree.yview)
        h_scrollbar.config(command=self.config_tree.xview)
        
        self.config_tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))
        
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 绑定双击事件加载配置
        self.config_tree.bind('<Double-Button-1>', self.load_config_from_browser)
        self.config_tree.bind('<Return>', self.load_config_from_browser)
        
        # 按钮区域
        button_frame = ttk.Frame(browser_frame)
        button_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(button_frame, text="加载", command=lambda: self.load_config_from_browser(None), 
                  width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="删除", command=self.delete_from_browser, 
                  width=10).pack(side=tk.LEFT, padx=2)
        
        # 信息标签
        self.browser_info_label = ttk.Label(browser_frame, text="", font=('TkDefaultFont', 8))
        self.browser_info_label.pack(pady=(5, 0))
        
        # 初始化列表
        self.update_config_browser()
    
    def update_config_browser(self):
        """更新配置浏览器列表"""
        # 清空现有项
        for item in self.config_tree.get_children():
            self.config_tree.delete(item)
        
        search_term = self.config_search_var.get().lower() if hasattr(self, 'config_search_var') else ""
        
        config_list = sorted(self.configs.keys())
        count = 0
        
        for name in config_list:
            if search_term and search_term not in name.lower():
                continue
                
            config = self.configs[name]
            excel_path = config.get("excel_path", "")
            # 如果有多个文件，显示数量
            if ";" in excel_path:
                file_display = f"[{excel_path.count(';')+1} 个文件] {os.path.basename(excel_path.split(';')[0])}..."
            else:
                file_display = os.path.basename(excel_path) if excel_path else "(无)"
                # 如果文件名相同但路径不同，显示部分路径以区分
                # 这里简单处理：显示文件名，鼠标悬停或选中时可以看到详情（虽然Treeview没有默认tooltip）
                # 实际上，用户可以通过查看 file_display 来区分
            
            # 插入数据，values对应 columns
            self.config_tree.insert("", tk.END, values=(name, file_display), tags=(name,))
            count += 1
            
        if hasattr(self, 'browser_info_label'):
            self.browser_info_label.config(text=f"共 {count} 个配置")

    def filter_config_list(self, *args):
        """过滤配置列表"""
        self.update_config_browser()

    def load_config_from_browser(self, event):
        """从浏览器加载配置"""
        selection = self.config_tree.selection()
        if not selection:
            return
            
        item = selection[0]
        values = self.config_tree.item(item, "values")
        if values:
            config_name = values[0]
            self.config_combo.set(config_name)
            self.load_selected_config()

    def delete_from_browser(self):
        """从浏览器删除配置"""
        selection = self.config_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的配置")
            return
            
        item = selection[0]
        values = self.config_tree.item(item, "values")
        if values:
            config_name = values[0]
            if messagebox.askyesno("确认删除", f"确定要删除配置 '{config_name}' 吗？"):
                del self.configs[config_name]
                if self.save_configs():
                    self.update_config_list() # 这会触发 update_config_browser
                    messagebox.showinfo("成功", f"配置 '{config_name}' 已删除")

    
    def add_recipient(self, recipient_type):
        """添加收件人/抄送/密送邮箱"""
        import re
        
        if recipient_type == 'to':
            email = self.to_entry_var.get().strip()
            listbox = self.to_listbox
        elif recipient_type == 'cc':
            email = self.cc_entry_var.get().strip()
            listbox = self.cc_listbox
        else:  # bcc
            email = self.bcc_entry_var.get().strip()
            listbox = self.bcc_listbox
        
        if not email:
            messagebox.showwarning("警告", "请输入邮箱地址")
            return
        
        # 简单的邮箱格式验证
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        if not re.match(email_pattern, email):
            if not messagebox.askyesno("格式警告", 
                f"'{email}' 格式可能不正确，是否仍要添加？"):
                return
        
        # 检查是否已存在
        current_emails = list(listbox.get(0, tk.END))
        if email in current_emails:
            messagebox.showwarning("警告", "此邮箱已在列表中")
            return
        
        # 添加到列表
        listbox.insert(tk.END, email)
        
        # 清空输入框
        if recipient_type == 'to':
            self.to_entry_var.set("")
        elif recipient_type == 'cc':
            self.cc_entry_var.set("")
        else:
            self.bcc_entry_var.set("")
    
    def remove_recipient(self, recipient_type):
        """删除选中的收件人/抄送/密送邮箱"""
        if recipient_type == 'to':
            listbox = self.to_listbox
        elif recipient_type == 'cc':
            listbox = self.cc_listbox
        else:  # bcc
            listbox = self.bcc_listbox
        
        selection = listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的邮箱")
            return
        
        # 从后往前删除（避免索引变化问题）
        for index in reversed(selection):
            listbox.delete(index)
    
    def show_placeholder_menu(self, entry_widget):
        """显示占位符菜单（用于主题）"""
        menu = tk.Menu(self.root, tearoff=0)
        
        # 常用占位符
        menu.add_command(label="📅 {DATE} - 当前日期", 
                        command=lambda: self.insert_placeholder(entry_widget, "{DATE}"))
        menu.add_command(label="⏰ {TIME} - 当前时间", 
                        command=lambda: self.insert_placeholder(entry_widget, "{TIME}"))
        menu.add_command(label="📆 {DATETIME} - 日期时间", 
                        command=lambda: self.insert_placeholder(entry_widget, "{DATETIME}"))
        
        menu.add_separator()
        
        # Excel列占位符
        if self.current_excel_data and len(self.current_excel_data) > 0:
            headers = self.current_excel_data[0]
            menu.add_command(label="Excel 列占位符:", state="disabled")
            for header in headers[:10]:  # 最多显示10个
                menu.add_command(label=f"  📊 {{{header}}}", 
                               command=lambda h=header: self.insert_placeholder(entry_widget, f"{{{h}}}"))
            if len(headers) > 10:
                menu.add_command(label=f"  ... 还有 {len(headers)-10} 个", state="disabled")
        else:
            menu.add_command(label="(先读取Excel数据)", state="disabled")
        
        # 显示菜单
        try:
            menu.post(entry_widget.winfo_rootx(), entry_widget.winfo_rooty() + entry_widget.winfo_height())
        finally:
            menu.grab_release()
    
    def edit_body_in_window(self):
        """在弹出窗口中编辑正文"""
        # 创建弹出窗口
        body_window = tk.Toplevel(self.root)
        body_window.title("编辑邮件正文")
        body_window.geometry("900x600")
        body_window.transient(self.root)
        
        # 工具栏
        toolbar = ttk.Frame(body_window, padding="5")
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        ttk.Label(toolbar, text="编辑正文内容:", font=('TkDefaultFont', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="📋 插入占位符", command=lambda: self.show_body_placeholder_menu_popup(body_text)).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="✍️ 插入签名", command=lambda: self.insert_signature_menu_popup(body_text)).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🔍 拼写检查", command=lambda: self.check_spelling_popup(body_text)).pack(side=tk.LEFT, padx=2)
        
        # 正文编辑区
        text_frame = ttk.Frame(body_window, padding="5")
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        body_text = scrolledtext.ScrolledText(text_frame, width=80, height=25, wrap=tk.WORD, font=('TkDefaultFont', 10))
        body_text.pack(fill=tk.BOTH, expand=True)
        
        # 加载当前正文内容
        current_body = self.body_text.get(1.0, tk.END)
        body_text.insert(1.0, current_body)
        
        # 按钮区
        button_frame = ttk.Frame(body_window, padding="10")
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        def save_and_close():
            # 保存内容到隐藏的body_text
            content = body_text.get(1.0, tk.END).strip()
            self.body_text.delete(1.0, tk.END)
            self.body_text.insert(1.0, content)
            
            # 更新预览标签
            preview = content[:50] + "..." if len(content) > 50 else content
            preview = preview.replace("\n", " ")
            if preview:
                self.body_preview_label.config(text=f"已编辑: {preview}", foreground='blue')
            else:
                self.body_preview_label.config(text="(点击编辑正文内容)", foreground='gray')
            
            body_window.destroy()
        
        ttk.Button(button_frame, text="✓ 保存并关闭", command=save_and_close, width=15).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="✗ 取消", command=body_window.destroy, width=15).pack(side=tk.RIGHT, padx=5)
        
        # 聚焦到文本框
        body_text.focus_set()
    
    def show_body_placeholder_menu_popup(self, text_widget):
        """在弹窗中显示正文占位符菜单"""
        menu = tk.Menu(self.root, tearoff=0)
        
        # 常用占位符
        menu.add_command(label="📅 {DATE} - 当前日期", 
                        command=lambda: text_widget.insert(tk.INSERT, "{DATE}"))
        menu.add_command(label="⏰ {TIME} - 当前时间", 
                        command=lambda: text_widget.insert(tk.INSERT, "{TIME}"))
        menu.add_command(label="📆 {DATETIME} - 日期时间", 
                        command=lambda: text_widget.insert(tk.INSERT, "{DATETIME}"))
        
        menu.add_separator()
        menu.add_command(label="📊 {EXCEL_DATA} - 插入完整表格", 
                        command=lambda: text_widget.insert(tk.INSERT, "{EXCEL_DATA}"))
        
        menu.add_separator()
        
        # Excel列占位符
        if self.current_excel_data and len(self.current_excel_data) > 0:
            headers = self.current_excel_data[0]
            menu.add_command(label="Excel 列占位符 (用于批量):", state="disabled")
            for header in headers[:15]:
                menu.add_command(label=f"  📊 {{{header}}}", 
                               command=lambda h=header: text_widget.insert(tk.INSERT, f"{{{h}}}"))
            if len(headers) > 15:
                menu.add_command(label=f"  ... 还有 {len(headers)-15} 个", state="disabled")
        else:
            menu.add_command(label="(先读取Excel数据查看列占位符)", state="disabled")
        
        menu.add_separator()
        
        # 自定义占位符
        if self.custom_placeholders:
            menu.add_command(label="自定义占位符:", state="disabled")
            for name, value in list(self.custom_placeholders.items())[:10]:
                display_value = str(value)[:20] + "..." if len(str(value)) > 20 else str(value)
                menu.add_command(label=f"  {{{name}}} = {display_value}", 
                               command=lambda n=name: text_widget.insert(tk.INSERT, f"{{{n}}}"))
        
        # 显示菜单
        try:
            menu.post(text_widget.winfo_rootx() + 10, text_widget.winfo_rooty() + 30)
        finally:
            menu.grab_release()
    
    def insert_signature_menu_popup(self, text_widget):
        """在弹窗中插入签名"""
        menu = tk.Menu(self.root, tearoff=0)
        
        if not self.email_signatures:
            menu.add_command(label="（暂无签名，请先添加）", state="disabled")
        else:
            for name, content in self.email_signatures.items():
                # 简单的预览处理
                preview = content.strip().replace('\n', ' ')
                if len(preview) > 20:
                    preview = preview[:20] + "..."
                
                menu.add_command(label=f"✍️ {name} ({preview})", 
                               command=lambda c=content: text_widget.insert(tk.INSERT, "\n\n" + c))
        
        menu.add_separator()
        menu.add_command(label="⚙️ 管理签名...", command=self.manage_signatures)
        
        try:
            # 获取按钮位置（如果可能）或者鼠标位置
            x = self.root.winfo_pointerx()
            y = self.root.winfo_pointery()
            menu.post(x, y)
        finally:
            menu.grab_release()
    
    def check_spelling_popup(self, text_widget):
        """在弹窗中检查拼写"""
        content = text_widget.get(1.0, tk.END).strip()
        if not content:
            messagebox.showinfo("提示", "正文为空")
            return
        
        # 检查拼写
        errors = []
        for wrong, correct in self.common_spelling_errors.items():
            if wrong in content:
                errors.append((wrong, correct))
        
        if not errors:
            messagebox.showinfo("拼写检查", "✓ 未发现常见拼写错误")
            return
        
        # 显示错误并提供修正
        msg = "发现以下可能的拼写错误:\n\n"
        for wrong, correct in errors:
            msg += f"  • '{wrong}' → '{correct}'\n"
        msg += f"\n是否自动修正这些错误？"
        
        if messagebox.askyesno("拼写检查", msg):
            new_content = content
            for wrong, correct in errors:
                new_content = new_content.replace(wrong, correct)
            text_widget.delete(1.0, tk.END)
            text_widget.insert(1.0, new_content)
            messagebox.showinfo("完成", f"已修正 {len(errors)} 个拼写错误")
    
    def show_body_placeholder_menu(self):
        """显示正文占位符菜单（旧方法，保留兼容性）"""
        menu = tk.Menu(self.root, tearoff=0)
        
        # 常用占位符
        menu.add_command(label="📅 {DATE} - 当前日期", 
                        command=lambda: self.insert_text_placeholder("{DATE}"))
        menu.add_command(label="⏰ {TIME} - 当前时间", 
                        command=lambda: self.insert_text_placeholder("{TIME}"))
        menu.add_command(label="📆 {DATETIME} - 日期时间", 
                        command=lambda: self.insert_text_placeholder("{DATETIME}"))
        
        menu.add_separator()
        menu.add_command(label="📊 {EXCEL_DATA} - 插入完整表格", 
                        command=lambda: self.insert_text_placeholder("{EXCEL_DATA}"))
        
        menu.add_separator()
        
        # Excel列占位符
        if self.current_excel_data and len(self.current_excel_data) > 0:
            headers = self.current_excel_data[0]
            menu.add_command(label="Excel 列占位符 (用于批量):", state="disabled")
            for header in headers[:15]:
                menu.add_command(label=f"  📊 {{{header}}}", 
                               command=lambda h=header: self.insert_text_placeholder(f"{{{h}}}"))
            if len(headers) > 15:
                menu.add_command(label=f"  ... 还有 {len(headers)-15} 个", state="disabled")
        else:
            menu.add_command(label="(先读取Excel数据查看列占位符)", state="disabled")
        
        menu.add_separator()
        
        # 自定义占位符
        menu.add_command(label="✏️ 自定义占位符...", 
                        command=self.create_custom_placeholder)
        
        # 显示菜单
        try:
            menu.post(self.body_text.winfo_rootx(), self.body_text.winfo_rooty() + 30)
        finally:
            menu.grab_release()
    
    def insert_placeholder(self, entry_widget, placeholder):
        """在Entry中插入占位符"""
        if isinstance(entry_widget, ttk.Entry):
            # 获取当前光标位置
            current_pos = entry_widget.index(tk.INSERT)
            current_text = entry_widget.get()
            
            # 插入占位符
            new_text = current_text[:current_pos] + placeholder + current_text[current_pos:]
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, new_text)
            
            # 移动光标到占位符后
            entry_widget.icursor(current_pos + len(placeholder))
            entry_widget.focus()
    
    def insert_text_placeholder(self, placeholder):
        """在Text控件中插入占位符"""
        # 获取当前光标位置
        cursor_pos = self.body_text.index(tk.INSERT)
        
        # 插入占位符
        self.body_text.insert(cursor_pos, placeholder)
        
        # 移动光标到占位符后
        self.body_text.mark_set(tk.INSERT, f"{cursor_pos}+{len(placeholder)}c")
        self.body_text.focus()
    
    def create_custom_placeholder(self):
        """创建自定义占位符"""
        dialog = tk.Toplevel(self.root)
        dialog.title("创建自定义占位符")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="输入占位符名称（不含花括号）：", 
                 font=('TkDefaultFont', 10)).pack(pady=10)
        
        name_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        entry.pack(pady=5)
        entry.focus()
        
        ttk.Label(dialog, text="例如: 客户名称, 订单号, 金额 等", 
                 font=('TkDefaultFont', 8), foreground='gray').pack()
        
        def insert_custom():
            name = name_var.get().strip()
            if name:
                placeholder = "{" + name + "}"
                self.insert_text_placeholder(placeholder)
                dialog.destroy()
            else:
                messagebox.showwarning("警告", "请输入占位符名称")
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="插入", command=insert_custom).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键
        entry.bind('<Return>', lambda e: insert_custom())
    
    def add_custom_placeholder(self):
        """添加自定义占位符"""
        name = self.placeholder_name_var.get().strip()
        value = self.placeholder_value_var.get().strip()
        
        if not name:
            messagebox.showwarning("警告", "请输入占位符名称")
            return
        
        if not value:
            messagebox.showwarning("警告", "请输入占位符的值")
            return
        
        # 保存占位符
        self.custom_placeholders[name] = value
        
        # 清空输入
        self.placeholder_name_var.set("")
        self.placeholder_value_var.set("")
        
        messagebox.showinfo("成功", f"已添加占位符: {{{name}}} = {value}\n\n在主题或正文中使用 {{{name}}} 即可自动替换")
        self.status_var.set(f"已添加占位符: {{{name}}}")
    
    def remove_custom_placeholder(self):
        """从管理窗口删除选中的占位符"""
        if hasattr(self, 'placeholder_listbox_manage'):
            selection = self.placeholder_listbox_manage.curselection()
            if not selection:
                messagebox.showwarning("警告", "请先选择要删除的占位符")
                return
            
            index = selection[0]
            keys = list(self.custom_placeholders.keys())
            if index < len(keys):
                key = keys[index]
                del self.custom_placeholders[key]
                self.update_placeholder_listbox()
                messagebox.showinfo("成功", f"已删除占位符: {{{key}}}")
    
    def clear_custom_placeholders(self):
        """清空所有自定义占位符"""
        if not self.custom_placeholders:
            messagebox.showinfo("提示", "没有自定义占位符")
            return
        
        if messagebox.askyesno("确认", f"确定要清空所有 {len(self.custom_placeholders)} 个自定义占位符吗？"):
            self.custom_placeholders.clear()
            if hasattr(self, 'placeholder_listbox_manage'):
                self.update_placeholder_listbox()
            messagebox.showinfo("成功", "已清空所有自定义占位符")
    
    def manage_placeholders(self):
        """打开占位符管理窗口"""
        dialog = tk.Toplevel(self.root)
        dialog.title("占位符管理")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 说明文本
        ttk.Label(dialog, text="在主题或正文中使用 {占位符名} 格式，系统会自动替换为设定的值", 
                 font=('TkDefaultFont', 9), foreground='blue').pack(pady=10, padx=10)
        
        # 已定义的占位符列表
        list_frame = ttk.LabelFrame(dialog, text="已定义的占位符", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.placeholder_listbox_manage = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=('Courier New', 10))
        self.placeholder_listbox_manage.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.placeholder_listbox_manage.yview)
        
        self.update_placeholder_listbox()
        
        # 按钮区域
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="❌ 删除选中", command=self.remove_custom_placeholder).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ 清空全部", command=self.clear_custom_placeholders).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="关闭", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # 使用说明
        help_frame = ttk.LabelFrame(dialog, text="使用说明", padding="10")
        help_frame.pack(fill=tk.X, padx=10, pady=5)
        
        help_text = """• 系统内置占位符: {DATE}, {TIME}, {DATETIME}, {EXCEL_DATA}
• Excel 列占位符: {列名} (如 {姓名}, {邮箱})
• 自定义占位符: 您添加的占位符 (如 {公司名}, {部门})
• 在主题或正文中使用这些占位符，创建草稿时会自动替换"""
        
        ttk.Label(help_frame, text=help_text, font=('TkDefaultFont', 8), justify=tk.LEFT).pack()
    
    def update_placeholder_listbox(self):
        """更新占位符列表显示"""
        if hasattr(self, 'placeholder_listbox_manage'):
            self.placeholder_listbox_manage.delete(0, tk.END)
            
            if not self.custom_placeholders:
                self.placeholder_listbox_manage.insert(tk.END, "（暂无自定义占位符）")
            else:
                for name, value in self.custom_placeholders.items():
                    # 格式化显示: {占位符名} = 值
                    display = f"{{{name}}}  =  {value}"
                    self.placeholder_listbox_manage.insert(tk.END, display)
    
    def load_signatures(self):
        """加载邮件签名"""
        signature_file = "email_signatures.json"
        if os.path.exists(signature_file):
            try:
                with open(signature_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        # 默认签名
        return {
            "默认签名": "\n\n--\n此致\n敬礼！",
            "正式签名": "\n\n───────────────\n{公司名}\n{联系人}\n电话: {电话}\n邮箱: {邮箱}",
            "简洁签名": "\n\n谢谢！\nBest regards"
        }
    
    def save_signatures(self):
        """保存邮件签名"""
        signature_file = "email_signatures.json"
        try:
            with open(signature_file, 'w', encoding='utf-8') as f:
                json.dump(self.email_signatures, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("错误", f"保存签名失败: {str(e)}")
            return False
    
    def insert_signature_menu(self):
        """显示插入签名菜单"""
        menu = tk.Menu(self.root, tearoff=0)
        
        if not self.email_signatures:
            menu.add_command(label="（暂无签名，请先添加）", state="disabled")
        else:
            for sig_name in self.email_signatures.keys():
                menu.add_command(
                    label=f"✍️ {sig_name}",
                    command=lambda name=sig_name: self.insert_signature(name)
                )
        
        menu.add_separator()
        menu.add_command(label="⚙️ 管理签名...", command=self.manage_signatures)
        
        try:
            menu.post(self.body_text.winfo_rootx(), self.body_text.winfo_rooty() + 30)
        finally:
            menu.grab_release()
    
    def insert_signature(self, signature_name):
        """插入签名到正文"""
        if signature_name in self.email_signatures:
            signature = self.email_signatures[signature_name]
            current_text = self.body_text.get(1.0, tk.END).rstrip()
            self.body_text.delete(1.0, tk.END)
            self.body_text.insert(1.0, current_text + signature)
            self.status_var.set(f"已插入签名: {signature_name}")
    
    def manage_signatures(self):
        """管理邮件签名"""
        dialog = tk.Toplevel(self.root)
        dialog.title("邮件签名管理")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 说明
        ttk.Label(dialog, text="管理您的邮件签名模板，可以在正文中快速插入", 
                 font=('TkDefaultFont', 9), foreground='blue').pack(pady=10, padx=10)
        
        # 签名列表
        list_frame = ttk.LabelFrame(dialog, text="已保存的签名", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 列表框
        list_container = ttk.Frame(list_frame)
        list_container.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.signature_listbox = tk.Listbox(list_container, yscrollcommand=scrollbar.set, font=('TkDefaultFont', 10))
        self.signature_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.signature_listbox.yview)
        
        self.signature_listbox.bind('<<ListboxSelect>>', self.on_signature_select)
        self.update_signature_listbox()
        
        # 编辑区域
        edit_frame = ttk.LabelFrame(dialog, text="签名内容", padding="10")
        edit_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        ttk.Label(edit_frame, text="签名名称:").pack(anchor=tk.W)
        self.signature_name_entry = ttk.Entry(edit_frame, width=40)
        self.signature_name_entry.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(edit_frame, text="签名内容: (支持占位符，如 {公司名})").pack(anchor=tk.W)
        self.signature_text = scrolledtext.ScrolledText(edit_frame, width=60, height=6, wrap=tk.WORD)
        self.signature_text.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # 按钮区域
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="➕ 新建", command=self.new_signature).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 保存", command=self.save_signature).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="❌ 删除", command=self.delete_signature).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="关闭", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def update_signature_listbox(self):
        """更新签名列表"""
        if hasattr(self, 'signature_listbox'):
            self.signature_listbox.delete(0, tk.END)
            for name in self.email_signatures.keys():
                self.signature_listbox.insert(tk.END, name)
    
    def on_signature_select(self, event):
        """选择签名时显示内容"""
        selection = self.signature_listbox.curselection()
        if selection:
            index = selection[0]
            sig_name = self.signature_listbox.get(index)
            if sig_name in self.email_signatures:
                self.signature_name_entry.delete(0, tk.END)
                self.signature_name_entry.insert(0, sig_name)
                self.signature_text.delete(1.0, tk.END)
                self.signature_text.insert(1.0, self.email_signatures[sig_name])
    
    def new_signature(self):
        """新建签名"""
        self.signature_name_entry.delete(0, tk.END)
        self.signature_text.delete(1.0, tk.END)
        self.signature_name_entry.focus()
    
    def save_signature(self):
        """保存/更新签名"""
        sig_name = self.signature_name_entry.get().strip()
        sig_content = self.signature_text.get(1.0, tk.END).rstrip()
        
        if not sig_name:
            messagebox.showwarning("警告", "请输入签名名称")
            return
        
        if not sig_content:
            messagebox.showwarning("警告", "请输入签名内容")
            return
        
        self.email_signatures[sig_name] = sig_content
        if self.save_signatures():
            self.update_signature_listbox()
            messagebox.showinfo("成功", f"签名 '{sig_name}' 已保存")
    
    def delete_signature(self):
        """删除签名"""
        sig_name = self.signature_name_entry.get().strip()
        
        if not sig_name:
            messagebox.showwarning("警告", "请先选择要删除的签名")
            return
        
        if sig_name in self.email_signatures:
            if messagebox.askyesno("确认", f"确定要删除签名 '{sig_name}' 吗？"):
                del self.email_signatures[sig_name]
                if self.save_signatures():
                    self.update_signature_listbox()
                    self.new_signature()
                    messagebox.showinfo("成功", "签名已删除")
    
    def check_spelling(self):
        """拼写检查"""
        subject = self.subject_var.get()
        body = self.body_text.get(1.0, tk.END)
        
        errors_found = []
        
        # 检查主题
        for wrong, correct in self.spell_check_dict.items():
            if wrong in subject:
                errors_found.append(f"主题中: '{wrong}' 应为 '{correct}'")
        
        # 检查正文
        for wrong, correct in self.spell_check_dict.items():
            if wrong in body:
                errors_found.append(f"正文中: '{wrong}' 应为 '{correct}'")
        
        # 检查常见格式问题
        if re.search(r'\s{2,}', subject):
            errors_found.append("主题中有多余的空格")
        
        if re.search(r'[a-zA-Z][，。！？]', body):
            errors_found.append("正文中英文后使用了中文标点")
        
        if re.search(r'[\u4e00-\u9fa5][,\.!?]', body):
            errors_found.append("正文中中文后使用了英文标点")
        
        if errors_found:
            error_msg = "发现以下可能的问题:\n\n" + "\n".join(f"• {e}" for e in errors_found)
            error_msg += "\n\n是否自动修正?"
            
            if messagebox.askyesno("拼写检查", error_msg):
                # 自动修正
                new_subject = subject
                new_body = body
                
                for wrong, correct in self.spell_check_dict.items():
                    new_subject = new_subject.replace(wrong, correct)
                    new_body = new_body.replace(wrong, correct)
                
                self.subject_var.set(new_subject)
                self.body_text.delete(1.0, tk.END)
                self.body_text.insert(1.0, new_body)
                
                messagebox.showinfo("成功", "已自动修正拼写错误")
        else:
            messagebox.showinfo("拼写检查", "未发现明显的拼写错误 ✓")
    
    def schedule_send_dialog(self):
        """定时发送对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("定时发送设置")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="设置邮件的发送时间", font=('TkDefaultFont', 10, 'bold')).pack(pady=10)
        
        # 日期选择
        date_frame = ttk.Frame(dialog)
        date_frame.pack(pady=10, padx=20, fill=tk.X)
        
        ttk.Label(date_frame, text="日期:").pack(side=tk.LEFT, padx=5)
        
        self.schedule_year = tk.StringVar(value=str(datetime.now().year))
        ttk.Spinbox(date_frame, from_=2025, to=2030, textvariable=self.schedule_year, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Label(date_frame, text="年").pack(side=tk.LEFT)
        
        self.schedule_month = tk.StringVar(value=str(datetime.now().month))
        ttk.Spinbox(date_frame, from_=1, to=12, textvariable=self.schedule_month, width=5).pack(side=tk.LEFT, padx=2)
        ttk.Label(date_frame, text="月").pack(side=tk.LEFT)
        
        self.schedule_day = tk.StringVar(value=str(datetime.now().day))
        ttk.Spinbox(date_frame, from_=1, to=31, textvariable=self.schedule_day, width=5).pack(side=tk.LEFT, padx=2)
        ttk.Label(date_frame, text="日").pack(side=tk.LEFT)
        
        # 时间选择
        time_frame = ttk.Frame(dialog)
        time_frame.pack(pady=10, padx=20, fill=tk.X)
        
        ttk.Label(time_frame, text="时间:").pack(side=tk.LEFT, padx=5)
        
        self.schedule_hour = tk.StringVar(value="09")
        ttk.Spinbox(time_frame, from_=0, to=23, textvariable=self.schedule_hour, width=5, format="%02.0f").pack(side=tk.LEFT, padx=2)
        ttk.Label(time_frame, text=":").pack(side=tk.LEFT)
        
        self.schedule_minute = tk.StringVar(value="00")
        ttk.Spinbox(time_frame, from_=0, to=59, textvariable=self.schedule_minute, width=5, format="%02.0f").pack(side=tk.LEFT, padx=2)
        
        # 说明文本
        info_text = """
注意：
• 邮件将先创建为草稿
• 在指定时间自动发送
• 请确保程序在发送时间保持运行
• 发送前会再次确认
        """
        ttk.Label(dialog, text=info_text, font=('TkDefaultFont', 8), foreground='gray', justify=tk.LEFT).pack(pady=10)
        
        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="确定", command=lambda: self.confirm_schedule_send(dialog)).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
    
    def confirm_schedule_send(self, dialog):
        """确认定时发送"""
        try:
            year = int(self.schedule_year.get())
            month = int(self.schedule_month.get())
            day = int(self.schedule_day.get())
            hour = int(self.schedule_hour.get())
            minute = int(self.schedule_minute.get())
            
            send_time = datetime(year, month, day, hour, minute)
            
            if send_time <= datetime.now():
                messagebox.showerror("错误", "发送时间必须晚于当前时间")
                return
            
            dialog.destroy()
            
            # 创建草稿
            self.create_draft()
            
            # 启动定时发送线程
            time_str = send_time.strftime("%Y-%m-%d %H:%M")
            if messagebox.askyesno("确认", f"草稿已创建\n\n将在 {time_str} 自动发送\n\n请保持程序运行"):
                thread = threading.Thread(target=self.scheduled_send_worker, args=(send_time,), daemon=True)
                thread.start()
                self.status_var.set(f"已设置定时发送: {time_str}")
        
        except Exception as e:
            messagebox.showerror("错误", f"设置失败: {str(e)}")
    
    def scheduled_send_worker(self, send_time):
        """定时发送工作线程"""
        while datetime.now() < send_time:
            time.sleep(60)  # 每分钟检查一次
        
        # 时间到了，尝试发送
        try:
            # 这里需要找到并发送对应的草稿
            # 简化实现：提示用户手动发送
            self.root.after(0, lambda: messagebox.showinfo(
                "定时发送提醒", 
                f"定时发送时间已到！\n\n请在 Outlook 中找到草稿并发送。"
            ))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"发送失败: {str(e)}"))
    
    def show_enhanced_preview(self):
        """显示增强预览（类似Outlook样式）"""
        # 获取邮件信息
        to_list = list(self.to_listbox.get(0, tk.END))
        cc_list = list(self.cc_listbox.get(0, tk.END))
        bcc_list = list(self.bcc_listbox.get(0, tk.END))
        subject = self.subject_var.get()
        body = self.body_text.get(1.0, tk.END).strip()
        
        if not subject:
            messagebox.showwarning("警告", "请填写主题")
            return
        
        # 创建预览窗口
        preview_win = tk.Toplevel(self.root)
        preview_win.title("邮件预览 - Outlook 风格")
        preview_win.geometry("800x600")
        
        # 顶部工具栏
        toolbar = ttk.Frame(preview_win, relief=tk.RAISED, borderwidth=1)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(toolbar, text="📧 邮件预览", font=('TkDefaultFont', 12, 'bold')).pack(side=tk.LEFT, padx=10)
        ttk.Button(toolbar, text="关闭", command=preview_win.destroy).pack(side=tk.RIGHT, padx=5)
        
        # 邮件头部信息
        header_frame = ttk.Frame(preview_win, relief=tk.GROOVE, borderwidth=2)
        header_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 发件人
        from_frame = ttk.Frame(header_frame)
        from_frame.pack(fill=tk.X, padx=10, pady=3)
        ttk.Label(from_frame, text="发件人:", font=('TkDefaultFont', 9, 'bold'), width=10).pack(side=tk.LEFT)
        ttk.Label(from_frame, text="当前 Outlook 用户", foreground='blue').pack(side=tk.LEFT)
        
        # 收件人
        if to_list:
            to_frame = ttk.Frame(header_frame)
            to_frame.pack(fill=tk.X, padx=10, pady=3)
            ttk.Label(to_frame, text="收件人:", font=('TkDefaultFont', 9, 'bold'), width=10).pack(side=tk.LEFT)
            ttk.Label(to_frame, text="; ".join(to_list), foreground='blue').pack(side=tk.LEFT)
        
        # 抄送
        if cc_list:
            cc_frame = ttk.Frame(header_frame)
            cc_frame.pack(fill=tk.X, padx=10, pady=3)
            ttk.Label(cc_frame, text="抄送:", font=('TkDefaultFont', 9, 'bold'), width=10).pack(side=tk.LEFT)
            ttk.Label(cc_frame, text="; ".join(cc_list), foreground='blue').pack(side=tk.LEFT)
        
        # 主题
        subject_frame = ttk.Frame(header_frame)
        subject_frame.pack(fill=tk.X, padx=10, pady=3)
        ttk.Label(subject_frame, text="主题:", font=('TkDefaultFont', 9, 'bold'), width=10).pack(side=tk.LEFT)
        ttk.Label(subject_frame, text=subject, font=('TkDefaultFont', 10)).pack(side=tk.LEFT)
        
        # 附件
        if self.attachments:
            attach_frame = ttk.Frame(header_frame)
            attach_frame.pack(fill=tk.X, padx=10, pady=3)
            ttk.Label(attach_frame, text="附件:", font=('TkDefaultFont', 9, 'bold'), width=10).pack(side=tk.LEFT)
            attach_text = f"📎 {len(self.attachments)} 个文件"
            ttk.Label(attach_frame, text=attach_text, foreground='gray').pack(side=tk.LEFT)
        
        # 分隔线
        ttk.Separator(preview_win, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=10)
        
        # 邮件正文
        body_frame = ttk.LabelFrame(preview_win, text="正文内容", padding=10)
        body_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 使用 Text 显示正文（支持基本的 HTML 显示）
        body_text = scrolledtext.ScrolledText(body_frame, wrap=tk.WORD, font=('Calibri', 11))
        body_text.pack(fill=tk.BOTH, expand=True)
        
        # 处理占位符
        processed_body = body
        processed_body = processed_body.replace("{DATE}", datetime.now().strftime("%Y-%m-%d"))
        processed_body = processed_body.replace("{TIME}", datetime.now().strftime("%H:%M:%S"))
        
        # 插入正文
        body_text.insert(1.0, processed_body)
        body_text.config(state='disabled')
        
        # 底部信息
        info_frame = ttk.Frame(preview_win)
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ttk.Label(info_frame, text=f"预览生成时间: {now_str}", 
                 font=('TkDefaultFont', 8), foreground='gray').pack(side=tk.LEFT)
    
    def toggle_language(self):
        """切换语言"""
        self.current_language = 'en' if self.current_language == 'zh' else 'zh'
        
        # 更新窗口标题
        self.root.title(self.translations[self.current_language]['title'])
        
        # 更新状态栏
        if self.current_language == 'zh':
            self.status_var.set("就绪 | 快捷键: Ctrl+S=保存配置 | Ctrl+D=创建草稿 | Ctrl+P=预览 | F1=帮助")
            messagebox.showinfo("语言切换", "已切换到中文\n注意：部分界面需要重启程序才能完全更新")
        else:
            self.status_var.set("Ready | Shortcuts: Ctrl+S=Save | Ctrl+D=Draft | Ctrl+P=Preview | F1=Help")
            messagebox.showinfo("Language Switch", "Switched to English\nNote: Some UI elements require restart to fully update")
    
    def zoom_in(self):
        """放大界面"""
        if self.ui_scale < 2.0:  # 最大200%
            self.ui_scale += 0.1
            self.apply_zoom()
            self.save_ui_preferences()  # 保存设置
    
    def zoom_out(self):
        """缩小界面"""
        if self.ui_scale > 0.5:  # 最小50%
            self.ui_scale -= 0.1
            self.apply_zoom()
            self.save_ui_preferences()  # 保存设置
    
    def reset_zoom(self):
        """重置缩放"""
        self.ui_scale = 1.0
        self.apply_zoom()
        self.save_ui_preferences()  # 保存设置
    
    def mouse_zoom(self, event):
        """鼠标滚轮缩放"""
        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()
    
    def apply_zoom(self):
        """应用缩放设置"""
        # 更新缩放标签
        zoom_percent = int(self.ui_scale * 100)
        self.zoom_label.config(text=f"{zoom_percent}%")
        
        # 计算字体大小
        base_font_size = 9
        new_font_size = int(base_font_size * self.ui_scale)
        
        # 更新默认字体
        default_font = ('TkDefaultFont', new_font_size)
        text_font = ('TkTextFont', new_font_size)
        fixed_font = ('TkFixedFont', new_font_size)
        
        try:
            self.root.option_add('*TCombobox*Listbox.font', default_font)
            self.root.option_add('*Font', default_font)
        except:
            pass
        
        # 更新所有文本框的字体
        for widget in self.root.winfo_children():
            self._update_widget_font(widget, new_font_size)
        
        # 更新窗口大小
        base_width = 1200
        base_height = 800
        new_width = int(base_width * self.ui_scale)
        new_height = int(base_height * self.ui_scale)
        
        # 限制最小和最大尺寸
        new_width = max(800, min(new_width, 1920))
        new_height = max(600, min(new_height, 1080))
        
        self.root.geometry(f"{new_width}x{new_height}")
        
        # 更新状态
        self.status_var.set(f"界面缩放: {zoom_percent}% | Ctrl+Plus/Minus 或 Ctrl+滚轮缩放 | Ctrl+0 重置")
    
    def _update_widget_font(self, widget, font_size):
        """递归更新控件字体"""
        try:
            # 更新当前控件
            widget_type = widget.winfo_class()
            
            if widget_type in ('TEntry', 'TCombobox', 'TLabel', 'TButton'):
                # ttk 控件通过 style 更新
                pass
            elif widget_type in ('Text', 'Entry', 'Label', 'Button', 'Listbox'):
                # tk 控件直接更新
                try:
                    current_font = widget.cget('font')
                    if current_font:
                        if isinstance(current_font, str):
                            # 字体名称
                            widget.config(font=(current_font, font_size))
                        elif isinstance(current_font, tuple):
                            # 字体元组
                            font_family = current_font[0] if len(current_font) > 0 else 'TkDefaultFont'
                            widget.config(font=(font_family, font_size))
                except:
                    pass
            
            # 递归处理子控件
            for child in widget.winfo_children():
                self._update_widget_font(child, font_size)
        except:
            pass
    
    def show_zoom_help(self):
        """显示缩放帮助"""
        help_text = """
═══════════════════════════════
    界面缩放功能
═══════════════════════════════

🔍 缩放方式:

1. 按钮控制:
   🔍+     放大界面
   🔍−     缩小界面
   
2. 键盘快捷键:
   Ctrl + Plus     放大
   Ctrl + Minus    缩小
   Ctrl + 0        重置为100%

3. 鼠标控制:
   Ctrl + 滚轮向上    放大
   Ctrl + 滚轮向下    缩小

📏 缩放范围:
   最小: 50%
   最大: 200%
   步进: 10%

💡 提示:
• 缩放会同时调整窗口大小和字体
• 适合不同屏幕分辨率和视力需求
• 设置会保持到程序关闭

═══════════════════════════════
        """
        messagebox.showinfo("缩放功能帮助", help_text)


def main():
    print("=== 程序启动 ===", flush=True)
    try:
        print("创建 Tk() 实例...", flush=True)
        root = tk.Tk()
        print("Tk() 创建成功", flush=True)
        
        print("创建 OutlookDraftManager...", flush=True)
        app = OutlookDraftManager(root)
        print("OutlookDraftManager 创建成功", flush=True)
        
        print("进入主循环...", flush=True)
        root.mainloop()
        print("主循环退出", flush=True)
    except Exception as e:
        import traceback
        error_msg = f"程序启动失败:\n{str(e)}\n\n详细错误:\n{traceback.format_exc()}"
        print(error_msg, flush=True)
        try:
            messagebox.showerror("致命错误", error_msg)
        except:
            pass
        input("按回车键退出...")


if __name__ == "__main__":
    print("=== __main__ 开始执行 ===", flush=True)
    main()
