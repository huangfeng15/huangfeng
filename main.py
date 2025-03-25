import tkinter as tk
from tkinter import ttk, filedialog
import os
import sys  # 确保这行导入存在
import pandas as pd
import re
from typing import List, Optional, Dict
from pdf_processor import PDFProcessor
from excel_exporter import ExcelExporter
from config_manager import ConfigManager
from datetime import datetime

# 设置工作目录，从simplified_main.py移植过来的代码
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

# 版本信息
VERSION = "1.0版 (2025年3月)"

class PDFExtractorGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.config_manager = ConfigManager()
        self.config = self.config_manager.load_config()
        
        # 初始化属性
        self.files = []
        self.key_file = self.config.get('key_file', '')
        
        # 设置默认字体和版本信息 - 使用更美观的字体
        default_font = ('Microsoft YaHei UI', 11)  # 增大基础字体
        title_font = ('Microsoft YaHei UI', 12, 'bold')  # 增大标题字体
        version_font = ('Microsoft YaHei UI', 16, 'bold')  # 增大版本号字体
        
        self.root.title(f"PDF信息提取工具")
        self.root.geometry("1100x720")  # 增大窗口尺寸以适应更大的字体和更好的布局
        
        # 设置窗口位置和大小
        if self.config['window_size']:
            self.root.geometry(self.config['window_size'])
        if self.config['window_position']:
            self.root.geometry(f"+{self.config['window_position'][0]}+{self.config['window_position'][1]}")
        
        # 创建一个顶部框架放置版本号
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=15, pady=10)
        
        # 版本号标签 - 居中放置并使用更大字体
        version_label = ttk.Label(top_frame, text=f"PDF信息提取工具 - {VERSION}", 
                                 font=version_font, foreground="#0066cc")
        version_label.pack(pady=10, anchor="center")
        
        self.root.option_add('*Font', default_font)
        
        # 设置按钮样式
        style = ttk.Style()
        style.configure('TButton', font=default_font, padding=(12, 6))  # 增大按钮尺寸
        style.configure('Title.TLabel', font=title_font)
        style.configure('TCheckbutton', font=default_font)
        style.configure('TLabelframe.Label', font=title_font)  # 为分组标题设置字体
        
        # 创建主框架来包含所有控件
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 文件选择框 - 优化布局
        self.file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding=10)
        self.file_frame.pack(fill="x", padx=10, pady=8)
        
        file_buttons_frame = ttk.Frame(self.file_frame)
        file_buttons_frame.pack(side="left", fill="y", padx=10)
        
        ttk.Button(file_buttons_frame, text="选择PDF文件", command=self.select_files, 
                  width=15).pack(side="left", padx=10)
        ttk.Button(file_buttons_frame, text="选择文件夹", command=self.select_folder, 
                  width=15).pack(side="left", padx=10)
        
        file_info_frame = ttk.Frame(self.file_frame)
        file_info_frame.pack(side="left", fill="both", expand=True, padx=10)
        
        self.file_label = ttk.Label(file_info_frame, text="未选择文件")
        self.file_label.pack(side="left", padx=10)
        
        self.subfolder_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(file_info_frame, text="包含子文件夹", 
                       variable=self.subfolder_var).pack(side="right", padx=15)
        
        # 过滤选项单独一行
        filter_frame = ttk.Frame(main_frame)
        filter_frame.pack(fill="x", padx=20, pady=8)
        
        ttk.Label(filter_frame, text="文件名关键词(用逗号分隔):").pack(side="left", padx=10)
        self.filter_var = tk.StringVar(value=self.config['filter_keywords'])
        self.filter_entry = ttk.Entry(filter_frame, textvariable=self.filter_var, width=50)
        self.filter_entry.pack(side="left", padx=10, fill="x", expand=True)
        
        # 创建配置面板 - 三栏布局
        config_frame = ttk.LabelFrame(main_frame, text="配置选项", padding=10)
        config_frame.pack(fill="x", padx=10, pady=8)
        
        # 键名配置
        key_frame = ttk.Frame(config_frame)
        key_frame.pack(side="left", fill="x", expand=True, padx=10)
        
        ttk.Button(key_frame, text="选择键名文件", command=self.select_key_file, 
                  width=15).pack(side="left", padx=10)
        self.key_label = ttk.Label(key_frame, 
                                text=f"已选择: {os.path.basename(self.key_file)}" if self.key_file 
                                else "未选择键名文件")
        self.key_label.pack(side="left", padx=10)
        
        # 中间放置阅读顺序
        order_frame = ttk.Frame(config_frame)
        order_frame.pack(side="left", padx=20)
        
        ttk.Label(order_frame, text="阅读顺序:").pack(side="left")
        self.read_order = tk.StringVar(value=self.config['read_order'])
        ttk.Radiobutton(order_frame, text="从上到下", variable=self.read_order, 
                       value="top_to_bottom").pack(side="left", padx=10)
        ttk.Radiobutton(order_frame, text="从左到右", variable=self.read_order, 
                       value="left_to_right").pack(side="left", padx=10)
        
        # 右侧放置提取选项
        options_frame = ttk.Frame(config_frame)
        options_frame.pack(side="right", padx=20)
        
        self.allow_empty = tk.BooleanVar(value=self.config['allow_empty'])
        ttk.Checkbutton(options_frame, text="允许值为空", 
                       variable=self.allow_empty).pack(side="left")

        # 项目处理选项
        project_frame = ttk.LabelFrame(main_frame, text="项目处理方式", padding=10)
        project_frame.pack(fill="x", padx=10, pady=8)
        
        self.project_mode = tk.StringVar(value=self.config.get('project_mode', 'same'))  # 从配置读取，默认为same
        ttk.Radiobutton(project_frame, text="每个文件夹作为不同项目", 
                       variable=self.project_mode, value="separate").pack(side="left", padx=20)
        ttk.Radiobutton(project_frame, text="所有文件夹作为同一项目", 
                       variable=self.project_mode, value="same").pack(side="left", padx=20)

        # Excel操作框
        excel_frame = ttk.LabelFrame(main_frame, text="Excel操作", padding=10)
        excel_frame.pack(fill="x", padx=10, pady=8)
        
        # Excel按钮区域
        excel_buttons_frame = ttk.Frame(excel_frame)
        excel_buttons_frame.pack(side="left", padx=10)
        
        ttk.Button(excel_buttons_frame, text="选择现有Excel", 
                  command=self.select_excel, width=15).pack(side="left", padx=10)
        
        # Excel选项区域
        excel_options_frame = ttk.Frame(excel_frame)
        excel_options_frame.pack(side="left", fill="x", expand=True, padx=10)
        
        ttk.Label(excel_options_frame, text="标题行:").pack(side="left", padx=5)
        self.header_var = tk.StringVar(value="1")
        ttk.Entry(excel_options_frame, textvariable=self.header_var, 
                 width=5).pack(side="left", padx=5)
        
        ttk.Label(excel_options_frame, text="工作表:").pack(side="left", padx=10)
        self.sheet_var = tk.StringVar()
        self.sheet_combobox = ttk.Combobox(excel_options_frame, textvariable=self.sheet_var, 
                                          width=20, state="disabled")
        self.sheet_combobox.pack(side="left", padx=5)
        
        self.append_button = ttk.Button(excel_options_frame, text="新增表格信息", 
                                      command=self.append_to_excel,
                                      width=15,
                                      state='disabled')
        self.append_button.pack(side="right", padx=10)
        
        # 创建底部框架
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(side="bottom", fill="x", padx=10, pady=15)
        
        # 进度条区域
        progress_frame = ttk.LabelFrame(bottom_frame, text="处理进度", padding=10)
        progress_frame.pack(fill="x", pady=8)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, 
                                          variable=self.progress_var, 
                                          maximum=100)
        self.progress_bar.pack(fill="x", padx=10, pady=5)
        
        # 底部按钮和状态区域
        button_frame = ttk.Frame(bottom_frame)
        button_frame.pack(fill="x", pady=10)
        
        # 状态栏
        self.status_var = tk.StringVar()
        status_label = ttk.Label(button_frame, textvariable=self.status_var, 
                                font=('Microsoft YaHei UI', 10))
        status_label.pack(side="left", fill="x", expand=True, padx=10)
        
        # 处理按钮 - 使用更大的尺寸
        self.process_button = ttk.Button(button_frame, text="处理并导出", 
                                       command=self.process_files,
                                       width=20)  # 增大按钮宽度
        self.process_button.pack(side="right", padx=10)
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def on_closing(self):
        """保存配置并关闭程序"""
        self.config_manager.update_window_state(self.root)
        config = {
            'read_order': self.read_order.get(),
            'allow_empty': self.allow_empty.get(),
            'filter_keywords': self.filter_var.get(),
            'key_file': self.key_file or '',
            'last_folder': os.path.dirname(self.files[0]) if self.files else '',
            'last_save_folder': self.config.get('last_save_folder', ''),
            'window_position': self.config.get('window_position'),
            'window_size': self.config.get('window_size'),
            'project_mode': self.project_mode.get()  # 添加项目处理方式的保存
        }
        self.config_manager.save_config(config)
        self.root.destroy()
        
    def select_files(self):
        kwargs = self.config_manager.get_file_dialog_kwargs('file')
        self.files = filedialog.askopenfilenames(
            filetypes=[("PDF文件", "*.pdf")],
            title="选择PDF文件",
            **kwargs
        )
        if self.files:
            self._update_file_label()
            # 更新最后访问的文件夹
            self.config['last_folder'] = os.path.dirname(self.files[0])
            self.config_manager.save_config(self.config)
        
    def select_folder(self):
        kwargs = self.config_manager.get_file_dialog_kwargs('file')
        folder = filedialog.askdirectory(
            title="选择包含PDF文件的文件夹",
            **kwargs
        )
        if folder:
            self.files = []
            keywords = [k.strip() for k in self.filter_var.get().split(',') if k.strip()]
            
            def process_folder(folder_path, is_root=True):
                folder_files = []
                for item in os.listdir(folder_path):
                    full_path = os.path.join(folder_path, item)
                    if os.path.isfile(full_path) and full_path.lower().endswith('.pdf'):
                        if not keywords or any(k in item for k in keywords):
                            folder_files.append(full_path)
                    elif os.path.isdir(full_path) and self.subfolder_var.get():
                        sub_files = process_folder(full_path, False)
                        folder_files.extend(sub_files)
                return folder_files

            self.files = process_folder(folder)
            self._update_file_label()
            self.config['last_folder'] = folder
            self.config_manager.save_config(self.config)
            
    def select_key_file(self):
        initial_dir = os.path.dirname(self.key_file) if self.key_file else None
        key_file = filedialog.askopenfilename(
            filetypes=[("文本文件", "*.txt")],
            title="选择键名配置文件",
            initialdir=initial_dir
        )
        if key_file:
            self.key_file = key_file
            self.key_label.config(text=f"已选择: {os.path.basename(key_file)}")
            self._save_current_config()
        else:
            self.key_label.config(text="必须选择键名文件")
            
    def _save_current_config(self):
        """保存当前配置"""
        config = {
            'read_order': self.read_order.get(),
            'allow_empty': self.allow_empty.get(),
            'filter_keywords': self.filter_var.get(),
            'key_file': self.key_file or '',
            'last_folder': os.path.dirname(self.files[0]) if self.files else ''
        }
        self.config_manager.save_config(config)
            
    def _update_file_label(self):
        self.file_label.config(text=f"已选择 {len(self.files)} 个文件")
    
    def select_excel(self):
        excel_file = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx")],
            title="选择Excel文件"
        )
        if excel_file:
            try:
                # 获取工作表列表
                exporter = ExcelExporter()
                sheet_names = exporter.get_excel_sheets(excel_file)
                
                if not sheet_names:
                    self.status_var.set("无法获取工作表信息，请检查Excel文件格式")
                    return
                
                # 更新工作表下拉列表
                self.sheet_combobox['values'] = sheet_names
                self.sheet_combobox.current(0)  # 默认选择第一个工作表
                self.sheet_combobox['state'] = 'readonly'  # 设置为只读状态
                
                # 存储Excel文件信息，但不立即读取标题行
                self.excel_file = excel_file
                
                # 提示用户选择工作表和标题行
                self.status_var.set(f"已选择Excel文件: {os.path.basename(excel_file)}，请选择工作表和标题行")
                
                # 启用新增按钮
                self.append_button.config(state='normal')
                self.process_button.config(state='disabled')
                
                # 绑定工作表选择事件
                self.sheet_var.trace_add("write", self._validate_header_row)
                self.header_var.trace_add("write", self._validate_header_row)
                
            except Exception as e:
                self.status_var.set(f"读取Excel出错: {str(e)}")
                self.excel_file = None
                self.sheet_combobox['state'] = 'disabled'
                self.append_button.config(state='disabled')
                self.process_button.config(state='normal')
                
    def _validate_header_row(self, *args):
        """验证标题行是否有效"""
        if not hasattr(self, 'excel_file') or not self.excel_file:
            return
            
        try:
            sheet_name = self.sheet_var.get()
            if not sheet_name:
                return
                
            # 尝试解析标题行
            try:
                header_row = int(self.header_var.get()) - 1
                if header_row < 0:
                    self.status_var.set("标题行必须大于0")
                    self.existing_excel = None
                    return
            except ValueError:
                self.status_var.set("标题行必须是数字")
                self.existing_excel = None
                return
                
            # 读取指定工作表和标题行
            df = pd.read_excel(self.excel_file, header=header_row, sheet_name=sheet_name)
            
            # 检查标题行是否有内容
            if df.columns.empty or all(str(col).strip() == '' for col in df.columns):
                self.status_var.set(f"错误：第 {header_row + 1} 行不包含任何标题，请重新选择标题行")
                self.existing_excel = None
                return
                
            # 存储有效的Excel信息
            self.existing_excel = {
                'file': self.excel_file,
                'header_row': header_row,
                'columns': list(df.columns),
                'sheet_name': sheet_name
            }
            
            self.status_var.set(f"已选择Excel文件: {os.path.basename(self.excel_file)} (工作表: {sheet_name}，标题行: {header_row + 1})")
            self.append_button.config(state='normal')
            
        except Exception as e:
            self.status_var.set(f"验证标题行出错: {str(e)}")
            self.existing_excel = None
                
    def append_to_excel(self):
        """将提取的信息新增到现有Excel"""
        if not hasattr(self, 'existing_excel') or not self.existing_excel:
            self.status_var.set("请先选择Excel文件")
            return
            
        if not self.files:
            self.status_var.set("请先选择PDF文件")
            return
        
        if not self.key_file:
            self.status_var.set("请先选择键名文件")
            return
            
        try:
            # 读取键名文件
            with open(self.key_file, 'r', encoding='utf-8') as f:
                key_names = [line.strip() for line in f if line.strip()]
                
            processor = PDFProcessor(
                read_order=self.read_order.get(),
                allow_empty=self.allow_empty.get(),
                custom_keys=key_names
            )
            
            # 计算总文件数用于进度显示
            total_files = len(self.files)
            processed_count = 0
            self.progress_var.set(0)  # 重置进度条
            
            # 按文件夹组织文件
            folder_files = {}
            for file in self.files:
                folder = os.path.dirname(file)
                if folder not in folder_files:
                    folder_files[folder] = []
                folder_files[folder].append(file)

            combined_folder_results = {}
            all_results = []
            skipped_folders = []

            for folder, files in folder_files.items():
                folder_results = {}
                
                for file in files:
                    self.status_var.set(f"正在处理: {os.path.basename(file)} ({processed_count + 1}/{total_files})")
                    
                    try:
                        results = processor.process_pdf(file)
                        if results:
                            for item in results:
                                # 优先使用新的非空值
                                if item['key'] not in folder_results or (
                                    item['value'].strip() and not folder_results[item['key']]['value'].strip()):
                                    folder_results[item['key']] = item
                        
                    except Exception as e:
                        self.status_var.set(f"处理文件出错: {str(e)}")
                    
                    processed_count += 1
                    # 更新进度条
                    progress = (processed_count / total_files) * 100
                    self.progress_var.set(progress)
                    self.root.update()

                # 检查采购项目名称是否为空
                has_project_name = False
                for key, item in folder_results.items():
                    if '采购项目名称' in key and item['value'].strip():
                        has_project_name = True
                        break
                
                if not has_project_name:
                    skipped_folders.append(os.path.basename(folder))
                    continue

                # 根据项目处理模式决定如何处理结果
                if self.project_mode.get() == "same":
                    # 合并到总结果中
                    for key, item in folder_results.items():
                        if key not in combined_folder_results or (
                            item['value'].strip() and not combined_folder_results[key]['value'].strip()):
                            combined_folder_results[key] = item
                else:
                    # 每个文件夹作为独立项目
                    row_data = {col: '' for col in self.existing_excel['columns']}
                    
                    # 将文件夹结果转换为Excel行
                    for item in folder_results.values():
                        matching_col = self._find_matching_column(item['key'], self.existing_excel['columns'])
                        if matching_col:
                            row_data[matching_col] = item['value']
                    
                    all_results.append(row_data)

            # 如果是合并模式，添加合并后的结果
            if self.project_mode.get() == "same" and combined_folder_results:
                # 检查合并模式下采购项目名称是否为空
                has_project_name = False
                for key, item in combined_folder_results.items():
                    if '采购项目名称' in key and item['value'].strip():
                        has_project_name = True
                        break
                
                if has_project_name:
                    row_data = {col: '' for col in self.existing_excel['columns']}
                    
                    # 将合并结果转换为Excel行
                    for item in combined_folder_results.values():
                        matching_col = self._find_matching_column(item['key'], self.existing_excel['columns'])
                        if matching_col:
                            row_data[matching_col] = item['value']
                    
                    all_results.append(row_data)

            if all_results:
                # 追加到现有Excel
                try:
                    exporter = ExcelExporter()
                    # 将工作表信息传递给exporter
                    exporter.export_to_excel(
                        all_results, 
                        self.existing_excel['file'],
                        existing_excel=self.existing_excel,
                        append_mode=True,
                        sheet_name=self.existing_excel.get('sheet_name')
                    )
                    if skipped_folders:
                        self.status_var.set(f"已成功新增数据到Excel。跳过了{len(skipped_folders)}个文件夹，因为采购项目名称为空。")
                    else:
                        self.status_var.set("已成功新增数据到Excel")
                except Exception as e:
                    self.status_var.set(f"保存Excel时出错: {str(e)}")
            else:
                if skipped_folders:
                    self.status_var.set(f"未找到可提取的内容。所有文件夹({len(skipped_folders)}个)的采购项目名称都为空。")
                else:
                    self.status_var.set("未找到可提取的内容")

            # 清除Excel选择
            self.existing_excel = None
            self.append_button.config(state='disabled')
            self.process_button.config(state='normal')
            # 完成后确保进度条显示100%
            self.progress_var.set(100)
            self.root.update()
            
        except Exception as e:
            self.status_var.set(f"处理出错: {str(e)}")
        finally:
            # 清理工作
            self.existing_excel = None
            self.append_button.config(state='disabled')
            self.process_button.config(state='normal')
            # 延迟重置进度条
            self.root.after(1000, lambda: self.progress_var.set(0))

    def _find_matching_column(self, key: str, columns: List[str]) -> Optional[str]:
        """查找与键名严格匹配的列名"""
        if key is None:
            return None
            
        # 确保key是字符串类型
        key = str(key) if key is not None else ""
        key = self._normalize_text(key)
        
        # 首先尝试完全匹配
        for col in columns:
            col_str = str(col) if col is not None else ""
            if self._normalize_text(col_str) == key:
                return col
                
        # 对于时间相关的键，需要特殊处理
        if '时间' in key:
            # 完全匹配所有可能的键名
            for col in columns:
                col_str = str(col) if col is not None else ""
                col_norm = self._normalize_text(col_str)
                # 只有当键名完全包含在列名中时才匹配
                if key == col_norm:
                    # 验证值是否符合时间格式
                    return col
                    
        return None

    def _is_valid_time_format(self, value: str) -> bool:
        """验证是否为有效的时间格式"""
        if not value:
            return False
            
        # 移除所有空白字符
        value = re.sub(r'\s+', '', value)
        
        # 常见的时间格式模式
        patterns = [
            r'^\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?$',  # 2023-01-01, 2023年01月01日
            r'^\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?\s*\d{1,2}[-:时]\d{1,2}分?$',  # 2023-01-01 10:30
            r'^\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?\s*\d{1,2}[-:时]\d{1,2}分?\d{1,2}秒?$',  # 2023-01-01 10:30:00
            r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}T\d{1,2}:\d{1,2}:\d{1,2}$',  # 2023-01-01T10:30:00
            r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}\s+\d{1,2}:\d{1,2}:\d{1,2}$',  # 2023-01-01 10:30:00
            r'^\d{4}年\d{1,2}月\d{1,2}日\d{1,2}时\d{1,2}分$',  # 2023年01月01日10时30分
        ]
        
        return any(re.match(pattern, value) for pattern in patterns)

    def _normalize_text(self, text: str) -> str:
        """标准化文本，移除所有空白字符但保留基本文本"""
        if not text:
            return ""
        # 移除所有空白字符
        text = re.sub(r'\s+', '', text)
        # 移除中英文冒号和括号
        text = text.rstrip('：:')
        # 统一括号格式
        text = text.replace('（', '(').replace('）', ')')
        # 移除可能的括号内容
        text = re.sub(r'\([^)]*\)', '', text)
        # 转换为小写
        return text.lower()

    def process_files(self):
        try:
            if not self.files:
                self.status_var.set("请先选择PDF文件")
                return
            
            if not self.key_file:
                self.status_var.set("请先选择键名文件")
                return
                
            # 读取键名文件
            key_names = []
            if self.key_file and os.path.exists(self.key_file):
                with open(self.key_file, 'r', encoding='utf-8') as f:
                    key_names = [line.strip() for line in f if line.strip()]
                    
            processor = PDFProcessor(
                read_order=self.read_order.get(),
                allow_empty=self.allow_empty.get(),
                custom_keys=key_names
            )
            results = []
            total_files = len(self.files)
            processed_count = 0
            self.progress_var.set(0)  # 重置进度条
            
            # 按文件夹组织文件
            folder_files = {}
            for file in self.files:
                folder = os.path.dirname(file)
                if folder not in folder_files:
                    folder_files[folder] = []
                folder_files[folder].append(file)

            current_count = 0
            for folder, files in folder_files.items():
                folder_results = []
                for file in files:
                    self.status_var.set(f"正在处理: {os.path.basename(file)} ({processed_count + 1}/{total_files})")
                    
                    try:
                        result = processor.process_pdf(file)
                        if result:
                            for item in result:
                                item['filename'] = os.path.basename(file)
                                item['folder'] = os.path.basename(folder)
                            folder_results.extend(result)
                    except Exception as e:
                        self.status_var.set(f"处理文件 {os.path.basename(file)} 时出错: {str(e)}")
                    processed_count += 1
                    # 更新进度条
                    progress = (processed_count / total_files) * 100
                    self.progress_var.set(progress)
                    self.root.update()

                # 根据项目处理模式决定是否合并结果
                if self.project_mode.get() == "same" and folder_results:
                    # 只保留每个键的第一个值
                    unique_results = {}
                    for item in folder_results:
                        if item['key'] not in unique_results:
                            unique_results[item['key']] = item
                    folder_results = list(unique_results.values())
                
                results.extend(folder_results)

            # 更新最终状态
            if results:
                kwargs = self.config_manager.get_file_dialog_kwargs('save')
                output_file = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel文件", "*.xlsx")],
                    title="保存Excel文件",
                    **kwargs
                )
                if output_file:
                    self.config['last_save_folder'] = os.path.dirname(output_file)
                    self.config_manager.save_config(self.config)
                    exporter = ExcelExporter()
                    exporter.export_to_excel(
                        results, 
                        output_file,
                        existing_excel=getattr(self, 'existing_excel', None),
                        append_mode=False  # 不传递 append_mode 参数
                    )
                    self.status_var.set("导出完成！")
            else:
                self.status_var.set("未找到可提取的内容")
            
            # 完成后确保进度条显示100%
            self.progress_var.set(100)
            self.root.update()

        except Exception as e:
            self.status_var.set(f"处理出错: {str(e)}")
        finally:
            # 延迟重置进度条
            self.root.after(1000, lambda: self.progress_var.set(0))

if __name__ == "__main__":
    app = PDFExtractorGUI()
    app.root.mainloop()
