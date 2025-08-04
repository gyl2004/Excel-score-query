"""
主窗口UI组件
"""
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable, List
import os

from utils.config_loader import ConfigLoader
from utils.logger import default_logger
from ui.file_selector import FileSelector


class MainWindow:
    """主应用窗口类"""
    
    def __init__(self):
        """初始化主窗口"""
        self.root = tk.Tk()
        config_loader = ConfigLoader()
        self.config = config_loader.load_config()
        
        # 文件路径变量
        self.position_file_path = tk.StringVar()
        self.interview_file_path = tk.StringVar()
        
        # 文件选择器
        self.file_selector = FileSelector(self.root)
        
        # 回调函数
        self.on_position_file_select: Optional[Callable] = None
        self.on_interview_file_select: Optional[Callable] = None
        self.on_start_processing: Optional[Callable] = None
        self.on_configure_columns: Optional[Callable] = None
        
        # 列映射配置
        self.column_mappings: Optional[dict] = None
        
        # 输出列选择配置
        self.selected_output_columns: Optional[List[str]] = None
        
        # 最近生成的结果文件路径
        self.last_result_file: Optional[str] = None
        
        # 最近文件列表（最多保存5个）
        self.recent_files: List[str] = []
        
        self.setup_ui()
        default_logger.info("主窗口初始化完成")
    
    def setup_ui(self):
        """设置用户界面"""
        # 窗口基本配置
        self.root.title(self.config['app_config']['window_title'])
        self.root.geometry(self.config['app_config']['window_size'])
        self.root.resizable(True, True)
        
        # 设置窗口居中
        self._center_window()
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 创建标题
        title_label = ttk.Label(
            main_frame, 
            text="Excel岗位分数查询工具",
            font=('Arial', 16, 'bold')
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 创建文件选择区域
        self._create_file_selection_area(main_frame)
        
        # 创建列映射配置区域
        self._create_column_mapping_area(main_frame)
        
        # 创建操作按钮区域
        self._create_action_area(main_frame)
        
        # 创建状态显示区域
        self._create_status_area(main_frame)
        
        default_logger.info("UI界面设置完成")
    
    def _center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def _create_file_selection_area(self, parent):
        """创建文件选择区域"""
        # 职位表文件选择
        ttk.Label(parent, text="职位表文件:").grid(
            row=1, column=0, sticky=tk.W, pady=(0, 10)
        )
        
        position_entry = ttk.Entry(
            parent, 
            textvariable=self.position_file_path,
            state='readonly',
            width=50
        )
        position_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 10), padx=(10, 10))
        
        position_btn = ttk.Button(
            parent,
            text="选择文件",
            command=self._on_select_position_file
        )
        position_btn.grid(row=1, column=2, pady=(0, 10))
        
        # 面试人员名单文件选择
        ttk.Label(parent, text="面试人员名单:").grid(
            row=2, column=0, sticky=tk.W, pady=(0, 10)
        )
        
        interview_entry = ttk.Entry(
            parent,
            textvariable=self.interview_file_path,
            state='readonly',
            width=50
        )
        interview_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=(0, 10), padx=(10, 10))
        
        interview_btn = ttk.Button(
            parent,
            text="选择文件",
            command=self._on_select_interview_file
        )
        interview_btn.grid(row=2, column=2, pady=(0, 10))
    
    def _create_column_mapping_area(self, parent):
        """创建列映射配置区域"""
        mapping_frame = ttk.LabelFrame(parent, text="配置选项", padding="10")
        mapping_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(20, 0))
        mapping_frame.columnconfigure(2, weight=1)
        
        # 配置说明
        desc_label = ttk.Label(
            mapping_frame,
            text="配置列匹配关系和输出文件包含的列",
            foreground="gray"
        )
        desc_label.grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=(0, 10))
        
        # 列映射配置按钮
        self.config_columns_btn = ttk.Button(
            mapping_frame,
            text="配置列映射",
            command=self._on_configure_columns,
            state='disabled'
        )
        self.config_columns_btn.grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        
        # 输出列选择按钮
        self.select_output_columns_btn = ttk.Button(
            mapping_frame,
            text="选择输出列",
            command=self._on_select_output_columns,
            state='disabled'
        )
        self.select_output_columns_btn.grid(row=1, column=1, sticky=tk.W, padx=(0, 10))
        
        # 配置状态显示
        self.mapping_status_label = ttk.Label(
            mapping_frame,
            text="请先选择两个Excel文件",
            foreground="gray"
        )
        self.mapping_status_label.grid(row=1, column=2, sticky=tk.W, padx=(20, 0))
        
        # 输出列状态显示
        self.output_columns_status_label = ttk.Label(
            mapping_frame,
            text="使用默认输出列",
            foreground="gray"
        )
        self.output_columns_status_label.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=(5, 0))
    
    def _create_action_area(self, parent):
        """创建操作按钮区域"""
        action_frame = ttk.Frame(parent)
        action_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        self.start_btn = ttk.Button(
            action_frame,
            text="开始处理",
            command=self._on_start_processing_click,
            state='disabled'
        )
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 快速打开文件按钮组
        open_frame = ttk.Frame(action_frame)
        open_frame.pack(side=tk.LEFT, padx=(0, 10))
        
        self.open_file_btn = ttk.Button(
            open_frame,
            text="打开结果文件",
            command=self._on_open_result_file,
            state='disabled'
        )
        self.open_file_btn.pack(side=tk.LEFT)
        
        # 下拉菜单按钮，用于显示最近的文件
        self.open_menu_btn = ttk.Button(
            open_frame,
            text="▼",
            width=3,
            command=self._show_recent_files_menu,
            state='disabled'
        )
        self.open_menu_btn.pack(side=tk.LEFT)
        
        clear_btn = ttk.Button(
            action_frame,
            text="清空选择",
            command=self._clear_selections
        )
        clear_btn.pack(side=tk.LEFT)
    
    def _create_status_area(self, parent):
        """创建状态显示区域"""
        status_frame = ttk.LabelFrame(parent, text="状态信息", padding="10")
        status_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 0))
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(1, weight=1)
        
        # 进度条区域
        progress_frame = ttk.Frame(status_frame)
        progress_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(1, weight=1)
        
        ttk.Label(progress_frame, text="处理进度:").grid(row=0, column=0, sticky=tk.W)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            length=300,
            mode='determinate'
        )
        self.progress_bar.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        
        self.progress_label = ttk.Label(progress_frame, text="0%")
        self.progress_label.grid(row=0, column=2, sticky=tk.W)
        
        # 状态文本显示区域
        self.status_text = tk.Text(
            status_frame,
            height=8,
            width=70,
            wrap=tk.WORD,
            state='disabled'
        )
        self.status_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # 初始状态信息
        self._update_status("请选择职位表文件和面试人员名单文件")
    
    def _on_select_position_file(self):
        """处理职位表文件选择"""
        file_path = self.file_selector.select_position_file()
        if file_path:
            self.set_position_file_path(file_path)
            if self.on_position_file_select:
                self.on_position_file_select(file_path)
        self._check_ready_to_process()
    
    def _on_select_interview_file(self):
        """处理面试人员名单文件选择"""
        file_path = self.file_selector.select_interview_file()
        if file_path:
            self.set_interview_file_path(file_path)
            if self.on_interview_file_select:
                self.on_interview_file_select(file_path)
        self._check_ready_to_process()
    
    def _on_start_processing_click(self):
        """处理开始处理按钮点击"""
        if self.on_start_processing:
            self.on_start_processing()
    
    def _clear_selections(self):
        """清空文件选择"""
        self.position_file_path.set("")
        self.interview_file_path.set("")
        self.start_btn.configure(state='disabled')
        self.open_file_btn.configure(state='disabled')
        self.open_menu_btn.configure(state='disabled')
        self.last_result_file = None
        self.column_mappings = None
        self.selected_output_columns = None
        self.mapping_status_label.config(text="请先选择两个Excel文件", foreground="gray")
        self.config_columns_btn.configure(state='disabled')
        self.select_output_columns_btn.configure(state='disabled')
        self._update_output_columns_status()
        self._update_status("已清空文件选择，请重新选择文件")
        default_logger.info("用户清空了文件选择")
    
    def _on_configure_columns(self):
        """处理列映射配置"""
        if self.on_configure_columns:
            self.on_configure_columns()
    
    def _on_select_output_columns(self):
        """处理输出列选择"""
        try:
            if not self.position_file_path.get() or not self.interview_file_path.get():
                self.show_error("请先选择两个Excel文件")
                return
            
            # 获取可用的列
            available_columns = self._get_available_output_columns()
            
            if not available_columns:
                self.show_error("无法获取可用的列信息")
                return
            
            # 显示列选择对话框
            from ui.column_selection_dialog import show_column_selection_dialog
            
            # 获取默认选择的列
            default_columns = self.selected_output_columns or self._get_default_output_columns()
            
            selected_columns = show_column_selection_dialog(
                self.root,
                available_columns,
                default_columns
            )
            
            if selected_columns is not None:
                self.selected_output_columns = selected_columns
                self._update_output_columns_status()
                self._update_status(f"已选择 {len(selected_columns)} 个输出列")
                default_logger.info(f"用户选择输出列: {selected_columns}")
            
        except Exception as e:
            error_msg = f"选择输出列时发生错误: {str(e)}"
            self.show_error(error_msg)
            default_logger.error(error_msg)
    
    def _get_available_output_columns(self) -> List[str]:
        """获取可用的输出列"""
        try:
            from services.excel_reader import ExcelReader
            
            excel_reader = ExcelReader()
            
            # 获取职位表的列
            position_columns = excel_reader.get_column_names(self.position_file_path.get())
            
            # 添加处理结果相关的列
            result_columns = ['最低面试分数', '面试人数', '状态']
            
            # 合并所有可用列，去重并保持顺序
            all_columns = []
            seen = set()
            
            # 先添加职位表的列
            for col in position_columns:
                if col not in seen:
                    all_columns.append(col)
                    seen.add(col)
            
            # 再添加结果列
            for col in result_columns:
                if col not in seen:
                    all_columns.append(col)
                    seen.add(col)
            
            return all_columns
            
        except Exception as e:
            default_logger.error(f"获取可用输出列失败: {e}")
            return []
    
    def _get_default_output_columns(self) -> List[str]:
        """获取默认的输出列"""
        return [
            '招考职位',
            '职位代码', 
            '用人司局',
            '部门代码',
            '最低面试分数',
            '面试人数',
            '状态'
        ]
    
    def _update_output_columns_status(self):
        """更新输出列状态显示"""
        if self.selected_output_columns:
            count = len(self.selected_output_columns)
            self.output_columns_status_label.config(
                text=f"已选择 {count} 个输出列",
                foreground="green"
            )
        else:
            self.output_columns_status_label.config(
                text="使用默认输出列",
                foreground="gray"
            )
    
    def _on_open_result_file(self):
        """处理打开结果文件按钮点击"""
        if self.last_result_file and os.path.exists(self.last_result_file):
            try:
                # 使用系统默认程序打开文件
                import subprocess
                import platform
                
                system = platform.system()
                if system == "Windows":
                    os.startfile(self.last_result_file)
                elif system == "Darwin":  # macOS
                    subprocess.run(["open", self.last_result_file])
                else:  # Linux
                    subprocess.run(["xdg-open", self.last_result_file])
                
                self._update_status(f"已打开文件: {os.path.basename(self.last_result_file)}")
                default_logger.info(f"打开结果文件: {self.last_result_file}")
                
            except Exception as e:
                error_msg = f"打开文件失败: {str(e)}"
                self.show_error(error_msg)
                default_logger.error(error_msg)
        else:
            # 如果没有最近的结果文件，显示文件选择对话框
            self._open_file_dialog()
    
    def _open_file_dialog(self):
        """打开文件选择对话框"""
        from tkinter import filedialog
        
        try:
            # 设置文件类型过滤器
            filetypes = [
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
            
            # 打开文件选择对话框
            file_path = filedialog.askopenfilename(
                title="选择要打开的文件",
                filetypes=filetypes,
                initialdir=os.getcwd()  # 从当前工作目录开始
            )
            
            if file_path:
                # 使用系统默认程序打开选择的文件
                import subprocess
                import platform
                
                system = platform.system()
                if system == "Windows":
                    os.startfile(file_path)
                elif system == "Darwin":  # macOS
                    subprocess.run(["open", file_path])
                else:  # Linux
                    subprocess.run(["xdg-open", file_path])
                
                self._update_status(f"已打开文件: {os.path.basename(file_path)}")
                default_logger.info(f"用户选择并打开文件: {file_path}")
                
        except Exception as e:
            error_msg = f"打开文件失败: {str(e)}"
            self.show_error(error_msg)
            default_logger.error(error_msg)
    
    def _show_recent_files_menu(self):
        """显示最近文件菜单"""
        try:
            # 创建弹出菜单
            menu = tk.Menu(self.root, tearoff=0)
            
            # 添加最近的结果文件
            if self.last_result_file and os.path.exists(self.last_result_file):
                filename = os.path.basename(self.last_result_file)
                menu.add_command(
                    label=f"📄 {filename} (最新)",
                    command=lambda: self._open_specific_file(self.last_result_file)
                )
                menu.add_separator()
            
            # 添加最近文件列表
            if self.recent_files:
                for file_path in self.recent_files:
                    if os.path.exists(file_path) and file_path != self.last_result_file:
                        filename = os.path.basename(file_path)
                        menu.add_command(
                            label=f"📄 {filename}",
                            command=lambda fp=file_path: self._open_specific_file(fp)
                        )
                
                if len(self.recent_files) > 0:
                    menu.add_separator()
            
            # 添加浏览文件选项
            menu.add_command(
                label="📁 浏览其他文件...",
                command=self._open_file_dialog
            )
            
            # 添加打开当前目录选项
            menu.add_command(
                label="📂 打开当前目录",
                command=self._open_current_directory
            )
            
            # 显示菜单
            try:
                # 获取按钮位置
                x = self.open_menu_btn.winfo_rootx()
                y = self.open_menu_btn.winfo_rooty() + self.open_menu_btn.winfo_height()
                menu.post(x, y)
            except tk.TclError:
                # 如果无法获取位置，使用鼠标位置
                menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
            
        except Exception as e:
            error_msg = f"显示最近文件菜单失败: {str(e)}"
            self.show_error(error_msg)
            default_logger.error(error_msg)
    
    def _open_specific_file(self, file_path: str):
        """打开指定的文件"""
        try:
            if os.path.exists(file_path):
                import subprocess
                import platform
                
                system = platform.system()
                if system == "Windows":
                    os.startfile(file_path)
                elif system == "Darwin":  # macOS
                    subprocess.run(["open", file_path])
                else:  # Linux
                    subprocess.run(["xdg-open", file_path])
                
                self._update_status(f"已打开文件: {os.path.basename(file_path)}")
                default_logger.info(f"打开指定文件: {file_path}")
            else:
                self.show_error(f"文件不存在: {os.path.basename(file_path)}")
                # 从最近文件列表中移除不存在的文件
                if file_path in self.recent_files:
                    self.recent_files.remove(file_path)
                
        except Exception as e:
            error_msg = f"打开文件失败: {str(e)}"
            self.show_error(error_msg)
            default_logger.error(error_msg)
    
    def _open_current_directory(self):
        """打开当前工作目录"""
        try:
            import subprocess
            import platform
            
            current_dir = os.getcwd()
            system = platform.system()
            
            if system == "Windows":
                os.startfile(current_dir)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", current_dir])
            else:  # Linux
                subprocess.run(["xdg-open", current_dir])
            
            self._update_status(f"已打开目录: {current_dir}")
            default_logger.info(f"打开当前目录: {current_dir}")
            
        except Exception as e:
            error_msg = f"打开目录失败: {str(e)}"
            self.show_error(error_msg)
            default_logger.error(error_msg)
    
    def _add_to_recent_files(self, file_path: str):
        """添加文件到最近文件列表"""
        try:
            # 如果文件已存在，先移除
            if file_path in self.recent_files:
                self.recent_files.remove(file_path)
            
            # 添加到列表开头
            self.recent_files.insert(0, file_path)
            
            # 保持列表最多5个文件
            if len(self.recent_files) > 5:
                self.recent_files = self.recent_files[:5]
            
            default_logger.info(f"添加到最近文件列表: {file_path}")
            
        except Exception as e:
            default_logger.error(f"添加到最近文件列表失败: {e}")
    
    def _check_ready_to_process(self):
        """检查是否准备好开始处理"""
        if self.position_file_path.get() and self.interview_file_path.get():
            # 启用配置按钮
            self.config_columns_btn.configure(state='normal')
            self.select_output_columns_btn.configure(state='normal')
            self.mapping_status_label.config(text="请配置列映射关系", foreground="orange")
            
            # 只有配置了列映射才能开始处理
            if self.column_mappings:
                self.start_btn.configure(state='normal')
                self._update_status("文件选择和列映射配置完成，可以开始处理")
            else:
                self.start_btn.configure(state='disabled')
                self._update_status("文件选择完成，请配置列映射关系")
        else:
            self.start_btn.configure(state='disabled')
            self.config_columns_btn.configure(state='disabled')
            self.select_output_columns_btn.configure(state='disabled')
            self.mapping_status_label.config(text="请先选择两个Excel文件", foreground="gray")
    
    def _update_status(self, message: str):
        """更新状态信息"""
        self.status_text.configure(state='normal')
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.configure(state='disabled')
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def set_position_file_path(self, file_path: str):
        """设置职位表文件路径"""
        self.position_file_path.set(file_path)
        filename = os.path.basename(file_path)
        self._update_status(f"已选择职位表文件: {filename}")
        default_logger.info(f"设置职位表文件路径: {file_path}")
    
    def set_interview_file_path(self, file_path: str):
        """设置面试人员名单文件路径"""
        self.interview_file_path.set(file_path)
        filename = os.path.basename(file_path)
        self._update_status(f"已选择面试人员名单文件: {filename}")
        default_logger.info(f"设置面试人员名单文件路径: {file_path}")
    
    def show_error(self, error_msg: str):
        """显示错误信息"""
        messagebox.showerror("错误", error_msg)
        self._update_status(f"错误: {error_msg}")
        default_logger.error(f"显示错误信息: {error_msg}")
    
    def show_info(self, info_msg: str):
        """显示信息提示"""
        messagebox.showinfo("信息", info_msg)
        self._update_status(f"信息: {info_msg}")
        default_logger.info(f"显示信息: {info_msg}")
    
    def run(self):
        """运行主窗口"""
        default_logger.info("启动主窗口事件循环")
        self.root.mainloop()
    
    def show_progress(self, progress: int, message: str = ""):
        """
        显示处理进度
        
        Args:
            progress: 进度百分比 (0-100)
            message: 进度消息
        """
        # 确保进度值在有效范围内
        progress = max(0, min(100, progress))
        
        self.progress_var.set(progress)
        self.progress_label.config(text=f"{progress}%")
        
        if message:
            self._update_status(f"[{progress}%] {message}")
        
        # 强制更新UI
        self.root.update_idletasks()
        default_logger.info(f"更新进度: {progress}% - {message}")
    
    def reset_progress(self):
        """重置进度条"""
        self.progress_var.set(0)
        self.progress_label.config(text="0%")
        default_logger.info("进度条已重置")
    
    def show_results(self, results: dict):
        """
        显示处理结果
        
        Args:
            results: 处理结果字典
        """
        try:
            self._update_status("=" * 50)
            self._update_status("处理完成！结果摘要:")
            self._update_status("=" * 50)
            
            # 显示基本统计信息
            if 'total_positions' in results:
                self._update_status(f"总岗位数: {results['total_positions']}")
            
            if 'processed_positions' in results:
                self._update_status(f"成功处理岗位数: {results['processed_positions']}")
            
            if 'failed_positions' in results:
                self._update_status(f"处理失败岗位数: {results['failed_positions']}")
            
            if 'total_candidates' in results:
                self._update_status(f"总面试人员数: {results['total_candidates']}")
            
            # 显示输出文件信息
            if 'output_file' in results:
                self._update_status(f"输出文件: {results['output_file']}")
                # 记录最近的结果文件路径并启用打开文件按钮
                self.last_result_file = results['output_file']
                self._add_to_recent_files(results['output_file'])
                self.open_file_btn.configure(state='normal')
                self.open_menu_btn.configure(state='normal')
            
            # 显示处理时间
            if 'processing_time' in results:
                self._update_status(f"处理耗时: {results['processing_time']:.2f}秒")
            
            # 显示错误信息（如果有）
            if 'errors' in results and results['errors']:
                self._update_status("\n发现以下问题:")
                for error in results['errors']:
                    self._update_status(f"- {error}")
            
            # 显示警告信息（如果有）
            if 'warnings' in results and results['warnings']:
                self._update_status("\n注意事项:")
                for warning in results['warnings']:
                    self._update_status(f"- {warning}")
            
            self._update_status("=" * 50)
            
            # 显示成功消息框
            success_msg = f"处理完成！\n成功处理 {results.get('processed_positions', 0)} 个岗位"
            if 'output_file' in results:
                success_msg += f"\n结果已保存到: {os.path.basename(results['output_file'])}"
            
            self.show_info(success_msg)
            
            default_logger.info(f"显示处理结果: {results}")
            
        except Exception as e:
            error_msg = f"显示结果时发生错误: {str(e)}"
            self.show_error(error_msg)
            default_logger.error(error_msg)
    
    def show_processing_error(self, error_msg: str, details: str = ""):
        """
        显示处理错误
        
        Args:
            error_msg: 错误消息
            details: 详细错误信息
        """
        self._update_status("=" * 50)
        self._update_status("处理失败！")
        self._update_status("=" * 50)
        self._update_status(f"错误: {error_msg}")
        
        if details:
            self._update_status(f"详细信息: {details}")
        
        self._update_status("=" * 50)
        
        # 重置进度条
        self.reset_progress()
        
        # 显示错误对话框
        full_error = error_msg
        if details:
            full_error += f"\n\n详细信息:\n{details}"
        
        self.show_error(full_error)
        default_logger.error(f"处理错误: {error_msg} - {details}")
    
    def clear_status(self):
        """清空状态显示"""
        self.status_text.configure(state='normal')
        self.status_text.delete(1.0, tk.END)
        self.status_text.configure(state='disabled')
        self.reset_progress()
        default_logger.info("状态显示已清空")
    
    def set_processing_state(self, is_processing: bool):
        """
        设置处理状态
        
        Args:
            is_processing: 是否正在处理
        """
        if is_processing:
            self.start_btn.configure(state='disabled', text="处理中...")
            self._update_status("开始处理，请稍候...")
        else:
            self.start_btn.configure(state='normal', text="开始处理")
            
        default_logger.info(f"设置处理状态: {is_processing}")
    
    def set_column_mappings(self, mappings: dict):
        """
        设置列映射配置
        
        Args:
            mappings: 列映射字典
        """
        self.column_mappings = mappings
        if mappings:
            mapping_count = len(mappings)
            self.mapping_status_label.config(
                text=f"已配置 {mapping_count} 个列映射关系", 
                foreground="green"
            )
            self._update_status(f"列映射配置完成，共 {mapping_count} 个映射关系")
            default_logger.info(f"设置列映射: {mappings}")
        else:
            self.mapping_status_label.config(
                text="列映射配置已清空", 
                foreground="gray"
            )
            self._update_status("列映射配置已清空")
        
        # 重新检查是否可以开始处理
        self._check_ready_to_process()
    
    def get_column_mappings(self) -> Optional[dict]:
        """获取当前的列映射配置"""
        return self.column_mappings
    
    def get_selected_output_columns(self) -> Optional[List[str]]:
        """获取当前选择的输出列"""
        return self.selected_output_columns
    
    def set_selected_output_columns(self, columns: List[str]):
        """设置选择的输出列"""
        self.selected_output_columns = columns
        self._update_output_columns_status()
        default_logger.info(f"设置输出列: {columns}")
    
    def destroy(self):
        """销毁窗口"""
        self.root.destroy()
        default_logger.info("主窗口已销毁")