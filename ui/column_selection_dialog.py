"""
列选择对话框
允许用户选择要在输出Excel文件中包含的列
"""
import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict, Optional, Set


class ColumnSelectionDialog:
    """列选择对话框类"""
    
    def __init__(self, parent, available_columns: List[str], default_columns: List[str] = None):
        """
        初始化列选择对话框
        
        Args:
            parent: 父窗口
            available_columns: 可用的列名列表
            default_columns: 默认选中的列名列表
        """
        self.parent = parent
        self.available_columns = available_columns
        self.default_columns = default_columns or []
        self.selected_columns = []
        self.result = None
        
        # 创建对话框窗口
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("选择输出列")
        self.dialog.geometry("500x600")
        self.dialog.resizable(True, True)
        
        # 设置模态
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 居中显示
        self._center_dialog()
        
        # 创建UI
        self._create_ui()
        
        # 设置默认选择
        self._set_default_selections()
    
    def _center_dialog(self):
        """将对话框居中显示"""
        self.dialog.update_idletasks()
        
        # 获取父窗口位置和大小
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        # 计算对话框位置
        dialog_width = self.dialog.winfo_width()
        dialog_height = self.dialog.winfo_height()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.dialog.geometry(f"+{x}+{y}")
    
    def _create_ui(self):
        """创建用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(
            main_frame,
            text="选择要在输出文件中包含的列",
            font=('Arial', 12, 'bold')
        )
        title_label.pack(pady=(0, 15))
        
        # 说明文字
        desc_label = ttk.Label(
            main_frame,
            text="勾选您希望在Excel输出文件中显示的列。默认已选择常用列。",
            foreground="gray"
        )
        desc_label.pack(pady=(0, 15))
        
        # 创建选择区域
        self._create_selection_area(main_frame)
        
        # 创建预设按钮区域
        self._create_preset_area(main_frame)
        
        # 创建按钮区域
        self._create_button_area(main_frame)
    
    def _create_selection_area(self, parent):
        """创建列选择区域"""
        # 选择框架
        selection_frame = ttk.LabelFrame(parent, text="可用列", padding="10")
        selection_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # 创建滚动区域
        canvas = tk.Canvas(selection_frame, height=300)
        scrollbar = ttk.Scrollbar(selection_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 布局滚动组件
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 存储复选框变量
        self.checkbox_vars = {}
        self.checkboxes = {}
        
        # 定义列的分类和描述
        column_categories = self._categorize_columns()
        
        # 创建分类的复选框
        for category, columns in column_categories.items():
            if not columns:
                continue
                
            # 分类标题
            category_label = ttk.Label(
                scrollable_frame,
                text=category,
                font=('Arial', 10, 'bold'),
                foreground="blue"
            )
            category_label.pack(anchor="w", pady=(10, 5))
            
            # 该分类下的列
            for column in columns:
                var = tk.BooleanVar()
                self.checkbox_vars[column] = var
                
                checkbox = ttk.Checkbutton(
                    scrollable_frame,
                    text=self._get_column_display_name(column),
                    variable=var,
                    command=self._on_selection_change
                )
                checkbox.pack(anchor="w", padx=(20, 0), pady=2)
                self.checkboxes[column] = checkbox
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind("<MouseWheel>", _on_mousewheel)
    
    def _categorize_columns(self) -> Dict[str, List[str]]:
        """将列按类别分组"""
        categories = {
            "基本信息": [],
            "岗位详情": [],
            "招录要求": [],
            "联系信息": [],
            "处理结果": [],
            "其他信息": []
        }
        
        # 定义列的分类规则
        basic_keywords = ['招考职位', '职位代码', '用人司局', '部门名称', '部门代码']
        detail_keywords = ['职位属性', '职位分布', '职位简介', '机构性质', '机构层级', '工作地点', '落户地点']
        requirement_keywords = ['招考人数', '专业', '学历', '学位', '政治面貌', '基层工作', '服务基层', '考试类别']
        contact_keywords = ['咨询电话', '部门网站', '联系方式']
        result_keywords = ['最低面试分数', '面试人数', '状态', '匹配', '处理']
        
        for column in self.available_columns:
            column_lower = column.lower()
            
            if any(keyword in column for keyword in basic_keywords):
                categories["基本信息"].append(column)
            elif any(keyword in column for keyword in detail_keywords):
                categories["岗位详情"].append(column)
            elif any(keyword in column for keyword in requirement_keywords):
                categories["招录要求"].append(column)
            elif any(keyword in column for keyword in contact_keywords):
                categories["联系信息"].append(column)
            elif any(keyword in column for keyword in result_keywords):
                categories["处理结果"].append(column)
            else:
                categories["其他信息"].append(column)
        
        return categories
    
    def _get_column_display_name(self, column: str) -> str:
        """获取列的显示名称（添加描述）"""
        descriptions = {
            '招考职位': '招考职位 (岗位名称)',
            '职位代码': '职位代码 (唯一标识)',
            '用人司局': '用人司局 (招录部门)',
            '部门代码': '部门代码',
            '部门名称': '部门名称',
            '最低面试分数': '最低面试分数 (处理结果)',
            '面试人数': '面试人数 (处理结果)',
            '状态': '状态 (处理结果)',
            '招考人数': '招考人数 (计划招录)',
            '专业': '专业要求',
            '学历': '学历要求',
            '学位': '学位要求',
            '工作地点': '工作地点',
            '咨询电话1': '咨询电话1',
            '咨询电话2': '咨询电话2',
            '咨询电话3': '咨询电话3'
        }
        
        return descriptions.get(column, column)
    
    def _create_preset_area(self, parent):
        """创建预设按钮区域"""
        preset_frame = ttk.LabelFrame(parent, text="快速选择", padding="10")
        preset_frame.pack(fill=tk.X, pady=(0, 15))
        
        button_frame = ttk.Frame(preset_frame)
        button_frame.pack()
        
        # 基本信息预设
        basic_btn = ttk.Button(
            button_frame,
            text="基本信息",
            command=self._select_basic_preset
        )
        basic_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 详细信息预设
        detailed_btn = ttk.Button(
            button_frame,
            text="详细信息",
            command=self._select_detailed_preset
        )
        detailed_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 全选按钮
        all_btn = ttk.Button(
            button_frame,
            text="全选",
            command=self._select_all
        )
        all_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 清空按钮
        none_btn = ttk.Button(
            button_frame,
            text="清空",
            command=self._select_none
        )
        none_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 恢复默认按钮
        default_btn = ttk.Button(
            button_frame,
            text="恢复默认",
            command=self._restore_default
        )
        default_btn.pack(side=tk.LEFT)
    
    def _create_button_area(self, parent):
        """创建按钮区域"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(15, 0))
        
        # 选择计数标签
        self.count_label = ttk.Label(
            button_frame,
            text="已选择 0 列",
            foreground="gray"
        )
        self.count_label.pack(side=tk.LEFT)
        
        # 右侧按钮
        right_buttons = ttk.Frame(button_frame)
        right_buttons.pack(side=tk.RIGHT)
        
        # 确定按钮
        ok_btn = ttk.Button(
            right_buttons,
            text="确定",
            command=self._on_ok
        )
        ok_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        # 取消按钮
        cancel_btn = ttk.Button(
            right_buttons,
            text="取消",
            command=self._on_cancel
        )
        cancel_btn.pack(side=tk.RIGHT)
    
    def _set_default_selections(self):
        """设置默认选择"""
        if self.default_columns:
            for column in self.default_columns:
                if column in self.checkbox_vars:
                    self.checkbox_vars[column].set(True)
        else:
            # 如果没有指定默认列，选择基本信息
            self._select_basic_preset()
        
        self._on_selection_change()
    
    def _select_basic_preset(self):
        """选择基本信息预设"""
        basic_columns = [
            '招考职位', '职位代码', '用人司局', '部门代码', '部门名称',
            '最低面试分数', '面试人数', '状态'
        ]
        
        # 清空所有选择
        for var in self.checkbox_vars.values():
            var.set(False)
        
        # 选择基本列
        for column in basic_columns:
            if column in self.checkbox_vars:
                self.checkbox_vars[column].set(True)
        
        self._on_selection_change()
    
    def _select_detailed_preset(self):
        """选择详细信息预设"""
        detailed_columns = [
            '招考职位', '职位代码', '用人司局', '部门代码', '部门名称',
            '招考人数', '专业', '学历', '学位', '工作地点',
            '最低面试分数', '面试人数', '状态'
        ]
        
        # 清空所有选择
        for var in self.checkbox_vars.values():
            var.set(False)
        
        # 选择详细列
        for column in detailed_columns:
            if column in self.checkbox_vars:
                self.checkbox_vars[column].set(True)
        
        self._on_selection_change()
    
    def _select_all(self):
        """全选"""
        for var in self.checkbox_vars.values():
            var.set(True)
        self._on_selection_change()
    
    def _select_none(self):
        """清空选择"""
        for var in self.checkbox_vars.values():
            var.set(False)
        self._on_selection_change()
    
    def _restore_default(self):
        """恢复默认选择"""
        self._set_default_selections()
    
    def _on_selection_change(self):
        """选择变化时的回调"""
        selected_count = sum(1 for var in self.checkbox_vars.values() if var.get())
        self.count_label.config(text=f"已选择 {selected_count} 列")
    
    def _on_ok(self):
        """确定按钮回调"""
        # 获取选中的列
        selected = [column for column, var in self.checkbox_vars.items() if var.get()]
        
        if not selected:
            messagebox.showwarning("警告", "请至少选择一列！")
            return
        
        self.selected_columns = selected
        self.result = selected
        self.dialog.destroy()
    
    def _on_cancel(self):
        """取消按钮回调"""
        self.result = None
        self.dialog.destroy()
    
    def show(self) -> Optional[List[str]]:
        """
        显示对话框并返回结果
        
        Returns:
            选中的列名列表，如果取消则返回None
        """
        # 等待对话框关闭
        self.dialog.wait_window()
        return self.result


def show_column_selection_dialog(parent, available_columns: List[str], 
                                default_columns: List[str] = None) -> Optional[List[str]]:
    """
    显示列选择对话框的便捷函数
    
    Args:
        parent: 父窗口
        available_columns: 可用的列名列表
        default_columns: 默认选中的列名列表
        
    Returns:
        选中的列名列表，如果取消则返回None
    """
    dialog = ColumnSelectionDialog(parent, available_columns, default_columns)
    return dialog.show()