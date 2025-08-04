"""
列映射配置对话框
允许用户选择两个表之间的列匹配关系
"""
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List, Optional, Tuple
import logging


class ColumnMappingDialog:
    """列映射配置对话框类"""
    
    def __init__(self, parent, position_columns: List[str], interview_columns: List[str]):
        """
        初始化列映射对话框
        
        Args:
            parent: 父窗口
            position_columns: 岗位表的列名列表
            interview_columns: 面试人员表的列名列表
        """
        self.parent = parent
        self.position_columns = position_columns
        self.interview_columns = interview_columns
        self.result = None
        self.dialog = None
        
        # 默认的匹配建议
        self.default_mappings = {
            "职位代码": ["职位代码", "岗位代码", "代码"],
            "用人单位": ["用人单位", "用人司局", "招录机关", "部门名称"],
            "招考职位": ["招考职位", "岗考职位", "职位名称"],
            "部门代码": ["部门代码", "机构代码"]
        }
        
        self._create_dialog()
    
    def _create_dialog(self):
        """创建对话框界面"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("配置列匹配关系")
        self.dialog.geometry("600x500")
        self.dialog.resizable(True, True)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # 居中显示
        self.dialog.geometry("+%d+%d" % (
            self.parent.winfo_rootx() + 50,
            self.parent.winfo_rooty() + 50
        ))
        
        self._create_widgets()
        self._setup_default_mappings()
    
    def _create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.dialog.columnconfigure(0, weight=1)
        self.dialog.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(3, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="请选择两个表之间的列匹配关系", 
                               font=("Arial", 12, "bold"))
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 20))
        
        # 说明文字
        desc_label = ttk.Label(main_frame, 
                              text="左侧选择岗位表的列，右侧选择面试人员表中对应的列",
                              foreground="gray")
        desc_label.grid(row=1, column=0, columnspan=4, pady=(0, 10))
        
        # 列标题
        ttk.Label(main_frame, text="岗位表列名", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Label(main_frame, text="面试人员表列名", font=("Arial", 10, "bold")).grid(
            row=2, column=2, sticky=tk.W, padx=(10, 0))
        
        # 创建滚动框架
        self._create_mapping_area(main_frame)
        
        # 按钮区域
        self._create_buttons(main_frame)
    
    def _create_mapping_area(self, parent):
        """创建映射配置区域"""
        # 创建滚动框架
        canvas = tk.Canvas(parent, height=300)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=3, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        scrollbar.grid(row=3, column=4, sticky=(tk.N, tk.S), pady=10)
        
        parent.rowconfigure(3, weight=1)
        
        # 存储映射组件
        self.mapping_vars = []
        self.mapping_frames = []
        
        # 创建映射行
        for i in range(max(len(self.position_columns), 6)):  # 至少显示6行
            self._create_mapping_row(scrollable_frame, i)
    
    def _create_mapping_row(self, parent, row_index):
        """创建一行映射配置"""
        frame = ttk.Frame(parent)
        frame.grid(row=row_index, column=0, sticky=(tk.W, tk.E), pady=2)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(3, weight=1)
        
        # 岗位表列选择
        pos_var = tk.StringVar()
        pos_combo = ttk.Combobox(frame, textvariable=pos_var, width=25)
        pos_combo['values'] = [''] + self.position_columns
        pos_combo.grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        
        # 箭头标签
        arrow_label = ttk.Label(frame, text="→")
        arrow_label.grid(row=0, column=1, padx=5)
        
        # 面试人员表列选择
        int_var = tk.StringVar()
        int_combo = ttk.Combobox(frame, textvariable=int_var, width=25)
        int_combo['values'] = [''] + self.interview_columns
        int_combo.grid(row=0, column=2, padx=(10, 10), sticky=tk.W)
        
        # 删除按钮
        delete_btn = ttk.Button(frame, text="删除", width=8,
                               command=lambda idx=row_index: self._delete_mapping_row(idx))
        delete_btn.grid(row=0, column=3, padx=(10, 0), sticky=tk.E)
        
        # 存储变量和组件
        self.mapping_vars.append((pos_var, int_var))
        self.mapping_frames.append(frame)
        
        # 绑定变化事件，自动建议匹配
        pos_combo.bind('<<ComboboxSelected>>', 
                      lambda e, idx=row_index: self._suggest_mapping(idx))
    
    def _suggest_mapping(self, row_index):
        """根据选择的岗位表列自动建议面试表列"""
        if row_index >= len(self.mapping_vars):
            return
        
        pos_var, int_var = self.mapping_vars[row_index]
        selected_pos_col = pos_var.get()
        
        if not selected_pos_col:
            return
        
        # 查找最佳匹配
        best_match = self._find_best_match(selected_pos_col)
        if best_match and not int_var.get():  # 只在未选择时自动填充
            int_var.set(best_match)
    
    def _find_best_match(self, pos_column: str) -> Optional[str]:
        """为岗位表列找到最佳的面试表列匹配"""
        # 首先检查默认映射
        for key, suggestions in self.default_mappings.items():
            if pos_column in suggestions:
                for suggestion in suggestions:
                    if suggestion in self.interview_columns:
                        return suggestion
        
        # 如果没有找到默认映射，尝试模糊匹配
        best_match = None
        best_score = 0.6  # 最低匹配分数
        
        for int_col in self.interview_columns:
            # 计算相似度
            score = self._calculate_similarity(pos_column, int_col)
            if score > best_score:
                best_score = score
                best_match = int_col
        
        return best_match
    
    def _calculate_similarity(self, str1: str, str2: str) -> float:
        """计算两个字符串的相似度"""
        from difflib import SequenceMatcher
        return SequenceMatcher(None, str1, str2).ratio()
    
    def _delete_mapping_row(self, row_index):
        """删除映射行"""
        if row_index < len(self.mapping_vars):
            pos_var, int_var = self.mapping_vars[row_index]
            pos_var.set('')
            int_var.set('')
    
    def _setup_default_mappings(self):
        """设置默认的映射关系"""
        mapping_index = 0
        
        # 尝试设置一些常见的默认映射
        common_mappings = [
            ("职位代码", ["职位代码", "岗位代码", "代码"]),
            ("用人单位", ["用人单位", "用人司局", "招录机关", "部门名称"]),
            ("招考职位", ["招考职位", "岗考职位", "职位名称"]),
            ("部门代码", ["部门代码", "机构代码"]),
            ("姓名", ["姓名", "考生姓名"]),
            ("准考证号", ["准考证号", "考号"])
        ]
        
        for pos_col_pattern, int_col_patterns in common_mappings:
            if mapping_index >= len(self.mapping_vars):
                break
            
            # 查找岗位表中匹配的列
            pos_col = None
            for col in self.position_columns:
                if pos_col_pattern in col or col in pos_col_pattern:
                    pos_col = col
                    break
            
            if pos_col:
                # 查找面试表中匹配的列
                int_col = None
                for pattern in int_col_patterns:
                    for col in self.interview_columns:
                        if pattern in col or col in pattern:
                            int_col = col
                            break
                    if int_col:
                        break
                
                if int_col:
                    pos_var, int_var = self.mapping_vars[mapping_index]
                    pos_var.set(pos_col)
                    int_var.set(int_col)
                    mapping_index += 1
    
    def _create_buttons(self, parent):
        """创建按钮区域"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=4, column=0, columnspan=4, pady=(20, 0))
        
        # 添加映射行按钮
        add_btn = ttk.Button(button_frame, text="添加映射行", 
                            command=self._add_mapping_row)
        add_btn.grid(row=0, column=0, padx=(0, 10))
        
        # 重置按钮
        reset_btn = ttk.Button(button_frame, text="重置", 
                              command=self._reset_mappings)
        reset_btn.grid(row=0, column=1, padx=(0, 10))
        
        # 预览按钮
        preview_btn = ttk.Button(button_frame, text="预览匹配", 
                                command=self._preview_mappings)
        preview_btn.grid(row=0, column=2, padx=(0, 20))
        
        # 确定和取消按钮
        ok_btn = ttk.Button(button_frame, text="确定", 
                           command=self._on_ok)
        ok_btn.grid(row=0, column=3, padx=(0, 10))
        
        cancel_btn = ttk.Button(button_frame, text="取消", 
                               command=self._on_cancel)
        cancel_btn.grid(row=0, column=4)
    
    def _add_mapping_row(self):
        """添加新的映射行"""
        row_index = len(self.mapping_vars)
        parent = self.mapping_frames[0].master if self.mapping_frames else None
        if parent:
            self._create_mapping_row(parent, row_index)
    
    def _reset_mappings(self):
        """重置所有映射"""
        for pos_var, int_var in self.mapping_vars:
            pos_var.set('')
            int_var.set('')
        self._setup_default_mappings()
    
    def _preview_mappings(self):
        """预览映射关系"""
        mappings = self._get_mappings()
        if not mappings:
            messagebox.showwarning("警告", "请至少配置一个映射关系")
            return
        
        # 创建预览窗口
        preview_window = tk.Toplevel(self.dialog)
        preview_window.title("映射关系预览")
        preview_window.geometry("500x400")
        preview_window.transient(self.dialog)
        
        # 创建文本框显示映射关系
        text_frame = ttk.Frame(preview_window, padding="10")
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 显示映射关系
        content = "配置的列映射关系：\n\n"
        for i, (pos_col, int_col) in enumerate(mappings.items(), 1):
            content += f"{i}. {pos_col} → {int_col}\n"
        
        content += f"\n总共配置了 {len(mappings)} 个映射关系。\n"
        content += "\n匹配时将使用这些列的组合来确定两个表中的记录是否对应同一个职位。"
        
        text_widget.insert(tk.END, content)
        text_widget.config(state=tk.DISABLED)
        
        # 关闭按钮
        close_btn = ttk.Button(preview_window, text="关闭", 
                              command=preview_window.destroy)
        close_btn.pack(pady=10)
    
    def _get_mappings(self) -> Dict[str, str]:
        """获取当前配置的映射关系"""
        mappings = {}
        for pos_var, int_var in self.mapping_vars:
            pos_col = pos_var.get().strip()
            int_col = int_var.get().strip()
            if pos_col and int_col:
                mappings[pos_col] = int_col
        return mappings
    
    def _validate_mappings(self) -> Tuple[bool, str]:
        """验证映射配置"""
        mappings = self._get_mappings()
        
        if not mappings:
            return False, "请至少配置一个映射关系"
        
        # 检查重复映射
        pos_cols = list(mappings.keys())
        int_cols = list(mappings.values())
        
        if len(set(pos_cols)) != len(pos_cols):
            return False, "岗位表中有重复的列被选择"
        
        if len(set(int_cols)) != len(int_cols):
            return False, "面试人员表中有重复的列被选择"
        
        # 检查列是否存在
        for pos_col in pos_cols:
            if pos_col not in self.position_columns:
                return False, f"岗位表中不存在列: {pos_col}"
        
        for int_col in int_cols:
            if int_col not in self.interview_columns:
                return False, f"面试人员表中不存在列: {int_col}"
        
        return True, ""
    
    def _on_ok(self):
        """确定按钮处理"""
        is_valid, error_msg = self._validate_mappings()
        if not is_valid:
            messagebox.showerror("配置错误", error_msg)
            return
        
        self.result = self._get_mappings()
        self.dialog.destroy()
    
    def _on_cancel(self):
        """取消按钮处理"""
        self.result = None
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict[str, str]]:
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result


def show_column_mapping_dialog(parent, position_columns: List[str], 
                              interview_columns: List[str]) -> Optional[Dict[str, str]]:
    """
    显示列映射配置对话框
    
    Args:
        parent: 父窗口
        position_columns: 岗位表列名列表
        interview_columns: 面试人员表列名列表
        
    Returns:
        映射关系字典，格式为 {岗位表列名: 面试表列名}，如果取消则返回None
    """
    dialog = ColumnMappingDialog(parent, position_columns, interview_columns)
    return dialog.show()