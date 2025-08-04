#!/usr/bin/env python3
"""
Excel岗位分数查询工具 - 主程序
整合了成功的日志解析逻辑
"""
import sys
import os
import time
import threading
import re
import pandas as pd
import logging
from typing import List, Dict, Any, Optional

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.main_window import MainWindow
from ui.column_mapping_dialog import ColumnMappingDialog
from services.processing_engine import ProcessingEngine
from services.optimized_processing_engine import OptimizedProcessingEngine
from services.fast_data_matcher import FastDataMatcher
from services.excel_reader import ExcelReader
from utils.logger import default_logger


class ExcelPositionQueryApp:
    """Excel岗位分数查询应用程序主类"""
    
    def __init__(self):
        """初始化应用程序"""
        self.main_window = MainWindow()
        self.excel_reader = ExcelReader()
        self.processing_engine = ProcessingEngine()
        self.optimized_engine = OptimizedProcessingEngine()
        
        # 绑定事件处理函数
        self.main_window.on_position_file_select = self.on_position_file_select
        self.main_window.on_interview_file_select = self.on_interview_file_select
        self.main_window.on_start_processing = self.on_start_processing
        self.main_window.on_configure_columns = self.on_configure_columns
        
        # 文件路径
        self.position_file_path: Optional[str] = None
        self.interview_file_path: Optional[str] = None
        
        # 性能优化选项
        self.use_optimized_engine = True  # 默认使用优化引擎
        
        default_logger.info("Excel岗位分数查询工具启动")
    
    def on_position_file_select(self, file_path: str):
        """处理职位表文件选择"""
        try:
            self.position_file_path = file_path
            # 验证文件格式
            columns = self.excel_reader.get_column_names(file_path)
            self.main_window._update_status(f"职位表包含 {len(columns)} 列")
            default_logger.info(f"职位表文件选择成功: {file_path}")
        except Exception as e:
            error_msg = f"读取职位表文件失败: {str(e)}"
            self.main_window.show_error(error_msg)
            default_logger.error(error_msg)
    
    def on_interview_file_select(self, file_path: str):
        """处理面试人员名单文件选择"""
        try:
            self.interview_file_path = file_path
            # 验证文件格式
            columns = self.excel_reader.get_column_names(file_path)
            self.main_window._update_status(f"面试人员名单包含 {len(columns)} 列")
            default_logger.info(f"面试人员名单文件选择成功: {file_path}")
        except Exception as e:
            error_msg = f"读取面试人员名单文件失败: {str(e)}"
            self.main_window.show_error(error_msg)
            default_logger.error(error_msg)
    
    def on_configure_columns(self):
        """处理列映射配置"""
        try:
            if not self.position_file_path or not self.interview_file_path:
                self.main_window.show_error("请先选择两个Excel文件")
                return
            
            # 获取两个文件的列名
            position_columns = self.excel_reader.get_column_names(self.position_file_path)
            interview_columns = self.excel_reader.get_column_names(self.interview_file_path)
            
            # 打开列映射配置对话框
            dialog = ColumnMappingDialog(
                self.main_window.root,
                position_columns,
                interview_columns
            )
            
            # 设置当前的映射配置（如果有）
            current_mappings = self.main_window.get_column_mappings()
            if current_mappings:
                dialog.set_mappings(current_mappings)
            
            # 显示对话框并获取结果
            mappings = dialog.show()
            if mappings:
                self.main_window.set_column_mappings(mappings)
                default_logger.info(f"列映射配置完成: {mappings}")
            
        except Exception as e:
            error_msg = f"配置列映射时发生错误: {str(e)}"
            self.main_window.show_error(error_msg)
            default_logger.error(error_msg)
    
    def on_start_processing(self):
        """开始处理数据"""
        try:
            # 验证必要条件
            if not self.position_file_path or not self.interview_file_path:
                self.main_window.show_error("请先选择两个Excel文件")
                return
            
            column_mappings = self.main_window.get_column_mappings()
            if not column_mappings:
                self.main_window.show_error("请先配置列映射关系")
                return
            
            # 获取输出列选择
            selected_output_columns = self.main_window.get_selected_output_columns()
            
            # 设置处理状态
            self.main_window.set_processing_state(True)
            self.main_window.clear_status()
            self.main_window._update_status("开始处理数据...")
            
            # 在后台线程中执行处理
            processing_thread = threading.Thread(
                target=self._process_data_with_fallback,
                args=(column_mappings, selected_output_columns),
                daemon=True
            )
            processing_thread.start()
            
        except Exception as e:
            error_msg = f"启动处理时发生错误: {str(e)}"
            self.main_window.show_processing_error(error_msg)
            self.main_window.set_processing_state(False)
            default_logger.error(error_msg)
    
    def _process_data_with_fallback(self, column_mappings: dict, selected_output_columns: Optional[List[str]] = None):
        """
        使用优化的处理流程或回退到验证过的日志方法
        """
        try:
            if self.use_optimized_engine:
                self.main_window._update_status("使用优化版高性能处理引擎...")
                success = self._process_with_optimized_engine(column_mappings, selected_output_columns)
            else:
                self.main_window._update_status("使用验证过的正确处理流程...")
                success = self._process_with_log_method(column_mappings, selected_output_columns)
            
            if not success:
                self.main_window.show_processing_error(
                    "处理失败",
                    "请检查文件格式和数据完整性"
                )
            
        except Exception as e:
            error_msg = f"处理过程中发生异常: {str(e)}"
            self.main_window.show_processing_error(error_msg)
            default_logger.error(error_msg)
        finally:
            self.main_window.set_processing_state(False)
    
    def _process_with_log_method(self, column_mappings: dict, selected_output_columns: Optional[List[str]] = None) -> bool:
        """使用验证过的日志处理方法"""
        try:
            # 第一步：生成处理日志
            self.main_window._update_status("第一步：生成处理日志...")
            self.main_window.show_progress(10, "开始生成处理日志")
            
            log_file = self._generate_processing_log(column_mappings)
            if not log_file:
                return False
            
            # 第二步：解析日志数据
            self.main_window._update_status("第二步：解析日志数据...")
            self.main_window.show_progress(70, "解析处理日志")
            
            positions_data = self._parse_log_data(log_file)
            if not positions_data:
                self.main_window._update_status("❌ 未能从日志中解析到有效数据")
                return False
            
            # 第三步：生成Excel报告
            self.main_window._update_status("第三步：生成Excel报告...")
            self.main_window.show_progress(90, "生成Excel报告")
            
            timestamp = int(time.time())
            output_file = f"岗位分数汇总_{timestamp}.xlsx"
            
            # 如果存在测试数据文件夹，将输出文件放到该文件夹中
            test_folders = [f for f in os.listdir('.') if f.startswith('测试数据_')]
            if test_folders:
                latest_folder = max(test_folders)  # 获取最新的测试数据文件夹
                output_file = os.path.join(latest_folder, output_file)
                default_logger.info(f"输出文件将保存到测试数据文件夹: {latest_folder}")
            self._create_excel_from_log_data(positions_data, output_file, selected_output_columns)
            
            # 显示成功结果
            self.main_window.show_progress(100, "处理完成")
            
            results = {
                'total_positions': len(positions_data),
                'processed_positions': len(positions_data),
                'failed_positions': 0,
                'total_candidates': sum(pos['candidate_count'] for pos in positions_data),
                'output_file': output_file,
                'processing_time': 0,  # 总时间在日志中
                'errors': [],
                'warnings': ['使用了验证过的日志解析方法生成结果']
            }
            
            self.main_window.show_results(results)
            
            # 清理临时日志文件
            try:
                if os.path.exists(log_file):
                    os.remove(log_file)
            except:
                pass
            
            return True
            
        except Exception as e:
            self.main_window._update_status(f"日志处理方法异常: {str(e)}")
            return False
    
    def _generate_processing_log(self, column_mappings: dict) -> str:
        """生成处理日志文件"""
        try:
            # 生成唯一的日志文件名
            timestamp = int(time.time())
            log_filename = f'temp_processing_{timestamp}.log'
            
            self.main_window._update_status(f"生成处理日志: {log_filename}")
            
            # 配置日志 - 与成功案例完全相同
            import logging
            
            # 清除现有的日志配置
            for handler in logging.root.handlers[:]:
                logging.root.removeHandler(handler)
            
            # 重新配置日志
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(log_filename, encoding='utf-8'),
                ],
                force=True
            )
            
            # 执行处理以生成日志
            self.main_window._update_status("执行处理以生成详细日志...")
            
            # 创建处理引擎
            processing_engine = ProcessingEngine()
            
            # 生成临时输出文件名
            temp_output = f"temp_output_{timestamp}.xlsx"
            
            # 执行处理
            result = processing_engine.process_files(
                position_file=self.position_file_path,
                interview_file=self.interview_file_path,
                column_mappings=column_mappings,
                output_path=temp_output
            )
            
            # 清理临时输出文件
            try:
                if os.path.exists(temp_output):
                    os.remove(temp_output)
            except:
                pass
            
            # 检查日志文件是否生成
            if os.path.exists(log_filename):
                file_size = os.path.getsize(log_filename)
                self.main_window._update_status(f"✅ 处理日志生成成功: {file_size:,} 字节")
                return log_filename
            else:
                self.main_window._update_status("❌ 处理日志生成失败")
                return None
                
        except Exception as e:
            self.main_window._update_status(f"生成处理日志时发生错误: {str(e)}")
            return None


    
    def _parse_log_data(self, log_file: str) -> List[Dict[str, Any]]:
        """解析日志文件，提取岗位和分数数据"""
        positions_data = []
        current_position = None
        candidates = []
        
        try:
            self.main_window._update_status(f"开始解析日志文件: {log_file}")
            
            with open(log_file, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            for line in lines:
                line = line.strip()
                
                # 匹配岗位开始标记
                if '=== 开始匹配岗位 ===' in line:
                    # 保存之前的岗位数据
                    if current_position and candidates:
                        min_score = min(float(c['score']) for c in candidates)
                        positions_data.append({
                            'position_name': current_position['name'],
                            'position_code': current_position['code'],
                            'department': current_position['department'],
                            'min_score': min_score,
                            'candidate_count': len(candidates),
                            'candidates': candidates.copy()
                        })
                    
                    # 重置当前岗位数据
                    current_position = {'name': '', 'code': '', 'department': ''}
                    candidates = []
                    continue
                
                # 匹配岗位信息
                if current_position is not None:
                    position_name_match = re.search(r'岗位名称: (.+)$', line)
                    if position_name_match:
                        current_position['name'] = position_name_match.group(1)
                        continue
                    
                    department_match = re.search(r'用人司局: (.+)$', line)
                    if department_match:
                        current_position['department'] = department_match.group(1)
                        continue
                    
                    code_match = re.search(r'职位代码: (.+)$', line)
                    if code_match:
                        current_position['code'] = code_match.group(1)
                        continue
                
                # 匹配候选人数据
                candidate_match = re.search(r'(\d+): (.+?) - 分数: ([\d.]+) - 司局: (.+?) - 岗位: (.+?)$', line)
                if candidate_match and current_position:
                    candidates.append({
                        'rank': int(candidate_match.group(1)),
                        'name': candidate_match.group(2),
                        'score': float(candidate_match.group(3)),
                        'department': candidate_match.group(4),
                        'position': candidate_match.group(5)
                    })
                    continue
                
                # 匹配成功完成的岗位
                completion_match = re.search(r"岗位 '(.+?)' 匹配完成: 成功", line)
                if completion_match and current_position and candidates:
                    if completion_match.group(1) == current_position['name']:
                        min_score = min(float(c['score']) for c in candidates)
                        positions_data.append({
                            'position_name': current_position['name'],
                            'position_code': current_position['code'],
                            'department': current_position['department'],
                            'min_score': min_score,
                            'candidate_count': len(candidates),
                            'candidates': candidates.copy()
                        })
                        current_position = None
                        candidates = []
            
            # 处理最后一个岗位
            if current_position and candidates and current_position['name']:
                min_score = min(float(c['score']) for c in candidates)
                positions_data.append({
                    'position_name': current_position['name'],
                    'position_code': current_position['code'],
                    'department': current_position['department'],
                    'min_score': min_score,
                    'candidate_count': len(candidates),
                    'candidates': candidates.copy()
                })
            
            self.main_window._update_status(f"✅ 日志解析完成，共找到 {len(positions_data)} 个岗位")
            
        except Exception as e:
            self.main_window._update_status(f"❌ 解析日志文件时发生错误: {str(e)}")
            return []
        
        return positions_data
    
    def _create_excel_from_log_data(self, positions_data: List[Dict[str, Any]], output_file: str, 
                                   selected_columns: Optional[List[str]] = None) -> str:
        """根据日志数据创建Excel文件"""
        try:
            self.main_window._update_status(f"开始创建Excel文件: {output_file}")
            
            # 创建汇总数据
            summary_data = []
            
            for pos_data in positions_data:
                # 构建完整的数据行
                full_row_data = {
                    '招考职位': pos_data['position_name'],
                    '职位代码': str(pos_data['position_code']),
                    '用人司局': pos_data['department'],
                    '最低面试分数': pos_data['min_score'],
                    '面试人数': pos_data['candidate_count'],
                    '状态': '成功'
                }
                
                # 确保职位代码以文本格式保存
                position_code = str(pos_data['position_code'])
                if position_code and position_code != 'nan':
                    if position_code.replace('.', '').replace('E', '').replace('+', '').replace('-', '').isdigit():
                        position_code = "'" + position_code
                full_row_data['职位代码'] = position_code
                
                # 根据用户选择的列过滤数据
                if selected_columns:
                    filtered_row = {}
                    for col in selected_columns:
                        filtered_row[col] = full_row_data.get(col, '')
                    summary_data.append(filtered_row)
                else:
                    # 使用默认列
                    summary_data.append(full_row_data)
            
            # 创建DataFrame并保存
            df = pd.DataFrame(summary_data)
            
            # 确保职位代码列为文本类型
            df['职位代码'] = df['职位代码'].astype(str)
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='岗位分数汇总', index=False)
                
                # 获取工作表对象，设置职位代码列为文本格式
                worksheet = writer.sheets['岗位分数汇总']
                
                # 设置职位代码列为文本格式
                from openpyxl.utils import get_column_letter
                
                # 找到职位代码列的位置
                headers = list(df.columns)
                if '职位代码' in headers:
                    col_idx = headers.index('职位代码') + 1  # Excel列从1开始
                    col_letter = get_column_letter(col_idx)
                    
                    # 设置整列为文本格式
                    for row in range(1, len(df) + 2):  # +2因为包含标题行
                        cell = worksheet[f'{col_letter}{row}']
                        cell.number_format = '@'  # '@'表示文本格式
                        
                        # 如果是数据行且包含职位代码，确保显示为完整数字
                        if row > 1:  # 跳过标题行
                            cell_value = str(cell.value)
                            if cell_value.startswith("'"):
                                cell.value = cell_value[1:]  # 移除前导单引号
                
                # 自动调整列宽
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 30)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # 验证文件是否成功创建
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                self.main_window._update_status(f"✅ Excel文件创建成功: {file_size:,} 字节")
            else:
                raise Exception("Excel文件创建失败")
            
            return output_file
            
        except Exception as e:
            raise Exception(f"创建Excel文件时发生错误: {str(e)}")
    
    def _process_with_optimized_engine(self, column_mappings: dict, selected_output_columns: Optional[List[str]] = None) -> bool:
        """使用优化版处理引擎"""
        try:
            start_time = time.time()
            
            # 设置进度回调
            def progress_callback(message, percentage):
                self.main_window._update_status(message)
                self.main_window.show_progress(percentage, message)
            
            self.optimized_engine.progress_callback = progress_callback
            
            # 生成输出文件名
            timestamp = int(time.time())
            output_file = f"岗位分数汇总_{timestamp}.xlsx"
            
            # 如果存在测试数据文件夹，将输出文件放到该文件夹中
            test_folders = [f for f in os.listdir('.') if f.startswith('测试数据_')]
            if test_folders:
                latest_folder = max(test_folders)  # 获取最新的测试数据文件夹
                output_file = os.path.join(latest_folder, output_file)
                default_logger.info(f"输出文件将保存到测试数据文件夹: {latest_folder}")
            
            # 如果存在测试数据文件夹，将输出文件放到该文件夹中
            test_folders = [f for f in os.listdir('.') if f.startswith('测试数据_')]
            if test_folders:
                latest_folder = max(test_folders)  # 获取最新的测试数据文件夹
                output_file = os.path.join(latest_folder, output_file)
                default_logger.info(f"输出文件将保存到测试数据文件夹: {latest_folder}")
            
            # 执行优化版处理
            result = self.optimized_engine.process_files_optimized(
                position_file=self.position_file_path,
                interview_file=self.interview_file_path,
                column_mappings=column_mappings,
                output_path=output_file,
                selected_columns=selected_output_columns
            )
            
            if result['success']:
                processing_time = time.time() - start_time
                
                # 显示成功结果
                self.main_window.show_progress(100, "处理完成")
                
                results = {
                    'total_positions': result['total_positions'],
                    'processed_positions': result['matched_positions'],
                    'failed_positions': result['unmatched_positions'],
                    'total_candidates': result['total_candidates'],
                    'output_file': output_file,
                    'processing_time': processing_time,
                    'errors': [],
                    'warnings': [f'使用优化版引擎处理，匹配率: {result["match_rate"]:.1%}']
                }
                
                self.main_window.show_results(results)
                return True
            else:
                self.main_window._update_status(f"优化版处理失败: {result['message']}")
                # 回退到日志方法
                self.main_window._update_status("回退到验证过的日志处理方法...")
                return self._process_with_log_method(column_mappings, selected_output_columns)
                
        except Exception as e:
            self.main_window._update_status(f"优化版处理异常: {str(e)}")
            # 回退到日志方法
            self.main_window._update_status("回退到验证过的日志处理方法...")
            return self._process_with_log_method(column_mappings, selected_output_columns)
    
    def run(self):
        """运行应用程序"""
        try:
            self.main_window.run()
        except KeyboardInterrupt:
            default_logger.info("用户中断程序")
        except Exception as e:
            error_msg = f"应用程序运行时发生错误: {str(e)}"
            default_logger.error(error_msg)
            print(error_msg)
        finally:
            default_logger.info("应用程序正常退出")


def main():
    """主函数"""
    try:
        app = ExcelPositionQueryApp()
        app.run()
    except Exception as e:
        print(f"启动应用程序时发生错误: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()