#!/usr/bin/env python3
"""
快速处理版本 - 专门针对大数据量优化
跳过UI界面，直接进行高性能批量处理
"""
import sys
import os
import time
import logging
from typing import Dict, Optional

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from services.fast_data_matcher import FastDataMatcher
from services.excel_reader import ExcelReader


class FastProcessingApp:
    """快速处理应用程序"""
    
    def __init__(self):
        """初始化快速处理应用"""
        # 配置简化的日志
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler(),
                logging.FileHandler(f'fast_processing_{int(time.time())}.log', encoding='utf-8')
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.excel_reader = ExcelReader()
        
        # 默认列映射（可以根据实际情况调整）
        self.default_column_mappings = {
            '职位代码': '职位代码',
            '招考职位': '招考职位',
            '部门代码': '部门代码',
            '用人司局': '用人司局'
        }
        
        self.logger.info("快速处理应用初始化完成")
    
    def process_files_fast(self, position_file: str, interview_file: str, 
                          output_file: Optional[str] = None,
                          column_mappings: Optional[Dict[str, str]] = None) -> bool:
        """
        快速处理文件
        
        Args:
            position_file: 职位表文件路径
            interview_file: 面试人员名单文件路径
            output_file: 输出文件路径
            column_mappings: 列映射配置
            
        Returns:
            bool: 处理是否成功
        """
        start_time = time.time()
        
        try:
            self.logger.info("=" * 60)
            self.logger.info("开始快速处理模式")
            self.logger.info(f"职位表文件: {position_file}")
            self.logger.info(f"面试表文件: {interview_file}")
            
            # 使用默认列映射（如果未提供）
            if column_mappings is None:
                column_mappings = self.default_column_mappings
                self.logger.info("使用默认列映射配置")
            
            self.logger.info(f"列映射配置: {column_mappings}")
            
            # 步骤1: 快速读取数据
            self.logger.info("步骤1: 读取数据文件...")
            read_start = time.time()
            
            position_df = self.excel_reader.read_excel_sheet(position_file)
            self.logger.info(f"职位表读取完成: {len(position_df)} 行，耗时: {time.time() - read_start:.2f}秒")
            
            interview_df = self.excel_reader.read_excel_sheet(interview_file)
            self.logger.info(f"面试表读取完成: {len(interview_df)} 行，耗时: {time.time() - read_start:.2f}秒")
            
            # 步骤2: 创建快速匹配器
            self.logger.info("步骤2: 初始化快速匹配器...")
            matcher = FastDataMatcher(column_mappings)
            
            # 步骤3: 构建索引
            self.logger.info("步骤3: 构建数据索引...")
            index_start = time.time()
            matcher.build_indices(interview_df)
            self.logger.info(f"索引构建完成，耗时: {time.time() - index_start:.2f}秒")
            
            # 步骤4: 执行匹配
            self.logger.info("步骤4: 执行快速匹配...")
            match_start = time.time()
            match_result = matcher.match_data_fast(position_df, interview_df)
            match_time = time.time() - match_start
            
            # 步骤5: 生成输出文件
            if output_file is None:
                timestamp = int(time.time())
                output_file = f"岗位分数汇总_{timestamp}.xlsx"
                
                # 如果存在测试数据文件夹，将输出文件放到该文件夹中
                test_folders = [f for f in os.listdir('.') if f.startswith('测试数据_')]
                if test_folders:
                    latest_folder = max(test_folders)  # 获取最新的测试数据文件夹
                    output_file = os.path.join(latest_folder, output_file)
                    self.logger.info(f"输出文件将保存到测试数据文件夹: {latest_folder}")
            
            self.logger.info("步骤5: 导出结果...")
            export_start = time.time()
            matcher.export_results_fast(match_result['match_results'], output_file)
            export_time = time.time() - export_start
            
            # 显示统计信息
            stats = match_result['statistics']
            total_time = time.time() - start_time
            
            self.logger.info("=" * 60)
            self.logger.info("处理完成统计:")
            self.logger.info(f"总岗位数: {stats['total_positions']}")
            self.logger.info(f"匹配成功: {stats['matched_positions']} 个")
            self.logger.info(f"匹配失败: {stats['unmatched_positions']} 个")
            self.logger.info(f"匹配率: {stats['match_rate']:.2%}")
            self.logger.info(f"总面试人数: {stats['total_candidates']}")
            avg_candidates = stats.get('avg_candidates_per_position', 0)
            self.logger.info(f"平均每岗位面试人数: {avg_candidates:.1f}")
            self.logger.info("=" * 60)
            self.logger.info("性能统计:")
            self.logger.info(f"数据匹配耗时: {match_time:.2f}秒")
            self.logger.info(f"结果导出耗时: {export_time:.2f}秒")
            self.logger.info(f"总处理耗时: {total_time:.2f}秒")
            self.logger.info(f"处理速度: {stats['total_positions'] / total_time:.1f} 岗位/秒")
            self.logger.info(f"输出文件: {output_file}")
            self.logger.info("=" * 60)
            
            # 清理内存
            matcher.clear_indices()
            
            return True
            
        except Exception as e:
            error_msg = f"快速处理失败: {str(e)}"
            self.logger.error(error_msg)
            return False
    
    def run_interactive(self):
        """交互式运行模式"""
        print("=" * 60)
        print("Excel岗位分数查询工具 - 快速处理版")
        print("=" * 60)
        
        try:
            # 获取文件路径
            position_file = input("请输入职位表文件路径: ").strip()
            if not position_file:
                print("错误: 职位表文件路径不能为空")
                return
            
            interview_file = input("请输入面试人员名单文件路径: ").strip()
            if not interview_file:
                print("错误: 面试人员名单文件路径不能为空")
                return
            
            output_file = input("请输入输出文件路径（回车使用默认）: ").strip()
            if not output_file:
                output_file = None
            
            # 询问是否使用自定义列映射
            use_custom = input("是否使用自定义列映射？(y/N): ").strip().lower()
            column_mappings = None
            
            if use_custom == 'y':
                print("请输入列映射配置（格式：岗位表列名=面试表列名，多个用逗号分隔）:")
                mapping_input = input().strip()
                if mapping_input:
                    try:
                        column_mappings = {}
                        for pair in mapping_input.split(','):
                            pos_col, int_col = pair.split('=')
                            column_mappings[pos_col.strip()] = int_col.strip()
                    except Exception as e:
                        print(f"列映射配置格式错误: {e}")
                        return
            
            # 开始处理
            print("\n开始快速处理...")
            success = self.process_files_fast(position_file, interview_file, output_file, column_mappings)
            
            if success:
                print("\n✅ 处理完成！请查看输出文件。")
            else:
                print("\n❌ 处理失败，请查看日志文件了解详情。")
                
        except KeyboardInterrupt:
            print("\n用户中断处理")
        except Exception as e:
            print(f"\n运行时发生错误: {e}")


def main():
    """主函数"""
    # 检查帮助参数
    if len(sys.argv) > 1 and sys.argv[1] in ['-h', '--help', 'help']:
        print("Excel岗位分数查询工具 - 快速处理版")
        print("=" * 50)
        print("用法:")
        print("  交互式模式:")
        print("    python fast_processing_main.py")
        print()
        print("  命令行模式:")
        print("    python fast_processing_main.py <职位表文件> <面试表文件> [输出文件]")
        print()
        print("示例:")
        print("    python fast_processing_main.py 工作簿1.xlsx 面试人员名单表.xlsx")
        print("    python fast_processing_main.py 工作簿1.xlsx 面试人员名单表.xlsx 结果.xlsx")
        print()
        print("特点:")
        print("  - 针对大数据量优化（支持10万+记录）")
        print("  - 使用预构建索引，处理速度快5-10倍")
        print("  - 自动生成处理日志")
        print("  - 支持自定义列映射")
        return
    
    app = FastProcessingApp()
    
    # 检查命令行参数
    if len(sys.argv) >= 3:
        # 命令行模式
        position_file = sys.argv[1]
        interview_file = sys.argv[2]
        output_file = sys.argv[3] if len(sys.argv) > 3 else None
        
        print(f"命令行模式: 处理 {position_file} 和 {interview_file}")
        success = app.process_files_fast(position_file, interview_file, output_file)
        sys.exit(0 if success else 1)
    else:
        # 交互式模式
        app.run_interactive()


if __name__ == "__main__":
    main()