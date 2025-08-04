#!/usr/bin/env python3
"""
性能对比测试脚本
比较原版和优化版的处理性能
"""
import sys
import os
import time
import logging
from typing import Dict, List

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from services.processing_engine import ProcessingEngine
from services.optimized_processing_engine import OptimizedProcessingEngine
from services.fast_data_matcher import FastDataMatcher
from services.excel_reader import ExcelReader


class PerformanceTest:
    """性能测试类"""
    
    def __init__(self):
        """初始化性能测试"""
        # 配置日志
        logging.basicConfig(
            level=logging.WARNING,  # 减少日志输出
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
        self.logger = logging.getLogger(__name__)
        self.excel_reader = ExcelReader()
        
        # 测试配置
        self.column_mappings = {
            '职位代码': '职位代码',
            '招考职位': '招考职位',
            '部门代码': '部门代码',
            '用人司局': '用人司局'
        }
    
    def test_original_engine(self, position_file: str, interview_file: str) -> Dict:
        """测试原版处理引擎"""
        print("测试原版处理引擎...")
        start_time = time.time()
        
        try:
            engine = ProcessingEngine()
            
            # 临时输出文件
            temp_output = f"temp_original_{int(time.time())}.xlsx"
            
            result = engine.process_files(
                position_file=position_file,
                interview_file=interview_file,
                column_mappings=self.column_mappings,
                output_path=temp_output
            )
            
            processing_time = time.time() - start_time
            
            # 清理临时文件
            try:
                if os.path.exists(temp_output):
                    os.remove(temp_output)
            except:
                pass
            
            return {
                'success': result.success,
                'processing_time': processing_time,
                'total_positions': len(result.results) if result.success else 0,
                'matched_positions': len([r for r in result.results if r.min_score > 0]) if result.success else 0,
                'error': None if result.success else result.message
            }
            
        except Exception as e:
            return {
                'success': False,
                'processing_time': time.time() - start_time,
                'total_positions': 0,
                'matched_positions': 0,
                'error': str(e)
            }
    
    def test_optimized_engine(self, position_file: str, interview_file: str) -> Dict:
        """测试优化版处理引擎"""
        print("测试优化版处理引擎...")
        start_time = time.time()
        
        try:
            engine = OptimizedProcessingEngine()
            
            # 临时输出文件
            temp_output = f"temp_optimized_{int(time.time())}.xlsx"
            
            result = engine.process_files_optimized(
                position_file=position_file,
                interview_file=interview_file,
                column_mappings=self.column_mappings,
                output_path=temp_output
            )
            
            processing_time = time.time() - start_time
            
            # 清理临时文件
            try:
                if os.path.exists(temp_output):
                    os.remove(temp_output)
            except:
                pass
            
            # 清理缓存
            engine.clear_cache()
            
            return {
                'success': result['success'],
                'processing_time': processing_time,
                'total_positions': result.get('total_positions', 0),
                'matched_positions': result.get('matched_positions', 0),
                'error': None if result['success'] else result.get('message', 'Unknown error')
            }
            
        except Exception as e:
            return {
                'success': False,
                'processing_time': time.time() - start_time,
                'total_positions': 0,
                'matched_positions': 0,
                'error': str(e)
            }
    
    def test_fast_matcher(self, position_file: str, interview_file: str) -> Dict:
        """测试快速匹配器"""
        print("测试快速匹配器...")
        start_time = time.time()
        
        try:
            # 读取数据
            position_df = self.excel_reader.read_excel_sheet(position_file)
            interview_df = self.excel_reader.read_excel_sheet(interview_file)
            
            # 创建快速匹配器
            matcher = FastDataMatcher(self.column_mappings)
            
            # 执行匹配
            result = matcher.match_data_fast(position_df, interview_df)
            
            processing_time = time.time() - start_time
            
            # 清理内存
            matcher.clear_indices()
            
            stats = result['statistics']
            
            return {
                'success': True,
                'processing_time': processing_time,
                'total_positions': stats['total_positions'],
                'matched_positions': stats['matched_positions'],
                'error': None
            }
            
        except Exception as e:
            return {
                'success': False,
                'processing_time': time.time() - start_time,
                'total_positions': 0,
                'matched_positions': 0,
                'error': str(e)
            }
    
    def run_performance_test(self, position_file: str, interview_file: str):
        """运行性能对比测试"""
        print("=" * 80)
        print("Excel岗位分数查询工具 - 性能对比测试")
        print("=" * 80)
        print(f"职位表文件: {position_file}")
        print(f"面试表文件: {interview_file}")
        print("=" * 80)
        
        # 验证文件存在
        if not os.path.exists(position_file):
            print(f"错误: 职位表文件不存在: {position_file}")
            return
        
        if not os.path.exists(interview_file):
            print(f"错误: 面试表文件不存在: {interview_file}")
            return
        
        # 获取文件大小信息
        pos_size = os.path.getsize(position_file) / 1024 / 1024  # MB
        int_size = os.path.getsize(interview_file) / 1024 / 1024  # MB
        
        print(f"职位表文件大小: {pos_size:.2f} MB")
        print(f"面试表文件大小: {int_size:.2f} MB")
        print()
        
        # 存储测试结果
        results = {}
        
        # 测试1: 快速匹配器（最快）
        print("🚀 测试1: 快速匹配器")
        results['fast_matcher'] = self.test_fast_matcher(position_file, interview_file)
        self._print_test_result("快速匹配器", results['fast_matcher'])
        print()
        
        # 测试2: 优化版引擎
        print("⚡ 测试2: 优化版处理引擎")
        results['optimized_engine'] = self.test_optimized_engine(position_file, interview_file)
        self._print_test_result("优化版引擎", results['optimized_engine'])
        print()
        
        # 测试3: 原版引擎（可能很慢，可选）
        print("🐌 测试3: 原版处理引擎")
        user_input = input("原版引擎可能很慢，是否继续测试？(y/N): ").strip().lower()
        
        if user_input == 'y':
            results['original_engine'] = self.test_original_engine(position_file, interview_file)
            self._print_test_result("原版引擎", results['original_engine'])
        else:
            print("跳过原版引擎测试")
            results['original_engine'] = None
        
        print()
        
        # 性能对比总结
        self._print_performance_summary(results)
    
    def _print_test_result(self, name: str, result: Dict):
        """打印测试结果"""
        if result['success']:
            print(f"✅ {name} - 成功")
            print(f"   处理时间: {result['processing_time']:.2f}秒")
            print(f"   总岗位数: {result['total_positions']}")
            print(f"   匹配成功: {result['matched_positions']}")
            if result['total_positions'] > 0:
                match_rate = result['matched_positions'] / result['total_positions']
                print(f"   匹配率: {match_rate:.2%}")
                speed = result['total_positions'] / result['processing_time']
                print(f"   处理速度: {speed:.1f} 岗位/秒")
        else:
            print(f"❌ {name} - 失败")
            print(f"   处理时间: {result['processing_time']:.2f}秒")
            print(f"   错误信息: {result['error']}")
    
    def _print_performance_summary(self, results: Dict):
        """打印性能对比总结"""
        print("=" * 80)
        print("性能对比总结")
        print("=" * 80)
        
        # 收集成功的测试结果
        successful_results = {k: v for k, v in results.items() if v and v['success']}
        
        if not successful_results:
            print("没有成功的测试结果")
            return
        
        # 按处理时间排序
        sorted_results = sorted(successful_results.items(), key=lambda x: x[1]['processing_time'])
        
        print("性能排名（按处理时间）:")
        for i, (name, result) in enumerate(sorted_results, 1):
            time_taken = result['processing_time']
            speed = result['total_positions'] / time_taken if time_taken > 0 else 0
            
            print(f"{i}. {name}")
            print(f"   处理时间: {time_taken:.2f}秒")
            print(f"   处理速度: {speed:.1f} 岗位/秒")
            
            # 计算相对于最快的倍数
            if i > 1:
                fastest_time = sorted_results[0][1]['processing_time']
                speedup = time_taken / fastest_time
                print(f"   相对最快: {speedup:.1f}x 时间")
            
            print()
        
        # 性能提升总结
        if len(sorted_results) >= 2:
            fastest = sorted_results[0]
            slowest = sorted_results[-1]
            
            improvement = slowest[1]['processing_time'] / fastest[1]['processing_time']
            
            print(f"🎯 性能提升总结:")
            print(f"   最快方案: {fastest[0]}")
            print(f"   最慢方案: {slowest[0]}")
            print(f"   性能提升: {improvement:.1f}x 倍")
            
            time_saved = slowest[1]['processing_time'] - fastest[1]['processing_time']
            print(f"   节省时间: {time_saved:.2f}秒")


def main():
    """主函数"""
    tester = PerformanceTest()
    
    if len(sys.argv) >= 3:
        # 命令行模式
        position_file = sys.argv[1]
        interview_file = sys.argv[2]
        tester.run_performance_test(position_file, interview_file)
    else:
        # 交互式模式
        print("Excel岗位分数查询工具 - 性能测试")
        print("请输入测试文件路径:")
        
        position_file = input("职位表文件路径: ").strip()
        interview_file = input("面试人员名单文件路径: ").strip()
        
        if position_file and interview_file:
            tester.run_performance_test(position_file, interview_file)
        else:
            print("错误: 文件路径不能为空")


if __name__ == "__main__":
    main()