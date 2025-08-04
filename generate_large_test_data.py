import pandas as pd
import random
import os
from datetime import datetime

def create_test_data_folder():
    """创建测试数据文件夹"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = f"测试数据_{timestamp}"
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    return folder_name

def generate_random_positions(count=200):
    """生成随机岗位表数据"""
    
    # 部门数据池
    departments = [
        ('财政部', '101'), ('教育部', '102'), ('工业和信息化部', '103'), ('交通运输部', '104'),
        ('国家卫生健康委员会', '105'), ('生态环境部', '106'), ('商务部', '107'), ('文化和旅游部', '108'),
        ('农业农村部', '109'), ('人力资源和社会保障部', '110'), ('住房和城乡建设部', '111'), ('自然资源部', '112'),
        ('应急管理部', '113'), ('国家市场监督管理总局', '114'), ('国家税务总局', '115'), ('国家发展和改革委员会', '116'),
        ('科学技术部', '117'), ('国防部', '118'), ('公安部', '119'), ('民政部', '120'),
        ('司法部', '121'), ('水利部', '122'), ('国家统计局', '123'), ('国家知识产权局', '124'),
        ('国家能源局', '125'), ('国家邮政局', '126'), ('国家外汇管理局', '127'), ('中国银行保险监督管理委员会', '128'),
        ('中国证券监督管理委员会', '129'), ('国家广播电视总局', '130')
    ]
    
    # 司局后缀
    bureau_suffixes = [
        '办公厅', '政策法规司', '综合司', '规划司', '财务司', '人事司', '监察司', '审计司',
        '国际合作司', '科技司', '信息中心', '服务中心', '管理局', '监督局', '执法局', '发展司',
        '建设司', '运营司', '技术司', '质量司', '安全司', '环保司', '宣传司', '教育司'
    ]
    
    # 职位级别
    position_levels = [
        '一级主任科员', '二级主任科员', '三级主任科员', '四级主任科员',
        '一级科员', '二级科员', '三级科员', '四级科员',
        '副处长', '处长', '副司长', '司长'
    ]
    
    # 专业领域
    majors = [
        '法学、政治学、行政管理', '经济学、金融学、会计学', '计算机科学与技术、软件工程、信息安全',
        '土木工程、建筑学、城乡规划', '机械工程、电气工程、自动化', '化学工程、材料科学、环境工程',
        '医学、生物学、药学', '教育学、心理学、社会学', '新闻传播学、中文、外语',
        '数学、统计学、物理学', '地理学、测绘工程、地质学', '农学、林学、畜牧学',
        '管理科学、工商管理、公共管理', '艺术学、设计学、音乐学', '历史学、哲学、考古学'
    ]
    
    # 学历要求
    education_levels = ['本科及以上', '硕士研究生及以上', '博士研究生及以上']
    degree_levels = ['学士', '硕士', '博士']
    
    # 工作地点
    work_locations = [
        '北京市', '上海市', '广州市', '深圳市', '天津市', '重庆市', '成都市', '西安市',
        '武汉市', '南京市', '杭州市', '青岛市', '大连市', '宁波市', '厦门市'
    ]
    
    positions_data = []
    
    for i in range(count):
        # 随机选择部门
        dept_name, dept_code = random.choice(departments)
        
        # 生成司局名称
        bureau_name = f"{dept_name}{random.choice(bureau_suffixes)}"
        
        # 生成职位名称
        position_name = f"{bureau_name}{random.choice(['管理', '技术', '综合', '专项', '执法', '监督'])}{random.choice(position_levels)}"
        
        # 生成职位代码 (确保唯一性)
        position_code = f"4001{dept_code}{i+1:04d}"
        
        positions_data.append({
            '招考职位': position_name,
            '职位代码': position_code,
            '用人司局': bureau_name,
            '部门代码': dept_code,
            '部门名称': dept_name,
            '招考人数': random.randint(1, 5),
            '专业': random.choice(majors),
            '学历': random.choice(education_levels),
            '学位': random.choice(degree_levels),
            '工作地点': random.choice(work_locations),
            '咨询电话1': f"010-{random.randint(60000000, 99999999)}",
            '咨询电话2': f"010-{random.randint(60000000, 99999999)}",
            '部门网站': f"http://www.{dept_name.replace('国家', '').replace('委员会', '').replace('总局', '').lower()}.gov.cn"
        })
    
    return pd.DataFrame(positions_data)

def generate_random_interview_data(positions_df, interview_count=2000):
    """生成随机面试人员数据"""
    
    # 姓氏和名字池
    surnames = [
        '王', '李', '张', '刘', '陈', '杨', '赵', '黄', '周', '吴', '徐', '孙', '胡', '朱', '高',
        '林', '何', '郭', '马', '罗', '梁', '宋', '郑', '谢', '韩', '唐', '冯', '于', '董', '萧',
        '程', '曹', '袁', '邓', '许', '傅', '沈', '曾', '彭', '吕', '苏', '卢', '蒋', '蔡', '贾',
        '丁', '魏', '薛', '叶', '阎', '余', '潘', '杜', '戴', '夏', '钟', '汪', '田', '任', '姜'
    ]
    
    given_names = [
        '伟', '芳', '娜', '敏', '静', '丽', '强', '磊', '军', '洋', '勇', '艳', '杰', '娟', '涛',
        '明', '超', '秀', '英', '华', '慧', '巧', '美', '娜', '静', '淑', '惠', '珠', '翠', '雅',
        '芝', '玉', '萍', '红', '娥', '玲', '菊', '兰', '凤', '洁', '梅', '琳', '素', '云', '莲',
        '真', '环', '雪', '荣', '爱', '妹', '霞', '香', '月', '莺', '媛', '艳', '瑞', '凡', '佳'
    ]
    
    interview_data = []
    position_codes = positions_df['职位代码'].tolist()
    
    # 确保每个职位都有面试人员
    for _, pos_row in positions_df.iterrows():
        position_code = pos_row['职位代码']
        position_name = pos_row['招考职位']
        department_name = pos_row['部门名称']
        department_code = pos_row['部门代码']
        bureau_name = pos_row['用人司局']
        recruit_count = pos_row['招考人数']
        
        # 每个职位至少生成 recruit_count + 1 到 recruit_count + 5 个面试人员
        interview_count_for_position = random.randint(recruit_count + 1, recruit_count + 8)
        
        for j in range(interview_count_for_position):
            # 生成随机姓名
            surname = random.choice(surnames)
            given_name = random.choice(given_names)
            if random.random() > 0.7:  # 30%概率生成双字名
                given_name += random.choice(given_names)
            full_name = surname + given_name
            
            # 生成准考证号
            exam_number = f"2024{position_code[-6:]}{j+1:03d}"
            
            # 生成最低面试分数
            min_score = round(random.uniform(60.0, 85.0), 1)
            
            interview_data.append({
                '准考证号': exam_number,
                '姓名': full_name,
                '招录机关': department_name,
                '部门代码': department_code,
                '用人司局': bureau_name,
                '招考职位': position_name,
                '职位代码': position_code,
                '最低面试分数': min_score
            })
    
    # 生成额外的不匹配数据（约占总数的30-40%）
    current_count = len(interview_data)
    additional_count = interview_count - current_count
    
    if additional_count > 0:
        for i in range(additional_count):
            # 生成随机姓名
            surname = random.choice(surnames)
            given_name = random.choice(given_names)
            if random.random() > 0.7:
                given_name += random.choice(given_names)
            full_name = surname + given_name
            
            # 生成不存在的职位代码（确保不匹配）
            fake_position_code = f"3001{random.randint(100, 999)}{random.randint(1000, 9999)}"
            fake_department_code = f"{random.randint(200, 299)}"
            fake_department_name = f"虚拟部门{random.randint(1, 50)}"
            fake_bureau_name = f"{fake_department_name}虚拟司局{random.randint(1, 20)}"
            fake_position_name = f"{fake_bureau_name}虚拟职位{random.randint(1, 100)}"
            
            # 生成准考证号
            exam_number = f"2024{fake_position_code[-6:]}{random.randint(1, 999):03d}"
            
            # 生成最低面试分数
            min_score = round(random.uniform(60.0, 85.0), 1)
            
            interview_data.append({
                '准考证号': exam_number,
                '姓名': full_name,
                '招录机关': fake_department_name,
                '部门代码': fake_department_code,
                '用人司局': fake_bureau_name,
                '招考职位': fake_position_name,
                '职位代码': fake_position_code,
                '最低面试分数': min_score
            })
    
    # 随机打乱数据顺序
    random.shuffle(interview_data)
    
    return pd.DataFrame(interview_data)

def main():
    """主函数"""
    print("开始生成大规模测试数据...")
    
    # 创建测试数据文件夹
    folder_name = create_test_data_folder()
    print(f"创建文件夹: {folder_name}")
    
    # 生成岗位表数据
    print("生成岗位表数据 (约200个职位)...")
    positions_df = generate_random_positions(200)
    positions_file = os.path.join(folder_name, "大规模测试岗位表.xlsx")
    positions_df.to_excel(positions_file, index=False, engine='openpyxl')
    print(f"岗位表生成完成: {len(positions_df)}个职位")
    
    # 生成面试人员数据
    print("生成面试人员数据 (约2000个人员)...")
    interview_df = generate_random_interview_data(positions_df, 2000)
    interview_file = os.path.join(folder_name, "大规模测试面试人员名单.xlsx")
    interview_df.to_excel(interview_file, index=False, engine='openpyxl')
    print(f"面试人员名单生成完成: {len(interview_df)}个人员")
    
    # 数据统计
    print("\n数据统计:")
    print(f"岗位表: {len(positions_df)}行, {len(positions_df.columns)}列")
    print(f"面试人员: {len(interview_df)}行, {len(interview_df.columns)}列")
    
    # 匹配性分析
    position_codes_in_positions = set(positions_df['职位代码'])
    position_codes_in_interviews = set(interview_df['职位代码'])
    
    matched_codes = position_codes_in_positions.intersection(position_codes_in_interviews)
    unmatched_codes = position_codes_in_interviews - position_codes_in_positions
    
    print(f"匹配的职位代码: {len(matched_codes)}个")
    print(f"不匹配的职位代码: {len(unmatched_codes)}个")
    print(f"匹配率: {len(matched_codes)/len(position_codes_in_interviews)*100:.1f}%")
    
    # 生成数据说明文件
    readme_content = f"""# 大规模测试数据说明

## 生成时间
{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## 文件说明
- **大规模测试岗位表.xlsx**: {len(positions_df)}个职位，{len(positions_df.columns)}个字段
- **大规模测试面试人员名单.xlsx**: {len(interview_df)}个面试人员，{len(interview_df.columns)}个字段

## 数据特点
1. **完全随机生成**: 所有数据都是随机生成，确保测试的客观性
2. **大规模数据**: 职位表约200个，面试人员约2000个，接近真实场景
3. **匹配性设计**: 
   - 匹配的职位代码: {len(matched_codes)}个
   - 不匹配的职位代码: {len(unmatched_codes)}个
   - 匹配率: {len(matched_codes)/len(position_codes_in_interviews)*100:.1f}%
4. **数据完整性**: 包含所有必要字段，支持完整功能测试

## 测试目的
- 验证程序处理大规模数据的性能
- 测试数据匹配算法的准确性
- 验证不匹配数据的处理能力
- 测试各种功能模块的稳定性

## 使用方法
1. 将两个Excel文件导入到程序中
2. 配置列映射关系
3. 选择需要的输出列
4. 运行处理程序
5. 查看处理结果和性能表现
"""
    
    readme_file = os.path.join(folder_name, "README.md")
    with open(readme_file, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    print(f"\n测试数据生成完成！")
    print(f"文件保存在: {folder_name}")
    print(f"包含文件:")
    print(f"   - 大规模测试岗位表.xlsx")
    print(f"   - 大规模测试面试人员名单.xlsx") 
    print(f"   - README.md")
    print(f"\n提示: 程序运行时生成的岗位分数汇总文件也会自动保存到此文件夹中")

if __name__ == "__main__":
    main()