# -*- coding: utf-8 -*-
"""简历提取器主入口"""

import json
import sys
from pathlib import Path
from extractor import ResumeExtractor


def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("用法: python main.py <Excel文件路径>")
        sys.exit(1)

    file_path = sys.argv[1]
    file_path_obj = Path(file_path)

    if not file_path_obj.exists():
        print(f"文件不存在: {file_path}")
        sys.exit(1)

    # 检查文件扩展名
    if file_path_obj.suffix.lower() not in [".xls", ".xlsx"]:
        print(f"不支持的文件格式: {file_path_obj.suffix}")
        print("请提供 .xls 或 .xlsx 文件")
        sys.exit(1)

    # 创建提取器实例
    extractor = ResumeExtractor()

    # 提取信息
    print(f"正在处理文件: {file_path}")
    print(f"文件格式: {file_path_obj.suffix}")

    result = extractor.extract_from_excel(file_path)

    # 显示结果
    if "error" not in result:
        print("\n✅ 提取结果:")
        print(f"   姓名: {result['name']}")
        print(f"   性别: {result['gender']}")
        print(f"   年龄: {result['age']}")
        print(f"   国籍: {result['nationality']}")
        print(f"   来日年份: {result['arrival_year_japan']}")
        print(f"   经验: {result['experience']}")
        print(f"   日语: {result['japanese_level']}")
        print(f"   技能({len(result['skills'])}个): {', '.join(result['skills'][:10])}")
        if len(result["skills"]) > 10:
            print(f"        ... 还有 {len(result['skills']) - 10} 个技能")
        print(f"   作业范围: {', '.join(result['work_scope'])}")
        print(f"   角色: {', '.join(result['roles'])}")

        print("\n📋 JSON格式:")
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"❌ 错误: {result['error']}")


if __name__ == "__main__":
    main()
