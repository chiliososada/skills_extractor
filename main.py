# -*- coding: utf-8 -*-
"""简历提取器主入口"""

import json
import sys
from pathlib import Path
from extractor import ResumeExtractor


def format_value(value):
    """格式化显示值"""
    if value is None:
        return "null"
    elif isinstance(value, list):
        if not value:
            return "null"
        else:
            return value
    else:
        return value


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
        print(f"   姓名: {format_value(result['name'])}")
        print(f"   性别: {format_value(result['gender'])}")
        print(f"   年龄: {format_value(result['age'])}")
        print(f"   出生年月日: {format_value(result['birthdate'])}")
        print(f"   国籍: {format_value(result['nationality'])}")
        print(f"   来日年份: {format_value(result['arrival_year_japan'])}")
        print(f"   经验: {format_value(result['experience'])}")
        print(f"   日语: {format_value(result['japanese_level'])}")

        # 技能特殊处理
        if result["skills"] is None:
            print(f"   技能: null")
        else:
            print(
                f"   技能({len(result['skills'])}个): {', '.join(result['skills'][:10])}"
            )
            if len(result["skills"]) > 10:
                print(f"        ... 还有 {len(result['skills']) - 10} 个技能")

        # 作业范围
        if result["work_scope"] is None:
            print(f"   作业范围: null")
        else:
            print(f"   作业范围: {', '.join(result['work_scope'])}")

        # 角色
        if result["roles"] is None:
            print(f"   角色: null")
        else:
            print(f"   角色: {', '.join(result['roles'])}")

        print("\n📋 JSON格式:")
        # 使用json.dumps时，ensure_ascii=False保留中文，indent=2格式化
        # 注意：Python的None会自动转换为JSON的null
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"❌ 错误: {result['error']}")


if __name__ == "__main__":
    main()
