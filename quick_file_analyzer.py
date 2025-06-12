# -*- coding: utf-8 -*-
"""快速Excel文件分析器 - 找出姓名在哪里"""

import pandas as pd
import re
from pathlib import Path


def analyze_excel_structure(file_path: str):
    """分析Excel文件结构，找出可能的姓名位置"""

    print("Excel文件结构分析器")
    print("=" * 50)

    if not Path(file_path).exists():
        print(f"❌ 文件不存在: {file_path}")
        return

    try:
        # 读取Excel文件
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")

        for sheet_name, df in all_sheets.items():
            print(f"\n📋 Sheet: {sheet_name}")
            print(f"大小: {df.shape[0]}行 x {df.shape[1]}列")

            # 输出前10行的内容
            print(f"\n前10行内容:")
            for i in range(min(10, len(df))):
                row_data = []
                for j in range(min(10, len(df.columns))):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()
                        # 限制显示长度
                        if len(cell_str) > 15:
                            cell_str = cell_str[:12] + "..."
                        row_data.append(f"[{j}]{cell_str}")
                    else:
                        row_data.append(f"[{j}]空")
                print(f"  行{i:2d}: {' | '.join(row_data)}")

            # 查找可能的姓名
            print(f"\n🔍 寻找可能的姓名:")
            find_possible_names(df)

            # 查找姓名相关关键词
            print(f"\n🔍 姓名关键词位置:")
            find_name_keywords(df)

    except Exception as e:
        print(f"❌ 分析文件时出错: {e}")


def find_possible_names(df: pd.DataFrame):
    """查找可能的姓名"""

    possible_names = []

    # 在前10行中查找可能的姓名
    for i in range(min(10, len(df))):
        for j in range(min(10, len(df.columns))):
            cell = df.iloc[i, j]
            if pd.notna(cell):
                cell_str = str(cell).strip()

                # 姓名模式匹配
                if is_likely_name(cell_str):
                    possible_names.append(
                        {
                            "value": cell_str,
                            "row": i,
                            "col": j,
                            "score": calculate_name_likelihood(cell_str),
                        }
                    )

    # 按评分排序
    possible_names.sort(key=lambda x: x["score"], reverse=True)

    if possible_names:
        print("    可能的姓名候选:")
        for name_info in possible_names[:5]:  # 只显示前5个
            print(
                f"      [{name_info['row']},{name_info['col']}] '{name_info['value']}' (评分: {name_info['score']:.1f})"
            )
    else:
        print("    ❌ 未找到明显的姓名候选")


def find_name_keywords(df: pd.DataFrame):
    """查找姓名相关关键词的位置"""

    name_keywords = ["氏名", "氏 名", "名前", "フリガナ", "ふりがな", "Name", "姓名"]

    found_keywords = []

    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = df.iloc[i, j]
            if pd.notna(cell):
                cell_str = str(cell).strip()

                for keyword in name_keywords:
                    if keyword in cell_str:
                        found_keywords.append(
                            {"keyword": keyword, "value": cell_str, "row": i, "col": j}
                        )

    if found_keywords:
        print("    找到的姓名关键词:")
        for kw_info in found_keywords:
            print(
                f"      [{kw_info['row']},{kw_info['col']}] '{kw_info['value']}' (包含: {kw_info['keyword']})"
            )

            # 查看关键词周围的内容
            print(f"        周围内容:")
            show_surrounding_cells(df, kw_info["row"], kw_info["col"])
    else:
        print("    ❌ 未找到姓名关键词")


def show_surrounding_cells(df: pd.DataFrame, center_row: int, center_col: int):
    """显示指定位置周围的单元格内容"""

    for r_offset in range(-1, 3):
        for c_offset in range(-1, 6):
            r = center_row + r_offset
            c = center_col + c_offset

            if 0 <= r < len(df) and 0 <= c < len(df.columns):
                cell = df.iloc[r, c]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    if cell_str:  # 只显示非空内容
                        marker = "🎯" if (r == center_row and c == center_col) else "  "
                        print(f"          {marker}[{r},{c}] '{cell_str}'")


def is_likely_name(text: str) -> bool:
    """判断文本是否可能是姓名"""

    text = text.strip()

    # 基本过滤
    if not text or len(text) > 15 or len(text) < 2:
        return False

    # 排除明显的标签词
    obvious_labels = [
        "氏名",
        "名前",
        "フリガナ",
        "性別",
        "年齢",
        "国籍",
        "得意分野",
        "専門分野",
        "技術分野",
        "開発経験",
        "業務経験",
        "使用技術",
        "開発環境",
        "役割",
        "職種",
        "自己PR",
        "男性",
        "女性",
    ]

    if any(label in text for label in obvious_labels):
        return False

    # 排除数字和日期
    if re.match(r"^\d+$", text) or re.match(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
        return False

    # 排除技术术语
    tech_terms = ["Java", "JavaScript", "PHP", "Python", "HTML", "CSS", "SQL"]
    if text in tech_terms:
        return False

    # 必须包含文字字符
    if not re.search(r"[一-龥ぁ-んァ-ンa-zA-Z]", text):
        return False

    return True


def calculate_name_likelihood(text: str) -> float:
    """计算文本是姓名的可能性评分"""

    score = 0.0

    # 长度评分
    if 2 <= len(text) <= 6:
        score += 2.0
    elif len(text) == 1:
        score += 0.5

    # 字符类型评分
    if re.search(r"[一-龥]", text):  # 包含汉字
        score += 3.0

    if re.search(r"[ぁ-んァ-ン]", text):  # 包含假名
        score += 2.0

    if re.search(r"[A-Z]", text):  # 包含大写字母
        score += 1.5

    if re.search(r"[a-z]", text):  # 包含小写字母
        score += 1.0

    # 姓名模式评分
    if re.match(r"^[一-龥]{1,3}[A-Z]{1,3}$", text):  # 中文姓+英文名
        score += 4.0

    if re.match(r"^[一-龥]{2,4}$", text):  # 纯中文姓名
        score += 3.0

    if re.match(r"^[A-Z]{2,4}$", text):  # 英文缩写
        score += 2.0

    # 位置评分（假设在前几行前几列）
    # 这个会在实际使用时添加

    return score


def quick_validation_test():
    """快速验证函数测试"""

    print(f"\n" + "=" * 50)
    print("快速验证函数测试")
    print("=" * 50)

    # 导入验证函数
    try:
        from utils.validation_utils import is_valid_name

        test_cases = [
            "庄YW",
            "庄 YW",
            "庄　YW",
            "ショウYW",
            "庄Young",
            "庄young",
            "YW",
            "庄",
            "得意分野",
            "フリガナ",
            "氏名",
            "男性",
            "Java",
            "30",
        ]

        print("验证函数测试结果:")
        for name in test_cases:
            result = is_valid_name(name)
            status = "✅" if result else "❌"
            likelihood = calculate_name_likelihood(name)
            print(f"{status} '{name}' -> 验证:{result}, 可能性:{likelihood:.1f}")

    except ImportError:
        print("无法导入验证函数，跳过测试")


def main():
    """主函数"""

    import sys

    if len(sys.argv) < 2:
        print("用法: python quick_file_analyzer.py <Excel文件路径>")
        print('示例: python quick_file_analyzer.py "【スキルシート】_庄YW_202411.xlsx"')
        sys.exit(1)

    file_path = sys.argv[1]

    # 分析文件结构
    analyze_excel_structure(file_path)

    # 验证函数测试
    quick_validation_test()


if __name__ == "__main__":
    main()
