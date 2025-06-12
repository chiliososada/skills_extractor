# diagnostic_resume_extractor.py - 诊断版简历提取器
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
import re
import json
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Union, Tuple, Set
from collections import defaultdict, Counter


class DiagnosticResumeExtractor:
    """诊断版简历提取器 - 显示所有数据结构来找问题"""

    def __init__(self):
        pass

    def diagnose_file(self, file_path: str) -> Dict:
        """诊断文件结构"""
        try:
            print(f"\n🔬 === 开始诊断文件: {file_path} ===")

            # 读取所有sheets
            all_sheets = pd.read_excel(file_path, sheet_name=None)

            for sheet_name, df in all_sheets.items():
                print(f"\n📋 工作表: {sheet_name}")
                print(f"   大小: {len(df)}行 x {len(df.columns)}列")

                # 1. 显示所有Date对象
                print(f"\n📅 所有Date对象:")
                date_count = 0
                for idx in range(min(50, len(df))):
                    for col in range(len(df.columns)):
                        cell = df.iloc[idx, col]
                        if isinstance(cell, datetime):
                            date_count += 1
                            print(f"   [{idx},{col}]: {cell} ({cell.year}年)")

                            # 显示同行内容
                            row_data = df.iloc[idx]
                            row_text = " | ".join(
                                [str(cell) for cell in row_data if pd.notna(cell)]
                            )
                            print(f"      同行: {row_text[:150]}...")

                if date_count == 0:
                    print("   ❌ 未找到任何Date对象")
                else:
                    print(f"   ✅ 共找到 {date_count} 个Date对象")

                # 2. 搜索特定关键词位置
                keywords_to_find = {
                    "来日": ["来日", "渡日", "入国"],
                    "年龄": ["年齢", "歳", "才", "生年月", "満"],
                    "国籍": ["国籍", "出身国", "出身地"],
                    "经验": ["経験年数", "実務経験", "開発経験", "IT経験"],
                }

                for category, keywords in keywords_to_find.items():
                    print(f"\n🔍 搜索 {category} 相关关键词:")
                    found_count = 0
                    for idx in range(min(50, len(df))):
                        for col in range(len(df.columns)):
                            cell = df.iloc[idx, col]
                            if pd.notna(cell):
                                cell_str = str(cell)
                                for keyword in keywords:
                                    if keyword in cell_str:
                                        found_count += 1
                                        print(
                                            f"   [{idx},{col}]: {keyword} -> {cell_str[:100]}"
                                        )

                                        # 显示周围内容
                                        print(f"      周围内容:")
                                        for r_off in range(-1, 2):
                                            for c_off in range(-2, 5):
                                                r = idx + r_off
                                                c = col + c_off
                                                if 0 <= r < len(df) and 0 <= c < len(
                                                    df.columns
                                                ):
                                                    nearby_cell = df.iloc[r, c]
                                                    if pd.notna(nearby_cell):
                                                        print(
                                                            f"        [{r},{c}]: {str(nearby_cell)[:50]}"
                                                        )
                                        break

                    if found_count == 0:
                        print(f"   ❌ 未找到 {category} 相关关键词")

                # 3. 显示前20行的主要内容
                print(f"\n📄 前20行主要内容:")
                for idx in range(min(20, len(df))):
                    row_data = df.iloc[idx]
                    row_text = " | ".join(
                        [
                            str(cell)
                            for cell in row_data
                            if pd.notna(cell) and str(cell).strip()
                        ]
                    )
                    if row_text:
                        print(f"   [{idx}]: {row_text[:200]}")

                # 4. 特别搜索数字模式（可能的年龄、年份、经验）
                print(f"\n🔢 数字模式分析:")
                number_patterns = [
                    (r"\b(19|20)\d{2}\b", "年份"),
                    (r"\b[12]?\d\s*[歳才]\b", "年龄"),
                    (r"\b\d+(?:\.\d+)?\s*年\b", "经验年数"),
                    (r"\b[12]?\d\b", "纯数字"),
                ]

                for pattern, desc in number_patterns:
                    print(f"   {desc} 模式:")
                    found_numbers = set()
                    for idx in range(min(30, len(df))):
                        for col in range(len(df.columns)):
                            cell = df.iloc[idx, col]
                            if pd.notna(cell):
                                cell_str = str(cell)
                                matches = re.findall(pattern, cell_str)
                                for match in matches:
                                    if isinstance(match, tuple):
                                        match = "".join(match)
                                    if match not in found_numbers:
                                        found_numbers.add(match)
                                        print(
                                            f"     [{idx},{col}]: {match} (在: {cell_str[:80]})"
                                        )

                    if not found_numbers:
                        print(f"     ❌ 未找到 {desc}")

                # 5. 特别诊断3个问题文件
                filename = file_path.split("/")[-1]
                if "ryu" in filename.lower():
                    self._diagnose_ryu_arrival(df)
                elif "fyy" in filename.lower():
                    self._diagnose_fyy_age(df)
                elif "gt" in filename.lower():
                    self._diagnose_gt_issues(df)

            return {"status": "诊断完成"}

        except Exception as e:
            print(f"❌ 诊断出错: {e}")
            import traceback

            traceback.print_exc()
            return {"error": str(e)}

    def _diagnose_ryu_arrival(self, df: pd.DataFrame):
        """专门诊断Ryu文件的来日年份问题"""
        print(f"\n🔬 === 专项诊断: Ryu文件来日年份 ===")

        # 寻找任何可能是2015年的信息
        print(f"🔍 搜索所有包含2015的内容:")
        for idx in range(len(df)):
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if pd.notna(cell):
                    cell_str = str(cell)
                    if "2015" in cell_str:
                        print(f"   [{idx},{col}]: {cell} (类型: {type(cell)})")
                        # 显示周围内容
                        for r_off in range(-1, 2):
                            for c_off in range(-3, 4):
                                r = idx + r_off
                                c = col + c_off
                                if 0 <= r < len(df) and 0 <= c < len(df.columns):
                                    nearby = df.iloc[r, c]
                                    if pd.notna(nearby):
                                        print(f"     [{r},{c}]: {nearby}")

        # 特别检查第4行（从之前分析看应该在这里）
        if len(df) > 3:
            print(f"\n📍 第4行详细内容:")
            row_4 = df.iloc[3]
            for col_idx, cell in enumerate(row_4):
                if pd.notna(cell):
                    print(f"   列{col_idx}: {cell} (类型: {type(cell)})")

    def _diagnose_fyy_age(self, df: pd.DataFrame):
        """专门诊断FYY文件的年龄问题"""
        print(f"\n🔬 === 专项诊断: FYY文件年龄问题 ===")

        # 寻找任何可能是1991年的信息
        print(f"🔍 搜索所有包含1991的内容:")
        for idx in range(len(df)):
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if pd.notna(cell):
                    cell_str = str(cell)
                    if "1991" in cell_str:
                        print(f"   [{idx},{col}]: {cell} (类型: {type(cell)})")
                        # 显示周围内容
                        for r_off in range(-1, 2):
                            for c_off in range(-3, 4):
                                r = idx + r_off
                                c = col + c_off
                                if 0 <= r < len(df) and 0 <= c < len(df.columns):
                                    nearby = df.iloc[r, c]
                                    if pd.notna(nearby):
                                        print(f"     [{r},{c}]: {nearby}")

        # 检查第5行（从之前分析看应该在这里）
        if len(df) > 4:
            print(f"\n📍 第5行详细内容:")
            row_5 = df.iloc[4]
            for col_idx, cell in enumerate(row_5):
                if pd.notna(cell):
                    print(f"   列{col_idx}: {cell} (类型: {type(cell)})")

    def _diagnose_gt_issues(self, df: pd.DataFrame):
        """专门诊断GT文件的问题"""
        print(f"\n🔬 === 专项诊断: GT文件国籍和经验问题 ===")

        # 寻找经验年数的正确位置
        print(f"🔍 搜索実務年数表格结构:")
        for idx in range(len(df)):
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if pd.notna(cell) and "実務年数" in str(cell):
                    print(f"   实务年数标题: [{idx},{col}]: {cell}")

                    # 检查下面几行的内容
                    print(f"   下方内容:")
                    for next_row in range(idx + 1, min(idx + 5, len(df))):
                        row_data = df.iloc[next_row]
                        row_text = " | ".join(
                            [str(cell) for cell in row_data[:10] if pd.notna(cell)]
                        )
                        if row_text:
                            print(f"     [{next_row}]: {row_text}")

        # 寻找可能的国籍信息（检查是否真的没有）
        print(f"\n🔍 全表搜索可能的国籍信息:")
        nationalities = [
            "中国",
            "日本",
            "韓国",
            "ベトナム",
            "フィリピン",
            "アメリカ",
            "台湾",
        ]
        for nationality in nationalities:
            for idx in range(len(df)):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell) and nationality in str(cell):
                        print(f"   找到国籍候选: [{idx},{col}]: {cell}")


# 测试函数
if __name__ == "__main__":
    extractor = DiagnosticResumeExtractor()

    test_files = [
        "スキルシート_ryu.xlsx",  # 来日年份问题
        "業務経歴書_ FYY_金町駅.xlsx",  # 年龄问题
        "スキルシートGT.xlsx",  # 国籍和经验问题
    ]

    for file in test_files:
        result = extractor.diagnose_file(file)
        print(f"\n" + "=" * 80)
