# -*- coding: utf-8 -*-
"""姓名提取器 - 最终修复版（针对スキルシート_付.xlsx问题）"""

from typing import List, Dict, Any
import pandas as pd
import re

from base.base_extractor import BaseExtractor
from base.constants import KEYWORDS
from utils.validation_utils import is_valid_name


class NameExtractor(BaseExtractor):
    """姓名信息提取器 - 最终修复版"""

    def __init__(self):
        super().__init__()
        # 学历相关关键词，用于避免在学历区域搜索姓名
        self.education_keywords = [
            "学歴",
            "最終学歴",
            "学校",
            "大学",
            "研究科",
            "学院",
            "専門学校",
            "高校",
            "卒業",
            "在学",
            "専攻",
            "学科",
            "学部",
            "研究室",
            "博士",
            "修士",
            "学士",
            "PhD",
            "Master",
            "Bachelor",
        ]

    def extract(self, all_data: List[Dict[str, Any]]) -> str:
        """提取姓名

        Args:
            all_data: 包含所有sheet数据的列表

        Returns:
            提取的姓名，如果未找到返回空字符串
        """
        candidates = []

        for data in all_data:
            df = data["df"]
            sheet_name = data.get("sheet_name", "Unknown")

            print(f"\n🔍 开始姓名提取 - Sheet: {sheet_name}")

            # 方法1: 精确搜索姓名关键词附近
            primary_candidates = self._search_name_by_keywords_enhanced(df)
            if primary_candidates:
                print(f"    通过关键词找到 {len(primary_candidates)} 个候选姓名")
                candidates.extend(primary_candidates)

            # 方法2: 如果主要方法失败，使用备用搜索（限制在前5行）
            if not candidates:
                print("    使用备用方法：前5行搜索")
                backup_candidates = self._search_name_in_top_rows(df)
                candidates.extend(backup_candidates)

        if candidates:
            # 按置信度排序，返回最佳候选
            candidates.sort(key=lambda x: x[1], reverse=True)
            best_name = candidates[0][0].strip()
            print(f"\n✅ 最终选择姓名: '{best_name}' (置信度: {candidates[0][1]:.2f})")
            return best_name

        print("\n❌ 未能提取到姓名")
        return ""

    def _search_name_by_keywords_enhanced(self, df: pd.DataFrame) -> List[tuple]:
        """通过姓名关键词搜索姓名 - 增强版"""
        candidates = []

        # 只搜索前10行，避免在学历等区域搜索
        for idx in range(min(10, len(df))):
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if pd.notna(cell) and any(k in str(cell) for k in KEYWORDS["name"]):
                    print(f"    找到姓名关键词 '{cell}' 在位置 [{idx}, {col}]")

                    # 在附近搜索姓名值 - 使用优化的搜索策略
                    nearby_candidates = self._search_name_nearby_optimized(df, idx, col)
                    candidates.extend(nearby_candidates)

        return candidates

    def _search_name_nearby_optimized(
        self, df: pd.DataFrame, row: int, col: int
    ) -> List[tuple]:
        """在指定位置附近搜索姓名 - 优化版

        Args:
            df: DataFrame对象
            row: 行索引
            col: 列索引

        Returns:
            候选姓名列表，格式为 [(姓名, 置信度), ...]
        """
        candidates = []

        # 优化搜索范围：减小范围以提高精度
        for r_offset in range(-1, 3):  # 缩小行搜索范围
            for c_offset in range(-1, 8):  # 缩小列搜索范围
                r = row + r_offset
                c = col + c_offset

                if 0 <= r < len(df) and 0 <= c < len(df.columns):
                    value = df.iloc[r, c]
                    if pd.notna(value):
                        value_str = str(value).strip()

                        # 跳过明显不是姓名的内容
                        if self._is_obviously_not_name(value_str):
                            continue

                        # 检查这个候选位置是否在学历区域
                        if self._is_in_education_area(df, r, c):
                            print(f"      跳过学历区域 [{r},{c}]: '{value_str}'")
                            continue

                        if is_valid_name(value_str):
                            confidence = self._calculate_confidence(
                                row, col, r, c, value_str
                            )
                            candidates.append((value_str, confidence))
                            print(
                                f"      找到候选姓名: '{value_str}' 行{r}, 列{c}, 置信度{confidence:.2f}"
                            )

        return candidates

    def _calculate_confidence(
        self,
        base_row: int,
        base_col: int,
        candidate_row: int,
        candidate_col: int,
        value: str,
    ) -> float:
        """计算姓名候选的置信度"""
        confidence = 1.0

        # 1. 位置因素
        if candidate_row == base_row:
            confidence *= 1.8  # 同行，但不要过高

        if candidate_col > base_col and candidate_col - base_col <= 3:
            confidence *= 1.3  # 右侧邻近

        # 2. 距离因素
        distance = abs(candidate_row - base_row) + abs(candidate_col - base_col)
        confidence *= 1.0 / (1 + distance * 0.15)

        # 3. 内容因素 - 关键改进
        trimmed_value = value.strip()

        # 优先选择实际的姓名而不是标记
        if len(trimmed_value) >= 2:
            confidence *= 1.5  # 多字符姓名获得更高置信度

        if re.search(r"[一-龥]", trimmed_value):
            confidence *= 1.2  # 汉字姓名获得更高置信度

        # 4. 避免选择单个假名标记
        kana_markers = [
            "ア",
            "イ",
            "ウ",
            "エ",
            "オ",
            "カ",
            "キ",
            "ク",
            "ケ",
            "コ",
            "サ",
            "シ",
            "ス",
            "セ",
            "ソ",
            "タ",
            "チ",
            "ツ",
            "テ",
            "ト",
            "ナ",
            "ニ",
            "ヌ",
            "ネ",
            "ノ",
            "ハ",
            "ヒ",
            "フ",
            "ヘ",
            "ホ",
            "マ",
            "ミ",
            "ム",
            "メ",
            "モ",
            "ヤ",
            "ユ",
            "ヨ",
            "ラ",
            "リ",
            "ル",
            "レ",
            "ロ",
            "ワ",
            "ヲ",
            "ン",
        ]

        if trimmed_value in kana_markers:
            confidence *= 0.1  # 大幅降低假名标记的置信度

        return confidence

    def _search_name_in_top_rows(self, df: pd.DataFrame) -> List[tuple]:
        """在前几行搜索可能的姓名（备用方法）"""
        candidates = []

        # 只搜索前5行，每行的前5列
        for row in range(min(5, len(df))):
            for col in range(min(5, len(df.columns))):
                cell = df.iloc[row, col]
                if pd.notna(cell):
                    cell_str = str(cell).strip()

                    # 跳过明显不是姓名的内容
                    if self._is_obviously_not_name(cell_str):
                        continue

                    # 跳过学历区域
                    if self._is_in_education_area(df, row, col):
                        continue

                    if is_valid_name(cell_str):
                        # 给予较低的置信度
                        confidence = 0.5

                        # 如果在个人信息可能出现的位置，提高置信度
                        if row <= 2 and col <= 3:
                            confidence += 0.3

                        # 多字符姓名获得更高置信度
                        if len(cell_str.strip()) >= 2:
                            confidence += 0.2

                        candidates.append((cell_str, confidence))
                        print(f"    备用方法找到候选: '{cell_str}' 行{row}, 列{col}")

        return candidates

    def _is_in_education_area(self, df: pd.DataFrame, row: int, col: int) -> bool:
        """检查是否在学历区域

        Args:
            df: DataFrame对象
            row: 行索引
            col: 列索引

        Returns:
            如果在学历区域返回True
        """
        # 检查当前行是否包含学历关键词
        if 0 <= row < len(df):
            row_data = df.iloc[row]
            row_text = " ".join([str(cell) for cell in row_data if pd.notna(cell)])

            # 如果当前行包含学历关键词，认为是学历区域
            if any(keyword in row_text for keyword in self.education_keywords):
                return True

        return False

    def _is_obviously_not_name(self, text: str) -> bool:
        """检查是否明显不是姓名

        Args:
            text: 待检查的文本

        Returns:
            如果明显不是姓名返回True
        """
        text = text.strip()

        # 空字符串
        if not text:
            return True

        # 太长的文本（可能是描述或组织名）
        if len(text) > 20:
            return True

        # 包含学历相关词汇
        if any(keyword in text for keyword in self.education_keywords):
            return True

        # 包含明显的非姓名标识符
        non_name_indicators = [
            "年",
            "月",
            "日",
            "歳",
            "才",
            "男",
            "女",
            "国",
            "県",
            "市",
            "区",
            "株式会社",
            "有限会社",
            "合同会社",
            "LLC",
            "Inc",
            "Corp",
            "Ltd",
            "TEL",
            "FAX",
            "Email",
            "住所",
            "〒",
            "番地",
            "丁目",
            "技術",
            "開発",
            "設計",
            "管理",
            "経験",
            "スキル",
            "言語",
            "ツール",
            "プロジェクト",
            "システム",
            "業務",
            "担当",
            "チーム",
            "部署",
        ]

        if any(indicator in text for indicator in non_name_indicators):
            return True

        # 全部是数字或符号
        if re.match(r"^[\d\W]+$", text):
            return True

        # 看起来像是标题或分类
        if (
            text.endswith("：")
            or text.endswith(":")
            or text.startswith("【")
            or text.endswith("】")
        ):
            return True

        return False
