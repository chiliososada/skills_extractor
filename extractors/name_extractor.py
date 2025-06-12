# -*- coding: utf-8 -*-
"""姓名提取器 - 针对性修复版：解决全角空格问题"""

from typing import List, Dict, Any
import pandas as pd
import re

from base.base_extractor import BaseExtractor
from base.constants import KEYWORDS
from utils.validation_utils import is_valid_name


class NameExtractor(BaseExtractor):
    """姓名信息提取器 - 针对性修复版"""

    def __init__(self):
        super().__init__()
        self.problematic_labels = [
            "得意分野",
            "得意 分野",
            "得意　分野",
            "氏名",
            "氏 名",
            "氏　名",
            "名前",
            "名 前",
            "名　前",
            "フリガナ",
            "ふりがな",
        ]

    def extract(self, all_data: List[Dict[str, Any]]) -> str:
        """提取姓名 - 针对性修复版"""
        candidates = []

        for data in all_data:
            df = data["df"]
            sheet_name = data.get("sheet_name", "Unknown")

            print(f"\n🔍 开始姓名提取 - Sheet: {sheet_name}")

            # 专门针对[2,0]位置的"氏　名"进行搜索
            specific_candidates = self._search_specific_position(df)
            if specific_candidates:
                print(f"    ✅ 在特定位置找到 {len(specific_candidates)} 个候选")
                candidates.extend(specific_candidates)

            # 通用关键词搜索
            general_candidates = self._search_by_keywords_enhanced(df)
            if general_candidates:
                print(f"    ✅ 通用搜索找到 {len(general_candidates)} 个候选")
                candidates.extend(general_candidates)

            # 备用搜索
            if not candidates:
                backup_candidates = self._search_top_rows_backup(df)
                candidates.extend(backup_candidates)

        if candidates:
            # 过滤和排序
            filtered_candidates = []
            for name, confidence in candidates:
                # 标准化姓名（处理全角空格）
                normalized_name = self._normalize_name(name)

                if not self._is_obvious_label(normalized_name):
                    # 验证标准化后的姓名
                    if is_valid_name(normalized_name):
                        filtered_candidates.append((normalized_name, confidence))
                        print(
                            f"    ✅ 保留候选: '{name}' -> '{normalized_name}' (置信度: {confidence:.2f})"
                        )
                    else:
                        print(f"    ❌ 验证失败: '{normalized_name}'")
                else:
                    print(f"    ❌ 过滤标签: '{name}'")

            if filtered_candidates:
                filtered_candidates.sort(key=lambda x: x[1], reverse=True)
                best_name = filtered_candidates[0][0]
                print(
                    f"\n✅ 最终选择姓名: '{best_name}' (置信度: {filtered_candidates[0][1]:.2f})"
                )
                return best_name

        print("\n❌ 未能提取到姓名")
        return ""

    def _search_specific_position(self, df: pd.DataFrame) -> List[tuple]:
        """专门搜索[2,0]位置的氏名对应的[2,3]位置"""
        candidates = []

        # 检查[2,0]是否包含"氏"
        if len(df) > 2 and len(df.columns) > 0:
            cell_2_0 = df.iloc[2, 0]
            if pd.notna(cell_2_0) and "氏" in str(cell_2_0):
                print(f"    🎯 在[2,0]发现氏名关键词: '{cell_2_0}'")

                # 检查[2,3]位置
                if len(df.columns) > 3:
                    cell_2_3 = df.iloc[2, 3]
                    if pd.notna(cell_2_3):
                        value = str(cell_2_3).strip()
                        print(f"    🎯 在[2,3]发现内容: '{value}'")

                        # 这就是我们要找的姓名！
                        normalized = self._normalize_name(value)
                        candidates.append((normalized, 5.0))  # 给予最高置信度
                        print(f"    ✅ 特定位置候选: '{value}' -> '{normalized}'")

        return candidates

    def _search_by_keywords_enhanced(self, df: pd.DataFrame) -> List[tuple]:
        """增强的关键词搜索"""
        candidates = []

        for idx in range(min(10, len(df))):
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if pd.notna(cell):
                    cell_str = str(cell).strip()

                    if any(k in cell_str for k in KEYWORDS["name"]):
                        print(f"    找到姓名关键词 '{cell_str}' 在位置 [{idx}, {col}]")

                        # 不要跳过任何内容，直接搜索附近
                        nearby_candidates = self._search_nearby_enhanced(df, idx, col)
                        candidates.extend(nearby_candidates)

        return candidates

    def _search_nearby_enhanced(
        self, df: pd.DataFrame, row: int, col: int
    ) -> List[tuple]:
        """增强的附近搜索"""
        candidates = []

        # 搜索更大的范围
        for r_offset in range(-2, 5):
            for c_offset in range(-2, 20):  # 扩大列搜索范围
                r = row + r_offset
                c = col + c_offset

                if 0 <= r < len(df) and 0 <= c < len(df.columns):
                    value = df.iloc[r, c]
                    if pd.notna(value):
                        value_str = str(value).strip()

                        if value_str and len(value_str) >= 2:
                            # 标准化处理
                            normalized = self._normalize_name(value_str)

                            # 简化验证：只排除明显的标签
                            if not self._is_obvious_label(normalized):
                                # 进一步验证
                                if is_valid_name(normalized):
                                    confidence = self._calculate_confidence(
                                        row, col, r, c, normalized
                                    )
                                    candidates.append((normalized, confidence))
                                    print(
                                        f"      候选: '{value_str}' -> '{normalized}' 行{r}, 列{c}, 置信度{confidence:.2f}"
                                    )

        return candidates

    def _search_top_rows_backup(self, df: pd.DataFrame) -> List[tuple]:
        """备用的前行搜索"""
        candidates = []

        for row in range(min(5, len(df))):
            for col in range(min(10, len(df.columns))):
                cell = df.iloc[row, col]
                if pd.notna(cell):
                    cell_str = str(cell).strip()

                    if cell_str and len(cell_str) >= 2:
                        normalized = self._normalize_name(cell_str)

                        if not self._is_obvious_label(normalized) and is_valid_name(
                            normalized
                        ):

                            confidence = 0.5
                            if row <= 2 and col <= 5:
                                confidence += 0.3
                            if len(normalized) >= 2:
                                confidence += 0.3

                            candidates.append((normalized, confidence))
                            print(
                                f"    备用: '{cell_str}' -> '{normalized}' 行{row}, 列{col}"
                            )

        return candidates

    def _normalize_name(self, name: str) -> str:
        """标准化姓名：处理各种空格"""
        if not name:
            return ""

        # 移除前后空格
        name = name.strip()

        # 将全角空格替换为半角空格
        name = name.replace("　", " ")

        # 移除多余的空格（保留单个空格用于姓名分隔）
        name = re.sub(r"\s+", " ", name)

        # 如果是单个空格分隔的姓名，保持原样
        # 否则移除所有空格
        parts = name.split(" ")
        if len(parts) == 2 and all(len(p) >= 1 for p in parts):
            # 姓名格式：保持空格
            return name
        else:
            # 其他情况：移除空格
            return name.replace(" ", "")

    def _is_obvious_label(self, text: str) -> bool:
        """检查是否是明显的标签"""
        if not text:
            return True

        text_clean = re.sub(r"\s+", "", text.lower())

        # 检查问题标签
        for label in self.problematic_labels:
            label_clean = re.sub(r"\s+", "", label.lower())
            if text_clean == label_clean:
                return True

        # 检查明显的非姓名词汇
        obvious_non_names = [
            "年齢",
            "性別",
            "国籍",
            "住所",
            "電話",
            "メール",
            "男性",
            "女性",
            "男",
            "女",
            "歳",
            "才",
            "年",
            "月",
            "日",
            "経験",
            "スキル",
            "技術",
            "開発",
            "業務",
            "プロジェクト",
            "システム",
            "言語",
            "ツール",
            "環境",
            "資格",
            "得意分野",
            "専門分野",
            "開発技術",
            "業務知識",
            "自己pr",
            "java",
            "javascript",
            "php",
            "python",
            "html",
            "css",
            "sql",
        ]

        for word in obvious_non_names:
            if word.lower() in text_clean:
                return True

        return False

    def _calculate_confidence(
        self,
        base_row: int,
        base_col: int,
        candidate_row: int,
        candidate_col: int,
        value: str,
    ) -> float:
        """计算置信度"""
        confidence = 1.0

        # 位置因素
        if candidate_row == base_row:
            confidence *= 1.5

        if candidate_col > base_col and candidate_col - base_col <= 5:
            confidence *= 1.3

        # 距离因素
        distance = abs(candidate_row - base_row) + abs(candidate_col - base_col)
        confidence *= 1.0 / (1 + distance * 0.1)

        # 内容因素
        if len(value) >= 2:
            confidence *= 1.5

        if re.search(r"[一-龥]", value):
            confidence *= 1.3

        if re.search(r"[A-Za-z]", value):
            confidence *= 1.1

        return confidence
