# -*- coding: utf-8 -*-
"""年龄提取器"""

from typing import List, Dict, Any, Optional
from datetime import datetime
from collections import defaultdict
import pandas as pd
import re

from base.base_extractor import BaseExtractor
from base.constants import KEYWORDS
from utils.date_utils import convert_excel_serial_to_date, calculate_age_from_birthdate


class AgeExtractor(BaseExtractor):
    """年龄信息提取器"""

    def extract(self, all_data: List[Dict[str, Any]]) -> str:
        """提取年龄

        Args:
            all_data: 包含所有sheet数据的列表

        Returns:
            年龄字符串，如果未找到返回空字符串
        """
        candidates = []

        for data in all_data:
            df = data["df"]

            # 方法1: 扫描所有Date对象
            candidates.extend(self._extract_from_date_objects(df))

            # 方法2: 扫描Excel序列日期数字
            candidates.extend(self._extract_from_serial_dates(df))

            # 方法3: 传统的年龄标签搜索
            candidates.extend(self._extract_from_age_labels(df))

        if candidates:
            # 统计每个年龄的总置信度
            age_scores = defaultdict(float)
            for age, conf in candidates:
                age_scores[age] += conf

            # 选择置信度最高的年龄
            best_age = max(age_scores.items(), key=lambda x: x[1])
            return best_age[0]

        return ""

    def _extract_from_date_objects(self, df: pd.DataFrame) -> List[tuple]:
        """从Date对象中提取年龄"""
        candidates = []

        for idx in range(min(30, len(df))):
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if isinstance(cell, datetime):
                    if 1950 <= cell.year <= 2010:
                        age = calculate_age_from_birthdate(cell)
                        if age:
                            context_score = self._get_age_context_score(df, idx, col)
                            confidence = 2.0 + context_score * 0.5
                            candidates.append((str(age), confidence))

        return candidates

    def _extract_from_serial_dates(self, df: pd.DataFrame) -> List[tuple]:
        """从Excel序列日期中提取年龄"""
        candidates = []

        for idx in range(min(30, len(df))):
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if pd.notna(cell) and isinstance(cell, (int, float)):
                    # 检查是否可能是Excel序列日期
                    if 18000 <= cell <= 50000:
                        converted_date = convert_excel_serial_to_date(cell)
                        if converted_date:
                            age = calculate_age_from_birthdate(converted_date)
                            if age:
                                # 检查上下文
                                if self._has_age_context(df, idx, col):
                                    candidates.append((str(age), 3.0))

        return candidates

    def _extract_from_age_labels(self, df: pd.DataFrame) -> List[tuple]:
        """从年龄标签附近提取年龄"""
        candidates = []

        for idx in range(min(30, len(df))):
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if pd.notna(cell):
                    cell_str = str(cell)
                    if any(k in cell_str for k in KEYWORDS["age"]):
                        # 搜索附近的年龄值
                        nearby_ages = self._search_age_nearby(df, idx, col)
                        candidates.extend(nearby_ages)

        return candidates

    def _search_age_nearby(self, df: pd.DataFrame, row: int, col: int) -> List[tuple]:
        """在指定位置附近搜索年龄值"""
        candidates = []

        for r_offset in range(-2, 6):
            for c_offset in range(-2, 10):
                r = row + r_offset
                c = col + c_offset

                if 0 <= r < len(df) and 0 <= c < len(df.columns):
                    value = df.iloc[r, c]
                    if pd.notna(value):
                        if isinstance(value, datetime):
                            age = calculate_age_from_birthdate(value)
                            if age:
                                candidates.append((str(age), 2.5))
                        else:
                            age = self._parse_age_value(str(value))
                            if age:
                                candidates.append((age, 2.0))

        return candidates

    def _parse_age_value(self, value: str) -> Optional[str]:
        """解析年龄值"""
        value = str(value).strip()

        # 排除明显不是年龄的内容
        if len(value) > 15 or "年月" in value or "西暦" in value or "19" in value:
            return None

        # 处理"満"字格式
        match = re.search(r"満\s*(\d{1,2})\s*[歳才]", value)
        if match:
            age = int(match.group(1))
            if 18 <= age <= 65:
                return str(age)

        # 优先匹配包含"歳"或"才"的
        match = re.search(r"(\d{1,2})\s*[歳才]", value)
        if match:
            age = int(match.group(1))
            if 18 <= age <= 65:
                return str(age)

        # 尝试提取纯数字
        match = re.search(r"^(\d{1,2})$", value)
        if match:
            age = int(match.group(1))
            if 18 <= age <= 65:
                return str(age)

        return None

    def _has_age_context(self, df: pd.DataFrame, row: int, col: int) -> bool:
        """检查是否有年龄相关的上下文"""
        age_keywords = ["生年月", "年齢", "歳", "才"]

        for r_off in range(-2, 3):
            for c_off in range(-5, 5):
                r = row + r_off
                c = col + c_off
                if 0 <= r < len(df) and 0 <= c < len(df.columns):
                    context_cell = df.iloc[r, c]
                    if pd.notna(context_cell):
                        context_str = str(context_cell)
                        if any(k in context_str for k in age_keywords):
                            return True

        return False

    def _get_age_context_score(self, df: pd.DataFrame, row: int, col: int) -> float:
        """获取年龄上下文评分"""
        personal_keywords = (
            KEYWORDS["name"]
            + KEYWORDS["gender"]
            + KEYWORDS["age"]
            + ["学歴", "最終学歴", "住所", "最寄駅", "最寄", "経験", "実務"]
        )

        return self.get_context_score(df, row, col, personal_keywords)
