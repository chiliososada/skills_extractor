# -*- coding: utf-8 -*-
"""来日年份提取器"""

from typing import List, Dict, Any, Optional
from datetime import datetime
from collections import defaultdict
import pandas as pd
import re

from base.base_extractor import BaseExtractor
from base.constants import KEYWORDS
from utils.date_utils import convert_excel_serial_to_date


class ArrivalYearExtractor(BaseExtractor):
    """来日年份信息提取器"""

    def extract(self, all_data: List[Dict[str, Any]]) -> Optional[str]:
        """提取来日年份

        Args:
            all_data: 包含所有sheet数据的列表

        Returns:
            来日年份字符串，如果未找到返回None
        """
        candidates = []

        for data in all_data:
            df = data["df"]

            # 方法1: 扫描所有Date对象
            candidates.extend(self._extract_from_date_objects(df))

            # 方法2: 扫描Excel序列日期数字
            candidates.extend(self._extract_from_serial_dates(df))

            # 方法3: 传统的来日关键词搜索
            candidates.extend(self._extract_from_arrival_labels(df))

        if candidates:
            # 统计每个年份的总置信度
            year_scores = defaultdict(float)
            for year, conf in candidates:
                year_scores[year] += conf

            if year_scores:
                best_year = max(year_scores.items(), key=lambda x: x[1])
                return best_year[0]

        return None

    def _extract_from_date_objects(self, df: pd.DataFrame) -> List[tuple]:
        """从Date对象中提取来日年份"""
        candidates = []

        for idx in range(min(30, len(df))):
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if isinstance(cell, datetime):
                    if 1990 <= cell.year <= 2024:
                        # 检查上下文
                        has_arrival_context = self._has_arrival_context(df, idx, col)
                        has_age_context = self._has_age_context(df, idx, col)

                        if has_arrival_context:
                            # 如果也有年龄上下文，可能是生年月日，降低置信度
                            confidence = 1.5 if has_age_context else 2.5
                            candidates.append((str(cell.year), confidence))

        return candidates

    def _extract_from_serial_dates(self, df: pd.DataFrame) -> List[tuple]:
        """从Excel序列日期中提取来日年份"""
        candidates = []

        for idx in range(min(30, len(df))):
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if pd.notna(cell) and isinstance(cell, (int, float)):
                    # 检查是否可能是Excel序列日期（1982-2037年的范围）
                    if 30000 <= cell <= 50000:
                        converted_date = convert_excel_serial_to_date(cell)
                        if converted_date and 1990 <= converted_date.year <= 2024:
                            if self._has_arrival_context(df, idx, col):
                                candidates.append((str(converted_date.year), 3.0))

        return candidates

    def _extract_from_arrival_labels(self, df: pd.DataFrame) -> List[tuple]:
        """从来日标签附近提取年份"""
        candidates = []

        for idx in range(min(40, len(df))):
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if pd.notna(cell):
                    cell_str = str(cell)
                    if any(k in cell_str for k in KEYWORDS["arrival"]):
                        # 搜索附近的年份
                        nearby_years = self._search_year_nearby(df, idx, col)
                        candidates.extend(nearby_years)

        return candidates

    def _search_year_nearby(self, df: pd.DataFrame, row: int, col: int) -> List[tuple]:
        """在指定位置附近搜索年份值"""
        candidates = []

        for r_off in range(-2, 5):
            for c_off in range(-2, 25):
                r = row + r_off
                c = col + c_off

                if 0 <= r < len(df) and 0 <= c < len(df.columns):
                    value = df.iloc[r, c]
                    if pd.notna(value):
                        if isinstance(value, datetime):
                            if 1990 <= value.year <= 2024:
                                candidates.append((str(value.year), 2.0))
                        else:
                            year = self._parse_year_value(str(value))
                            if year:
                                candidates.append((year, 1.8))

        return candidates

    def _parse_year_value(self, value: str) -> Optional[str]:
        """解析年份值"""
        value_str = str(value).strip()

        # 直接的年份格式
        if re.match(r"^20\d{2}$", value_str):
            return value_str

        # 年月格式
        match = re.search(r"(20\d{2})[年/月]", value_str)
        if match:
            return match.group(1)

        # 2016年4月格式
        match = re.search(r"(20\d{2})年\d+月", value_str)
        if match:
            return match.group(1)

        # 和暦
        if "平成" in value_str:
            match = re.search(r"平成\s*(\d+)", value_str)
            if match:
                return str(1988 + int(match.group(1)))
        elif "令和" in value_str:
            match = re.search(r"令和\s*(\d+)", value_str)
            if match:
                return str(2018 + int(match.group(1)))

        return None

    def _has_arrival_context(self, df: pd.DataFrame, row: int, col: int) -> bool:
        """检查是否有来日相关的上下文"""
        return self.has_nearby_keyword(df, row, col, KEYWORDS["arrival"], radius=5)

    def _has_age_context(self, df: pd.DataFrame, row: int, col: int) -> bool:
        """检查是否有年龄相关的上下文"""
        age_keywords = ["生年月", "年齢", "歳", "才"]
        return self.has_nearby_keyword(df, row, col, age_keywords, radius=5)
