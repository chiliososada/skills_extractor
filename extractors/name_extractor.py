# -*- coding: utf-8 -*-
"""姓名提取器"""

from typing import List, Dict, Any
import pandas as pd
import re

from base.base_extractor import BaseExtractor
from base.constants import KEYWORDS
from utils.validation_utils import is_valid_name


class NameExtractor(BaseExtractor):
    """姓名信息提取器"""

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

            # 搜索姓名关键词
            for idx in range(min(20, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell) and any(k in str(cell) for k in KEYWORDS["name"]):
                        # 在附近搜索姓名值
                        candidates.extend(self._search_name_nearby(df, idx, col))

        if candidates:
            # 按置信度排序，返回最佳候选
            candidates.sort(key=lambda x: x[1], reverse=True)
            return candidates[0][0]

        return ""

    def _search_name_nearby(self, df: pd.DataFrame, row: int, col: int) -> List[tuple]:
        """在指定位置附近搜索姓名

        Args:
            df: DataFrame对象
            row: 行索引
            col: 列索引

        Returns:
            候选姓名列表，格式为 [(姓名, 置信度), ...]
        """
        candidates = []

        # 搜索范围
        for r_offset in range(-2, 5):
            for c_offset in range(-2, 15):
                r = row + r_offset
                c = col + c_offset

                if 0 <= r < len(df) and 0 <= c < len(df.columns):
                    value = df.iloc[r, c]
                    if pd.notna(value) and is_valid_name(str(value)):
                        confidence = 1.0

                        # 同行的置信度更高
                        if r == row:
                            confidence *= 1.5

                        # 右侧邻近的置信度更高
                        if c > col and c - col <= 5:
                            confidence *= 1.3

                        candidates.append((str(value).strip(), confidence))

        return candidates
