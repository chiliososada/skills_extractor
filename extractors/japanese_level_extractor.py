# -*- coding: utf-8 -*-
"""日语水平提取器"""

from typing import List, Dict, Any
import pandas as pd
import re

from base.base_extractor import BaseExtractor
from base.constants import KEYWORDS


class JapaneseLevelExtractor(BaseExtractor):
    """日语水平信息提取器"""

    def extract(self, all_data: List[Dict[str, Any]]) -> str:
        """提取日语水平

        Args:
            all_data: 包含所有sheet数据的列表

        Returns:
            日语水平字符串，如果未找到返回空字符串
        """
        candidates = []

        for data in all_data:
            text = data["text"]
            df = data["df"]

            # 搜索JLPT等级
            candidates.extend(self._extract_jlpt_levels(text))

            # 搜索其他日语水平描述
            candidates.extend(self._extract_other_levels(text))

        if candidates:
            # 按置信度排序，返回最高的
            candidates.sort(key=lambda x: x[1], reverse=True)
            return candidates[0][0]

        return ""

    def _extract_jlpt_levels(self, text: str) -> List[tuple]:
        """提取JLPT等级"""
        candidates = []

        # JLPT模式列表
        jlpt_patterns = [
            (r"JLPT\s*[NnＮ]([1-5１-５])", 2.0),
            (r"[NnＮ]([1-5１-５])\s*(?:合格|取得|レベル|級)", 1.8),
            (r"日本語能力試験\s*[NnＮ]?([1-5１-５])\s*級?", 1.5),
            (r"(?:^|\s)[NnＮ]([1-5１-５])(?:\s|$|[\(（])", 1.0),
            (r"日本語.*?([一二三四五])級", 1.3),
            (r"([一二三四五])級.*?日本語", 1.3),
        ]

        for pattern, confidence in jlpt_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE | re.MULTILINE)
            for match in matches:
                level_str = match.group(1)

                # 汉字数字转换
                kanji_to_num = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5"}

                if level_str in kanji_to_num:
                    level_num = kanji_to_num[level_str]
                else:
                    level_num = level_str.translate(self.trans_table)

                level = f"N{level_num}"
                candidates.append((level, confidence))

        return candidates

    def _extract_other_levels(self, text: str) -> List[tuple]:
        """提取其他日语水平描述"""
        candidates = []

        # 检查商务级别
        if "ビジネス" in text and any(jp in text for jp in ["日本語", "日語"]):
            candidates.append(("ビジネスレベル", 0.8))

        # 检查上级
        elif "上級" in text and any(jp in text for jp in ["日本語", "日語"]):
            candidates.append(("上級", 0.8))

        # 检查中级
        elif "中級" in text and any(jp in text for jp in ["日本語", "日語"]):
            candidates.append(("中級", 0.7))

        return candidates
