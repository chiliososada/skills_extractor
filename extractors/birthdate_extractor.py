# -*- coding: utf-8 -*-
"""出生年月日提取器 - 改进版"""

from typing import List, Dict, Any, Optional, Tuple, Set
from datetime import datetime
import pandas as pd
import re
import numpy as np

from base.base_extractor import BaseExtractor
from utils.date_utils import convert_excel_serial_to_date


class BirthdateExtractor(BaseExtractor):
    """出生年月日信息提取器"""

    def __init__(self):
        super().__init__()
        # 生年月日关键词
        self.birthdate_keywords = [
            "生年月日",
            "生年月",
            "生年",
            "誕生日",
            "出生日期",
            "出生年月日",
            "Birth Date",
            "Birthday",
            "DOB",
            "Date of Birth",
        ]

        # 和暦转换表
        self.wareki_to_seireki = {
            "明治": 1867,  # 明治元年 = 1868
            "大正": 1911,  # 大正元年 = 1912
            "昭和": 1925,  # 昭和元年 = 1926
            "平成": 1988,  # 平成元年 = 1989
            "令和": 2018,  # 令和元年 = 2019
        }

    def extract(self, all_data: List[Dict[str, Any]]) -> Optional[str]:
        """提取出生年月日

        Args:
            all_data: 包含所有sheet数据的列表

        Returns:
            出生年月日字符串（yyyy-mm-dd格式），如果未找到返回None
        """
        for data in all_data:
            df = data["df"]
            sheet_name = data.get("sheet_name", "Unknown")

            print(f"\n🔍 开始出生年月日提取 - Sheet: {sheet_name}")
            print(f"    表格大小: {df.shape[0]}行 x {df.shape[1]}列")

            # 查找生年月关键字的位置
            keyword_positions = self._find_birthdate_keyword_positions(df)

            if not keyword_positions:
                print("    未找到生年月关键字")
                continue

            print(f"    找到 {len(keyword_positions)} 个生年月关键字位置")

            # 对每个关键字位置进行搜索
            for pos in keyword_positions:
                print(
                    f"\n    检查位置: 行{pos['row']}, 列{pos['col']}, 内容: '{pos['value']}'"
                )

                # 检查是否是合并单元格
                merged_info = self._check_merged_cell(df, pos["row"], pos["col"])
                if merged_info:
                    print(
                        f"      检测到合并单元格: 列{merged_info['start_col']}到列{merged_info['end_col']}"
                    )

                # 向下搜索年份
                year_result = self._search_year_below(df, pos, merged_info)

                if year_result:
                    # 根据找到的年份构建完整日期
                    # 这里简单处理，如果只有年份，设置为1月1日
                    year = year_result["year"]
                    month = year_result.get("month", 1)
                    day = year_result.get("day", 1)

                    try:
                        date_obj = datetime(year, month, day)
                        date_str = date_obj.strftime("%Y-%m-%d")

                        # 验证日期合理性
                        if self._validate_birthdate(date_str):
                            print(f"\n✅ 成功提取出生年月日: {date_str}")
                            return date_str
                        else:
                            print(f"    日期验证失败: {date_str}")
                    except ValueError:
                        print(f"    无法构建有效日期: {year}-{month}-{day}")

        print("\n❌ 未能提取到出生年月日")
        return None

    def _find_birthdate_keyword_positions(self, df: pd.DataFrame) -> List[Dict]:
        """查找生年月关键字的位置"""
        positions = []

        # 扫描整个表格查找关键字
        for row in range(len(df)):
            for col in range(len(df.columns)):
                cell = df.iloc[row, col]
                if pd.notna(cell):
                    cell_str = str(cell).strip()

                    # 检查是否包含生年月关键词
                    for keyword in self.birthdate_keywords:
                        if keyword in cell_str:
                            positions.append(
                                {
                                    "row": row,
                                    "col": col,
                                    "value": cell_str,
                                    "keyword": keyword,
                                }
                            )
                            print(f"      发现关键字 '{keyword}' 在: 行{row}, 列{col}")
                            break

        return positions

    def _check_merged_cell(
        self, df: pd.DataFrame, row: int, col: int
    ) -> Optional[Dict]:
        """检查是否是合并单元格

        注意：pandas读取Excel时不会保留合并单元格信息，
        这里通过检查相邻单元格是否为空来推断合并情况
        """
        # 检查右侧是否有连续的空单元格（可能是合并的）
        start_col = col
        end_col = col

        # 向右检查
        for c in range(col + 1, min(col + 10, len(df.columns))):  # 最多检查右侧10列
            if pd.isna(df.iloc[row, c]):
                # 如果是空的，可能是合并单元格的一部分
                # 但需要检查下一行对应位置是否有值
                if row + 1 < len(df):
                    if pd.notna(df.iloc[row + 1, c]):
                        # 下一行有值，说明不是合并单元格
                        break
                end_col = c
            else:
                # 遇到非空单元格，停止
                break

        # 如果发现可能的合并单元格
        if end_col > start_col:
            return {
                "start_col": start_col,
                "end_col": end_col,
                "width": end_col - start_col + 1,
            }

        return None

    def _search_year_below(
        self, df: pd.DataFrame, keyword_pos: Dict, merged_info: Optional[Dict]
    ) -> Optional[Dict]:
        """在关键字下方搜索年份信息"""
        start_row = keyword_pos["row"] + 1  # 从下一行开始
        start_col = keyword_pos["col"]

        # 确定搜索的列范围
        if merged_info:
            # 如果是合并单元格，搜索整个合并范围
            col_range = range(merged_info["start_col"], merged_info["end_col"] + 1)
        else:
            # 否则只搜索当前列和右侧几列
            col_range = range(start_col, min(start_col + 5, len(df.columns)))

        print(f"      搜索范围: 行{start_row}开始, 列{list(col_range)}")

        # 向下搜索（最多搜索20行）
        for row_offset in range(20):
            row = start_row + row_offset
            if row >= len(df):
                break

            # 搜索指定的列范围
            for col in col_range:
                cell = df.iloc[row, col]

                if pd.notna(cell):
                    # 尝试提取年份
                    year_info = self._extract_year_from_cell(cell, row, col)

                    if year_info:
                        print(f"      ✓ 在行{row}, 列{col}找到年份: {year_info}")
                        return year_info

            # 如果是合并单元格，还要检查合并后的整行文本
            if merged_info:
                # 收集整行合并范围内的文本
                row_text = ""
                for col in col_range:
                    if pd.notna(df.iloc[row, col]):
                        row_text += str(df.iloc[row, col]).strip() + " "

                if row_text.strip():
                    # 尝试从合并的文本中提取年份
                    year_info = self._extract_year_from_text(row_text.strip())
                    if year_info:
                        print(f"      ✓ 在行{row}的合并文本中找到年份: {year_info}")
                        return year_info

        return None

    def _extract_year_from_cell(self, cell, row: int, col: int) -> Optional[Dict]:
        """从单元格中提取年份信息"""
        # 处理日期对象
        if isinstance(cell, (datetime, pd.Timestamp)):
            if 1950 <= cell.year <= 2010:
                return {
                    "year": cell.year,
                    "month": cell.month,
                    "day": cell.day,
                    "source": "datetime_object",
                }

        # 处理Excel序列日期
        if isinstance(cell, (int, float)) and 18000 <= cell <= 50000:
            converted_date = convert_excel_serial_to_date(cell)
            if converted_date and 1950 <= converted_date.year <= 2010:
                return {
                    "year": converted_date.year,
                    "month": converted_date.month,
                    "day": converted_date.day,
                    "source": "excel_serial",
                }

        # 处理文本
        if isinstance(cell, str) or pd.api.types.is_string_dtype(type(cell)):
            return self._extract_year_from_text(str(cell))

        return None

    def _extract_year_from_text(self, text: str) -> Optional[Dict]:
        """从文本中提取年份信息"""
        text = str(text).strip()

        # 1. 完整日期格式
        date_patterns = [
            # yyyy年mm月dd日
            (
                r"(19[5-9]\d|20[0-1]\d)[年/](0?[1-9]|1[0-2])[月/](0?[1-9]|[12]\d|3[01])日?",
                lambda m: {
                    "year": int(m.group(1)),
                    "month": int(m.group(2)),
                    "day": int(m.group(3)),
                    "source": "text_ymd",
                },
            ),
            # yyyy.mm.dd 或 yyyy-mm-dd 或 yyyy/mm/dd
            (
                r"(19[5-9]\d|20[0-1]\d)[\.\-/](0?[1-9]|1[0-2])[\.\-/](0?[1-9]|[12]\d|3[01])",
                lambda m: {
                    "year": int(m.group(1)),
                    "month": int(m.group(2)),
                    "day": int(m.group(3)),
                    "source": "text_numeric",
                },
            ),
        ]

        for pattern, parser in date_patterns:
            match = re.search(pattern, text)
            if match:
                return parser(match)

        # 2. 和暦格式
        for era, base_year in self.wareki_to_seireki.items():
            # 和暦带日期
            pattern = rf"{era}\s*(\d{{1,2}})[年/]\s*(0?[1-9]|1[0-2])[月/]\s*(0?[1-9]|[12]\d|3[01])日?"
            match = re.search(pattern, text)
            if match:
                year = base_year + int(match.group(1))
                if 1950 <= year <= 2010:
                    return {
                        "year": year,
                        "month": int(match.group(2)),
                        "day": int(match.group(3)),
                        "source": f"wareki_{era}",
                    }

            # 和暦只有年
            pattern = rf"{era}\s*(\d{{1,2}})年"
            match = re.search(pattern, text)
            if match:
                year = base_year + int(match.group(1))
                if 1950 <= year <= 2010:
                    return {"year": year, "source": f"wareki_{era}_year_only"}

        # 3. 只有4位年份
        match = re.search(r"\b(19[5-9]\d|20[0-1]\d)\b", text)
        if match:
            year = int(match.group(1))
            return {"year": year, "source": "year_only"}

        return None

    def _validate_birthdate(self, date_str: str) -> bool:
        """验证出生日期是否合理"""
        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")

            # 检查年份范围
            if not (1950 <= date_obj.year <= 2010):
                return False

            # 检查是否是未来日期
            if date_obj > datetime.now():
                return False

            # 检查年龄是否在合理范围内（14-74岁）
            age = datetime.now().year - date_obj.year
            if not (14 <= age <= 74):
                return False

            return True
        except ValueError:
            return False
