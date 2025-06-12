# -*- coding: utf-8 -*-
"""角色（役割）提取器"""

from typing import List, Dict, Any, Set, Optional
import pandas as pd
import re

from base.base_extractor import BaseExtractor


class RoleExtractor(BaseExtractor):
    """角色信息提取器"""

    def __init__(self):
        super().__init__()
        # 工程阶段关键词（用于定位作业范围）
        self.design_keywords = [
            "基本設計",
            "詳細設計",
            "製造",
            "単体テスト",
            "結合テスト",
            "総合テスト",
            "運用保守",
            "要件定義",
            "作業範囲",
        ]

        # 角色关键词
        self.role_keywords = ["SE", "PG", "PL", "PM", "TL"]

        # 角色级别映射（数字越大，级别越高）
        self.role_levels = {
            "PM": 5,  # Project Manager - 最高级别
            "PL": 4,  # Project Leader
            "TL": 3,  # Team Leader
            "SE": 2,  # System Engineer
            "PG": 1,  # Programmer - 最低级别
        }

    def extract(self, all_data: List[Dict[str, Any]]) -> List[str]:
        """提取角色

        Args:
            all_data: 包含所有sheet数据的列表

        Returns:
            角色列表（按级别排序）
        """
        all_roles = set()

        for data in all_data:
            df = data["df"]
            sheet_name = data.get("sheet_name", "Unknown")

            print(f"\n🔍 开始角色提取 - Sheet: {sheet_name}")
            print(f"    表格大小: {df.shape[0]}行 x {df.shape[1]}列")

            # 从底部向上查找作业范围
            design_row = self._find_design_row_from_bottom(df)

            if design_row is not None:
                print(f"    找到作业范围行: 行{design_row}")

                # 找到角色所在的列
                role_column = self._find_role_column_from_design_row(df, design_row)

                if role_column is not None:
                    print(f"    找到角色列: 列{role_column}")

                    # 在该列中提取所有角色
                    roles = self._extract_roles_from_column(df, role_column, design_row)
                    all_roles.update(roles)

                    if roles:
                        print(f"    ✓ 发现角色: {roles}")
                else:
                    print("    未找到角色列")
            else:
                print("    未找到作业范围行")

        # 按照职位级别排序
        sorted_roles = self._sort_roles_by_level(list(all_roles))

        print(f"\n✅ 最终提取的角色: {sorted_roles}")
        return sorted_roles

    def _find_design_row_from_bottom(self, df: pd.DataFrame) -> Optional[int]:
        """从底部向上查找包含作业范围的行"""
        # 从底部向上遍历
        for row in range(len(df) - 1, -1, -1):
            row_text = ""
            design_count = 0

            # 检查该行的所有列
            for col in range(len(df.columns)):
                cell = df.iloc[row, col]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    row_text += " " + cell_str

                    # 计算包含的工程阶段关键词数量
                    for keyword in self.design_keywords:
                        if keyword in cell_str:
                            design_count += 1
                            break

            # 如果该行包含多个工程阶段关键词或包含"作業範囲"，认为找到了
            if design_count >= 3 or "作業範囲" in row_text:
                return row

        return None

    def _find_role_column_from_design_row(
        self, df: pd.DataFrame, design_row: int
    ) -> Optional[int]:
        """从作业范围行开始，从右向左查找包含角色的列"""
        # 从最右侧开始向左查找
        for col in range(len(df.columns) - 1, -1, -1):
            # 从作业范围行开始向下查找一定范围
            for row_offset in range(1, min(20, len(df) - design_row)):
                row = design_row + row_offset
                cell = df.iloc[row, col]

                if pd.notna(cell):
                    cell_str = str(cell).strip()

                    # 检查是否包含角色关键词
                    if self._cell_contains_role(cell_str):
                        return col

        return None

    def _cell_contains_role(self, cell_str: str) -> bool:
        """检查单元格是否包含角色关键词"""
        # 检查每个角色关键词
        for role in self.role_keywords:
            # 使用正则表达式确保是独立的角色词
            patterns = [
                rf"^{role}$",  # 完全匹配
                rf"(?:^|[^A-Za-z]){role}(?:[^A-Za-z]|$)",  # 前后非字母
            ]

            for pattern in patterns:
                if re.search(pattern, cell_str, re.IGNORECASE):
                    return True

        return False

    def _extract_roles_from_column(
        self, df: pd.DataFrame, role_column: int, start_row: int
    ) -> Set[str]:
        """从指定列中提取所有角色"""
        roles = set()

        # 从start_row开始向下查找
        for row in range(start_row + 1, len(df)):
            cell = df.iloc[row, role_column]
            if pd.notna(cell):
                cell_str = str(cell).strip()

                # 提取角色
                for role in self.role_keywords:
                    patterns = [
                        rf"^{role}$",  # 完全匹配
                        rf"(?:^|[^A-Za-z]){role}(?:[^A-Za-z]|$)",  # 前后非字母
                    ]

                    for pattern in patterns:
                        if re.search(pattern, cell_str, re.IGNORECASE):
                            roles.add(role.upper())
                            break

        return roles

    def _sort_roles_by_level(self, roles: List[str]) -> List[str]:
        """按照职位级别排序（从高到低）"""
        # 确保所有角色都是大写
        roles = [role.upper() for role in roles]

        # 去重
        roles = list(set(roles))

        # 按照级别排序
        sorted_roles = sorted(
            roles,
            key=lambda x: self.role_levels.get(x, 0),
            reverse=True,  # 降序，级别高的在前
        )

        return sorted_roles
