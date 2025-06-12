# -*- coding: utf-8 -*-
"""主提取器类"""

import pandas as pd
from typing import Dict, List
import traceback
from pathlib import Path

from extractors.name_extractor import NameExtractor
from extractors.gender_extractor import GenderExtractor
from extractors.age_extractor import AgeExtractor
from extractors.nationality_extractor import NationalityExtractor
from extractors.arrival_year_extractor import ArrivalYearExtractor
from extractors.experience_extractor import ExperienceExtractor
from extractors.japanese_level_extractor import JapaneseLevelExtractor
from extractors.skills_extractor import SkillsExtractor
from extractors.work_scope_extractor import WorkScopeExtractor
from extractors.role_extractor import RoleExtractor
from utils.text_utils import dataframe_to_text


class ResumeExtractor:
    """简历信息提取器主类"""

    def __init__(self):
        """初始化提取器"""
        self.template = {
            "name": "",
            "gender": None,
            "age": "",
            "nationality": None,
            "arrival_year_japan": None,
            "skills": [],
            "experience": "",
            "japanese_level": "",
            "work_scope": [],  # 新增作业范围字段
        }

        # 初始化各个提取器
        self.name_extractor = NameExtractor()
        self.gender_extractor = GenderExtractor()
        self.age_extractor = AgeExtractor()
        self.nationality_extractor = NationalityExtractor()
        self.arrival_year_extractor = ArrivalYearExtractor()
        self.experience_extractor = ExperienceExtractor()
        self.japanese_level_extractor = JapaneseLevelExtractor()
        self.skills_extractor = SkillsExtractor()
        self.work_scope_extractor = WorkScopeExtractor()  # 新增作业范围提取器
        self.role_extractor = RoleExtractor()  # 新增角色提取器

    def extract_from_excel(self, file_path: str) -> Dict:
        """从Excel文件提取简历信息

        Args:
            file_path: Excel文件路径（支持.xls和.xlsx格式）

        Returns:
            提取的简历信息字典
        """
        try:
            # 检查文件扩展名
            file_ext = Path(file_path).suffix.lower()

            # 根据文件类型设置引擎
            if file_ext == ".xls":
                engine = "xlrd"
                print(f"使用xlrd引擎读取xls文件")
            elif file_ext == ".xlsx":
                engine = "openpyxl"
                print(f"使用openpyxl引擎读取xlsx文件")
            else:
                # pandas会自动选择引擎
                engine = None
                print(f"自动选择引擎读取文件")

            # 读取所有sheets
            try:
                all_sheets = pd.read_excel(file_path, sheet_name=None, engine=engine)
            except ImportError as e:
                if "xlrd" in str(e):
                    return {"error": "缺少xlrd库。请运行: pip install xlrd==2.0.1"}
                elif "openpyxl" in str(e):
                    return {"error": "缺少openpyxl库。请运行: pip install openpyxl"}
                else:
                    raise

            result = self.template.copy()

            # 收集所有数据
            all_data = []
            for sheet_name, df in all_sheets.items():
                # 跳过空的sheet
                if df.empty:
                    print(f"跳过空sheet: {sheet_name}")
                    continue

                all_data.append(
                    {"sheet_name": sheet_name, "df": df, "text": dataframe_to_text(df)}
                )

            if not all_data:
                return {"error": "Excel文件中没有有效的数据"}

            # 提取各个字段
            print(f"开始提取文件: {file_path}")
            print(f"包含 {len(all_data)} 个有效sheet")

            result["name"] = self.name_extractor.extract(all_data)
            print(f"✓ 姓名: {result['name']}")

            result["gender"] = self.gender_extractor.extract(all_data)
            print(f"✓ 性别: {result['gender']}")

            result["age"] = self.age_extractor.extract(all_data)
            print(f"✓ 年龄: {result['age']}")

            result["nationality"] = self.nationality_extractor.extract(all_data)
            print(f"✓ 国籍: {result['nationality']}")

            result["arrival_year_japan"] = self.arrival_year_extractor.extract(all_data)
            print(f"✓ 来日年份: {result['arrival_year_japan']}")

            result["experience"] = self.experience_extractor.extract(all_data)
            print(f"✓ 经验: {result['experience']}")

            result["japanese_level"] = self.japanese_level_extractor.extract(all_data)
            print(f"✓ 日语: {result['japanese_level']}")

            result["skills"] = self.skills_extractor.extract(all_data)

            print(f"✓ 技能: {len(result['skills'])}个")
            result["work_scope"] = self.work_scope_extractor.extract(
                all_data
            )  # 新增作业范围提取
            print(f"✓ 作业范围: {result['work_scope']}")

            result["roles"] = self.role_extractor.extract(all_data)  # 新增角色提取
            print(f"✓ 角色: {result['roles']}")
            return result

        except Exception as e:
            print(f"处理文件时出错: {e}")
            traceback.print_exc()

            # 提供更详细的错误信息
            error_msg = str(e)
            if "Excel file format cannot be determined" in error_msg:
                error_msg = "无法识别的Excel文件格式。请确保文件是有效的.xls或.xlsx文件"
            elif "No module named" in error_msg:
                if "xlrd" in error_msg:
                    error_msg = "缺少xlrd库。请运行: pip install xlrd==2.0.1"
                elif "openpyxl" in error_msg:
                    error_msg = "缺少openpyxl库。请运行: pip install openpyxl"

            return {"error": error_msg}
