# -*- coding: utf-8 -*-
"""主提取器类 - 修复版：统一返回null"""

import pandas as pd
from typing import Dict, List, Optional, Any
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
from extractors.birthdate_extractor import BirthdateExtractor
from utils.text_utils import dataframe_to_text


class ResumeExtractor:
    """简历信息提取器主类 - 修复版"""

    def __init__(self):
        """初始化提取器"""
        # 修改模板：所有字段默认为None
        self.template = {
            "name": None,
            "gender": None,
            "age": None,
            "birthdate": None,
            "nationality": None,
            "arrival_year_japan": None,
            "skills": None,
            "experience": None,
            "japanese_level": None,
            "work_scope": None,
            "roles": None,
        }

        # 初始化各个提取器
        self.name_extractor = NameExtractor()
        self.gender_extractor = GenderExtractor()
        self.age_extractor = AgeExtractor()
        self.birthdate_extractor = BirthdateExtractor()
        self.nationality_extractor = NationalityExtractor()
        self.arrival_year_extractor = ArrivalYearExtractor()
        self.experience_extractor = ExperienceExtractor()
        self.japanese_level_extractor = JapaneseLevelExtractor()
        self.skills_extractor = SkillsExtractor()
        self.work_scope_extractor = WorkScopeExtractor()
        self.role_extractor = RoleExtractor()

    def _normalize_result(self, value: Any) -> Optional[Any]:
        """标准化提取结果

        Args:
            value: 提取的原始值

        Returns:
            标准化后的值：
            - 空字符串 -> None
            - 空列表 -> None
            - 其他空值 -> None
            - 有效值 -> 原值
        """
        if value is None:
            return None

        # 处理字符串
        if isinstance(value, str):
            value = value.strip()
            return value if value else None

        # 处理列表
        if isinstance(value, list):
            return value if value else None

        # 其他类型直接返回
        return value

    def extract_from_excel(self, file_path: str) -> Dict:
        """从Excel文件提取简历信息 - 修复版

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

            # 按优化顺序提取各个字段
            print(f"开始提取文件: {file_path}")
            print(f"包含 {len(all_data)} 个有效sheet")

            # 基本信息
            name_result = self.name_extractor.extract(all_data)
            result["name"] = self._normalize_result(name_result)
            print(f"✓ 姓名: {result['name']}")

            gender_result = self.gender_extractor.extract(all_data)
            result["gender"] = self._normalize_result(gender_result)
            print(f"✓ 性别: {result['gender']}")

            # 先提取生年月日，再用它来计算年龄
            birthdate_result = self.birthdate_extractor.extract(all_data)
            result["birthdate"] = self._normalize_result(birthdate_result)
            print(f"✓ 出生年月日: {result['birthdate']}")

            # 修复：年龄提取器现在可以接受生年月日参数
            age_result = self.age_extractor.extract(all_data, result["birthdate"])
            result["age"] = self._normalize_result(age_result)
            print(f"✓ 年龄: {result['age']}")

            nationality_result = self.nationality_extractor.extract(all_data)
            result["nationality"] = self._normalize_result(nationality_result)
            print(f"✓ 国籍: {result['nationality']}")

            # 修复：来日年份提取器现在可以排除出生年份
            arrival_result = self.arrival_year_extractor.extract(
                all_data, result["birthdate"]
            )
            result["arrival_year_japan"] = self._normalize_result(arrival_result)
            print(f"✓ 来日年份: {result['arrival_year_japan']}")

            experience_result = self.experience_extractor.extract(all_data)
            result["experience"] = self._normalize_result(experience_result)
            print(f"✓ 经验: {result['experience']}")

            # 修复：使用改进的日语水平提取器
            japanese_result = self.japanese_level_extractor.extract(all_data)
            result["japanese_level"] = self._normalize_result(japanese_result)
            print(f"✓ 日语: {result['japanese_level']}")

            skills_result = self.skills_extractor.extract(all_data)
            result["skills"] = self._normalize_result(skills_result)
            print(f"✓ 技能: {len(skills_result) if skills_result else 0}个")

            work_scope_result = self.work_scope_extractor.extract(all_data)
            result["work_scope"] = self._normalize_result(work_scope_result)
            print(f"✓ 作业范围: {result['work_scope']}")

            roles_result = self.role_extractor.extract(all_data)
            result["roles"] = self._normalize_result(roles_result)
            print(f"✓ 角色: {result['roles']}")

            # 后处理：如果某些字段仍然有问题，进行最后修复
            result = self._post_process_result(result)

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

    def _post_process_result(self, result: Dict) -> Dict:
        """后处理结果，进行最终修复"""
        print("\n🔧 后处理阶段...")

        # 修复1：如果年龄仍为空但有生年月日，计算年龄
        if result.get("age") is None and result.get("birthdate"):
            try:
                from datetime import datetime

                birthdate = datetime.strptime(result["birthdate"], "%Y-%m-%d")
                current_date = datetime.now()
                age = current_date.year - birthdate.year
                if (current_date.month, current_date.day) < (
                    birthdate.month,
                    birthdate.day,
                ):
                    age -= 1
                if 15 <= age <= 80:
                    result["age"] = str(age)
                    print(f"    修复年龄: {result['age']}")
            except:
                pass

        # 修复2：如果来日年份等于出生年份，清空（明显错误）
        if (
            result.get("arrival_year_japan")
            and result.get("birthdate")
            and result["arrival_year_japan"] == result["birthdate"][:4]
        ):
            print(f"    发现错误的来日年份（等于出生年份），清空")
            result["arrival_year_japan"] = None

        # 修复3：如果日语水平为空但经验丰富，给出合理推测
        if (
            result.get("japanese_level") is None
            and result.get("experience")
            and any(char.isdigit() for char in result["experience"])
        ):
            try:
                # 提取经验年数
                import re

                exp_match = re.search(r"(\d+(?:\.\d+)?)", result["experience"])
                if exp_match:
                    exp_years = float(exp_match.group(1))
                    if exp_years >= 5:
                        result["japanese_level"] = "N2以上"
                        print(f"    根据经验推测日语水平: {result['japanese_level']}")
            except:
                pass

        # 修复4：技能去重和清理
        if result.get("skills"):
            # 去重（保持顺序）
            seen = set()
            unique_skills = []
            for skill in result["skills"]:
                if skill.lower() not in seen:
                    seen.add(skill.lower())
                    unique_skills.append(skill)

            if len(unique_skills) != len(result["skills"]):
                print(f"    技能去重: {len(result['skills'])} → {len(unique_skills)}")
                result["skills"] = unique_skills

        # 最终确保所有空值都是None
        for key, value in result.items():
            result[key] = self._normalize_result(value)

        print("✅ 后处理完成")
        return result
