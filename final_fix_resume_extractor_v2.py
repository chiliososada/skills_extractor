# final_fix_resume_extractor_v2.py - 优化合并单元格中的技能提取
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
import re
import json
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Union, Tuple, Set
from collections import defaultdict, Counter


class FinalFixResumeExtractor:
    """最终修复版简历提取器 - V2版优化合并单元格技能提取"""

    def __init__(self):
        self.template = {
            "name": "",
            "gender": None,
            "age": "",
            "nationality": None,
            "arrival_year_japan": None,
            "skills": [],
            "experience": "",
            "japanese_level": "",
        }

        # 全角转半角的转换表
        self.trans_table = str.maketrans("０１２３４５６７８９", "0123456789")

        # 扩展的技能标记符号
        self.skill_marks = ["◎", "○", "△", "×", "★", "●", "◯", "▲", "※"]

        # Excel日期起始点：1900年1月1日
        self.excel_epoch = datetime(1900, 1, 1)

        # 真正的技能关键词（标准化后的名称）
        self.valid_skills = {
            # 编程语言
            "Java",
            "JAVA",
            "Python",
            "JavaScript",
            "TypeScript",
            "C",
            "C++",
            "C#",
            "C/C++",
            "PHP",
            "Ruby",
            "Go",
            "Kotlin",
            "Swift",
            "Scala",
            "Rust",
            "VB.NET",
            "VB",
            "VBA",
            "COBOL",
            "Perl",
            "R",
            "Objective-C",
            # 前端
            "HTML",
            "HTML5",
            "CSS",
            "CSS3",
            "React",
            "React.js",
            "Vue",
            "Vue.js",
            "Angular",
            "jQuery",
            "Bootstrap",
            "Webpack",
            "Sass",
            "Less",
            # 后端框架
            "Spring",
            "SpringBoot",
            "SpringMVC",
            "Struts",
            "Struts2",
            "Django",
            "Flask",
            "Rails",
            "Express",
            "Node.js",
            ".NET",
            "ASP.NET",
            "Laravel",
            "Thymeleaf",
            "JSF",
            "JSP",
            "Servlet",
            "Java Servlet",
            "Hibernate",
            "MyBatis",
            "Mybatis",
            "JPA",
            # 数据库
            "MySQL",
            "PostgreSQL",
            "Oracle",
            "SQL Server",
            "SQLServer",
            "MongoDB",
            "Redis",
            "DB2",
            "SQLite",
            "Access",
            "Sybase",
            "Azure SQL DB",
            "PL/SQL",
            "Aurora",
            # 工具和平台
            "AWS",
            "Azure",
            "GCP",
            "Docker",
            "Kubernetes",
            "Git",
            "GitHub",
            "GitLab",
            "SVN",
            "TortoiseSVN",
            "Jenkins",
            "Maven",
            "Gradle",
            # OS
            "Windows",
            "Windows 10",
            "Linux",
            "Unix",
            "Solaris",
            "DOS/V",
            "Win95/98",
            "win10",
            "Ubuntu",
            "CentOS",
            "RedHat",
            # 服务器
            "Apache",
            "Nginx",
            "Tomcat",
            "WebSphere",
            "JBoss",
            "JBOSS",
            "IIS",
            # IDE和工具
            "Eclipse",
            "IntelliJ",
            "VS Code",
            "Visual Studio",
            "Android Studio",
            "NetBeans",
            "Xcode",
            "A5M2",
            "WinMerge",
            "WinSCP",
            "Sourcetree",
            "Postman",
            "Fiddler",
            "Charles",
            "Form Designer",
            # 测试工具
            "JUnit",
            "Junit",
            "Selenium",
            "JMeter",
            "Jmeter",
            "Spock",
            # 其他技术
            "XML",
            "JSON",
            "REST",
            "SOAP",
            "Ajax",
            "Microservices",
            "Shell",
            "Bash",
            "PowerShell",
            "VBScript",
            # 移动开发
            "Android",
            "iOS",
            "React Native",
            "Flutter",
            # 其他框架和库
            "TERASOLUNA",
            "OutSystems",
            "Wacs",
            # Web相关
            "PHP/HTML",
            "Jquery",
            "JAVASCRIPT",
        }

        # 需要排除的非技能内容
        self.exclude_patterns = [
            r"^\d{4}[-/]\d{2}[-/]\d{2}",
            r"^\d{2,4}年\d{1,2}月",
            r"\d+人",
            r"人以[下上]",
            r"^(基本|詳細)設計$",
            r"^製造$",
            r"^(単体|結合|総合|運用)[試験テスト]",
            r"^保守運用$",
            r"^要件定義$",
            r"^(SE|PG|PL|PM|TL)$",
            r"^管理$",
            r"管理経験",
            r"^教\s*育$",
            r"業務経歴",
            r"プロジェクト",
            r"^その他$",
            r"^過去の",
            r"【.*】",
            r"^[A-BD-Z]$",
            r"^携帯$",
            r"^E-mail$",
        ]

        # 关键词定义
        self.keywords = {
            "name": ["氏名", "氏 名", "名前", "フリガナ", "Name", "名　前", "姓名"],
            "age": [
                "年齢",
                "年龄",
                "年令",
                "歳",
                "才",
                "Age",
                "年　齢",
                "生年月",
                "満",
            ],
            "gender": ["性別", "性别", "Gender", "性　別"],
            "nationality": ["国籍", "出身国", "出身地", "Nationality", "国　籍"],
            "experience": [
                "経験年数",
                "実務経験",
                "開発経験",
                "ソフト関連業務経験年数",
                "IT経験",
                "業務経験",
                "経験",
                "実務年数",
                "Experience",
                "エンジニア経験",
                "経験年月",
                "職歴",
                "IT経験年数",
                "コンピュータソフトウエア関連業務",
                "実務年数",
            ],
            "arrival": [
                "来日",
                "渡日",
                "入国",
                "日本滞在年数",
                "滞在年数",
                "在日年数",
                "来日年",
                "来日時期",
                "来日年月",
                "来日年度",
            ],
            "japanese": [
                "日本語",
                "日語",
                "JLPT",
                "日本語能力",
                "語学力",
                "言語能力",
                "日本語レベル",
                "Japanese",
            ],
            "education": [
                "学歴",
                "学校",
                "大学",
                "卒業",
                "専門学校",
                "高校",
                "最終学歴",
            ],
            "skills": [
                "技術",
                "スキル",
                "言語",
                "DB",
                "OS",
                "フレームワーク",
                "ツール",
                "Skills",
                "开发语言",
                "プログラミング言語",
                "データベース",
                "開発環境",
                "技術経験",
            ],
        }

    def extract_from_excel(self, file_path: str) -> Dict:
        """从Excel文件提取简历信息"""
        try:
            # 读取所有sheets
            all_sheets = pd.read_excel(file_path, sheet_name=None)

            result = self.template.copy()

            # 收集所有数据
            all_data = []
            for sheet_name, df in all_sheets.items():
                all_data.append(
                    {
                        "sheet_name": sheet_name,
                        "df": df,
                        "text": self._dataframe_to_text(df),
                    }
                )

            # 提取各个字段
            result["name"] = self._extract_name(all_data)
            result["gender"] = self._extract_gender(all_data)
            result["age"] = self._extract_age_final_fix(all_data)
            result["nationality"] = self._extract_nationality_final_fix(all_data)
            result["arrival_year_japan"] = self._extract_arrival_year_final_fix(
                all_data
            )
            result["experience"] = self._extract_experience_final_fix(all_data)
            result["japanese_level"] = self._extract_japanese_level(all_data)
            result["skills"] = self._extract_skills_v2(all_data)  # 使用V2版本

            return result

        except Exception as e:
            print(f"处理文件时出错: {e}")
            import traceback

            traceback.print_exc()
            return {"error": str(e)}

    def _extract_skills_v2(self, all_data: List[Dict]) -> List[str]:
        """基于表格结构的技术关键字提取V2版 - 优化合并单元格处理"""
        all_skills = []

        for data in all_data:
            df = data["df"]
            sheet_name = data.get("sheet_name", "Unknown")

            print(f"\n🔍 开始技术关键字提取V2 - Sheet: {sheet_name}")
            print(f"    表格大小: {df.shape[0]}行 x {df.shape[1]}列")

            # 方法1: 直接搜索包含多个技能的合并单元格
            merged_cell_skills = self._find_skills_in_merged_cells(df)
            if merged_cell_skills:
                print(f"    ✓ 从合并单元格提取到 {len(merged_cell_skills)} 个技能")
                all_skills.extend(merged_cell_skills)

            # 方法2: 基于设计列的结构化搜索
            design_positions = self._find_design_columns(df)
            if design_positions:
                print(f"    ✓ 找到设计列位置: {len(design_positions)}个")

                # 在设计列附近查找技能
                skill_start_position = self._find_skill_start_row_v2(
                    df, design_positions
                )
                if skill_start_position:
                    print(
                        f"    ✓ 技术关键字起始位置: 行{skill_start_position['row']}, 列{skill_start_position['col']}"
                    )
                    skills = self._extract_skills_from_column_v2(
                        df, skill_start_position
                    )
                    all_skills.extend(skills)

            # 方法3: 备用方法
            if len(all_skills) < 3:  # 如果提取的技能太少，使用备用方法
                print(f"    使用备用方法补充技能提取")
                fallback_skills = self._extract_skills_fallback(df)
                all_skills.extend(fallback_skills)

        # 去重和标准化
        final_skills = self._process_and_deduplicate_skills(all_skills)
        return final_skills

    def _find_skills_in_merged_cells(self, df: pd.DataFrame) -> List[str]:
        """专门查找包含多个技能的合并单元格"""
        skills = []
        print(f"    查找合并单元格中的技能...")

        # 遍历前50行的所有单元格
        for row in range(min(50, len(df))):
            for col in range(len(df.columns)):
                cell = df.iloc[row, col]
                if pd.notna(cell):
                    cell_str = str(cell).strip()

                    # 检查是否包含换行符（合并单元格的特征）
                    if "\n" in cell_str:
                        lines = cell_str.split("\n")

                        # 统计包含技能的行数
                        skill_lines = []
                        for line in lines:
                            line = line.strip()
                            if line and self._line_contains_skill(line):
                                skill_lines.append(line)

                        # 如果超过3行包含技能，很可能是技能列表
                        if len(skill_lines) >= 3:
                            print(
                                f"      发现技能合并单元格 [{row},{col}]: {len(skill_lines)}个技能行"
                            )

                            # 提取所有技能
                            for line in lines:
                                line_skills = self._extract_skills_from_line(line)
                                skills.extend(line_skills)

                            # 打印前几个技能
                            if skills:
                                print(f"      提取到的技能: {', '.join(skills[:5])}...")

        return skills

    def _line_contains_skill(self, line: str) -> bool:
        """检查一行是否包含技能"""
        line = line.strip().upper()

        # 快速检查是否包含常见技能关键词
        skill_keywords = [
            "JAVA",
            "PYTHON",
            "SCRIPT",
            "HTML",
            "CSS",
            "SQL",
            "SPRING",
            "REACT",
            "VUE",
            "ANGULAR",
            "NODE",
            "PHP",
            "RUBY",
            "KOTLIN",
            "SWIFT",
            "LINUX",
            "WINDOWS",
            "DOCKER",
        ]

        for keyword in skill_keywords:
            if keyword in line:
                return True

        # 检查是否匹配有效技能
        for skill in self.valid_skills:
            if skill.upper() in line:
                return True

        return False

    def _extract_skills_from_line(self, line: str) -> List[str]:
        """从一行文本中提取技能"""
        skills = []
        line = line.strip()

        # 移除开头的标记符号
        line = re.sub(r"^[◎○△×★●◯▲※・\-\s]+", "", line)

        if not line:
            return []

        # 尝试多种分隔符
        items = re.split(r"[、,，/／\s]+", line)

        for item in items:
            item = item.strip()
            if item and self._is_valid_skill_v2(item):
                normalized = self._normalize_skill_name(item)
                skills.append(normalized)

        # 如果没有找到技能，但行本身可能就是一个技能
        if not skills and line and self._is_valid_skill_v2(line):
            normalized = self._normalize_skill_name(line)
            skills.append(normalized)

        return skills

    def _find_skill_start_row_v2(
        self, df: pd.DataFrame, design_positions: List[Dict]
    ) -> Optional[Dict]:
        """查找技术关键字起始行 - V2版本"""
        if not design_positions:
            return None

        # 获取设计列的位置信息
        design_rows = [pos["row"] for pos in design_positions]
        design_cols = [pos["col"] for pos in design_positions]

        min_design_row = min(design_rows)
        max_design_row = max(design_rows)
        min_design_col = min(design_cols)

        print(
            f"    设计列范围: 行{min_design_row}-{max_design_row}, 最左列{min_design_col}"
        )

        # 扩大搜索范围
        candidates = []

        # 搜索设计行前后的区域
        for row in range(
            max(0, min_design_row - 15), min(len(df), max_design_row + 20)
        ):
            # 重点搜索设计列左侧
            for col in range(0, min(len(df.columns), min_design_col + 3)):
                cell = df.iloc[row, col]
                if pd.notna(cell):
                    cell_str = str(cell).strip()

                    # 检查单元格内容
                    found_skills = []

                    # 如果包含换行符，按行检查
                    if "\n" in cell_str:
                        lines = cell_str.split("\n")
                        for line in lines:
                            line_skills = self._extract_skills_from_line(line)
                            found_skills.extend(line_skills)
                    else:
                        # 直接提取技能
                        found_skills = self._extract_skills_from_line(cell_str)

                    if found_skills:
                        # 计算优先级
                        score = len(found_skills)

                        # 在设计列左侧的优先级更高
                        if col < min_design_col:
                            score *= 2

                        # 靠近设计行的优先级更高
                        distance_to_design = min(abs(row - dr) for dr in design_rows)
                        if distance_to_design <= 5:
                            score *= 1.5

                        candidates.append(
                            {
                                "row": row,
                                "col": col,
                                "found_skills": found_skills,
                                "score": score,
                            }
                        )

        if candidates:
            # 选择最佳候选
            best = max(candidates, key=lambda x: x["score"])
            print(
                f"    选择最佳位置: [{best['row']},{best['col']}] 包含{len(best['found_skills'])}个技能"
            )
            return best

        return None

    def _extract_skills_from_column_v2(
        self, df: pd.DataFrame, start_position: Dict
    ) -> List[str]:
        """从指定位置提取技能 - V2版本"""
        skills = []
        start_row = start_position["row"]
        start_col = start_position["col"]

        # 首先提取起始单元格的所有技能
        start_cell = df.iloc[start_row, start_col]
        if pd.notna(start_cell):
            cell_str = str(start_cell).strip()

            # 如果是多行单元格（合并单元格）
            if "\n" in cell_str:
                lines = cell_str.split("\n")
                print(f"    起始单元格包含{len(lines)}行")

                for line in lines:
                    line_skills = self._extract_skills_from_line(line)
                    skills.extend(line_skills)
            else:
                # 单行技能
                cell_skills = self._extract_skills_from_line(cell_str)
                skills.extend(cell_skills)

        # 如果起始单元格已经包含足够的技能，可能不需要继续搜索
        if len(skills) >= 5:
            print(f"    起始单元格已包含{len(skills)}个技能")
            return skills

        # 否则继续向下搜索
        for row in range(start_row + 1, min(len(df), start_row + 30)):
            found_skill_in_row = False

            # 检查当前列和相邻列
            for col_offset in range(-2, 3):
                col = start_col + col_offset
                if 0 <= col < len(df.columns):
                    cell = df.iloc[row, col]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()

                        # 检查是否到达结束
                        if self._is_skill_section_end(cell_str):
                            if col == start_col:
                                return skills

                        # 提取技能
                        if "\n" in cell_str:
                            # 多行单元格
                            lines = cell_str.split("\n")
                            for line in lines:
                                line_skills = self._extract_skills_from_line(line)
                                if line_skills:
                                    skills.extend(line_skills)
                                    found_skill_in_row = True
                        else:
                            # 单行
                            cell_skills = self._extract_skills_from_line(cell_str)
                            if cell_skills:
                                skills.extend(cell_skills)
                                found_skill_in_row = True

            # 如果连续几行没有技能，停止搜索
            if not found_skill_in_row:
                # 但要继续检查几行，因为可能有空行
                empty_rows = 0
                for check_row in range(row, min(row + 5, len(df))):
                    has_content = False
                    for col in range(
                        max(0, start_col - 2), min(start_col + 3, len(df.columns))
                    ):
                        if pd.notna(df.iloc[check_row, col]):
                            has_content = True
                            break
                    if not has_content:
                        empty_rows += 1

                if empty_rows >= 3:
                    break

        return skills

    def _is_valid_skill_v2(self, skill: str) -> bool:
        """验证是否为有效技能 - V2版本，更宽松的匹配"""
        if not skill or not isinstance(skill, str):
            return False

        skill = skill.strip()

        # 基本检查
        if len(skill) < 1 or len(skill) > 50:
            return False

        # 排除明显不是技能的模式
        for pattern in self.exclude_patterns:
            if re.match(pattern, skill, re.IGNORECASE):
                return False

        # 特殊情况：单字符"C"是有效的
        if skill.upper() == "C":
            return True

        # 检查是否在有效技能列表中（大小写不敏感）
        skill_upper = skill.upper()
        for valid_skill in self.valid_skills:
            if valid_skill.upper() == skill_upper:
                return True

        # 检查是否包含技术相关的关键词
        tech_keywords = [
            "JAVA",
            "SCRIPT",
            "SQL",
            "HTML",
            "CSS",
            "SPRING",
            "BOOT",
            "REACT",
            "VUE",
            "NODE",
            "PYTHON",
            "RUBY",
            "PHP",
        ]
        for keyword in tech_keywords:
            if keyword in skill_upper:
                return True

        # 如果包含英文字符且看起来像技术名称
        if re.search(r"[a-zA-Z]", skill):
            # 排除一些明显不是技能的词
            exclude_words = ["設計", "製造", "試験", "テスト", "管理", "経験", "担当"]
            if not any(word in skill for word in exclude_words):
                # 如果长度合适且包含技术相关的模式
                if 2 <= len(skill) <= 30:
                    return True

        return False

    def _find_design_columns(self, df: pd.DataFrame) -> List[Dict]:
        """查找基本設計和詳細設計所在的列"""
        design_keywords = ["基本設計", "詳細設計", "基本设计", "详细设计"]
        positions = []

        # 遍历所有单元格查找设计关键词
        for idx in range(min(len(df), 100)):  # 限制搜索范围
            for col in range(len(df.columns)):
                cell = df.iloc[idx, col]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    if any(keyword in cell_str for keyword in design_keywords):
                        positions.append({"row": idx, "col": col, "value": cell_str})

        return positions

    def _is_skill_section_end(self, cell_content: str) -> bool:
        """判断是否到达技能区域的结束"""
        end_markers = [
            "業務内容",
            "業務內容",
            "プロジェクト",
            "職歴",
            "経歴",
            "実績",
            "担当業務",
            "開発実績",
            "主な開発",
            "参画プロジェクト",
            "作業内容",
            "期間",
            "案件",
            "概要",
            "工程",
            "【",
            "■",
            "◆",
            "▼",
            "●プロジェクト",
        ]

        cell_content = cell_content.strip()

        # 检查是否包含结束标记
        for marker in end_markers:
            if marker in cell_content:
                return True

        # 检查是否是日期格式（项目开始）
        if re.match(r"^\d{4}[年/]\d{1,2}[月/]", cell_content):
            return True

        return False

    def _extract_skills_fallback(self, df: pd.DataFrame) -> List[str]:
        """备用技能提取方法"""
        skills = []

        # 全文搜索已知技能
        text = self._dataframe_to_text(df)

        # 使用更宽松的匹配
        for skill in self.valid_skills:
            # 创建多种可能的模式
            patterns = [
                rf"\b{re.escape(skill)}\b",  # 完整单词边界
                rf"(?:^|\s|[、,，]){re.escape(skill)}(?:$|\s|[、,，])",  # 带分隔符
                rf"{re.escape(skill)}(?:[^a-zA-Z]|$)",  # 后面不是字母
            ]

            for pattern in patterns:
                if re.search(pattern, text, re.IGNORECASE):
                    skills.append(skill)
                    break

        return skills

    def _process_and_deduplicate_skills(self, skills: List[str]) -> List[str]:
        """处理和去重技能列表"""
        final_skills = []
        seen = set()
        seen_lower = set()

        for skill in skills:
            skill = skill.strip()
            if not skill:
                continue

            # 标准化
            normalized = self._normalize_skill_name(skill)
            normalized_lower = normalized.lower()

            # 去重（大小写不敏感）
            if normalized_lower not in seen_lower:
                seen_lower.add(normalized_lower)
                final_skills.append(normalized)

        # 按优先级排序
        def skill_priority(skill):
            programming_langs = [
                "Java",
                "Python",
                "JavaScript",
                "TypeScript",
                "C",
                "C++",
                "C#",
                "PHP",
                "Ruby",
                "Go",
                "Kotlin",
                "Swift",
                "Scala",
                "Rust",
                "VB.NET",
                "VB",
                "COBOL",
            ]
            frameworks = [
                "Spring",
                "SpringBoot",
                "React",
                "Vue",
                "Angular",
                "Django",
                "Flask",
                ".NET",
                "Rails",
                "Laravel",
                "Express",
            ]
            databases = [
                "MySQL",
                "PostgreSQL",
                "Oracle",
                "SQL Server",
                "MongoDB",
                "Redis",
                "DB2",
                "SQLite",
                "Access",
            ]
            cloud_platforms = ["AWS", "Azure", "GCP"]

            if skill in programming_langs:
                return (
                    0,
                    (
                        programming_langs.index(skill)
                        if skill in programming_langs
                        else 999
                    ),
                )
            elif skill in frameworks:
                return (1, frameworks.index(skill) if skill in frameworks else 999)
            elif skill in databases:
                return (2, databases.index(skill) if skill in databases else 999)
            elif skill in cloud_platforms:
                return (
                    3,
                    cloud_platforms.index(skill) if skill in cloud_platforms else 999,
                )
            else:
                return (4, skill)

        final_skills.sort(key=skill_priority)

        print(f"    最终提取技能数量: {len(final_skills)}")
        return final_skills

    # 保留所有其他方法不变...
    def _convert_excel_serial_to_date(
        self, serial_number: Union[int, float]
    ) -> Optional[datetime]:
        """将Excel序列数字转换为日期"""
        try:
            if serial_number < 1:
                return None
            if serial_number >= 61:
                serial_number -= 1
            base_date = datetime(1900, 1, 1)
            result_date = base_date + timedelta(days=serial_number - 1)
            if 1950 <= result_date.year <= 2030:
                return result_date
        except Exception as e:
            pass
        return None

    def _extract_age_final_fix(self, all_data: List[Dict]) -> str:
        """最终修复版年龄提取"""
        candidates = []

        for data in all_data:
            df = data["df"]

            # 方法1: 扫描所有Date对象
            for idx in range(min(30, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if isinstance(cell, datetime):
                        if 1950 <= cell.year <= 2010:
                            age = self._calculate_age_from_birthdate(cell)
                            if age:
                                context_score = self._get_personal_info_context_score(
                                    df, idx, col
                                )
                                confidence = 2.0 + context_score * 0.5
                                candidates.append((str(age), confidence))

            # 方法2: 扫描Excel序列日期数字
            for idx in range(min(30, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell) and isinstance(cell, (int, float)):
                        if 18000 <= cell <= 50000:
                            converted_date = self._convert_excel_serial_to_date(cell)
                            if converted_date:
                                age = self._calculate_age_from_birthdate(converted_date)
                                if age:
                                    has_age_context = False
                                    for r_off in range(-2, 3):
                                        for c_off in range(-5, 5):
                                            r = idx + r_off
                                            c = col + c_off
                                            if 0 <= r < len(df) and 0 <= c < len(
                                                df.columns
                                            ):
                                                context_cell = df.iloc[r, c]
                                                if pd.notna(context_cell):
                                                    context_str = str(context_cell)
                                                    if any(
                                                        k in context_str
                                                        for k in [
                                                            "生年月",
                                                            "年齢",
                                                            "歳",
                                                            "才",
                                                        ]
                                                    ):
                                                        has_age_context = True
                                                        break
                                        if has_age_context:
                                            break

                                    if has_age_context:
                                        confidence = 3.0
                                        candidates.append((str(age), confidence))

            # 方法3: 传统的年龄标签搜索
            for idx in range(min(30, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell):
                        cell_str = str(cell)
                        if any(k in cell_str for k in self.keywords["age"]):
                            for r_offset in range(-2, 6):
                                for c_offset in range(-2, 10):
                                    r = idx + r_offset
                                    c = col + c_offset
                                    if 0 <= r < len(df) and 0 <= c < len(df.columns):
                                        value = df.iloc[r, c]
                                        if pd.notna(value):
                                            if isinstance(value, datetime):
                                                age = (
                                                    self._calculate_age_from_birthdate(
                                                        value
                                                    )
                                                )
                                                if age:
                                                    candidates.append((str(age), 2.5))
                                            else:
                                                age = self._parse_age_value_improved(
                                                    str(value)
                                                )
                                                if age:
                                                    candidates.append((age, 2.0))

        if candidates:
            age_scores = defaultdict(float)
            for age, conf in candidates:
                age_scores[age] += conf
            best_age = max(age_scores.items(), key=lambda x: x[1])
            return best_age[0]

        return ""

    def _extract_arrival_year_final_fix(self, all_data: List[Dict]) -> Optional[str]:
        """最终修复版来日年份提取"""
        candidates = []

        for data in all_data:
            df = data["df"]

            # 方法1: 扫描所有Date对象
            for idx in range(min(30, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if isinstance(cell, datetime):
                        if 1990 <= cell.year <= 2024:
                            has_arrival_context = False
                            has_age_context = False
                            for r_off in range(-2, 3):
                                for c_off in range(-5, 5):
                                    r = idx + r_off
                                    c = col + c_off
                                    if 0 <= r < len(df) and 0 <= c < len(df.columns):
                                        context_cell = df.iloc[r, c]
                                        if pd.notna(context_cell):
                                            context_str = str(context_cell)
                                            if any(
                                                k in context_str
                                                for k in self.keywords["arrival"]
                                            ):
                                                has_arrival_context = True
                                            elif any(
                                                k in context_str
                                                for k in ["生年月", "年齢", "歳", "才"]
                                            ):
                                                has_age_context = True
                                if has_arrival_context:
                                    break

                            if has_arrival_context:
                                confidence = 1.5 if has_age_context else 2.5
                                candidates.append((str(cell.year), confidence))

            # 方法2: 扫描Excel序列日期数字
            for idx in range(min(30, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell) and isinstance(cell, (int, float)):
                        if 30000 <= cell <= 50000:
                            converted_date = self._convert_excel_serial_to_date(cell)
                            if converted_date and 1990 <= converted_date.year <= 2024:
                                has_arrival_context = False
                                for r_off in range(-2, 3):
                                    for c_off in range(-5, 5):
                                        r = idx + r_off
                                        c = col + c_off
                                        if 0 <= r < len(df) and 0 <= c < len(
                                            df.columns
                                        ):
                                            context_cell = df.iloc[r, c]
                                            if pd.notna(context_cell):
                                                context_str = str(context_cell)
                                                if any(
                                                    k in context_str
                                                    for k in self.keywords["arrival"]
                                                ):
                                                    has_arrival_context = True
                                                    break
                                        if has_arrival_context:
                                            break
                                    if has_arrival_context:
                                        break

                                if has_arrival_context:
                                    confidence = 3.0
                                    candidates.append(
                                        (str(converted_date.year), confidence)
                                    )

            # 方法3: 传统的来日关键词搜索
            for idx in range(min(40, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell):
                        cell_str = str(cell)
                        if any(k in cell_str for k in self.keywords["arrival"]):
                            for r_off in range(-2, 5):
                                for c_off in range(-2, 25):
                                    r = idx + r_off
                                    c = col + c_off
                                    if 0 <= r < len(df) and 0 <= c < len(df.columns):
                                        value = df.iloc[r, c]
                                        if pd.notna(value):
                                            if isinstance(value, datetime):
                                                if 1990 <= value.year <= 2024:
                                                    candidates.append(
                                                        (str(value.year), 2.0)
                                                    )
                                            else:
                                                year = self._parse_year_value_improved(
                                                    str(value)
                                                )
                                                if year:
                                                    candidates.append((year, 1.8))

        if candidates:
            year_scores = defaultdict(float)
            for year, conf in candidates:
                year_scores[year] += conf
            if year_scores:
                best_year = max(year_scores.items(), key=lambda x: x[1])
                return best_year[0]

        return None

    def _extract_nationality_final_fix(self, all_data: List[Dict]) -> Optional[str]:
        """最终修复版国籍提取"""
        valid_nationalities = [
            "中国",
            "日本",
            "韓国",
            "ベトナム",
            "フィリピン",
            "インド",
            "ネパール",
            "アメリカ",
            "ブラジル",
            "台湾",
            "タイ",
            "インドネシア",
            "バングラデシュ",
            "スリランカ",
            "ミャンマー",
            "カンボジア",
            "ラオス",
            "モンゴル",
        ]

        candidates = []

        for data in all_data:
            df = data["df"]

            # 方法1: 扫描整个表格的所有国籍关键词
            for idx in range(min(50, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()

                        # 直接检查是否是国籍值
                        if cell_str in valid_nationalities:
                            # 计算上下文评分
                            context_score = 0

                            # 检查同行是否有个人信息
                            row_data = df.iloc[idx]
                            row_text = " ".join(
                                [str(cell) for cell in row_data if pd.notna(cell)]
                            )
                            if any(
                                keyword in row_text
                                for keyword in [
                                    "氏名",
                                    "性別",
                                    "年齢",
                                    "最寄",
                                    "住所",
                                    "男",
                                    "女",
                                ]
                            ):
                                context_score += 2.0

                            # 检查周围是否有个人信息
                            for r_off in range(-3, 4):
                                for c_off in range(-5, 6):
                                    r = idx + r_off
                                    c = col + c_off
                                    if 0 <= r < len(df) and 0 <= c < len(df.columns):
                                        nearby_cell = df.iloc[r, c]
                                        if pd.notna(nearby_cell):
                                            nearby_text = str(nearby_cell)
                                            if any(
                                                keyword in nearby_text
                                                for keyword in [
                                                    "氏名",
                                                    "性別",
                                                    "年齢",
                                                    "学歴",
                                                ]
                                            ):
                                                context_score += 1.0

                            total_confidence = max(1.0, context_score)
                            candidates.append((cell_str, total_confidence))

            # 方法2: 查找国籍标签附近的值
            for idx in range(min(40, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell) and any(
                        k in str(cell) for k in self.keywords["nationality"]
                    ):
                        for r_off in range(-3, 6):
                            for c_off in range(-3, 15):
                                r = idx + r_off
                                c = col + c_off
                                if 0 <= r < len(df) and 0 <= c < len(df.columns):
                                    value = df.iloc[r, c]
                                    if pd.notna(value):
                                        value_str = str(value).strip()
                                        if value_str in valid_nationalities:
                                            confidence = 3.0
                                            candidates.append((value_str, confidence))

        if candidates:
            nationality_scores = defaultdict(float)
            for nationality, conf in candidates:
                nationality_scores[nationality] += conf
            if nationality_scores:
                best_nationality = max(nationality_scores.items(), key=lambda x: x[1])
                return best_nationality[0]

        return None

    def _extract_experience_final_fix(self, all_data: List[Dict]) -> str:
        """最终修复版经验提取"""
        candidates = []

        for data in all_data:
            df = data["df"]

            # 方法1: 查找经验关键词
            for idx in range(min(60, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell):
                        cell_str = str(cell)
                        if any(k in cell_str for k in self.keywords["experience"]):
                            # 排除说明文字
                            if any(
                                ex in cell_str
                                for ex in [
                                    "以上",
                                    "未満",
                                    "◎",
                                    "○",
                                    "△",
                                    "指導",
                                    "精通",
                                    "できる",
                                ]
                            ):
                                continue

                            # 搜索数值
                            for r_off in range(-3, 6):
                                for c_off in range(-3, 30):
                                    r = idx + r_off
                                    c = col + c_off
                                    if 0 <= r < len(df) and 0 <= c < len(df.columns):
                                        value = df.iloc[r, c]
                                        if pd.notna(value):
                                            exp = self._parse_experience_value(
                                                str(value)
                                            )
                                            if exp:
                                                confidence = 1.0

                                                # 根据关键词类型调整置信度
                                                if "ソフト関連業務経験年数" in cell_str:
                                                    confidence *= 3.0
                                                elif "IT経験年数" in cell_str:
                                                    confidence *= 2.5
                                                elif "実務経験" in cell_str:
                                                    confidence *= 2.0

                                                candidates.append((exp, confidence))

            # 方法2: 从项目日期推算经验
            for idx in range(min(50, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if isinstance(cell, datetime):
                        # 检查是否是项目开始日期（可以推算经验）
                        if 2015 <= cell.year <= 2024:
                            # 检查同行是否有项目描述
                            row_data = df.iloc[idx]
                            row_text = " ".join(
                                [str(cell) for cell in row_data if pd.notna(cell)]
                            )
                            if any(
                                keyword in row_text
                                for keyword in [
                                    "システム",
                                    "開発",
                                    "業務",
                                    "プロジェクト",
                                ]
                            ):
                                # 从最早的项目日期推算经验年数
                                experience_years = 2024 - cell.year
                                if 1 <= experience_years <= 15:
                                    # 对于合理的项目经验，给予更高置信度
                                    confidence = 1.5 if experience_years >= 5 else 1.2
                                    exp_str = f"{experience_years}年"
                                    candidates.append((exp_str, confidence))

        if candidates:
            candidates.sort(key=lambda x: x[1], reverse=True)
            result = candidates[0][0].translate(self.trans_table)
            return result

        return ""

    def _calculate_age_from_birthdate(self, birthdate: datetime) -> Optional[int]:
        """从生年月日计算年龄"""
        try:
            current_date = datetime(2024, 11, 1)
            age = current_date.year - birthdate.year
            if (current_date.month, current_date.day) < (
                birthdate.month,
                birthdate.day,
            ):
                age -= 1
            if 15 <= age <= 75:
                return age
        except:
            pass
        return None

    def _parse_age_value_improved(self, value: str) -> Optional[str]:
        """改进的年龄值解析"""
        value = str(value).strip()
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

    def _parse_year_value_improved(self, value: str) -> Optional[str]:
        """改进的年份值解析"""
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

    def _get_personal_info_context_score(
        self, df: pd.DataFrame, row: int, col: int
    ) -> float:
        """获取个人信息上下文评分"""
        score = 0.0
        personal_keywords = (
            self.keywords["name"]
            + self.keywords["gender"]
            + self.keywords["age"]
            + ["学歴", "最終学歴", "住所", "最寄駅", "最寄", "経験", "実務"]
        )

        for r in range(max(0, row - 3), min(len(df), row + 4)):
            for c in range(max(0, col - 5), min(len(df.columns), col + 6)):
                cell = df.iloc[r, c]
                if pd.notna(cell):
                    cell_str = str(cell)
                    for keyword in personal_keywords:
                        if keyword in cell_str:
                            score += 1.0

        return score

    # 保留原有的其他方法
    def _extract_name(self, all_data: List[Dict]) -> str:
        """提取姓名"""
        candidates = []

        for data in all_data:
            df = data["df"]

            for idx in range(min(20, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell) and any(
                        k in str(cell) for k in self.keywords["name"]
                    ):
                        for r_offset in range(-2, 5):
                            for c_offset in range(-2, 15):
                                r = idx + r_offset
                                c = col + c_offset
                                if 0 <= r < len(df) and 0 <= c < len(df.columns):
                                    value = df.iloc[r, c]
                                    if pd.notna(value) and self._is_valid_name(
                                        str(value)
                                    ):
                                        confidence = 1.0
                                        if r == idx:
                                            confidence *= 1.5
                                        if c > col and c - col <= 5:
                                            confidence *= 1.3
                                        candidates.append(
                                            (str(value).strip(), confidence)
                                        )

        if candidates:
            candidates.sort(key=lambda x: x[1], reverse=True)
            return candidates[0][0]
        return ""

    def _extract_gender(self, all_data: List[Dict]) -> Optional[str]:
        """提取性别"""
        for data in all_data:
            df = data["df"]

            for idx in range(min(30, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()

                        if cell_str in ["男", "男性"]:
                            if self._has_nearby_keyword(
                                df, idx, col, self.keywords["gender"], radius=5
                            ):
                                return "男性"
                        elif cell_str in ["女", "女性"]:
                            if self._has_nearby_keyword(
                                df, idx, col, self.keywords["gender"], radius=5
                            ):
                                return "女性"
                        elif any(k in cell_str for k in self.keywords["gender"]):
                            for r_off in range(-2, 3):
                                for c_off in range(-2, 10):
                                    r = idx + r_off
                                    c = col + c_off
                                    if 0 <= r < len(df) and 0 <= c < len(df.columns):
                                        value = df.iloc[r, c]
                                        if pd.notna(value):
                                            v_str = str(value).strip()
                                            if v_str in ["男", "男性", "M", "Male"]:
                                                return "男性"
                                            elif v_str in ["女", "女性", "F", "Female"]:
                                                return "女性"
        return None

    def _extract_japanese_level(self, all_data: List[Dict]) -> str:
        """提取日语水平"""
        candidates = []

        for data in all_data:
            text = data["text"]
            df = data["df"]

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
                    kanji_to_num = {
                        "一": "1",
                        "二": "2",
                        "三": "3",
                        "四": "4",
                        "五": "5",
                    }
                    if level_str in kanji_to_num:
                        level_num = kanji_to_num[level_str]
                    else:
                        level_num = level_str.translate(self.trans_table)
                    level = f"N{level_num}"
                    candidates.append((level, confidence))

            if "ビジネス" in text and any(jp in text for jp in ["日本語", "日語"]):
                candidates.append(("ビジネスレベル", 0.8))
            elif "上級" in text and any(jp in text for jp in ["日本語", "日語"]):
                candidates.append(("上級", 0.8))
            elif "中級" in text and any(jp in text for jp in ["日本語", "日語"]):
                candidates.append(("中級", 0.7))

        if candidates:
            candidates.sort(key=lambda x: x[1], reverse=True)
            return candidates[0][0]
        return ""

    def _normalize_skill_name(self, skill: str) -> str:
        skill = skill.strip()
        skill_mapping = {
            "JAVA": "Java",
            "java": "Java",
            "Javascript": "JavaScript",
            "javascript": "JavaScript",
            "JAVASCRIPT": "JavaScript",
            "MySql": "MySQL",
            "mysql": "MySQL",
            "Intellij": "IntelliJ",
            "intellij": "IntelliJ",
            "Junit": "JUnit",
            "junit": "JUnit",
            "python": "Python",
            "React.js": "React",
            "Vue.js": "Vue",
            "Node.JS": "Node.js",
            "node.js": "Node.js",
            "LINUX": "Linux",
            "linux": "Linux",
            "UNIX": "Unix",
            "unix": "Unix",
            "ORACLE": "Oracle",
            "oracle": "Oracle",
            "eclipse": "Eclipse",
            "spring": "Spring",
            "springboot": "SpringBoot",
            "SpringBoot": "SpringBoot",
            "postman": "Postman",
            "git": "Git",
            "github": "GitHub",
            "struct2": "Struts2",
            "PG": "PostgreSQL",
            "JS": "JavaScript",
            "ＯＳ": "OS",
            "ＤＢ": "DB",
            "Jquery": "jQuery",
            "jquery": "jQuery",
            "JBOSS": "JBoss",
            "jboss": "JBoss",
            "SqlServer": "SQL Server",
            "sqlserver": "SQL Server",
            "Form Desigener": "Form Designer",
            "win10": "Windows 10",
            "Win10": "Windows 10",
        }
        if skill in skill_mapping:
            return skill_mapping[skill]
        skill_lower = skill.lower()
        for k, v in skill_mapping.items():
            if k.lower() == skill_lower:
                return v
        skill = re.sub(r"\s+", " ", skill)
        if "ＯＳ" in skill:
            skill = skill.replace("ＯＳ", "OS")
        if "ＤＢ" in skill:
            skill = skill.replace("ＤＢ", "DB")
        if skill in self.valid_skills:
            return skill
        for valid_skill in self.valid_skills:
            if valid_skill.lower() == skill_lower:
                return valid_skill
        return skill

    def _is_skill_mark(self, value) -> bool:
        if pd.isna(value):
            return False
        value_str = str(value).strip()
        return value_str in self.skill_marks

    def _is_valid_name(self, name: str) -> bool:
        name = str(name).strip()
        exclude_words = [
            "氏名",
            "名前",
            "フリガナ",
            "性別",
            "年齢",
            "国籍",
            "男",
            "女",
            "歳",
            "才",
            "経験",
            "資格",
            "学歴",
            "住所",
            "電話",
            "メール",
            "現在",
            "スキルシート",
            "履歴書",
            "職務経歴書",
            "技術",
            "年月",
            "生年月",
        ]
        for word in exclude_words:
            if word in name:
                return False
        if len(name) < 2 or len(name) > 15:
            return False
        if name.replace(" ", "").replace("　", "").isdigit():
            return False
        if not re.search(r"[一-龥ぁ-んァ-ンa-zA-Z]", name):
            return False
        if re.match(r"^[A-Z]{2,4}$", name):
            return True
        if " " in name or "　" in name:
            parts = re.split(r"[\s　]+", name)
            if len(parts) == 2 and all(len(p) >= 1 for p in parts):
                return True
        return True

    def _parse_experience_value(self, value: str) -> Optional[str]:
        value = str(value).strip()
        if any(exclude in value for exclude in ["以上", "未満", "◎", "○", "△", "経験"]):
            return None
        value = value.translate(self.trans_table)
        patterns = [
            (
                r"^(\d+)\s*年\s*(\d+)\s*ヶ月$",
                lambda m: f"{m.group(1)}年{m.group(2)}ヶ月",
            ),
            (
                r"^(\d+(?:\.\d+)?)\s*年$",
                lambda m: (
                    f"{float(m.group(1)):.0f}年"
                    if float(m.group(1)) == int(float(m.group(1)))
                    else f"{m.group(1)}年"
                ),
            ),
            (
                r"^(\d+(?:\.\d+)?)\s*$",
                lambda m: (
                    f"{float(m.group(1)):.0f}年"
                    if 1 <= float(m.group(1)) <= 40
                    else None
                ),
            ),
            (r"(\d+)\s*年\s*(\d+)\s*ヶ月", lambda m: f"{m.group(1)}年{m.group(2)}ヶ月"),
            (r"(\d+)\s*年", lambda m: f"{m.group(1)}年"),
        ]
        for pattern, formatter in patterns:
            match = re.search(pattern, value)
            if match:
                result = formatter(match)
                if result:
                    return result
        return None

    def _has_nearby_keyword(
        self, df: pd.DataFrame, row: int, col: int, keywords: List[str], radius: int = 5
    ) -> bool:
        for r in range(max(0, row - radius), min(len(df), row + radius + 1)):
            for c in range(
                max(0, col - radius), min(len(df.columns), col + radius + 1)
            ):
                cell = df.iloc[r, c]
                if pd.notna(cell) and any(k in str(cell) for k in keywords):
                    return True
        return False

    def _dataframe_to_text(self, df: pd.DataFrame) -> str:
        text_parts = []
        for idx, row in df.iterrows():
            row_text = " ".join([str(cell) for cell in row if pd.notna(cell)])
            if row_text.strip():
                text_parts.append(row_text)
        return "\n".join(text_parts)


# 测试函数
if __name__ == "__main__":
    extractor = FinalFixResumeExtractor()

    test_files = [
        "スキルシート_ryu.xlsx",  # 重点测试这个文件的技能提取
        "業務経歴書_ FYY_金町駅.xlsx",
        "スキルシートGT.xlsx",
    ]

    for file in test_files:
        print(f"\n{'='*60}")
        print(f"🔧 最终修复测试文件V2: {file}")
        print("=" * 60)

        try:
            result = extractor.extract_from_excel(file)

            if "error" not in result:
                print(f"\n✅ 提取结果:")
                print(f"   姓名: {result['name']}")
                print(f"   性别: {result['gender']}")
                print(f"   年龄: {result['age']}")
                print(f"   国籍: {result['nationality']}")
                print(f"   来日年份: {result['arrival_year_japan']}")
                print(f"   经验: {result['experience']}")
                print(f"   日语: {result['japanese_level']}")
                print(f"   技能({len(result['skills'])}个): ")

                # 分组显示技能
                if result["skills"]:
                    # 前10个技能
                    print(f"     主要技能: {', '.join(result['skills'][:10])}")
                    if len(result["skills"]) > 10:
                        print(f"     其他技能: {', '.join(result['skills'][10:20])}")
                        if len(result["skills"]) > 20:
                            print(f"     ... 还有 {len(result['skills']) - 20} 个技能")

                print("\n📋 JSON格式:")
                print(json.dumps(result, ensure_ascii=False, indent=2))
            else:
                print(f"❌ 错误: {result['error']}")
        except Exception as e:
            print(f"❌ 错误: {e}")
            import traceback

            traceback.print_exc()
