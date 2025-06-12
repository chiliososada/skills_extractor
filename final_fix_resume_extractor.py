# final_fix_resume_extractor.py - 最终修复版简历提取器
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
import re
import json
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Union, Tuple, Set
from collections import defaultdict, Counter


class FinalFixResumeExtractor:
    """最终修复版简历提取器 - 解决Excel序列日期问题"""

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
            # 语言（特殊处理）
            "日本語",
            "英語",
            "中国語",
            # Web相关
            "PHP/HTML",
            "Jquery",
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
            result["skills"] = self._extract_skills(all_data)

            return result

        except Exception as e:
            print(f"处理文件时出错: {e}")
            import traceback

            traceback.print_exc()
            return {"error": str(e)}

    def _convert_excel_serial_to_date(
        self, serial_number: Union[int, float]
    ) -> Optional[datetime]:
        """将Excel序列数字转换为日期"""
        try:
            # Excel的序列日期：从1900年1月1日开始的天数
            # 注意：Excel错误地认为1900年是闰年，所以需要调整
            if serial_number < 1:
                return None

            # Excel的bug：1900年被错误地认为是闰年，所以1900年3月1日之后需要减去1天
            if serial_number >= 61:  # 1900年3月1日对应61
                serial_number -= 1

            base_date = datetime(1900, 1, 1)
            result_date = base_date + timedelta(days=serial_number - 1)

            # 验证日期是否合理
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

            print(f"🔍 开始最终修复版年龄提取...")

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
                                print(
                                    f"    Date对象年龄候选: [{idx},{col}] {cell} -> 年龄 {age}"
                                )

            # 方法2: **关键修复** - 扫描Excel序列日期数字
            for idx in range(min(30, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell) and isinstance(cell, (int, float)):
                        # 检查是否可能是Excel序列日期（范围大概在18000-50000之间）
                        if 18000 <= cell <= 50000:
                            converted_date = self._convert_excel_serial_to_date(cell)
                            if converted_date:
                                age = self._calculate_age_from_birthdate(converted_date)
                                if age:
                                    # 检查周围是否有生年月日相关标签
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
                                        confidence = 3.0  # 序列日期+上下文的高置信度
                                        candidates.append((str(age), confidence))
                                        print(
                                            f"    Excel序列日期年龄候选: [{idx},{col}] {cell} -> {converted_date} -> 年龄 {age}"
                                        )

            # 方法3: 传统的年龄标签搜索
            for idx in range(min(30, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell):
                        cell_str = str(cell)
                        if any(k in cell_str for k in self.keywords["age"]):
                            # 搜索附近的年龄值
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
            print(f"    年龄候选总数: {len(candidates)}")
            for age, conf in candidates:
                print(f"      候选: 年龄={age}, 置信度={conf}")

            age_scores = defaultdict(float)
            for age, conf in candidates:
                age_scores[age] += conf
            best_age = max(age_scores.items(), key=lambda x: x[1])
            print(f"    最终选择: 年龄={best_age[0]}, 总置信度={best_age[1]}")
            return best_age[0]

        print(f"    未找到年龄信息")
        return ""

    def _extract_arrival_year_final_fix(self, all_data: List[Dict]) -> Optional[str]:
        """最终修复版来日年份提取"""
        candidates = []

        for data in all_data:
            df = data["df"]

            print(f"🔍 开始最终修复版来日年份提取...")

            # 方法1: 扫描所有Date对象
            for idx in range(min(30, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if isinstance(cell, datetime):
                        if 1990 <= cell.year <= 2024:
                            # 检查是否有来日上下文
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
                                # 如果也有年龄上下文，可能是生年月日，降低置信度
                                confidence = 1.5 if has_age_context else 2.5
                                candidates.append((str(cell.year), confidence))
                                print(
                                    f"    Date对象来日候选: [{idx},{col}] {cell} -> 年份 {cell.year} (置信度: {confidence})"
                                )

            # 方法2: **关键修复** - 扫描Excel序列日期数字
            for idx in range(min(30, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell) and isinstance(cell, (int, float)):
                        # 检查是否可能是Excel序列日期
                        if 30000 <= cell <= 50000:  # 大概1982-2037年的范围
                            converted_date = self._convert_excel_serial_to_date(cell)
                            if converted_date and 1990 <= converted_date.year <= 2024:
                                # 检查周围是否有来日相关标签
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
                                    confidence = 3.0  # 序列日期+上下文的高置信度
                                    candidates.append(
                                        (str(converted_date.year), confidence)
                                    )
                                    print(
                                        f"    Excel序列日期来日候选: [{idx},{col}] {cell} -> {converted_date} -> 年份 {converted_date.year}"
                                    )

            # 方法3: 传统的来日关键词搜索
            for idx in range(min(40, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell):
                        cell_str = str(cell)
                        if any(k in cell_str for k in self.keywords["arrival"]):
                            # 搜索附近的年份
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
            print(f"    来日年份候选总数: {len(candidates)}")
            for year, conf in candidates:
                print(f"      候选: 年份={year}, 置信度={conf}")

            year_scores = defaultdict(float)
            for year, conf in candidates:
                year_scores[year] += conf
            if year_scores:
                best_year = max(year_scores.items(), key=lambda x: x[1])
                print(f"    最终选择: 年份={best_year[0]}, 总置信度={best_year[1]}")
                return best_year[0]

        print(f"    未找到来日年份信息")
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

            print(f"🔍 开始最终修复版国籍提取...")

            # 方法1: 扫描整个表格的所有国籍关键词
            for idx in range(min(50, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()

                        # 直接检查是否是国籍值
                        if cell_str in valid_nationalities:
                            print(f"    发现国籍候选: [{idx},{col}] {cell_str}")

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
                                print(f"      同行有个人信息")

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
                        print(f"    发现国籍标签: [{idx},{col}] {cell}")
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
                                            print(f"      找到国籍值: {value_str}")
                                            candidates.append((value_str, confidence))

        if candidates:
            print(f"    国籍候选总数: {len(candidates)}")
            for nationality, conf in candidates:
                print(f"      候选: 国籍={nationality}, 置信度={conf}")

            nationality_scores = defaultdict(float)
            for nationality, conf in candidates:
                nationality_scores[nationality] += conf
            if nationality_scores:
                best_nationality = max(nationality_scores.items(), key=lambda x: x[1])
                print(
                    f"    最终选择: 国籍={best_nationality[0]}, 总置信度={best_nationality[1]}"
                )
                return best_nationality[0]

        print(f"    未找到国籍信息")
        return None

    def _extract_experience_final_fix(self, all_data: List[Dict]) -> str:
        """最终修复版经验提取"""
        candidates = []

        for data in all_data:
            df = data["df"]

            print(f"🔍 开始最终修复版经验提取...")

            # 方法1: 查找经验关键词
            for idx in range(min(60, len(df))):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if pd.notna(cell):
                        cell_str = str(cell)
                        if any(k in cell_str for k in self.keywords["experience"]):
                            print(f"    发现经验关键词: [{idx},{col}] {cell_str}")

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
                                print(f"      跳过说明文字")
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
                                                print(
                                                    f"      找到经验值: [{r},{c}] {value} -> {exp}"
                                                )

                                                # 根据关键词类型调整置信度
                                                if "ソフト関連業務経験年数" in cell_str:
                                                    confidence *= 3.0
                                                elif "IT経験年数" in cell_str:
                                                    confidence *= 2.5
                                                elif "実務経験" in cell_str:
                                                    confidence *= 2.0

                                                candidates.append((exp, confidence))

            # 方法2: **GT文件特殊处理** - 検查実務年数表格
            # 对于GT文件，実務年数可能没有具体数字，而是用符号表示
            print(f"    检查実務年数表格格式...")

            # 在GT文件中，経験可能以其他形式存在，比如项目开始时间推算
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
                                    print(
                                        f"      从项目日期推算经验: {cell} -> {exp_str} (置信度: {confidence})"
                                    )

        if candidates:
            print(f"    经验候选总数: {len(candidates)}")
            for exp, conf in candidates:
                print(f"      候选: 经验={exp}, 置信度={conf}")

            candidates.sort(key=lambda x: x[1], reverse=True)
            result = candidates[0][0].translate(self.trans_table)
            print(f"    最终选择: 经验={result}, 置信度={candidates[0][1]}")
            return result

        print(f"    未找到经验信息")
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

    def _extract_skills(self, all_data: List[Dict]) -> List[str]:
        """提取技能"""
        all_skills = []

        for data in all_data:
            df = data["df"]

            for idx in range(len(df)):
                for col in range(len(df.columns)):
                    cell = df.iloc[idx, col]
                    if self._is_skill_mark(cell):
                        search_positions = [
                            (-1, 0),
                            (0, -1),
                            (-1, -1),
                            (-2, 0),
                            (0, -2),
                            (0, -3),
                            (0, -4),
                            (0, -5),
                        ]

                        for r_off, c_off in search_positions:
                            r = idx + r_off
                            c = col + c_off
                            if 0 <= r < len(df) and 0 <= c < len(df.columns):
                                skill_cell = df.iloc[r, c]
                                if pd.notna(skill_cell):
                                    skills = self._extract_skills_from_cell(
                                        str(skill_cell)
                                    )
                                    if skills:
                                        all_skills.extend(skills)
                                        break

            text = data["text"]
            for skill in self.valid_skills:
                if re.search(rf"\b{re.escape(skill)}\b", text, re.IGNORECASE):
                    all_skills.append(skill)

        final_skills = []
        seen = set()

        for skill in all_skills:
            skill = skill.strip()
            if skill and skill not in seen and self._is_valid_skill(skill):
                normalized = self._normalize_skill_name(skill)
                if normalized and normalized not in seen:
                    seen.add(normalized)
                    final_skills.append(normalized)

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
            ]
            databases = ["MySQL", "PostgreSQL", "Oracle", "MongoDB", "Redis", "DB2"]

            if skill in programming_langs:
                return (0, programming_langs.index(skill))
            elif skill in frameworks:
                return (1, frameworks.index(skill))
            elif skill in databases:
                return (2, databases.index(skill))
            else:
                return (3, skill)

        final_skills.sort(key=skill_priority)
        return final_skills

    # 保留所有辅助方法
    def _extract_skills_from_cell(self, cell_content: str) -> List[str]:
        if not cell_content:
            return []
        category_labels = [
            "ＯＳ",
            "OS",
            "ＤＢ",
            "DB",
            "言語",
            "開発言語",
            "フレームワーク",
            "ツール",
            "技術経験",
            "その他",
        ]
        if cell_content.strip() in category_labels:
            return []
        skills = re.split(r"[\n\r,，、\t]+", cell_content)
        valid_skills = []
        for skill in skills:
            skill = skill.strip()
            if skill and self._is_valid_skill(skill):
                valid_skills.append(skill)
        return valid_skills

    def _is_skill_mark(self, value) -> bool:
        if pd.isna(value):
            return False
        value_str = str(value).strip()
        return value_str in self.skill_marks

    def _is_valid_skill(self, skill: str) -> bool:
        if not skill or not isinstance(skill, str):
            return False
        skill = skill.strip()
        if len(skill) < 1 or len(skill) > 50:
            return False
        for pattern in self.exclude_patterns:
            if re.match(pattern, skill, re.IGNORECASE):
                return False
        if skill.isdigit():
            return False
        if re.search(r"\d{4}[-/]\d{2}[-/]\d{2}", skill):
            return False
        if skill == "C":
            return True
        if len(skill) == 1 and skill not in self.skill_marks:
            return False
        normalized = self._normalize_skill_name(skill)
        if normalized in self.valid_skills:
            return True
        skill_lower = skill.lower()
        for valid_skill in self.valid_skills:
            if valid_skill.lower() == skill_lower:
                return True
        if re.search(r"[a-zA-Z]", skill) and len(skill) <= 30:
            exclude_words = [
                "設計",
                "製造",
                "試験",
                "テスト",
                "管理",
                "経験",
                "担当",
                "年",
                "月",
            ]
            if not any(word in skill for word in exclude_words):
                return True
        os_names = ["Windows", "Linux", "Unix", "Solaris", "Ubuntu", "CentOS", "RedHat"]
        if any(os.lower() in skill_lower for os in os_names):
            return True
        return False

    def _normalize_skill_name(self, skill: str) -> str:
        skill = skill.strip()
        skill_mapping = {
            "JAVA": "Java",
            "java": "Java",
            "Javascript": "JavaScript",
            "javascript": "JavaScript",
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
        "スキルシート_ryu.xlsx",  # 测试Excel序列日期来日年份
        "業務経歴書_ FYY_金町駅.xlsx",  # 测试Excel序列日期年龄
        "スキルシートGT.xlsx",  # 测试国籍和经验
    ]

    for file in test_files:
        print(f"\n{'='*60}")
        print(f"🔧 最终修复测试文件: {file}")
        print("=" * 60)

        try:
            result = extractor.extract_from_excel(file)

            if "error" not in result:
                print(f"✅ 姓名: {result['name']}")
                print(f"✅ 性别: {result['gender']}")
                print(f"✅ 年龄: {result['age']}")
                print(f"✅ 国籍: {result['nationality']}")
                print(f"✅ 来日年份: {result['arrival_year_japan']}")
                print(f"✅ 经验: {result['experience']}")
                print(f"✅ 日语: {result['japanese_level']}")
                print(
                    f"✅ 技能({len(result['skills'])}个): {', '.join(result['skills'][:5])}"
                )

                print("\n📋 JSON格式:")
                print(json.dumps(result, ensure_ascii=False, indent=2))
            else:
                print(f"❌ 错误: {result['error']}")
        except Exception as e:
            print(f"❌ 错误: {e}")
            import traceback

            traceback.print_exc()
