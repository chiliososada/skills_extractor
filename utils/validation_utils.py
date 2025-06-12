# -*- coding: utf-8 -*-
"""验证工具"""

import re
from typing import List


# -*- coding: utf-8 -*-
"""验证工具 - 最终修复版（针对スキルシート_付.xlsx问题）"""

import re
from typing import List


def is_valid_name(name: str) -> bool:
    """验证是否为有效的姓名 - 最终修复版

    Args:
        name: 待验证的姓名

    Returns:
        是否为有效姓名
    """
    name = str(name).strip()

    # 需要排除的词 - 包含学校相关词汇
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
        # 新増学校相关词汇
        "大学",
        "学校",
        "研究科",
        "学院",
        "専門学校",
        "高校",
        "中学校",
        "小学校",
        "大学院",
        "学部",
        "研究室",
        "工学部",
        "理学部",
        "文学部",
        "法学部",
        "経済学部",
        "医学部",
        "薬学部",
        "農学部",
        "教育学部",
        "商学部",
        "博士",
        "修士",
        "学士",
        "卒業",
        "在学",
        "専攻",
        "学科",
        # 学位相关
        "PhD",
        "Master",
        "Bachelor",
        "MBA",
        "修了",
        "取得",
    ]

    # 检查是否包含排除词
    for word in exclude_words:
        if word in name:
            return False

    # 检查是否是学校名称的模式
    school_patterns = [
        r".*大学.*",  # 包含"大学"
        r".*学院.*",  # 包含"学院"
        r".*研究科.*",  # 包含"研究科"
        r".*学校.*",  # 包含"学校"
        r".*専門.*",  # 包含"専門"
        r".*高等.*",  # 包含"高等"
        r".*University.*",  # 英文大学
        r".*College.*",  # 英文学院
        r".*Institute.*",  # 英文研究所
        r".*School.*",  # 英文学校
    ]

    for pattern in school_patterns:
        if re.match(pattern, name, re.IGNORECASE):
            return False

    # 排除假名标记（单个假名字符，通常是标记用）
    kana_markers = [
        "ア",
        "イ",
        "ウ",
        "エ",
        "オ",
        "カ",
        "キ",
        "ク",
        "ケ",
        "コ",
        "サ",
        "シ",
        "ス",
        "セ",
        "ソ",
        "タ",
        "チ",
        "ツ",
        "テ",
        "ト",
        "ナ",
        "ニ",
        "ヌ",
        "ネ",
        "ノ",
        "ハ",
        "ヒ",
        "フ",
        "ヘ",
        "ホ",
        "マ",
        "ミ",
        "ム",
        "メ",
        "モ",
        "ヤ",
        "ユ",
        "ヨ",
        "ラ",
        "リ",
        "ル",
        "レ",
        "ロ",
        "ワ",
        "ヲ",
        "ン",
    ]

    if name in kana_markers:
        return False

    # 长度检查
    if len(name) < 1 or len(name) > 15:
        return False

    # 不能全是数字
    if name.replace(" ", "").replace("　", "").isdigit():
        return False

    # 必须包含文字字符
    if not re.search(r"[一-龥ぁ-んァ-ンa-zA-Z]", name):
        return False

    # 特殊情况：缩写名
    if re.match(r"^[A-Z]{2,4}$", name):
        return True

    # 空格分隔的姓名
    if " " in name or "　" in name:
        parts = re.split(r"[\s　]+", name)
        if len(parts) == 2 and all(len(p) >= 1 for p in parts):
            return True

    # 排除过长的组织名称（包含特殊符号）
    if len(name) > 10 and any(char in name for char in ["・", "・", "-", "ー"]):
        return False

    return True


def is_valid_skill(skill: str, valid_skills: set, exclude_patterns: List[str]) -> bool:
    """验证是否为有效的技能

    Args:
        skill: 待验证的技能
        valid_skills: 有效技能集合
        exclude_patterns: 排除模式列表

    Returns:
        是否为有效技能
    """
    if not skill or not isinstance(skill, str):
        return False

    skill = skill.strip()

    # 长度检查
    if len(skill) < 1 or len(skill) > 50:
        return False

    # 检查排除模式
    for pattern in exclude_patterns:
        if re.match(pattern, skill, re.IGNORECASE):
            return False

    # 特殊情况
    if skill == "C":
        return True

    # 单字符且不是技能标记
    if len(skill) == 1:
        return False

    # 检查是否在有效技能列表中
    if skill in valid_skills or skill.lower() in {s.lower() for s in valid_skills}:
        return True

    # 包含英文字符的可能是技能
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

    return False
