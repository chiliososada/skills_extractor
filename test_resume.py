# -*- coding: utf-8 -*-
"""
简历提取测试脚本
用于测试 EnhancedResumeExtractor 的功能
"""

import os
import json
from datetime import datetime
from resume_extractor import EnhancedResumeExtractor, batch_process_resumes


def test_single_file(file_path):
    """测试单个文件的提取"""
    print(f"\n{'='*60}")
    print(f"测试文件: {file_path}")
    print("=" * 60)

    if not os.path.exists(file_path):
        print(f"❌ 错误：文件不存在 - {file_path}")
        return None

    extractor = EnhancedResumeExtractor()

    try:
        # 提取信息
        result = extractor.extract_from_excel(file_path)

        # 显示结果
        if "error" not in result:
            print("\n✅ 提取成功！")
            print(f"姓名: {result.get('name', '未找到')}")
            print(f"性别: {result.get('gender', '未找到')}")
            print(f"年龄: {result.get('age', '未找到')}")
            print(f"国籍: {result.get('nationality', '未找到')}")
            print(f"来日年份: {result.get('arrival_year_japan', '未找到')}")
            print(f"工作经验: {result.get('experience', '未找到')}")
            print(f"日语水平: {result.get('japanese_level', '未找到')}")

            skills = result.get("skills", [])
            print(f"\n技能 ({len(skills)}个):")
            if skills:
                # 分组显示技能
                for i in range(0, len(skills), 5):
                    print(f"  {', '.join(skills[i:i+5])}")
            else:
                print("  未找到技能")

            return result
        else:
            print(f"❌ 提取失败: {result['error']}")
            return None

    except Exception as e:
        print(f"❌ 处理出错: {str(e)}")
        import traceback

        traceback.print_exc()
        return None


def test_all_files():
    """测试所有简历文件"""
    print("\n" + "=" * 80)
    print("批量测试所有简历文件")
    print("=" * 80)

    # 定义测试文件列表
    test_files = [
        "蕨_馬S.xlsx",
        "【スキルシート】_庄YW_202411.xlsx",
        "スキルシート_ryu.xlsx",
        "本八幡_スキルシート_JY.xlsx",
        "業務経歴書_ FYY_金町駅.xlsx",
        "スキルシートGT.xlsx",
    ]

    # 检查哪些文件存在
    existing_files = []
    missing_files = []

    for file in test_files:
        if os.path.exists(file):
            existing_files.append(file)
        else:
            missing_files.append(file)

    if missing_files:
        print("\n⚠️  缺少以下文件:")
        for file in missing_files:
            print(f"  - {file}")

    if not existing_files:
        print("\n❌ 没有找到任何测试文件！")
        return

    print(f"\n找到 {len(existing_files)} 个文件，开始处理...\n")

    # 批量处理
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"简历提取结果_{timestamp}.json"

    results = batch_process_resumes(existing_files, output_file)

    # 显示统计信息
    print("\n" + "=" * 60)
    print("处理统计")
    print("=" * 60)

    total = len(results)
    success = sum(1 for r in results if "error" not in r)
    failed = total - success

    print(f"总文件数: {total}")
    print(f"成功提取: {success}")
    print(f"失败数量: {failed}")
    print(f"成功率: {success/total*100:.1f}%")

    # 显示失败的文件
    if failed > 0:
        print("\n失败的文件:")
        for r in results:
            if "error" in r:
                print(f"  - {r['source_file']}: {r['error']}")

    print(f"\n📁 结果已保存到: {output_file}")

    return results


def interactive_test():
    """交互式测试"""
    print("\n" + "=" * 80)
    print("简历提取系统 - 交互式测试")
    print("=" * 80)

    while True:
        print("\n请选择操作:")
        print("1. 测试单个文件")
        print("2. 批量测试所有文件")
        print("3. 查看使用说明")
        print("0. 退出")

        choice = input("\n请输入选项 (0-3): ").strip()

        if choice == "0":
            print("\n感谢使用！再见！")
            break

        elif choice == "1":
            file_path = input("\n请输入文件路径 (或直接回车使用默认文件): ").strip()
            if not file_path:
                # 使用默认文件
                default_files = [
                    "蕨_馬S.xlsx",
                    "【スキルシート】_庄YW_202411.xlsx",
                    "スキルシート_ryu.xlsx",
                ]
                for f in default_files:
                    if os.path.exists(f):
                        file_path = f
                        print(f"使用默认文件: {file_path}")
                        break
                else:
                    print("❌ 没有找到默认文件")
                    continue

            test_single_file(file_path)

        elif choice == "2":
            test_all_files()

        elif choice == "3":
            print("\n" + "=" * 60)
            print("使用说明")
            print("=" * 60)
            print(
                """
1. 支持的文件格式：Excel (.xlsx)

2. 可提取的信息：
   - 姓名 (name)
   - 性别 (gender)
   - 年龄 (age)
   - 国籍 (nationality)
   - 来日年份 (arrival_year_japan)
   - 工作经验 (experience)
   - 日语水平 (japanese_level)
   - 技能列表 (skills)

3. 支持的日语等级格式：
   - JLPT: N1-N5, 1級-5級
   - 标记: ◎○△×, ●★, ABCD
   - 描述: ビジネスレベル, 日常会話等

4. 工作经验计算：
   - 自动排除教育期间
   - 自动排除来日准备期间
   - 支持多种格式 (X年, X年Yヶ月, X.Y年)

5. 输出格式：JSON文件
            """
            )
        else:
            print("❌ 无效的选项，请重新选择")


def quick_test():
    """快速测试所有功能"""
    print("\n🚀 开始快速测试...\n")

    # 测试第一个存在的文件
    test_files = [
        "蕨_馬S.xlsx",
        "【スキルシート】_庄YW_202411.xlsx",
        "スキルシート_ryu.xlsx",
        "本八幡_スキルシート_JY.xlsx",
        "業務経歴書_ FYY_金町駅.xlsx",
        "スキルシートGT.xlsx",
    ]

    found = False
    for file in test_files:
        if os.path.exists(file):
            print(f"找到测试文件: {file}")
            result = test_single_file(file)
            if result:
                print("\n📋 JSON格式输出:")
                print(json.dumps(result, ensure_ascii=False, indent=2))
            found = True
            break

    if not found:
        print("❌ 没有找到任何测试文件！")
        print("\n请确保以下文件至少有一个在当前目录:")
        for file in test_files:
            print(f"  - {file}")


if __name__ == "__main__":
    # 显示欢迎信息
    print("=" * 80)
    print("日文简历信息提取系统 - 测试程序")
    print("=" * 80)

    # 检查是否有命令行参数
    import sys

    if len(sys.argv) > 1:
        # 如果有参数，测试指定的文件
        for file_path in sys.argv[1:]:
            test_single_file(file_path)
    else:
        # 没有参数，运行交互式测试
        try:
            # 先运行快速测试
            quick_test()

            # 询问是否继续
            print("\n" + "-" * 60)
            choice = input("\n是否进入交互式菜单？(y/n): ").strip().lower()
            if choice == "y" or choice == "yes":
                interactive_test()
            else:
                # 直接运行批量测试
                print("\n运行批量测试...")
                test_all_files()

        except KeyboardInterrupt:
            print("\n\n程序被中断")
        except Exception as e:
            print(f"\n❌ 发生错误: {e}")
            import traceback

            traceback.print_exc()
