#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公共建筑碳排放计算报告系统 - 精细化版

入口脚本，用于生成和管理精细化计算单元（39个独立工作表）。

用法:
    python main.py                  # 生成精细化Excel
    python main.py --regenerate     # 重新生成
    python main.py --sync           # 同步Excel数据到Word
    python main.py --example        # 运行同步示例
"""

import os
import sys
import subprocess
import argparse


def run_script(script_name, description):
    """运行指定的Python脚本"""
    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), script_name)
    if not os.path.exists(script_path):
        print(f"⚠️  脚本不存在: {script_name}")
        return False

    print(f"\n{'='*60}")
    print(f"📝 {description}")
    print(f"{'='*60}")

    try:
        result = subprocess.run([sys.executable, script_path],
                              capture_output=False, text=True)
        return result.returncode == 0
    except Exception as e:
        print(f"❌ 执行失败: {e}")
        return False


def sync_excel_word(excel_path, word_path, output_path):
    """Excel-Word 同步（OpenXML方式）"""
    from sync import SyncManager, SyncConfig, CellMapping

    print(f"\n{'='*60}")
    print("🔄 同步 Excel → Word (OpenXML)")
    print(f"{'='*60}")

    try:
        config = SyncConfig(default_sheet_name="项目基本信息")
        config.mappings = [
            CellMapping("D4", "项目名称"),
            CellMapping("D5", "建筑地址"),
            CellMapping("D6", "建筑类型"),
            CellMapping("D7", "建筑面积"),
            CellMapping("D14", "计算年度"),
            CellMapping("D18", "年用电量"),
            CellMapping("D19", "年用气量"),
            CellMapping("D20", "年用热量"),
        ]

        sync = SyncManager(config)
        sync.load_excel(excel_path)
        sync.load_word(word_path)

        stats = sync.sync_excel_to_word(output_path)
        print(f"✅ 同步完成:")
        print(f"   单元格同步: {stats['cells_synced']} 个")
        print(f"   表格同步: {stats['tables_synced']} 个")
        print(f"   输出文件: {output_path}")
        return True
    except Exception as e:
        print(f"❌ 同步失败: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(
        description="公共建筑碳排放计算报告系统 - 精细化版（39个独立计算单元）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
    python main.py                  # 生成精细化Excel
    python main.py --regenerate     # 重新生成Excel
    python main.py --sync           # 同步Excel数据到Word
    python main.py --example        # 运行同步示例
        """
    )
    parser.add_argument("--regenerate", action="store_true",
                       help="重新生成精细化Excel计算表格")
    parser.add_argument("--word", action="store_true",
                       help="生成Word报告模板")
    parser.add_argument("--sync", action="store_true",
                       help="同步Excel数据到Word")
    parser.add_argument("--example", action="store_true",
                       help="运行同步示例")
    parser.add_argument("--output", "-o", help="同步输出文件路径")

    args = parser.parse_args()
    output_dir = os.path.dirname(os.path.abspath(__file__))

    # Word模板模式
    if args.word:
        run_script("generate_word_template.py", "生成Word报告模板")
        return

    # 同步模式
    if args.sync:
        templates_dir = os.path.join(output_dir, "templates")
        excel_path = os.path.join(templates_dir, "公共建筑碳排放计算表_精细化版.xlsx")
        word_path = os.path.join(templates_dir, "公共建筑碳排放计算报告.docx")
        output_path = args.output or os.path.join(output_dir, "公共建筑碳排放计算报告_生成版.docx")

        if not os.path.exists(excel_path):
            print(f"⚠️  精细化Excel不存在，正在生成...")
            run_script("generate_granular.py", "生成精细化Excel")

        if not os.path.exists(word_path):
            print(f"❌ Word模板不存在: {word_path}")
            return

        sync_excel_word(excel_path, word_path, output_path)
        return

    # 示例模式
    if args.example:
        run_script("example_sync.py", "运行同步示例")
        return

    # 默认：生成精细化Excel
    run_script("generate_granular.py", "生成精细化计算单元Excel")

    print(f"\n{'='*60}")
    print("💡 使用说明:")
    print(f"   python main.py --sync     # 同步数据到Word")
    print(f"   python main.py --example  # 运行同步示例")
    print(f"   python main.py -h         # 查看帮助")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
