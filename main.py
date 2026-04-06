#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公共建筑碳排放计算报告系统 - 统一入口

用法:
    python main.py                  # 生成所有文件
    python main.py --excel          # 仅生成Excel
    python main.py --word           # 仅生成Word
    python main.py --vba            # 仅生成VBA宏代码
    python main.py --all            # 生成所有文件（默认）
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
    parser = argparse.ArgumentParser(description="公共建筑碳排放计算报告系统")
    parser.add_argument("--excel", action="store_true", help="仅生成Excel计算表格")
    parser.add_argument("--word", action="store_true", help="仅生成Word报告模板")
    parser.add_argument("--vba", action="store_true", help="仅生成VBA宏代码")
    parser.add_argument("--sync", action="store_true", help="同步Excel数据到Word")
    parser.add_argument("--all", action="store_true", help="生成所有文件（默认）")
    parser.add_argument("--output", "-o", help="同步输出文件路径")
    
    args = parser.parse_args()
    
    output_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 同步模式
    if args.sync:
        excel_path = os.path.join(output_dir, "公共建筑碳排放计算表.xlsx")
        word_path = os.path.join(output_dir, "公共建筑碳排放计算报告.docx")
        output_path = args.output or os.path.join(output_dir, "公共建筑碳排放计算报告_生成版.docx")
        
        sync_excel_word(excel_path, word_path, output_path)
        return
    
    # 如果没有指定任何参数，默认生成所有
    if not any([args.excel, args.word, args.vba, args.all]):
        args.all = True
    
    success_count = 0
    total_count = 0
    
    if args.excel or args.all:
        total_count += 1
        if run_script("generate_excel.py", "生成Excel计算表格"):
            success_count += 1
    
    if args.word or args.all:
        total_count += 1
        if run_script("generate_word.py", "生成Word报告模板"):
            success_count += 1
    
    if args.vba or args.all:
        total_count += 1
        if run_script("create_vba_macros.py", "生成VBA宏代码"):
            success_count += 1
    
    # 打印总结
    print(f"\n{'='*60}")
    print(f"🎉 生成完成！成功 {success_count}/{total_count}")
    print(f"{'='*60}")
    
    # 列出生成的文件
    output_dir = os.path.dirname(os.path.abspath(__file__))
    output_files = [
        "公共建筑碳排放计算表.xlsx",
        "公共建筑碳排放计算报告.docx",
        "Excel_Word_Sync_Macros.bas",
    ]
    
    print("\n📁 生成的文件:")
    for filename in output_files:
        filepath = os.path.join(output_dir, filename)
        if os.path.exists(filepath):
            size = os.path.getsize(filepath)
            print(f"   ✅ {filename} ({size/1024:.1f} KB)")
        else:
            print(f"   ❌ {filename} (未生成)")
    
    print(f"\n💡 使用说明:")
    print(f"   1. 打开 Excel 文件并填写黄色单元格")
    print(f"   2. 同步数据到 Word: python main.py --sync")
    print(f"   3. 或使用 VBA 宏（需要 Excel 启用宏）")
    print(f"   4. 查看 README.md 获取详细说明")


if __name__ == "__main__":
    main()
