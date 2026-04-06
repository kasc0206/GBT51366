#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
同步工具使用示例

演示如何使用 sync.py 模块实现 Excel-Word 数据同步
"""

import os
from sync import SyncManager, SyncConfig, CellMapping, TableMapping

def example_basic_sync():
    """示例1: 基础单元格同步"""
    print("=" * 60)
    print("示例1: 基础单元格同步")
    print("=" * 60)
    
    # 1. 创建配置
    config = SyncConfig(default_sheet_name="项目基本信息")
    
    # 2. 添加映射关系
    config.mappings = [
        # Excel单元格 → Word占位符
        CellMapping("D4", "项目名称"),
        CellMapping("D5", "建筑地址"),
        CellMapping("D6", "建筑类型"),
        CellMapping("D7", "建筑面积"),
        CellMapping("D14", "计算年度"),
    ]
    
    # 3. 执行同步
    sync = SyncManager(config)
    sync.load_excel("公共建筑碳排放计算表.xlsx")
    sync.load_word("公共建筑碳排放计算报告.docx")
    
    stats = sync.sync_excel_to_word("输出_基础同步.docx")
    print(f"✅ 同步完成: {stats}")


def example_table_sync():
    """示例2: 表格数据同步"""
    print("\n" + "=" * 60)
    print("示例2: 表格数据同步")
    print("=" * 60)
    
    config = SyncConfig()
    
    # 添加表格映射
    config.table_mappings = [
        # Excel范围 → Word表格索引
        TableMapping("A5:D15", 0, sheet_name="运行阶段-能源", header_row=True),
        TableMapping("A5:H60", 1, sheet_name="建材生产运输", header_row=True),
    ]
    
    sync = SyncManager(config)
    sync.load_excel("公共建筑碳排放计算表.xlsx")
    sync.load_word("公共建筑碳排放计算报告.docx")
    
    stats = sync.sync_excel_to_word("输出_表格同步.docx")
    print(f"✅ 同步完成: {stats}")


def example_full_sync():
    """示例3: 完整同步（单元格+表格）"""
    print("\n" + "=" * 60)
    print("示例3: 完整同步（单元格+表格）")
    print("=" * 60)
    
    config = SyncConfig(default_sheet_name="项目基本信息")
    
    # 单元格映射
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
    
    # 表格映射
    config.table_mappings = [
        TableMapping("A5:J20", 0, sheet_name="运行阶段-能源", header_row=True),
    ]
    
    sync = SyncManager(config)
    sync.load_excel("公共建筑碳排放计算表.xlsx")
    sync.load_word("公共建筑碳排放计算报告.docx")
    
    stats = sync.sync_excel_to_word("输出_完整同步.docx")
    print(f"✅ 同步完成:")
    print(f"   单元格同步: {stats['cells_synced']} 个")
    print(f"   表格同步: {stats['tables_synced']} 个")


def example_command_line():
    """示例4: 命令行用法"""
    print("\n" + "=" * 60)
    print("示例4: 命令行用法")
    print("=" * 60)
    print("""
# 基础同步
python sync.py --excel 公共建筑碳排放计算表.xlsx \\
               --word 公共建筑碳排放计算报告.docx \\
               --output 输出_命令行.docx

# 详细输出
python sync.py -e 公共建筑碳排放计算表.xlsx \\
               -w 公共建筑碳排放计算报告.docx \\
               -o 输出_详细.docx \\
               -v

# 反向同步（待实现）
python sync.py -e data.xlsx \\
               -w report.docx \\
               -o output.xlsx \\
               -d word_to_excel
    """)


if __name__ == "__main__":
    # 运行所有示例
    example_basic_sync()
    example_table_sync()
    example_full_sync()
    example_command_line()
