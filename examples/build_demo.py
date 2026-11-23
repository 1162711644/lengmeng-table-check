#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成一份故意做脏的示例表，用于验证 table_check.py 的效果。

用法：
    python3 examples/build_demo.py
产出：
    examples/demo_销售明细.xlsx
"""

import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "demo_销售明细.xlsx")


def main():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "销售明细"

    # 检查项9：首行合并大标题占多列
    ws.merge_cells("A1:J1")
    ws["A1"] = "2026年第一季度销售台账"

    heads = ["部门", "订单号", "客户信息", "商品", "数量", "金额", "日期", "状态", "备注", "月"]
    for i, h in enumerate(heads, 1):
        ws.cell(row=2, column=i, value=h)

    rows = [
        ["销售一部", "DD001", "张三 13800138000", "A件", 5, "1,200元", "3月2日", "已完成", None, "数据"],
        [None, "DD002", "李四 13900139000 北京市", "B件", "5件", "1200", "2026-03-02", "已完成", None, "数据"],
        [None, "DD003", "王五", "C件", 5, "约1200", "03/02", "已完成", None, "数据"],
        [None, "DD004", "赵六", "D件", "5个", "￥1300", "2026/3/5", "已完成", None, None],
        ["销售二部", "DD005", "孙七", "E件", None, "1500", "3月8日", "已完成", None, None],
        [None, "DD006", "周八", "F件", 3, "1500", "3月9日", "已完成", None, None],
        [None, None, "吴九", "G件", 3, None, "3月10日", "已完成", None, None],
        ["销售三部", "DD001", "张三 13800138000", "A件", 5, "1,200元", "3月2日", "已完成", None, "数据"],
        [None, "DD008", "郑十", "H件", 2, "1000", "3月12日", "已完成", None, None],
    ]
    r = 3
    for row in rows:
        for i, v in enumerate(row, 1):
            if v is not None:
                ws.cell(row=r, column=i, value=v)
        r += 1

    # 检查项1：分类列合并
    ws.merge_cells("A3:A5")

    # 检查项7：表内空行 + 游离合计行
    ws.cell(row=r, column=1, value="合计")
    ws.cell(row=r, column=6, value="9200")

    # 第二个 sheet：字段名与主表不一致（检查项10）
    ws2 = wb.create_sheet("库存")
    heads2 = ["部门", "订单编号", "客户名称", "数量", "金额（元）", "日期"]
    for i, h in enumerate(heads2, 1):
        ws2.cell(row=1, column=i, value=h)
    for i in range(2, 6):
        ws2.cell(row=i, column=1, value="销售一部")
        ws2.cell(row=i, column=2, value=f"DD{i:03d}")
        ws2.cell(row=i, column=3, value="张三")
        ws2.cell(row=i, column=4, value=3)
        ws2.cell(row=i, column=5, value=100)
        ws2.cell(row=i, column=6, value="2026-03-01")

    wb.save(OUT)
    print(f"已生成示例表：{OUT}")
    print(f"试跑：python3 scripts/table_check.py {os.path.relpath(OUT)}")


if __name__ == "__main__":
    main()
