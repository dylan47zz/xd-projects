#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用xlwings打开Excel，刷新计算，然后验证CI值"""
import xlwings as xw
import os, time

OUTPUT = os.path.expanduser('~/Desktop/钢铁行业模型V11.1.xlsx')

YEARS = list(range(2024, 2061))
YC = 3

dashboard = {
    '积极情景': {2024:1.9500,2025:1.9470,2026:1.9050,2027:1.8530,2028:1.8120,2029:1.7800,
                 2030:1.7440,2035:1.3700,2040:0.8500,2050:0.1700,2060:0.0100},
    '适度情景': {2024:1.9500,2025:1.9540,2026:1.9230,2027:1.8910,2028:1.8600,2029:1.8290,
                 2030:1.7980,2035:1.6070,2040:1.1500,2050:0.4000,2060:0.0500},
    '保守情景': {2024:1.9500,2025:1.9580,2026:1.9410,2027:1.9240,2028:1.9070,2029:1.8910,
                 2030:1.8740,2035:1.7510,2040:1.4500,2050:0.7500,2060:0.1200}
}

print("打开Excel并刷新计算...")
app = xw.App(visible=False, add_book=False)
try:
    wb = app.books.open(OUTPUT)
    wb.app.calculate()
    time.sleep(2)
    
    print("\n===== 验证CI（Excel重新计算后）=====")
    all_ok = True
    for sc in ['积极情景', '适度情景', '保守情景']:
        ws = wb.sheets[sc]
        print(f"\n【{sc}】")
        for y in [2024,2025,2026,2027,2028,2029,2030,2035,2040,2050,2060]:
            col = YC + (y-2024)
            # Row 63 = CI, Row 13 = 产量
            ci = ws.cells(63, col).value
            tgt = dashboard[sc].get(y)
            P = ws.cells(13, col).value
            if ci and tgt:
                dev = (ci-tgt)/tgt*100
                flag = "✓" if abs(dev)<=1 else f"⚠️ dev={dev:.1f}%"
                print(f"  {y}: CI={ci:.4f}, 目标={tgt:.4f}, 偏差={dev:+.2f}% {flag}")
                if abs(dev) > 1: all_ok = False
            else:
                print(f"  {y}: CI={ci}, 目标={tgt}")
    
    if all_ok:
        print("\n✅ 所有关键年份偏差 ≤ 1%")
    else:
        print("\n⚠️ 部分年份偏差 > 1%，需要调整")
    
    wb.save()
    wb.close()
finally:
    app.quit()
print("完成")
