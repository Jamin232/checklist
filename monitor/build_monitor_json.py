#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_monitor_json.py — 物流运营监控看板「数据生成」ETL
======================================================
把最新《产品跟踪表》xlsx 转成前端消费的 data.json。

设计要点（对齐 logistics-dashboard 范式，但更轻量）：
  - 仅依赖 openpyxl（不依赖 pandas），便于在本机轻量部署。
  - 采用「原始行透传」策略：Python 只做「读取 + 列头保留 + 值清洗」，
    把每一行原样输出为 {"列头": 值, ...} 的数组；前端 chayan.js 的
    processData / daily.js 的 parseDailyRows 仍各自完成全部派生计算
    （双维查验、事件链、滞留、周月聚合等），仅数据源由「浏览器上传 xlsx」
    改为「fetch data.json」。这样计算口径与旧版完全一致，且无需在 Python
    侧重写复杂的物流算法，风险最低。
  - 日期统一清洗为 YYYY-MM-DD，数字整数化，空/NaN 转 ''，直接可被
    前端 parseDate / safeNum 消费。

用法：
    python build_monitor_json.py --input "D:/素芸/跟踪表/0905 产品跟踪表.xlsx" --output data.json
    python build_monitor_json.py --dir "D:/素芸/跟踪表"            # 自动取最新
    python build_monitor_json.py                                       # 默认 dir + 默认 output
"""

import argparse
import datetime
import json
import os
import sys

try:
    from openpyxl import load_workbook
except ImportError:
    sys.stderr.write(
        "❌ 缺少依赖 openpyxl。请先安装：\n"
        "   python -m pip install openpyxl\n"
    )
    sys.exit(10)

# ---- 路径 / 匹配配置 ----
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SEARCH_DIR = r"D:/素芸/跟踪表"
DEFAULT_OUTPUT = os.path.join(SCRIPT_DIR, "data.json")

# 主数据 sheet 选择：沿用 chayan.js 的口径（名含该前缀则优先，否则取第一个）
SHEET_HINT = "空运+快递+陆运"
# 候选 Excel 文件名匹配（命中任意一个即视为跟踪表）
NAME_HINTS = ["产品跟踪表", "跟踪表"]


def find_latest_excel(search_dir):
    """在 search_dir 下找最新的跟踪表 xlsx（忽略临时 ~$ 文件）。"""
    if not os.path.isdir(search_dir):
        return None
    import glob
    candidates = []
    for path in glob.glob(os.path.join(search_dir, "*.xlsx")):
        name = os.path.basename(path)
        if name.startswith("~$"):
            continue
        if any(h in name for h in NAME_HINTS):
            candidates.append(path)
    if not candidates:
        candidates = [p for p in glob.glob(os.path.join(search_dir, "*.xlsx"))
                     if not os.path.basename(p).startswith("~$")]
    if not candidates:
        return None
    return max(candidates, key=lambda p: os.path.getmtime(p))


def clean_cell(v):
    """把 openpyxl 读取到的单元格值清洗为前端友好的形式。"""
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        # 日期 → YYYY-MM-DD（前端 parseDate 可直接消费）
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        # 整数值浮点 → int，避免 3.0 之类；其余保留（前端 safeNum 处理）
        if v.is_integer():
            return int(v)
        return v
    if isinstance(v, (int,)):
        return v
    s = str(v).strip()
    # 去掉形如 "YYYY-MM-DD HH:MM:SS" 的尾部时间，仅保留日期
    if len(s) >= 19 and s[4] == "-" and s[10] == " ":
        return s[:10]
    low = s.lower()
    if low in ("nan", "none", "nat", "#na", "null"):
        return ""
    return s


def find_sheet(wb):
    for s in wb.sheetnames:
        if s.strip().startswith(SHEET_HINT):
            return s
    return wb.sheetnames[0]


def build(excel_path, output_path):
    """读取 xlsx 主 sheet，输出 data.json。返回 (meta, row_count)。"""
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    sheet = find_sheet(wb)
    ws = wb[sheet]

    rows_iter = ws.iter_rows(values_only=True)
    try:
        header_raw = list(next(rows_iter))
    except StopIteration:
        wb.close()
        raise ValueError("工作表为空，无法读取表头")

    # 表头：None（合并/空列）→ 占位列名；其余去空格
    header = []
    for i, h in enumerate(header_raw):
        if h is None or str(h).strip() == "":
            header.append("__col%d" % i)
        else:
            header.append(str(h).strip())

    # 定位"仓库出货日期"列（用于 meta.dataDate）
    ship_idx = None
    for i, h in enumerate(header):
        if "仓库出货日期" in h:
            ship_idx = i
            break

    records = []
    max_ship = None
    for r in rows_iter:
        # 补齐到表头长度
        row = list(r) + [""] * (len(header) - len(r))
        obj = {}
        for i, h in enumerate(header):
            val = row[i] if i < len(row) else ""
            obj[h] = clean_cell(val)
        records.append(obj)
        # 跟踪最大发货日（仅当该单元格是真实日期）
        if ship_idx is not None:
            sv = row[ship_idx] if ship_idx < len(row) else None
            if isinstance(sv, (datetime.datetime, datetime.date)):
                d = sv.date() if isinstance(sv, datetime.datetime) else sv
                if max_ship is None or d > max_ship:
                    max_ship = d

    wb.close()

    meta = {
        "sourceFile": os.path.basename(excel_path),
        "sheet": sheet,
        "generatedAt": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "rowCount": len(records),
        "dataDate": max_ship.strftime("%Y-%m-%d") if max_ship else "",
    }
    out = {"meta": meta, "rows": records}
    with open(output_path, "w", encoding="utf-8") as f:
        # 紧凑分隔符，减小体积（前端 fetch 后 JSON.parse 仍很快）
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))

    return meta, len(records)


def main():
    ap = argparse.ArgumentParser(description="产品跟踪表 → data.json ETL")
    ap.add_argument("--input", help="指定跟踪表 Excel（覆盖自动发现）")
    ap.add_argument("--dir", default=DEFAULT_SEARCH_DIR, help="搜索目录，默认 D:/素芸/跟踪表")
    ap.add_argument("--output", default=DEFAULT_OUTPUT, help="输出 data.json 路径")
    args = ap.parse_args()

    excel = args.input or find_latest_excel(args.dir)
    if not excel or not os.path.exists(excel):
        sys.stderr.write("[FAIL] 未在 %s 找到跟踪表 Excel，也未通过 --input 指定。\n" % args.dir)
        sys.exit(1)

    try:
        meta, n = build(excel, args.output)
    except Exception as e:
        sys.stderr.write("[FAIL] 生成 data.json 失败: %s\n" % e)
        sys.exit(2)

    print("[OK] 数据源: %s (sheet=%s)" % (meta["sourceFile"], meta["sheet"]))
    print("[OK] 记录数: %d | 数据基准日: %s" % (n, meta["dataDate"] or "未知"))
    print("[OK] 已写出: %s (%d 字节)" % (args.output, os.path.getsize(args.output)))


if __name__ == "__main__":
    main()
