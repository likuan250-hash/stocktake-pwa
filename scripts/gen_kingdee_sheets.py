#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
金蝶盘点单静态同步库生成脚本
================================
把「金蝶云星空」盘点单(STK_StockCountInput)的物料主数据，生成为前端只读静态库
js/kingdee-sheets.js（window.KINGDEE_SHEETS）。App 内「从金蝶盘点单导入」按单号选单导入。

数据准备（需开发者经 KingdeeMCP 手动拉取，无法全自动）：
  1. 用 KingdeeMCP 查询盘点单：form_id=STK_StockCountInput，
     过滤 FStockOrgId.FName in ('测试89','中央厨房') 且 FDocumentStatus='C'（已审核）。
  2. 用 kingdee_view_bill 拉每张单的物料分录，保存为「干净 entries JSON」，
     字段为：code / name / spec / unit / warehouse（每行一个物料）。
     ⚠️ 关键坑：中央厨房单的 Specification 是 list-of-dicts
        [{Key:2052, Value:'中文规格'}]，务必取 Key==2052 的 Value，
        直接 str(list) 会把整段 Python repr 写进规格字段（曾污染 1356 条）。
  3. 把各单 entries JSON 放到 KD_DATA_DIR（默认 scripts/_data/），
     命名如 pdzy016370_entries.json。
  4. 在 CENTRAL 列表登记单号/组织/日期/fid/文件名；测试89 单(0010012)写死在 TEST89_RAW。

运行：
  python scripts/gen_kingdee_sheets.py
  KD_DATA_DIR=/path/to/entries python scripts/gen_kingdee_sheets.py

依赖：Python 3.8+，仅标准库（json / os / datetime）。
"""
import json, os, datetime

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(REPO_ROOT, "js", "kingdee-sheets.js")
DATA_DIR = os.environ.get("KD_DATA_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "_data"))

# 中央厨房 4 张盘点单（全部已授权纳入同步库）
# fid / 日期来自 KingdeeMCP 查询 STK_StockCountInput（FStockOrgId.FName='中央厨房'）
CENTRAL = [
    {"billNo": "PDZY016370", "org": "中央厨房", "date": "2026-06-30", "fid": "116421",
     "file": "pdzy016370_entries.json"},
    {"billNo": "PDZY016366", "org": "中央厨房", "date": "2026-05-31", "fid": "116417",
     "file": "pdzy016366_entries.json"},
    {"billNo": "PDZY016365", "org": "中央厨房", "date": "2026-04-28", "fid": "116416",
     "file": "pdzy016365_entries.json"},
    {"billNo": "PDZY016362", "org": "中央厨房", "date": "2026-04-20", "fid": "116413",
     "file": "pdzy016362_entries.json"},
]

# 测试89 PDZY016356（来自 kingdee_view_bill 实测，单条）
TEST89_RAW = [
    {"code": "0010012", "name": "海天柱候酱", "spec": "6.5kg*2桶/件", "unit": "桶", "warehouse": "干货库"},
]


def clean(v):
    return (v or "").strip()


def to_spec(v):
    # 兼容 list-of-dicts（中央厨房 spec 原始形态 [{Key:2052,Value:'中文'}]）
    if isinstance(v, list):
        for it in v:
            if isinstance(it, dict) and it.get("Key") == 2052:
                return clean(it.get("Value"))
        for it in v:
            if isinstance(it, dict) and it.get("Value"):
                return clean(it.get("Value"))
        return ""
    return clean(v)


def to_mat(r):
    return {
        "code": clean(r.get("code")),
        "name": clean(r.get("name")),
        "spec": to_spec(r.get("spec")),
        "unit": clean(r.get("unit")),
        "warehouse": clean(r.get("warehouse")),
    }


def dedupe(mats):
    seen = {}
    out = []
    dup = 0
    for m in mats:
        c = m["code"]
        if not c:
            continue
        if c in seen:
            dup += 1
            continue
        seen[c] = True
        out.append(m)
    return out, dup


def main():
    if not os.path.isdir(DATA_DIR):
        print("⚠️  数据目录不存在: " + DATA_DIR)
        print("    请先把各单 entries JSON 放到该目录（见脚本顶部说明），或用 KD_DATA_DIR 指定。")
        raise SystemExit(1)

    sheets = []
    total = 0
    log_lines = []

    tz_mats, tz_dup = dedupe([to_mat(r) for r in TEST89_RAW])
    sheets.append({
        "billNo": "PDZY016356", "org": "测试89", "date": "2026-04-03", "fid": "116407",
        "materials": tz_mats,
    })
    total += len(tz_mats)
    log_lines.append("   测试89 PDZY016356: %d 条 (去重跳过 %d)" % (len(tz_mats), tz_dup))

    for cfg in CENTRAL:
        path = os.path.join(DATA_DIR, cfg["file"])
        if not os.path.isfile(path):
            print("⚠️  缺少文件: " + path + "（跳过 " + cfg["billNo"] + "）")
            continue
        with open(path, "r", encoding="utf-8") as f:
            rows = json.load(f)
        mats, dup = dedupe([to_mat(r) for r in rows])
        sheets.append({
            "billNo": cfg["billNo"], "org": cfg["org"], "date": cfg["date"], "fid": cfg["fid"],
            "materials": mats,
        })
        total += len(mats)
        log_lines.append("   %s %s: %d 条 (去重跳过 %d)" % (cfg["org"], cfg["billNo"], len(mats), dup))

    gen = datetime.date.today().isoformat()
    header = (
        "// 金蝶盘点单静态同步库（由 Senior Developer 通过 KingdeeMCP 按单号拉取生成）\n"
        "// 生成日期: %s | 来源表单: STK_StockCountInput\n" % gen +
        "// 授权组织: 测试89 + 中央厨房(全部单号已永久授权纳入同步库)\n"
        "// 数据用途: 物料档案页「从金蝶盘点单导入」按单号选单 → upsert 进物料主数据\n"
        "// 仅含主数据字段(code/name/spec/unit/warehouse)，不含盘点数量(属盘点单流程)\n"
        "// 重新生成: 改 CENTRAL/TEST89_RAW 后运行 python scripts/gen_kingdee_sheets.py\n"
    )

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(header)
        f.write("window.KINGDEE_SHEETS = ")
        json.dump(sheets, f, ensure_ascii=False, indent=2)
        f.write(";\n")

    print("✅ 已生成 " + OUT)
    print("   盘点单数量: %d" % len(sheets))
    print("   物料去重后总数: %d" % total)
    for line in log_lines:
        print(line)


if __name__ == "__main__":
    main()
