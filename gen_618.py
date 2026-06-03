#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
端午食品饮料赛道趋势报告 · 数据加工脚本
================================================
作用：
  - 输入：原始 xlsx（友望视频号小店导出）+ 品类清洗规则 JSON
  - 输出：data/records.json + data/stats.json，供 index.html 通过 fetch 加载

用法：
  python3 gen_618.py
  python3 gen_618.py --xlsx /path/to/source.xlsx --out ./data

设计原则：
  1. 数据与展示分离：HTML 不再硬编码 JS 字面量，全部从 JSON 加载
  2. 品类清洗：透传商家原始类目，但允许通过关键词规则二次校正
     （解决"雪媚娘错归粮油调味"这类问题）
  3. 可复用：本脚本是 seasonal-trend-template 的实例，
     换 xlsx + 改规则即可生成中秋/七夕等同类报告

字段映射（基于源表 16 列）：
  排行 → rank
  品类 → cat (会经过 category-rules.json 二次校正)
  商品名称 → name
  商品主图链接 → img
  商品价格 → price
  商品来源 → shop
  销量 → sales
  销售额 → gmv
  直播销量 → live_sales
  直播销售额 → live_gmv
  关联直播数 → live_count
  关联视频数 → video_count
  是否带货中心 → in_dm
  数据来源 → data_source
  是否有类似在投品 → ad_status (是→matched, 否→nomatch)
  类似在投商品举例 → ad_example
"""

import argparse
import json
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("缺少 openpyxl，请先运行: pip install openpyxl", file=sys.stderr)
    sys.exit(1)


# ============================================================
# 1. 路径与默认配置
# ============================================================

ROOT = Path(__file__).parent
DEFAULT_XLSX = ROOT.parent / "2025年视频号端午食品饮料商品总表.xlsx"
DEFAULT_OUT = ROOT / "data"
RULES_FILE = ROOT / "category-rules.json"


# ============================================================
# 2. 内置品类清洗规则（首次运行时落盘到 category-rules.json）
# ============================================================

DEFAULT_RULES = {
    "_comment": "品类二次校正规则。优先级：keyword 命中即覆盖商家原始类目。",
    "_priority": [
        "休闲食品",
        "生鲜",
        "饮料冲调",
        "乳品冷饮",
        "传统滋补",
        "保健食品/营养补充",
        "茗茶",
        "酒类",
        "宠物生活",
        "粮油调味"
    ],
    "rules": {
        "休闲食品": [
            "雪媚娘", "大福", "糯米糍", "凤梨酥", "蛋黄酥", "月饼",
            "麻薯", "饼干", "曲奇", "薯片", "薯条", "锅巴", "辣条",
            "蛋卷", "面包", "吐司", "欧包", "蛋糕", "巧克力", "糖果",
            "果冻", "果干", "西梅", "蜜饯", "话梅", "山楂", "肉脯",
            "肉松", "肉干", "牛肉干", "猪肉脯", "鱼干", "鱼片",
            "魔芋爽", "魔芋丝", "虾片", "海苔", "坚果", "瓜子",
            "开心果", "巴旦木", "夏威夷果", "腰果", "核桃", "杏仁",
            "豆干", "鹌鹑蛋", "卤蛋", "卤味", "卤鹅", "卤鸡爪",
            "酱牛肉", "盐焗鸡", "鸭脖", "鸭翅", "鸡爪", "猪蹄",
            "粽子", "甜品"
        ],
        "生鲜": [
            "生鲜", "冷冻", "鲜活", "活虾", "活蟹", "鲜虾", "虾仁",
            "虾滑", "鱼", "三文鱼", "鳕鱼", "带鱼", "黄花鱼",
            "海鲜", "扇贝", "鲍鱼", "海参", "蛤蜊", "牡蛎", "海螺",
            "螃蟹", "大闸蟹", "梭子蟹", "龙虾", "皮皮虾",
            "猪肉", "牛肉", "羊肉", "鸡肉", "鸭肉", "鹅肉",
            "排骨", "猪蹄", "猪骨", "牛排", "牛腩", "牛腱",
            "鸡腿", "鸡翅", "鸡胸", "整鸡", "童子鸡", "乌鸡",
            "藜麦鸡排", "鸡排", "牛仔骨",
            "蔬菜", "水果", "榴莲", "山竹", "椰子", "椰青",
            "苹果", "葡萄", "草莓", "蓝莓", "樱桃", "车厘子",
            "西瓜", "哈密瓜", "甜瓜", "桃子", "李子", "杏",
            "海鸭蛋", "咸鸭蛋", "鸭蛋"
        ],
        "饮料冲调": [
            "果汁", "NFC", "椰子水", "椰汁", "苏打水", "气泡水",
            "可乐", "雪碧", "矿泉水", "纯净水",
            "咖啡", "速溶咖啡", "挂耳咖啡", "冷萃",
            "奶茶", "奶昔", "豆奶", "豆浆",
            "燕麦片", "麦片", "代餐粉", "蛋白粉", "营养粉",
            "酸梅汤", "凉茶", "刺梨原液"
        ],
        "乳品冷饮": [
            "牛奶", "纯牛奶", "酸奶", "奶酪", "芝士",
            "黄油", "炼乳", "奶油", "鲜奶",
            "冰淇淋", "雪糕", "冰棒", "冰激凌",
            "牛乳", "AD钙奶"
        ],
        "传统滋补": [
            "燕窝", "阿胶", "人参", "西洋参", "鹿茸", "海参",
            "灵芝", "虫草", "花胶", "鱼胶", "藏红花",
            "枸杞", "黑芝麻", "黑豆", "红枣", "桂圆", "莲子",
            "薏米", "芡实", "山药", "百合", "银耳",
            "膏方", "丸剂", "滋补", "炖品", "汤料包",
            "八珍糕", "八珍粉", "茯苓糕", "桂花糕"
        ],
        "保健食品/营养补充": [
            "维生素", "VC", "VE", "复合维生素", "钙片", "铁剂",
            "鱼油", "DHA", "Omega", "深海鱼油",
            "益生菌", "酵素", "胶原蛋白", "玻尿酸",
            "蛋白粉", "氨基酸", "辅酶Q10", "葡萄籽",
            "螺旋藻", "蜂胶", "蜂王浆", "保健"
        ],
        "茗茶": [
            "茶叶", "绿茶", "红茶", "乌龙茶", "普洱", "白茶", "黄茶",
            "铁观音", "大红袍", "金骏眉", "正山小种", "龙井",
            "碧螺春", "毛峰", "信阳毛尖", "六安瓜片",
            "桑叶茶", "冬瓜茶", "荷叶茶", "草本花茶", "花草茶",
            "玫瑰花茶", "菊花茶", "茉莉花茶"
        ],
        "酒类": [
            "白酒", "红酒", "葡萄酒", "啤酒", "黄酒", "米酒",
            "果酒", "梅酒", "清酒", "伏特加", "威士忌", "白兰地",
            "茅台", "五粮液", "汾酒", "洋河", "剑南春", "泸州老窖"
        ],
        "宠物生活": [
            "猫粮", "狗粮", "宠物", "猫砂", "猫罐头", "狗罐头",
            "宠物零食", "猫条", "冻干", "磨牙棒",
            "宠物用品", "狗笼", "猫窝", "宠物玩具"
        ],
        "粮油调味": [
            "大米", "面粉", "面条", "挂面", "意面", "米线",
            "食用油", "花生油", "菜籽油", "橄榄油", "玉米油",
            "酱油", "生抽", "老抽", "蚝油", "醋", "陈醋", "香醋",
            "料酒", "花雕", "黄酒料酒", "豆瓣酱", "辣椒酱",
            "芝麻酱", "花生酱", "沙拉酱", "番茄酱",
            "盐", "糖", "白砂糖", "冰糖", "红糖", "蜂蜜",
            "调料", "调味", "火锅底料", "干锅酱",
            "白凉粉", "仙草粉", "龟苓膏粉", "凉粉",
            "粉丝", "粉条", "酸辣粉", "螺蛳粉", "小面"
        ]
    },
    "exclude_keywords": {
        "_comment": "如果商品名同时含 exclude_keywords[cat] 中任一词，跳过该 cat 的命中（用于避免误伤）",
        "休闲食品": ["猪肉粽", "鲜肉粽", "蛋黄粽"],
        "生鲜": ["即食", "罐头", "肉脯", "肉松", "肉干"]
    }
}


# ============================================================
# 3. 工具函数
# ============================================================

def load_rules() -> dict:
    """加载品类清洗规则；首次运行自动落盘默认规则"""
    if not RULES_FILE.exists():
        RULES_FILE.write_text(
            json.dumps(DEFAULT_RULES, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
        print(f"  · 首次运行，已生成默认规则文件: {RULES_FILE.name}")
        return DEFAULT_RULES
    with RULES_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


def reclassify(name: str, original_cat: str, rules: dict) -> tuple[str, str]:
    """
    返回 (final_cat, reason)
      - final_cat: 最终品类
      - reason: 'origin' / 'rule:关键词'
    按 _priority 顺序匹配，命中即返回；都不命中则保留原始类目。
    """
    if not name:
        return original_cat, "origin"

    priority = rules.get("_priority", [])
    rule_map = rules.get("rules", {})
    excludes = rules.get("exclude_keywords", {})

    for cat in priority:
        keywords = rule_map.get(cat, [])
        excl = excludes.get(cat, [])
        # 先排除
        if any(e in name for e in excl):
            continue
        # 再匹配
        for kw in keywords:
            if kw in name:
                return cat, f"rule:{kw}"

    return original_cat, "origin"


def parse_xlsx(xlsx_path: Path, rules: dict) -> tuple[list, dict]:
    """解析 xlsx，返回 (records, stats)"""
    print(f"  · 加载 {xlsx_path.name}")
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        raise ValueError("xlsx 数据为空")

    header = rows[0]
    print(f"  · 表头: {header}")
    print(f"  · 总行数（含表头）: {len(rows)}")

    # 字段下标映射
    col = {h: i for i, h in enumerate(header) if h is not None}
    required = ["排行", "品类", "商品名称", "销量", "销售额"]
    for r in required:
        if r not in col:
            raise ValueError(f"缺少必需字段: {r}")

    records = []
    reclassified_count = 0
    reclassified_examples = []

    for row in rows[1:]:
        if row[col["排行"]] is None:
            continue

        name = row[col["商品名称"]] or ""
        original_cat = row[col["品类"]] or "未分类"

        # 二次清洗
        final_cat, reason = reclassify(name, original_cat, rules)
        if reason != "origin" and final_cat != original_cat:
            reclassified_count += 1
            if len(reclassified_examples) < 8:
                reclassified_examples.append(
                    f"    [{original_cat} → {final_cat}] {name[:40]} ({reason})"
                )

        rec = {
            "rank": row[col["排行"]],
            "cat": final_cat,
            "cat_origin": original_cat,
            "cat_reason": reason,
            "name": name,
            "img": row[col.get("商品主图链接", -1)] or "" if "商品主图链接" in col else "",
            "price": row[col.get("商品价格", -1)] or "" if "商品价格" in col else "",
            "shop": row[col.get("商品来源", -1)] or "" if "商品来源" in col else "",
            "sales": row[col["销量"]] or 0,
            "gmv": round(float(row[col["销售额"]] or 0), 2),
            "live_sales": row[col.get("直播销量", -1)] or 0 if "直播销量" in col else 0,
            "live_gmv": round(float(row[col.get("直播销售额", -1)] or 0), 2) if "直播销售额" in col else 0,
            "live_count": row[col.get("关联直播数", -1)] or 0 if "关联直播数" in col else 0,
            "video_count": row[col.get("关联视频数", -1)] or 0 if "关联视频数" in col else 0,
            "ad_status": "matched" if (row[col.get("是否有类似在投品", -1)] or "") == "是" else "nomatch",
            "ad_example": row[col.get("类似在投商品举例", -1)] or "" if "类似在投商品举例" in col else "",
        }
        records.append(rec)

    print(f"  · 有效记录数: {len(records)}")
    print(f"  · 品类重分类记录: {reclassified_count} 条")
    if reclassified_examples:
        print("  · 重分类样例:")
        for line in reclassified_examples:
            print(line)

    # 聚合 STATS
    stats = {}
    for r in records:
        cat = r["cat"]
        if cat not in stats:
            stats[cat] = {"count": 0, "total_sales": 0, "total_gmv": 0.0}
        stats[cat]["count"] += 1
        stats[cat]["total_sales"] += r["sales"]
        stats[cat]["total_gmv"] += r["gmv"]

    # 排序：按 GMV 降序
    stats = dict(sorted(stats.items(), key=lambda x: -x[1]["total_gmv"]))
    for cat, s in stats.items():
        s["total_gmv"] = round(s["total_gmv"], 2)

    return records, stats


# ============================================================
# 4. 主流程
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="端午食饮趋势数据加工")
    parser.add_argument("--xlsx", type=Path, default=DEFAULT_XLSX, help="原始 xlsx 路径")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT, help="输出目录")
    args = parser.parse_args()

    print("【端午食饮趋势 · 数据加工】")
    print(f"  · 输入: {args.xlsx}")
    print(f"  · 输出: {args.out}")
    print()

    if not args.xlsx.exists():
        print(f"❌ 输入文件不存在: {args.xlsx}", file=sys.stderr)
        sys.exit(1)

    args.out.mkdir(parents=True, exist_ok=True)

    # Step 1: 加载规则
    print("[1/3] 加载品类清洗规则")
    rules = load_rules()
    print()

    # Step 2: 解析 xlsx + 二次清洗
    print("[2/3] 解析 xlsx 并执行品类二次校正")
    records, stats = parse_xlsx(args.xlsx, rules)
    print()

    # Step 3: 输出 JSON
    print("[3/3] 写出 JSON")
    records_path = args.out / "records.json"
    stats_path = args.out / "stats.json"

    with records_path.open("w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, separators=(",", ":"))
    with stats_path.open("w", encoding="utf-8") as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)

    print(f"  · {records_path.relative_to(ROOT)} ({records_path.stat().st_size / 1024:.1f} KB)")
    print(f"  · {stats_path.relative_to(ROOT)} ({stats_path.stat().st_size / 1024:.1f} KB)")
    print()

    print("✅ 完成。品类汇总：")
    for cat, s in stats.items():
        gmv_w = s["total_gmv"] / 10000
        print(f"  - {cat:<14} {s['count']:>5} 个 SKU · 销量 {s['total_sales']:>10,} · GMV ¥{gmv_w:>8,.1f}w")


if __name__ == "__main__":
    main()
