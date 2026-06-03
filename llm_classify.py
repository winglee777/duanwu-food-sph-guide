#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLM 兜底分类脚本（方案 3）
================================================
作用：
  对 gen_618.py 跑出的 records.json 中 cat_reason="origin" 且 GMV >= 1万
  的记录调用 DeepSeek 重新判断品类，结果汇总成"建议规则"追加到
  category-rules.json 的 _llm_suggestions 区，等待人工审核 → 升级为正式规则。

工作流：
  1. python3 gen_618.py             ← 关键词规则跑一遍
  2. python3 llm_classify.py        ← 对 origin + 头部 GMV 记录调 LLM
  3. 人工 review category-rules.json 里的 _llm_suggestions
  4. 把确认有效的关键词移入 rules.{品类}.keywords
  5. 再跑一次 python3 gen_618.py    ← 应用新规则

设计原则：
  - 缓存：data/llm_cache.json，相同商品名不重复调用
  - 阈值：GMV >= 10000 元（在 GMV_THRESHOLD 调整）
  - 并发：5 路并发请求，速度可控
  - 输出：不直接覆盖 cat 字段，只生成"建议关键词"，人工审核后再升级
"""

import argparse
import json
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

try:
    from openai import OpenAI  # DeepSeek 兼容 OpenAI 协议
except ImportError:
    print("缺少 openai 库，请运行: pip install openai", file=sys.stderr)
    sys.exit(1)


# ============================================================
# 配置
# ============================================================

ROOT = Path(__file__).parent
RECORDS_FILE = ROOT / "data" / "records.json"
RULES_FILE = ROOT / "category-rules.json"
CACHE_FILE = ROOT / "data" / "llm_cache.json"

GMV_THRESHOLD = 10000      # 1万元
CONCURRENCY = 5            # 并发数
MAX_RETRIES = 2            # 失败重试

# 标准品类（必须与 category-rules.json 保持一致）
STANDARD_CATS = [
    "粮油调味", "生鲜", "休闲食品", "饮料冲调", "乳品冷饮",
    "传统滋补", "酒类", "宠物生活", "保健食品/营养补充", "茗茶"
]


# ============================================================
# DeepSeek 客户端
# ============================================================

def get_client():
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        print("❌ 未找到环境变量 DEEPSEEK_API_KEY", file=sys.stderr)
        sys.exit(1)
    return OpenAI(
        api_key=api_key,
        base_url="https://api.deepseek.com/v1"
    )


# ============================================================
# Prompt 设计
# ============================================================

SYSTEM_PROMPT = """你是腾讯广告食品饮料行业商品分类专家。
你需要根据商品标题判断它真实属于哪个品类，并提取最具识别度的关键词。

可选品类（10 选 1，必须严格使用以下名称）：
- 粮油调味（米/面/油/酱油/醋/料酒/调料）
- 生鲜（生肉/海鲜/水果/蔬菜/未加工）
- 休闲食品（零食/糕点/饼干/卤味即食/坚果/糖果/月饼/雪媚娘）
- 饮料冲调（果汁/咖啡/奶茶/苏打水/代餐粉）
- 乳品冷饮（牛奶/酸奶/奶酪/冰淇淋）
- 传统滋补（燕窝/阿胶/人参/海参/枸杞/灵芝）
- 酒类（白酒/红酒/啤酒/果酒）
- 宠物生活（猫粮/狗粮/宠物用品）
- 保健食品/营养补充（维生素/鱼油/益生菌/蛋白粉）
- 茗茶（茶叶/花茶/草本茶）

判断原则：
1. 看商品的实际形态：即食肉脯归休闲食品，生鲜冷冻肉归生鲜
2. 加工品 > 原料：苹果汁=饮料，海参=传统滋补
3. 复合词以主体词为准：猪肉粽=休闲食品（粽子），不归生鲜（猪肉）

返回严格的 JSON 格式（不要多余文字）：
{"cat": "品类名", "keyword": "最有识别度的1-3字关键词"}"""


def build_user_prompt(name: str, original_cat: str, sample_count: int = 1) -> str:
    return f"""商品标题：{name}
商家原始类目：{original_cat}

请判断真实品类并给出最具识别度的关键词（如"雪媚娘"、"鸭脖"、"NFC果汁"）。"""


# ============================================================
# LLM 调用 + 缓存
# ============================================================

def load_cache() -> dict:
    if CACHE_FILE.exists():
        return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    return {}


def save_cache(cache: dict):
    CACHE_FILE.write_text(
        json.dumps(cache, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def parse_llm_response(text):  # type: (str) -> dict | None
    """从 LLM 返回里提取 JSON"""
    # 去掉 ```json ... ``` 包裹
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.MULTILINE)
    try:
        obj = json.loads(text)
        if "cat" in obj and "keyword" in obj:
            cat = obj["cat"].strip()
            kw = obj["keyword"].strip()
            if cat in STANDARD_CATS and kw:
                return {"cat": cat, "keyword": kw}
    except Exception:
        pass
    return None


def classify_one(client, record: dict, cache: dict) -> dict:
    """单条分类，带缓存与重试"""
    name = record["name"]
    if name in cache:
        return {**record, "llm": cache[name], "from_cache": True}

    user_prompt = build_user_prompt(name, record["cat_origin"])

    last_err = None
    for attempt in range(MAX_RETRIES + 1):
        try:
            resp = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.0,
                max_tokens=80,
                timeout=20
            )
            content = resp.choices[0].message.content
            parsed = parse_llm_response(content)
            if parsed:
                cache[name] = parsed
                return {**record, "llm": parsed, "from_cache": False}
            last_err = f"无法解析: {content[:100]}"
        except Exception as e:
            last_err = str(e)[:120]

    return {**record, "llm": None, "error": last_err}


# ============================================================
# 主流程
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="LLM 兜底分类")
    parser.add_argument("--threshold", type=float, default=GMV_THRESHOLD,
                        help=f"GMV 阈值，默认 {GMV_THRESHOLD}")
    parser.add_argument("--limit", type=int, default=0,
                        help="测试限流，>0 时只跑前 N 条")
    parser.add_argument("--concurrency", type=int, default=CONCURRENCY,
                        help=f"并发数，默认 {CONCURRENCY}")
    args = parser.parse_args()

    print("【LLM 兜底分类 · DeepSeek】\n")

    if not RECORDS_FILE.exists():
        print(f"❌ {RECORDS_FILE.name} 不存在，请先跑 gen_618.py", file=sys.stderr)
        sys.exit(1)

    # 1. 加载数据
    print("[1/5] 加载数据")
    records = json.loads(RECORDS_FILE.read_text(encoding="utf-8"))
    rules = json.loads(RULES_FILE.read_text(encoding="utf-8"))
    cache = load_cache()
    print(f"  · 总记录: {len(records)}")
    print(f"  · 缓存命中候选: {len(cache)}")

    # 2. 筛选
    print(f"\n[2/5] 筛选待 LLM 处理的记录（origin + GMV >= {args.threshold}）")
    targets = [
        r for r in records
        if r["cat_reason"] == "origin" and r["gmv"] >= args.threshold
    ]
    targets.sort(key=lambda x: -x["gmv"])  # GMV 降序，让重要的先跑
    if args.limit > 0:
        targets = targets[:args.limit]
    print(f"  · 待处理: {len(targets)} 条")
    if not targets:
        print("  · 无需处理，退出")
        return

    # 3. 并发调用 LLM
    print(f"\n[3/5] LLM 分类（并发 {args.concurrency}）")
    client = get_client()
    results = []
    cache_hits = 0
    api_calls = 0
    errors = 0

    with ThreadPoolExecutor(max_workers=args.concurrency) as ex:
        futures = {ex.submit(classify_one, client, r, cache): r for r in targets}
        for i, fut in enumerate(as_completed(futures), 1):
            res = fut.result()
            results.append(res)
            if res.get("llm"):
                if res.get("from_cache"):
                    cache_hits += 1
                else:
                    api_calls += 1
            else:
                errors += 1

            if i % 50 == 0 or i == len(targets):
                print(f"  · 进度 {i}/{len(targets)} · "
                      f"缓存命中 {cache_hits} · API 调用 {api_calls} · 错误 {errors}")

    # 落盘缓存
    save_cache(cache)
    print(f"  · 缓存已写入 {CACHE_FILE.name} ({len(cache)} 条)")

    # 4. 聚合关键词建议（按品类分组，只保留 LLM 给出的、不在现有规则里的关键词）
    print("\n[4/5] 聚合 LLM 建议为新规则")
    existing_keywords = set()
    for cat in STANDARD_CATS:
        for kw in rules.get("rules", {}).get(cat, []):
            existing_keywords.add(kw)

    suggestions = {}  # {cat: {keyword: [(name, gmv, original_cat), ...]}}
    cat_changes = {}  # {(原cat, 新cat): count}

    for r in results:
        llm = r.get("llm")
        if not llm:
            continue
        new_cat = llm["cat"]
        kw = llm["keyword"]
        old_cat = r["cat"]
        if new_cat == old_cat:
            continue  # LLM 同意现状，不需要建议
        if kw in existing_keywords:
            continue  # 已有规则覆盖，不重复

        suggestions.setdefault(new_cat, {}).setdefault(kw, []).append({
            "name": r["name"],
            "gmv": r["gmv"],
            "rank": r["rank"],
            "original_cat": old_cat
        })
        cat_changes[(old_cat, new_cat)] = cat_changes.get((old_cat, new_cat), 0) + 1

    print(f"  · LLM 给出建议覆盖 {sum(c for c in cat_changes.values())} 条记录")
    print(f"  · 涉及新关键词: {sum(len(v) for v in suggestions.values())} 个")
    print(f"  · 主要重分类方向（Top 10）:")
    for (old, new), n in sorted(cat_changes.items(), key=lambda x: -x[1])[:10]:
        print(f"      [{old} → {new}] {n} 条")

    # 5. 写回 category-rules.json 的 _llm_suggestions 区
    print("\n[5/5] 写回 category-rules.json")
    # 整理建议格式
    formatted = {
        "_comment": "LLM 自动生成的关键词建议，需人工审核后再移入 rules.{品类}.keywords",
        "_generated_at": __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "_total_records_affected": sum(c for c in cat_changes.values()),
        "_total_new_keywords": sum(len(v) for v in suggestions.values()),
        "by_category": {}
    }
    for cat in sorted(suggestions.keys()):
        kws = suggestions[cat]
        # 每个关键词带上前 3 个样例和总命中数
        kw_list = []
        for kw in sorted(kws.keys(), key=lambda k: -sum(s["gmv"] for s in kws[k])):
            samples = sorted(kws[kw], key=lambda x: -x["gmv"])
            total_gmv = sum(s["gmv"] for s in samples)
            kw_list.append({
                "keyword": kw,
                "hit_count": len(samples),
                "total_gmv": round(total_gmv, 2),
                "examples": [
                    {
                        "name": s["name"][:60],
                        "gmv_w": round(s["gmv"] / 10000, 1),
                        "from_cat": s["original_cat"]
                    } for s in samples[:3]
                ]
            })
        formatted["by_category"][cat] = kw_list

    rules["_llm_suggestions"] = formatted
    RULES_FILE.write_text(
        json.dumps(rules, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    print(f"  · 建议已写入 {RULES_FILE.name} → _llm_suggestions 区")
    print(f"\n✅ 完成。请人工审核 _llm_suggestions，确认有效后移入 rules.{{品类}}.keywords，再重跑 gen_618.py")


if __name__ == "__main__":
    main()
