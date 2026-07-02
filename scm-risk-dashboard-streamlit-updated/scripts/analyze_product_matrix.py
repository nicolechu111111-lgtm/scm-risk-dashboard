import json
import math
import re
from collections import Counter, defaultdict

import pandas as pd


SRC = "/Users/blue/Downloads/SCM Follow Up.xlsx"
OUT = "outputs/product_matrix/product_matrix_data.json"


def clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def num(v):
    try:
        if clean(v) == "":
            return None
        return float(v)
    except Exception:
        return None


def classify(name, sku, material, cost):
    n = name.upper()
    m = material.upper()

    if any(k in n for k in ["VIEWING CAGE", "VIEW CAGE", "TV VIEWING CAGE", "TUNNEL CAGE", "RABBIT CAGE", "OUTDOOR CAGE", "TRAVEL CAGE", "MODERN VIEWING CAGE"]):
        main = "笼具与外出箱"
        if "TRAVEL" in n or "OUTDOOR" in n:
            sub = "外出/临时安置笼"
        elif "RABBIT CAGE" in n:
            sub = "兔用基础笼"
        else:
            sub = "展示型主笼"
    elif "WHEEL" in n:
        main = "运动跑轮"
        if any(k in n for k in ["XL", "11.5"]):
            sub = "XL 大尺寸静音跑轮"
        elif "6.7" in n:
            sub = "小尺寸静音跑轮"
        else:
            sub = "8.5 英寸静音跑轮"
    elif any(k in n for k in ["WATER BOTTLE", "DRINKING BOTTLE", "WTR BTL", "WATER DISH", "BOTTLE HLDR"]):
        main = "饮水与喂食"
        if "HLDR" in n or "HOLDER" in n:
            sub = "水瓶支架"
        elif "DISH" in n:
            sub = "重力水碗/水盘"
        else:
            sub = "饮水瓶"
    elif any(k in n for k in ["BOWL", "DISH", "FEEDER", "HAY"]):
        main = "饮水与喂食"
        if "HAY" in n or "FEEDER" in n:
            sub = "草架/喂草器"
        else:
            sub = "食碗/食盆"
    elif any(k in n for k in ["BEDDING", "PAPER TWIST", "PAPER PEBBLES", "NESTING", "COTTON"]):
        main = "垫料与窝材"
        if any(k in n for k in ["COTTON", "NESTING"]):
            sub = "棉球/筑窝材料"
        elif "PEBBLES" in n:
            sub = "纸粒垫料"
        elif "TWIST" in n:
            sub = "纸绳/纸条垫料"
        elif "FLAKE" in n:
            sub = "纸片垫料"
        elif "STRAW" in n:
            sub = "纸吸管垫料"
        else:
            sub = "纸质垫料"
    elif any(k in n for k in ["TOILET", "POTTY"]):
        main = "厕所与清洁"
        if "RABBIT" in n or "BIG" in n:
            sub = "兔用厕所"
        else:
            sub = "小宠厕所"
    elif any(k in n for k in ["CHEW", "TREAT"]):
        main = "咀嚼/磨牙"
        if any(k in n for k in ["DECOR", "FOREST", "CAMP", "LUNCH", "TREE"]):
            sub = "场景化磨牙组合"
        else:
            sub = "基础磨牙单品"
    elif any(k in n for k in ["HIDE", "CABIN", "HOUSE", "CASTLE", "BURROW", "BRIDGE", "LADDER", "FENCE", "STAIRCASE", "SWING", "SHOPPE", "MOUNTAIN", "SCUBA", "CORAL"]):
        main = "木制窝屋与陈列玩具"
        if any(k in n for k in ["BRIDGE", "LADDER", "FENCE", "STAIRCASE"]):
            sub = "桥梯/围栏"
        elif any(k in n for k in ["HIDE", "CABIN", "HOUSE", "CASTLE", "BURROW", "SHOPPE", "MOUNTAIN"]):
            sub = "窝屋/躲避屋"
        else:
            sub = "陈列互动玩具"
    else:
        if "CERAMIC" in m:
            main, sub = "饮水与喂食", "陶瓷食器/造型件"
        elif "WOOD" in m or "BASSWOOD" in m:
            main, sub = "木制窝屋与陈列玩具", "陈列互动玩具"
        else:
            main, sub = "其他小宠配件", "待归类配件"

    high = main in ["笼具与外出箱"] or (main == "运动跑轮" and ("XL" in sub or (cost or 0) >= 9)) or (cost or 0) >= 9
    traffic = main in ["垫料与窝材", "厕所与清洁"] or sub in ["饮水瓶", "食碗/食盆", "草架/喂草器"] and (cost or 99) <= 1.6 or (cost or 99) <= 1.2
    if high:
        channel = "高端形象款"
    elif traffic:
        channel = "流通走量款"
    else:
        channel = "中端利润款"
    return main, sub, channel


def animal_line(name, sku, main, sub):
    n = name.upper()
    s = sku.upper()
    if (
        "RABBIT" in n
        or s.startswith("RJ")
        or sku in {"CPTL0800", "CPCJ0100-OR", "CPCS1600-OR"}
        or "HAY FEEDER" in n
        or "GRASS FEEDER" in n
    ):
        return "兔/豚鼠类"
    if main == "垫料与窝材" or sub in {"饮水瓶", "食碗/食盆", "基础磨牙单品"}:
        return "通用小宠耗材"
    return "仓鼠/小型啮齿类"


def redundancy_group(name, sku):
    n = name.upper()
    if re.search(r"SILENT.*RUN|WHEEL", n):
        if "XL" in n or "11.5" in n:
            return "静音跑轮 XL：颜色/渠道重复"
        if "6.7" in n:
            return "静音跑轮 6.7：颜色重复"
        return "静音跑轮 8.5：颜色/渠道重复"
    if any(k in n for k in ["PAPER PEBBLES", "PAPER PEBBLE"]):
        return "纸粒垫料：颜色/规格重复"
    if "PAPER TWIST" in n or "PAPER TWISTS" in n:
        return "纸绳垫料：颜色重复"
    if "COTTON" in n or "NESTING BALLS" in n:
        return "棉球筑窝材料：颜色重复"
    if "TRAVEL CAGE" in n:
        return "外出笼：颜色/主题重复"
    if "CRYSTAL OUTDOOR CAGE" in n:
        return "Crystal Outdoor Cage：颜色重复"
    if "RABBIT DRINKING BOTTLE" in n:
        return "兔用饮水瓶：容量阶梯"
    if "SIMPLE RABBIT BOWL" in n:
        return "兔碗：颜色/材质重复"
    if "WOODLAND POTTY" in n:
        return "Woodland Potty：颜色/渠道重复"
    if "TULIP WATER BOTTLE" in n or "BUBBLE WATER BOTTLE" in n or "CLOUD WATER BOTTLE" in n or "MAG" in n and "WTR" in n:
        return "小宠饮水瓶：造型功能重叠"
    if any(k in n for k in ["BRIDGE", "LADDER", "FENCE", "STAIRCASE"]):
        return "桥梯围栏：功能重叠"
    if any(k in n for k in ["CANDY", "MUSHROOM", "TREEHOUSE", "COTTAGE HIDE", "TREE STUMP", "CHEESE SHOP", "CASTLE", "CAKE HOUSE", "CHOCOLATE MOUNTAIN"]) and not any(k in n for k in ["BOWL", "DISH"]):
        return "木制/陶瓷躲避屋：主题功能重叠"
    return ""


def package_attr(row, channel):
    cp = clean(row.get("Case Pack"))
    moq = clean(row.get("MOQ"))
    case_type = clean(row.get("Case Type"))
    customer = clean(row.get("Customer"))
    parts = []
    if channel == "流通走量款":
        parts.append("门店基础常备，适合整箱补货")
    elif channel == "中端利润款":
        parts.append("作为利润补充与主题陈列款")
    else:
        parts.append("控制门店样机/少量陈列，按订单补货")
    if cp:
        parts.append(f"{cp}/箱")
    if moq:
        parts.append(f"MOQ {moq}")
    if case_type:
        parts.append(f"箱型 {case_type}")
    if customer:
        parts.append(f"现有渠道：{customer}")
    return "；".join(parts)


product = pd.read_excel(SRC, sheet_name="Product List", dtype=str)
summary = pd.read_excel(SRC, sheet_name="Sum", dtype=str)
summary = summary.rename(columns=lambda x: str(x).replace("\n", " ").strip())

sum_cols = ["Item Code", "Onhand  Stock", "In-Transit", "Onhand  Available", "Weekly  Forecast", "Inventory  Turnover", "Purchase  Suggestion"]
sum_map = {}
for _, r in summary.iterrows():
    sku = clean(r.get("Item Code"))
    if not sku:
        continue
    sum_map[sku] = {c: clean(r.get(c)) for c in sum_cols if c in summary.columns}

rows = []
for _, r in product.iterrows():
    sku = clean(r.get("Item Code"))
    if not sku:
        continue
    cost = num(r.get("Cost"))
    main, sub, channel = classify(clean(r.get("Product Name")), sku, clean(r.get("Material")), cost)
    animal = animal_line(clean(r.get("Product Name")), sku, main, sub)
    stock = sum_map.get(sku, {})
    core = []
    if clean(r.get("Material")):
        core.append(f"材质：{clean(r.get('Material'))}")
    if cost is not None:
        core.append(f"成本：{cost:.2f}")
    if clean(r.get("Length")) and clean(r.get("Width")) and clean(r.get("Height")):
        core.append(f"裸品尺寸：{clean(r.get('Length'))}x{clean(r.get('Width'))}x{clean(r.get('Height'))}")
    if clean(r.get("Case Pack")):
        core.append(f"箱规：{clean(r.get('Case Pack'))}/箱")
    if clean(stock.get("Weekly  Forecast")):
        core.append(f"周预测：{clean(stock.get('Weekly  Forecast'))}")
    group = redundancy_group(clean(r.get("Product Name")), sku)
    rows.append({
        "一级大类": main,
        "二级细分品类": sub,
        "适用动物线": animal,
        "渠道定位": channel,
        "SKU 编码": sku,
        "产品名称": clean(r.get("Product Name")),
        "产品核心属性": "；".join(core),
        "库存铺货属性": package_attr(r, channel),
        "冗余组": group,
        "成本": cost,
        "现有渠道": clean(r.get("Customer")),
        "Case Pack": clean(r.get("Case Pack")),
        "MOQ": clean(r.get("MOQ")),
    })

rows.sort(key=lambda x: (x["适用动物线"], x["一级大类"], x["二级细分品类"], {"流通走量款": 0, "中端利润款": 1, "高端形象款": 2}.get(x["渠道定位"], 9), x["SKU 编码"]))

cat_counts = []
counter = Counter((r["适用动物线"], r["一级大类"], r["二级细分品类"]) for r in rows)
for (animal, main, sub), count in sorted(counter.items()):
    cat_counts.append({
        "适用动物线": animal,
        "一级大类": main,
        "二级细分品类": sub,
        "SKU 数量": count,
        "均衡判断": "SKU 扎堆，需控色/控主题" if count >= 8 else "覆盖较薄，可补品类" if count <= 2 else "结构基本合理",
    })

grouped = defaultdict(list)
for r in rows:
    if r["冗余组"]:
        grouped[r["冗余组"]].append(r)

redundancy = []
for group, items in sorted(grouped.items(), key=lambda kv: (-len(kv[1]), kv[0])):
    if len(items) < 2:
        continue
    keep = []
    merge = []
    for item in items:
        if item["渠道定位"] == "流通走量款" or any(c in item["现有渠道"] for c in ["PetcoUS", "Pet Supermarket", "Petsmart"]):
            keep.append(item["SKU 编码"])
        else:
            merge.append(item["SKU 编码"])
    redundancy.append({
        "冗余组": group,
        "涉及 SKU": ", ".join(i["SKU 编码"] for i in items),
        "建议保留": ", ".join(keep[:4]) if keep else items[0]["SKU 编码"],
        "建议精简/合并": ", ".join(merge) if merge else "保留销量验证，后续按渠道动销淘汰尾色",
        "判断口径": "同功能、同尺寸或仅颜色/主题差异，线下门店货架不宜过多占面",
    })

gaps = [
    ["兔/豚鼠大规格刚需", "大容量草架、重型陶瓷碗、耐咬饮水瓶、尿垫/除味", "现有兔用品多在厕所/饮水瓶/外出箱，日常消耗和高复购清洁耗材不足"],
    ["清洁除味耗材", "笼具清洁湿巾、除味喷雾、尿砂/尿垫组合", "线下商超适合与垫料同架联动，提高复购"],
    ["成套 starter kit", "入门套装：小笼/外出箱+水瓶+碗+垫料+躲避屋", "经销商更易按套订货，门店陈列更清晰"],
    ["安全天然咬胶扩展", "苹果木枝、草编球、无糖天然磨牙多包装", "现有咬胶偏造型，基础高频耗材不足"],
    ["季节/节庆陈列包", "春夏清凉陶瓷窝、秋冬保暖窝材、节庆限定小陈列", "线下货架需要主题更新，但应控制 SKU 数量"],
]

sets = [
    ["基础小宠门店 1 米货架", "WJ-4 / WJ-46 / T-1-BL-WH / ZM005-BL", "BOBO22-WH / R-8 / R-12 / T-41", "低价水瓶食碗和垫料引流，搭配跑轮、木屋、围栏形成客单"],
    ["垫料耗材补货包", "CFDL1090-BR / CFDL1090-PK / T-1-BL-WH", "ZM003-BL-WH / ZM1.2-PL-BL / ZM005-MIX", "按纸粒、纸绳、棉球三类铺货，减少同色重复"],
    ["形象陈列门店", "WJ-4 / WJ-46 / ZM005-BL", "HL-4 / SP-4 / WJ-39-PK / CPWJ87000", "用笼具和 XL 跑轮做视觉展示，基础耗材承担动销"],
    ["兔用品补充包", "RJ722 / RJ723 / RJ183-BL / RJ121", "RJ186-BL / RJ714-BL / CPTL0800 / RJ122-BL", "容量阶梯覆盖基础需求，大件控制样机和订单补货"],
    ["木制主题利润包", "CPWJ0014 / CPWJ181002RD / R-62", "CPWJ0011 / CPWJ0013 / R-71 / R-72", "桥梯走量、窝屋做利润，但每店控制主题数量"],
]

tree_rows = []
for animal in ["仓鼠/小型啮齿类", "兔/豚鼠类", "通用小宠耗材"]:
    animal_rows = [r for r in rows if r["适用动物线"] == animal]
    if not animal_rows:
        continue
    tree_rows.append([animal, "", "", "", "", "", ""])
    for main in sorted({r["一级大类"] for r in animal_rows}):
        tree_rows.append(["", main, "", "", "", "", ""])
        for sub in sorted({r["二级细分品类"] for r in animal_rows if r["一级大类"] == main}):
            tree_rows.append(["", "", sub, "", "", "", ""])
            for channel in ["流通走量款", "中端利润款", "高端形象款"]:
                skus = [r for r in animal_rows if r["一级大类"] == main and r["二级细分品类"] == sub and r["渠道定位"] == channel]
                if not skus:
                    continue
                tree_rows.append(["", "", "", channel, len(skus), ", ".join(r["SKU 编码"] for r in skus), "；".join(r["产品名称"] for r in skus)])

animal_summary = []
for animal in ["仓鼠/小型啮齿类", "兔/豚鼠类", "通用小宠耗材"]:
    subset = [r for r in rows if r["适用动物线"] == animal]
    if not subset:
        continue
    channel_counter = Counter(r["渠道定位"] for r in subset)
    main_counter = Counter(r["一级大类"] for r in subset)
    animal_summary.append({
        "适用动物线": animal,
        "SKU 数量": len(subset),
        "流通走量款": channel_counter.get("流通走量款", 0),
        "中端利润款": channel_counter.get("中端利润款", 0),
        "高端形象款": channel_counter.get("高端形象款", 0),
        "主要品类": "；".join(f"{k}{v}" for k, v in main_counter.most_common(4)),
        "矩阵判断": "主力专用品线，需控制造型重复" if animal == "仓鼠/小型啮齿类" else "兔用品覆盖偏基础，大件需少量陈列" if animal == "兔/豚鼠类" else "高频补货线，适合作为经销商基础盘",
    })

output = {
    "detail": rows,
    "tree": tree_rows,
    "cat_counts": cat_counts,
    "redundancy": redundancy,
    "gaps": gaps,
    "sets": sets,
    "animal_summary": animal_summary,
    "summary": {
        "total_sku": len(rows),
        "animal_counts": Counter(r["适用动物线"] for r in rows),
        "main_counts": Counter(r["一级大类"] for r in rows),
        "channel_counts": Counter(r["渠道定位"] for r in rows),
    },
}

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(OUT)
