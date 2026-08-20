# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
import time
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from config import (
    INPUT_A, INPUT_B, DATA_DIR, OUTPUT_DIR, AUTO_DISCOVER_TWO_INPUTS,
    COLUMN_ALIASES, REQUIRED_COLUMNS, CAMERA_ERROR_NAMES,
    GROUP_CHANNEL_RULES, DEFAULT_CHANNEL_RULES, CORE_SYSTEMS,
    CORE_SYSTEM_LAYER, ORDINARY_SYSTEM_LAYER, REVIEW_SYSTEM_LAYER,
    SYSTEM_WEIGHTS, PRIORITY_CONFIG, PRIORITY_ORDER, PRIORITY_ADVICE,
    REVIEW_SYSTEMS, CHANGE_CONFIG, OUTPUT_FILE_NAME, TEMP_FILE_NAME,
    TOP_FOCUS_COUNT, TOP100_COUNT,
)


def text(v: Any) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v).strip()


def safe_int(v: Any) -> int | None:
    s = text(v)
    if not s:
        return None
    try:
        return int(float(s))
    except (TypeError, ValueError):
        return channel_from_content(s)


def channel_from_content(v: Any) -> int | None:
    s = text(v)
    if not s:
        return None
    patterns = [
        r"通道号\s*[:：#]?\s*(\d+)",
        r"通道\s*[:：#]?\s*(\d+)",
        r"channel\s*[:：#_\-]?\s*(\d+)",
        r"ch(?:annel)?\s*[:：#_\-]?\s*(\d+)",
        r"\bCH\s*[-_:#：]?\s*(\d+)\b",
        r"(\d+)\s*通道\b",
    ]
    for pattern in patterns:
        m = re.search(pattern, s, flags=re.I)
        if m:
            return int(m.group(1))
    return None


def detect_report_type(df: pd.DataFrame, path: Path | None = None) -> str:
    """
    自动判断 Excel 报表类型。

    返回：
    - video_anomaly：视频异常明细报表
    - mileage_summary：车辆运行统计报表
    - unknown：无法识别
    """
    columns = {text(c) for c in df.columns}

    # 视频异常明细报表的核心字段
    anomaly_fields = {
        "状态名称",
        "状态类型",
        "设备编号",
        "归属车组",
        "状态内容",
    }

    # 车辆运行统计报表的典型字段
    mileage_fields = {
        "统计日期",
        "总运行车辆数",
        "总运行里程（KM）",
    }

    # 先判断视频异常明细报表
    if anomaly_fields.issubset(columns):
        return "video_anomaly"

    # 再判断车辆运行统计报表
    if mileage_fields.issubset(columns):
        return "mileage_summary"

    # 根据部分字段进一步判断
    anomaly_score = len(columns & anomaly_fields)
    mileage_score = len(columns & mileage_fields)

    if anomaly_score >= 3 and anomaly_score > mileage_score:
        return "video_anomaly"

    if mileage_score >= 2 and mileage_score > anomaly_score:
        return "mileage_summary"

    return "unknown"


def report_type_name(report_type: str) -> str:
    """将内部报表类型转换成用户容易理解的中文名称。"""
    names = {
        "video_anomaly": "视频异常明细报表",
        "mileage_summary": "车辆运行统计报表",
        "unknown": "未知报表",
    }
    return names.get(report_type, "未知报表")


def detect_report_type(df: pd.DataFrame, path: Path | None = None) -> str:
    """
    自动判断 Excel 报表类型。

    返回：
    - video_anomaly：视频异常明细报表
    - mileage_summary：车辆运行统计报表
    - unknown：无法识别
    """
    columns = {text(c) for c in df.columns}

    # 视频异常明细报表的核心字段
    anomaly_fields = {
        "状态名称",
        "状态类型",
        "设备编号",
        "归属车组",
        "状态内容",
    }

    # 车辆运行统计报表的典型字段
    mileage_fields = {
        "统计日期",
        "总运行车辆数",
        "总运行里程（KM）",
    }

    # 先判断视频异常明细报表
    if anomaly_fields.issubset(columns):
        return "video_anomaly"

    # 再判断车辆运行统计报表
    if mileage_fields.issubset(columns):
        return "mileage_summary"

    # 根据部分字段进一步判断
    anomaly_score = len(columns & anomaly_fields)
    mileage_score = len(columns & mileage_fields)

    if anomaly_score >= 3 and anomaly_score > mileage_score:
        return "video_anomaly"

    if mileage_score >= 2 and mileage_score > anomaly_score:
        return "mileage_summary"

    return "unknown"


def report_type_name(report_type: str) -> str:
    """将内部报表类型转换成用户容易理解的中文名称。"""
    names = {
        "video_anomaly": "视频异常明细报表",
        "mileage_summary": "车辆运行统计报表",
        "unknown": "未知报表",
    }
    return names.get(report_type, "未知报表")


def canonicalize_columns(df: pd.DataFrame, path: Path | None = None) -> pd.DataFrame:
    df = df.copy()

    # 清理 Excel 表头
    df.columns = [text(c) for c in df.columns]

    lower_to_actual = {c.lower(): c for c in df.columns}

    rename_map: dict[str, str] = {}

    # 根据 config.py 中的 COLUMN_ALIASES 自动统一字段名称
    for canonical, aliases in COLUMN_ALIASES.items():
        if canonical in df.columns:
            continue

        for alias in aliases:
            actual = lower_to_actual.get(alias.lower())

            if actual is not None:
                rename_map[actual] = canonical
                break

    if rename_map:
        df = df.rename(columns=rename_map)

    # 自动识别报表类型
    report_type = detect_report_type(df, path)

    # 车辆运行统计报表
    if report_type == "mileage_summary":
        file_name = path.name if path else "当前文件"

        raise ValueError(
            f"\n检测到：{report_type_name(report_type)}\n"
            f"文件：{file_name}\n\n"
            "当前文件包含车辆运行统计字段：\n"
            "  - 统计日期\n"
            "  - 总运行车辆数\n"
            "  - 总运行里程（KM）\n\n"
            "该文件不属于 V8 Pro 视频异常明细报表，"
            "无法用于当前的视频异常分析。\n\n"
            "请提供包含以下字段的视频异常明细报表：\n"
            "  - 状态名称\n"
            "  - 状态类型\n"
            "  - 设备编号\n"
            "  - 归属车组\n"
            "  - 状态内容"
        )

    # 完全无法识别
    if report_type == "unknown":
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]

        file_name = path.name if path else "当前文件"

        raise ValueError(
            f"\n无法识别输入报表类型。\n"
            f"文件：{file_name}\n\n"
            f"当前字段：{list(df.columns)}\n\n"
            f"缺少视频异常分析所需字段：{missing}\n\n"
            "请确认输入的是视频异常明细报表。"
        )

    # 视频异常明细报表
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]

    if missing:
        raise ValueError(
            f"\n检测到：{report_type_name(report_type)}\n"
            f"但输入报表字段仍不完整。\n"
            f"缺少：{missing}\n"
            f"当前字段：{list(df.columns)}"
        )

    # 统一字段数据类型
    for c in REQUIRED_COLUMNS:
        df[c] = df[c].map(text)

    if "通道号" in df.columns:
        df["通道号"] = df["通道号"].map(safe_int)

    return df

def discover_inputs() -> tuple[Path, Path]:
    if INPUT_A.exists() and INPUT_B.exists():
        return INPUT_A, INPUT_B
    if not AUTO_DISCOVER_TWO_INPUTS:
        raise FileNotFoundError(
            f"请把两份Excel放入 data：\n{INPUT_A}\n{INPUT_B}"
        )
    candidates = []
    for p in DATA_DIR.glob("*.xlsx"):
        name = p.name
        if name.startswith("~$"):
            continue
        if "V8" in name or name == OUTPUT_FILE_NAME or name == TEMP_FILE_NAME:
            continue
        candidates.append(p)
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    if len(candidates) < 2:
        raise FileNotFoundError(
            "data 文件夹中不足两份可用 Excel。\n"
            "推荐直接命名为 data_A.xlsx 和 data_B.xlsx。"
        )
    return candidates[1], candidates[0]


def classify(group: Any, channel: Any) -> tuple[str, str, str, str]:
    g = text(group)
    ch = safe_int(channel)
    if ch is None:
        return "未知", REVIEW_SYSTEM_LAYER, "无法提取通道号", "待确认"
    for rule in GROUP_CHANNEL_RULES:
        if rule["pattern"] in g:
            if ch in rule["channels"]:
                system = rule["channels"][ch]
                layer = CORE_SYSTEM_LAYER if system in CORE_SYSTEMS else ORDINARY_SYSTEM_LAYER
                return system, layer, f"车组专属规则：{rule['name']}", "已确认"
            return "未配置", REVIEW_SYSTEM_LAYER, f"车组“{rule['name']}”未覆盖通道{ch}", "待确认"
    if ch in DEFAULT_CHANNEL_RULES:
        system = DEFAULT_CHANNEL_RULES[ch]
        layer = CORE_SYSTEM_LAYER if system in CORE_SYSTEMS else ORDINARY_SYSTEM_LAYER
        return system, layer, "通用通道规则", "已确认"
    return "未知", REVIEW_SYSTEM_LAYER, f"规则库未覆盖通道{ch}", "待确认"


def load_one(path: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    raw = pd.read_excel(path)
    raw = canonicalize_columns(raw,path)
    camera_names = {text(x) for x in CAMERA_ERROR_NAMES}
    mask = (
        raw["状态名称"].isin(camera_names)
        | raw["状态名称"].str.contains("摄像头|视频丢失|视频异常|视频丢帧|视频中断", case=False, na=False)
        | raw["状态内容"].str.contains("摄像头异常|视频丢失|视频异常|视频丢帧|视频中断", case=False, na=False)
    )
    camera = raw.loc[mask].copy()
    if camera.empty:
        detail = pd.DataFrame(columns=["设备编号","归属车组","状态名称","状态类型","通道号","系统类型","系统层级","分类依据","规则状态","状态内容"])
        stats = empty_stats()
        quality = quality_frame(len(raw), 0, stats)
        return detail, stats, quality

    if "通道号" in camera.columns:
        camera["通道号"] = camera["通道号"].map(safe_int)
        extracted = camera["状态内容"].map(channel_from_content)
        camera["通道号"] = camera["通道号"].where(camera["通道号"].notna(), extracted)
    else:
        camera["通道号"] = camera["状态内容"].map(channel_from_content)

    classified = camera.apply(lambda r: classify(r["归属车组"], r["通道号"]), axis=1, result_type="expand")
    classified.columns = ["系统类型", "系统层级", "分类依据", "规则状态"]
    camera[["系统类型","系统层级","分类依据","规则状态"]] = classified

    detail_cols = ["设备编号","归属车组","状态名称","状态类型","通道号","系统类型","系统层级","分类依据","规则状态","状态内容"]
    detail = camera[detail_cols].copy()
    stats = (
        camera.groupby(["归属车组","通道号","系统类型","系统层级","分类依据","规则状态"], dropna=False)
        .size().reset_index(name="视频丢失次数")
    )
    quality = quality_frame(len(raw), len(detail), stats)
    return detail, stats, quality


def empty_stats() -> pd.DataFrame:
    return pd.DataFrame(columns=["归属车组","通道号","系统类型","系统层级","分类依据","规则状态","视频丢失次数"])


def pct_rank(s: pd.Series) -> pd.Series:
    if len(s) <= 1:
        return pd.Series(1.0, index=s.index, dtype=float)
    return s.rank(method="average", pct=True)


def normalize(s: pd.Series) -> pd.Series:
    s = pd.to_numeric(s, errors="coerce").fillna(0.0)
    if len(s) == 0:
        return s.astype(float)
    lo, hi = float(s.min()), float(s.max())
    if hi <= lo:
        return pd.Series(0.5, index=s.index, dtype=float)
    return ((s - lo) / (hi - lo)).clip(0, 1)


def dynamic_cut(n: int, share: float) -> int:
    return 0 if n <= 0 else max(1, int(np.ceil(n * share)))


def validate_priority_config() -> None:
    c = PRIORITY_CONFIG
    shares = [float(c[k]) for k in ("P1_MAX_SHARE","P2_MAX_SHARE","P3_MAX_SHARE")]
    if not all(0 < x <= 1 for x in shares) or not (shares[0] <= shares[1] <= shares[2]):
        raise ValueError("P1/P2/P3比例配置错误，必须满足 0<P1<=P2<=P3<=1。")
    weights = [k for k in c if k.startswith("WEIGHT_")]
    if sum(float(c[k]) for k in weights) <= 0:
        raise ValueError("优先级权重不能全部为0。")


def score_priority(stats: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    validate_priority_config()
    df = stats.copy()
    if df.empty:
        cols = ["车组异常总次数","系统异常总次数","车组异常通道数","系统权重","系统重要性分","次数排名分","绝对次数分","车组影响分","系统影响分","扩散分","核心系统加分","动态风险分","最终优先级","规则判断","处理建议","异常等级"]
        for col in cols:
            df[col] = pd.Series(dtype=float if col.endswith("分") or col.endswith("次数") else object)
        summary = pd.DataFrame({"最终优先级":["P1","P2","P3","P4"],"异常通道数":[0,0,0,0]})
        return df, summary
    df["视频丢失次数"] = pd.to_numeric(df["视频丢失次数"], errors="coerce").fillna(0).clip(lower=0).astype(int)
    group_total = df.groupby("归属车组", dropna=False)["视频丢失次数"].sum().rename("车组异常总次数")
    system_total = df.groupby("系统类型", dropna=False)["视频丢失次数"].sum().rename("系统异常总次数")
    spread = df.groupby("归属车组", dropna=False)["通道号"].nunique(dropna=True).rename("车组异常通道数")
    df = df.join(group_total, on="归属车组").join(system_total, on="系统类型").join(spread, on="归属车组")
    df["系统权重"] = df["系统类型"].map(SYSTEM_WEIGHTS).fillna(0.10)
    max_w = max(SYSTEM_WEIGHTS.values()) if SYSTEM_WEIGHTS else 1.0
    df["系统重要性分"] = (df["系统权重"] / max_w).clip(0,1)
    df["次数排名分"] = pct_rank(df["视频丢失次数"])
    df["绝对次数分"] = normalize(df["视频丢失次数"])
    df["车组影响分"] = normalize(df["车组异常总次数"])
    df["系统影响分"] = normalize(df["系统异常总次数"])
    df["扩散分"] = normalize(df["车组异常通道数"])
    df["核心系统加分"] = df["系统类型"].isin(CORE_SYSTEMS).astype(float)
    c = PRIORITY_CONFIG
    weights = {
        "次数排名分": c["WEIGHT_COUNT_RANK"], "绝对次数分": c["WEIGHT_COUNT_ABSOLUTE"],
        "车组影响分": c["WEIGHT_GROUP_IMPACT"], "系统影响分": c["WEIGHT_SYSTEM_IMPACT"],
        "扩散分": c["WEIGHT_SPREAD"], "系统重要性分": c["WEIGHT_SYSTEM_IMPORTANCE"], "核心系统加分": c["WEIGHT_CORE_SYSTEM"],
    }
    total_w = sum(float(x) for x in weights.values())
    df["动态风险分"] = (sum(df[k] * float(v) for k,v in weights.items()) / total_w * 100).round(2)
    n = len(df)
    rank = df["动态风险分"].rank(method="first", ascending=False)
    p1_cut = dynamic_cut(n, c["P1_MAX_SHARE"]); p2_cut = max(p1_cut, dynamic_cut(n,c["P2_MAX_SHARE"])); p3_cut=max(p2_cut,dynamic_cut(n,c["P3_MAX_SHARE"]))
    df["最终优先级"] = "P4"
    confirmed = ~df["系统类型"].isin(REVIEW_SYSTEMS)
    p1_min=min(int(c["P1_MIN_COUNT"]), max(1,int(np.ceil(n*0.02))))
    p2_min=min(int(c["P2_MIN_COUNT"]), max(1,int(np.ceil(n*0.01))))
    df.loc[confirmed & (rank<=p1_cut) & (df["视频丢失次数"]>=p1_min),"最终优先级"]="P1"
    df.loc[confirmed & (rank<=p2_cut) & (df["视频丢失次数"]>=p2_min) & (df["最终优先级"]!="P1"),"最终优先级"]="P2"
    df.loc[confirmed & (rank<=p3_cut) & (df["视频丢失次数"]>=1) & df["最终优先级"].isin(["P3","P4"]),"最终优先级"]="P3"
    df["规则判断"] = np.where(df["系统类型"].isin(REVIEW_SYSTEMS), "待确认：当前规则库无法可靠映射", "已按当前规则库分类")
    df["处理建议"] = df["最终优先级"].map(PRIORITY_ADVICE).fillna("常规跟踪")
    df.loc[df["系统类型"].isin(REVIEW_SYSTEMS),"处理建议"]="先确认车组/通道规则，再判断是否为真实异常"
    df["异常等级"] = pd.cut(df["视频丢失次数"],[-1,0,4,9,19,np.inf],labels=["无","轻微","一般","较严重","严重"]).astype(str)
    df["_p"] = df["最终优先级"].map(PRIORITY_ORDER).fillna(99)
    # 核心系统层级先于普通摄像头层级；同层再看风险分。
    df["_layer"] = df["系统层级"].map({CORE_SYSTEM_LAYER:1, ORDINARY_SYSTEM_LAYER:2, REVIEW_SYSTEM_LAYER:3}).fillna(4)
    df = df.sort_values(["_layer","_p","动态风险分","视频丢失次数","归属车组","通道号"],ascending=[True,True,False,False,True,True],na_position="last").drop(columns=["_p","_layer"])
    summary = df["最终优先级"].value_counts().reindex(["P1","P2","P3","P4"],fill_value=0).rename_axis("最终优先级").reset_index(name="异常通道数")
    return df, summary


def priority_layer_summary(priority: pd.DataFrame) -> pd.DataFrame:
    if priority.empty:
        return pd.DataFrame({"系统层级":[CORE_SYSTEM_LAYER,ORDINARY_SYSTEM_LAYER,REVIEW_SYSTEM_LAYER],"P1":[0,0,0],"P2":[0,0,0],"P3":[0,0,0],"P4":[0,0,0]})
    x = pd.crosstab(priority["系统层级"], priority["最终优先级"]).reindex(columns=["P1","P2","P3","P4"],fill_value=0).reset_index()
    return x


def quality_frame(raw_count: int, detail_count: int, stats: pd.DataFrame) -> pd.DataFrame:
    p = score_priority(stats)[1]
    cmap=dict(zip(p["最终优先级"],p["异常通道数"]))
    return pd.DataFrame({"检查项目":["原始记录数","识别出的摄像头/视频异常记录数","统计后的异常车组-通道数","待确认规则车组-通道数","P1通道数","P2通道数","P3通道数","P4通道数"],"结果":[raw_count,detail_count,len(stats),int((stats["规则状态"]=="待确认").sum()) if len(stats) else 0,cmap.get("P1",0),cmap.get("P2",0),cmap.get("P3",0),cmap.get("P4",0)]})


def system_summary(stats: pd.DataFrame) -> pd.DataFrame:
    if stats.empty:
        return pd.DataFrame(columns=["系统层级","系统类型","视频丢失次数"])
    return stats.groupby(["系统层级","系统类型"],dropna=False)["视频丢失次数"].sum().reset_index().sort_values(["系统层级","视频丢失次数"],ascending=[True,False])


def group_summary(stats: pd.DataFrame) -> pd.DataFrame:
    if stats.empty:
        return pd.DataFrame(columns=["归属车组","系统层级","视频丢失次数"])
    return stats.groupby(["归属车组","系统层级"],dropna=False)["视频丢失次数"].sum().reset_index().sort_values("视频丢失次数",ascending=False)


def core_summary(stats: pd.DataFrame) -> pd.DataFrame:
    if stats.empty:
        return pd.DataFrame(columns=["系统类型","归属车组","视频丢失次数"])
    return stats[stats["系统类型"].isin(CORE_SYSTEMS)].groupby(["系统类型","归属车组"],dropna=False)["视频丢失次数"].sum().reset_index().sort_values("视频丢失次数",ascending=False)


def review_summary(stats: pd.DataFrame) -> pd.DataFrame:
    if stats.empty:
        return stats.copy()
    return stats[stats["规则状态"]=="待确认"].sort_values("视频丢失次数",ascending=False).copy()


def rule_coverage(stats: pd.DataFrame) -> pd.DataFrame:
    if stats.empty:
        return pd.DataFrame(columns=["系统层级","系统类型","分类依据","规则状态","视频丢失次数","车组数","通道数"])
    return stats.groupby(["系统层级","系统类型","分类依据","规则状态"],dropna=False).agg(视频丢失次数=("视频丢失次数","sum"),车组数=("归属车组","nunique"),通道数=("通道号","nunique")).reset_index().sort_values("视频丢失次数",ascending=False)


def analysis_bundle(path: Path) -> dict[str, Any]:
    detail, stats, _ = load_one(path)
    priority, psummary = score_priority(stats)
    return {
        "path": path, "raw_count": None, "detail": detail, "stats": stats,
        "priority": priority, "priority_summary": psummary,
        "system": system_summary(stats), "group": group_summary(stats), "core": core_summary(stats),
        "focus": priority.head(TOP_FOCUS_COUNT).copy(), "top100": priority.head(TOP100_COUNT).copy(),
        "layer_summary": priority_layer_summary(priority), "review": review_summary(stats),
        "rule_check": rule_coverage(stats),
    }


def build_key(df: pd.DataFrame) -> pd.Series:
    if df.empty:
        return pd.Series(index=df.index,dtype=object)
    return df["归属车组"].map(text)+"||"+df["通道号"].map(lambda x:"" if pd.isna(x) else str(int(x)))


def compare_long(a: pd.DataFrame, b: pd.DataFrame, metric: str, extra_cols: list[str]) -> pd.DataFrame:
    # 对比的主键统一为“车组 + 通道”，系统类型变化也会被展示。
    aa=a.copy(); bb=b.copy()
    for df in (aa,bb):
        if not df.empty: df["对比键"]=build_key(df)
    base_cols=["归属车组","通道号","系统类型","系统层级","视频丢失次数"]+extra_cols
    aa=aa.reindex(columns=[c for c in base_cols if c in aa.columns]+(["对比键"] if "对比键" in aa.columns else []))
    bb=bb.reindex(columns=[c for c in base_cols if c in bb.columns]+(["对比键"] if "对比键" in bb.columns else []))
    if aa.empty and bb.empty:
        return pd.DataFrame(columns=["归属车组","通道号","A系统类型","B系统类型","A系统层级","B系统层级","A次数","B次数","变化次数","变化率","变化状态"])
    if aa.empty:
        aa=pd.DataFrame(columns=bb.columns)
    if bb.empty:
        bb=pd.DataFrame(columns=aa.columns)
    aa=aa.rename(columns={"系统类型":"A系统类型","系统层级":"A系统层级","视频丢失次数":"A次数"})
    bb=bb.rename(columns={"系统类型":"B系统类型","系统层级":"B系统层级","视频丢失次数":"B次数"})
    # 如果同一车组+通道出现多条系统分类，按次数汇总，避免 merge 重复。
    def prep(x):
        if x.empty: return x
        x["A次数" if "A次数" in x.columns else "B次数"]=pd.to_numeric(x["A次数" if "A次数" in x.columns else "B次数"],errors="coerce").fillna(0)
        return x.groupby("对比键",as_index=False).agg({"归属车组":"first","通道号":"first", **({"A系统类型":"first","A系统层级":"first","A次数":"sum"} if "A次数" in x.columns else {"B系统类型":"first","B系统层级":"first","B次数":"sum"})})
    aa=prep(aa); bb=prep(bb)
    m=aa.merge(bb,on="对比键",how="outer",suffixes=("","_B"))
    for c in ["归属车组","通道号"]:
        m[c]=m[c].where(m[c].notna(),m.get(c+"_B"))
    for c in ["A次数","B次数"]:
        if c not in m.columns: m[c]=0
        m[c]=pd.to_numeric(m[c],errors="coerce").fillna(0)
    m["变化次数"]=(m["B次数"]-m["A次数"]).astype(int)
    m["变化率"]=np.where(m["A次数"]>0,m["变化次数"]/m["A次数"],np.where(m["B次数"]>0,np.inf,0))
    sig_n=int(CHANGE_CONFIG["SIGNIFICANT_COUNT_CHANGE"]); sig_p=float(CHANGE_CONFIG["SIGNIFICANT_PERCENT_CHANGE"])
    def status(r):
        a,b=r["A次数"],r["B次数"]
        if a==0 and b>0:return "新增异常"
        if a>0 and b==0:return "异常消失"
        d=b-a
        if d==0:return "持续稳定"
        if d>0 and (d>=sig_n or (a>0 and d/a>=sig_p)):return "明显恶化"
        if d<0 and ((-d)>=sig_n or (a>0 and (-d)/a>=sig_p)):return "明显改善"
        return "小幅变化"
    m["变化状态"]=m.apply(status,axis=1)
    if "A系统类型" not in m.columns:m["A系统类型"]=""
    if "B系统类型" not in m.columns:m["B系统类型"]=""
    if "A系统层级" not in m.columns:m["A系统层级"]=""
    if "B系统层级" not in m.columns:m["B系统层级"]=""
    # 系统层级变化：核心系统必须优先于普通摄像头；若分类发生变化单独标记。
    def layer_status(r):
        a = "" if pd.isna(r["A系统层级"]) else text(r["A系统层级"])
        b = "" if pd.isna(r["B系统层级"]) else text(r["B系统层级"])
        if not a and b:
            return "新增分类"
        if a and not b:
            return "原有分类消失"
        if a and b and a != b:
            return "系统层级发生变化"
        return "无"
    m["分类变化"]=m.apply(layer_status,axis=1)
    cols=["归属车组","通道号","A系统类型","B系统类型","A系统层级","B系统层级","A次数","B次数","变化次数","变化率","变化状态","分类变化"]
    return m[cols].sort_values(["变化状态","变化次数"],ascending=[True,False],na_position="last").reset_index(drop=True)


def compare_system(a: pd.DataFrame,b: pd.DataFrame)->pd.DataFrame:
    x=a.groupby(["系统层级","系统类型"],dropna=False)["视频丢失次数"].sum().rename("A次数") if not a.empty else pd.Series(dtype=float)
    y=b.groupby(["系统层级","系统类型"],dropna=False)["视频丢失次数"].sum().rename("B次数") if not b.empty else pd.Series(dtype=float)
    m=pd.concat([x,y],axis=1).fillna(0).reset_index()
    m["变化次数"]=(m["B次数"]-m["A次数"]).astype(int); m["变化率"]=np.where(m["A次数"]>0,m["变化次数"]/m["A次数"],np.where(m["B次数"]>0,np.inf,0))
    return m.sort_values("变化次数",ascending=False)


def compare_groups(a: pd.DataFrame,b: pd.DataFrame)->pd.DataFrame:
    x=a.groupby(["归属车组","系统层级"],dropna=False)["视频丢失次数"].sum().rename("A次数") if not a.empty else pd.Series(dtype=float)
    y=b.groupby(["归属车组","系统层级"],dropna=False)["视频丢失次数"].sum().rename("B次数") if not b.empty else pd.Series(dtype=float)
    m=pd.concat([x,y],axis=1).fillna(0).reset_index(); m["变化次数"]=(m["B次数"]-m["A次数"]).astype(int); m["变化率"]=np.where(m["A次数"]>0,m["变化次数"]/m["A次数"],np.where(m["B次数"]>0,np.inf,0)); return m.sort_values("变化次数",ascending=False)


def compare_core(a: pd.DataFrame,b: pd.DataFrame)->pd.DataFrame:
    return compare_groups(a[a["系统类型"].isin(CORE_SYSTEMS)] if not a.empty else a,b[b["系统类型"].isin(CORE_SYSTEMS)] if not b.empty else b).rename(columns={"系统层级":"系统层级"})


def compare_priority(a: pd.DataFrame,b: pd.DataFrame)->pd.DataFrame:
    aa=a[["归属车组","通道号","系统类型","系统层级","视频丢失次数","动态风险分","最终优先级"]].copy() if not a.empty else pd.DataFrame(columns=["归属车组","通道号","系统类型","系统层级","视频丢失次数","动态风险分","最终优先级"])
    bb=b[["归属车组","通道号","系统类型","系统层级","视频丢失次数","动态风险分","最终优先级"]].copy() if not b.empty else aa.iloc[0:0].copy()
    aa["对比键"]=build_key(aa); bb["对比键"]=build_key(bb)
    aa=aa.rename(columns={"系统类型":"A系统类型","系统层级":"A系统层级","视频丢失次数":"A次数","动态风险分":"A风险分","最终优先级":"A优先级"})
    bb=bb.rename(columns={"系统类型":"B系统类型","系统层级":"B系统层级","视频丢失次数":"B次数","动态风险分":"B风险分","最终优先级":"B优先级"})
    aa=aa.groupby("对比键",as_index=False).agg({"归属车组":"first","通道号":"first","A系统类型":"first","A系统层级":"first","A次数":"sum","A风险分":"max","A优先级":"first"})
    bb=bb.groupby("对比键",as_index=False).agg({"归属车组":"first","通道号":"first","B系统类型":"first","B系统层级":"first","B次数":"sum","B风险分":"max","B优先级":"first"})
    m=aa.merge(bb,on="对比键",how="outer",suffixes=("","_b"))
    m["归属车组"]=m["归属车组"].fillna(m.get("归属车组_b")); m["通道号"]=m["通道号"].fillna(m.get("通道号_b"))
    for c in ["A次数","B次数","A风险分","B风险分"]: m[c]=pd.to_numeric(m.get(c,0),errors="coerce").fillna(0)
    m["次数变化"]=m["B次数"]-m["A次数"]; m["风险分变化"]=(m["B风险分"]-m["A风险分"]).round(2)
    m["优先级变化"]=m["A优先级"].fillna("无")+" → "+m["B优先级"].fillna("无")
    def level(r):
        if r["A系统层级"]!=r["B系统层级"] and pd.notna(r["A系统层级"]) and pd.notna(r["B系统层级"]): return "系统层级变化"
        if pd.isna(r["A次数"]) or r["A次数"]==0: return "新增异常"
        if pd.isna(r["B次数"]) or r["B次数"]==0: return "异常消失"
        if r["B次数"]>r["A次数"]: return "风险上升"
        if r["B次数"]<r["A次数"]: return "风险下降"
        return "无变化"
    m["变化状态"]=m.apply(level,axis=1)
    return m[["归属车组","通道号","A系统类型","B系统类型","A系统层级","B系统层级","A次数","B次数","次数变化","A风险分","B风险分","风险分变化","A优先级","B优先级","优先级变化","变化状态"]].sort_values(["变化状态","风险分变化"],ascending=[True,False])


def comparison_overview(a: dict[str,Any], b: dict[str,Any], item: pd.DataFrame) -> pd.DataFrame:
    def total(d): return int(d["stats"]["视频丢失次数"].sum()) if not d["stats"].empty else 0
    def core(d): return int(d["stats"].loc[d["stats"]["系统类型"].isin(CORE_SYSTEMS),"视频丢失次数"].sum()) if not d["stats"].empty else 0
    vals=[
        ("原始记录数",None,None),
        ("视频/摄像头异常记录数",len(a["detail"]),len(b["detail"])),
        ("视频丢失次数",total(a),total(b)),
        ("核心系统异常次数",core(a),core(b)),
        ("异常车组-通道数",len(a["stats"]),len(b["stats"])),
        ("核心系统异常车组-通道数",int((a["stats"]["系统类型"].isin(CORE_SYSTEMS)).sum()) if not a["stats"].empty else 0,int((b["stats"]["系统类型"].isin(CORE_SYSTEMS)).sum()) if not b["stats"].empty else 0),
        ("待确认规则数",int((a["stats"]["规则状态"]=="待确认").sum()) if not a["stats"].empty else 0,int((b["stats"]["规则状态"]=="待确认").sum()) if not b["stats"].empty else 0),
    ]
    rows=[]
    for name,x,y in vals:
        if x is None: continue
        rows.append({"指标":name,"A值":x,"B值":y,"变化":y-x,"变化率":(y-x)/x if x else (np.inf if y else 0)})
    return pd.DataFrame(rows)


def format_excel(path: Path) -> None:
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    wb=load_workbook(path)
    if not wb.worksheets: wb.create_sheet("对比总览")
    for ws in wb.worksheets:
        ws.sheet_state="visible" if ws.sheet_state!="visible" else ws.sheet_state
        ws.freeze_panes="A2"
        if ws.max_row>=1 and ws.max_column>=1: ws.auto_filter.ref=ws.dimensions
        for c in ws[1]: c.font=Font(bold=True)
        for col in ws.iter_cols():
            letter=col[0].column_letter; ml=max([len(str(c.value)) if c.value is not None else 0 for c in list(col)[:200]]+[10]); ws.column_dimensions[letter].width=min(42,max(10,ml+2))
    wb.active=0; wb.save(path); wb.close()


def write_output(a: dict[str,Any], b: dict[str,Any]) -> Path:
    OUTPUT_DIR.mkdir(parents=True,exist_ok=True)
    out=OUTPUT_DIR/OUTPUT_FILE_NAME; tmp=OUTPUT_DIR/TEMP_FILE_NAME
    if tmp.exists():
        try: tmp.unlink()
        except OSError: pass
    ca=compare_long(a["stats"],b["stats"],"视频丢失次数",[])
    cp=compare_priority(a["priority"],b["priority"])
    sheets=[
        ("对比总览",comparison_overview(a,b,ca)),
        ("对比_总体系统",compare_system(a["stats"],b["stats"])),
        ("对比_车组异常",compare_groups(a["stats"],b["stats"])),
        ("对比_DMS_ADAS_DSC",compare_core(a["stats"],b["stats"])),
        ("对比_车组通道变化",ca),
        ("对比_优先级变化",cp),
        ("对比_新增异常",ca[ca["变化状态"]=="新增异常"]),
        ("对比_持续恶化",ca[ca["变化状态"]=="明显恶化"]),
        ("对比_明显改善",ca[ca["变化状态"]=="明显改善"]),
        ("对比_异常消失",ca[ca["变化状态"]=="异常消失"]),
        ("A_总体系统统计",a["system"]),
        ("A_车组异常统计",a["group"]),
        ("A_DMS_ADAS_DSC",a["core"]),
        ("A_最高优先级重点",a["focus"]),
        ("A_TOP100",a["top100"]),
        ("A_完整动态优先级",a["priority"]),
        ("A_优先级层级统计",a["layer_summary"]),
        ("A_待确认规则",a["review"]),
        ("A_规则覆盖检查",a["rule_check"]),
        ("A_摄像头异常明细",a["detail"]),
        ("B_总体系统统计",b["system"]),
        ("B_车组异常统计",b["group"]),
        ("B_DMS_ADAS_DSC",b["core"]),
        ("B_最高优先级重点",b["focus"]),
        ("B_TOP100",b["top100"]),
        ("B_完整动态优先级",b["priority"]),
        ("B_优先级层级统计",b["layer_summary"]),
        ("B_待确认规则",b["review"]),
        ("B_规则覆盖检查",b["rule_check"]),
        ("B_摄像头异常明细",b["detail"]),
    ]
    with pd.ExcelWriter(tmp,engine="openpyxl") as writer:
        for name,df in sheets: df.to_excel(writer,sheet_name=name,index=False)
    format_excel(tmp)
    try: os.replace(tmp,out); return out
    except PermissionError:
        stamp=time.strftime("%Y%m%d_%H%M%S"); fallback=OUTPUT_DIR/f"视频异常巡检V8_双报表分析与对比_{stamp}.xlsx"; os.replace(tmp,fallback); return fallback


def main() -> None:
    print("="*78); print("视频异常巡检 V8 - 双报表独立分析 + 自动对比"); print("="*78)
    pa,pb=discover_inputs(); print(f"A文件：{pa}"); print(f"B文件：{pb}")
    a=analysis_bundle(pa); b=analysis_bundle(pb)
    print(f"A：异常明细 {len(a['detail'])}，车组-通道 {len(a['stats'])}")
    print(f"B：异常明细 {len(b['detail'])}，车组-通道 {len(b['stats'])}")
    out=write_output(a,b)
    cmp=compare_long(a["stats"],b["stats"],"视频丢失次数",[])
    print("\n===== 对比变化统计 =====")
    if cmp.empty: print("没有可比较的异常对象。")
    else: print(cmp["变化状态"].value_counts().to_string())
    print(f"\n结果文件：{out}"); print("="*78)


if __name__=="__main__":
    try: main()
    except Exception as exc:
        print("\n程序运行失败。")
        print(f"错误类型：{type(exc).__name__}")
        print(f"错误信息：{exc}")
        raise
