from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from config import (
    INPUT_FILE, OUTPUT_DIR, CAMERA_ERROR_NAMES, REQUIRED_COLUMNS,
    GROUP_CHANNEL_RULES, DEFAULT_CHANNEL_RULES, CORE_SYSTEMS,
    SYSTEM_WEIGHTS, PRIORITY_CONFIG, PRIORITY_ORDER, REVIEW_SYSTEMS,
)


def text(v: Any) -> str:
    return "" if pd.isna(v) else str(v).strip()


def channel_from_content(v: Any) -> int | None:
    s = text(v)
    patterns = [
        r"通道号\s*[:：]?\s*(\d+)",
        r"通道\s*[:：]?\s*(\d+)",
        r"channel\s*[:：]?\s*(\d+)",
    ]
    for p in patterns:
        m = re.search(p, s, flags=re.I)
        if m:
            return int(m.group(1))
    return None


def classify(group: Any, channel: Any) -> tuple[str, str, str]:
    """返回 系统类型、分类依据、规则状态。"""
    g = text(group)
    try:
        ch = None if pd.isna(channel) else int(float(channel))
    except (TypeError, ValueError):
        ch = None

    if ch is None:
        return "未知", "无法提取通道号", "待确认"

    for rule in GROUP_CHANNEL_RULES:
        if rule["pattern"] in g:
            if ch in rule["channels"]:
                return rule["channels"][ch], f"车组专属规则：{rule['name']}", "已确认"
            return "未配置", f"车组“{rule['name']}”未覆盖通道{ch}", "待确认"

    if ch in DEFAULT_CHANNEL_RULES:
        return DEFAULT_CHANNEL_RULES[ch], "通用通道规则", "已确认"

    return "未知", f"规则库未覆盖通道{ch}", "待确认"


def load_data() -> pd.DataFrame:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"找不到输入文件：{INPUT_FILE}\n"
            "请把 Excel 放入 data 文件夹并命名为 test_data.xlsx。"
        )
    df = pd.read_excel(INPUT_FILE)
    df.columns = [text(c) for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"缺少必要字段：{missing}\n当前字段：{list(df.columns)}"
        )
    for c in REQUIRED_COLUMNS:
        df[c] = df[c].map(text)
    return df


def build_stats(raw: pd.DataFrame):
    camera = raw[raw["状态名称"].isin(CAMERA_ERROR_NAMES)].copy()
    camera["通道号"] = camera["状态内容"].map(channel_from_content)

    classified = camera.apply(
        lambda r: classify(r["归属车组"], r["通道号"]),
        axis=1, result_type="expand"
    )
    classified.columns = ["系统类型", "分类依据", "规则状态"]
    camera[["系统类型", "分类依据", "规则状态"]] = classified

    detail_cols = [
        "设备编号", "归属车组", "状态名称", "状态类型",
        "通道号", "系统类型", "分类依据", "规则状态", "状态内容"
    ]
    detail = camera[detail_cols].copy()

    stats = (
        camera.groupby(
            ["归属车组", "通道号", "系统类型", "分类依据", "规则状态"],
            dropna=False
        ).size().reset_index(name="视频丢失次数")
    )
    return detail, stats


def pct_rank(s: pd.Series) -> pd.Series:
    if len(s) <= 1:
        return pd.Series(1.0, index=s.index)
    return s.rank(method="average", pct=True)


def normalize(s: pd.Series) -> pd.Series:
    s = pd.to_numeric(s, errors="coerce").fillna(0)
    lo, hi = float(s.min()), float(s.max())
    if hi <= lo:
        return pd.Series(0.5, index=s.index)
    return (s - lo) / (hi - lo)


def score_and_priority(stats: pd.DataFrame):
    df = stats.copy()
    df["视频丢失次数"] = pd.to_numeric(df["视频丢失次数"], errors="coerce").fillna(0).astype(int)

    group = df.groupby("归属车组", dropna=False)["视频丢失次数"].sum().rename("车组异常总次数")
    sys = df.groupby("系统类型", dropna=False)["视频丢失次数"].sum().rename("系统异常总次数")
    spread = df.groupby("归属车组", dropna=False)["通道号"].nunique().rename("车组异常通道数")

    df = df.join(group, on="归属车组").join(sys, on="系统类型").join(spread, on="归属车组")

    df["系统权重"] = df["系统类型"].map(SYSTEM_WEIGHTS).fillna(0.10)
    max_w = max(SYSTEM_WEIGHTS.values())
    df["系统重要性分"] = (df["系统权重"] / max_w).clip(0, 1)

    df["次数排名分"] = pct_rank(df["视频丢失次数"])
    df["车组影响分"] = normalize(df["车组异常总次数"])
    df["系统影响分"] = normalize(df["系统异常总次数"])
    df["扩散分"] = normalize(df["车组异常通道数"])
    df["核心系统加分"] = df["系统类型"].isin(CORE_SYSTEMS).astype(float)

    # V4 风险分：绝对次数 + 当前报表排名 + 车组/系统影响 + 核心系统。
    df["动态风险分"] = (
        df["次数排名分"] * 0.40
        + normalize(df["视频丢失次数"]) * 0.20
        + df["车组影响分"] * 0.15
        + df["系统影响分"] * 0.10
        + df["扩散分"] * 0.05
        + df["系统重要性分"] * 0.07
        + df["核心系统加分"] * 0.03
    ) * 100
    df["动态风险分"] = df["动态风险分"].round(2)

    n = len(df)
    if n == 0:
        return df, pd.DataFrame(columns=["最终优先级", "异常通道数"])

    rank = df["动态风险分"].rank(method="first", ascending=False)
    p1_cut = max(1, int(np.ceil(n * PRIORITY_CONFIG["P1_MAX_SHARE"])))
    p2_cut = max(p1_cut, int(np.ceil(n * PRIORITY_CONFIG["P2_MAX_SHARE"])))
    p3_cut = max(p2_cut, int(np.ceil(n * PRIORITY_CONFIG["P3_MAX_SHARE"])))

    # 最低次数门槛；规则未确认的数据即使风险高也进入“待确认”，不自动当真实故障P1。
    df["最终优先级"] = "P4"
    confirmed = ~df["系统类型"].isin(REVIEW_SYSTEMS)

    df.loc[
        confirmed & (rank <= p1_cut) &
        (df["视频丢失次数"] >= PRIORITY_CONFIG["P1_MIN_COUNT"]),
        "最终优先级"
    ] = "P1"

    df.loc[
        confirmed & (rank <= p2_cut) &
        (df["视频丢失次数"] >= PRIORITY_CONFIG["P2_MIN_COUNT"]) &
        (df["最终优先级"] != "P1"),
        "最终优先级"
    ] = "P2"

    df.loc[
        confirmed & (rank <= p3_cut) &
        (df["视频丢失次数"] >= PRIORITY_CONFIG["P3_MIN_COUNT"]) &
        (df["最终优先级"].isin(["P3", "P4"])),
        "最终优先级"
    ] = "P3"

    df["处理建议"] = np.select(
        [
            df["系统类型"].isin(REVIEW_SYSTEMS),
            df["最终优先级"].eq("P1"),
            df["最终优先级"].eq("P2"),
            df["最终优先级"].eq("P3"),
        ],
        [
            "先确认车组/通道规则，再判断是否为真实异常",
            "立即重点处理",
            "优先处理",
            "持续关注",
        ],
        default="常规跟踪",
    )

    df["异常等级"] = pd.cut(
        df["视频丢失次数"],
        [-1, 0, 4, 9, 19, np.inf],
        labels=["无", "轻微", "一般", "较严重", "严重"]
    ).astype(str)

    df["优先级排序"] = df["最终优先级"].map(PRIORITY_ORDER).fillna(99)
    df = df.sort_values(
        ["优先级排序", "动态风险分", "视频丢失次数"],
        ascending=[True, False, False]
    ).drop(columns=["优先级排序"])

    summary = (
        df["最终优先级"].value_counts()
        .reindex(["P1", "P2", "P3", "P4"], fill_value=0)
        .rename_axis("最终优先级").reset_index(name="异常通道数")
    )
    return df, summary


def build_output(detail, stats, priority, priority_summary):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output = OUTPUT_DIR / "视频异常巡检最终汇总_V4.xlsx"

    system_summary = (
        stats.groupby("系统类型")["视频丢失次数"].sum()
        .reset_index().sort_values("视频丢失次数", ascending=False)
    )
    group_summary = (
        stats.groupby("归属车组")["视频丢失次数"].sum()
        .reset_index().sort_values("视频丢失次数", ascending=False)
    )
    core_summary = (
        stats[stats["系统类型"].isin(CORE_SYSTEMS)]
        .groupby(["系统类型", "归属车组"])["视频丢失次数"].sum()
        .reset_index().sort_values("视频丢失次数", ascending=False)
    )

    review = stats[stats["规则状态"] == "待确认"].copy()
    if len(review):
        review = review.sort_values("视频丢失次数", ascending=False)

    top100 = priority.head(100).copy()

    rule_check = (
        stats.groupby(["系统类型", "分类依据", "规则状态"], dropna=False)
        .agg(
            视频丢失次数=("视频丢失次数", "sum"),
            车组数=("归属车组", "nunique"),
            通道数=("通道号", "nunique"),
        ).reset_index().sort_values("视频丢失次数", ascending=False)
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        system_summary.to_excel(writer, "总体系统统计", index=False)
        group_summary.to_excel(writer, "车组异常统计", index=False)
        core_summary.to_excel(writer, "DMS_ADAS_DSC统计", index=False)
        top100.to_excel(writer, "动态优先级TOP100", index=False)
        priority.to_excel(writer, "完整动态优先级", index=False)
        priority_summary.to_excel(writer, "优先级统计", index=False)
        review.to_excel(writer, "待确认规则", index=False)
        rule_check.to_excel(writer, "规则覆盖检查", index=False)
        detail.to_excel(writer, "摄像头异常明细", index=False)
    return output


def main():
    print("=" * 72)
    print("视频异常巡检 - 动态优先级 V4")
    print("=" * 72)

    raw = load_data()
    print(f"输入文件：{INPUT_FILE}")
    print(f"原始数据量：{len(raw)}")

    detail, stats = build_stats(raw)
    priority, summary = score_and_priority(stats)
    output = build_output(detail, stats, priority, summary)

    print("\n===== 优先级统计 =====")
    print(summary.to_string(index=False))

    print("\n===== 待确认规则 =====")
    review = stats[stats["规则状态"] == "待确认"]
    print(f"待确认车组/通道：{len(review)}")

    print("\n===== TOP20 =====")
    cols = ["归属车组", "通道号", "视频丢失次数", "系统类型",
            "动态风险分", "最终优先级", "处理建议"]
    print(priority[cols].head(20).to_string(index=False) if len(priority) else "无异常")

    print(f"\n结果文件：{output}")
    print("=" * 72)


if __name__ == "__main__":
    main()
