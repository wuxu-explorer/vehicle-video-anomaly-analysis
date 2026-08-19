# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from config import (
    INPUT_FILE,
    OUTPUT_DIR,
    CAMERA_ERROR_NAMES,
    REQUIRED_COLUMNS,
    GROUP_CHANNEL_RULES,
    DEFAULT_CHANNEL_RULES,
    CORE_SYSTEMS,
    SYSTEM_WEIGHTS,
    PRIORITY_CONFIG,
    PRIORITY_ORDER,
    PRIORITY_ADVICE,
    REVIEW_SYSTEMS,
)


# ============================================================
# 基础工具
# ============================================================

def text(v: Any) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def safe_int(v: Any) -> int | None:
    try:
        if pd.isna(v) or text(v) == "":
            return None
        return int(float(v))
    except (TypeError, ValueError):
        return None


def channel_from_content(v: Any) -> int | None:
    """从状态内容中尽可能稳健地提取通道号。"""
    s = text(v)
    if not s:
        return None

    patterns = [
        r"通道号\s*[:：]?\s*(\d+)",
        r"通道\s*[:：]?\s*(\d+)",
        r"channel\s*[:：]?\s*(\d+)",
        r"ch(?:annel)?\s*[:：]?\s*(\d+)",
    ]
    for p in patterns:
        m = re.search(p, s, flags=re.I)
        if m:
            return int(m.group(1))

    # 某些报表可能直接写“CH5”或“5通道”。
    patterns2 = [
        r"\bCH\s*[-_：:]?\s*(\d+)\b",
        r"\b(\d+)\s*通道\b",
    ]
    for p in patterns2:
        m = re.search(p, s, flags=re.I)
        if m:
            return int(m.group(1))

    return None


# ============================================================
# 车组 + 通道分类
# ============================================================

def classify(group: Any, channel: Any) -> tuple[str, str, str]:
    """
    返回：
    系统类型、分类依据、规则状态

    规则状态：
    已确认 / 待确认
    """
    g = text(group)
    ch = safe_int(channel)

    if ch is None:
        return "未知", "无法提取通道号", "待确认"

    # 特殊车组优先。
    for rule in GROUP_CHANNEL_RULES:
        if rule["pattern"] in g:
            if ch in rule["channels"]:
                return (
                    rule["channels"][ch],
                    f"车组专属规则：{rule['name']}",
                    "已确认",
                )
            return (
                "未配置",
                f"车组“{rule['name']}”未覆盖通道{ch}",
                "待确认",
            )

    if ch in DEFAULT_CHANNEL_RULES:
        return DEFAULT_CHANNEL_RULES[ch], "通用通道规则", "已确认"

    return "未知", f"规则库未覆盖通道{ch}", "待确认"


# ============================================================
# 输入数据
# ============================================================

def load_data() -> pd.DataFrame:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"\n找不到输入文件：{INPUT_FILE}\n"
            f"请把新的原始Excel放入 data 文件夹，并命名为 test_data.xlsx。"
        )

    df = pd.read_excel(INPUT_FILE)
    df.columns = [text(c) for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "\n输入报表字段不完整。\n"
            f"缺少字段：{missing}\n"
            f"当前字段：{list(df.columns)}\n\n"
            "当前版本要求原始异常表至少包含："
            f"{REQUIRED_COLUMNS}"
        )

    for c in REQUIRED_COLUMNS:
        df[c] = df[c].map(text)

    return df


# ============================================================
# 异常明细 -> 车组/通道统计
# ============================================================

def build_stats(raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    # 状态名称放宽识别，兼容“摄像头异常/视频异常/视频丢失”。
    camera_names = {text(x) for x in CAMERA_ERROR_NAMES}
    camera = raw[raw["状态名称"].isin(camera_names)].copy()

    # 如果严格名称没有命中，尝试从状态内容/状态名称中识别。
    if camera.empty:
        mask = (
            raw["状态名称"].str.contains(
                "摄像头|视频丢失|视频异常", case=False, na=False
            )
            | raw["状态内容"].str.contains(
                "摄像头异常|视频丢失|视频异常", case=False, na=False
            )
        )
        camera = raw[mask].copy()

    camera["通道号"] = camera["状态内容"].map(channel_from_content)

    classified = camera.apply(
        lambda r: classify(r["归属车组"], r["通道号"]),
        axis=1,
        result_type="expand",
    )
    classified.columns = ["系统类型", "分类依据", "规则状态"]
    camera[["系统类型", "分类依据", "规则状态"]] = classified

    detail_cols = [
        "设备编号",
        "归属车组",
        "状态名称",
        "状态类型",
        "通道号",
        "系统类型",
        "分类依据",
        "规则状态",
        "状态内容",
    ]
    detail = camera[detail_cols].copy()

    if camera.empty:
        stats = pd.DataFrame(
            columns=[
                "归属车组",
                "通道号",
                "系统类型",
                "分类依据",
                "规则状态",
                "视频丢失次数",
            ]
        )
        return detail, stats

    stats = (
        camera.groupby(
            [
                "归属车组",
                "通道号",
                "系统类型",
                "分类依据",
                "规则状态",
            ],
            dropna=False,
        )
        .size()
        .reset_index(name="视频丢失次数")
    )

    return detail, stats


# ============================================================
# 动态评分
# ============================================================

def pct_rank(s: pd.Series) -> pd.Series:
    if len(s) <= 1:
        return pd.Series(1.0, index=s.index)
    return s.rank(method="average", pct=True)


def normalize(s: pd.Series) -> pd.Series:
    s = pd.to_numeric(s, errors="coerce").fillna(0)
    lo = float(s.min()) if len(s) else 0.0
    hi = float(s.max()) if len(s) else 0.0

    if hi <= lo:
        return pd.Series(0.5, index=s.index)

    return (s - lo) / (hi - lo)


def dynamic_cut(n: int, share: float) -> int:
    if n <= 0:
        return 0
    return max(1, int(np.ceil(n * share)))


def score_and_priority(stats: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = stats.copy()

    if df.empty:
        return (
            df,
            pd.DataFrame(
                {
                    "最终优先级": ["P1", "P2", "P3", "P4"],
                    "异常通道数": [0, 0, 0, 0],
                }
            ),
        )

    df["视频丢失次数"] = (
        pd.to_numeric(df["视频丢失次数"], errors="coerce")
        .fillna(0)
        .astype(int)
    )

    # 车组总体影响。
    group_total = (
        df.groupby("归属车组", dropna=False)["视频丢失次数"]
        .sum()
        .rename("车组异常总次数")
    )

    # 系统总体影响。
    system_total = (
        df.groupby("系统类型", dropna=False)["视频丢失次数"]
        .sum()
        .rename("系统异常总次数")
    )

    # 一个车组涉及多少个异常通道。
    spread = (
        df.groupby("归属车组", dropna=False)["通道号"]
        .nunique(dropna=True)
        .rename("车组异常通道数")
    )

    df = (
        df.join(group_total, on="归属车组")
        .join(system_total, on="系统类型")
        .join(spread, on="归属车组")
    )

    # 系统重要性。
    df["系统权重"] = df["系统类型"].map(SYSTEM_WEIGHTS).fillna(0.10)
    max_w = max(SYSTEM_WEIGHTS.values()) if SYSTEM_WEIGHTS else 1.0
    df["系统重要性分"] = (df["系统权重"] / max_w).clip(0, 1)

    # 动态指标。
    df["次数排名分"] = pct_rank(df["视频丢失次数"])
    df["绝对次数分"] = normalize(df["视频丢失次数"])
    df["车组影响分"] = normalize(df["车组异常总次数"])
    df["系统影响分"] = normalize(df["系统异常总次数"])
    df["扩散分"] = normalize(df["车组异常通道数"])
    df["核心系统加分"] = df["系统类型"].isin(CORE_SYSTEMS).astype(float)

    cfg = PRIORITY_CONFIG
    total_weight = (
        cfg["WEIGHT_COUNT_RANK"]
        + cfg["WEIGHT_COUNT_ABSOLUTE"]
        + cfg["WEIGHT_GROUP_IMPACT"]
        + cfg["WEIGHT_SYSTEM_IMPACT"]
        + cfg["WEIGHT_SPREAD"]
        + cfg["WEIGHT_SYSTEM_IMPORTANCE"]
        + cfg["WEIGHT_CORE_SYSTEM"]
    )

    if total_weight <= 0:
        total_weight = 1.0

    # 自动归一化权重，防止以后改配置时总和不等于1。
    risk = (
        df["次数排名分"] * cfg["WEIGHT_COUNT_RANK"]
        + df["绝对次数分"] * cfg["WEIGHT_COUNT_ABSOLUTE"]
        + df["车组影响分"] * cfg["WEIGHT_GROUP_IMPACT"]
        + df["系统影响分"] * cfg["WEIGHT_SYSTEM_IMPACT"]
        + df["扩散分"] * cfg["WEIGHT_SPREAD"]
        + df["系统重要性分"] * cfg["WEIGHT_SYSTEM_IMPORTANCE"]
        + df["核心系统加分"] * cfg["WEIGHT_CORE_SYSTEM"]
    ) / total_weight

    df["动态风险分"] = (risk * 100).round(2)

    # --------------------------------------------------------
    # 动态优先级
    # --------------------------------------------------------
    n = len(df)
    rank = df["动态风险分"].rank(method="first", ascending=False)

    p1_cut = dynamic_cut(n, cfg["P1_MAX_SHARE"])
    p2_cut = max(p1_cut, dynamic_cut(n, cfg["P2_MAX_SHARE"]))
    p3_cut = max(p2_cut, dynamic_cut(n, cfg["P3_MAX_SHARE"]))

    df["最终优先级"] = "P4"

    # 规则未确认：不自动认定为真实核心系统故障。
    confirmed = ~df["系统类型"].isin(REVIEW_SYSTEMS)

    # 小数据集自动放宽最低次数门槛，避免“只有1~2条数据时永远没有P1”。
    p1_min = min(cfg["P1_MIN_COUNT"], max(1, int(np.ceil(n * 0.02))))
    p2_min = min(cfg["P2_MIN_COUNT"], max(1, int(np.ceil(n * 0.01))))
    p3_min = min(cfg["P3_MIN_COUNT"], 1)

    df.loc[
        confirmed
        & (rank <= p1_cut)
        & (df["视频丢失次数"] >= p1_min),
        "最终优先级",
    ] = "P1"

    df.loc[
        confirmed
        & (rank <= p2_cut)
        & (df["视频丢失次数"] >= p2_min)
        & (df["最终优先级"] != "P1"),
        "最终优先级",
    ] = "P2"

    df.loc[
        confirmed
        & (rank <= p3_cut)
        & (df["视频丢失次数"] >= p3_min)
        & (df["最终优先级"].isin(["P3", "P4"])),
        "最终优先级",
    ] = "P3"

    # 待确认规则即使风险很高，也保留“待确认”标识。
    df["规则判断"] = np.where(
        df["系统类型"].isin(REVIEW_SYSTEMS),
        "待确认：不能直接认定为真实系统故障",
        "已按当前规则库分类",
    )

    df["处理建议"] = df["最终优先级"].map(PRIORITY_ADVICE).fillna("常规跟踪")
    df.loc[
        df["系统类型"].isin(REVIEW_SYSTEMS),
        "处理建议",
    ] = "先确认车组/通道规则，再判断是否为真实异常"

    # 异常等级仅用于辅助展示，不参与最终优先级硬切分。
    df["异常等级"] = pd.cut(
        df["视频丢失次数"],
        [-1, 0, 4, 9, 19, np.inf],
        labels=["无", "轻微", "一般", "较严重", "严重"],
    ).astype(str)

    df["优先级排序"] = df["最终优先级"].map(PRIORITY_ORDER).fillna(99)

    df = df.sort_values(
        ["优先级排序", "动态风险分", "视频丢失次数"],
        ascending=[True, False, False],
    ).drop(columns=["优先级排序"])

    summary = (
        df["最终优先级"]
        .value_counts()
        .reindex(["P1", "P2", "P3", "P4"], fill_value=0)
        .rename_axis("最终优先级")
        .reset_index(name="异常通道数")
    )

    return df, summary


# ============================================================
# Excel输出
# ============================================================

def autosize_and_freeze(writer: pd.ExcelWriter) -> None:
    """统一优化Excel可读性。"""
    try:
        from openpyxl.styles import Font
    except ImportError:
        return

    wb = writer.book

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # 表头加粗。
        for cell in ws[1]:
            cell.font = Font(bold=True)

        # 自动列宽，设置合理上限。
        for column_cells in ws.columns:
            max_len = 0
            col_letter = column_cells[0].column_letter
            for cell in column_cells[:200]:
                value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(value))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 42)


def build_output(
    detail: pd.DataFrame,
    stats: pd.DataFrame,
    priority: pd.DataFrame,
    priority_summary: pd.DataFrame,
) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    output = OUTPUT_DIR / "视频异常巡检最终汇总_动态版.xlsx"

    system_summary = (
        stats.groupby("系统类型", dropna=False)["视频丢失次数"]
        .sum()
        .reset_index()
        .sort_values("视频丢失次数", ascending=False)
    )

    group_summary = (
        stats.groupby("归属车组", dropna=False)["视频丢失次数"]
        .sum()
        .reset_index()
        .sort_values("视频丢失次数", ascending=False)
    )

    core_summary = (
        stats[stats["系统类型"].isin(CORE_SYSTEMS)]
        .groupby(["系统类型", "归属车组"], dropna=False)["视频丢失次数"]
        .sum()
        .reset_index()
        .sort_values("视频丢失次数", ascending=False)
    )

    review = stats[stats["规则状态"] == "待确认"].copy()
    review = review.sort_values("视频丢失次数", ascending=False)

    top100 = priority.head(100).copy()

    rule_check = (
        stats.groupby(
            ["系统类型", "分类依据", "规则状态"],
            dropna=False,
        )
        .agg(
            视频丢失次数=("视频丢失次数", "sum"),
            车组数=("归属车组", "nunique"),
            通道数=("通道号", "nunique"),
        )
        .reset_index()
        .sort_values("视频丢失次数", ascending=False)
    )

    # 新增“最高优先级重点”工作表，直接给人工巡检使用。
    focus_cols = [
        "归属车组",
        "通道号",
        "视频丢失次数",
        "系统类型",
        "规则状态",
        "动态风险分",
        "最终优先级",
        "异常等级",
        "处理建议",
    ]
    focus = priority[focus_cols].head(30).copy()

    # 数据质量检查。
    quality = pd.DataFrame(
        {
            "检查项目": [
                "原始记录数",
                "识别出的摄像头/视频异常记录数",
                "统计后的异常车组-通道数",
                "待确认规则车组-通道数",
                "P1通道数",
                "P2通道数",
                "P3通道数",
                "P4通道数",
            ],
            "结果": [
                _RAW_COUNT,
                len(detail),
                len(stats),
                int((stats["规则状态"] == "待确认").sum()) if len(stats) else 0,
                int(
                    priority_summary.loc[
                        priority_summary["最终优先级"] == "P1",
                        "异常通道数",
                    ].iloc[0]
                ),
                int(
                    priority_summary.loc[
                        priority_summary["最终优先级"] == "P2",
                        "异常通道数",
                    ].iloc[0]
                ),
                int(
                    priority_summary.loc[
                        priority_summary["最终优先级"] == "P3",
                        "异常通道数",
                    ].iloc[0]
                ),
                int(
                    priority_summary.loc[
                        priority_summary["最终优先级"] == "P4",
                        "异常通道数",
                    ].iloc[0]
                ),
            ],
        }
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        system_summary.to_excel(writer, "总体系统统计", index=False)
        group_summary.to_excel(writer, "车组异常统计", index=False)
        core_summary.to_excel(writer, "DMS_ADAS_DSC统计", index=False)
        focus.to_excel(writer, "最高优先级重点", index=False)
        top100.to_excel(writer, "动态优先级TOP100", index=False)
        priority.to_excel(writer, "完整动态优先级", index=False)
        priority_summary.to_excel(writer, "优先级统计", index=False)
        review.to_excel(writer, "待确认规则", index=False)
        rule_check.to_excel(writer, "规则覆盖检查", index=False)
        detail.to_excel(writer, "摄像头异常明细", index=False)
        quality.to_excel(writer, "数据质量检查", index=False)

        autosize_and_freeze(writer)

    return output


_RAW_COUNT = 0


def main() -> None:
    global _RAW_COUNT

    print("=" * 72)
    print("视频异常巡检 - 动态优先级 V5")
    print("=" * 72)

    raw = load_data()
    _RAW_COUNT = len(raw)

    print(f"输入文件：{INPUT_FILE}")
    print(f"原始数据量：{len(raw)}")

    detail, stats = build_stats(raw)

    print(f"识别视频/摄像头异常：{len(detail)} 条")
    print(f"统计异常车组-通道：{len(stats)} 个")

    priority, summary = score_and_priority(stats)

    output = build_output(
        detail,
        stats,
        priority,
        summary,
    )

    print("\n===== 动态优先级统计 =====")
    print(summary.to_string(index=False))

    print("\n===== 待确认规则 =====")
    review_count = int((stats["规则状态"] == "待确认").sum()) if len(stats) else 0
    print(f"待确认车组/通道：{review_count}")

    print("\n===== 最高优先级 TOP20 =====")
    cols = [
        "归属车组",
        "通道号",
        "视频丢失次数",
        "系统类型",
        "动态风险分",
        "最终优先级",
        "处理建议",
    ]

    if len(priority):
        print(priority[cols].head(20).to_string(index=False))
    else:
        print("没有识别到视频/摄像头异常。")

    print(f"\n结果文件：{output}")
    print("=" * 72)


if __name__ == "__main__":
    main()
