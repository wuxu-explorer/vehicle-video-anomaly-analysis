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
    CONFIG_LIKE_SYSTEMS,
    PRIORITY_THRESHOLDS,
    PRIORITY_ORDER,
)


def normalize_text(value: Any) -> str:
    """统一清洗 Excel 中的文本值。"""
    if pd.isna(value):
        return ""
    return str(value).strip()


def extract_channel(text: Any) -> int | None:
    """从状态内容中提取通道号，兼容“通道号12 / 通道号:12 / 通道12”等写法。"""
    text = normalize_text(text)

    patterns = [
        r"通道号\s*[:：]?\s*(\d+)",
        r"通道\s*[:：]?\s*(\d+)",
        r"channel\s*[:：]?\s*(\d+)",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return int(match.group(1))

    return None


def classify_system(group: Any, channel: Any) -> tuple[str, str]:
    """
    根据“车组 + 通道号”判断系统类型。
    返回：(系统类型, 分类依据)

    车组专属规则优先；没有匹配时才使用通用通道规则。
    """

    group_text = normalize_text(group)

    if channel is None or str(channel).strip() == "" or pd.isna(channel):
        return "未知", "未提取到通道号"

    try:
        channel_num = int(float(channel))
    except (TypeError, ValueError):
        return "未知", "通道号不是有效数字"

    # 1. 车组专属规则优先
    for rule in GROUP_CHANNEL_RULES:
        if rule["pattern"] in group_text:
            mapping = rule["channels"]

            if channel_num in mapping:
                return mapping[channel_num], f"车组专属规则：{rule['name']}"

            return "未配置", f"已匹配车组规则，但通道 {channel_num} 未配置"

    # 2. 通用规则兜底
    if channel_num in DEFAULT_CHANNEL_RULES:
        return DEFAULT_CHANNEL_RULES[channel_num], "通用通道规则"

    return "未知", "没有匹配到分类规则"


def require_columns(
    df: pd.DataFrame,
    columns: list[str],
    stage: str,
) -> None:
    """检查输入表是否包含必要字段。"""

    missing = [column for column in columns if column not in df.columns]

    if missing:
        raise ValueError(
            f"{stage}缺少必要字段：{', '.join(missing)}\n"
            f"当前字段：{', '.join(map(str, df.columns))}"
        )


def clean_input(df: pd.DataFrame) -> pd.DataFrame:
    """清洗输入 Excel。"""

    df = df.copy()

    # 清理列名空格
    df.columns = [normalize_text(column) for column in df.columns]

    require_columns(df, REQUIRED_COLUMNS, "输入数据")

    for column in [
        "状态名称",
        "状态类型",
        "设备编号",
        "归属车组",
        "状态内容",
    ]:
        df[column] = df[column].map(normalize_text)

    return df


def build_channel_stats(
    data: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    从原始数据生成：
    1. 摄像头异常明细
    2. 车组 + 通道 + 系统类型统计

    注意：这里不再 drop_duplicates()，
    防止原始重复异常记录被错误地当成一条。
    """

    camera = data[
        data["状态名称"].isin(CAMERA_ERROR_NAMES)
    ].copy()

    camera["通道号"] = camera["状态内容"].map(extract_channel)

    classified = camera.apply(
        lambda row: classify_system(
            row["归属车组"],
            row["通道号"],
        ),
        axis=1,
        result_type="expand",
    )

    classified.columns = [
        "系统类型",
        "分类依据",
    ]

    camera[
        ["系统类型", "分类依据"]
    ] = classified

    detail_columns = [
        "设备编号",
        "归属车组",
        "状态名称",
        "状态类型",
        "通道号",
        "系统类型",
        "分类依据",
        "状态内容",
    ]

    detail = camera[detail_columns].copy()

    # 每一条原始异常记录计 1 次。
    stats = (
        camera.assign(
            通道号=camera["通道号"]
            .fillna(-1)
            .astype(int)
        )
        .groupby(
            [
                "归属车组",
                "通道号",
                "系统类型",
                "分类依据",
            ],
            dropna=False,
        )
        .size()
        .reset_index(name="视频丢失次数")
    )

    stats["通道号"] = stats["通道号"].replace(
        -1,
        np.nan,
    )

    return detail, stats


def minmax(series: pd.Series) -> pd.Series:
    """0~1 归一化；如果所有值相同，则给 0.5。"""

    values = pd.to_numeric(
        series,
        errors="coerce",
    ).fillna(0.0)

    low = float(values.min())
    high = float(values.max())

    if high <= low:
        return pd.Series(
            0.5,
            index=values.index,
        )

    return (values - low) / (high - low)


def percentile_score(series: pd.Series) -> pd.Series:
    """
    当前报表内部排名分。
    新报表换了数量级后，会自动重新计算。
    """

    values = pd.to_numeric(
        series,
        errors="coerce",
    ).fillna(0.0)

    if len(values) <= 1:
        return pd.Series(
            1.0,
            index=values.index,
        )

    return values.rank(
        method="average",
        pct=True,
    )


def dynamic_priority(
    stats: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    动态风险评分核心。

    不依赖固定的“100次=P1”之类绝对阈值，
    而是根据本次报表内部的相对风险重新排名。

    因此下一张报表即使：
    - 总异常次数改变
    - 最大次数改变
    - 车组数量改变
    - DMS/ADAS/DSC数量比例改变

    仍会重新计算最高优先级。
    """

    df = stats.copy()

    df["视频丢失次数"] = pd.to_numeric(
        df["视频丢失次数"],
        errors="coerce",
    ).fillna(0).astype(int)

    # -----------------------------
    # 车组维度
    # -----------------------------
    group_stats = (
        df.groupby(
            "归属车组",
            dropna=False,
        )
        .agg(
            车组异常总次数=(
                "视频丢失次数",
                "sum",
            ),
            异常通道数量=(
                "通道号",
                "nunique",
            ),
        )
        .reset_index()
    )

    # -----------------------------
    # 系统维度
    # -----------------------------
    system_stats = (
        df.groupby(
            "系统类型",
            dropna=False,
        )
        .agg(
            系统异常总次数=(
                "视频丢失次数",
                "sum",
            ),
            系统异常通道数=(
                "通道号",
                "nunique",
            ),
        )
        .reset_index()
    )

    df = df.merge(
        group_stats,
        on="归属车组",
        how="left",
    )

    df = df.merge(
        system_stats,
        on="系统类型",
        how="left",
    )

    # -----------------------------
    # 系统基础权重
    # -----------------------------
    df["系统权重"] = (
        df["系统类型"]
        .map(SYSTEM_WEIGHTS)
        .fillna(0.20)
    )

    max_system_weight = (
        max(SYSTEM_WEIGHTS.values())
        if SYSTEM_WEIGHTS
        else 1.0
    )

    df["系统基础分"] = (
        df["系统权重"]
        / max_system_weight
    ).clip(0, 1)

    # DMS / ADAS / DSC 属于核心系统
    df["核心系统加分"] = (
        df["系统类型"]
        .isin(CORE_SYSTEMS)
        .astype(float)
        * 0.10
    )

    # -----------------------------
    # 本次报表动态指标
    # -----------------------------
    df["次数相对分"] = percentile_score(
        df["视频丢失次数"]
    )

    df["车组集中分"] = minmax(
        df["车组异常总次数"]
    )

    df["通道扩散分"] = minmax(
        df["异常通道数量"]
    )

    df["系统集中分"] = minmax(
        df["系统异常总次数"]
    )

    # -----------------------------
    # 综合风险分
    # -----------------------------
    df["动态风险分"] = (
        df["次数相对分"] * 0.45
        + df["车组集中分"] * 0.20
        + df["通道扩散分"] * 0.10
        + df["系统集中分"] * 0.10
        + df["系统基础分"] * 0.10
        + df["核心系统加分"] * 0.05
    ) * 100

    df["动态风险分"] = df[
        "动态风险分"
    ].round(2).clip(0, 100)

    # -----------------------------
    # 配置/真实异常提示
    # -----------------------------
    df["配置判断"] = np.where(
        df["系统类型"].isin(
            CONFIG_LIKE_SYSTEMS
        ),
        "需要确认配置/通道方案",
        np.where(
            df["系统类型"].isin(
                [
                    "未知",
                    "未配置",
                ]
            ),
            "规则未覆盖，需人工确认",
            "疑似实际异常",
        ),
    )

    # -----------------------------
    # 动态优先级
    # -----------------------------
    if len(df) == 0:
        df["最终优先级"] = pd.Series(
            dtype=str
        )
    elif len(df) == 1:
        df["最终优先级"] = "P1"
    else:
        # 风险最高的记录一定排第一。
        rank = df[
            "动态风险分"
        ].rank(
            method="first",
            ascending=False,
        )

        count = len(df)

        p1_cut = max(
            1,
            int(
                np.ceil(
                    count
                    * PRIORITY_THRESHOLDS["P1"]
                )
            ),
        )

        p2_cut = max(
            p1_cut,
            int(
                np.ceil(
                    count
                    * PRIORITY_THRESHOLDS["P2"]
                )
            ),
        )

        p3_cut = max(
            p2_cut,
            int(
                np.ceil(
                    count
                    * PRIORITY_THRESHOLDS["P3"]
                )
            ),
        )

        df["最终优先级"] = np.select(
            [
                rank <= p1_cut,
                rank <= p2_cut,
                rank <= p3_cut,
            ],
            [
                "P1",
                "P2",
                "P3",
            ],
            default="P4",
        )

    # -----------------------------
    # 辅助异常等级
    # -----------------------------
    df["异常等级"] = pd.cut(
        df["视频丢失次数"],
        bins=[
            -1,
            0,
            4,
            9,
            np.inf,
        ],
        labels=[
            "无",
            "轻微",
            "较严重",
            "严重",
        ],
    ).astype(str)

    # -----------------------------
    # 最终排序
    # -----------------------------
    df["优先级排序"] = (
        df["最终优先级"]
        .map(PRIORITY_ORDER)
        .fillna(99)
    )

    df = df.sort_values(
        [
            "优先级排序",
            "动态风险分",
            "视频丢失次数",
        ],
        ascending=[
            True,
            False,
            False,
        ],
    )

    priority_summary = (
        df["最终优先级"]
        .value_counts()
        .reindex(
            [
                "P1",
                "P2",
                "P3",
                "P4",
            ],
            fill_value=0,
        )
        .rename_axis(
            "最终优先级"
        )
        .reset_index(
            name="异常通道数"
        )
    )

    return (
        df.drop(
            columns=[
                "优先级排序"
            ]
        ),
        priority_summary,
    )


def build_workbook(
    detail: pd.DataFrame,
    stats: pd.DataFrame,
    priority: pd.DataFrame,
    priority_summary: pd.DataFrame,
) -> Path:
    """生成最终 Excel。"""

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    output = (
        OUTPUT_DIR
        / "视频异常巡检最终汇总_动态版.xlsx"
    )

    # 总体系统统计
    system_summary = (
        stats.groupby(
            "系统类型"
        )["视频丢失次数"]
        .sum()
        .reset_index()
        .sort_values(
            "视频丢失次数",
            ascending=False,
        )
    )

    # 车组统计
    group_summary = (
        stats.groupby(
            "归属车组"
        )["视频丢失次数"]
        .sum()
        .reset_index()
        .sort_values(
            "视频丢失次数",
            ascending=False,
        )
    )

    # DMS / ADAS / DSC
    core_summary = (
        stats[
            stats["系统类型"].isin(
                CORE_SYSTEMS
            )
        ]
        .groupby(
            [
                "系统类型",
                "归属车组",
            ]
        )["视频丢失次数"]
        .sum()
        .reset_index()
        .sort_values(
            "视频丢失次数",
            ascending=False,
        )
    )

    # TOP100
    top100_columns = [
        "归属车组",
        "通道号",
        "视频丢失次数",
        "系统类型",
        "动态风险分",
        "最终优先级",
        "配置判断",
        "分类依据",
    ]

    top100 = priority[
        top100_columns
    ].head(100)

    # 规则覆盖检查
    rule_check = (
        stats.groupby(
            [
                "系统类型",
                "分类依据",
            ],
            dropna=False,
        )
        .agg(
            视频丢失次数=(
                "视频丢失次数",
                "sum",
            ),
            车组数=(
                "归属车组",
                "nunique",
            ),
            通道数=(
                "通道号",
                "nunique",
            ),
        )
        .reset_index()
        .sort_values(
            "视频丢失次数",
            ascending=False,
        )
    )

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        system_summary.to_excel(
            writer,
            sheet_name="总体系统统计",
            index=False,
        )

        group_summary.to_excel(
            writer,
            sheet_name="车组异常统计",
            index=False,
        )

        core_summary.to_excel(
            writer,
            sheet_name="DMS_ADAS_DSC统计",
            index=False,
        )

        top100.to_excel(
            writer,
            sheet_name="动态优先级TOP100",
            index=False,
        )

        priority.to_excel(
            writer,
            sheet_name="完整动态优先级",
            index=False,
        )

        priority_summary.to_excel(
            writer,
            sheet_name="优先级统计",
            index=False,
        )

        rule_check.to_excel(
            writer,
            sheet_name="规则覆盖检查",
            index=False,
        )

        detail.to_excel(
            writer,
            sheet_name="摄像头异常明细",
            index=False,
        )

    return output


def main() -> None:
    print("=" * 70)
    print("视频异常巡检 - 动态优先级 V3")
    print("=" * 70)

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"找不到输入文件：{INPUT_FILE}\n"
            "请把新的报表放入 data 文件夹，"
            "并命名为 test_data.xlsx。"
        )

    data = pd.read_excel(
        INPUT_FILE
    )

    data = clean_input(
        data
    )

    print(
        f"输入文件：{INPUT_FILE}"
    )

    print(
        f"原始数据量：{len(data)}"
    )

    detail, stats = (
        build_channel_stats(data)
    )

    priority, priority_summary = (
        dynamic_priority(stats)
    )

    output = build_workbook(
        detail,
        stats,
        priority,
        priority_summary,
    )

    print("\n===== 本次报表优先级统计 =====")
    print(
        priority_summary.to_string(
            index=False
        )
    )

    print("\n===== 动态优先级 TOP20 =====")

    preview_columns = [
        "归属车组",
        "通道号",
        "视频丢失次数",
        "系统类型",
        "动态风险分",
        "最终优先级",
        "配置判断",
    ]

    if len(priority) > 0:
        print(
            priority[
                preview_columns
            ]
            .head(20)
            .to_string(
                index=False
            )
        )
    else:
        print("本次没有检测到摄像头异常。")

    print("\n" + "=" * 70)
    print("分析完成")
    print(
        f"结果文件：{output}"
    )
    print("=" * 70)


if __name__ == "__main__":
    main()
