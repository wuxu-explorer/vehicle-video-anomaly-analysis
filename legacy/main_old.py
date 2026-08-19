import pandas as pd
from config import (
    COUNT_LEVELS,
    COUNT_SCORE,
    GROUP_CHANNEL_LEVELS,
    FINAL_PRIORITY as PRIORITY_MAP
)

# 读取 Excel
data = pd.read_excel("data/test_data.xlsx")

print("===== 前5行数据 =====")
print(data.head())

print("\n===== 数据规模 =====")
print(data.shape)

print("\n===== 列名 =====")
print(data.columns.tolist())

print("\n===== 数据基本信息 =====")
data.info()
print("\n===== 状态名称 =====")
print(data["状态名称"].value_counts())

print("\n===== 状态类型 =====")
print(data["状态类型"].value_counts())
print("\n===== 摄像头异常：设备编号 =====")
camera_error = data[data["状态名称"] == "摄像头异常"]
print(camera_error["设备编号"].value_counts())

print("\n===== 摄像头异常：车组 =====")
print(camera_error["归属车组"].value_counts())

print("\n===== 摄像头异常：状态内容 =====")
print(camera_error["状态内容"].value_counts())
print("\n===== 摄像头异常：设备编号 + 状态内容 =====")

camera_detail = camera_error[
    ["设备编号", "状态内容"]
].drop_duplicates()

print(camera_detail.head(30).to_string(index=False))
# ==============================
# 生成摄像头异常明细
# ==============================

camera_detail = camera_error[
    ["设备编号", "归属车组", "状态名称", "状态内容"]
].drop_duplicates()

print("\n===== 摄像头异常明细 =====")
print(camera_detail.head(30).to_string(index=False))
# ==============================
# 自动判断 DMS / ADAS / DSC
# ==============================

import re


def get_channel(text):
    """从状态内容中提取通道号"""
    text = str(text)

    match = re.search(r"通道号\s*(\d+)", text)

    if match:
        return int(match.group(1))

    match = re.search(r"通道\s*(\d+)", text)

    if match:
        return int(match.group(1))

    return None


def get_camera_type(group, channel):
    """根据车组 + 通道号判断 DMS / ADAS / DSC"""

    group = str(group)

    # --------------------------
    # 多宝山：HD运输九班
    # 通道1：ADAS
    # 通道5：DMS
    # --------------------------
    if "HD运输九班" in group:
        mapping = {
            1: "ADAS",
            5: "DMS"
        }

        return mapping.get(channel, "其他/未知")


    # --------------------------
    # 西藏巨龙：集为辅助设备
    # 通道2：ADAS
    # 通道5：DMS
    # 通道6：DSC
    # --------------------------
    if "集为辅助设备" in group:
        mapping = {
            2: "ADAS",
            5: "DMS",
            6: "DSC"
        }

        return mapping.get(channel, "其他/未知")


    # --------------------------
    # 紫金锌业：建设宽体车
    # 通道5：DMS
    # 通道6：ADAS
    # 通道7：DSC
    # --------------------------
    if "建设宽体车" in group:
        mapping = {
            5: "DMS",
            6: "ADAS",
            7: "DSC"
        }

        return mapping.get(channel, "其他/未知")


    # --------------------------
    # 如果暂时没有匹配规则
    # --------------------------
    return "未配置"


# 提取通道号
camera_detail["通道号"] = camera_detail["状态内容"].apply(get_channel)

# 判断系统类型
camera_detail["系统类型"] = camera_detail.apply(
    lambda row: get_camera_type(
        row["归属车组"],
        row["通道号"]
    ),
    axis=1
)


print("\n===== 通道 + 系统类型 =====")

print(
    camera_detail[
        ["设备编号", "归属车组", "通道号", "系统类型", "状态内容"]
    ].head(50).to_string(index=False)
)


# ==============================
# 统计 DMS / ADAS / DSC
# ==============================

print("\n===== 系统类型统计 =====")

print(
    camera_detail["系统类型"].value_counts()
)


# ==============================
# 保存最终结果
# ==============================

camera_detail.to_excel(
    "output/摄像头异常分类结果.xlsx",
    index=False
)

print("\n===== 分类完成 =====")
print("结果已保存到：output/摄像头异常分类结果.xlsx")

# 保存结果
camera_detail.to_excel(
    "output/摄像头异常明细.xlsx",
    index=False
)

print("\n===== 已完成 =====")
print("结果已保存到：output/摄像头异常明细.xlsx")
# ==============================
# 按车组 + 通道号判断系统类型
# ==============================

def get_system_type(row):
    group = str(row["归属车组"])
    content = str(row["状态内容"])

    # 从“状态内容”中提取通道号
    import re
    match = re.search(r"通道号\s*(\d+)", content)

    if not match:
        return "未配置"

    channel = int(match.group(1))

    # HD运输九班
    if group == "HD运输九班":
        if channel == 1:
            return "ADAS"
        elif channel == 5:
            return "DMS"

    # 集为辅助设备
    elif group == "集为辅助设备":
        if channel == 2:
            return "ADAS"
        elif channel == 5:
            return "DMS"
        elif channel == 6:
            return "DSC"

    # 建设宽体车
    elif group == "建设宽体车":
        if channel == 5:
            return "DMS"
        elif channel == 6:
            return "ADAS"
        elif channel == 7:
            return "DSC"

    return "未配置"


# 添加系统类型
camera_detail["系统类型"] = camera_detail.apply(
    get_system_type,
    axis=1
)

# 查看系统类型统计
print("\n===== 最终系统类型统计 =====")
print(camera_detail["系统类型"].value_counts())

# 保存最终结果
camera_detail.to_excel(
    "output/摄像头异常最终分类.xlsx",
    index=False
)

print("\n===== 最终分类完成 =====")
print("结果已保存到：output/摄像头异常最终分类.xlsx")
# ==============================
# 统计所有车组 + 通道号
# ==============================

import re

def get_channel(content):
    content = str(content)

    match = re.search(r"通道号\s*(\d+)", content)

    if match:
        return int(match.group(1))

    return None


camera_detail["通道号"] = camera_detail["状态内容"].apply(get_channel)

# 统计车组 + 通道号
channel_statistics = (
    camera_detail
    .dropna(subset=["通道号"])
    .groupby(["归属车组", "通道号"])
    .size()
    .reset_index(name="视频丢失次数")
    .sort_values("视频丢失次数", ascending=False)
)

print("\n===== 车组 + 通道号统计 =====")
print(channel_statistics.to_string(index=False))

# 保存统计结果
channel_statistics.to_excel(
    "output/车组通道号统计.xlsx",
    index=False
)

print("\n===== 统计完成 =====")
print("结果已保存到：output/车组通道号统计.xlsx")
# ==========================================
# 根据通道号判断系统类型
# ==========================================

def get_system_type(channel):
    channel = int(channel)

    if channel == 2:
        return "ADAS"

    elif channel == 5:
        return "DMS"

    elif channel == 6:
        return "DSC"

    elif channel == 1:
        return "原车倒车"

    elif channel in [3, 4, 9]:
        return "未接摄像机"

    elif channel in [7, 8, 10, 11]:
        return "盲区摄像头"

    elif channel == 12:
        return "前摄像机"

    else:
        return "未知"


# 添加系统类型
df = pd.read_excel("output/车组通道号统计.xlsx")

df["系统类型"] = df["通道号"].apply(get_system_type)


# 查看分类结果
print("\n===== 通道号系统类型统计 =====")
print(df["系统类型"].value_counts())


# 查看 DMS / ADAS / DSC 的详细情况
print("\n===== DMS / ADAS / DSC 统计 =====")
system_data = df[df["系统类型"].isin(["DMS", "ADAS", "DSC"])]

print(
    system_data
    .groupby(["归属车组", "通道号", "系统类型"])["视频丢失次数"]
    .sum()
    .reset_index()
    .sort_values("视频丢失次数", ascending=False)
    .to_string(index=False)
)


# 保存结果
df.to_excel(
    "output/车组通道号系统分类.xlsx",
    index=False
)

print("\n===== 分类完成 =====")
print("结果已保存到：output/车组通道号系统分类.xlsx")
# ==========================================
# 最终巡检汇总
# ==========================================

print("\n===== 开始生成最终巡检汇总 =====")

# 读取已经完成系统分类的数据
final_data = pd.read_excel("output/车组通道号系统分类.xlsx")

# ------------------------------------------
# 1. 总体系统类型统计
# ------------------------------------------
system_summary = (
    final_data
    .groupby("系统类型")["视频丢失次数"]
    .sum()
    .reset_index()
    .sort_values("视频丢失次数", ascending=False)
)

print("\n===== 系统类型总体统计 =====")
print(system_summary.to_string(index=False))


# ------------------------------------------
# 2. 车组统计
# ------------------------------------------
group_summary = (
    final_data
    .groupby("归属车组")["视频丢失次数"]
    .sum()
    .reset_index()
    .sort_values("视频丢失次数", ascending=False)
)

print("\n===== 车组异常统计 TOP20 =====")
print(group_summary.head(20).to_string(index=False))


# ------------------------------------------
# 3. DMS / ADAS / DSC 专项统计
# ------------------------------------------
main_systems = ["DMS", "ADAS", "DSC"]

system_detail = final_data[
    final_data["系统类型"].isin(main_systems)
].copy()

system_detail_summary = (
    system_detail
    .groupby(["系统类型", "归属车组"])["视频丢失次数"]
    .sum()
    .reset_index()
    .sort_values(
        ["系统类型", "视频丢失次数"],
        ascending=[True, False]
    )
)

print("\n===== DMS / ADAS / DSC 统计 =====")
print(system_detail_summary.to_string(index=False))


# ------------------------------------------
# 4. 异常 TOP100
# ------------------------------------------
top_errors = (
    final_data[
        ["归属车组", "通道号", "视频丢失次数", "系统类型"]
    ]
    .sort_values("视频丢失次数", ascending=False)
    .head(100)
)

print("\n===== 视频丢失 TOP100 =====")
print(top_errors.head(20).to_string(index=False))


# ------------------------------------------
# 5. 保存到一个 Excel 的多个工作表
# ------------------------------------------
output_file = "output/视频异常巡检最终汇总.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

    system_summary.to_excel(
        writer,
        sheet_name="总体系统统计",
        index=False
    )

    group_summary.to_excel(
        writer,
        sheet_name="车组异常统计",
        index=False
    )

    system_detail_summary.to_excel(
        writer,
        sheet_name="DMS_ADAS_DSC统计",
        index=False
    )

    top_errors.to_excel(
        writer,
        sheet_name="异常TOP100",
        index=False
    )


print("\n===== 最终巡检汇总完成 =====")
print("结果已保存到：output/视频异常巡检最终汇总.xlsx")
# ==========================================
# 自动计算异常等级
# ==========================================

def get_count_level(count):

    if count >= COUNT_LEVELS["严重"]:
        return "严重"

    elif count >= COUNT_LEVELS["较严重"]:
        return "较严重"

    elif count >= COUNT_LEVELS["一般"]:
        return "一般"

    else:
        return "轻微"


# ==========================================
# 计算异常优先级
# ==========================================

def calculate_priority(row):
    global FINAL_PRIORITY

    system = str(row["系统类型"])
    count = int(row["视频丢失次数"])

    # 系统基础优先级
    FINAL_PRIORITY = PRIORITY_MAP.get(
        system,
        FINAL_PRIORITY["未知"]
    )

    # 次数等级
    count_level = get_count_level(count)

    # 次数分数
    count_score = COUNT_SCORE[count_level]

    # 系统优先级转分数
    system_score = 5 - FINAL_PRIORITY

    # 最终分数
    total_score = system_score + count_score

    # 判断最终等级
    if total_score >= FINAL_PRIORITY["P1"]:
        priority = "P1"

    elif total_score >= FINAL_PRIORITY["P2"]:
        priority = "P2"

    elif total_score >= FINAL_PRIORITY["P3"]:
        priority = "P3"

    else:
        priority = "P4"

    return pd.Series([
        count_level,
        total_score,
        priority
    ])


# ==========================================
# 读取最终分类数据
# ==========================================

priority_data = pd.read_excel(
    "output/车组通道号系统分类.xlsx"
)


# 计算优先级
priority_data[
    ["异常等级", "优先级分数", "最终优先级"]
] = priority_data.apply(
    calculate_priority,
    axis=1
)


# ==========================================
# 按优先级 + 异常次数排序
# ==========================================

priority_order = {
    "P1": 1,
    "P2": 2,
    "P3": 3,
    "P4": 4
}

priority_data["排序"] = (
    priority_data["最终优先级"]
    .map(priority_order)
)

priority_data = priority_data.sort_values(
    ["排序", "视频丢失次数"],
    ascending=[True, False]
)


# ==========================================
# 保存优先级分析
# ==========================================

priority_data = priority_data.drop(
    columns=["排序"]
)

priority_data.to_excel(
    "output/视频异常优先级分析.xlsx",
    index=False
)


print("\n===== 优先级分析完成 =====")

print(
    priority_data[
        [
            "归属车组",
            "通道号",
            "视频丢失次数",
            "系统类型",
            "异常等级",
            "优先级分数",
            "最终优先级"
        ]
    ].head(30).to_string(index=False)
)

print(
    "\n结果已保存到："
    "output/视频异常优先级分析.xlsx"
)
# ==========================================
# 动态优先级 V2
# 数据驱动风险评分
# ==========================================

from config import (
    CORE_SYSTEMS,
    NORMAL_CONFIG_TYPES,
    SYSTEM_WEIGHT,
    COUNT_WEIGHT,
    GROUP_WEIGHT,
    SYSTEM_WEIGHT_RATIO,
    CORE_SYSTEM_BONUS,
    PRIORITY_LEVELS
)


print("\n")
print("=" * 60)
print("开始进行动态优先级 V2 分析")
print("=" * 60)


# ==========================================
# 读取系统分类结果
# ==========================================

dynamic_data = pd.read_excel(
    "output/车组通道号系统分类.xlsx"
)


# ==========================================
# 数据基本检查
# ==========================================

required_columns = [
    "归属车组",
    "通道号",
    "视频丢失次数",
    "系统类型"
]

for column in required_columns:

    if column not in dynamic_data.columns:

        raise ValueError(
            f"数据中缺少必要字段：{column}"
        )


# ==========================================
# 1. 统计当前报表
# ==========================================

total_count = dynamic_data[
    "视频丢失次数"
].sum()

max_count = dynamic_data[
    "视频丢失次数"
].max()


if max_count <= 0:
    max_count = 1


# ==========================================
# 2. 统计每个车组
# ==========================================

group_stats = (
    dynamic_data
    .groupby("归属车组")
    .agg(
        车组异常总次数=(
            "视频丢失次数",
            "sum"
        ),
        异常通道数量=(
            "通道号",
            "nunique"
        )
    )
    .reset_index()
)


max_group_count = group_stats[
    "车组异常总次数"
].max()

max_channel_count = group_stats[
    "异常通道数量"
].max()


if max_group_count <= 0:
    max_group_count = 1

if max_channel_count <= 0:
    max_channel_count = 1


# ==========================================
# 3. 统计每种系统的异常次数
# ==========================================

system_stats = (
    dynamic_data
    .groupby("系统类型")[
        "视频丢失次数"
    ]
    .sum()
    .reset_index()
)


system_stats = system_stats.rename(
    columns={
        "视频丢失次数": "系统异常总次数"
    }
)


# ==========================================
# 4. 合并车组统计
# ==========================================

dynamic_data = dynamic_data.merge(
    group_stats,
    on="归属车组",
    how="left"
)


# ==========================================
# 5. 合并系统统计
# ==========================================

dynamic_data = dynamic_data.merge(
    system_stats,
    on="系统类型",
    how="left"
)


# ==========================================
# 6. 动态风险评分
# ==========================================

def calculate_risk(row):

    count = float(
        row["视频丢失次数"]
    )

    group_total = float(
        row["车组异常总次数"]
    )

    channel_count = float(
        row["异常通道数量"]
    )

    system = str(
        row["系统类型"]
    )


    # --------------------------------------
    # A. 异常次数风险
    #
    # 当前数据中越接近最高异常次数，
    # 风险越高
    # --------------------------------------

    count_score = (
        count / max_count
    ) * COUNT_WEIGHT * 100


    # --------------------------------------
    # B. 车组集中度
    #
    # 同一个车组问题越集中，
    # 风险越高
    # --------------------------------------

    group_score = (
        group_total / max_group_count
    ) * GROUP_WEIGHT * 100


    # --------------------------------------
    # C. 系统因素
    # --------------------------------------

    system_weight = SYSTEM_WEIGHT.get(
        system,
        0.2
    )

    max_system_weight = max(
        SYSTEM_WEIGHT.values()
    )

    system_score = (
        system_weight
        / max_system_weight
    ) * SYSTEM_WEIGHT_RATIO * 100


    # --------------------------------------
    # D. 核心系统额外权重
    # --------------------------------------

    core_bonus = 0

    if system in CORE_SYSTEMS:

        core_bonus = (
            CORE_SYSTEM_BONUS * 100
        )


    # --------------------------------------
    # E. 最终风险分
    # --------------------------------------

    risk_score = (
        count_score
        + group_score
        + system_score
        + core_bonus
    )


    # 防止超过100
    risk_score = min(
        risk_score,
        100
    )


    # --------------------------------------
    # F. 配置判断
    # --------------------------------------

    if system in NORMAL_CONFIG_TYPES:

        config_status = "需要确认配置"

    else:

        config_status = "疑似实际异常"


    # --------------------------------------
    # G. 最终优先级
    # --------------------------------------

    if risk_score >= PRIORITY_LEVELS["P1"]:

        priority = "P1"

    elif risk_score >= PRIORITY_LEVELS["P2"]:

        priority = "P2"

    elif risk_score >= PRIORITY_LEVELS["P3"]:

        priority = "P3"

    else:

        priority = "P4"


    return pd.Series([
        round(risk_score, 2),
        priority,
        config_status
    ])


# ==========================================
# 7. 执行风险计算
# ==========================================

dynamic_data[
    [
        "动态风险分",
        "最终优先级",
        "配置判断"
    ]
] = dynamic_data.apply(
    calculate_risk,
    axis=1
)


# ==========================================
# 8. 排序
# ==========================================

priority_order = {
    "P1": 1,
    "P2": 2,
    "P3": 3,
    "P4": 4
}


dynamic_data["优先级排序"] = (
    dynamic_data[
        "最终优先级"
    ].map(priority_order)
)


dynamic_data = dynamic_data.sort_values(
    [
        "优先级排序",
        "动态风险分",
        "视频丢失次数"
    ],
    ascending=[
        True,
        False,
        False
    ]
)


dynamic_data = dynamic_data.drop(
    columns=[
        "优先级排序"
    ]
)


# ==========================================
# 9. 输出结果
# ==========================================

dynamic_output = (
    "output/视频异常动态优先级V2.xlsx"
)


dynamic_data.to_excel(
    dynamic_output,
    index=False
)


# ==========================================
# 10. 输出 TOP30
# ==========================================

print("\n")
print("=" * 60)
print("动态优先级 V2 TOP30")
print("=" * 60)


print(
    dynamic_data[
        [
            "归属车组",
            "通道号",
            "视频丢失次数",
            "系统类型",
            "车组异常总次数",
            "异常通道数量",
            "动态风险分",
            "最终优先级",
            "配置判断"
        ]
    ]
    .head(30)
    .to_string(index=False)
)


# ==========================================
# 11. 优先级统计
# ==========================================

print("\n")
print("=" * 60)
print("优先级统计")
print("=" * 60)


print(
    dynamic_data[
        "最终优先级"
    ]
    .value_counts()
    .sort_index()
)


print("\n")
print("=" * 60)
print("动态优先级 V2 分析完成")
print("=" * 60)

print(
    f"本次报表总异常次数：{total_count}"
)

print(
    f"结果文件：{dynamic_output}"
)