# -*- coding: utf-8 -*-
"""
视频异常巡检 V8 - 配置文件

V8 支持：
1. 两份原始 Excel 分别独立分析；
2. 自动比较两份分析结果；
3. DMS / ADAS / DSC 始终属于核心系统层，优先于普通摄像头；
4. 新车组/新通道进入“待确认规则”，不会被误判；
5. Excel 被占用时自动生成带时间戳的备用结果。
"""
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"

# 两份输入文件：建议明确命名，避免程序误把旧结果当输入。
INPUT_A = DATA_DIR / "data_A.xlsx"
INPUT_B = DATA_DIR / "data_B.xlsx"

# 如果 data_A/data_B 不存在，自动从 data 中寻找最新两个 xlsx。
# 会忽略临时文件和 V8 输出文件。
AUTO_DISCOVER_TWO_INPUTS = True

COLUMN_ALIASES = {
    "状态名称": ["状态名称", "状态名", "异常名称", "告警名称", "事件名称"],
    "状态类型": ["状态类型", "异常类型", "告警类型", "事件类型"],
    "设备编号": ["设备编号", "设备号", "设备ID", "设备id", "设备 Id"],
    "归属车组": ["归属车组", "车组", "车组名称", "所属车组", "车辆组"],
    "状态内容": ["状态内容", "异常内容", "告警内容", "描述", "详情", "事件内容"],
    "通道号": ["通道号", "通道", "channel", "Channel", "CH"],
}

REQUIRED_COLUMNS = ["状态名称", "状态类型", "设备编号", "归属车组", "状态内容"]

CAMERA_ERROR_NAMES = {
    "摄像头异常", "视频异常", "视频丢失", "视频丢帧", "视频中断"
}

# 特殊车组规则优先于通用通道规则。
GROUP_CHANNEL_RULES = [
    {
        "name": "HD运输九班",
        "pattern": "HD运输九班",
        "channels": {5: "DMS", 6: "ADAS", 7: "DSC"},
    },
    {
        "name": "集为辅助设备",
        "pattern": "集为辅助设备",
        "channels": {2: "ADAS", 5: "DMS", 6: "DSC"},
    },
    {
        "name": "建设宽体车",
        "pattern": "建设宽体车",
        "channels": {5: "DMS", 6: "ADAS", 7: "DSC"},
    },
]

DEFAULT_CHANNEL_RULES = {
    1: "原车倒车",
    2: "ADAS",
    3: "未接摄像机",
    4: "未接摄像机",
    5: "DMS",
    6: "DSC",
    7: "盲区摄像头",
    8: "盲区摄像头",
    9: "未接摄像机",
    10: "盲区摄像头",
    11: "盲区摄像头",
    12: "前摄像机",
}

CORE_SYSTEMS = {"DMS", "ADAS", "DSC"}
CORE_SYSTEM_LAYER = "核心系统"
ORDINARY_SYSTEM_LAYER = "普通摄像头"
REVIEW_SYSTEM_LAYER = "待确认"

SYSTEM_WEIGHTS = {
    "DMS": 1.00,
    "ADAS": 1.00,
    "DSC": 1.00,
    "原车倒车": 0.70,
    "盲区摄像头": 0.70,
    "前摄像机": 0.70,
    "未接摄像机": 0.30,
    "未配置": 0.15,
    "未知": 0.10,
}

REVIEW_SYSTEMS = {"未知", "未配置"}

PRIORITY_CONFIG = {
    "P1_MAX_SHARE": 0.10,
    "P2_MAX_SHARE": 0.30,
    "P3_MAX_SHARE": 0.60,
    "P1_MIN_COUNT": 3,
    "P2_MIN_COUNT": 2,
    "P3_MIN_COUNT": 1,
    "WEIGHT_COUNT_RANK": 0.35,
    "WEIGHT_COUNT_ABSOLUTE": 0.20,
    "WEIGHT_GROUP_IMPACT": 0.15,
    "WEIGHT_SYSTEM_IMPACT": 0.10,
    "WEIGHT_SPREAD": 0.07,
    "WEIGHT_SYSTEM_IMPORTANCE": 0.08,
    "WEIGHT_CORE_SYSTEM": 0.05,
}

PRIORITY_ORDER = {"P1": 1, "P2": 2, "P3": 3, "P4": 4}
PRIORITY_ADVICE = {
    "P1": "立即重点处理",
    "P2": "优先处理",
    "P3": "持续关注并复核",
    "P4": "常规跟踪",
}

# 两份报表比较时的变化阈值。
CHANGE_CONFIG = {
    "SIGNIFICANT_COUNT_CHANGE": 5,
    "SIGNIFICANT_PERCENT_CHANGE": 0.20,
}

OUTPUT_FILE_NAME = "视频异常巡检V8_双报表分析与对比.xlsx"
TEMP_FILE_NAME = "视频异常巡检V8_双报表分析与对比.tmp.xlsx"
TOP_FOCUS_COUNT = 30
TOP100_COUNT = 100
