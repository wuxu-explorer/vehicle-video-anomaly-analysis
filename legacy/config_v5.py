# -*- coding: utf-8 -*-
"""
视频异常巡检 - 动态优先级配置
说明：
1. main.py 负责计算，不要在 main.py 里硬编码优先级规则。
2. 如果以后新增/调整车组通道，只修改本文件的 GROUP_CHANNEL_RULES。
3. 优先级不是固定“次数>=100就是P1”，而是结合当前报表的相对排名、
   异常次数、车组影响、系统影响、规则确认状态动态计算。
"""

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "data" / "test_data.xlsx"
OUTPUT_DIR = BASE_DIR / "output"

# 原始异常报表至少需要这些字段。
REQUIRED_COLUMNS = [
    "状态名称", "状态类型", "设备编号", "归属车组", "状态内容"
]

# 识别为摄像头异常的状态名称。
CAMERA_ERROR_NAMES = {
    "摄像头异常",
    "视频异常",
    "视频丢失",
}

# 车组专属通道规则。
# 新报表如果车组规则发生变化，只需要在这里增加/修改。
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

# 没有命中特殊车组时使用。
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

# 只作为“系统重要性”的辅助因子，不会单独决定P1。
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

# 规则无法确认时，不直接把它当成真实核心系统故障。
REVIEW_SYSTEMS = {"未知", "未配置", "未接摄像机"}

# 动态优先级参数。
# 这些是“权重”，不是固定次数阈值，因此下一份报表数据变化后，
# P1/P2/P3 会随着当前数据的分布自动变化。
PRIORITY_CONFIG = {
    "P1_MAX_SHARE": 0.10,
    "P2_MAX_SHARE": 0.30,
    "P3_MAX_SHARE": 0.60,

    # 极少量数据时仍允许出现P1/P2/P3；这里是最低异常次数保护线。
    "P1_MIN_COUNT": 3,
    "P2_MIN_COUNT": 2,
    "P3_MIN_COUNT": 1,

    # 风险分组成：
    # 当前通道异常次数排名、绝对次数、车组影响、系统影响、
    # 车组异常扩散、系统重要性、核心系统。
    "WEIGHT_COUNT_RANK": 0.35,
    "WEIGHT_COUNT_ABSOLUTE": 0.20,
    "WEIGHT_GROUP_IMPACT": 0.15,
    "WEIGHT_SYSTEM_IMPACT": 0.10,
    "WEIGHT_SPREAD": 0.07,
    "WEIGHT_SYSTEM_IMPORTANCE": 0.08,
    "WEIGHT_CORE_SYSTEM": 0.05,
}

PRIORITY_ORDER = {
    "P1": 1,
    "P2": 2,
    "P3": 3,
    "P4": 4,
}

PRIORITY_ADVICE = {
    "P1": "立即重点处理",
    "P2": "优先处理",
    "P3": "持续关注并复核",
    "P4": "常规跟踪",
}
