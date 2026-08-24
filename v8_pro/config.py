# -*- coding: utf-8 -*-
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"

INPUT_A = DATA_DIR / "data_A.xlsx"
INPUT_B = DATA_DIR / "data_B.xlsx"
AUTO_DISCOVER_TWO_INPUTS = True

OUTPUT_FILE_NAME = "车辆运营与报警月度报告_V8_最终版.xlsx"
TEMP_FILE_NAME = "~车辆运营与报警月度报告_V8_临时.xlsx"
WORD_REPORT_FILE = "车辆运营与报警月度报告_V8_最终版.docx"

TREND_EPSILON = 0.03

ALARM_TYPE_RULES = [
    ("车距过近", ["车距过近", "车距过小", "跟车过近"]),
    ("疲劳驾驶", ["疲劳驾驶", "疲劳"]),
    ("驾驶员分心", ["驾驶员分心", "分心驾驶", "驾驶员分心驾驶"]),
    ("无驾驶员", ["无驾驶员"]),
    ("违规打电话", ["驾驶员打电话", "违规打电话", "打电话"]),
    ("违规抽烟", ["驾驶员抽烟", "违规抽烟", "抽烟"]),
    ("碰撞报警", ["前车碰撞", "碰撞报警", "前向碰撞", "碰撞预警"]),
    ("左右盲区报警", ["左右盲区", "左盲区", "右盲区", "盲区报警", "盲区检测"]),
    ("驾驶员打哈欠", ["驾驶员打哈欠", "打哈欠", "哈欠"]),
]

UNIT_RULES = {
    "新华都": ["HD", "新华都"],
    "兴万祥": ["WX", "兴万祥"],
    "富达": ["FD", "富达"],
}

TOP_LIMITS = {
    "车距过近": 5,
    "疲劳驾驶": 10,
    "打电话和抽烟": 5,
    "碰撞报警": 5,
    "左右盲区报警": 5,
    "驾驶员分心": 10,
}

# V8 视频异常巡检兼容配置
COLUMN_ALIASES = {
    "状态名称": ["状态名称", "异常名称", "异常类型", "通道名称"],
    "状态类型": ["状态类型", "系统类型", "设备类型", "类型"],
    "设备编号": ["设备编号", "车牌号码", "车牌", "车辆编号", "车号"],
    "归属车组": ["归属车组", "归属分组", "车组", "班组", "所属车组", "所属分组"],
    "状态内容": ["状态内容", "报警内容", "内容", "描述", "异常描述"],
    "通道号": ["通道号", "通道", "channel", "CH"],
}

REQUIRED_COLUMNS = ["状态名称", "状态类型", "设备编号", "归属车组", "状态内容"]

CAMERA_ERROR_NAMES = ["摄像头异常", "视频丢失", "视频异常", "视频丢帧", "视频中断"]

CORE_SYSTEMS = ["DMS", "ADAS", "DSC"]
CORE_SYSTEM_LAYER = "核心系统"
ORDINARY_SYSTEM_LAYER = "普通摄像头"
REVIEW_SYSTEM_LAYER = "待确认"

SYSTEM_WEIGHTS = {"DMS": 1.00, "ADAS": 0.95, "DSC": 0.90, "普通摄像头": 0.30, "未知": 0.10}
REVIEW_SYSTEMS = ["未知", "未配置"]

GROUP_CHANNEL_RULES = []
DEFAULT_CHANNEL_RULES = {}

PRIORITY_CONFIG = {
    "P1_MAX_SHARE": 0.05, "P2_MAX_SHARE": 0.15, "P3_MAX_SHARE": 0.40,
    "P1_MIN_COUNT": 5, "P2_MIN_COUNT": 3,
    "WEIGHT_COUNT_RANK": 0.20, "WEIGHT_COUNT_ABSOLUTE": 0.20,
    "WEIGHT_GROUP_IMPACT": 0.15, "WEIGHT_SYSTEM_IMPACT": 0.15,
    "WEIGHT_SPREAD": 0.10, "WEIGHT_SYSTEM_IMPORTANCE": 0.10,
    "WEIGHT_CORE_SYSTEM": 0.10,
}
PRIORITY_ORDER = {"P1": 1, "P2": 2, "P3": 3, "P4": 4}
PRIORITY_ADVICE = {
    "P1": "优先处理，建议立即核查设备、线路及通道状态",
    "P2": "重点处理，建议安排现场或远程核查",
    "P3": "持续跟踪，结合后续巡检结果处理",
    "P4": "常规关注，纳入日常巡检",
}
CHANGE_CONFIG = {"SIGNIFICANT_COUNT_CHANGE": 5, "SIGNIFICANT_PERCENT_CHANGE": 0.30}
TOP_FOCUS_COUNT = 20
TOP100_COUNT = 100
