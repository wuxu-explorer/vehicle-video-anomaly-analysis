from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "data" / "test_data.xlsx"
OUTPUT_DIR = BASE_DIR / "output"

REQUIRED_COLUMNS = [
    "状态名称", "状态类型", "设备编号", "归属车组", "状态内容"
]

CAMERA_ERROR_NAMES = {"摄像头异常"}

# 已确认的车组专属规则；新规则只需要在这里增加。
GROUP_CHANNEL_RULES = [
    {"name": "HD运输九班", "pattern": "HD运输九班",
     "channels": {5: "DMS", 6: "ADAS", 7: "DSC"}},
    {"name": "集为辅助设备", "pattern": "集为辅助设备",
     "channels": {2: "ADAS", 5: "DMS", 6: "DSC"}},
    {"name": "建设宽体车", "pattern": "建设宽体车",
     "channels": {5: "DMS", 6: "ADAS", 7: "DSC"}},
]

DEFAULT_CHANNEL_RULES = {
    1: "原车倒车", 2: "ADAS", 3: "未接摄像机", 4: "未接摄像机",
    5: "DMS", 6: "DSC", 7: "盲区摄像头", 8: "盲区摄像头",
    9: "未接摄像机", 10: "盲区摄像头", 11: "盲区摄像头",
    12: "前摄像机",
}

CORE_SYSTEMS = {"DMS", "ADAS", "DSC"}
SYSTEM_WEIGHTS = {
    "DMS": 1.00, "ADAS": 1.00, "DSC": 1.00,
    "原车倒车": 0.70, "盲区摄像头": 0.70, "前摄像机": 0.70,
    "未接摄像机": 0.30, "未配置": 0.15, "未知": 0.10,
}

# V4：P1 不再单纯按前10%切。
# 同时满足风险排名和最低异常次数门槛；若数据整体很少，门槛会自动放宽。
PRIORITY_CONFIG = {
    "P1_MAX_SHARE": 0.10,
    "P2_MAX_SHARE": 0.30,
    "P3_MAX_SHARE": 0.60,
    "P1_MIN_COUNT": 10,
    "P2_MIN_COUNT": 5,
    "P3_MIN_COUNT": 2,
}

PRIORITY_ORDER = {"P1": 1, "P2": 2, "P3": 3, "P4": 4}

# 这些类型不会因为排名高就自动认定为真实系统故障。
REVIEW_SYSTEMS = {"未知", "未配置", "未接摄像机"}
