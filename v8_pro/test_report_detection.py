# -*- coding: utf-8 -*-

from pathlib import Path
import pandas as pd

from main import detect_report_type, report_type_name


TEST_FILE = Path("test_data/unknown_test.xlsx")


def main():
    print("=" * 70)
    print("V1.1 Excel 报表类型自动识别测试")
    print("=" * 70)

    if not TEST_FILE.exists():
        print(f"测试文件不存在：{TEST_FILE}")
        return

    print(f"测试文件：{TEST_FILE}")

    df = pd.read_excel(TEST_FILE)

    print("\n检测到的 Excel 字段：")
    for column in df.columns:
        print(f"  - {column}")

    report_type = detect_report_type(df, TEST_FILE)

    print("\n===== 自动识别结果 =====")
    print(f"内部类型：{report_type}")
    print(f"报表类型：{report_type_name(report_type)}")

    if report_type == "mileage_summary":
        print("\n✅ 测试通过！")
        print("程序已经成功识别：车辆运行统计报表。")

    elif report_type == "video_anomaly":
        print("\n❌ 测试失败！")
        print("程序错误地将该文件识别成了视频异常明细报表。")

    elif report_type == "unknown":
        print("\n✅ 测试通过！")
        print("程序已经成功识别：未知报表。")

    else:
        print("\n❌ 测试失败！")
        print("程序返回了未定义的报表类型。")

    print("=" * 70)


if __name__ == "__main__":
    main()