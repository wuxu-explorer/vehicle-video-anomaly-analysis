# 车辆视频异常数据分析系统

一个基于 Python、Pandas、OpenPyXL 的车辆视频异常数据分析工具，用于从 Excel 巡检数据中识别视频异常、判断 DMS / ADAS / DSC 等系统类型，并生成统计分析结果。

> **GitHub 安全说明**
>
> 本仓库整理版不包含原项目中的 `.venv`、真实业务 Excel 数据和运行生成的报告文件。
> 如果该项目来自公司/实习业务，请在公开仓库前确认你拥有发布相关源代码的权限，并确保不包含公司内部数据、接口、账号或其他保密信息。

## 版本

- **V7**：单报表视频异常分析。
- **V8 Lite**：精简版双报表分析。
- **V8 Pro**：较完整的双报表分析与对比版本。
- **legacy**：历史测试版本，仅用于查看开发过程，不保证直接运行。

## 主要功能

- Excel 视频异常数据读取与清洗
- 根据车组 + 通道号识别 DMS / ADAS / DSC 等系统
- 支持车组专属通道规则与通用规则
- 异常次数、车组影响、系统影响等维度的风险/优先级分析
- 双报表数据对比（V8）
- 生成 Excel 分析结果
- 配置与分析逻辑分离，便于后续扩展

## 环境

建议 Python 3.10+。

安装依赖：

```bash
pip install -r requirements.txt
```

## 使用方式

### V7

将待分析的 Excel 文件放入：

```text
v7/data/
```

然后运行：

```bash
python v7/main.py
```

### V8 Lite

将两个待比较的 Excel 文件放入：

```text
v8_lite/data/
```

然后运行：

```bash
python v8_lite/main.py
```

### V8 Pro

将两个待比较的 Excel 文件放入：

```text
v8_pro/data/
```

然后运行：

```bash
python v8_pro/main.py
```

> 当前整理版刻意没有上传原始业务 Excel。不同版本的输入文件命名和字段要求仍以各自代码中的配置为准。

## 项目结构

```text
vehicle-video-anomaly-analysis/
├── v7/
├── v8_lite/
├── v8_pro/
├── legacy/
├── demo_data/
├── tests/
├── requirements.txt
├── .gitignore
└── README.md
```

## 后续优化方向

1. 自动识别 Excel 文件及报表日期，减少手动改文件名。
2. 增加自动化测试。
3. 把 V7/V8 的重复代码进一步模块化。
4. 增加脱敏示例数据。
5. 完善异常规则配置与规则校验。
6. 使用 Git 分支和 Pull Request 管理版本。
