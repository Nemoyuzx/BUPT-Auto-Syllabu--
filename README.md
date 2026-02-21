# BUPT-Auto-Syllabu

自动获取北邮课表并生成日历文件（ICS）与 16 周可视化课表图。

## Fork 说明

- 官方仓库：<https://github.com/LAWTED/BUPT-Auto-Syllabu>
- 当前维护仓库（本仓库）：<https://github.com/Nemoyuzx/BUPT-Auto-Syllabu-->
- 本仓库在官方基础上新增了：
	- `config.py` 读取账号与学期配置
	- 学期起始日期精确映射（支持指定 `term_start_date`）
	- 16 周图表导出（CSV / Markdown / PNG）
	- 所有生成产物统一输出到 `output/` 目录

## 环境要求

- Python 3.9+
- 校园网环境（或可访问教务系统）

安装依赖：

```bash
pip install -r requirements.txt
```

如果要生成课表图片：

```bash
pip install matplotlib chinese-calendar
```

## 新版使用方法

### 1) 配置账号与学期参数

编辑 `config.py`：

```python
account = "你的学号"
password = "你的教务密码"

xueqi = "2025-2026-2"
year = "2026"
term_start_date = "2026-03-02"   # 学期第1周周一

Combine_Trigger = True
show_week_mapping = True
output_dir = "output"
```

### 2) 生成课表 ICS 与文本导入链接

```bash
python process.py
```

运行后会在 `output/` 生成：

- `calendar.ics`
- `direct.txt`
- `semester_16week_chart.md`
- `semester_16week_chart.csv`
- `fetched_kb.xls`（调试用抓取快照）

### 3) 生成 16 周可视化课表图

```bash
python generate_weekly_image.py
```

输出文件：

- `output/semester_16week_vertical.png`

## 导入说明

### macOS / iOS

- 直接打开 `output/calendar.ics` 导入系统日历。

### 其他设备（兼容旧流程）

- 可使用 `output/direct.txt`（data URI）方式手动导入。

## 备注

- `config.py` 与 `output/` 已加入 `.gitignore`，默认不会被提交。
- 若课程日期偏移，优先检查 `config.py` 中的 `term_start_date`。
