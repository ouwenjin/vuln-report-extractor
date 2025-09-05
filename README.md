# vuln-report-extractor — 漏洞报告解析与合并工具集


## 一、概述

`vuln-report-extractor` 是一套用于解析、标准化和合并多种安全扫描器（RSAS/绿盟、AWVS、Nessus、Nmap 等）输出的 Python 工具集合。目标是把不同来源的扫描结果转换为统一格式的可导出表格（Excel/CSV），并生成便于上报与复核的“中/高危”清单与端口调研表。

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE) [![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org)

---

## 二、仓库总体结构（建议且已修正的清晰版）

> 下面为建议的语义化目录结构（把代码模块与整理结果明确区分）。如果你当前仓库里的目录名不同，请按实际文件名替换命令中的路径或在脚本里修改常量。

nessus缺失示例补充:https://www.tenablecloud.cn/plugins

```
vuln-report-extractor/
├─ controller.py              # 可选：任务编排器（按顺序运行各模块）
├─ 整理结果/                   
│  ├─ RSAS/
│  │  └─ rsas.py              # 绿盟/RSAS HTML 或 Excel 报告解析与合并
│  ├─ awvs/
│  │  └─ awvs.py              # AWVS 报告（HTML/XLSX）解析、筛选、导出 web 漏洞汇总
│  ├─ nessus/
│  │  └─ nessus.py            # Nessus CSV 合并、字段映射、导出中高危
│  ├─ nmap/
│  │  └─ nmap.py              # Nmap XML 合并、端口调研表生成
│  ├─ 文件/
│  │  ├─ move.py              # ZIP 解压、文件移动/重命名/编码修正、读取工具
│  └─ └─ 输入文件
│    
├─ README.md (本文件)
├─ requirements.txt          
└─ LICENSE
```
---

## 三、每个模块（文件）功能详解、参数与示例

> 下列说明基于仓库中常见脚本。

### 1) `整理结果/RSAS/rsas.py`

**功能**
解析绿盟（RSAS）导出的 Excel/HTML 报表，标准化列名，抽取每条漏洞记录，导出整理后的报告与中/高危清单。

**期望输入**

* `combined_reports/漏洞报告.xlsx`（默认）或包含 RSAS HTML 报告的目录。

**输出**

* `漏洞报告_整理.xlsx`：标准化后的完整表。
* `中高危漏洞.xlsx`：仅包含风险等级为“中”或“高”的记录。

**命令行示例**

```bash
python 整理结果/RSAS/rsas.py --input combined_reports/漏洞报告.xlsx --output-dir ../整理结果/ --min-risk 中
```

**实现要点 / 建议**

* 使用列名候选映射（`COLUMN_CANDIDATES`）兼容不同模板。
* 风险等级识别应支持中文与英文（`高/中/低`、`Critical/High/Medium`）。
* 对空表或无中高危结果做容错并写入占位说明（便于 downstream 检查）。

**常见故障排查**

* 输出为空：确认输入文件内是否含漏洞记录且列名是否命中映射候选。

---

### 2) `整理结果/awvs/awvs.py`

**功能**
解析 AWVS 导出的 Excel 或 HTML 报告，合并 web 漏洞记录，筛选并导出 `web漏洞汇总表.xlsx`，可按 IP 汇总。

**期望输入**

* `awvs/AwvsReport.xlsx` 或 `awvs/html_reports/` 中的一组 HTML/XLSX 文件。

**输出**

* `web漏洞汇总表.xlsx`（并建议将其复制/移动到 `../整理结果/` 下）。

**命令行示例**

```bash
python 整理结果/awvs/awvs.py --input awvs/AwvsReport.xlsx --out ../整理结果/web漏洞汇总表.xlsx
```

**实现要点 / 建议**

* 解析字段：URL、漏洞名称、风险、证据、建议、影响范围、插件ID 等。
* 对 `Low/Medium/High` 做中文映射为 `低/中/高`。
* 合并同一 URL 的多条记录时保留最严重的风险等级并合并/去重证据文本。

**高级建议**

* 同时输出 `web漏洞按IP汇总.xlsx`（统计每主机漏洞数），便于快速查看重点目标。

---

### 3) `整理结果/nessus/nessus.py`

**功能**
合并多个 Nessus 导出的 CSV（或转换 `.nessus` XML 的 CSV），做字段标准化与增强（使用映射表），生成带时间戳的总表与中高危清单，并可导出 IP 列表。

**期望输入**

* 目录内的一组 `.csv`（Nessus 导出），或 `.nessus`（若脚本支持解析）。

**输出**

* `漏洞扫描结果-<timestamp>.xlsx`
* `中高危漏洞.xlsx`
* `ip.xlsx`（可选）

**命令行示例**

```bash
python 整理结果/nessus/nessus.py --src nessus/csvs/ --out ../整理结果/ --min-risk Medium
```

**实现要点 / 建议**

* 合并 CSV 时要处理常见编码问题（`utf-8`、`utf-8-sig`、`gbk`），脚本可按序尝试多种编码读取。
* 若存在 `Nessus中文报告.xlsx` 映射表，可用来补全插件描述、CVE 映射等信息，增强结果可读性。

---

### 4) `整理结果/nmap/nmap.py`

**功能**
解析并合并 Nmap 的 XML 输出（或读取预制的 `开放端口.xlsx`），生成 `端口调研表.xlsx`，支持按端口/服务/协议统计。

**期望输入**

* 多个 Nmap XML 文件或 `开放端口.xlsx`。

**输出**

* `端口调研表.xlsx`（包含：IP、端口、服务、协议、发现时间、来源文件等字段）。

**命令行示例**

```bash
python 结果整理/nmap/nmap.py --xml-dir nmap/xmls/ --out ../整理结果/端口调研表.xlsx
```

**实现要点 / 建议**

* 使用 `lxml` 或 `xml.etree.ElementTree` 解析 XML；`lxml` 更健壮且在复杂 XML 上表现更好。
* 支持合并同一 IP 多次扫描结果，并统计端口出现频次，便于判断常开端口与临时性开放端口。

---

### 5) `整理结果/utils/move.py`

**功能**
辅助脚本：递归解压 ZIP（包括内嵌 ZIP）、自动识别并按规则移动文件到对应模块目录、尝试修复文件名编码问题。

**常见用法**

```bash
python 结果整理/utils/move.py --src uploads/scan_zips/ --dest 结果整理/ --unpack
```

**实现要点 / 建议**

* 解压后识别典型文件名（如 `AwvsReport.xlsx`、`漏洞报告.xlsx`、`*.csv`、`*.xml`），并自动移动到 `结果整理/awvs/`、`结果整理/RSAS/`、`结果整理/nessus/`、`结果整理/nmap/` 等目录。
* 对文件名出现乱码时可通过 `chardet` 探测原始编码并重命名为 UTF-8 可读名（注意保留原始文件备份以免数据丢失）。

---

## 四、输入/输出字段定义（建议的标准字段）

为便于把不同扫描器的数据标准化，建议目标表（统一格式）包含下列字段（优先级由上到下）：

1. 序号
2. IP
3. 端口
4. 协议（TCP/UDP/…）
5. 漏洞名称
6. 风险等级（高/中/低 / Critical/High/Medium）
7. 漏洞说明（短摘要）
8. 加固建议（建议措施）
9. CVE（若有）
10. 证据 / Proof（抓取的样本证明或响应）
11. 来源文件（扫描器名 + 原始文件名）
12. 首次发现时间
13. 最后扫描时间

> 脚本应尽量填充这些字段，对不存在的字段留空或写 `N/A`，便于后续自动化处理与人工复核。

---

## 五、示例工作流（从原始 ZIP 到整理结果）

1. 把扫描器导出的 ZIP/CSV/XLS 等放入 `uploads/`（或任意你习惯的目录）。
2. （可选）运行 `move.py` 把压缩包内的扫描器输出分发到对应模块：

   ```bash
   python 结果整理/utils/move.py --src uploads/ --dest 结果整理/ --unpack
   ```
3. 逐个运行解析脚本（或使用 `controller.py` 编排执行）：

   ```bash
   # 单个模块处理（示例）
   python 整理结果/RSAS/rsas.py --input combined_reports/漏洞报告.xlsx --output-dir ../整理结果/
   python 整理结果/awvs/awvs.py --input awvs/AwvsReport.xlsx --out ../整理结果/web漏洞汇总表.xlsx
   python 整理结果/nessus/nessus.py --src nessus/csvs/ --out ../整理结果/
   python 整理结果/nmap/nmap.py --xml-dir nmap/xmls/ --out ../整理结果/端口调研表.xlsx

   # 一键编排（若 controller.py 支持）
   python controller.py --base-dir 结果整理/ --out ../整理结果/
   ```
4. 在 `整理结果/` 中复核输出文件（特别是 `中高危漏洞.xlsx` ），按需脱敏后上报/归档。

---

## 六、常见 Issue 与快速排查清单

1. **Unicode / 编码错误（CSV/文件名乱码）**

   * 尝试读取时使用 `encoding='utf-8-sig'` 或 `encoding='gbk'`。
   * 使用 `chardet` 探测编码并进行转换。
2. **Excel 文件被占用无法写入**

   * 关闭打开的 Excel 程序或以管理员权限运行脚本；在写入前删除目标文件或使用临时文件后替换。
3. **字段对不上（列名不同）**

   * 编辑脚本顶部的 `COLUMN_CANDIDATES` 或映射表，把实际列名加入候选列表以便自动匹配。
4. **脚本运行慢 / 内存占用高**

   * 对大 CSV 使用 `pandas.read_csv(..., chunksize=...)` 分块处理并增量写入。
5. **解压后目录出现乱码（例如 `╒√└φ╜ß╣√/`）**

   * 这是编码问题导致的文件名显示异常。建议重命名为 `结果整理/` 并在脚本中使用相对路径或 CLI 参数避免硬编码目录名。
6. **输出为空或少量数据**

   * 检查输入文件是否是真实扫描结果（而非空模板），并确认脚本的过滤阈值（例如 `--min-risk`）是否设置过高。
