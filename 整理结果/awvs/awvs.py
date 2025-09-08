#!/usr/bin/env python3
import sys
from pathlib import Path
import pandas as pd
import shutil
import zipfile

# ---------- 目录定位 ----------
try:
    current_folder = Path(__file__).resolve().parent
except NameError:
    current_folder = Path.cwd()

# 先查找 HTML 文件（决定是否保存输出）
html_files = [p for p in current_folder.iterdir() if p.is_file() and p.suffix.lower() == ".html"]

if not html_files:
    print("未发现 HTML 文件，按要求不保存任何输出文件。脚本结束。")
    sys.exit(0)

# 如果有 HTML，再继续处理 Excel 与输出
input_file = current_folder / "AwvsReport.xlsx"
output_folder = (current_folder / ".." / "整理结果").resolve()
output_file = output_folder / "web漏洞汇总表.xlsx"

TARGET_RISKS = {"中", "高"}

def read_excel_safe(path: Path):
    if not path.exists():
        raise FileNotFoundError(f"输入文件未找到: {path}")
    try:
        return pd.read_excel(path, engine="openpyxl")
    except Exception as e:
        raise RuntimeError(f"读取 Excel 失败: {e}")

def ensure_col(df, col_name):
    if col_name not in df.columns:
        raise KeyError(f"Excel 中缺少必要列: '{col_name}'. 现有列: {list(df.columns)}")

def write_with_note(df, out_path: Path, note_text: str, note_col: str = None):
    if df.empty:
        cols = list(df.columns) or ["说明"]
        note_col = note_col or cols[0]
        row = {c: "" for c in cols}
        row[note_col] = note_text
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_excel(out_path, index=False, engine="openpyxl")

# 主流程
try:
    print(f"发现 HTML 文件数量: {len(html_files)}，开始处理 Excel 与输出...")
    df = read_excel_safe(input_file)

    risk_col = "风险等级"
    ensure_col(df, risk_col)

    risk_series = df[risk_col].astype(str).str.strip()
    filtered_df = df[risk_series.isin(TARGET_RISKS)].copy()

    # 创建输出文件夹（既然要保存，才创建文件夹）
    output_folder.mkdir(parents=True, exist_ok=True)

    if filtered_df.empty:
        note = "暂未发现中高危风险"
        write_with_note(pd.DataFrame(columns=df.columns), output_file, note_text=note)
        print(f"未发现中/高风险，已在 {output_file} 写入提示行。")
    else:
        write_with_note(filtered_df, output_file, note_text="")
        print(f"已导出 {len(filtered_df)} 条中/高风险记录 到: {output_file}")

    # HTML 处理（由于至少存在一个 html，这里会进行压缩或复制）
    if len(html_files) > 1:
        zip_path = output_folder / "awvs.zip"
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for html_path in html_files:
                zipf.write(html_path, arcname=html_path.name)
        print(f"已将多个 HTML 文件打包为: {zip_path}")
    else:
        src_file = html_files[0]
        dst_file = output_folder / "awvs.html"
        shutil.copy2(src_file, dst_file)
        print(f"已将单个 HTML 文件复制为: {dst_file}")

except FileNotFoundError as e:
    print("错误:", e)
    sys.exit(2)
except KeyError as e:
    print("错误:", e)
    sys.exit(3)
except Exception as e:
    print("发生未处理的异常:", str(e))
    sys.exit(1)
