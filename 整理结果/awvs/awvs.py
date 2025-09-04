import os
import pandas as pd
import shutil
import zipfile

# 当前文件夹路径
current_folder = os.path.dirname(os.path.abspath(__file__))

# 输入文件路径
input_file = os.path.join(current_folder, "AwvsReport.xlsx")

# 输出文件夹路径（上级目录的整理结果）
output_folder = os.path.join(current_folder, "..", "整理结果")
os.makedirs(output_folder, exist_ok=True)

# 输出文件路径
output_file = os.path.join(output_folder, "web漏洞汇总表.xlsx")

# 读取Excel文件
df = pd.read_excel(input_file)

# 保留风险等级为中、高的行
filtered_df = df[df['风险等级'].isin(['中', '高'])]

# 如果没有数据 -> 只保留表头，并新增一行提示
if filtered_df.empty:
    filtered_df = pd.DataFrame(columns=df.columns)
    filtered_df = pd.concat([filtered_df, pd.DataFrame([{ '风险等级': '暂未发现中高危风险' }])],
                            ignore_index=True)

# 保存处理后的Excel
filtered_df.to_excel(output_file, index=False)

print(f"处理完成，文件已保存至: {output_file}")

# ----------------- HTML 文件处理 -----------------
html_files = [f for f in os.listdir(current_folder) if f.lower().endswith(".html")]

if len(html_files) > 1:
    # 压缩所有 html 文件
    zip_path = os.path.join(output_folder, "awvs.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for html_file in html_files:
            html_path = os.path.join(current_folder, html_file)
            zipf.write(html_path, arcname=html_file)
    print(f"已将多个 HTML 文件打包为: {zip_path}")

elif len(html_files) == 1:
    # 复制并重命名
    src_file = os.path.join(current_folder, html_files[0])
    dst_file = os.path.join(output_folder, "awvs.html")
    shutil.copy2(src_file, dst_file)
    print(f"已将单个 HTML 文件复制为: {dst_file}")

else:
    print("未发现 HTML 文件。")
