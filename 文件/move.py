import os
import re
import zipfile
import pandas as pd

# 匹配规则：文件名中按序出现 r e s u l t（中间可有任意字符），不区分大小写
RESULT_PATTERN = re.compile(r"r.*?e.*?s.*?u.*?l.*?t", re.I)

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path

def move_file_overwrite(src, dst_dir):
    """移动文件到目标文件夹，如果存在同名文件覆盖并返回目标路径"""
    filename = os.path.basename(src)
    target = os.path.join(dst_dir, filename)
    if os.path.exists(target):
        try:
            os.remove(target)
        except Exception as e:
            print(f"[!] 无法删除已存在文件 {target}：{e}")
    try:
        os.replace(src, target)
        print(f"[+] 移动: {src} -> {target}")
    except PermissionError:
        print(f"[!] 权限错误，无法移动: {src} -> {target}")
    except Exception as e:
        print(f"[!] 移动失败 {src} -> {target}：{e}")
    return target

def extract_and_move_zip(zip_path, rsas_dir):
    """解压 ZIP：若内部有内嵌 ZIP 则提取内嵌 ZIP 到 rsas_dir 并删除源 ZIP；否则移动到 rsas_dir"""
    inner_zip_found = False
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            for name in zf.namelist():
                if name.lower().endswith(".zip"):
                    inner_zip_found = True
                    try:
                        extracted = zf.extract(name, rsas_dir)
                        print(f"[+] 提取内嵌ZIP: {extracted}")
                    except Exception as e:
                        print(f"[!] 提取内嵌ZIP失败 {name}：{e}")
    except zipfile.BadZipFile:
        print(f"[!] 非法或损坏的 ZIP 文件: {zip_path}")
        return

    if inner_zip_found:
        try:
            os.remove(zip_path)
            print(f"[+] 删除原 ZIP: {zip_path}")
        except Exception as e:
            print(f"[!] 无法删除原 ZIP {zip_path}：{e}")
    else:
        move_file_overwrite(zip_path, rsas_dir)

def read_csv_robust(path):
    """尝试使用多种编码读取 CSV，返回 DataFrame 或 None"""
    encodings = ("utf-8", "utf-8-sig", "gbk", "latin1")
    for enc in encodings:
        try:
            df = pd.read_csv(path, encoding=enc)
            print(f"[+] 以编码 {enc} 读取 CSV: {path}")
            return df
        except Exception:
            continue
    try:
        df = pd.read_csv(path, engine='python')
        print(f"[+] 使用 engine=python 读取 CSV: {path}")
        return df
    except Exception as e:
        print(f"[!] 无法读取 CSV 文件 {path}：{e}")
        return None

def merge_csv_files(csv_files, output_file):
    """
    合并 CSV 并去重（整行去重），列不一致时自动补空
    最终保存为 Excel (xlsx)，并删除源 CSV 文件
    """
    if not csv_files:
        print("[*] 没有 CSV 文件需要合并")
        return
    dfs = []
    for f in csv_files:
        df = read_csv_robust(f)
        if df is None:
            print(f"[!] 跳过无法读取的文件: {f}")
            continue
        dfs.append(df)
    if not dfs:
        print("[*] 没有可用的 DataFrame 可合并")
        return

    combined = pd.concat(dfs, ignore_index=True, sort=False)
    before = len(combined)
    combined.drop_duplicates(inplace=True)
    after = len(combined)
    print(f"[+] 合并完成：合并前行数={before}, 去重后行数={after}")

    try:
        combined.to_excel(output_file, index=False, engine="openpyxl")
        print(f"[+] 已写出 results -> {output_file}")
        # 删除源 CSV
        for f in csv_files:
            try:
                os.remove(f)
                print(f"[+] 已删除源文件: {f}")
            except Exception as e:
                print(f"[!] 删除源文件失败 {f}：{e}")
    except Exception as e:
        print(f"[!] 写出文件失败 {output_file}：{e}")

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = ensure_dir(os.path.join(base_dir, "..", "整理结果"))

    dirs = {
        "rsas": ensure_dir(os.path.join(output_dir, "RSAS")),
        "awvs": ensure_dir(os.path.join(output_dir, "awvs")),
        "nessus": ensure_dir(os.path.join(output_dir, "nessus")),
        "nmap": ensure_dir(os.path.join(output_dir, "nmap")),
        "整理结果": ensure_dir(os.path.join(output_dir, "整理结果"))
    }

    html_files = []
    csv_files_to_merge = []

    for fname in os.listdir(base_dir):
        fpath = os.path.join(base_dir, fname)
        if not os.path.isfile(fpath):
            continue
        lower_name = fname.lower()

        # ZIP 文件
        if lower_name.endswith(".zip"):
            extract_and_move_zip(fpath, dirs["rsas"])
            continue

        # HTML 文件
        if lower_name.endswith(".html"):
            if "affected" in lower_name:
                move_file_overwrite(fpath, dirs["awvs"])
            else:
                html_files.append(fpath)
            continue

        # CSV 文件
        if lower_name.endswith(".csv"):
            if RESULT_PATTERN.search(lower_name):  # 匹配 result
                moved = move_file_overwrite(fpath, dirs["整理结果"])
                csv_files_to_merge.append(moved)
            else:
                move_file_overwrite(fpath, dirs["nessus"])
            continue

        # XML 文件
        if lower_name.endswith(".xml"):
            move_file_overwrite(fpath, dirs["nmap"])
            continue

    # 处理 HTML 剩余文件
    if len(html_files) == 1:
        move_file_overwrite(html_files[0], dirs["整理结果"])
    elif len(html_files) > 1:
        zip_target = os.path.join(dirs["整理结果"], "nessus.zip")
        if os.path.exists(zip_target):
            try:
                os.remove(zip_target)
            except Exception as e:
                print(f"[!] 无法删除已有压缩包 {zip_target}：{e}")
        try:
            with zipfile.ZipFile(zip_target, 'w', zipfile.ZIP_DEFLATED) as zf:
                for h in html_files:
                    zf.write(h, os.path.basename(h))
                    try:
                        os.remove(h)
                    except Exception:
                        pass
            print(f"[+] 多个HTML已压缩 -> {zip_target}")
        except Exception as e:
            print(f"[!] 创建 HTML 压缩包失败：{e}")

    # 合并 CSV 文件为 result.xlsx 并去重
    if csv_files_to_merge:
        result_file = os.path.join(dirs["整理结果"], "result.xlsx")
        merge_csv_files(csv_files_to_merge, result_file)
    else:
        print("[*] 未发现匹配 result 的 CSV 文件，未生成 result.xlsx")

if __name__ == "__main__":
    main()

'''

zhkali127

'''