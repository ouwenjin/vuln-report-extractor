import argparse
import re
from pathlib import Path
from bs4 import BeautifulSoup
import pandas as pd
import logging
from datetime import datetime
from typing import List, Dict
import sys
import os
import unicodedata
import textwrap

# 日志配置
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# 颜色定义
ANSI = {
    'reset': "\033[0m",
    'bold': "\033[1m",
    'cyan': "\033[36m",
    'magenta': "\033[35m",
    'green': "\033[32m",
    'yellow': "\033[33m",
}

AUTHOR = 'zhkali'
REPOS = [
    'https://github.com/ouwenjin/awvs-report-extractor',
    'https://gitee.com/zhkali/awvs-report-extractor'
]

# 用于去除 ANSI 控制码的正则
_ansi_re = re.compile(r'\x1B\[[0-?]*[ -/]*[@-~]')

def supports_color() -> bool:
    """
    简单检测终端是否支持 ANSI 颜色（Windows 上做了基础兼容判断）
    """
    if sys.platform.startswith('win'):
        return os.getenv('ANSICON') is not None or 'WT_SESSION' in os.environ or sys.stdout.isatty()
    return sys.stdout.isatty()

_COLOR = supports_color()

def strip_ansi(s: str) -> str:
    """去掉 ANSI 控制码，用于准确计算可见长度"""
    return _ansi_re.sub('', s)

def visible_width(s: str) -> int:
    """
    计算字符串在终端中的可见列宽（考虑中文等宽字符为 2 列、组合字符为 0 列）。
    传入字符串可以包含 ANSI 码，函数会先移除 ANSI 码再计算。
    """
    s2 = strip_ansi(s)
    w = 0
    for ch in s2:
        # 跳过不可见的组合字符（比如重音组合符）
        if unicodedata.combining(ch):
            continue
        ea = unicodedata.east_asian_width(ch)
        # 'F' (Fullwidth), 'W' (Wide) 视作 2 列；其余视作 1 列
        if ea in ('F', 'W'):
            w += 2
        else:
            w += 1
    return w

def pad_visible(s: str, target_visible_len: int) -> str:
    """
    在带颜色字符串的右侧补空格，使其可见长度达到 target_visible_len。
    空格为普通 ASCII 空格（宽度 1）。
    """
    cur = visible_width(s)
    if cur >= target_visible_len:
        return s
    return s + ' ' * (target_visible_len - cur)

def make_lines():
    """返回未着色的行（保留艺术字的前导空格）"""
    big_name = r"""
   ███████╗██╗  ██╗██╗  ██╗ █████╗ ██╗      ██╗        
   ╚══███╔╝██║  ██║██║ ██╔╝██╔══██╗██║      ██║        
     ███╔╝ ███████║█████╔╝ ███████║██║      ██║        
    ███╔╝  ██╔══██║██╔═██╗ ██╔══██║██║      ██║        
   ███████╗██║  ██║██║  ██╗██║  ██║███████╗ ██║       
   ╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝ ╚═╝        
"""
    art = textwrap.dedent(big_name)
    art_lines = [ln.rstrip('\n') for ln in art.splitlines() if ln != '']
    author_line = f"作者： {AUTHOR}"
    repo1 = REPOS[0]
    repo2 = REPOS[1]
    return art_lines + [''] + [author_line, repo1, repo2]

def print_banner(use_unicode: bool = True, outer_margin: int = 0, inner_pad: int = 1):
    # 选择字符集
    if use_unicode:
        tl, tr, bl, br, hor, ver = '┌','┐','└','┘','─','│'
    else:
        tl, tr, bl, br, hor, ver = '+','+','+','+','-','|'

    c_reset = ANSI.get('reset','')
    c_bold = ANSI.get('bold','')
    c_cyan = ANSI.get('cyan','')
    c_green = ANSI.get('green','')
    c_yellow = ANSI.get('yellow','')

    raw_lines = make_lines()

    # 着色（仅在支持颜色的终端）
    colored = []
    for ln in raw_lines:
        if ln.startswith('作者'):
            colored.append((c_bold + c_green + ln + c_reset) if _COLOR else ln)
        elif ln.startswith('http'):
            colored.append((c_yellow + ln + c_reset) if _COLOR else ln)
        else:
            if ln.strip() == '':
                colored.append(ln)
            else:
                colored.append((c_bold + c_cyan + ln + c_reset) if _COLOR else ln)

    # 计算可见最大宽度（使用 visible_width 来正确处理中文宽度）
    content_max = max((visible_width(x) for x in colored), default=0)

    # 预先把每行（带颜色的）右侧填充到 content_max（保证每行实际可见宽度相同）
    padded_lines = [pad_visible(ln, content_max) for ln in colored]

    # line_content = inner_pad + padded_line + inner_pad
    total_inner = inner_pad * 2 + content_max
    width = total_inner + 2  # 两侧竖线占 2

    # 构造顶部与底部边框
    top = tl + (hor * (width - 2)) + tr
    bottom = bl + (hor * (width - 2)) + br

    pad = ' ' * max(0, outer_margin)

    # 打印顶部（统一颜色）
    if _COLOR and use_unicode:
        print(pad + (c_cyan + top + c_reset))
    else:
        print(pad + top)

    # 打印所有内容行（左对齐：艺术字本身的前导空格会保留）
    left_bar = (c_cyan + ver + c_reset) if _COLOR else ver
    right_bar = (c_cyan + ver + c_reset) if _COLOR else ver
    for pl in padded_lines:
        line_content = (' ' * inner_pad) + pl + (' ' * inner_pad)
        print(pad + left_bar + line_content + right_bar)

    # 打印底部
    if _COLOR and use_unicode:
        print(pad + (c_cyan + bottom + c_reset))
    else:
        print(pad + bottom)

# 输出列（与 AwvsReport.xlsx 保持一致）
COLS = ['风险目标','风险名称','风险地址','风险等级','风险描述','风险详细','风险请求','整改意见']

# 关键字映射（中英文变体）
KEYWORD_MAP = {
    '风险名称': ['alert group', 'alert', '警报组', '漏洞名称', '告警组', '警报', '漏洞'],
    '风险等级': ['severity', 'risk level', 'risk', '严重性', '风险等级', '等级'],
    '风险描述': ['description', '描述', '漏洞描述', '说明', '详情说明'],
    '整改意见': ['recommendations', 'recommendation', 'recommend', '建议', '修复建议', '整改建议'],
    '风险详细': ['details', 'detail', '详细', '详情', '漏洞详情', '更多信息'],
    '风险请求': ['request', 'requests', '请求', '请求报文', 'http请求', '请求/响应', '请求响应'],
    '风险地址': ['risk address', 'risk地址', '风险地址', '位置', '风险位置', '位置/路径', '风险地址/位置'],
}

TARGET_PATTERNS = [
    r'Scan of\s*(.+)',
    r'扫描目标[:：]?\s*(.+)',
    r'开始 URL[:：]?\s*(.+)',
    r'Start url[:：]?\s*(.+)',
    r'Scan target[:：]?\s*(.+)',
]

def normalize_key(s: str) -> str:
    if not s:
        return ""
    s2 = s.strip()
    s2 = re.sub(r'[\r\n\t]+', ' ', s2)
    s2 = re.sub(r'\s+', ' ', s2)
    return s2.lower()

def key_matches_column(key_text: str, column: str) -> bool:
    if not key_text:
        return False
    key_norm = normalize_key(key_text)
    candidates = KEYWORD_MAP.get(column, [])
    for cand in candidates:
        cand_norm = normalize_key(cand)
        if cand_norm in key_norm or key_norm in cand_norm:
            return True
    return False

def extract_text(node):
    if node is None:
        return ""
    return node.get_text(separator="\n", strip=True)

def extract_request_from_node(node):
    if node is None:
        return ""
    for tagname in ('code', 'pre', 'textarea'):
        found = node.find(tagname)
        if found:
            return found.get_text("\n", strip=True)
    txt = extract_text(node)
    if re.search(r'^(GET|POST|PUT|DELETE|HEAD)\s+', txt, flags=re.MULTILINE):
        return txt
    return ""

def find_target_from_document(soup: BeautifulSoup) -> str:
    for header_tag in soup.find_all(['h1','h2','h3','div','p','span','td']):
        txt = header_tag.get_text(" ", strip=True)
        for pat in TARGET_PATTERNS:
            m = re.search(pat, txt, flags=re.IGNORECASE)
            if m:
                if m.groups():
                    return m.group(1).strip()
                return txt.strip()
    maybe = soup.find(lambda el: el.name in ('td','th','div','span') and re.search(r'(Start url|开始 url|开始URL|开始 URL|Start URL)', el.get_text(), flags=re.IGNORECASE))
    if maybe:
        sib = maybe.find_next_sibling()
        if sib:
            return extract_text(sib)
    return ""

def is_affected_table(table_tag: BeautifulSoup) -> bool:
    txt = table_tag.get_text(" ", strip=True).lower()
    keywords = ['alert group','severity','description','recommend','details','警报','严重性','描述','建议','详情','修复建议','漏洞描述','漏洞']
    score = 0
    for kw in keywords:
        if kw in txt:
            score += 1
    return score >= 1

def parse_single_html(path: Path) -> List[Dict]:
    html = path.read_text(encoding='utf-8', errors='ignore')
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    current_target = find_target_from_document(soup)

    for table in soup.find_all('table'):
        if not is_affected_table(table):
            continue

        item = {c: "" for c in COLS}
        item['风险目标'] = current_target or ""

        first_td = table.find('td')
        if first_td:
            t0 = extract_text(first_td)
            if len(t0) < 60 and not re.search(r'(alert|severity|description|details|recommend|警报|严重|描述|详情|建议)', t0, flags=re.IGNORECASE):
                item['风险地址'] = t0

        for tr in table.find_all('tr'):
            cells = tr.find_all(['th','td'])
            if not cells:
                continue
            if len(cells) == 1:
                continue
            key_cell = cells[0]
            val_cell = cells[1] if len(cells) > 1 else None
            key_text = extract_text(key_cell)
            val_text = extract_text(val_cell) if val_cell else ""

            mapped = False
            for col in ['风险名称','风险等级','风险描述','整改意见','风险详细','风险请求','风险地址']:
                if key_matches_column(key_text, col):
                    if col == '风险请求':
                        req = extract_request_from_node(val_cell)
                        item[col] = req or val_text
                    else:
                        item[col] = val_text
                    mapped = True
                    break
            if mapped:
                continue

            if re.search(r'^(GET|POST|PUT|DELETE|HEAD)\s+', val_text, flags=re.MULTILINE):
                item['风险请求'] = val_text
                continue

            key_low = normalize_key(key_text)
            if 'severity' in key_low or '严重' in key_low:
                item['风险等级'] = val_text
            elif 'alert' in key_low or '警报' in key_low or '漏洞' in key_low:
                if not item['风险名称']:
                    item['风险名称'] = val_text
            elif 'description' in key_low or '描述' in key_low:
                item['风险描述'] = val_text
            elif 'recommend' in key_low or '建议' in key_low or '修复' in key_low:
                item['整改意见'] = val_text
            else:
                existing = item.get('风险详细','') or ''
                if existing:
                    existing += "\n\n"
                existing += f"{key_text}: {val_text}"
                item['风险详细'] = existing

        if not item['风险请求']:
            req = extract_request_from_node(table)
            if req:
                item['风险请求'] = req

        if not item['风险名称']:
            bold = table.find(lambda t: t.name in ('b','strong') and len(t.get_text(strip=True))>2)
            if bold:
                item['风险名称'] = extract_text(bold)

        rows.append(item)

    return rows

def parse_files(input_paths: List[Path]) -> pd.DataFrame:
    all_rows = []
    for p in input_paths:
        logging.info(f"解析文件: {p.name}")
        parsed = parse_single_html(p)
        logging.info(f"  在 {p.name} 中找到 {len(parsed)} 个受影响项")
        all_rows.extend(parsed)
    if not all_rows:
        logging.warning("未解析到任何受影响项。请检查输入是否为 Acunetix/AWVS Affected Items 报告。")
    df = pd.DataFrame(all_rows, columns=COLS)
    return df

def backup_if_exists(path: Path):
    if path.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        bk = path.with_name(path.stem + f"_backup_{ts}" + path.suffix)
        path.replace(bk)
        logging.info(f"已备份原文件到: {bk}")

def main():
    parser = argparse.ArgumentParser(description="自动解析 AWVS HTML 报告为 Excel（默认输出 web漏洞汇总表.xlsx），并打印作者横幅")
    parser.add_argument("inputs", nargs="*", help="要解析的 HTML 文件（可选，不指定时解析当前目录所有 html/htm）")
    parser.add_argument("-o", "--output", default="web漏洞汇总表.xlsx", help="输出 Excel 文件名，默认 web漏洞汇总表.xlsx")
    parser.add_argument('--no-unicode', dest='no_unicode', action='store_true',
                        help='强制使用 ASCII 框（不使用 Unicode 盒绘字符）')
    parser.add_argument('--margin', type=int, default=0, help='横幅左侧外边距空格数（默认 0）')
    parser.add_argument('--pad', type=int, default=1, help='横幅内部左右边距（默认 1）')
    args = parser.parse_args()

    # 打印横幅
    print_banner(use_unicode=not args.no_unicode, outer_margin=args.margin, inner_pad=max(0, args.pad))

    # 如果没有传入 inputs，则在当前目录查找所有 .html/.htm
    if not args.inputs:
        cwd = Path.cwd()
        htmls = sorted([p for p in cwd.glob("*.html")] + [p for p in cwd.glob("*.htm")])
        if not htmls:
            logging.error("当前目录没有找到任何 .html 或 .htm 文件。请把 HTML 报告放到当前目录或传入文件路径。")
            return
        input_paths = htmls
    else:
        input_paths = [Path(p) for p in args.inputs]

    # 过滤不存在或空文件
    input_paths = [p for p in input_paths if p.exists() and p.stat().st_size > 0]
    if not input_paths:
        logging.error("没有可用的输入文件（文件不存在或为空）。")
        return

    out_path = Path(args.output)
    # 先备份已存在的输出文件
    if out_path.exists():
        backup_if_exists(out_path)

    df = parse_files(input_paths)
    # 写入 Excel（若 DataFrame 为空，仍写入空文件以便后续检查）
    df.to_excel(out_path, index=False)
    logging.info(f"已保存 {len(df)} 行到 {out_path.resolve()}")

if __name__ == "__main__":
    main()