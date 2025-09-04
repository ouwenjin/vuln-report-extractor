from __future__ import annotations
import os
import sys
import zipfile
import shutil
import tempfile
import traceback
import re
import json
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
import argparse
import unicodedata
import textwrap

# Optional deps
try:
    import pandas as pd
except Exception:
    pd = None

try:
    from bs4 import BeautifulSoup
except Exception:
    BeautifulSoup = None

try:
    from tqdm import tqdm
except Exception:
    def tqdm(x, **kw):
        return x  # 如果没安装 tqdm，则退化为普通迭代

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
    'https://github.com/ouwenjin/nsfocus-report-extractor',
    'https://gitee.com/zhkali/nsfocus-report-extractor'
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

# ----------------------------- Utils -----------------------------

def log(msg: str) -> None:
    print(msg)

def ensure_env():
    if pd is None:
        raise RuntimeError("pandas 未安装。请运行: pip install pandas openpyxl")
    if BeautifulSoup is None:
        raise RuntimeError("beautifulsoup4 未安装。请运行: pip install beautifulsoup4 lxml")

def unzip_all(zips: List[Path], tempdir: Path) -> List[Path]:
    out = []
    for z in zips:
        try:
            with zipfile.ZipFile(z, 'r') as zz:
                dst = tempdir / z.stem
                dst.mkdir(parents=True, exist_ok=True)
                zz.extractall(path=str(dst))
                out.append(dst)
                log(f"解压: {z} -> {dst}")
        except Exception as e:
            log(f"解压失败 {z}: {e}")
    return out

def find_files(base: Path, exts: Tuple[str, ...]) -> List[Path]:
    res = []
    for root, _, files in os.walk(base):
        for f in files:
            if f.lower().endswith(exts):
                res.append(Path(root) / f)
    return res

# ----------------------------- HTML 解析 -----------------------------

HEADER_MAP = {
    'ip': {'ip','主机','地址','资产地址','host','ip地址'},
    'port': {'端口','服务端口','port','端口号','协议/端口','协议端口','服务','protocol/port'},
    'name': {'漏洞名称','名称','漏洞','问题','title','name'},
    'risk': {'风险等级','风险','severity','威胁等级','等级'},
    'desc': {'漏洞说明','漏洞描述','描述','说明','description','detail','details'},
    'fix': {'加固建议','修复建议','解决办法','整改建议','处理建议','建议','修复措施','处置建议','remediation','solution'},
    'cve': {'cve','编号','cve编号','参考编号','reference','vul id'},
}

CVE_RE = re.compile(r'(CVE[-_:\s]*\d{4}[-_]\d{4,7})', re.I)
IP_RE = re.compile(r'((?:\d{1,3}\.){3}\d{1,3})')
PORT_RE = re.compile(r'(?:port|端口)[\s:：]*([0-9]{1,5})', re.I)
SEV_INLINE_RE = re.compile(r'(高危|中危|低危|严重|高|中|低|high|medium|low)', re.I)

def normalize_key(k: str) -> str:
    k = (k or '').strip().lower()
    for canon, aliases in HEADER_MAP.items():
        if k in {a.lower() for a in aliases}:
            return canon
    return k

def _clean_text_list_or_str(v: Any) -> str:
    """把 list 或者形如 "['x', 'y']" 的字符串清成纯文本。"""
    if isinstance(v, list):
        return "\n".join(str(x).strip().strip("[]'\"") for x in v if x is not None)
    s = str(v) if v is not None else ''
    return s.strip().strip("[]'\"")

def extract_from_table(table) -> Tuple[List[Dict[str, Any]], List[Tuple[str,str]]]:
    """从一个 <table> 同时提取漏洞记录与端口记录。返回 (vuln_rows, port_pairs)。"""
    headers = [th.get_text(strip=True) for th in table.find_all('th')]
    if not headers:
        first = table.find('tr')
        if first:
            headers = [td.get_text(strip=True) for td in first.find_all(['td','th'])]
            data_trs = table.find_all('tr')[1:]
        else:
            data_trs = table.find_all('tr')
    else:
        data_trs = table.find_all('tr')[1:]

    headers_norm = [normalize_key(h) for h in headers]

    vuln_rows: List[Dict[str, Any]] = []
    port_pairs: List[Tuple[str,str]] = []

    for tr in data_trs:
        cols = [td.get_text(" ", strip=True) for td in tr.find_all(['td','th'])]
        if not cols:
            continue
        row = {headers_norm[i] if i < len(headers_norm) else f'col{i}': cols[i] for i in range(len(cols))}
        # 端口行判定：含 IP+端口 且不含明显“漏洞名称/风险”等
        has_ip = any(k in row and IP_RE.search(str(row[k])) for k in row)
        has_port = any(k in row and re.search(r'^\d{1,5}$', str(row[k])) for k in row)
        has_vuln_keys = any(k in row for k in ('name','risk','desc','fix','cve'))

        if has_ip and has_port and not has_vuln_keys:
            ip = None; port = None
            for k, v in row.items():
                if isinstance(v, str) and IP_RE.search(v):
                    ip = IP_RE.search(v).group(1)
                if isinstance(v, str) and re.fullmatch(r'\d{1,5}', v):
                    port = v
            if ip and port:
                port_pairs.append((ip, port))
            continue

        # 作为漏洞行处理
        vr = {'IP':'', '端口':'', '漏洞名称':'', '风险等级':'', '漏洞说明':'', '加固建议':'', 'CVE':''}
        for k, v in row.items():
            if v is None:
                continue
            key = k
            txt = _clean_text_list_or_str(v)
            if key == 'ip' and not vr['IP']:
                m = IP_RE.search(txt)
                vr['IP'] = m.group(1) if m else txt
            elif key == 'port' and not vr['端口']:
                m = re.search(r'\d{1,5}', txt)
                vr['端口'] = m.group(0) if m else txt
            elif key == 'name' and not vr['漏洞名称']:
                vr['漏洞名称'] = txt
            elif key == 'risk' and not vr['风险等级']:
                m = SEV_INLINE_RE.search(txt)
                vr['风险等级'] = m.group(1) if m else txt
            elif key == 'desc' and not vr['漏洞说明']:
                vr['漏洞说明'] = txt
            elif key == 'fix' and not vr['加固建议']:
                vr['加固建议'] = txt
            elif key == 'cve' and not vr['CVE']:
                m = CVE_RE.search(txt)
                vr['CVE'] = m.group(1).upper().replace('_','-') if m else txt
        # 再从整行抓 CVE
        if not vr['CVE']:
            m = CVE_RE.search(" ".join(cols))
            if m:
                vr['CVE'] = m.group(1).upper().replace('_','-')
        # 只有存在“漏洞名称/描述/CVE/风险”等之一才算漏洞行
        if any(vr[k] for k in ('漏洞名称','漏洞说明','CVE','风险等级')):
            vuln_rows.append(vr)

    return vuln_rows, port_pairs

def extract_from_blocks(soup) -> Tuple[List[Dict[str, Any]], List[Tuple[str,str]]]:
    """从非表格的文本块里尽力提取漏洞与端口。"""
    vulns: List[Dict[str, Any]] = []
    ports: List[Tuple[str,str]] = []
    blocks = soup.find_all(['div','section','li','p'])
    for b in blocks:
        text = b.get_text(' ', strip=True)
        if len(text) < 20:
            continue
        # 端口对
        ip_m = IP_RE.search(text)
        port_m = PORT_RE.search(text)
        if ip_m and port_m:
            ports.append((ip_m.group(1), port_m.group(1)))
        # 漏洞
        cve_m = CVE_RE.search(text)
        sev_m = SEV_INLINE_RE.search(text)
        if cve_m or ('漏洞' in text) or sev_m:
            vr = {'IP':'', '端口':'', '漏洞名称':'', '风险等级':'', '漏洞说明':'', '加固建议':'', 'CVE':''}
            if ip_m:
                vr['IP'] = ip_m.group(1)
            if port_m:
                vr['端口'] = port_m.group(1)
            vr['CVE'] = cve_m.group(1).upper().replace('_','-') if cve_m else ''
            vr['风险等级'] = sev_m.group(1) if sev_m else ''
            # 名称/描述/修复建议的启发式
            label_map = {
                '漏洞名称': re.compile(r'漏洞名称|名称|问题|title', re.I),
                '漏洞说明': re.compile(r'漏洞说明|漏洞描述|描述|说明|description', re.I),
                '加固建议': re.compile(r'加固建议|修复建议|整改建议|解决办法|处理建议|solution|remediation', re.I),
            }
            for sib in b.next_siblings:
                if getattr(sib, 'get_text', None):
                    st = sib.get_text(' ', strip=True)
                    if not st:
                        continue
                    if not vr['漏洞名称'] and label_map['漏洞名称'].search(st):
                        vr['漏洞名称'] = st
                    if not vr['漏洞说明'] and label_map['漏洞说明'].search(st):
                        vr['漏洞说明'] = st
                    if not vr['加固建议'] and label_map['加固建议'].search(st):
                        vr['加固建议'] = _clean_text_list_or_str(st)
                    if vr['漏洞名称'] and vr['漏洞说明'] and vr['加固建议']:
                        break
            # 若仍为空，退化处理
            if not vr['漏洞名称']:
                vr['漏洞名称'] = text[:120]
            if not vr['漏洞说明']:
                vr['漏洞说明'] = text[:1500]
            vulns.append(vr)
    return vulns, ports

def _extract_js_object_by_marker(text: str, marker: str) -> Optional[str]:
    """Find JS object literal after marker (e.g. window.data). Return substring or None."""
    idx = text.find(marker)
    if idx == -1:
        return None
    brace_start = text.find('{', idx)
    if brace_start == -1:
        return None
    i = brace_start
    depth = 0
    while i < len(text):
        ch = text[i]
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                return text[brace_start:i+1]
        i += 1
    return None

def _walk_for_lists(obj: Any, keys: Tuple[str, ...]) -> List[Any]:
    found: List[Any] = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k in keys:
                found.append(v)
            else:
                found.extend(_walk_for_lists(v, keys))
    elif isinstance(obj, list):
        for it in obj:
            found.extend(_walk_for_lists(it, keys))
    return found

def normalize_risk(raw: Optional[str]) -> str:
    if not raw:
        return ''
    s = str(raw).strip().lower()
    if any(x in s for x in ('critical', '严重', '高危', 'high')):
        return '高'
    if any(x in s for x in ('中高', '中危', 'medium', '中')):
        return '中'
    if any(x in s for x in ('low', '低')):
        return '低'
    if any(x in s for x in ('info', '信息', 'informational')):
        return '信息'
    if re.search(r'\b(4|5)\b', s):
        return '高'
    return raw or ''

def normalize_record(d: Dict[str, Any]) -> Dict[str, Any]:
    """
    将任意解析结果字典映射为统一字段：
    '序号','IP','端口','漏洞名称','风险等级','漏洞说明','加固建议','CVE'
    ✅ 特别处理：加固建议字段可能是 list 或形如 ['...'] 的字符串，这里统一清理。
    """
    out = {
        '序号': '',
        'IP': '',
        '端口': '',
        '漏洞名称': '',
        '风险等级': '',
        '漏洞说明': '',
        '加固建议': '',
        'CVE': ''
    }
    if not isinstance(d, dict):
        return out

    for k, v in d.items():
        if v is None:
            continue
        key = str(k).strip().lower()
        # list 或字符串统一清洗
        val = _clean_text_list_or_str(v)

        if key in ('ip', 'host', 'ipaddress', '地址'):
            out['IP'] = val
        elif key in ('port', '端口'):
            m = re.search(r'(\d{1,5})', val)
            out['端口'] = m.group(1) if m else val
        elif key in ('name', 'title', '漏洞名称', 'vuln', '漏洞'):
            out['漏洞名称'] = val
        elif key in ('risk', '风险等级', 'severity'):
            out['风险等级'] = normalize_risk(val)
        elif key in ('description', '漏洞说明', 'desc'):
            out['漏洞说明'] = val
        elif key in ('recommendation', '建议', '加固建议', 'fix', 'solution'):
            out['加固建议'] = val
        elif 'cve' in key:
            m = CVE_RE.search(val)
            out['CVE'] = m.group(1).upper().replace('_', '-') if m else val
        else:
            # 额外探测
            m = CVE_RE.search(val)
            if m and not out['CVE']:
                out['CVE'] = m.group(1).upper().replace('_', '-')
            if not out['风险等级']:
                m2 = SEV_INLINE_RE.search(val)
                if m2:
                    out['风险等级'] = normalize_risk(m2.group(1))
            if len(val) > 10 and not out['漏洞说明']:
                out['漏洞说明'] = val

    # 保证非 None
    for k in out:
        if out[k] is None:
            out[k] = ''
    return out

def parse_html_file(path: Path) -> Tuple[List[Dict[str, Any]], List[Tuple[str,str]]]:
    """
    解析单个 HTML 文件：若嵌入 JSON 则优先解析；否则回退到表格/文本。
    返回 (vuln_records, port_pairs)
    """
    try:
        txt = path.read_text(encoding='utf-8', errors='ignore')
    except Exception:
        txt = path.read_text(encoding='gbk', errors='ignore')
    soup = BeautifulSoup(txt, 'lxml') if BeautifulSoup else None

    vulns_all: List[Dict[str, Any]] = []
    ports_all: List[Tuple[str,str]] = []

    # 1) 尝试解析嵌入 JSON（如 window.data）
    parsed = None
    js_text = _extract_js_object_by_marker(txt, "window.data")
    if js_text:
        try:
            parsed = json.loads(js_text)
        except Exception:
            try:
                cleaned = re.sub(r',\s*(?=[}\]])', '', js_text)
                parsed = json.loads(cleaned)
            except Exception:
                parsed = None

    if parsed is None and soup is not None:
        for script in soup.find_all("script"):
            st = script.string or script.get_text() or ""
            if not st:
                continue
            if any(k in st.lower() for k in ("vul_items", "vuls", "vul", "vulner")):
                # 寻找对象文本
                br_idx = st.find('{')
                if br_idx != -1:
                    cand = _extract_js_object_by_marker(st, "{")
                else:
                    cand = None
                if cand:
                    try:
                        parsed = json.loads(cand)
                        break
                    except Exception:
                        try:
                            cand2 = re.sub(r',\s*(?=[}\]])', '', cand)
                            parsed = json.loads(cand2)
                            break
                        except Exception:
                            parsed = None
                            continue

    if parsed is not None:
        keynames = ("vul_items", "vuls", "solve_items", "vul_info", "vulnerabilities", "vuls_list")
        lists = _walk_for_lists(parsed, keynames)
        for lst in lists:
            items: List[Any] = []
            if isinstance(lst, list):
                items = lst
            elif isinstance(lst, dict):
                for v in lst.values():
                    if isinstance(v, list):
                        items.extend(v)
            for entry in items:
                if not isinstance(entry, dict):
                    continue
                if 'vuls' in entry and isinstance(entry['vuls'], list):
                    port = entry.get('port') or entry.get('service_port') or ''
                    for subv in entry['vuls']:
                        vm = subv.get('vul_msg') or subv
                        ip = vm.get('host_ip') or vm.get('host') or entry.get('host_ip','') or ''
                        cve = vm.get('cve_id') or vm.get('cve') or vm.get('cncve','') or ''
                        name = vm.get('i18n_name') or vm.get('threat') or subv.get('i18n_name','') or ''
                        if isinstance(name, list):
                            name = name[0] if name else ''
                        # 描述与修复建议可能是 list
                        if isinstance(vm.get('i18n_description'), list):
                            desc = "\n".join(_clean_text_list_or_str(x) for x in vm.get('i18n_description'))
                        else:
                            desc = vm.get('i18n_description') or vm.get('exp_desc') or subv.get('exp_desc','') or ''
                        sol = vm.get('i18n_solution')
                        fix_txt = _clean_text_list_or_str(sol) if sol is not None else ''
                        level = subv.get('vul_level') or subv.get('threat_level') or vm.get('vul_level') or vm.get('threat_level','')
                        rec = {
                            'IP': ip or '',
                            '端口': str(port) if port else '',
                            '漏洞名称': _clean_text_list_or_str(name) if name else '',
                            '风险等级': normalize_risk(level),
                            '漏洞说明': _clean_text_list_or_str(desc) if desc else '',
                            '加固建议': fix_txt,
                            'CVE': _clean_text_list_or_str(cve) if cve else ''
                        }
                        vulns_all.append(rec)
                        if ip and port:
                            ports_all.append((ip, str(port)))
                else:
                    vm = entry.get('vul_msg') or entry
                    ip = vm.get('host_ip') or vm.get('host') or entry.get('host_ip','') or ''
                    port = entry.get('port') or vm.get('port') or ''
                    cve = vm.get('cve_id') or vm.get('cve') or vm.get('cncve','') or ''
                    name = vm.get('i18n_name') or vm.get('threat') or entry.get('i18n_name','') or ''
                    if isinstance(name, list):
                        name = name[0] if name else ''
                    if isinstance(vm.get('i18n_description'), list):
                        desc = "\n".join(_clean_text_list_or_str(x) for x in vm.get('i18n_description'))
                    else:
                        desc = vm.get('i18n_description') or vm.get('exp_desc') or entry.get('exp_desc','') or ''
                    sol = vm.get('i18n_solution')
                    fix_txt = _clean_text_list_or_str(sol) if sol is not None else ''
                    level = entry.get('vul_level') or vm.get('vul_level') or entry.get('threat_level') or vm.get('threat_level','')
                    rec = {
                        'IP': _clean_text_list_or_str(ip) if ip else '',
                        '端口': str(port) if port else '',
                        '漏洞名称': _clean_text_list_or_str(name) if name else '',
                        '风险等级': normalize_risk(level),
                        '漏洞说明': _clean_text_list_or_str(desc) if desc else '',
                        '加固建议': fix_txt,
                        'CVE': _clean_text_list_or_str(cve) if cve else ''
                    }
                    vulns_all.append(rec)
                    if ip and port:
                        ports_all.append((ip, str(port)))
        # 统一归一化
        return [normalize_record(r) for r in vulns_all], ports_all

    # 2) 回退到表格/文本解析
    if soup is None:
        return [], []

    for table in soup.find_all('table'):
        try:
            vrows, ppairs = extract_from_table(table)
            vulns_all.extend(vrows)
            ports_all.extend(ppairs)
        except Exception:
            log(f"表格解析出错: {traceback.format_exc()}")

    try:
        v2, p2 = extract_from_blocks(soup)
        vulns_all.extend(v2)
        ports_all.extend(p2)
    except Exception:
        log(f"文本块解析出错: {traceback.format_exc()}")

    return [normalize_record(r) for r in vulns_all], ports_all

# ----------------------------- 合并 & 输出 -----------------------------

def merge_vulns(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen = set(); out = []
    for r in records:
        key = (r.get('IP','').strip(), r.get('端口','').strip(), r.get('漏洞名称','').strip(), r.get('CVE','').strip())
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    for i, r in enumerate(out, start=1):
        r['序号'] = i
    for r in out:
        for k in ['漏洞名称','风险等级','漏洞说明','加固建议','CVE','IP','端口']:
            if not r.get(k):
                r[k] = ''
    return out

def merge_ports(pairs: List[Tuple[str,str]]) -> List[Dict[str, Any]]:
    seen = set(); out = []
    for ip, port in pairs:
        if not ip or not port:
            continue
        key = (ip.strip(), port.strip())
        if key in seen:
            continue
        seen.add(key)
        out.append({'序号': 0, 'IP': key[0], '端口': key[1]})
    for i, r in enumerate(out, start=1):
        r['序号'] = i
    return out

def save_excels(vulns: List[Dict[str, Any]], ports: List[Dict[str, Any]], outdir: Path) -> None:
    ensure_env()
    outdir.mkdir(parents=True, exist_ok=True)
    # 漏洞报告
    df_v = pd.DataFrame(vulns)
    cols_v = ['序号','IP','端口','漏洞名称','风险等级','漏洞说明','加固建议','CVE']
    for c in cols_v:
        if c not in df_v.columns:
            df_v[c] = ''
    df_v = df_v[cols_v]
    (outdir / '漏洞报告.xlsx').unlink(missing_ok=True)
    df_v.to_excel(outdir / '漏洞报告.xlsx', index=False)

    # 开放端口
    df_p = pd.DataFrame(ports)
    cols_p = ['序号','IP','端口']
    for c in cols_p:
        if c not in df_p.columns:
            df_p[c] = ''
    df_p = df_p[cols_p]
    (outdir / '开放端口.xlsx').unlink(missing_ok=True)
    df_p.to_excel(outdir / '开放端口.xlsx', index=False)

    log(f"已输出：{outdir / '漏洞报告.xlsx'}  和  {outdir / '开放端口.xlsx'}")

# ----------------------------- Orchestrator -----------------------------

def use_existing_outputs_if_present(base: Path, outdir: Path) -> bool:
    ok = False
    for fname in ('漏洞报告.xlsx','开放端口.xlsx'):
        src = base / fname
        if src.exists():
            try:
                outdir.mkdir(parents=True, exist_ok=True)
                dst = outdir / src.name
                shutil.copy2(str(src), str(dst))
                log(f"发现现成文件并复制: {src} -> {dst}")
                ok = True
            except Exception as e:
                log(f"复制现成文件失败 {src}: {e}")
    return ok

def process_folder(base: Path, output_folder: Path, force_regenerate: bool=True) -> None:
    ensure_env()

    files = [p for p in base.iterdir() if p.is_file()]
    zips = [p for p in files if p.suffix.lower() == '.zip']
    log(f"发现文件数: {len(files)}，zip: {len(zips)}")

    if not force_regenerate and use_existing_outputs_if_present(base, output_folder):
        log("优先使用现成输出文件，跳过解析（若需重新解析请使用 --force-regenerate）")
        return

    tempdir = Path(tempfile.mkdtemp(prefix='rsas2_html_'))
    try:
        targets = [base]
        if zips:
            targets += unzip_all(zips, tempdir)

        # 收集所有 html
        html_files: List[Path] = []
        for t in targets:
            html_files += find_files(t, ('.html', '.htm'))
        log(f"发现 HTML 报告: {len(html_files)} 个")

        # 实时进度条遍历 HTML 文件
        all_vulns: List[Dict[str, Any]] = []
        all_ports_pairs: List[Tuple[str,str]] = []
        for hp in tqdm(html_files, desc='解析 HTML', unit='file'):
            try:
                v, p = parse_html_file(hp)
                if v:
                    all_vulns.extend(v)
                if p:
                    all_ports_pairs.extend(p)
            except Exception as e:
                log(f"解析 {hp} 出错: {e}\n{traceback.format_exc()}")

        # 合并与导出
        merged_vulns = merge_vulns(all_vulns)
        merged_ports = merge_ports(all_ports_pairs)
        save_excels(merged_vulns, merged_ports, output_folder)

    finally:
        try:
            shutil.rmtree(tempdir, ignore_errors=True)
        except Exception:
            pass

def main():
    parser = argparse.ArgumentParser(description='解析HTML报告并生成 漏洞报告.xlsx 和 开放端口.xlsx，并打印作者横幅')
    parser.add_argument('--input', '-i', default='.', help='输入目录（默认当前）')
    parser.add_argument('--output', '-o', default='./整理结果', help='输出目录（默认 ./整理结果）')
    parser.add_argument('--force-regenerate', dest='force', action='store_true', default=True,
                        help='强制重新解析所有文件（默认开启）')
    parser.add_argument('--no-force-regenerate', dest='force', action='store_false',
                        help='关闭强制解析，若目录下已有 漏洞报告.xlsx/开放端口.xlsx 则直接复用')
    parser.add_argument('--no-unicode', dest='no_unicode', action='store_true',
                        help='强制使用 ASCII 框（不使用 Unicode 盒绘字符）')
    parser.add_argument('--margin', type=int, default=0, help='横幅左侧外边距空格数（默认 0）')
    parser.add_argument('--pad', type=int, default=1, help='横幅内部左右边距（默认 1）')
    args = parser.parse_args()

    # 打印横幅
    print_banner(use_unicode=not args.no_unicode, outer_margin=args.margin, inner_pad=max(0, args.pad))

    # 处理报告
    base = Path(args.input).resolve()
    outdir = Path(args.output).resolve()
    log(f"输入目录: {base}，输出目录: {outdir}，force_regenerate={args.force}")
    try:
        process_folder(base, outdir, force_regenerate=args.force)
    except Exception as e:
        log(f"主流程异常: {e}\n{traceback.format_exc()}")
        sys.exit(2)

if __name__ == '__main__':
    main()