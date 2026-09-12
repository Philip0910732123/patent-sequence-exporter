#!/usr/bin/env python3
"""
专利生物序列导出器 - Flask Web 版
本地运行，通过浏览器界面调用智慧芽 Bio OpenAPI 提取专利序列并下载。
支持 PyInstaller 打包：自动检测 frozen 模式，使用 sys._MEIPASS 定位模板。
集成专利号标准化 API（patent-search-pn），支持非标准专利号智能查找。
"""

import json
import time
import sys
import os
import re
import threading
import webbrowser
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, send_file, send_from_directory
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
import requests

app = Flask(__name__)

# ── 配置 ──────────────────────────────────────────────────
API_URL = "https://connect.zhihuiya.com/bio-openapi/patent/extract/single"
PATENT_SEARCH_URL = "https://connect.zhihuiya.com/search/patent/patent-search-pn"
PAGE_SIZE = 100
MAX_RETRIES = 3
RETRY_INTERVAL = 3

# PyInstaller 兼容：frozen 时用临时解包目录找模板，否则用脚本目录
if getattr(sys, 'frozen', False):
    _BASE_DIR = Path(sys._MEIPASS)
    _EXE_DIR = Path(sys.executable).parent
else:
    _BASE_DIR = Path(__file__).parent
    _EXE_DIR = Path(__file__).parent

TEMPLATE_DIR = _BASE_DIR / "templates"
OUTPUT_DIR = _EXE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

# ── 历史记录 ──────────────────────────────────────────────
HISTORY_DIR = _EXE_DIR / "history"
HISTORY_DIR.mkdir(exist_ok=True)

import re as _re
_HISTORY_NAME_RE = _re.compile(r"^history_\d{8}_\d{6}\.json$")


def _save_history(extract_data, sequences, patent_label, files_info):
    """提取成功后自动保存历史记录到 exe 同目录 history/ 下。"""
    ts = datetime.now()
    ts_str = ts.strftime("%Y%m%d_%H%M%S")
    ts_display = ts.strftime("%Y-%m-%d %H:%M:%S")

    type_dist = {}
    for s in sequences:
        st = s.get("sequence_type", "UNKNOWN")
        type_dist[st] = type_dist.get(st, 0) + 1

    history_entry = {
        "timestamp": ts_str,
        "timestamp_display": ts_display,
        "patent": patent_label,
        "patent_number": extract_data.get("patent_number", ""),
        "seq_id_range": extract_data.get("seq_id_range", ""),
        "output_format": extract_data.get("output_format", "excel"),
        "sequence_type": extract_data.get("sequence_type") or "",
        "seq_length_start": extract_data.get("seq_length_start"),
        "seq_length_end": extract_data.get("seq_length_end"),
        "total": len(sequences),
        "type_distribution": type_dist,
        "files": files_info,
        "sequences": sequences,
    }

    filename = f"history_{ts_str}.json"
    filepath = HISTORY_DIR / filename
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(history_entry, f, ensure_ascii=False, indent=2)

    return filename


def _timestamped_path(path: str) -> str:
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    p = Path(path)
    return str(p.parent / f"{p.stem}_{ts}{p.suffix}")


def parse_seq_id_range(seq_input: str) -> tuple:
    raw = seq_input.strip()
    # 支持空格分隔，自动转换为逗号
    if " " in raw and "," not in raw and "-" not in raw:
        parts = raw.split()
        raw = ",".join(parts)
    lower = raw.lower()
    if lower in ("all", "全部", "all sequences", "所有", "全部序列"):
        return "EXTRACT_ALL", "all"
    if "-" in raw:
        return "EXTRACT_RANGE", raw
    return "EXTRACT_NUM", raw


def read_api_key(api_key, api_key_file):
    """从直接输入或文件路径获取 API Key。"""
    if api_key:
        return api_key.strip()
    if api_key_file:
        p = Path(api_key_file)
        if p.exists():
            content = p.read_text(encoding="utf-8-sig").strip()
            m = re.search(r'sk-[A-Za-z0-9]+', content)
            if m:
                return m.group(0)
            for line in content.splitlines():
                line = line.strip()
                if line and not line.startswith("#") and not line.startswith("//"):
                    if ":" in line:
                        val = line.split(":", 1)[1].strip()
                        if val and val.isascii():
                            return val
                    if "=" in line:
                        val = line.split("=", 1)[1].strip()
                        if val and val.isascii():
                            return val
                    if line.isascii() and len(line) > 10:
                        return line
            raise ValueError("无法从文件中提取 API Key")
        raise ValueError(f"文件不存在: {api_key_file}")
    raise ValueError("未提供 API Key")


# ── 专利号标准化 API ─────────────────────────────────────

def search_patent_pn(api_key, query, authority=None):
    """
    调用专利号标准化 API（patent-search-pn）。
    先用 pn（公开号）查询，无结果再用 apno（申请号）查询。
    返回标准化专利号列表。
    """
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json;charset=UTF-8",
    }

    results = []

    # 尝试1: 作为公开号 (pn) 查询
    payload_pn = {
        "pn": query,
        "limit": 20,
        "offset": 0,
    }
    if authority:
        payload_pn["authority"] = authority

    body_pn = _do_patent_search(headers, payload_pn)
    results.extend(_extract_patent_list(body_pn, query))

    # 尝试2: 作为申请号 (apno) 查询
    payload_apno = {
        "apno": query,
        "limit": 20,
        "offset": 0,
    }
    if authority:
        payload_apno["authority"] = authority

    body_apno = _do_patent_search(headers, payload_apno)
    results.extend(_extract_patent_list(body_apno, query))

    # 去重
    seen = set()
    unique = []
    for r in results:
        key = r.get("pn", "")
        if key and key not in seen:
            seen.add(key)
            unique.append(r)

    return unique


def _do_patent_search(headers, payload):
    """执行一次专利号搜索请求，带重试。"""
    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.post(PATENT_SEARCH_URL, headers=headers, json=payload, timeout=60)
            if resp.status_code != 200:
                raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:500]}")
            body = resp.json()
            # 某些错误返回 status=false 但 HTTP 200
            if isinstance(body, dict) and body.get("error_code") and body.get("error_code") != 0:
                # 不是真正的错误，可能是无结果的提示，返回空
                return {"data": {"results": []}}
            return body
        except Exception as e:
            last_err = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_INTERVAL)
            else:
                raise RuntimeError(f"专利号查找请求{MAX_RETRIES}次重试后仍失败: {last_err}")


def _extract_patent_list(body, query):
    """从 API 响应中提取专利列表，统一格式。"""
    if not body or not isinstance(body, dict):
        return []

    # 尝试多种可能的响应结构
    data = body.get("data") or body
    results = None

    # 常见结构1: data.results
    if isinstance(data, dict):
        results = data.get("results") or data.get("list") or data.get("patents")
    # 常见结构2: body.results
    if not results:
        results = body.get("results") or body.get("list") or body.get("patents")

    if not results or not isinstance(results, list):
        return []

    extracted = []
    for item in results:
        if not isinstance(item, dict):
            continue
        pn = item.get("pn") or item.get("patent_number") or item.get("publication_number") or ""
        patent_id = item.get("patent_id") or item.get("id") or ""
        title = item.get("title") or item.get("patent_title") or ""
        applicant = item.get("applicant") or item.get("assignee") or ""
        apno = item.get("apno") or item.get("application_number") or ""
        authority_code = item.get("authority") or item.get("country") or item.get("office") or ""

        if pn:
            extracted.append({
                "pn": pn,
                "patent_id": str(patent_id) if patent_id else "",
                "title": title,
                "applicant": applicant,
                "apno": apno,
                "authority": authority_code,
            })

    return extracted


# ── 序列提取 API ──────────────────────────────────────────

def fetch_page(api_key, patent_number, patent_id, extract_type,
               sequence_no_range, sequence_type, sequence_length, page):
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json;charset=UTF-8",
    }
    payload = {
        "extract_type": extract_type,
        "page": page,
        "size": PAGE_SIZE,
        "sequence_no_range": sequence_no_range,
    }
    if patent_id:
        payload["patent_id"] = patent_id
    elif patent_number:
        payload["patent_number"] = patent_number
    else:
        raise ValueError("patent_number 和 patent_id 至少提供一个")
    if sequence_type:
        payload["sequence_type"] = sequence_type
    if sequence_length:
        payload["sequence_length"] = sequence_length

    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.post(API_URL, headers=headers, json=payload, timeout=120)
            if resp.status_code != 200:
                raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:500]}")
            body = resp.json()
            if not body.get("status", False):
                err_code = body.get("error_code", "unknown")
                err_msg = body.get("error_msg") or body.get("error_message") or ""
                raise RuntimeError(f"API错误 error_code={err_code}, error_msg={err_msg}")
            return body
        except Exception as e:
            last_err = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_INTERVAL)
            else:
                raise RuntimeError(f"API请求{MAX_RETRIES}次重试后仍失败: {last_err}")


def fetch_all(api_key, patent_number, patent_id, extract_type,
              sequence_no_range, sequence_type, sequence_length):
    all_results = []
    total = None
    page = 1
    while True:
        body = fetch_page(api_key, patent_number, patent_id, extract_type,
                         sequence_no_range, sequence_type, sequence_length, page=page)
        data = body.get("data") or {}
        if total is None:
            total = data.get("total", 0)
        results = data.get("results") or []
        all_results.extend(results)
        if not results or len(all_results) >= total:
            break
        page += 1
    return all_results, total


# ── 导出函数 ──────────────────────────────────────────────

def export_excel(sequences, output_path, patent_label):
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Summary"
    type_dist = {}
    for s in sequences:
        st = s.get("sequence_type", "UNKNOWN")
        type_dist[st] = type_dist.get(st, 0) + 1
    summary_rows = [
        ("专利号", patent_label),
        ("序列总数", len(sequences)),
        ("导出时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("", ""),
        ("序列类型分布", ""),
    ]
    for st, cnt in sorted(type_dist.items()):
        summary_rows.append((st, cnt))
    s_font = Font(bold=True, color="FFFFFF")
    s_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    for r, (k, v) in enumerate(summary_rows, 1):
        c1 = ws_summary.cell(row=r, column=1, value=k)
        c2 = ws_summary.cell(row=r, column=2, value=v)
        if r == 1 or (k and not v and r == 5):
            c1.font = s_font
            c1.fill = s_fill
        elif k and not isinstance(v, (int, float)) and not v:
            c1.font = Font(bold=True)
    ws_summary.column_dimensions["A"].width = 20
    ws_summary.column_dimensions["B"].width = 35

    ws = wb.create_sheet("Sequences")
    loc_keys = set()
    for s in sequences:
        loc = s.get("location") or {}
        loc_keys.update(loc.keys())
    loc_keys = sorted(loc_keys)
    headers = (
        ["Seq ID", "Sequence ID Number", "Sequence Type", "Length"]
        + [f"Location_{k}" for k in loc_keys]
        + ["Sequence"]
    )
    h_font = Font(bold=True, color="FFFFFF")
    h_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    h_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = h_font
        c.fill = h_fill
        c.alignment = h_align
    for r, s in enumerate(sequences, 2):
        sid = s.get("seq_id", "")
        sid_num = s.get("sequence_id_number", [])
        sid_num_str = ", ".join(str(x) for x in sid_num) if isinstance(sid_num, list) else str(sid_num)
        stype = s.get("sequence_type", "")
        slen = s.get("s_len", "")
        seq = s.get("s_sequence", "")
        loc = s.get("location") or {}
        col = 1
        ws.cell(row=r, column=col, value=sid); col += 1
        ws.cell(row=r, column=col, value=sid_num_str); col += 1
        ws.cell(row=r, column=col, value=stype); col += 1
        ws.cell(row=r, column=col, value=slen); col += 1
        for lk in loc_keys:
            lv = loc.get(lk, [])
            lv_str = ", ".join(str(x) for x in lv) if isinstance(lv, list) else str(lv)
            ws.cell(row=r, column=col, value=lv_str); col += 1
        cell = ws.cell(row=r, column=col, value=seq)
        cell.alignment = Alignment(wrap_text=False)
    for ci in range(1, len(headers) + 1):
        max_len = len(str(headers[ci - 1]))
        for ri in range(2, min(len(sequences) + 2, 50)):
            v = ws.cell(row=ri, column=ci).value
            if v is not None:
                max_len = max(max_len, min(len(str(v)), 40))
        ws.column_dimensions[ws.cell(row=1, column=ci).column_letter].width = max_len + 2
    ws.freeze_panes = "A2"
    try:
        wb.save(output_path)
    except PermissionError:
        output_path = _timestamped_path(output_path)
        wb.save(output_path)
    return output_path


def export_fasta(sequences, output_path, patent_label):
    def write_inner(f):
        for s in sequences:
            sid = s.get("seq_id", "")
            sid_num = s.get("sequence_id_number", [])
            sid_num_str = ",".join(str(x) for x in sid_num) if isinstance(sid_num, list) else str(sid_num)
            stype = s.get("sequence_type", "UNKNOWN")
            slen = s.get("s_len", 0)
            seq = s.get("s_sequence", "")
            header = f">{patent_label}|seq_id_no:{sid_num_str}|type:{stype}|len:{slen}|seq_id:{sid}"
            f.write(header + "\n")
            for i in range(0, len(seq), 60):
                f.write(seq[i:i + 60] + "\n")
    try:
        with open(output_path, "w") as f:
            write_inner(f)
    except PermissionError:
        output_path = _timestamped_path(output_path)
        with open(output_path, "w") as f:
            write_inner(f)
    return output_path


# ── Flask 路由 ────────────────────────────────────────────

@app.route("/")
def index():
    html_path = TEMPLATE_DIR / "index.html"
    if html_path.exists():
        return send_file(str(html_path))
    return "<h1>templates/index.html not found</h1>", 404


@app.route("/api/search-pn", methods=["POST"])
def search_pn():
    """专利号智能查找：接收任意格式的专利号，返回标准公开号列表。"""
    try:
        data = request.get_json()
        api_key_input = data.get("api_key", "").strip()
        api_key_file = data.get("api_key_file", "").strip()
        query = data.get("query", "").strip()
        authority = data.get("authority")

        if not query:
            return jsonify({"success": False, "error": "请输入专利号"})

        try:
            api_key = read_api_key(api_key_input, api_key_file)
        except Exception as e:
            return jsonify({"success": False, "error": f"API Key 获取失败: {e}"})

        results = search_patent_pn(api_key, query, authority)

        if not results:
            return jsonify({
                "success": False,
                "error": f"未找到与 \"{query}\" 匹配的专利，请检查输入或尝试其他格式"
            })

        return jsonify({
            "success": True,
            "count": len(results),
            "results": results,
            "query": query,
        })
    except Exception as e:
        err_msg = str(e)
        if "not found" in err_msg.lower() or "error" in err_msg.lower():
            err_msg += "，请点击「核对专利号」按钮确认专利号是否正确"
        return jsonify({"success": False, "error": err_msg})


@app.route("/api/extract", methods=["POST"])
def extract():
    try:
        data = request.get_json()
        api_key_input = data.get("api_key", "").strip()
        api_key_file = data.get("api_key_file", "").strip()
        patent_number = data.get("patent_number", "").strip()
        patent_id = data.get("patent_id", "").strip()
        seq_id_range = data.get("seq_id_range", "").strip()
        output_format = data.get("output_format", "excel")
        sequence_type = data.get("sequence_type", "").strip() or None
        seq_length_start = data.get("seq_length_start")
        seq_length_end = data.get("seq_length_end")

        if not patent_number and not patent_id:
            return jsonify({"success": False, "error": "必须提供专利公开号或专利 UUID"})
        if not seq_id_range:
            return jsonify({"success": False, "error": "必须提供 Seq ID 范围"})

        try:
            api_key = read_api_key(api_key_input, api_key_file)
        except Exception as e:
            return jsonify({"success": False, "error": f"API Key 获取失败: {e}"})

        # 空格分隔自动转逗号
        seq_id_range_normalized = seq_id_range
        if " " in seq_id_range_normalized and "," not in seq_id_range_normalized and "-" not in seq_id_range_normalized:
            parts = seq_id_range_normalized.split()
            seq_id_range_normalized = ",".join(parts)

        # 自动将专利公开号转换为 UUID
        resolved_patent_id = patent_id
        if not resolved_patent_id and patent_number:
            try:
                pn_results = search_patent_pn(api_key, patent_number)
                if not pn_results:
                    return jsonify({
                        "success": False,
                        "error": f"未找到专利号 \"{patent_number}\" 对应的记录，请点击「核对专利号」按钮确认正确的专利公开号"
                    })
                if len(pn_results) == 1:
                    resolved_patent_id = pn_results[0].get("patent_id", "")
                    if not resolved_patent_id:
                        return jsonify({
                            "success": False,
                            "error": f"专利号 \"{patent_number}\" 查找到了但未返回 UUID，请点击「核对专利号」按钮手动选择"
                        })
                else:
                    # 多个结果，提示用户手动选择
                    return jsonify({
                        "success": False,
                        "error": f"专利号 \"{patent_number}\" 匹配到 {len(pn_results)} 条结果，请点击「核对专利号」按钮选择正确的专利",
                        "multiple_matches": True,
                        "matches": pn_results
                    })
            except Exception as e:
                return jsonify({
                    "success": False,
                    "error": f"专利号转换 UUID 失败: {e}，请点击「核对专利号」按钮确认专利号"
                })

        patent_label = resolved_patent_id or patent_number
        extract_type, seq_no_range = parse_seq_id_range(seq_id_range_normalized)

        seq_len = None
        if seq_length_start is not None or seq_length_end is not None:
            seq_len = {
                "start": int(seq_length_start) if seq_length_start is not None else 0,
                "end": int(seq_length_end) if seq_length_end is not None else 999999,
            }

        sequences, total = fetch_all(
            api_key=api_key,
            patent_number=patent_number if not resolved_patent_id else "",
            patent_id=resolved_patent_id,
            extract_type=extract_type,
            sequence_no_range=seq_no_range,
            sequence_type=sequence_type,
            sequence_length=seq_len,
        )

        if not sequences:
            return jsonify({"success": False, "error": "未找到符合条件的序列，请点击「核对专利号」按钮确认专利号是否正确"})

        files = []
        if output_format in ("excel", "both"):
            p = str(OUTPUT_DIR / f"{patent_label}_sequences.xlsx")
            actual_path = export_excel(sequences, p, patent_label)
            fname = Path(actual_path).name
            files.append({"name": fname, "url": f"/download/{fname}"})

        if output_format in ("fasta", "both"):
            p = str(OUTPUT_DIR / f"{patent_label}_sequences.fasta")
            actual_path = export_fasta(sequences, p, patent_label)
            fname = Path(actual_path).name
            files.append({"name": fname, "url": f"/download/{fname}"})

        type_dist = {}
        for s in sequences:
            st = s.get("sequence_type", "UNKNOWN")
            type_dist[st] = type_dist.get(st, 0) + 1

        table_sequences = []
        for s in sequences:
            sid_num = s.get("sequence_id_number", [])
            sid_num_str = ", ".join(str(x) for x in sid_num) if isinstance(sid_num, list) and sid_num else ""
            loc = s.get("location") or {}
            loc_flat = {}
            for k, v in loc.items():
                loc_flat[k] = ", ".join(str(x) for x in v) if isinstance(v, list) else str(v)
            table_sequences.append({
                "seq_id": s.get("seq_id", ""),
                "seq_id_no": sid_num_str,
                "sequence_type": s.get("sequence_type", ""),
                "s_len": s.get("s_len", 0),
                "s_sequence": s.get("s_sequence", ""),
                "location": loc_flat,
            })

        # 保存历史记录
        history_filename = _save_history(
            extract_data=data,
            sequences=table_sequences,
            patent_label=patent_label,
            files_info=files,
        )

        return jsonify({
            "success": True,
            "total": total,
            "fetched": len(sequences),
            "files": files,
            "type_distribution": type_dist,
            "patent": patent_label,
            "sequences": table_sequences,
            "history_saved": history_filename,
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(str(OUTPUT_DIR), filename, as_attachment=True)


# ── 历史记录 API ──────────────────────────────────────────

@app.route("/api/history")
def list_history():
    """列出所有历史记录（按时间倒序，仅摘要）。"""
    files = sorted(HISTORY_DIR.glob("history_*.json"), reverse=True)
    records = []
    for fp in files:
        try:
            with open(fp, encoding="utf-8") as f:
                data = json.load(f)
            records.append({
                "filename": fp.name,
                "timestamp": data.get("timestamp", ""),
                "timestamp_display": data.get("timestamp_display", ""),
                "patent": data.get("patent", ""),
                "patent_number": data.get("patent_number", ""),
                "seq_id_range": data.get("seq_id_range", ""),
                "output_format": data.get("output_format", ""),
                "total": data.get("total", 0),
                "type_distribution": data.get("type_distribution", {}),
                "files": data.get("files", []),
            })
        except Exception:
            continue
    return jsonify({"success": True, "count": len(records), "history": records})


@app.route("/api/history/<filename>")
def get_history(filename):
    """获取某条历史记录完整结果。"""
    if not _HISTORY_NAME_RE.match(filename):
        return jsonify({"success": False, "error": "非法文件名"}), 400
    filepath = HISTORY_DIR / filename
    if not filepath.exists():
        return jsonify({"success": False, "error": "记录不存在"}), 404
    with open(filepath, encoding="utf-8") as f:
        data = json.load(f)
    return jsonify({"success": True, "data": data})


@app.route("/api/history/<filename>/download/<fmt>")
def download_history(filename, fmt):
    """下载历史记录的 Excel/FASTA/JSON。"""
    if not _HISTORY_NAME_RE.match(filename):
        return jsonify({"success": False, "error": "非法文件名"}), 400
    filepath = HISTORY_DIR / filename
    if not filepath.exists():
        return jsonify({"success": False, "error": "记录不存在"}), 404
    with open(filepath, encoding="utf-8") as f:
        data = json.load(f)
    sequences = data.get("sequences", [])
    patent_label = data.get("patent", "patent")

    if fmt == "json":
        download_path = str(OUTPUT_DIR / f"{patent_label}_history.json")
        with open(download_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return send_file(download_path, as_attachment=True)

    if fmt == "fasta":
        download_path = str(OUTPUT_DIR / f"{patent_label}_history.fasta")
        with open(download_path, "w") as f:
            for s in sequences:
                sid_num = s.get("seq_id_no", "")
                stype = s.get("sequence_type", "UNKNOWN")
                slen = s.get("s_len", 0)
                seq = s.get("s_sequence", "")
                f.write(f">{patent_label}|seq_id_no:{sid_num}|type:{stype}|len:{slen}\n")
                for i in range(0, len(seq), 60):
                    f.write(seq[i:i + 60] + "\n")
        return send_file(download_path, as_attachment=True)

    if fmt == "excel":
        download_path = str(OUTPUT_DIR / f"{patent_label}_history.xlsx")
        export_excel(sequences, download_path, patent_label)
        return send_file(download_path, as_attachment=True)

    return jsonify({"success": False, "error": f"不支持的格式: {fmt}"}), 400


@app.route("/api/history/<filename>", methods=["DELETE"])
def delete_history(filename):
    """删除某条历史记录。"""
    if not _HISTORY_NAME_RE.match(filename):
        return jsonify({"success": False, "error": "非法文件名"}), 400
    filepath = HISTORY_DIR / filename
    if not filepath.exists():
        return jsonify({"success": False, "error": "记录不存在"}), 404
    filepath.unlink()
    return jsonify({"success": True})


def open_browser():
    """延迟 1.5 秒打开浏览器，等待 Flask 启动"""
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == "__main__":
    print("=" * 50)
    print("  专利生物序列导出器 - Web 版 (v6)")
    print(f"  输出目录: {OUTPUT_DIR}")
    print("  访问 http://127.0.0.1:5000 开始使用")
    print("  按 Ctrl+C 退出")
    print("=" * 50)

    # 自动打开浏览器
    threading.Thread(target=open_browser, daemon=True).start()

    app.run(host="127.0.0.1", port=5000, debug=False)
