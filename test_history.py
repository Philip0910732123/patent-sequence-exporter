#!/usr/bin/env python3
"""patent-sequence-exporter v6 历史记录功能冒烟测试"""
import json
import sys
import os
import tempfile
from pathlib import Path

# 导入 app 模块
sys.path.insert(0, str(Path(__file__).parent))
os.chdir(Path(__file__).parent)

# 模拟 frozen 环境
import app as appmod

# 使用临时目录
tmpdir = Path(tempfile.mkdtemp())
appmod.HISTORY_DIR = tmpdir / "history"
appmod.HISTORY_DIR.mkdir(exist_ok=True)
appmod.OUTPUT_DIR = tmpdir / "outputs"
appmod.OUTPUT_DIR.mkdir(exist_ok=True)

from flask import Flask
test_app = appmod.app
test_app.config['TESTING'] = True
client = test_app.test_client()

passed = 0
failed = 0

def check(name, condition, detail=""):
    global passed, failed
    if condition:
        passed += 1
        print(f"[PASS] {name}")
    else:
        failed += 1
        print(f"[FAIL] {name} - {detail}")

# Test 1: history 目录自动创建
check("history 目录自动创建", appmod.HISTORY_DIR.exists())

# Test 2: 模拟保存历史记录
fake_sequences = [
    {"seq_id": "seq1", "seq_id_no": "1", "sequence_type": "PROTEIN", "s_len": 100, "s_sequence": "MKTAYIAKQRQ", "location": {"loc": ["seq1"]}},
    {"seq_id": "seq2", "seq_id_no": "2", "sequence_type": "NUCLEOTIDE", "s_len": 200, "s_sequence": "ATGCATGCATGC", "location": {"loc": ["seq2"]}},
]
fake_data = {
    "patent_number": "WO2024000001A1",
    "seq_id_range": "1-2",
    "output_format": "excel",
    "sequence_type": "",
    "seq_length_start": None,
    "seq_length_end": None,
}
fake_files = [{"name": "test_sequences.xlsx", "url": "/download/test_sequences.xlsx"}]
hist_file = appmod._save_history(fake_data, fake_sequences, "WO2024000001A1", fake_files)
check("保存历史记录", hist_file.startswith("history_") and hist_file.endswith(".json"))

# Test 3: 验证历史记录内容
hist_path = appmod.HISTORY_DIR / hist_file
with open(hist_path, encoding="utf-8") as f:
    hist_content = json.load(f)
check("历史记录字段完整", all(k in hist_content for k in ["timestamp", "patent", "sequences", "type_distribution", "files"]), f"缺失字段: {set(['timestamp','patent','sequences','type_distribution','files']) - set(hist_content.keys())}")
check("历史记录序列数正确", hist_content["total"] == 2)

# Test 4: GET /api/history 列表
resp = client.get("/api/history")
data = resp.get_json()
check("GET /api/history 返回成功", data.get("success") is True)
check("GET /api/history 返回记录", data.get("count") == 1 and len(data.get("history", [])) == 1)
check("GET /api/history 摘要不含 sequences", "sequences" not in data["history"][0])

# Test 5: GET /api/history/<filename> 详情
resp = client.get(f"/api/history/{hist_file}")
data = resp.get_json()
check("GET /api/history/<filename> 返回成功", data.get("success") is True)
check("GET /api/history/<filename> 包含 sequences", len(data.get("data", {}).get("sequences", [])) == 2)

# Test 6: 非法文件名拦截
resp = client.get("/api/history/../../etc/passwd")
check("非法文件名拦截", resp.status_code == 400)

# Test 7: DELETE /api/history/<filename>
resp = client.delete(f"/api/history/{hist_file}")
data = resp.get_json()
check("DELETE 历史记录", data.get("success") is True)

# Test 8: 删除后列表为空
resp = client.get("/api/history")
data = resp.get_json()
check("删除后列表为空", data.get("count") == 0)

# Test 9: 下载历史记录 JSON
# 先重新保存一条
hist_file2 = appmod._save_history(fake_data, fake_sequences, "WO2024000001A1", fake_files)
resp = client.get(f"/api/history/{hist_file2}/download/json")
check("下载历史记录 JSON", resp.status_code == 200)

# Test 10: 下载历史记录 FASTA
resp = client.get(f"/api/history/{hist_file2}/download/fasta")
check("下载历史记录 FASTA", resp.status_code == 200)

# Test 11: 下载历史记录 Excel
resp = client.get(f"/api/history/{hist_file2}/download/excel")
check("下载历史记录 Excel", resp.status_code == 200)

print(f"\n=== Results: {passed} passed, {failed} failed ===")
sys.exit(0 if failed == 0 else 1)
