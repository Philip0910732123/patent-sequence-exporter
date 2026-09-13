#!/usr/bin/env python3
"""
patent-sequence-exporter selftest - 自检脚本
验证核心函数的正确性，不依赖网络和 API Key。
"""
import sys
import os
import json
import tempfile
from pathlib import Path

# 确保能导入 batch_extract
sys.path.insert(0, str(Path(__file__).parent))
from batch_extract import (
    parse_patent_ids,
    parse_seq_id_range,
    read_api_key,
    _timestamped_path,
    append_ndjson,
    load_batch_state,
    save_batch_state,
)


def test_parse_patent_ids():
    """测试专利号列表解析：逗号/分号/换行分隔 + 去空格/斜杠 + 去重"""
    ids = parse_patent_ids('WO2020123456, WO2021123456; US2022001234A1\nWO2023/456789')
    assert len(ids) == 4, f'Expected 4, got {len(ids)}: {ids}'
    assert ids[0] == 'WO2020123456', f'Expected WO2020123456, got {ids[0]}'
    assert ids[1] == 'WO2021123456', f'Expected WO2021123456, got {ids[1]}'
    assert ids[2] == 'US2022001234A1', f'Expected US2022001234A1, got {ids[2]}'
    assert ids[3] == 'WO2023456789', f'Expected WO2023456789, got {ids[3]}'
    print('[PASS] test_parse_patent_ids')

    # 测试去重
    ids2 = parse_patent_ids('WO2020123456, WO2020123456, WO2020123456')
    assert len(ids2) == 1, f'Expected 1 (deduped), got {len(ids2)}'
    print('[PASS] test_parse_patent_ids_dedup')

    # 测试空输入
    ids3 = parse_patent_ids(',,,;;;')
    assert len(ids3) == 0, f'Expected 0, got {len(ids3)}'
    print('[PASS] test_parse_patent_ids_empty')


def test_parse_seq_id_range():
    """测试 Seq ID 范围解析"""
    et, sr = parse_seq_id_range('all')
    assert et == 'EXTRACT_ALL' and sr == 'all', f'Expected EXTRACT_ALL/all, got {et}/{sr}'
    print('[PASS] test_parse_seq_id_range_all')

    et, sr = parse_seq_id_range('1-50')
    assert et == 'EXTRACT_RANGE' and sr == '1-50', f'Expected EXTRACT_RANGE/1-50, got {et}/{sr}'
    print('[PASS] test_parse_seq_id_range_range')

    et, sr = parse_seq_id_range('1,2,3')
    assert et == 'EXTRACT_NUM' and sr == '1,2,3', f'Expected EXTRACT_NUM/1,2,3, got {et}/{sr}'
    print('[PASS] test_parse_seq_id_range_num')

    # 空格自动转逗号
    et, sr = parse_seq_id_range('1 2 3')
    assert et == 'EXTRACT_NUM' and sr == '1,2,3', f'Expected 1,2,3, got {sr}'
    print('[PASS] test_parse_seq_id_range_space_to_comma')


def test_read_api_key():
    """测试 API Key 读取"""
    # 直接输入
    key = read_api_key('sk-test123')
    assert key == 'sk-test123', f'Expected sk-test123, got {key}'
    print('[PASS] test_read_api_key_direct')

    # 从文件读取
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as f:
        f.write('sk-fromfile456')
        tmpfile = f.name
    try:
        key = read_api_key('', tmpfile)
        assert key == 'sk-fromfile456', f'Expected sk-fromfile456, got {key}'
        print('[PASS] test_read_api_key_file')
    finally:
        os.unlink(tmpfile)

    # 空输入
    try:
        read_api_key('')
        assert False, 'Should have raised ValueError'
    except ValueError:
        print('[PASS] test_read_api_key_empty_raises')


def test_ndjson_append():
    """测试 NDJSON 即时落盘"""
    with tempfile.TemporaryDirectory() as tmpdir:
        ndjson_path = Path(tmpdir) / 'test.ndjson'
        append_ndjson(ndjson_path, {'seq_id': 'SEQ001', 'type': 'DNA'})
        append_ndjson(ndjson_path, {'seq_id': 'SEQ002', 'type': 'Protein'})

        lines = ndjson_path.read_text(encoding='utf-8').strip().split('\n')
        assert len(lines) == 2, f'Expected 2 lines, got {len(lines)}'

        r1 = json.loads(lines[0])
        assert r1['seq_id'] == 'SEQ001', f'Expected SEQ001, got {r1["seq_id"]}'
        assert r1['type'] == 'DNA', f'Expected DNA, got {r1["type"]}'

        r2 = json.loads(lines[1])
        assert r2['seq_id'] == 'SEQ002', f'Expected SEQ002, got {r2["seq_id"]}'
        print('[PASS] test_ndjson_append')


def test_batch_state():
    """测试断点恢复状态文件"""
    with tempfile.TemporaryDirectory() as tmpdir:
        state_path = Path(tmpdir) / 'batch_state.json'

        # 初始状态
        state = load_batch_state(state_path)
        assert state['completed'] == [], f'Expected empty, got {state["completed"]}'
        print('[PASS] test_batch_state_init')

        # 保存状态
        state['completed'].append('WO2020123456')
        state['failed'].append('BAD_PATENT')
        save_batch_state(state_path, state)

        # 重新加载
        state2 = load_batch_state(state_path)
        assert 'WO2020123456' in state2['completed'], 'WO2020123456 should be in completed'
        assert 'BAD_PATENT' in state2['failed'], 'BAD_PATENT should be in failed'
        print('[PASS] test_batch_state_save_load')

        # 损坏文件恢复
        state_path.write_text('corrupted json {{{', encoding='utf-8')
        state3 = load_batch_state(state_path)
        assert state3['completed'] == [], 'Corrupted file should return empty state'
        print('[PASS] test_batch_state_corrupt_recovery')


def test_timestamped_path():
    """测试时间戳路径生成"""
    ts_path = _timestamped_path('/tmp/test.xlsx')
    assert '_20' in ts_path, f'Expected timestamp in path, got {ts_path}'
    assert ts_path.endswith('.xlsx'), f'Expected .xlsx extension, got {ts_path}'
    print('[PASS] test_timestamped_path')


def test_export_fasta():
    """测试 FASTA 导出（离线）"""
    from batch_extract import export_fasta
    sequences = [
        {'seq_id': 'SEQ001', 'sequence_id_number': [1], 'sequence_type': 'DNA', 's_len': 6, 's_sequence': 'ATGCAT'},
        {'seq_id': 'SEQ002', 'sequence_id_number': [2], 'sequence_type': 'Protein', 's_len': 3, 's_sequence': 'MKV'},
    ]
    with tempfile.TemporaryDirectory() as tmpdir:
        fasta_path = str(Path(tmpdir) / 'test.fasta')
        actual = export_fasta(sequences, fasta_path, 'WO_TEST')
        content = Path(actual).read_text()
        assert '>WO_TEST|seq_id_no:1|type:DNA|len:6|seq_id:SEQ001' in content, 'FASTA header 1 missing'
        assert 'ATGCAT' in content, 'Sequence 1 missing'
        assert '>WO_TEST|seq_id_no:2|type:Protein|len:3|seq_id:SEQ002' in content, 'FASTA header 2 missing'
        assert 'MKV' in content, 'Sequence 2 missing'
        print('[PASS] test_export_fasta')


def test_export_json():
    """测试 JSON 导出（离线）"""
    from batch_extract import export_json
    sequences = [
        {'seq_id': 'SEQ001', 'sequence_type': 'DNA', 's_len': 3, 's_sequence': 'ATG'},
    ]
    with tempfile.TemporaryDirectory() as tmpdir:
        json_path = str(Path(tmpdir) / 'test.json')
        actual = export_json(sequences, json_path, 'WO_TEST')
        data = json.loads(Path(actual).read_text(encoding='utf-8'))
        assert data['patent'] == 'WO_TEST', f'Expected WO_TEST, got {data["patent"]}'
        assert data['total'] == 1, f'Expected 1, got {data["total"]}'
        assert data['sequences'][0]['seq_id'] == 'SEQ001'
        print('[PASS] test_export_json')


if __name__ == '__main__':
    print('=== patent-sequence-exporter selftest ===')
    test_parse_patent_ids()
    test_parse_seq_id_range()
    test_read_api_key()
    test_ndjson_append()
    test_batch_state()
    test_timestamped_path()
    test_export_fasta()
    test_export_json()
    print('\n=== Results: all tests passed ===')
