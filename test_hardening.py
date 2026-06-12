"""Saglamlastirma testleri: satir uydurma (tekrar mould) ve D/CMD uyarisi."""
from app import drop_duplicate_mould_rows, build_d_value, drop_phantom_cmd_rows


def test_phantom_cmd_dropped_cmd_code_recovered():
    rows = [
        {'mould': 8, 'row_type': '7d'},
        {'mould': None, 'row_type': 'cmd', 'cmd_code': '7410'},  # fantom extra
        {'mould': 187, 'row_type': '28d'},
    ]
    out, recovered = drop_phantom_cmd_rows(rows)
    assert [r['mould'] for r in out] == [8, 187]
    assert recovered == '7410'


def test_real_cmd_row_with_mould_kept():
    # cube 795 gibi: cmd satiri gercek mould'la gelir -> KORUNUR
    rows = [{'mould': 34, 'row_type': '7d'}, {'mould': 73, 'row_type': 'cmd', 'age': 7}]
    out, recovered = drop_phantom_cmd_rows(rows)
    assert [r['mould'] for r in out] == [34, 73]
    assert recovered is None



def _logs():
    msgs = []
    return msgs, lambda m='': msgs.append(m)


def test_drops_duplicate_mould_keeps_first():
    rows = [
        {'mould': 5, 'row_type': '7d'},
        {'mould': 69, 'row_type': '7d'},
        {'mould': 5, 'row_type': '28d'},   # uydurma: mould 5 tekrar
    ]
    msgs, log = _logs()
    out = drop_duplicate_mould_rows(rows, cube_no=713, log=log)
    assert [r['mould'] for r in out] == [5, 69]
    assert out[0]['row_type'] == '7d'  # ilk gecis korundu
    assert any('UYDURMA' in m for m in msgs)


def test_keeps_all_null_mould_rows():
    # ft satiri / okunamayan mould = null; bunlar deduplenmez
    rows = [
        {'mould': None, 'row_type': 'ft'},
        {'mould': None, 'row_type': 'ft'},
        {'mould': 12, 'row_type': '28d'},
    ]
    msgs, log = _logs()
    out = drop_duplicate_mould_rows(rows, cube_no=1, log=log)
    assert len(out) == 3


def test_no_duplicates_unchanged():
    rows = [{'mould': 1, 'row_type': '7d'}, {'mould': 2, 'row_type': '28d'}]
    msgs, log = _logs()
    out = drop_duplicate_mould_rows(rows, cube_no=1, log=log)
    assert len(out) == 2
    assert not any('UYDURMA' in m for m in msgs)


def test_d_value_still_works():
    assert build_d_value({'supplier': None, 'cmd_code': '7304'}) == 'S2A BP\nCMD-7304'


if __name__ == '__main__':
    for name, fn in list(globals().items()):
        if name.startswith('test_'):
            fn()
            print(f'OK {name}')
    print('ALL PASS')
