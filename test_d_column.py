"""D sutunu regression testi: supplier null gelince 'S2A BP' default'a dusmeli."""
from app import build_d_value


def test_supplier_null_falls_back_to_default():
    # Gemini anti-uydurma prompt'u supplier'i null birakabilir
    assert build_d_value({'supplier': None, 'cmd_code': '7304'}) == 'S2A BP\nCMD-7304'


def test_supplier_null_no_cmd():
    assert build_d_value({'supplier': None}) == 'S2A BP'


def test_supplier_missing_key():
    assert build_d_value({'cmd_code': '7011'}) == 'S2A BP\nCMD-7011'


def test_supplier_real_value_kept():
    assert build_d_value({'supplier': 'Acme', 'cmd_code': '7011'}) == 'Acme\nCMD-7011'


def test_empty_string_supplier_falls_back():
    assert build_d_value({'supplier': '', 'cmd_code': '7011'}) == 'S2A BP\nCMD-7011'


if __name__ == '__main__':
    for name, fn in list(globals().items()):
        if name.startswith('test_'):
            fn()
            print(f'OK {name}')
    print('ALL PASS')
