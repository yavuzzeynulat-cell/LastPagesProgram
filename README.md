# Last Pages Program - Cube Block Creator (test)

El yazisi concrete sampling formundan Excel "Concrete" sheet'ine komple yeni
bloklar yazar. Mevcut blok formatini (merge, stil, formul) birebir taklit eder.

## Hizli kullanim (is bilgisayari)

### Tek seferlik kurulum
1. Python 3.10+ kurulu olsun (yoksa python.org'dan indir)
2. `run.bat`'i bu klasore at (veya tum klasoru indir)

### Her sefer
- `run.bat`'a Excel dosyasini suruklesin **veya** cift tikla, sorunca yolu yapistir
- Her calistirmada otomatik olarak en yeni surumu GitHub'dan indirir
- Excel yolunu bir kere kaydeder (`config.txt`), tekrar sormaz

### Cikti
`<excel_adi>.with_blocks.xlsx` — orijinal dosya degisemez.

## Dosyalar

| Dosya | Ne ise yarar |
|-------|--------------|
| `run.bat` | Tek tik calistir + auto-update |
| `create_blocks.py` | Ana writer scripti |
| `blocks_data.json` | Blok verileri (su an test icin 713-716) |
| `config.txt` | Excel dosya yolu (otomatik kaydedilir) |

## Test yontemi

1. Excel'in YEDEGINI AL
2. Test Excel'de **sadece 713-716 bloklarini sil** (698-712'ye dokunma)
3. `run.bat` cift tikla, Excel yolunu yapistir
4. Cikan dosyayi ac, 713-716 yerli yerinde mi bak:
   - A/B/E/F/H/J/P merge mi
   - D iki parca mi (supplier + CMD-XXXX)
   - K sutunu otomatik hesaplanmis mi (sampling+7 / sampling+28)
   - O sutunu formul mu
   - Site satirinda P="Site", F/T satirinda N="(15 Adet)"

## Sorun cikarsa
Hatayi (komut ciktisini) yapistir, hizli duzeltirim ve yeni surumu push'larim.
`run.bat` bir sonraki calistirmada otomatik en yenisini indirir.

## Repo
https://github.com/yavuzzeynulat-cell/LastPagesProgram
