# Ledger Block Creator — El Yazısından Excel'e Yeni Blok Ekleme

## Problem

Mevcut CubeLogReader uygulaması el yazısı formu Gemini AI ile okuyup **var olan** Excel bloklarına sadece Weight (M) ve Load (N) değerlerini yazıyor. Kullanıcı ise el yazısı formdan okunan verilerle **komple yeni bloklar oluşturulmasını** istiyor — "Concrete" sheet'inde tablonun sonuna tüm sütunları doldurarak yeni satırlar eklemek.

---

## Detaylı Blok Deseni Analizi

Fotoğrafları satır satır inceledim. **Blok yapısı değişken** — her bloğun satır sayısı, mould sayısı ve test türü farklı olabiliyor. 3 örnek:

### Blok Tipi 1 — Kısa Blok (Cube 716, 7 satır)
Basit yapı: 3×7-gün + CMD satırı + 3×28-gün = 7 satır

| # | Mould | Supplier | Testing Date | Age | Tested By | Not |
|---|-------|----------|-------------|-----|-----------|-----|
| 1 | 93 | S2A BP | 02.06.26 | 7 | | İlk satır: Cube No, Sample Mark, Location, Section, Batch, Grade, Sampling Date, Sampled By yazılır |
| 2 | 138 | | " | 7 | | |
| 3 | 24 | CMD | " | 7 | | CMD satırı |
| 4 | 75 | 7410 | 23.06.26 | 28 | | CMD kodu |
| 5 | 63 | | " | 28 | | |
| 6 | 74 | | " | 28 | | |

> [!NOTE]
> Bu tipte Site satırı, WP satırı, 2-Day satırı yok.

---

### Blok Tipi 2 — Orta Blok (Cube 715, 10 satır)
3×7-gün + CMD satırı + 3×28-gün + 3×WP = ~10 satır

| # | Mould | Supplier | Testing Date | Age | Tested By | Not |
|---|-------|----------|-------------|-----|-----------|-----|
| 1 | 139 | S2A BP | 02.06.26 | 7 | | Cube No, Sample Mark, Location, Section, Batch, Grade, Sampling Date |
| 2 | 90 | | " | 7 | | |
| 3 | 157 | CMD | " | 7 | | CMD satırı |
| 4 | 136 | 7302 | 23.06.26 | 28 | | CMD kodu + 28-gün test tarih değişimi |
| 5 | 71 | | " | 28 | | Sampled By bu civarda |
| 6 | 48 | | " | 28 | | |
| 7 | 180 | | 23.06.26 | WP | | WP = "Waiting Period" satırları |
| 8 | 29 | | " | WP | | |
| 9 | 130 | | " | WP | | |

> [!NOTE]
> WP satırlarında test tarihi var ama Age sütununda "WP" yazıyor, Weight/Load boş.

---

### Blok Tipi 3 — Uzun Blok (Cube 714, 15 satır)
3×7-gün + CMD + 3×28-gün + 2×WP + 1×2Day + 3×Site + 1×F/T = ~15 satır

| # | Mould | Supplier | Testing Date | Age | Tested By | Not |
|---|-------|----------|-------------|-----|-----------|-----|
| 1 | 171 | S2A BP | 02.06.26 | 7 | | Cube No, Sample Mark, Location, Section, Batch, Grade, Sampling Date |
| 2 | 92 | | " | 7 | | |
| 3 | 117 | CMD | " | 7 | | CMD satırı |
| 4 | 02 | 7008 | 23.06.26 | 28 | | CMD kodu + 28-gün test tarih |
| 5 | 16 | | " | 28 | | Sampled By bu civarda |
| 6 | 186 | | " | 28 | | |
| 7 | 48 | | 23.06.26 | WP | | WP satırı |
| 8 | 44 | | " | WP | | WP satırı |
| 9 | 31 | | " | (boş) | | — boş age |
| 10 | 60 | | 28.05.26 | 2 Day | | 2-günlük test |
| 11 | 08 | | | | Site | Sahaya gönderilen |
| 12 | 121 | | | | Site | Sahaya gönderilen |
| 13 | — | | | | Site | Mould no yok, 3. site satırı |
| 14 | (boş) | | 23.06.26 | F/T | (15 Adet) | **Özet satırı** — bu farklı! |

> [!WARNING]
> **Kritik farklılıklar bu blokta:**
> - WP satırları (Waiting Period)
> - "2 Day" = 2 günlük test
> - "Site" satırları — Testing Date ve Age yok, Tested By'da "Site" yazıyor
> - Son satır **F/T (15 Adet)** — toplu test özetidir. Age sütununda "F/T", Load sütununda "(15 Adet)" yazıyor. Bu satırda mould no yok.
> - Bazı satırlarda age tamamen boş

---

### Blok Tipi 4 — Çok Kısa Blok (Cube 713, ~6 satır, ayrı fotoğraftan)
Önceki fotoğraftaki Cube 713 (G26-CON-759): 3×7-gün + CMD + 3×28-gün

| # | Mould | Supplier | Testing Date | Age |
|---|-------|----------|-------------|-----|
| 1 | 05 | S2A BP | 02.06.26 | 7 |
| 2 | 61 | | " | 7 |
| 3 | 80 | CMD | " | 7 |
| 4 | 132 | 7015 | 23.06.26 | 28 |
| 5 | 115 | | " | 28 |
| 6 | 96 | | " | 28 |

---

### Excel Karşılaştırması — Blok 714 Detay

Excel'deki satırlar (5907-5917) ile el yazısı karşılaştırma:

| Excel Row | C (Mould) | D (Supplier) | I (Sampling) | K (Testing) | L (Age) | P (Tested By) |
|-----------|-----------|--------------|-------------|-------------|---------|----------------|
| 5907 | 117 | | 26/05/2026 | 02/06/2026 | 7 | |
| 5908 | 2 | | 26/05/2026 | 23/06/2026 | 28 | |
| 5909 | 16 | S2A BP | 26/05/2026 | 23/06/2026 | 28 | |
| 5910 | 186 | CMD-7008 | 26/05/2026 | 23/06/2026 | 28 | |
| 5911 | 48 | | 26/05/2026 | 23/06/2026 | 28 | |
| 5912 | 44 | | 26/05/2026 | 23/06/2026 | WP | |
| 5913 | 31 | | 26/05/2026 | 23/06/2026 | WP | |
| 5914 | 60 | | 26/05/2026 | 28/05/2026 | 2 | Site |
| 5915 | 8 | | 26/05/2026 | | | Site |
| 5916 | 121 | | 26/05/2026 | | | Site |
| 5917 | | | 26/05/2026 | 23/06/2026 | F/T | (15 Adet) |

> [!IMPORTANT]
> **Kritik gözlemler (Excel vs el yazısı farklılıkları):**
> 1. Excel'de **Mould sırası farklı** — el yazısından farklı sıralanmış! (ör: 171, 92 el yazısında var ama Excel'de 117'den başlıyor, 7-day moulds Excel'de görünmüyor. Aslında Row 5907'de cube 714 için 7-gün moulds eksik!)
> 2. Excel'de **D sütunu** "S2A BP" ve "CMD-7008" şeklinde birleşik yazılıyor
> 3. **"2 Day"** el yazısında → Excel'de Age=**2**, Tested By=**Site**
> 4. **F/T satırı**: Age="F/T", Load sütununda "(15 Adet)"
> 5. **WP satırları**: Age sütununda "WP"
> 6. **Sampled By (J)**: Bloğun ortasındaki satırda (ör: row 5911 = S.B)
> 7. Her satırda **Date of Sampling (I)** tekrarlanıyor
> 8. **O (Compressive Strength)**: Her yerde **0.00** — muhtemelen formül veya sabit

Aslında dikkatli bakınca, Excel'deki **Cube 714 satırları (5907-5917)** toplamda **11 satır**. El yazısında da yaklaşık 15 row var (bazıları Excel'e 7-gün satırları olarak girilmiş fakat bu anlık fotoğrafta mould 171, 92 üstte ayrı satırlarda). 

**Doğru yorum**: Cube 714 Excel'de satır 5907'den başlıyor ve satır 5917'de bitiyor. Tüm mould'lar sırasıyla: 117, 2, 16, 186, 48, 44, 31, 60, 8, 121, (boş). Bu 7-day moulds (171, 92) Excel'de **yok** çünkü onlar muhtemelen önceki satırlardaymış. 

Hayır, tekrar bakıyorum: Cube 714 A sütununda satır **5911**'de yazıyor. O hâlde:
- 5907-5910 arasındaki mould'lar (117, 2, 16, 186) → bunlar Cube 714'ün 7-day satırları
- 5911'de A=714, B=G26-CON-0761 → başlık satırı
- Ve aşağı doğru devam ediyor

Bu karmaşık yapı, Gemini'nin el yazısından okuduğu tüm satırları sırayla Excel'e aktarmamız gerektiğini gösteriyor.

---

## Proposed Changes

### 1. Yeni Gemini Prompt — Blok Oluşturma Modu

#### [MODIFY] [reader.py](file:///c:/Users/Yafka/Desktop/CubeLogReader/reader.py)

Mevcut `PROMPT`'a dokunmadan, **yeni bir prompt** (`BLOCK_PROMPT`) ekliyoruz. Bu prompt tüm blok bilgilerini çıkarır:

**Yeni JSON çıktı formatı:**
```json
{
  "cubes": [
    {
      "cube_no": "716",
      "sample_mark": "G26-CON-763",
      "concrete_supplier": "S2A BP",
      "cmd_code": "7410",
      "site_location": "Km:35+935 - KM 37+960 Section-1 Irrigation Channel Concrete",
      "section": "1",
      "batch_ticket": "15080",
      "c_grade": "C25/30",
      "date_of_sampling": "26.05.26",
      "sampled_by": "Y.A",
      "rows": [
        {"mould_no": "93", "date_of_testing": "02.06.26", "age": "7", "row_type": "7d"},
        {"mould_no": "138", "date_of_testing": "02.06.26", "age": "7", "row_type": "7d"},
        {"mould_no": "24", "date_of_testing": "02.06.26", "age": "7", "row_type": "7d"},
        {"mould_no": "75", "date_of_testing": "23.06.26", "age": "28", "row_type": "28d"},
        {"mould_no": "63", "date_of_testing": "23.06.26", "age": "28", "row_type": "28d"},
        {"mould_no": "74", "date_of_testing": "23.06.26", "age": "28", "row_type": "28d"}
      ]
    },
    {
      "cube_no": "714",
      "sample_mark": "G26-CON-761",
      "concrete_supplier": "S2A BP",
      "cmd_code": "7008",
      "site_location": "Beam Yard Km:5+800 B0080 Bridge Girder L2G",
      "section": "2A",
      "batch_ticket": "15062",
      "c_grade": "C40/50",
      "date_of_sampling": "26.05.26",
      "sampled_by": "S.B",
      "rows": [
        {"mould_no": "171", "date_of_testing": "02.06.26", "age": "7", "row_type": "7d"},
        {"mould_no": "92", "date_of_testing": "02.06.26", "age": "7", "row_type": "7d"},
        {"mould_no": "117", "date_of_testing": "02.06.26", "age": "7", "row_type": "7d"},
        {"mould_no": "02", "date_of_testing": "23.06.26", "age": "28", "row_type": "28d"},
        {"mould_no": "16", "date_of_testing": "23.06.26", "age": "28", "row_type": "28d"},
        {"mould_no": "186", "date_of_testing": "23.06.26", "age": "28", "row_type": "28d"},
        {"mould_no": "48", "date_of_testing": "23.06.26", "age": "WP", "row_type": "wp"},
        {"mould_no": "44", "date_of_testing": "23.06.26", "age": "WP", "row_type": "wp"},
        {"mould_no": "31", "date_of_testing": null, "age": null, "row_type": "wp"},
        {"mould_no": "60", "date_of_testing": "28.05.26", "age": "2", "row_type": "other"},
        {"mould_no": "08", "date_of_testing": null, "age": null, "row_type": "site"},
        {"mould_no": "121", "date_of_testing": null, "age": null, "row_type": "site"},
        {"mould_no": null, "date_of_testing": null, "age": null, "row_type": "site"},
        {"mould_no": null, "date_of_testing": "23.06.26", "age": "F/T", "row_type": "ft", "ft_note": "15 Adet"}
      ]
    }
  ]
}
```

**Yeni fonksiyon:** `read_notebook_for_blocks(file_path, progress_cb)` — mevcut `read_notebook()` ile aynı altyapı ama `BLOCK_PROMPT` kullanır, farklı cache prefix ile saklar.

> [!NOTE]
> Mevcut `read_notebook()` ve `PROMPT` hiç değişmeyecek. Backward compatibility korunacak.

---

### 2. Yeni Blok Yazıcı

#### [NEW] [block_writer.py](file:///c:/Users/Yafka/Desktop/CubeLogReader/block_writer.py)

"Concrete" sheet'ine yeni bloklar ekleyen modül.

**Ana fonksiyonlar:**

```python
def find_last_data_row(ws) -> int
    """Concrete sheet'te son dolu satırı bul (A-L sütunları taranarak)."""

def write_new_block(ws, cube_data: dict, start_row: int) -> dict
    """Bir cube'un tüm satırlarını start_row'dan itibaren yaz.
    
    Yazma mantığı — her satır için:
    - A: Cube No (sadece ilk satırda)
    - B: Sample Mark (sadece ilk satırda)
    - C: Mould No
    - D: Concrete Supplier düzeni:
        * İlk 7d satırda: "S2A BP" (veya ana supplier)
        * CMD satırında: "CMD"  
        * CMD kodu satırında: "7410" (veya kod)
        * Diğer satırlarda: boş
    - E: Site Location (sadece ilk satırda, merge etmeden)
    - F: Section (sadece ilk satırda)
    - G: Batch Ticket (sadece ilk satırda)
    - H: C Grade (sadece ilk satırda)
    - I: Date of Sampling (her satırda, DD/MM/YYYY formatında)
    - J: Sampled by (bloğun orta satırında)
    - K: Date of Testing (7d/28d/WP/F/T satırlarında)
    - L: Age (7, 28, WP, 2, F/T vb.)
    - M: Weight — boş
    - N: Load — boş (F/T satırında "(X Adet)" yazılır)
    - O: Compressive Strength — 0.00
    - P: Tested By — Site satırlarında "Site"
    - Q: Engineer Signature — boş
    
    Returns: {wrote: [...], errors: [...], start_row, end_row}
    """

def write_multiple_blocks(ws, cubes: list[dict]) -> list[dict]
    """Birden fazla bloğu sırayla yaz, her birini son satırın altına ekle."""
```

**D sütunu (Concrete Supplier) yazma kuralları — Excel'den gözlemlenen:**

El yazısı formda Supplier iki satırda yazılıyor:
1. Ana tedarikçi: "S2A BP" (ilk satırda)
2. CMD kodu: "CMD" + "\n7008" (ayrı satırlarda)

Excel'de bu D sütununa şöyle yansıyor:
- Bazı bloklarda "S2A BP" ve "CMD-7008" olarak
- Bazılarında ilk satırda "S2A BP", alt satırda "CMD" ayrı, onun altında "7008" ayrı

Güvenli yaklaşım: **Excel'deki mevcut formatı taklit et** — son blokların D sütunundaki yazım stilini oku ve aynısını uygula.

---

### 3. Önizleme ve Düzenleme UI

#### [NEW] [block_preview.py](file:///c:/Users/Yafka/Desktop/CubeLogReader/block_preview.py)

CustomTkinter tabanlı önizleme penceresi:

- Her cube için bir **kart**: Cube No, Sample Mark, Location, Section, Batch, Grade, Dates
- Her kart içinde **satır tablosu**: Mould No, Testing Date, Age, Row Type
- Tüm alanlar düzenlenebilir (CTkEntry)
- Satır ekle / sil butonları (blok boyutu ayarlanabilir)
- Fotoğraf önizleme paneli (sol tarafta, mevcut PreviewWindow gibi)
- **"Write Blocks to Excel"** butonu

---

### 4. Ana UI Entegrasyonu

#### [MODIFY] [main.py](file:///c:/Users/Yafka/Desktop/CubeLogReader/main.py)

Ana ekrana **"Add New Blocks"** butonu ekleme:

Mevcut flow:
```
[Pick File] → Gemini reads → [PreviewWindow] → Write to Excel (values only)
```

Yeni flow:
```
[Pick File] → [Add New Blocks] butonu → Gemini reads (BLOCK_PROMPT) → [BlockPreviewWindow] → Write blocks to Excel
```

---

## Resolved Decisions (2026-05-29)

> [!NOTE]
> **Referans blok: Cube 716, satır 25927-25932** (kullanıcı tarafından onaylandı). Tüm yeni bloklar bu formatı birebir taklit etmeli.

### ✅ Karar 1 — D sütunu (Supplier) formatı
İki merge edilmiş hücre kullanılacak:
- **Üst merge** = supplier adı (ör: "S2A BP") → 7-gün non-CMD satırları kapsar
- **Alt merge** = "CMD-XXXX" birleşik formatı (ör: "CMD-7410") → CMD satırı + 28-gün satırları kapsar

Blok 716 örneği:
- `D25927:D25928` = "S2A BP" (2 satır)
- `D25929:D25932` = "CMD-7410" (4 satır)

### ✅ Karar 2 — O sütunu (Compressive Strength)
**Formül** yazılacak. Bir önceki bloğun O sütunundaki formülü okuyup yeni satırlara kopyalayacağız (openpyxl ile relative reference güncellemesi).

### ✅ Karar 3 — Hücre formatlama: HYBRID yaklaşım

> [!IMPORTANT]
> Son blok yapısı yeni bloktan farklı olabileceği için, "son bloğu birebir kopyala" çalışmaz. Bunun yerine **stil** mevcut bloktan kopyalanır, **merge'ler** kuraldan hesaplanır.

**Strateji:**

**1. Stil kopyalama (cell-level):**
- Sheet'teki mevcut bir bloktan (en son dolu blok) her sütun için **referans hücre stili** oku: font, fill, border, alignment, number_format
- Yeni blok hücrelerine bu stilleri uygula (openpyxl `copy()` ile)
- Row heights ve column widths zaten sheet seviyesinde, dokunma

**2. Merge aralıkları (kuraldan):**
```
n = blok satır sayısı (Gemini'nin döndürdüğü rows listesi uzunluğu)
start = first_new_row, end = start + n - 1
cmd_idx = CMD satırının index'i (yoksa None)

# Tüm blok merge:
merge(A, start, end)
merge(B, start, end)
merge(E, start, end)  # boş kalacak
merge(F, start, end)
merge(G, start, end)
merge(H, start, end)
merge(J, start, end)
merge(P, start, end)  # boş kalacak

# D iki parça:
if cmd_idx is not None:
    merge(D, start, start + cmd_idx - 1)   # üst: supplier (S2A BP)
    merge(D, start + cmd_idx, end)         # alt: CMD-XXXX
else:
    merge(D, start, end)                    # tek parça

# C, I, K, L, M, N, O, Q: merge yok
```

**3. Veri yazma:**
- A: Cube No, B: Sample Mark
- C: her satıra mould no
- D üst merge: supplier (ör. "S2A BP")
- D alt merge: "CMD-XXXX" (ör. "CMD-7410")
- F: Section, G: Batch Ticket, H: C Grade
- I: her satırda sampling date (DD/MM/YYYY)
- J: Sampled by
- K: testing date — **otomatik hesaplanır** (7d → sampling+7, 28d → sampling+28, WP/Site/F/T için Gemini'den gelen değer)
- L: age (7, 28, WP, F/T, 2, vb.)
- **E ve P: BOŞ** (merge edilmiş ama içerik yok)
- M, N, Q: boş
- **O: formül** — mevcut bloktan oku, relative shift ile yeni satıra uygula (ör. `=N5907/(150*150)*1000` → `=N{new_row}/(150*150)*1000`)

**Genel merge kuralları (blok boyutundan bağımsız):**

| Sütun | Merge davranışı | İçerik | Not |
|-------|-----------------|--------|-----|
| A | tüm blok (start→end) | Cube No | Gemini doldurur |
| B | tüm blok | Sample Mark | Gemini doldurur |
| C | merge yok, her satır ayrı | Mould No | Gemini doldurur |
| D | **iki parça**: üst = supplier (CMD'den önce), alt = "CMD-XXXX" (CMD'den sonra) | Supplier/CMD | Gemini doldurur |
| **E** | **tüm blok merge — ama BOŞ bırakılır** | Site Location | **Kullanıcı manuel doldurur** |
| F | tüm blok | Section | Gemini doldurur |
| G | tüm blok | Batch Ticket | Gemini doldurur |
| H | tüm blok | C Grade | Gemini doldurur |
| I | her satırda tekrar (merge yok) — *doğrulanacak* | Sampling Date | Gemini doldurur |
| J | tüm blok | Sampled by | Gemini doldurur |
| K | her satır ayrı | Testing Date | Otomatik hesaplanır |
| L | her satır ayrı | Age | Gemini doldurur |
| M, N | merge yok | Weight / Load | Boş |
| O | merge yok | Compressive Strength | Formül (önceki bloktan kopyala) |
| **P** | **tüm blok merge — BOŞ** | Tested By Signature | Sonradan imzalanır |
| Q | merge yok | Engineer Signature | Boş |

**Dinamik merge hesaplama algoritması (Gemini'nin döndürdüğü `rows` listesine göre):**
```
n = len(rows)
start = first_data_row
end = start + n - 1

# Whole-block merges:
merge(A, start, end)
merge(B, start, end)
merge(E, start, end)  # ve F, G, H, J
merge(F, start, end)
merge(G, start, end)
merge(H, start, end)
merge(J, start, end)

# D sütunu iki parça: CMD satırının indexini bul
cmd_idx = first index where row.row_type == "cmd" or row.supplier_marker == "CMD"
if cmd_idx is not None:
    merge(D, start, start + cmd_idx - 1)   # üst: supplier (S2A BP)
    merge(D, start + cmd_idx, end)         # alt: "CMD-XXXX"
else:
    merge(D, start, end)                    # CMD yoksa tek parça

# C, I, K, L, M-Q: merge yok
```

**Referans örnek — Blok 716 (sadece kısa blok formatını göstermek için):**
- 6 satır (3×7d + 3×28d, CMD ortada)
- A25927:A25932 (Cube No 716)
- D25927:D25928 = "S2A BP", D25929:D25932 = "CMD-7410"
- Diğer bloklar farklı satır sayısına sahip olabilir; algoritma uyum sağlar.

### ✅ Karar 4 — Ayrı program
Bu özellik **tamamen ayrı bir program** olarak geliştirilecek. Klasör: `c:\Users\Yafka\Desktop\Last pages Program\`. CubeLogReader'a dokunulmayacak.

### ✅ Karar 5 — Test tarihleri
**Otomatik hesaplanacak**:
- 7-gün test tarihi = `date_of_sampling + 7 gün`
- 28-gün test tarihi = `date_of_sampling + 28 gün`

Gemini sadece sampling date'i okuyacak, test tarihlerini hesaplamayacak. Bu Gemini'nin hata yapma alanını azaltır.

## Verification Plan

### Automated Tests
- El yazısı form fotoğrafını Gemini'ye gönderip block JSON çıktısını doğrulama
- Excel'e yazılan blokların mevcut bloklarla sütun eşleşmesini test etme
- Farklı blok tipleri (kısa 7 satır, uzun 15 satır) için yazma testi
- Mevcut CubeLogReader işlevselliğinin bozulmadığını kontrol etme

### Manual Verification
- Gerçek Excel dosyasına test bloğu ekleme
- Eklenen bloğun mevcut bloklarla aynı formatta olduğunu görsel olarak doğrulama (kullanıcıdan onay)
- Birden fazla blok ekleme senaryosu test etme
