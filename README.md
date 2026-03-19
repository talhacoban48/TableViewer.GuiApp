# Table Viewer

Excel ve CSV dosyalarını görüntülemek, düzenlemek ve kaydetmek için PyQt5 tabanlı masaüstü uygulaması.

---

## İndir

**[TableViewer-Setup-v1.0.0.exe](https://github.com/talhacoban48/TableViewer.GuiApp/releases/download/v.1.0.0/TableViewer-Setup-v1.0.0.exe)**


---

## Özellikler

### Dosya Desteği
- `.xlsx`, `.xls`, `.csv` dosyalarını açar
- Çoklu sayfa (sheet) desteği — alt sekme çubuğu ile sayfalar arası geçiş
- CSV dosyalarını otomatik encoding tespiti ile okur (`utf-8-sig`, `utf-8`, `cp1254`, `latin-1`, `iso-8859-9`)
- Dosyayı pencereye **sürükle-bırak** ile açabilirsiniz
- **Excel (.xlsx)** veya **CSV (.csv)** olarak kayıt
- Windows dosya ilişkilendirmesi — `.xlsx`, `.xls`, `.csv` dosyalarını çift tıkla açacak şekilde kayıt

### Başlangıç
- Uygulama açıldığında herhangi bir dosya açılmamışsa otomatik olarak **boş 10 satır × 5 sütun** bir tablo gösterilir

### Tablo Düzenleme
- Hücrelere çift tıklayarak veya doğrudan yazarak düzenleme
- **Satır ekle** — seçili satırın altına boş satır ekler
- **Satır sil** — seçili satır(lar)ı siler
- **Sütun ekle** — seçili sütunun sağına isim vererek yeni sütun ekler
- **Sütun sil** — seçili sütun(lar)ı siler
- Yapısal değişiklikler (ekle/sil) kayıt sırasında hücre biçimlendirmesini doğru konuma taşır

### Kes / Kopyala / Yapıştır
- **Ctrl+C** — seçili bölgeyi kopyalar (biçimlendirme dahil)
- **Ctrl+X** — seçili bölgeyi keser; yapıştırma sonrasında kaynak hücreler temizlenir
- **Ctrl+V** — seçili başlangıç hücresine yapıştırır (biçimlendirme korunur)
- **Escape** — kopyala/kes işlemini iptal eder
- Kopyalanan/kesilen bölge etrafında **marching ants** (hareketli noktalı çerçeve) animasyonu
- Harici uygulamalardan (Excel, Notepad vb.) sekme-ayrımlı metin yapıştırma desteği
- Kopyalanan metin aynı zamanda sistem panosuna da yazılır (harici yapıştırma için)

### Geri Al (Undo)
- **Ctrl+Z** — son hücre düzenlemesini geri alır (sınırsız adım)

### Biçimlendirme Araç Çubuğu
| Kontrol | İşlev |
|---|---|
| **B** | Kalın (Bold) |
| *I* | İtalik |
| Size | Font boyutu (6–72 pt) |
| **A** | Metin rengi |
| ■ | Hücre arka plan rengi |

- Biçimlendirme seçili tüm hücrelere uygulanır
- Farklı bir hücreye tıklandığında araç çubuğu o hücrenin biçimini yansıtır
- Biçimlendirme `.xlsx` olarak kaydedildiğinde korunur (orijinal font ailesi, alt çizgi vb. bozulmadan, sadece değiştirilen özellikler yazılır)

### Sıralama & Filtreleme
Her sütun başlığında sağ tarafta üç ikon bulunur:

| İkon | İşlev |
|---|---|
| Sırala | Sütunu artan/azalan sıraya göre sıralar |
| Filtrele | Filtre popup'ını açar |
| Temizle | Aktif filtreyi kaldırır (sadece filtre varken görünür) |

**Filtre Popup:**
- Tüm benzersiz değerleri listeler, checkbox ile seçim
- Arama kutusu ile değer arama
- "(Tümünü Seç)" üçlü durum checkbox'ı
- Sayısal sütunlarda ek **Sayı Filtresi** sekmesi: `=`, `≠`, `>`, `>=`, `<`, `<=`, `between`

### Global Arama
- Üst kısımdaki arama kutusuna yazarak tüm sütunlarda eş zamanlı arama
- 500 ms debounce (geciktirme) ile performans koruması
- Temizleme butonu (×) metin girildiğinde görünür

### Durum Çubuğu
- Dosya adı, toplam satır sayısı ve sütun sayısı
- Filtre aktifken: `görünen / toplam` satır bilgisi

---

## Proje Yapısı

```
TableViewer.GuiApp/
├── main.py                 # Giriş noktası
├── assets/                 # Uygulama ikonları
└── tableviewer/            # Ana paket
    ├── __init__.py
    ├── constants.py        # Sabitler (uzantılar, encoding listesi, rol sabitleri)
    ├── utils.py            # Yardımcı fonksiyonlar (ikon yükleme, tür dönüşümü)
    ├── models.py           # MultiColumnFilterProxyModel
    ├── filter_popup.py     # FilterPopup (Excel tarzı filtre penceresi)
    ├── overlay.py          # MarchingAntsOverlay (kopyala/kes animasyonu)
    ├── header.py           # SortFilterHeaderView (özel başlık görünümü)
    └── app.py              # TableViewerApp (ana pencere)
```

---

## Gereksinimler

```
PyQt5
pandas
openpyxl
```

## Kurulum & Çalıştırma

```bash
pip install PyQt5 pandas openpyxl
python main.py
# veya bir dosyayla doğrudan açmak için:
python main.py dosya.xlsx
```
