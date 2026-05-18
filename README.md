# 📋 Puantaj Sistemi

Küçük ve orta ölçekli işletmeler için geliştirilmiş, masaüstü tabanlı **personel devam takibi ve maaş hesaplama** uygulaması.

---

## 🖥️ Ekran Görüntüsü

> Flask + pywebview ile çalışan, koyu temalı modern bir arayüze sahiptir.

---

## 🚀 Özellikler

- 👤 **Personel Yönetimi** — Personel ekleme, düzenleme, pasife alma ve yeniden aktifleştirme
- 🗓️ **Puantaj Takibi** — Günlük çalışma saati ve gün türü (normal, pazar, tatil, yıllık izin, rapor, ücretsiz izin) girişi
- 💰 **Otomatik Maaş Hesaplama** — Saatlik ücret, fazla mesai (%50 zam), banka/elden ödeme ayrımı
- 🏖️ **Resmi Tatil Yönetimi** — Tam gün ve yarım gün tatil tanımlama
- 💳 **Avans Takibi** — Personel bazında aylık avans kaydı
- 📊 **Excel Raporu** — Aylık puantaj ve maaş detaylarını `.xlsx` olarak dışa aktarma
- 🔄 **Geri Alma (Undo)** — Son işlemleri tek tıkla geri alma
- 💾 **Veritabanı Yedekleme** — Otomatik ve manuel yedek alma / geri yükleme
- 🔒 **Tek Örnek Kontrolü** — Uygulamanın birden fazla açılmasını engeller (Windows Mutex)

---

## 🛠️ Teknolojiler

| Katman | Teknoloji |
|---|---|
| Backend | Python 3, Flask, Flask-CORS |
| Veritabanı | SQLite3 |
| Arayüz | HTML, CSS, Vanilla JS (pywebview) |
| Raporlama | openpyxl |
| Masaüstü | pywebview |

---

## ⚙️ Kurulum

### Gereksinimler

```bash
pip install flask flask-cors openpyxl pywebview
```

### Çalıştırma

```bash
python app.py
```

Uygulama otomatik olarak masaüstü penceresi açar. `pywebview` yüklü değilse varsayılan tarayıcıda `http://localhost:5000` adresini açar.

---

## 📁 Proje Yapısı

```
puantaj-sistemi/
├── app.py          # Flask backend ve iş mantığı
├── index.html      # Tek sayfalık uygulama arayüzü
├── puantaj.db      # SQLite veritabanı (otomatik oluşur)
└── yedekler/       # Veritabanı yedekleri (otomatik oluşur)
```

---

## 🧮 Maaş Hesaplama Mantığı

| Gün Türü | Hesaplama |
|---|---|
| Normal gün | Girilen saat normal; 7,5 saati aşan kısım fazla mesai |
| Pazar günü | `0` girilirse → 7,5 saat normal; değer girilirse o kadar normal |
| Tam tatil | Çalışılan saat → 7,5 normal + aşan kısım fazla mesai |
| Yarım tatil | Çalışılan saat → 7,5 normal + 4 üstü fazla mesai |
| Yıllık izin | 7,5 saat normal sayılır |
| Rapor / Ücretsiz izin | Hesaba dahil edilmez |

**Fazla mesai çarpanı:** `1,5x`  
**Aylık standart saat:** `225 saat`  
**Tavan saat (günlük):** `7,5 saat`

---

## 🔌 API Referansı

| Method | Endpoint | Açıklama |
|---|---|---|
| GET | `/api/personel` | Aktif personel listesi |
| POST | `/api/personel` | Yeni personel ekle |
| PUT | `/api/personel/<id>` | Personel güncelle |
| DELETE | `/api/personel/<id>` | Personeli kalıcı sil |
| POST | `/api/personel/<id>/pasife_al` | Pasife al |
| POST | `/api/personel/<id>/aktife_al` | Aktife al |
| GET/POST | `/api/maas/<id>` | Maaş geçmişi |
| GET/POST | `/api/puantaj` | Puantaj kayıtları |
| GET/POST | `/api/tatil` | Tatil tanımları |
| GET/POST | `/api/avans` | Avans kayıtları |
| GET | `/api/hesapla/<id>/<yil>/<ay>` | Aylık maaş hesapla |
| GET | `/api/rapor/excel/<yil>/<ay>` | Excel raporu indir |
| POST | `/api/yedek` | Yedek oluştur |
| GET | `/api/yedek` | Yedek listesi |
| POST | `/api/yedek/geri/<dosya>` | Yedeği geri yükle |
| POST | `/api/geri_al/<log_id>` | İşlemi geri al |

---

## 📝 Notlar

- Uygulama **yalnızca Windows** üzerinde tek örnek kontrolü yapar (Windows Mutex API).
- Veritabanı dosyası (`puantaj.db`) proje kök dizininde saklanır.
- Yedekler `yedekler/` klasörüne `puantaj_YYYYMMDD_HHMMSS.db` formatında kaydedilir.

---

## 📄 Lisans

MIT License — dilediğiniz gibi kullanabilir, değiştirebilir ve dağıtabilirsiniz.
