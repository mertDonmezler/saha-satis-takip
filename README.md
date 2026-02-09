# Saha Satis Takip Sistemi

Excel tabanli saha satis ekibi takip araci. Haftalik ziyaret planlari, gerceklesen ziyaretler ve siparis formlarini otomatik tarayip tek bir MASTER_DATA.xlsx dosyasinda birlestirir.

## Ozellikler

- **Tam Otomatik Algilama** - Hafta, temsilci ve dosya tipi dosya adlarindan otomatik tespit edilir. Hardcoded deger yok.
- **8 Sayfalik Master Rapor** - Dosya envanteri, haftalik takvim, planlanan/yapilan ziyaretler, siparisler, musteri master, ozet ve veri kalite sorunlari
- **Watch Mode** - Klasore yeni Excel eklendi mi? Otomatik gunceller (debounce destekli)
- **Akilli Veri Cekme** - Merged cell/carry-forward, bos satir filtreleme, dedup, gomulu siparis tespiti
- **Dosya Kilidi Korumasi** - Excel acikken yazma hatasi yerine temp dosyaya kaydeder
- **Log Dosyasi** - Tum islemler `master_data.log` dosyasina kaydedilir

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanim

### Tek seferlik guncelleme
```bash
python master_data_olustur.py
```
veya `GUNCELLE.bat` dosyasini cift tiklayin.

### Surekli izleme (watch mode)
```bash
python master_data_olustur.py --izle
```
veya `IZLE.bat` dosyasini cift tiklayin.

## Dosya Adlandirma Kurali

Script su formatlari otomatik tanir:

| Format | Ornek |
|--------|-------|
| `DD.MM.YYYY-DD.MM.YYYY ISIM TIP.xlsx` | `26.01.2026-30.01.2026 ALI YILMAZ PLANLANAN ZIYARET.xlsx` |
| `DD-DD AY TIP ISIM.xlsx` | `26-30 OCAK HAFTALIK SIPARIS FORMU MEHMET DEMIR.xlsx` |

Desteklenen tipler: `PLANLANAN ZIYARET`, `YAPILAN ZIYARET`, `SIPARIS FORMU`, `ZIYARET PLANI`

## Cikti: MASTER_DATA.xlsx

| Sheet | Icerik |
|-------|--------|
| Dosya Envanteri | Hangi dosyalar mevcut/eksik |
| Haftalik Takvim | Gun bazli ziyaret takvimi |
| Planlanan Ziyaretler | Tum planlanan ziyaretler |
| Yapilan Ziyaretler | Gerceklesen ziyaretler (dedup edilmis) |
| Siparisler | Urun/adet/fiyat detaylari |
| Musteri Master | Benzersiz musteri listesi |
| Ozet | Temsilci bazli performans tablosu |
| Veri Kalite Sorunlari | Eksik dosyalar ve tarih hatalari |

## Teknolojiler

- Python 3
- openpyxl (Excel okuma/yazma)
- Logging modulu (dosya + konsol log)
- Windows batch dosyalari

## Lisans

MIT
