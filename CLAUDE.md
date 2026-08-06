# Fenotipik İndeks Hesaplama

Karacabey Merinos koyunları için fenotipik indeks hesaplayan Streamlit uygulaması.
Ham veri → veri kalite kontrolü → fenotipik indeks → Excel/CSV indir.

GitHub: `cemaldayanikli/koyun-fenotipik-indeks` (**public**)

## ⚠️ İŞE BAŞLAMADAN ÖNCE

`AI_GELISTIRICI_HAFIZASI.md` dosyasını oku. Geliştirme kararları ve "neden böyle
yapıldı" bilgisi orada, en yeni bölüm en üstte. **İş bitince en üste tarih başlıklı
yeni bir bölüm ekle.**

`README.md` hesaplama mantığının kullanıcıya dönük özetini tutuyor — ağırlık veya
eşik değiştirirsen **oradaki tabloyu da güncelle**, yoksa belge koddan sapar.

## Zincirdeki yeri

Bu uygulama üç parçalı bir akışın **ilk** halkasıdır:

```
Fenotipik_Indeks_App  (bu proje)      → Fenotipik_Indeks_YYYY.xlsx   (yılda bir kez)
        ↓
koyun-fenotipik-puan-birleştirme      → Puan_Birlesim_*.xlsx/.tsv    (çoklu yıl ortalaması)
        ↓
damızlık_reforme_seçim_app            → sahada seçim ekranı
```

Çıktı sütun adlarını değiştirirsen **bir sonraki halka kırılır**. Değiştirmen
gerekiyorsa kardeş klasördeki `koyun-fenotipik-puan-birleştirme/app.py`'nin okuma
tarafını da güncelle ve iki projeyi birlikte yayınla.

## Proje yapısı

| Dosya | Ne |
|---|---|
| `app.py` | Uygulamanın tamamı, tek dosya (~47 KB) |
| `requirements.txt` | Paket bağımlılıkları |
| `README.md` | Kullanıcıya dönük anlatım + hesaplama tablosu |

`app.py` tek dosya olduğu için baştan sona okumaya çalışma; `grep -n "^def \|^# ==="`
ile harita çıkarıp ilgili bölümü düzenle.

## Hesaplama mantığı — kırmızı çizgiler

Altı kriter ağırlığı ve eşikleri **sidebar'dan değiştirilebilir**; koddaki değerler
yalnızca varsayılandır (kaynak: `Is_Akisi_Sablonu_Final.xlsm`).

| Kriter | Varsayılan ağırlık |
|---|---|
| DKV (Doğum Tipi) | 17 |
| 90KV (Kuzu Sayısı) | 25 |
| Toplam 90CA | 20 |
| OrtCAA | 12 |
| Erken Kızgınlık | 14 |
| Embriyonik Kayıp | 12 |

- **Varsayılanları sessizce değiştirme.** Geçmiş yılların çıktıları bu değerlerle
  üretildi; değişirse yıllar arası karşılaştırma anlamsızlaşır. Değiştireceksen
  hafızaya *hangi yıldan itibaren* geçerli olduğunu yaz.
- **Erken Kızgınlık grubu** Koç Aşım Tarihi'nden **veya** Kuzu Doğum Tarihi'nden
  hesaplanabilir; hangisinin kullanıldığı sidebar'dan seçilir ve **sonucu değiştirir**.
  Her iki grup numarası da tabloda görünür — çıktıyı yorumlarken hangisine
  bakıldığını doğrula.
- Doğum tipi (tek/ikiz/üçüz/dört) normalize edilerek okunur ve sütun eşleşmesi
  boşluklara toleranslıdır (son commit'in konusu buydu) — girdi dosyasındaki
  başlıklar birebir aynı olmak zorunda değil.

## Dağıtım

```powershell
.\deploy.ps1 "Ne değişti"
```

Önce `py_compile` ile derler; hata varsa **hiçbir şey push edilmez**. Sonra
`git commit` + `git push`. Ayrıntı: `.claude/commands/deploy.md`.

### ⚠️ İki nokta

1. **`main`e push = canlıya çıkış.** Streamlit Community Cloud `main` dalını izler ve
   push'tan ~1 dk sonra uygulamayı yeniden başlatır. Yarım işi commit etme.
2. **Bu repo public.** `.gitignore` veri dosyalarını (`*.xlsx`, `*.xlsm`, `*.xls`,
   `*.csv`) kasten dışarıda tutuyor — hayvan verisi herkese açık olmasın diye.
   **Bu kuralları gevşetme.** Örnek veri gerekiyorsa uydurma satırlarla, ayrı ve
   açıkça "örnek" adlı bir dosyayla ekle.

## Yerel çalıştırma

```bash
pip install -r requirements.txt
streamlit run app.py
```

Tarayıcıda `http://localhost:8501` açılır.
