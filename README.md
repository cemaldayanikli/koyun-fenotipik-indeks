# Fenotipik İndeks Hesaplama

Karacabey Merinos koyunları için fenotipik indeks hesaplayan Streamlit uygulaması.
Ham veri → Veri kalite kontrol → Fenotipik indeks → Excel/CSV indir.

## Yerel çalıştırma

```bash
pip install -r requirements.txt
streamlit run app.py
```

Tarayıcıda `http://localhost:8501` otomatik açılır.

## Streamlit Community Cloud'a Deploy

1. Bu klasörü kendi başına bir GitHub repo'su olarak yükle (aşağıdaki adımlar).
2. https://share.streamlit.io adresine git, GitHub ile giriş yap.
3. **New app** → Repository: seçtiğin repo → Branch: `main` → Main file: `app.py` → **Deploy**.

### GitHub'a yükleme

```bash
cd Fenotipik_Indeks_App
git init
git add .
git commit -m "Initial commit"
# GitHub'da yeni repo oluşturduktan sonra:
git remote add origin https://github.com/<KULLANICI>/<REPO>.git
git branch -M main
git push -u origin main
```

`gh` CLI yüklüyse tek komutla:

```bash
gh repo create koyun-fenotipik-indeks --public --source=. --push
```

## İçerik

- `app.py` — Streamlit uygulaması (tek dosya)
- `requirements.txt` — paket bağımlılıkları
- `.gitignore` — veri dosyalarını dışlar

## Hesaplama Mantığı

AYARLAR sayfasındaki 6 kriter ağırlığını kullanır (varsayılanlar `Is_Akisi_Sablonu_Final.xlsm`'den):

| Kriter | Ağırlık | Mantık |
|---|---|---|
| DKV (Doğum Tipi) | 17 | T → ×0.3, İ → ×0.7, Ü → ×1.0 |
| 90KV (Kuzu Sayısı) | 25 | 1 kuzu → ×0.7, 2 → ×0.9, 3 → ×1.0 |
| Toplam 90CA | 20 | 0–27.41 → 0.5 · 27.41–50 → 0.85 · >50 → 1.0 |
| OrtCAA | 12 | 0.112–0.311 → 0.5 · 0.311–0.35 → 0.85 · >0.35 → 1.0 |
| Erken Kızgınlık | 14 | Hafta ≤23 → ×1.0, =24 → ×0.85, ≤26 → ×0.3, sonra → 0 |
| Embriyonik Kayıp | 12 | Doğum var → 1.0 (sabit) |

**Erken Kızgınlık grubu** Koç Aşım Tarihi veya Kuzu Doğum Tarihi'nden hesaplanabilir (sidebar'dan seçilir). Her iki grup numarası da tabloda görünür.

Tüm ağırlıklar ve eşikler sidebar'dan değiştirilebilir.
