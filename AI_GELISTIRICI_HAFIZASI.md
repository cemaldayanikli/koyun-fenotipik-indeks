# AI Geliştirici Hafızası — Fenotipik İndeks Hesaplama

Bu dosya, projede yapılan geliştirmelerin kalıcı kaydıdır. Amaç: projeye sonradan bakan
kişinin (veya yapay zekâ asistanının) **neyin neden yapıldığını** koddan okumadan
anlayabilmesi.

## Kurallar

- Her geliştirme, **en yeni en üstte** olacak şekilde bir başlık altında yazılır.
- Her kayıt şunları içerir: **tarih**, **talep/belirti**, **teşhis veya karar**,
  **dokunulan dosyalar**, **doğrulama**, **yayın durumu**.
- Sadece "ne yapıldı" değil, **neden öyle yapıldı** da yazılır.
- Aşağıdaki *Kalıcı referans* bölümü kronolojik değildir, **silinmez**.

---

## Kalıcı referans (kronolojik değil — silme)

| | |
|---|---|
| Teknoloji | Streamlit, tek dosya (`app.py`, ~47 KB) |
| GitHub | `cemaldayanikli/koyun-fenotipik-indeks` — **PUBLIC** |
| Yayın | Streamlit Community Cloud, `main` dalını izler → **push = canlıya çıkış** |
| Zincirdeki yeri | İlk halka: bu app → `koyun-fenotipik-puan-birleştirme` → `damızlık_reforme_seçim_app` |
| Çıktı | `Fenotipik_Indeks_YYYY.xlsx` (yılda bir kez üretilir) |

### Ağırlıklar — değiştirmeden önce oku

Altı kriter (varsayılanlar `Is_Akisi_Sablonu_Final.xlsm`'den): DKV 17 · 90KV 25 ·
Toplam 90CA 20 · OrtCAA 12 · Erken Kızgınlık 14 · Embriyonik Kayıp 12.
Hepsi sidebar'dan değiştirilebilir; koddaki değerler yalnızca varsayılandır.

⚠️ **Varsayılanları sessizce değiştirme.** Geçmiş yılların çıktıları bu değerlerle
üretildi; değişirse yıllar arası karşılaştırma anlamsızlaşır. Değiştireceksen buraya
*hangi yıldan itibaren* geçerli olduğunu yaz.

⚠️ **Erken Kızgınlık grubu** Koç Aşım Tarihi'nden **veya** Kuzu Doğum Tarihi'nden
hesaplanabilir; seçim sidebar'dan yapılır ve **sonucu değiştirir**. Her iki grup
numarası da tabloda görünür.

### 🔒 Bu repo public

`.gitignore` veri dosyalarını (`*.xlsx`, `*.xlsm`, `*.xls`, `*.csv`) kasten dışarıda
tutuyor — hayvan verisi herkese açık olmasın diye. **Bu kuralları gevşetme.**

---

## 6 Ağustos 2026 — GELİŞTİRME ALTYAPISI KURULDU (kod değişmedi)

**Talep (kullanıcı):** "İlgili klasördeki tüm projelere her proje klasörüne AI
geliştirici hafızası ekle, her projeye GitHub reposu ve commit ekle, ayrıca her projede
*deploy yap* dediğimde otomatik olarak clasp ile push edip deploy yapacak hem de
GitHub'a commit edecek şekilde ayarlama yap."

Bu bir Streamlit projesi olduğu için clasp söz konusu değil; burada "deploy" =
**doğrulama + commit + push**. `app.py` değişmedi.

### Eklenenler

| Dosya | Ne işe yarar |
|---|---|
| `CLAUDE.md` | Proje kuralları: zincirdeki yeri, ağırlıklar, public repo uyarısı |
| `AI_GELISTIRICI_HAFIZASI.md` | Bu dosya |
| `deploy.ps1` | `py_compile` ile tüm `.py`'yi derler → `git commit` + `git push` |
| `.claude/commands/deploy.md` | "deploy yap" denince izlenecek akış |

### Kararlar ve nedenleri

- **`.gitattributes` eklenmedi.** Diğer (Apps Script) projelere `* -text` konuldu ki
  clasp'taki uzak kopyayla bayt eşitliği bozulmasın. Ama bu depoda
  `core.autocrlf = true` ve geçmiş LF olarak saklanmış — `* -text` eklenseydi bir
  sonraki commit **tüm dosyaları değişmiş** gösterirdi.
- **`.gitignore` ezilmedi, üzerine eklendi.** İlk denemede ortak şablonla
  değiştirilmişti; bu, `*.xlsx` / `*.csv` gibi **veri dosyası** kurallarını siliyordu
  ve repo public olduğu için hayvan verisi sızabilirdi. Geri alınıp yalnızca eksik
  ortak kurallar (`.env.*`, `desktop.ini`, `~$*` …) sonuna eklendi.
- **`deploy.ps1` önce derler, sonra push eder.** `main`e push = canlıya çıkış olduğu
  için sözdizimi hatası doğrudan sahaya iner; kapı burada tutuluyor.

### Doğrulama (yapıldı)

`deploy.ps1` sözdizimi PowerShell ayrıştırıcısıyla temiz. `git status` yalnızca
beklenen dosyaları gösterdi (`.gitignore` değişikliği + `.claude/` + `deploy.ps1`);
veri dosyalarının hiçbiri izlenmeye alınmadı.

**Yayın yapılmadı** — bu oturumda uygulama kodu değişmedi.
