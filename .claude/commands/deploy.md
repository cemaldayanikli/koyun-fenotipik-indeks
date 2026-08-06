---
description: Projeyi dogrula ve GitHub'a yayinla
---

Kullanici "deploy" / "deploy yap" / "yayinla" dediginde bu akisi calistir.

## Adimlar

1. **Once degisiklikleri ozetle.** `git status` ve `git diff --stat` ile ne degistigini
   gor; tek cumlelik bir commit mesaji cikar (Turkce, ne degistigini soyleyen).
2. **Betigi calistir:** `.\deploy.ps1 "<cikardigin mesaj>"`
3. **Ciktiyi oku** ve sonucu bildir: sozdizimi denetimi gecti mi, commit ve push gitti mi.
4. **Hafizaya yaz.** `AI_GELISTIRICI_HAFIZASI.md` dosyasinin EN USTUNE tarih baslikli
   yeni bir bolum ekle: talep, karar, dokunulan dosyalar, dogrulama, commit ozeti.

## Bu projeye ozel (Streamlit)

Canliya cikis GitHub uzerinden olur - Streamlit Community Cloud `main` dalini izler ve
push'tan ~1 dk sonra kendini yeniden baslatir. Ayri bir "deploy" adimi YOKTUR;
push = yayin. Bu yuzden main'e giden her commit canliya ciktigini bilerek atilmalidir.

`deploy.ps1` once `py_compile` ile tum `.py` dosyalarini derler; hata varsa
**hicbir sey push edilmez**.

## Yerine gecemeyecegi sey

Kullanici "yayinlama" dediyse **1. adimda dur**: kodu yaz, dogrula - ama `deploy.ps1`i
CALISTIRMA. `main`e push = canliya cikis oldugu icin commit bile kullanicinin
onayiyla atilmali.