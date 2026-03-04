# Word-merge

Basit bir masaüstü uygulaması ile Word template'lerine fotoğraf ekler ve tek bir `.docx` dosyasında birleştirir.

## Özellikler
- 4 kategori desteği:
  - Pre photos
  - Post photos
  - Teardown photos
  - Handle side cover
- Her kategoride fotoğraflar **sırayı koruyarak** işlenir.
- Her template sayfası için 6 görsel slotu vardır (`{{IMG1}} ... {{IMG6}}`).
- 6'dan fazla fotoğraf gelirse aynı template, otomatik sayfalara bölünerek çoğaltılır.
- Tüm oluşturulan ara dosyalar `tempfiles/` altında tutulur.
- Uygulama başlarken mevcut `tempfiles` için temizleme sorusu sorar.
- Uygulamadan çıkarken `tempfiles` silinsin mi diye sorar.
- Ortak progress bar hem yükleme hem merge ilerlemesini gösterir.
- Çıktı dosya adını kullanıcı seçer; varsayılan ad seçilen opsiyonlardan üretilir (örn. `pre_post_teardown.docx`).

## Kurulum
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Template hazırlığı
`templates/` klasörüne aşağıdaki dosyaları ekleyin:
- `pre.docx`
- `post.docx`
- `teardown.docx`
- `handle_side_cover.docx`

Her template içinde fotoğraf yerleştirilecek alanlara aşağıdaki placeholder metinlerini koyun:
- `{{IMG1}}`
- `{{IMG2}}`
- `{{IMG3}}`
- `{{IMG4}}`
- `{{IMG5}}`
- `{{IMG6}}`

> Not: Uygulama görselleri bu placeholder'ları bulup o noktaya ekler. Tabloların kaymaması için template tablolarını sabit boyutlu hazırlamanız önerilir.

## Çalıştırma
```bash
python app.py
```
