# Word Merge Uygulaması

Basit bir masaüstü arayüzü ile fotoğrafları kategori bazlı yükleyip Word şablonlarına yerleştirir ve tek bir `.docx` dosyası olarak birleştirir.

## Özellikler
- 4 kategori desteği:
  - Pre photos
  - Post photos
  - Teardown photos
  - Handle side cover
- Her template sayfası için 6 fotoğraf slotu (`IMG_1` ... `IMG_6`)
- 6'dan fazla fotoğrafta template otomatik çoğaltılır (2,3,4... sayfa)
- Kategori sırası korunur: **pre -> post -> teardown -> handle_side_cover**
- Ortak progress bar:
  - Fotoğraf yükleme sırasında her fotoğrafta ilerler
  - Merge sırasında her sayfada ilerler
- `tempfiles` klasörüne çalışma dosyaları kaydedilir
- Uygulama açılışında var olan `tempfiles` temizleme sorusu
- Uygulama kapanışında `tempfiles` temizleme sorusu
- Merge çıktısı için:
  - Dosya yolu ve adı kullanıcı tarafından seçilir
  - Varsayılan ad seçilen kategorilere göre üretilir (örn. `pre_post_teardown.docx`)

## Kurulum
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Çalıştırma
```bash
python app.py
```

## Template hazırlama
`templates` klasörüne şu dosyaları ekleyin:
- `pre_photos.docx`
- `post_photos.docx`
- `teardown_photos.docx`
- `handle_side_cover.docx`

Her template içinde fotoğraf konulacak alanlarda şu yer tutucular bulunmalı:
- `{{IMG_1}}`
- `{{IMG_2}}`
- `{{IMG_3}}`
- `{{IMG_4}}`
- `{{IMG_5}}`
- `{{IMG_6}}`

> Öneri: Bu yer tutucuları tablo hücrelerine koyun. Böylece tablo düzeni sabit kalır.

## Notlar
- Uygulama görsel yerleşimi korumak için resimleri sabit kutu ölçüsüyle yerleştirir.
- Çıktı Word dosyaları sırayı bozmaz ve kategorileri seçtiğiniz düzende birleştirir.
