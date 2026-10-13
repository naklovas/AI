# YZ Ekip Atama

Yapay zeka geliştirme sürecinin aksiyon/kontrol maddelerini, servis kataloğuna bakarak kurum içi YZ modeline sorar ve her madde için **işi yapması gereken ekip(ler)i ve görevlerini**, ayrıca **kararın gerekçesini** önerir.

Proje iki dosyadan oluşur: `Program.cs` (.NET 8 minimal API) ve `wwwroot/index.html` (dış kütüphane kullanmayan, filtreli bir grid). Veritabanı yoktur. İşler bellekte tutulur ve 12 saat sonra silinir.

## Çalıştırma

İlk kurulumda örnek ayar dosyasını kopyalayıp kendi değerlerinizi yazın:

```cmd
copy appsettings.ornek.json appsettings.json
notepad appsettings.json
dotnet run
```

Uygulama `http://localhost:5080` adresinde açılır.

`appsettings.json` git'te tutulmaz (`.gitignore`), `git pull` bu dosyaya hiç dokunmaz. Yeni bir ayar eklendiğinde `appsettings.ornek.json` dosyası güncellenir; yeni ayarı oradan kendi dosyanıza elle eklersiniz. Eklemezseniz koddaki varsayılan değer kullanılır.

`QwenApiUrl` (OpenAI uyumlu `chat/completions` adresinin tamamı) ve `QwenApiKey` değerleri `YzAyarlari` bölümüne ya da kök seviyeye yazılabilir. Model Qwen olmasa da ayar adları değişmez.

| Ayar | Açıklama |
|---|---|
| `AiModel` | Model adı (varsayılan `zt-ga-small-0`) |
| `Paralellik` | Aynı anda modele gönderilen madde sayısı |
| `OnElemeYap` | `false` (varsayılan): her maddede 114 ekibin tamamı tüm servisleriyle gönderilir. `true`: anahtar kelime benzerliğiyle `AdayEkipSayisi` kadar ekip, en ilgili `EkipBasinaServisSayisi` servisiyle gönderilir (ekip adlarının tamamı yine de gider). |
| `MaksToken` | Yanıt için token sınırı; düşünen modellerde `<think>` bölümü de buna dahildir. `0` = sınır yok. |
| `EkipBasinaServisSayisi` | Her aday ekip için gönderilen servis sayısı. Maddeyle en ilgili servisler önce gelir. |
| `MaksEkipSayisi` | Bir madde için önerilebilecek en fazla ekip sayısı (ilk ekip "Sorumlu", diğerleri "Destek") |
| `DusunmeyiKapat` | Qwen3 gibi düşünen modellerde `<think>` bölümünü kapatır (`chat_template_kwargs.enable_thinking=false`). Sunucu tanımazsa `false` yapın. |
| `ServisAciklamasiGonder` / `ServisAciklamaMaksKarakter` | Katalogdaki servis açıklamaları (başlığında "Açıklama" geçen kolon, yoksa `ExcelAyarlari:KatalogAciklamaKolonu`, varsayılan F) servis adının yanında gönderilir; her biri en fazla bu kadar karakter (`0` = tamamı). Prompt'u en çok büyüten ayar budur; sayfadaki "Prompt boyutu" kutusundan izleyin. |
| `DenemeSayisi` | Hata alındığında veya yanıt ayrıştırılamadığında toplam deneme sayısı |
| `ExcelAyarlari:KullaniciEkipKolonu / KullaniciNotKolonu` | Sizin doldurduğunuz E ve F kolonları. Bu kolonlar modele **yalnızca ipucu** olarak gider. |

## Kullanım

- **Servis kataloğu** bir kez yüklenir. Uygulama `veri/son_katalog.xlsx` dosyasında saklar; uygulama yeniden başlasa da tekrar yüklemek gerekmez. Yeni bir dosya yüklenene kadar bu katalog kullanılır.
- **Metinle sor** sekmesine yapılacak işi yazıp **Ekipleri Bul**'a basın. Sorumlu ve destek ekipler, her birinin görevi ve karar gerekçesi gösterilir. "Her satırı ayrı bir iş olarak değerlendir" seçiliyse her satır ayrı bir madde olur; baştaki `1.`, `-` gibi işaretler atılır. Sonuçlar gridde listelenir ve Excel olarak indirilebilir.
- **Excel ile toplu** sekmesinde kontrol listesi yüklenir ve bütün maddeler işlenir.

## Uydurma ekip adına karşı önlemler

- Her ekibe prompt'ta `[E012]` gibi bir kod verilir. Modelden ekibi koduyla birlikte yazması istenir. Ad hatalı yazılsa da kod doğruysa eşleşme yapılır.
- Model katalogda olmayan bir birim adı yazarsa yanıt kabul edilmez. Geçersiz adlar modele bildirilip düzeltmesi istenir (`DenemeSayisi`, varsayılan 3).
- Son denemede de uydurma ad varsa bu adlar sonuçtan çıkarılır ve gerekçeye not düşülür. Hiç geçerli ekip kalmazsa madde "Hata" olarak işaretlenir.

## Akış

1. Kontrol listesi ve servis kataloğu Excel dosyaları yüklenir. Başlık satırı otomatik bulunur (Kontrol listesinde "Aksiyon/Kontrol", katalogda "Servis Adı", "Servis Sahibi Bölüm" ve varsa "Açıklama" başlıkları aranır). Birleştirilmiş hücreler desteklenir.
2. Katalog "Servis Sahibi Bölüm" bazında gruplanarak ekip profilleri çıkarılır.
3. Her madde için anahtar kelime benzerliğiyle aday ekipler seçilir. E kolonunda yazdığınız ekipler her zaman adaylara eklenir. Ardından model çağrılır. Model `EKIP: ad | Sorumlu | görev`, `GUVEN: %80`, `GEREKCE: ...` satırlarıyla yanıt verir; JSON dönerse o da okunur. `<think>` blokları atlanır.
4. Modelin yazdığı ekip adı katalogla eşleştirilir (büyük-küçük harf, Türkçe karakter ve kısmi ad farkları tolere edilir). Katalogda bulunamayan ekip kırmızı ile işaretlenir.
5. **Excel'i İndir** düğmesi, yüklediğiniz orijinal dosyanın sonuna iki kolon ekleyip indirir: *YZ Önerilen Ekipler ve Görevleri* ve *YZ Karar Gerekçesi*. Diğer hücrelere dokunulmaz. Çıktı dosyasını tekrar yüklerseniz bu iki kolon yeni kolon açılmadan güncellenir.

Gridde arama, alt kategori, ekip, rol, durum ve güven aralığına göre filtreleme yapılabilir. Ayrıca E kolonundaki görüşünüzle karşılaştırma (Aynı / Kısmi / Farklı) gösterilir. Her madde tek tek yeniden çalıştırılabilir. Detay penceresinde modele gönderilen aday ekipler ve modelin ham yanıtı görülebilir.
