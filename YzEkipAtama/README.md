# YZ Ekip Atama

Yapay zeka geliştirme sürecinin aksiyon/kontrol maddelerini, servis kataloğuna bakarak kurum içi YZ modeline sorar ve her madde için **işi yapması gereken ekip(ler)i ve görevlerini**, ayrıca **kararın gerekçesini** önerir.

Proje iki dosyadan oluşur: `Program.cs` (.NET 8 minimal API) ve `wwwroot/index.html` (dış kütüphane kullanmayan, filtreli bir grid). Veritabanı yoktur. İşler bellekte tutulur ve 12 saat sonra silinir.

## Çalıştırma

```bash
dotnet run
# http://localhost:5080
```

`appsettings.json` → `YzAyarlari` bölümünde `QwenApiUrl` (OpenAI uyumlu `chat/completions` adresinin tamamı) ve `QwenApiKey` alanlarını doldurun. Bu iki ayar kök seviyede de yazılabilir; model Qwen olmasa da ayar adları değişmez.

| Ayar | Açıklama |
|---|---|
| `AiModel` | Model adı (varsayılan `zt-ga-small-0`) |
| `Paralellik` | Aynı anda modele gönderilen madde sayısı |
| `AdayEkipSayisi` | Prompt'a servis listesiyle birlikte konulan en olası ekip sayısı. Ekiplerin yalnızca **adları** her zaman eksiksiz gönderilir. `0` verilirse bütün ekipler tüm servisleriyle gönderilir (modelin bağlam penceresi büyükse kullanın). |
| `EkipBasinaServisSayisi` | Her aday ekip için gönderilen servis sayısı. Maddeyle en ilgili servisler önce gelir. |
| `MaksEkipSayisi` | Bir madde için önerilebilecek en fazla ekip sayısı (ilk ekip "Sorumlu", diğerleri "Destek") |
| `DusunmeyiKapat` | Qwen3 gibi düşünen modellerde `<think>` bölümünü kapatır (`chat_template_kwargs.enable_thinking=false`). Sunucu tanımazsa `false` yapın. |
| `DenemeSayisi` | Hata alındığında veya yanıt ayrıştırılamadığında toplam deneme sayısı |
| `ExcelAyarlari:KullaniciEkipKolonu / KullaniciNotKolonu` | Sizin doldurduğunuz E ve F kolonları. Bu kolonlar modele **yalnızca ipucu** olarak gider. |

## Akış

1. Kontrol listesi ve servis kataloğu Excel dosyaları yüklenir. Başlık satırı otomatik bulunur (Kontrol listesinde "Aksiyon/Kontrol", katalogda "Servis Adı" ve "Servis Sahibi Bölüm" başlıkları aranır). Birleştirilmiş hücreler desteklenir.
2. Katalog "Servis Sahibi Bölüm" bazında gruplanarak ekip profilleri çıkarılır.
3. Her madde için anahtar kelime benzerliğiyle aday ekipler seçilir. E kolonunda yazdığınız ekipler her zaman adaylara eklenir. Ardından model çağrılır. Model `EKIP: ad | Sorumlu | görev`, `GUVEN: %80`, `GEREKCE: ...` satırlarıyla yanıt verir; JSON dönerse o da okunur. `<think>` blokları atlanır.
4. Modelin yazdığı ekip adı katalogla eşleştirilir (büyük-küçük harf, Türkçe karakter ve kısmi ad farkları tolere edilir). Katalogda bulunamayan ekip kırmızı ile işaretlenir.
5. **Excel'i İndir** düğmesi, yüklediğiniz orijinal dosyanın sonuna iki kolon ekleyip indirir: *YZ Önerilen Ekipler ve Görevleri* ve *YZ Karar Gerekçesi*. Diğer hücrelere dokunulmaz. Çıktı dosyasını tekrar yüklerseniz bu iki kolon yeni kolon açılmadan güncellenir.

Gridde arama, alt kategori, ekip, rol, durum ve güven aralığına göre filtreleme yapılabilir. Ayrıca E kolonundaki görüşünüzle karşılaştırma (Aynı / Kısmi / Farklı) gösterilir. Her madde tek tek yeniden çalıştırılabilir. Detay penceresinde modele gönderilen aday ekipler ve modelin ham yanıtı görülebilir.
