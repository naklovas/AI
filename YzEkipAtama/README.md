# ATAMAN · YZ Ekip Atama

<img src="wwwroot/ataman-logo.svg" alt="ATAMAN" width="360">

Yapay zeka geliştirme sürecinin aksiyon/kontrol maddelerini, servis kataloğuna bakarak kurum içi YZ modeline sorar ve her madde için **işi yapması gereken ekip(ler)i ve görevlerini**, ayrıca **kararın gerekçesini** önerir.

Proje `Program.cs` (.NET 8 minimal API), `AuthHelper.cs` ve `wwwroot/index.html` (dış kütüphane kullanmayan, filtreli bir grid; açık/koyu tema) dosyalarından oluşur. Servis kataloğu ve prosedürler SQL Server'dan okunur; analiz işleri bellekte tutulur ve 12 saat sonra silinir.

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
| `ServisAciklamasiGonder` / `ServisAciklamaMaksKarakter` | Katalogdaki servis açıklamaları (`SERVIS_ACIKLAMASI`) servis adının yanında gönderilir; her biri en fazla bu kadar karakter (`0` = tamamı). Prompt'u en çok büyüten ayar budur; sayfadaki "Prompt boyutu" kutusundan izleyin. |
| `DenemeSayisi` | Hata alındığında veya yanıt ayrıştırılamadığında toplam deneme sayısı |
| `ExcelAyarlari:KontrolKolonu / TahminiEkipKolonu` | Kontrol listesi kolonları (varsayılan A = Kontrol, B = Tahmini Ekip). Tahmini ekip boş olabilir; modele **yalnızca ipucu** olarak gider. |

## Kullanım

- **Servis kataloğu veritabanından:** `KatalogVeritabani:BaglantiMetni` doluysa katalog açılışta `[SM9PROD].[dbo].[ZTSERVISKATALOGUM1]` tablosundan okunur. Kullanılan kolonlar: `SERVIS_KODU`, `SERVIS_ADI`, `SERVIS_SAHIBI_BOLUM`, `YONETICILER`, `SERVIS_KATEGORISI`, `SERVIS_ACIKLAMASI`. Sayfadaki **Veritabanından yenile** düğmesi katalogu yeniden okur. "Kayıt durumları" listesi `DURUM` / `ONAY_DURUMU` / `RECORD_ACTIVE` dağılımını gösterir; buna bakarak `Filtre` ayarına bir WHERE koşulu yazılabilir. Tamamen farklı bir sorgu gerekiyorsa `Sorgu` ayarı kullanılır.
- **Prosedür listesi veritabanından:** `ProsedurVeritabani:BaglantiMetni` doluysa prosedürler açılışta `[doku].[dbo].[prosedur]` tablosundan okunur (`Durum` = 0/False olanlar atlanır). Sayfada **Veritabanından yenile** düğmesi vardır.
- **Yetki kontrolü:** `AppSettings:YetkiKontrolu = true` ise kullanıcı Windows kimliğiyle doğrulanır, `AuthHelper` ile DokuPanel'den `CurrentSiteName` için yetki listesi sorulur. Sonuç kullanıcı başına `YetkiOnbellekDakika` boyunca önbellekte tutulur. Yetkisiz kullanıcıya "Erişim Reddedildi" sayfası, API isteklerine 403 döner.
- **Prosedürler:** Her madde için anahtar kelimeyle en ilgili `IlgiliProsedurSayisi` (varsayılan 15) prosedür, süreç sahibi birimiyle birlikte modele gönderilir. Süreç sahibi birim, katalogdaki ekiple eşleştirilir. Model ilgili prosedürleri `PROSEDUR: <doküman kodu> | <ilişki>` satırlarıyla döner; listede olmayan kodlar atılır. Sonuçlar gridde (linkli) ve Excel'de üçüncü kolon (*YZ İlgili Mevcut Prosedürler*) olarak gösterilir.
- **Metinle sor** kartına yapılacak işi yazıp **Ekipleri Bul**'a basın. Sorumlu ve destek ekipler, her birinin görevi ve karar gerekçesi gösterilir. "Her satırı ayrı bir iş olarak değerlendir" seçiliyse her satır ayrı bir madde olur; baştaki `1.`, `-` gibi işaretler atılır. Sonuçlar gridde listelenir ve Excel olarak indirilebilir.
- **Toplu analiz · Kontrol listesi:** Excel iki kolonludur: **A = Kontrol**, **B = Tahmini Ekip** (boş olabilir). İlk satır başlıksa atlanır; başlık yoksa çıktıya eklenir. Sonuçlar C, D (ve prosedür varsa E) kolonlarına yazılır. Kontrol listesi kartına yüklenip **Analizi Başlat** ile bütün maddeler işlenir (dosya seçilmemişse düğme dosya seçtirir).

## İsim benzerliği ve katalog dışı ekipler

- Sürecin bütün maddeleri zaten YZ ile ilgili olduğu için "yapay / zeka / YZ" kelimeleri anahtar kelime puanlamasında yok sayılır.
- Prompt, ekibin işin türüne göre seçilmesini ister. `GenelAdliEkipler` listesindeki ekipler (varsayılan: YAPAY ZEKA UYGULAMALARI) yalnızca iş doğrudan model geliştirme veya işletme ise sorumlu yapılır.
- `ExcelAyarlari:KatalogaEklenecekEkipler` listesindeki ekipler ile ekiplere ek görev tanımı verilir: ad katalogdaki ekiple eşleşirse tanım o ekibe eklenir, eşleşmezse katalog dışı yeni ekip olarak eklenir. Varsayılan: STRATEJİK PLANLAMA ve KOORDİNASYON için dar bir görev tanımı (talep yönetimi/Pergel, kapsam, paydaş, iletişim).
- `EkKurallar` listesindeki kurum kuralları prompt'a öncelikli kural olarak eklenir. `{EKİP ADI}` yazılan yerler ekibin `[E...]` koduyla değiştirilir. Varsayılan kural yalnızca maddenin asıl konusu talep (Pergel), kapsam, paydaş veya proje iletişimi ise STRATEJİK PLANLAMA ve KOORDİNASYON'u sorumlu yapar; risk, uyum, güvenlik, veri gibi konularda uygulanmaz.
- Denge kuralı: hiçbir ekip varsayılan ekip değildir; ekip yalnızca servis/prosedür içeriği maddeyle doğrudan uyuşuyorsa seçilir.
- Katalogda uygun ekip yoksa model alakasız ekip seçmek yerine `EKIP: YOK` yazar; sonuç "KATALOGDA UYGUN EKİP YOK" ve olması gereken işlevle gösterilir.
- Mevcut süreç önceliği: maddedeki faaliyet kurumda zaten bir serviste veya prosedürde yürütülüyorsa sorumlu o süreci yürüten ekiptir; YZ ekibi en fazla destek olur.

## Hücre içinde birden fazla madde

`HucreMaddeleriniAyir` açıkken (varsayılan) şu durumlarda hücre maddelere bölünür ve her madde ayrı değerlendirilir:
- Satırlar `-`, `•`, `1.`, `1)`, `a)` gibi işaretlerle başlıyorsa,
- Hücrede alt alta yazılmış ayrı cümleler varsa,
- Aynı satırda `1) … 2) …` biçiminde numaralandırma varsa.

"Aşağıdakiler yapılmalıdır:" gibi bir giriş satırı her maddeye bağlam olarak eklenir. Sonuçlar Excel'de aynı hücreye `[1/3] …`, `[2/3] …` başlıklarıyla yazılır; gridde her madde ayrı satırdır.

## Uydurma ekip adına karşı önlemler

- Her ekibe prompt'ta `[E012]` gibi bir kod verilir. Modelden ekibi koduyla birlikte yazması istenir. Ad hatalı yazılsa da kod doğruysa eşleşme yapılır.
- Model katalogda olmayan bir birim adı yazarsa yanıt kabul edilmez. Geçersiz adlar modele bildirilip düzeltmesi istenir (`DenemeSayisi`, varsayılan 3).
- Son denemede de uydurma ad varsa bu adlar sonuçtan çıkarılır ve gerekçeye not düşülür. Hiç geçerli ekip kalmazsa madde "Hata" olarak işaretlenir.

## Akış

1. Servis kataloğu ve prosedürler veritabanından okunur. Kontrol listesi iki kolonlu Excel olarak yüklenir (A = Kontrol, B = Tahmini Ekip).
2. Katalog "Servis Sahibi Bölüm" bazında gruplanarak ekip profilleri çıkarılır.
3. Her madde için anahtar kelime benzerliğiyle aday ekipler seçilir. Tahmini ekip kolonundaki ekipler her zaman adaylara eklenir. Ardından model çağrılır. Model `EKIP: ad | Sorumlu | görev`, `GUVEN: %80`, `GEREKCE: ...` satırlarıyla yanıt verir; JSON dönerse o da okunur. `<think>` blokları atlanır.
4. Modelin yazdığı ekip adı katalogla eşleştirilir (büyük-küçük harf, Türkçe karakter ve kısmi ad farkları tolere edilir). Katalogda bulunamayan ekip kırmızı ile işaretlenir.
5. **Excel'i İndir** düğmesi, yüklediğiniz orijinal dosyanın sonuna iki kolon ekleyip indirir: *YZ Önerilen Ekipler ve Görevleri* ve *YZ Karar Gerekçesi*. Diğer hücrelere dokunulmaz. Çıktı dosyasını tekrar yüklerseniz bu iki kolon yeni kolon açılmadan güncellenir.

Gridde arama, ekip, rol, durum ve güven aralığına göre filtreleme yapılabilir. Ayrıca tahmini ekiple karşılaştırma (Aynı / Kısmi / Farklı) gösterilir. Her madde tek tek yeniden çalıştırılabilir. Detay penceresinde modele gönderilen aday ekipler ve modelin ham yanıtı görülebilir.
