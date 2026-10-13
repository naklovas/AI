using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

var builder = WebApplication.CreateBuilder(args);

// appsettings.json git'te tutulmaz (appsettings.ornek.json'dan kopyalanır); dosya yoksa da varsayılanlarla açılır
if (string.IsNullOrWhiteSpace(builder.Configuration["Urls"])) builder.WebHost.UseUrls("http://0.0.0.0:5080");

var yzAyar = builder.Configuration.GetSection("YzAyarlari").Get<YzAyarlari>() ?? new YzAyarlari();
// QwenApiUrl / QwenApiKey YzAyarlari bölümünde yoksa kök seviyeden okunur
if (string.IsNullOrWhiteSpace(yzAyar.QwenApiUrl)) yzAyar.QwenApiUrl = builder.Configuration["QwenApiUrl"] ?? "";
if (string.IsNullOrWhiteSpace(yzAyar.QwenApiKey)) yzAyar.QwenApiKey = builder.Configuration["QwenApiKey"] ?? "";
var excelAyar = builder.Configuration.GetSection("ExcelAyarlari").Get<ExcelAyarlari>() ?? new ExcelAyarlari();

builder.Services.AddSingleton(yzAyar);
builder.Services.AddSingleton(excelAyar);
builder.Services.AddSingleton<IsYoneticisi>();
builder.Services.AddHttpClient("InsecureClient", c => c.Timeout = TimeSpan.FromMinutes(yzAyar.ZamanAsimiDakika))
    .ConfigurePrimaryHttpMessageHandler(() => yzAyar.SertifikaDogrulamasiniAtla
        ? new HttpClientHandler { ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator }
        : new HttpClientHandler());
builder.Services.ConfigureHttpJsonOptions(o =>
    o.SerializerOptions.Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping);

var app = builder.Build();
app.UseDefaultFiles();
app.UseStaticFiles();

app.MapGet("/api/ayarlar", (YzAyarlari a, ExcelAyarlari e) => new
{
    model = a.AiModel,
    paralellik = a.Paralellik,
    onElemeYap = a.OnElemeYap,
    adayEkipSayisi = a.AdayEkipSayisi,
    servisAciklamasi = a.ServisAciklamasiGonder ? (a.ServisAciklamaMaksKarakter > 0 ? $"ilk {a.ServisAciklamaMaksKarakter} karakter" : "tamamı") : "gönderilmiyor",
    maksEkipSayisi = a.MaksEkipSayisi,
    kullaniciEkipKolonu = e.KullaniciEkipKolonu,
    kullaniciNotKolonu = e.KullaniciNotKolonu,
    apiTanimli = !string.IsNullOrWhiteSpace(a.QwenApiUrl)
});

app.MapGet("/api/katalog", (IsYoneticisi yonetici) => yonetici.KatalogBilgisi());

app.MapPost("/api/katalog", async (HttpRequest req, IsYoneticisi yonetici) =>
{
    if (!req.HasFormContentType || req.Form.Files["katalog"] is not { } dosya)
        return Results.BadRequest(new { hata = "Servis kataloğu Excel dosyası gerekli." });
    try
    {
        yonetici.KatalogYukle(await OkuAsync(dosya), dosya.FileName);
        return Results.Ok(yonetici.KatalogBilgisi());
    }
    catch (ExcelOkumaHatasi ex)
    {
        return Results.BadRequest(new { hata = ex.Message });
    }
});

app.MapGet("/api/prosedur", (IsYoneticisi yonetici) => yonetici.ProsedurBilgisi());

app.MapPost("/api/prosedur", async (HttpRequest req, IsYoneticisi yonetici) =>
{
    if (!req.HasFormContentType || req.Form.Files["prosedur"] is not { } dosya)
        return Results.BadRequest(new { hata = "Prosedür listesi Excel dosyası gerekli." });
    try
    {
        yonetici.ProsedurYukle(await OkuAsync(dosya), dosya.FileName);
        return Results.Ok(yonetici.ProsedurBilgisi());
    }
    catch (ExcelOkumaHatasi ex)
    {
        return Results.BadRequest(new { hata = ex.Message });
    }
});

app.MapDelete("/api/prosedur", (IsYoneticisi yonetici) =>
{
    yonetici.ProsedurKaldir();
    return Results.Ok(yonetici.ProsedurBilgisi());
});

// Excel ile toplu analiz. Katalog gönderilmezse son yüklenen katalog kullanılır.
app.MapPost("/api/is", async (HttpRequest req, IsYoneticisi yonetici) =>
{
    if (!req.HasFormContentType)
        return Results.BadRequest(new { hata = "multipart/form-data bekleniyor." });
    var form = await req.ReadFormAsync();
    var kontrolDosya = form.Files["kontrol"];
    var katalogDosya = form.Files["katalog"];
    if (kontrolDosya is null)
        return Results.BadRequest(new { hata = "Kontrol listesi Excel dosyası gerekli." });

    try
    {
        if (katalogDosya is not null) yonetici.KatalogYukle(await OkuAsync(katalogDosya), katalogDosya.FileName);
        var analiz = yonetici.ExcelIleBaslat(await OkuAsync(kontrolDosya), kontrolDosya.FileName);
        return Results.Ok(analiz.Durum());
    }
    catch (ExcelOkumaHatasi ex)
    {
        return Results.BadRequest(new { hata = ex.Message });
    }
});

// Serbest metinle soru: işi anlat, ekipleri öğren. HerSatirAyriMadde=true ise her dolu satır ayrı madde olur.
app.MapPost("/api/sor", (MetinSorusu soru, IsYoneticisi yonetici) =>
{
    try
    {
        return Results.Ok(yonetici.MetinleBaslat(soru.Metin ?? "", soru.HerSatirAyriMadde).Durum());
    }
    catch (ExcelOkumaHatasi ex)
    {
        return Results.BadRequest(new { hata = ex.Message });
    }
});

app.MapGet("/api/is/{id}", (string id, IsYoneticisi yonetici) =>
    yonetici.Getir(id) is { } analiz ? Results.Ok(analiz.Durum()) : Results.NotFound(new { hata = "İş bulunamadı (uygulama yeniden başlatılmış olabilir)." }));

app.MapPost("/api/is/{id}/iptal", (string id, IsYoneticisi yonetici) =>
{
    if (yonetici.Getir(id) is not { } analiz) return Results.NotFound();
    analiz.Iptal.Cancel();
    return Results.Ok();
});

app.MapPost("/api/is/{id}/satir/{sira:int}/yeniden", (string id, int sira, IsYoneticisi yonetici) =>
    yonetici.YenidenCalistir(id, sira) ? Results.Ok() : Results.NotFound());

app.MapGet("/api/is/{id}/excel", (string id, IsYoneticisi yonetici, ExcelAyarlari e) =>
{
    if (yonetici.Getir(id) is not { } analiz) return Results.NotFound();
    var dosyaAdi = Path.GetFileNameWithoutExtension(analiz.DosyaAdi) + "_YZ_Ekip_Atama.xlsx";
    return Results.File(ExcelIslemleri.SonucYaz(analiz, e),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", dosyaAdi);
});

app.Run();

static async Task<byte[]> OkuAsync(IFormFile dosya)
{
    using var ms = new MemoryStream();
    await dosya.CopyToAsync(ms);
    return ms.ToArray();
}

// =====================================================================
// AYARLAR
// =====================================================================

record MetinSorusu(string? Metin, bool HerSatirAyriMadde);

class EkEkip
{
    public string Ad { get; set; } = "";
    public string Aciklama { get; set; } = "";
}

class YzAyarlari
{
    public string QwenApiUrl { get; set; } = "";
    public string QwenApiKey { get; set; } = "";
    public string AiModel { get; set; } = "zt-ga-small-0";
    public double Temperature { get; set; } = 0.1;
    public int MaksToken { get; set; } = 8000;
    public bool JsonModu { get; set; }
    public bool DusunmeyiKapat { get; set; }
    public int ZamanAsimiDakika { get; set; } = 5;
    public bool SertifikaDogrulamasiniAtla { get; set; } = true;
    public int Paralellik { get; set; } = 3;
    public int DenemeSayisi { get; set; } = 3;
    public bool OnElemeYap { get; set; }
    public int AdayEkipSayisi { get; set; } = 20;
    public int EkipBasinaServisSayisi { get; set; } = 12;
    public int MaksEkipSayisi { get; set; } = 3;
    public bool ServisAciklamasiGonder { get; set; } = true;
    public int ServisAciklamaMaksKarakter { get; set; } = 250;
    /// Prosedür listesi yüklüyse her maddeyle en ilgili kaç prosedür modele gönderilsin (0 = gönderme)
    public int IlgiliProsedurSayisi { get; set; } = 15;
    /// Adı çok genel olduğu için her işi kendine çeken ekipler; prompt'ta bu ekipler için ayrıca uyarı verilir
    public List<string>? GenelAdliEkipler { get; set; }
    public List<string> GenelEkipler() => GenelAdliEkipler ?? ["YAPAY ZEKA UYGULAMALARI"];
    /// Prompt'a eklenen kurum kuralları. {EKİP ADI} yazılan yerler ekibin [E...] etiketiyle değiştirilir.
    public List<string>? EkKurallar { get; set; }
    public List<string> Kurallar() => EkKurallar ??
    [
        "İş birimlerinden gelen taleplerin alınması ve yönetimi Pergel uygulaması üzerinden {PROJE YÖNETİMİ OFİSİ} tarafından yürütülür. " +
        "Bu nedenle talep yönetimi, kapsamın belirlenmesi ve kapsam değişiklikleri, paydaşların belirlenmesi ve sürece dahil edilmesi, " +
        "paydaş ve proje iletişimi (bilgilendirme, toplantı, raporlama) konularının geneli {PROJE YÖNETİMİ OFİSİ} sorumluluğundadır; " +
        "bu maddelerde Sorumlu ekip {PROJE YÖNETİMİ OFİSİ} olmalı, konunun uzmanı ekipler gerekiyorsa Destek olarak eklenmelidir."
    ];
}

class ExcelAyarlari
{
    public string KullaniciEkipKolonu { get; set; } = "E";
    public string KullaniciNotKolonu { get; set; } = "F";
    public string CiktiEkipBasligi { get; set; } = "YZ Önerilen Ekipler ve Görevleri";
    public string CiktiGerekceBasligi { get; set; } = "YZ Karar Gerekçesi";
    public string KatalogAciklamaKolonu { get; set; } = "F";
    public string CiktiProsedurBasligi { get; set; } = "YZ İlgili Mevcut Prosedürler";
    /// Prosedür listesinde başlıklar bulunamazsa kullanılacak kolon harfleri
    public string ProsedurAdKolonu { get; set; } = "D";
    public string ProsedurBirimKolonu { get; set; } = "H";
    /// Kontrol hücresinde birden fazla madde varsa (-, •, 1., a) veya alt alta satırlar) her biri ayrı değerlendirilir
    public bool HucreMaddeleriniAyir { get; set; } = true;
    /// Katalogda olmayan ama kurumda var olan ekipler (ör. Proje Yönetimi Ofisi). Boş bırakılırsa varsayılan liste kullanılır.
    public List<EkEkip>? KatalogaEklenecekEkipler { get; set; }

    public List<EkEkip> EkEkipler() => KatalogaEklenecekEkipler ??
    [
        new EkEkip
        {
            Ad = "PROJE YÖNETİMİ OFİSİ",
            Aciklama = "Kurum genelinde proje yönetimi ofisi (PMO): iş birimlerinden gelen taleplerin Pergel uygulaması üzerinden alınması, değerlendirilmesi ve takibi (talep yönetimi); proje başlatma ve onay süreçleri, kapsam yönetimi ve kapsam değişiklikleri, paydaşların belirlenmesi ve sürece dahil edilmesi, paydaş ve proje iletişimi (bilgilendirme, toplantı, raporlama), proje planlama, kaynak ve bütçe takibi, proje portföy yönetimi, proje yönetişimi, kilometre taşı ve risk takibi, proje durum raporlaması, ekip görevlendirme koordinasyonu."
        }
    ];
}

// =====================================================================
// MODELLER
// =====================================================================

record Servis(string Kod, string Ad, string Bolum, string Yonetici, string Kategori, string Aciklama);

record Mesaj(string role, string content);

record EkipOnerisi(string Ekip, string Rol, string Gorev, bool KatalogdaVar);

record ProsedurOnerisi(string Kod, string Ad, string Birim, string Link, string Iliski);

class Prosedur
{
    public string Tip { get; init; } = "";
    public string Kod { get; init; } = "";
    public string DokumanKodu { get; init; } = "";
    public string Ad { get; init; } = "";
    public string Birim { get; init; } = "";
    public string SurecModeli { get; init; } = "";
    public string Link { get; init; } = "";
    /// Prompt'ta ve yanıtta kullanılan kimlik: Doküman Kodu, yoksa Kod
    public string Kimlik => DokumanKodu.Length > 0 ? DokumanKodu : Kod;
}

class Ekip
{
    public required string Ad { get; init; }
    public required string Norm { get; init; }
    /// Prompt'ta kullanılan kısa kod (E001...). Küçük modeller uzun adı yanlış kopyalasa da kodu doğru yazar.
    public string Kod { get; set; } = "";
    public string Etiket => $"[{Kod}] {Ad}";
    public List<Servis> Servisler { get; } = [];
    public HashSet<string> AdKokleri { get; set; } = [];
    public Dictionary<string, double> Kokler { get; } = [];
    public double Uzunluk { get; set; }
}

class KontrolSatiri
{
    public int Sira { get; init; }
    public int SatirNo { get; init; }
    public string AltKategori { get; init; } = "";
    public string Aksiyon { get; init; } = "";
    public List<KeyValuePair<string, string>> Alanlar { get; init; } = [];
    public string KullaniciEkip { get; init; } = "";
    public string KullaniciNot { get; init; } = "";
    public bool Serbest { get; init; }
    /// Hücre birden fazla maddeye bölündüyse: bu maddenin sırası, toplam madde sayısı ve hücrenin tamamı
    public int AltNo { get; init; } = 1;
    public int AltToplam { get; init; } = 1;
    public string HucreMetni { get; init; } = "";

    // İşlem sonucu (arka plan görevleri tarafından güncellenir)
    public volatile string Durum = "Bekliyor";
    public List<EkipOnerisi> Ekipler = [];
    public string Gerekce = "";
    public int? Guven;
    public string Uyum = "";
    public string Hata = "";
    public string HamCevap = "";
    public List<string> Adaylar = [];
    public List<ProsedurOnerisi> Prosedurler = [];
    public List<string> GonderilenProsedurler = [];
    public int PromptKarakter;

    public string EkipMetni() => string.Join("\n", Ekipler.Select(e =>
        $"• {e.Ekip} ({e.Rol}){(e.KatalogdaVar ? "" : " [katalogda bulunamadı]")}: {e.Gorev}"));

    public string ProsedurMetni() => Durum != "Tamam" ? "" : Prosedurler.Count == 0
        ? (GonderilenProsedurler.Count > 0 ? "İlgili mevcut prosedür bulunamadı." : "")
        : string.Join("\n", Prosedurler.Select(p => $"• {p.Kod} {p.Ad} ({p.Birim}){(p.Iliski.Length > 0 ? ": " + p.Iliski : "")}"));

    public string GerekceMetni() => Durum == "Tamam"
        ? $"Güven: %{Guven ?? 0}\n{Gerekce}"
        : (string.IsNullOrEmpty(Hata) ? "" : $"HATA: {Hata}");
}

class AnalizIsi
{
    public string Id { get; } = Guid.NewGuid().ToString("N")[..12];
    public DateTime Baslangic { get; } = DateTime.Now;
    public required string DosyaAdi { get; init; }
    /// Metinle sorulan işlerde null'dır; Excel çıktısı sıfırdan üretilir.
    public byte[]? KontrolBytes { get; init; }
    public required KontrolDosyasi Kontrol { get; init; }
    public required Katalog Katalog { get; init; }
    public ProsedurKutuphanesi? Prosedurler { get; init; }
    public CancellationTokenSource Iptal { get; } = new();
    public volatile string GenelDurum = "Çalışıyor";

    public object Durum() => new
    {
        id = Id,
        tur = KontrolBytes is null ? "Metin" : "Excel",
        prosedurVar = Prosedurler is not null,
        dosyaAdi = DosyaAdi,
        baslangic = Baslangic,
        durum = GenelDurum,
        toplam = Kontrol.Satirlar.Count,
        tamamlanan = Kontrol.Satirlar.Count(s => s.Durum == "Tamam"),
        hatali = Kontrol.Satirlar.Count(s => s.Durum == "Hata"),
        ekipSayisi = Katalog.Ekipler.Count,
        servisSayisi = Katalog.Ekipler.Sum(e => e.Servisler.Count),
        aciklamaliServis = Katalog.Ekipler.Sum(e => e.Servisler.Count(sv => sv.Aciklama.Length > 0)),
        ekipler = Katalog.Ekipler.Select(e => e.Ad),
        satirlar = Kontrol.Satirlar.Select(s => new
        {
            sira = s.Sira,
            satirNo = s.SatirNo,
            altNo = s.AltNo,
            altToplam = s.AltToplam,
            altKategori = s.AltKategori,
            aksiyon = s.Aksiyon,
            alanlar = s.Alanlar.Select(a => new { ad = a.Key, deger = a.Value }),
            kullaniciEkip = s.KullaniciEkip,
            kullaniciNot = s.KullaniciNot,
            durum = s.Durum,
            ekipler = s.Ekipler.Select(e => new { ekip = e.Ekip, rol = e.Rol, gorev = e.Gorev, katalogdaVar = e.KatalogdaVar }),
            gerekce = s.Gerekce,
            guven = s.Guven,
            uyum = s.Uyum,
            hata = s.Hata,
            hamCevap = s.HamCevap,
            adaylar = s.Adaylar,
            prosedurler = s.Prosedurler.Select(p => new { kod = p.Kod, ad = p.Ad, birim = p.Birim, link = p.Link, iliski = p.Iliski }),
            gonderilenProsedurler = s.GonderilenProsedurler,
            promptKarakter = s.PromptKarakter
        })
    };
}

class ExcelOkumaHatasi(string mesaj) : Exception(mesaj);

// =====================================================================
// METİN YARDIMCILARI (Türkçe normalizasyon, kök bulma)
// =====================================================================

static class Metin
{
    static readonly CultureInfo Tr = new("tr-TR");

    static readonly HashSet<string> Durdurma = new(new[]
    {
        "ve", "veya", "ile", "için", "bir", "bu", "şu", "da", "de", "ki", "olarak", "olan", "olmalı", "olmalıdır",
        "edilmelidir", "yapılmalıdır", "tarafından", "gibi", "vb", "vs", "ilgili", "her", "tüm", "göre", "kadar",
        "daha", "en", "mi", "ne", "çok", "ise", "hem", "ama", "fakat", "sonra", "önce", "üzerinde", "üzerinden",
        "içinde", "içerisinde", "konusunda", "yapılması", "edilmesi", "sağlanmalıdır", "bulunmalıdır", "gerekir",
        "gereklidir", "olup", "olacak", "var", "yok", "aksiyon", "kontrol", "the", "and", "for", "of", "to", "in",
        // Sürecin tamamı YZ ile ilgili; bu kelimeler ayırt edici değildir ve adında geçen ekibi kayırır
        "yapay", "zeka", "yz", "ai", "yapayzeka"
    }.Select(Normalize));

    public static string Normalize(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        var sb = new StringBuilder(s.Length);
        foreach (var ch in s.ToLower(Tr))
        {
            sb.Append(ch switch
            {
                'ı' => 'i', 'ğ' => 'g', 'ü' => 'u', 'ş' => 's', 'ö' => 'o', 'ç' => 'c', 'â' => 'a', 'î' => 'i', 'û' => 'u',
                _ => char.IsLetterOrDigit(ch) ? ch : ' '
            });
        }
        return Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
    }

    /// Türkçe eklerden etkilenmemek için kelimenin ilk 5 harfini kök kabul eder.
    public static IEnumerable<string> Kokler(string? s) =>
        Normalize(s).Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Where(t => t.Length >= 2 && !Durdurma.Contains(t) && !t.All(char.IsDigit))
            .Select(t => t.Length > 5 ? t[..5] : t);

    public static string Kisalt(string? s, int n) =>
        string.IsNullOrEmpty(s) ? "" : s.Length <= n ? s : s[..n] + "…";
}

// =====================================================================
// SERVİS KATALOĞU (ekip profilleri + aday ön eleme + isim eşleştirme)
// =====================================================================

class Katalog
{
    public List<Ekip> Ekipler { get; }
    readonly Dictionary<string, double> idf = [];
    readonly double ortUzunluk;

    public Katalog(IEnumerable<Servis> servisler)
    {
        var gruplar = new Dictionary<string, Ekip>();
        foreach (var sv in servisler)
        {
            var norm = Metin.Normalize(sv.Bolum);
            if (norm.Length == 0) continue;
            if (!gruplar.TryGetValue(norm, out var ekip))
                gruplar[norm] = ekip = new Ekip { Ad = sv.Bolum, Norm = norm };
            ekip.Servisler.Add(sv);
        }
        Ekipler = gruplar.Values.OrderBy(e => e.Ad, StringComparer.Create(new CultureInfo("tr-TR"), true)).ToList();
        for (int i = 0; i < Ekipler.Count; i++) Ekipler[i].Kod = $"E{i + 1:000}";

        foreach (var e in Ekipler)
        {
            e.AdKokleri = Metin.Kokler(e.Ad).ToHashSet();
            foreach (var k in e.AdKokleri) e.Kokler[k] = e.Kokler.GetValueOrDefault(k) + 3; // ekip adı daha ağırlıklı
            foreach (var sv in e.Servisler)
            {
                foreach (var k in Metin.Kokler(sv.Ad))
                    e.Kokler[k] = e.Kokler.GetValueOrDefault(k) + 1;
                foreach (var k in Metin.Kokler(sv.Aciklama))
                    e.Kokler[k] = e.Kokler.GetValueOrDefault(k) + 0.5; // açıklama daha uzun ve gürültülü
            }
            e.Uzunluk = e.Kokler.Values.Sum();
        }
        ortUzunluk = Ekipler.Count == 0 ? 1 : Math.Max(1, Ekipler.Average(e => e.Uzunluk));

        var df = new Dictionary<string, int>();
        foreach (var k in Ekipler.SelectMany(e => e.Kokler.Keys)) df[k] = df.GetValueOrDefault(k) + 1;
        foreach (var (k, n) in df) idf[k] = Math.Log(1 + (double)Ekipler.Count / n);
    }

    /// Model ya da kullanıcı tarafından yazılmış bir ekip adını katalogdaki ekiple eşleştirir.
    static readonly Regex KodDeseni = new(@"^\s*\[?\s*E\s*0*(\d{1,4})\s*\]?\s*[-:.)]?\s*", RegexOptions.IgnoreCase);

    public Ekip? EkipBul(string? ad)
    {
        // Önce [E012] kodu: model adı bozsa bile kod doğruysa eşleşir
        if (ad is not null && KodDeseni.Match(ad) is { Success: true } km)
        {
            var no = int.Parse(km.Groups[1].Value);
            if (no >= 1 && no <= Ekipler.Count) return Ekipler[no - 1];
            ad = ad[km.Length..];
        }
        var n = Metin.Normalize(ad);
        if (n.Length < 2) return null;

        var tam = Ekipler.FirstOrDefault(e => e.Norm == n);
        if (tam is not null) return tam;

        if (n.Length >= 4)
        {
            var icerir = Ekipler.Where(e => e.Norm.Contains(n) || n.Contains(e.Norm))
                .OrderBy(e => Math.Abs(e.Norm.Length - n.Length)).FirstOrDefault();
            if (icerir is not null) return icerir;
        }

        var kokler = Metin.Kokler(ad).ToHashSet();
        if (kokler.Count == 0) return null;
        var (enIyi, skor) = Ekipler
            .Select(e =>
            {
                var ortak = e.AdKokleri.Intersect(kokler).Count();
                var birlesim = e.AdKokleri.Union(kokler).Count();
                var kapsama = (double)ortak / Math.Max(1, Math.Min(e.AdKokleri.Count, kokler.Count));
                var jaccard = (double)ortak / Math.Max(1, birlesim);
                return (e, Math.Max(jaccard, ortak >= 2 ? kapsama * 0.8 : 0));
            })
            .OrderByDescending(x => x.Item2).FirstOrDefault();
        return skor >= 0.6 ? enIyi : null;
    }

    /// "A, B; C" biçimindeki serbest metindeki ekip adlarını eşleştirir.
    public List<Ekip> MetindenEkipler(string? metin) =>
        string.IsNullOrWhiteSpace(metin) ? [] :
        Regex.Split(metin, @"[,;/\n\r]+|\s-\s")
            .Select(p => EkipBul(p.Trim()))
            .OfType<Ekip>().Distinct().ToList();

    public Dictionary<string, double> Sorgu(KontrolSatiri s)
    {
        var q = new Dictionary<string, double>();
        void Ekle(string? metin, double w)
        {
            foreach (var k in Metin.Kokler(metin)) q[k] = Math.Max(q.GetValueOrDefault(k), w);
        }
        foreach (var a in s.Alanlar) Ekle(a.Value, 1);
        Ekle(s.KullaniciNot, 1);
        Ekle(s.KullaniciEkip, 1.5);
        Ekle(s.AltKategori, 1.5);
        Ekle(s.Aksiyon, 2);
        return q;
    }

    /// Basit BM25 benzeri puanlama ile en olası aday ekipleri seçer.
    /// Kullanıcının E kolonunda yazdığı ekipler her zaman aday listesine eklenir.
    public List<Ekip> AdaylariSec(KontrolSatiri s, int adet)
    {
        if (adet <= 0 || adet >= Ekipler.Count) return Ekipler;

        var q = Sorgu(s);
        var kullanici = MetindenEkipler(s.KullaniciEkip).ToHashSet();
        return Ekipler
            .Select(e =>
            {
                double skor = 0;
                foreach (var (k, w) in q)
                    if (e.Kokler.TryGetValue(k, out var tf))
                        skor += w * idf[k] * (tf * 2.2) / (tf + 1.2 * (0.25 + 0.75 * e.Uzunluk / ortUzunluk));
                return (e, skor: skor + (kullanici.Contains(e) ? 1000 : 0));
            })
            .OrderByDescending(x => x.skor)
            .Take(adet + kullanici.Count)
            .Where((x, i) => i < adet || kullanici.Contains(x.e))
            .Select(x => x.e)
            .ToList();
    }

    /// Bir ekibin servislerini maddeyle en ilgili olanlar önde olacak şekilde sıralar.
    public IEnumerable<Servis> IlgiliServisler(Ekip e, Dictionary<string, double> q) =>
        e.Servisler.OrderByDescending(sv =>
            Metin.Kokler(sv.Ad).Distinct().Sum(k => q.TryGetValue(k, out var w) ? w * idf.GetValueOrDefault(k) : 0) +
            0.5 * Metin.Kokler(sv.Aciklama).Distinct().Sum(k => q.TryGetValue(k, out var w) ? w * idf.GetValueOrDefault(k) : 0));
}

// =====================================================================
// PROSEDÜR KÜTÜPHANESİ (mevcut süreç dokümanları; maddeyle en ilgilileri seçilir)
// =====================================================================

class ProsedurKutuphanesi
{
    public List<Prosedur> Liste { get; }
    readonly Katalog katalog;
    readonly Dictionary<Prosedur, Ekip?> birimEkip = [];
    readonly Dictionary<Prosedur, Dictionary<string, double>> kokler = [];
    readonly Dictionary<string, double> idf = [];

    public ProsedurKutuphanesi(List<Prosedur> liste, Katalog katalog)
    {
        (Liste, this.katalog) = (liste, katalog);
        var df = new Dictionary<string, int>();
        foreach (var p in liste)
        {
            birimEkip[p] = katalog.EkipBul(p.Birim);
            var k = new Dictionary<string, double>();
            foreach (var kok in Metin.Kokler(p.Ad)) k[kok] = k.GetValueOrDefault(kok) + 1;
            foreach (var kok in Metin.Kokler(p.SurecModeli)) k[kok] = k.GetValueOrDefault(kok) + 0.5;
            kokler[p] = k;
            foreach (var kok in k.Keys) df[kok] = df.GetValueOrDefault(kok) + 1;
        }
        foreach (var (kok, n) in df) idf[kok] = Math.Log(1 + (double)liste.Count / n);
    }

    public int EslesmeyenBirimSayisi => Liste.Select(p => p.Birim).Distinct().Count(b => katalog.EkipBul(b) is null);
    public Ekip? Ekibi(Prosedur p) => birimEkip.GetValueOrDefault(p);

    public List<Prosedur> IlgilileriSec(Dictionary<string, double> q, int adet)
    {
        if (adet <= 0) return [];
        return Liste
            .Select(p => (p, skor: kokler[p].Sum(kv => q.TryGetValue(kv.Key, out var w) ? w * idf[kv.Key] * (1 + Math.Log(kv.Value + 0.5)) : 0)
                                  / Math.Sqrt(1 + kokler[p].Count)))
            .Where(x => x.skor > 0)
            .OrderByDescending(x => x.skor)
            .Take(adet)
            .Select(x => x.p)
            .ToList();
    }

    public Prosedur? Bul(string? kimlik)
    {
        var n = Metin.Normalize(kimlik);
        if (n.Length == 0) return null;
        return Liste.FirstOrDefault(p => Metin.Normalize(p.DokumanKodu) == n || Metin.Normalize(p.Kod) == n);
    }
}

// =====================================================================
// HÜCRE İÇİ MADDE AYIRMA
// =====================================================================

static class MaddeAyirici
{
    // Satır başı madde işaretleri: -, •, *, 1. 1) 1-, a) a.
    static readonly Regex Isaret = new(@"^\s*(?:[-–—•*▪●○◦·►✓]|\d{1,2}\s*[.)\-]|[a-zA-ZçğıöşüÇĞİÖŞÜ]\s*[)])\s*(?=\S)");
    // Aynı satır içinde "1) ... 2) ..." veya "a) ... b) ..."
    static readonly Regex SatirIci = new(@"(?:^|(?<=[\s;:,.]))(?:\d{1,2}|[a-zçğıöşü])\)\s+");

    public static List<string> Ayir(string metin)
    {
        var satirlar = metin.Split('\n').Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
        if (satirlar.Count >= 2)
        {
            // "Aşağıdakiler yapılmalıdır:" gibi giriş satırı maddelere bağlam olarak eklenir
            var giris = satirlar[0].EndsWith(':') && !Isaret.IsMatch(satirlar[0]) ? satirlar[0] : "";
            var govde = giris.Length > 0 ? satirlar.Skip(1).ToList() : satirlar;
            var maddeler = new List<string>();
            if (govde.Count(x => Isaret.IsMatch(x)) >= 2)
            {
                foreach (var x in govde)
                {
                    if (Isaret.IsMatch(x) || maddeler.Count == 0) maddeler.Add(Isaret.Replace(x, "", 1).Trim());
                    else maddeler[^1] += " " + x; // işaretsiz satır önceki maddenin devamıdır
                }
            }
            else if (govde.Count >= 2 && govde.All(x => x.Length >= 15))
            {
                maddeler = govde; // işaretsiz ama alt alta yazılmış ayrı cümleler
            }
            if (maddeler.Count >= 2)
                return maddeler.Select(m => giris.Length > 0 ? $"{giris} {m}" : m).ToList();
        }

        var eslesmeler = SatirIci.Matches(metin);
        if (eslesmeler.Count >= 2)
        {
            var giris = metin[..eslesmeler[0].Index].Trim();
            var maddeler = new List<string>();
            for (int i = 0; i < eslesmeler.Count; i++)
            {
                int bas = eslesmeler[i].Index + eslesmeler[i].Length;
                int son = i + 1 < eslesmeler.Count ? eslesmeler[i + 1].Index : metin.Length;
                var m = metin[bas..son].Trim().TrimEnd(';', ',').Trim();
                if (m.Length > 0) maddeler.Add(giris.Length > 0 ? $"{giris} {m}" : m);
            }
            if (maddeler.Count >= 2) return maddeler;
        }
        return [metin];
    }
}

// =====================================================================
// EXCEL OKUMA / YAZMA
// =====================================================================

class KontrolDosyasi
{
    public required string SayfaAdi { get; init; }
    public int BaslikSatiri { get; init; }
    public int CiktiEkipKolonu { get; init; }
    public int CiktiGerekceKolonu { get; init; }
    public int CiktiProsedurKolonu { get; init; }
    public required List<KontrolSatiri> Satirlar { get; init; }
}

static class ExcelIslemleri
{
    static string Deger(IXLCell c)
    {
        var hucre = c.IsMerged() ? c.MergedRange().FirstCell() : c;
        return Regex.Replace(hucre.GetFormattedString() ?? "", @"[ \t]+", " ").Trim();
    }

    static int KolonNo(string harf) => XLHelper.GetColumnNumberFromLetter(harf.Trim().ToUpperInvariant());

    static XLWorkbook Ac(byte[] bytes, string tur)
    {
        try { return new XLWorkbook(new MemoryStream(bytes)); }
        catch (Exception ex) { throw new ExcelOkumaHatasi($"{tur} dosyası okunamadı (.xlsx olmalı): {ex.Message}"); }
    }

    /// İlk 20 satırda koşulu sağlayan başlık hücreleri olan ilk satırı bulur.
    static (IXLWorksheet ws, int satir)? BaslikBul(XLWorkbook wb, Func<List<string>, bool> kosul)
    {
        foreach (var ws in wb.Worksheets)
        {
            var son = Math.Min(20, ws.LastRowUsed()?.RowNumber() ?? 0);
            for (int r = 1; r <= son; r++)
            {
                var basliklar = ws.Row(r).CellsUsed().Select(c => Metin.Normalize(c.GetFormattedString())).ToList();
                if (kosul(basliklar)) return (ws, r);
            }
        }
        return null;
    }

    public static Katalog KatalogOku(byte[] bytes, ExcelAyarlari ayar)
    {
        using var wb = Ac(bytes, "Servis kataloğu");
        var bulunan = BaslikBul(wb, b => b.Any(x => x.Contains("servis")) && b.Any(x => x.Contains("sahibi") || x.Contains("bolum")))
            ?? throw new ExcelOkumaHatasi("Servis kataloğunda 'Servis Adı' ve 'Servis Sahibi Bölüm' başlıkları bulunamadı.");
        var (ws, baslik) = bulunan;

        int Kolon(params string[] anahtarlar) => ws.Row(baslik).CellsUsed()
            .FirstOrDefault(c => anahtarlar.Any(a => Metin.Normalize(c.GetFormattedString()).Contains(a)))?.Address.ColumnNumber ?? 0;

        int kod = Kolon("servis kodu", "kod"), ad = Kolon("servis adi", "servis ad"), bolum = Kolon("sahibi bolum", "sahibi", "bolum"),
            yonetici = Kolon("yonetici"), kategori = Kolon("kategori");
        // Servis açıklaması: önce "Açıklama" başlığı aranır, bulunamazsa ayardaki kolon harfi (varsayılan F) kullanılır
        int aciklama = Kolon("aciklama", "tanim");
        if (aciklama == 0 && !string.IsNullOrWhiteSpace(ayar.KatalogAciklamaKolonu)) aciklama = KolonNo(ayar.KatalogAciklamaKolonu);
        if (ad == 0 || bolum == 0)
            throw new ExcelOkumaHatasi("Servis kataloğunda 'Servis Adı' veya 'Servis Sahibi Bölüm' kolonu bulunamadı.");

        string Al(IXLRow r, int k) => k == 0 ? "" : Deger(r.Cell(k));
        var servisler = new List<Servis>();
        var sonSatir = ws.LastRowUsed()?.RowNumber() ?? baslik;
        for (int i = baslik + 1; i <= sonSatir; i++)
        {
            var r = ws.Row(i);
            var sv = new Servis(Al(r, kod), Al(r, ad), Al(r, bolum), Al(r, yonetici), Al(r, kategori),
                Regex.Replace(Al(r, aciklama), @"\s+", " ").Trim());
            if (sv.Ad.Length > 0 && sv.Bolum.Length > 0) servisler.Add(sv);
        }

        if (servisler.Count == 0) throw new ExcelOkumaHatasi("Servis kataloğunda hiç ekip (Servis Sahibi Bölüm) bulunamadı.");

        // Katalogda olmayan ama kurumda var olan ekipler; adı katalogdaki bir ekiple eşleşirse o ekibe görev tanımı olarak eklenir
        var ilk = new Katalog(servisler);
        foreach (var ek in ayar.EkEkipler().Where(e => !string.IsNullOrWhiteSpace(e.Ad)))
            servisler.Add(new Servis("EK", "Genel görev tanımı", ilk.EkipBul(ek.Ad)?.Ad ?? ek.Ad.Trim(), "", "Ek tanım (katalog dışı)", ek.Aciklama.Trim()));
        return new Katalog(servisler);
    }

    public static List<Prosedur> ProsedurOku(byte[] bytes, ExcelAyarlari ayar)
    {
        using var wb = Ac(bytes, "Prosedür listesi");
        // Başlıklar bulunamazsa ilk sayfanın 1. satırı başlık kabul edilir; ad ve birim D / H kolonlarından okunur
        var (ws, baslik) = BaslikBul(wb, b => b.Any(x => x.Contains("dokuman adi")) && b.Any(x => x.Contains("surec sahibi") || x.Contains("birim")))
            ?? (wb.Worksheet(1), 1);
        var basliklar = ws.Row(baslik).CellsUsed().Select(c => (no: c.Address.ColumnNumber, ad: Metin.Normalize(c.GetFormattedString()))).ToList();
        int Kolon(Func<string, bool> kosul) => basliklar.FirstOrDefault(b => kosul(b.ad)).no;

        int tip = Kolon(a => a.Contains("mevzuat tipi") || a == "tip"), kod = Kolon(a => a == "kod"), dkod = Kolon(a => a.Contains("dokuman kodu")),
            ad = Kolon(a => a.Contains("dokuman adi")), birim = Kolon(a => a.Contains("surec sahibi") || a.Contains("birim")),
            model = Kolon(a => a.Contains("surec modeli")), durum = Kolon(a => a == "durum"), link = Kolon(a => a.Contains("link"));
        if (ad == 0) ad = KolonNo(ayar.ProsedurAdKolonu);
        if (birim == 0) birim = KolonNo(ayar.ProsedurBirimKolonu);

        string Al(IXLRow r, int k) => k == 0 ? "" : Deger(r.Cell(k));
        string LinkAl(IXLRow r)
        {
            if (link == 0) return "";
            var c = r.Cell(link);
            try { if (c.HasHyperlink && c.GetHyperlink().ExternalAddress is { } u) return u.ToString(); } catch { }
            var metin = Deger(c);
            return metin.StartsWith("http", StringComparison.OrdinalIgnoreCase) ? metin : "";
        }

        var liste = new List<Prosedur>();
        var sonSatir = ws.LastRowUsed()?.RowNumber() ?? baslik;
        for (int i = baslik + 1; i <= sonSatir; i++)
        {
            var r = ws.Row(i);
            if (Al(r, ad).Length == 0) continue;
            // Pasif dokümanlar alınmaz
            if (durum != 0 && Metin.Normalize(Al(r, durum)) is "false" or "0" or "pasif" or "hayir" or "yururlukten kalkti") continue;
            liste.Add(new Prosedur
            {
                Tip = Al(r, tip), Kod = Al(r, kod), DokumanKodu = Al(r, dkod), Ad = Al(r, ad),
                Birim = Al(r, birim), SurecModeli = Al(r, model), Link = LinkAl(r)
            });
        }
        if (liste.Count == 0) throw new ExcelOkumaHatasi("Prosedür listesinde geçerli (Durum = True) satır bulunamadı.");
        return liste;
    }

    public static KontrolDosyasi KontrolOku(byte[] bytes, ExcelAyarlari ayar)
    {
        using var wb = Ac(bytes, "Kontrol listesi");
        var bulunan = BaslikBul(wb, b => b.Any(x => x.Contains("aksiyon") || x.Contains("kontrol")) && b.Count >= 2)
            ?? throw new ExcelOkumaHatasi("Kontrol listesinde 'Aksiyon / Kontrol' başlığı bulunamadı.");
        var (ws, baslik) = bulunan;

        int ekipKolon = KolonNo(ayar.KullaniciEkipKolonu), notKolon = KolonNo(ayar.KullaniciNotKolonu);
        var cikti1Norm = Metin.Normalize(ayar.CiktiEkipBasligi);
        var cikti2Norm = Metin.Normalize(ayar.CiktiGerekceBasligi);
        var cikti3Norm = Metin.Normalize(ayar.CiktiProsedurBasligi);

        var basliklar = ws.Row(baslik).CellsUsed().ToDictionary(c => c.Address.ColumnNumber, c => Deger(c));
        int sonKolon = Math.Max(basliklar.Keys.DefaultIfEmpty(1).Max(), notKolon);

        // Dosya daha önce bu programdan çıkmışsa çıktı kolonlarını yeniden kullan
        int cikti1 = basliklar.FirstOrDefault(b => Metin.Normalize(b.Value) == cikti1Norm).Key;
        int cikti2 = basliklar.FirstOrDefault(b => Metin.Normalize(b.Value) == cikti2Norm).Key;
        if (cikti1 == 0) cikti1 = sonKolon + 1;
        if (cikti2 == 0) cikti2 = Math.Max(sonKolon, cikti1) + 1;
        int cikti3 = basliklar.FirstOrDefault(b => Metin.Normalize(b.Value) == cikti3Norm).Key;
        if (cikti3 == 0) cikti3 = Math.Max(sonKolon, Math.Max(cikti1, cikti2)) + 1;

        var haric = new HashSet<int> { ekipKolon, notKolon, cikti1, cikti2, cikti3 };
        var aksiyonKolon = basliklar.Where(b => !haric.Contains(b.Key))
            .OrderByDescending(b => Metin.Normalize(b.Value).Contains("aksiyon") ? 2 : Metin.Normalize(b.Value).Contains("kontrol") ? 1 : 0)
            .First().Key;
        var altKategoriKolon = basliklar.FirstOrDefault(b => Metin.Normalize(b.Value).Contains("alt kategori")).Key;
        var bilgiKolonlari = basliklar.Keys.Where(k => !haric.Contains(k) && k != aksiyonKolon && k != altKategoriKolon).OrderBy(k => k).ToList();

        var satirlar = new List<KontrolSatiri>();
        var sonSatir = ws.LastRowUsed()?.RowNumber() ?? baslik;
        for (int i = baslik + 1; i <= sonSatir; i++)
        {
            var r = ws.Row(i);
            var aksiyon = Deger(r.Cell(aksiyonKolon));
            if (aksiyon.Length == 0) continue;
            var altKategori = altKategoriKolon == 0 ? "" : Deger(r.Cell(altKategoriKolon));
            var alanlar = bilgiKolonlari
                .Select(k => new KeyValuePair<string, string>(string.IsNullOrEmpty(basliklar[k]) ? $"Kolon {XLHelper.GetColumnLetterFromNumber(k)}" : basliklar[k], Deger(r.Cell(k))))
                .Where(kv => kv.Value.Length > 0).ToList();
            var maddeler = ayar.HucreMaddeleriniAyir ? MaddeAyirici.Ayir(aksiyon) : [aksiyon];
            for (int j = 0; j < maddeler.Count; j++)
            {
                satirlar.Add(new KontrolSatiri
                {
                    Sira = satirlar.Count + 1,
                    SatirNo = i,
                    Aksiyon = maddeler[j],
                    AltNo = j + 1,
                    AltToplam = maddeler.Count,
                    HucreMetni = maddeler.Count > 1 ? aksiyon : "",
                    AltKategori = altKategori,
                    Alanlar = alanlar,
                    KullaniciEkip = Deger(r.Cell(ekipKolon)),
                    KullaniciNot = Deger(r.Cell(notKolon))
                });
            }
        }
        if (satirlar.Count == 0) throw new ExcelOkumaHatasi("Kontrol listesinde 'Aksiyon / Kontrol' kolonu dolu hiç satır bulunamadı.");

        return new KontrolDosyasi
        {
            SayfaAdi = ws.Name,
            BaslikSatiri = baslik,
            CiktiEkipKolonu = cikti1,
            CiktiGerekceKolonu = cikti2,
            CiktiProsedurKolonu = cikti3,
            Satirlar = satirlar
        };
    }

    /// Metinle sorulan işler için: İş / Ekipler / Gerekçe kolonlu yeni bir dosya üretir.
    static byte[] MetinSonucuYaz(AnalizIsi analiz, ExcelAyarlari ayar)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sonuç");
        var basliklar = new List<string> { "No", "İş", ayar.CiktiEkipBasligi, ayar.CiktiGerekceBasligi };
        if (analiz.Prosedurler is not null) basliklar.Add(ayar.CiktiProsedurBasligi);
        for (int i = 0; i < basliklar.Count; i++)
        {
            var h = ws.Cell(1, i + 1);
            h.Value = basliklar[i];
            h.Style.Font.Bold = true;
            h.Style.Fill.BackgroundColor = XLColor.FromHtml("#DDEBF7");
        }
        int r = 2;
        foreach (var s in analiz.Kontrol.Satirlar)
        {
            ws.Cell(r, 1).Value = s.Sira;
            ws.Cell(r, 2).Value = s.Aksiyon;
            ws.Cell(r, 3).Value = s.EkipMetni();
            ws.Cell(r, 4).Value = s.GerekceMetni();
            if (analiz.Prosedurler is not null) ws.Cell(r, 5).Value = s.ProsedurMetni();
            r++;
        }
        ws.Column(1).Width = 6; ws.Column(2).Width = 60; ws.Column(3).Width = 70; ws.Column(4).Width = 80; ws.Column(5).Width = 70;
        ws.Range(2, 1, Math.Max(2, r - 1), basliklar.Count).Style.Alignment.WrapText = true;
        ws.Range(2, 1, Math.Max(2, r - 1), basliklar.Count).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        using var cikti = new MemoryStream();
        wb.SaveAs(cikti);
        return cikti.ToArray();
    }

    /// Yüklenen orijinal dosyaya iki sonuç kolonunu ekleyip döner; diğer hücrelere dokunmaz.
    public static byte[] SonucYaz(AnalizIsi analiz, ExcelAyarlari ayar)
    {
        if (analiz.KontrolBytes is null) return MetinSonucuYaz(analiz, ayar);
        using var wb = new XLWorkbook(new MemoryStream(analiz.KontrolBytes));
        var k = analiz.Kontrol;
        var ws = wb.Worksheet(k.SayfaAdi);

        var kolonlar = new List<(int kolon, string baslik, double genislik)> { (k.CiktiEkipKolonu, ayar.CiktiEkipBasligi, 70.0), (k.CiktiGerekceKolonu, ayar.CiktiGerekceBasligi, 80.0) };
        if (analiz.Prosedurler is not null) kolonlar.Add((k.CiktiProsedurKolonu, ayar.CiktiProsedurBasligi, 70.0));
        foreach (var (kolon, baslikMetni, genislik) in kolonlar)
        {
            var hucre = ws.Cell(k.BaslikSatiri, kolon);
            if (kolon > 1) hucre.Style = ws.Cell(k.BaslikSatiri, kolon - 1).Style;
            hucre.Value = baslikMetni;
            hucre.Style.Font.Bold = true;
            hucre.Style.Fill.BackgroundColor = XLColor.FromHtml("#DDEBF7");
            hucre.Style.Alignment.WrapText = true;
            ws.Column(kolon).Width = genislik;
        }

        // Bir hücre birden fazla maddeye bölündüyse sonuçlar aynı hücreye [1/3], [2/3]... başlıklarıyla yazılır
        foreach (var grup in k.Satirlar.GroupBy(s => s.SatirNo))
        {
            var liste = grup.OrderBy(s => s.AltNo).ToList();
            string Birlestir(Func<KontrolSatiri, string> f) => liste.Count == 1 ? f(liste[0])
                : string.Join("\n\n", liste.Select(s => $"[{s.AltNo}/{s.AltToplam}] {Metin.Kisalt(s.Aksiyon, 90)}\n{f(s)}"));
            var degerler = new List<(int, string)> { (k.CiktiEkipKolonu, Birlestir(s => s.EkipMetni())), (k.CiktiGerekceKolonu, Birlestir(s => s.GerekceMetni())) };
            if (analiz.Prosedurler is not null) degerler.Add((k.CiktiProsedurKolonu, Birlestir(s => s.ProsedurMetni())));
            foreach (var (kolon, deger) in degerler)
            {
                var hucre = ws.Cell(grup.Key, kolon);
                hucre.Value = deger;
                hucre.Style.Alignment.WrapText = true;
                hucre.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            }
        }

        using var cikti = new MemoryStream();
        wb.SaveAs(cikti);
        return cikti.ToArray();
    }
}

// =====================================================================
// PROMPT
// =====================================================================

static class PromptOlusturucu
{
    public const string Sistem =
        "Sen bir bankanın BT organizasyonunu, servis kataloğunu ve yapay zeka yönetişimini (NIST AI RMF, ISO/IEC 42001, BDDK düzenlemeleri) " +
        "iyi bilen kıdemli bir süreç ve yönetişim analistisin. Görevin, yapay zeka geliştirme sürecine ait aksiyon/kontrol maddelerini " +
        "veya kullanıcının anlattığı işleri kurumun servis kataloğuna bakarak işi fiilen yapması gereken ekiplere atamaktır. Yalnızca istenen satır formatında, Türkçe yanıt verirsin.";

    static string GenelEkipKurali(Katalog katalog, YzAyarlari ayar)
    {
        var ekipler = ayar.GenelEkipler().Select(katalog.EkipBul).OfType<Ekip>().Distinct().Select(e => e.Etiket).ToList();
        return ekipler.Count == 0 ? "" :
            $"\n    {string.Join(", ", ekipler)} ekibini YALNIZCA işin kendisi model geliştirme, model eğitimi, YZ uygulamasının kodlanması veya işletilmesi ise Sorumlu yap; " +
            "diğer işlerde gerekiyorsa en fazla Destek olarak ekle.";
    }

    static string KurumKurallari(Katalog katalog, YzAyarlari ayar)
    {
        var kurallar = ayar.Kurallar().Where(k => !string.IsNullOrWhiteSpace(k))
            .Select(k => Regex.Replace(k, @"\{([^{}]+)\}", m => katalog.EkipBul(m.Groups[1].Value)?.Etiket ?? m.Groups[1].Value))
            .ToList();
        return kurallar.Count == 0 ? "" : "\n1b. KURUM KURALLARI (bu kurallar katalogdaki servis adlarından ÖNCELİKLİDİR):" +
            string.Concat(kurallar.Select(k => "\n    - " + k));
    }

    public static string Olustur(KontrolSatiri s, Katalog katalog, List<Ekip> adaylar, YzAyarlari ayar,
        List<Prosedur>? ilgiliProsedurler = null, ProsedurKutuphanesi? kutuphane = null)
    {
        var sb = new StringBuilder();
        var q = katalog.Sorgu(s);
        bool tumu = adaylar.Count >= katalog.Ekipler.Count;
        // Ön eleme kapalıysa servis kırpılmaz; açıkken 0 da "tüm servisler" demektir
        int servisSiniri = !ayar.OnElemeYap || ayar.EkipBasinaServisSayisi <= 0 ? int.MaxValue : ayar.EkipBasinaServisSayisi;

        if (tumu)
        {
            // Katalog her satırda aynı sırayla en başta gider; sunucu önbelleği (prefix cache) varsa hızlanır
            sb.AppendLine($"KURUMDAKİ TÜM EKİPLER ({katalog.Ekipler.Count} adet) VE SAHİBİ OLDUKLARI SERVİSLER. EKİPLERİ YALNIZCA ### BAŞLIKLARINDAN, KODUYLA BİRLİKTE SEÇ:");
        }
        else
        {
            sb.AppendLine($"KURUMDAKİ TÜM EKİPLER ({katalog.Ekipler.Count} adet). EKİPLERİ YALNIZCA BU LİSTEDEN, KODUYLA BİRLİKTE SEÇ:");
            foreach (var e in katalog.Ekipler) sb.AppendLine($"- {e.Etiket}");
            sb.AppendLine();
            sb.AppendLine($"EN OLASI {adaylar.Count} ADAY EKİBİN SAHİBİ OLDUĞU SERVİSLER (anahtar kelime benzerliğiyle otomatik seçildi; doğru ekip burada yoksa yukarıdaki tüm ekip listesinden seç):");
        }
        foreach (var e in adaylar)
        {
            sb.AppendLine($"### {e.Etiket}");
            var servisler = tumu ? e.Servisler : katalog.IlgiliServisler(e, q);
            foreach (var sv in servisler.Take(servisSiniri))
                sb.AppendLine(ayar.ServisAciklamasiGonder && sv.Aciklama.Length > 0
                    ? $"  - {sv.Ad}: {(ayar.ServisAciklamaMaksKarakter > 0 ? Metin.Kisalt(sv.Aciklama, ayar.ServisAciklamaMaksKarakter) : sv.Aciklama)}"
                    : $"  - {sv.Ad}");
            if (e.Servisler.Count > servisSiniri)
                sb.AppendLine($"  - ... (+{e.Servisler.Count - servisSiniri} servis daha)");
        }
        sb.AppendLine();

        if (s.Serbest)
        {
            sb.AppendLine("KULLANICININ ANLATTIĞI, YAPILMASI GEREKEN İŞ:");
            sb.AppendLine(s.Aksiyon);
        }
        else
        {
            sb.AppendLine("DEĞERLENDİRİLECEK MADDE:");
            if (s.AltKategori.Length > 0) sb.AppendLine($"- Alt Kategori: {s.AltKategori}");
            if (s.AltToplam > 1)
            {
                sb.AppendLine($"- AKSİYON / KONTROL (hücredeki {s.AltToplam} ayrı maddeden {s.AltNo}. madde; YALNIZCA BU MADDE için ekip seç): {s.Aksiyon}");
                sb.AppendLine($"- Hücrenin tamamı (yalnızca bağlam için): {s.HucreMetni.ReplaceLineEndings(" ")}");
            }
            else sb.AppendLine($"- AKSİYON / KONTROL: {s.Aksiyon}");
        }
        foreach (var a in s.Alanlar) sb.AppendLine($"- {a.Key}: {a.Value}");
        sb.AppendLine();

        bool prosedurVar = ilgiliProsedurler is { Count: > 0 } && kutuphane is not null;
        if (prosedurVar)
        {
            sb.AppendLine("KURUMDA YÜRÜRLÜKTE OLAN, BU İŞLE İLGİLİ OLABİLECEK PROSEDÜRLER (anahtar kelime benzerliğiyle seçildi; hepsi ilgili olmayabilir):");
            foreach (var p in ilgiliProsedurler!)
            {
                var ekip = kutuphane!.Ekibi(p);
                var sahip = ekip is not null ? ekip.Etiket : "(süreç sahibi birim katalogda yok, atanamaz)";
                sb.AppendLine($"- {p.Kimlik} | {p.Ad}{(p.SurecModeli.Length > 0 && p.SurecModeli != p.Ad ? $" (Süreç: {p.SurecModeli})" : "")} | Süreç sahibi: {sahip}");
            }
            sb.AppendLine();
        }

        if (s.KullaniciEkip.Length > 0 || s.KullaniciNot.Length > 0)
        {
            sb.AppendLine("KULLANICININ ÖN GÖRÜŞÜ (yarım kalmış bir ön çalışmadır; doğruluğu garanti değildir, YALNIZCA İPUCU olarak kullan):");
            if (s.KullaniciEkip.Length > 0) sb.AppendLine($"- Kullanıcının düşündüğü ekipler: {s.KullaniciEkip}");
            if (s.KullaniciNot.Length > 0) sb.AppendLine($"- Mevcut süreçte var mı / nerede: {s.KullaniciNot}");
            sb.AppendLine();
        }

        sb.AppendLine($@"KURALLAR:
1. Bu maddeyi hayata geçirmek, uygulamak ve sürdürmek için işi FİİLEN yapması gereken ekip(ler)i seç. Sadece adı benzeyen ekibi değil, servis kataloğundaki görev alanı bu işe uyan ekibi seç.
1a. İSİM BENZERLİĞİ TUZAĞI (ÇOK KRİTİK): Bu maddelerin TAMAMI zaten yapay zeka geliştirme sürecine aittir. Bu yüzden maddede ""yapay zeka"", ""YZ"", ""model"" geçmesi veya ekip adında ""Yapay Zeka"" bulunması, işi o ekibe vermek için TEK BAŞINA SEBEP DEĞİLDİR.
    Ekibi işin TÜRÜNE göre seç: ekip görevlendirme/organizasyon/süreç tanımlama, proje planlama ve takip, bilgi güvenliği ve sızma testi, kişisel veri/KVKK, hukuk/uyum, risk, altyapı/sunucu, değişiklik ve sürüm yönetimi, eğitim, satın alma gibi işler bu görev alanlarına sahip ekiplere aittir.{GenelEkipKurali(katalog, ayar)}{KurumKurallari(katalog, ayar)}
2. YALNIZCA yukarıdaki listede [E...] koduyla verilen ekipleri seçebilirsin. Her ekibi KODU VE ADIYLA, listede yazıldığı gibi AYNEN yaz (örn: [E001] ...).
   Listede olmayan birim, müdürlük, komite, ofis veya ekip adı UYDURMA; kendi bildiğin genel birim adlarını (""Risk Yönetimi"", ""YZ Ofisi"", ""Hukuk"" vb.) listede yoksa ASLA yazma. Listede olmayan ad yazılan yanıt GEÇERSİZ sayılır.
   İşi birebir karşılayan ekip yoksa görev alanı en yakın olan listedeki ekibi seç ve GUVEN değerini düşür.
3. En az 1, en fazla {ayar.MaksEkipSayisi} ekip seç. İlk ekibin rolü ""Sorumlu"" (işin sahibi), diğerlerinin rolü ""Destek"" olsun. Gereksiz yere ekip ekleme.
4. Her ekip için bu maddede somut olarak NE YAPACAĞINI 1-2 cümleyle yaz (örn: ""Atama yazısını hazırlar ve süreç sahibini belirler"").
5. Kullanıcının ön görüşüne körü körüne uyma. Servis kataloğuyla uyumluysa destekleyici kanıt olarak kullan; uyumsuzsa gerekçede nedenini kısaca belirt.
6. Gerekçede kararı hangi servis(ler)e veya ekibin hangi görev alanına dayanarak verdiğini açıkça yaz (2-4 cümle).
7. GUVEN 0-100 arası bir sayıdır. Katalogda bu işi açıkça kapsayan bir servis yoksa 60'ın altında ver.{(prosedurVar ? @"
7a. PROSEDÜRLER: Yukarıdaki prosedür listesinden bu işi bugün zaten kapsayan veya doğrudan ilişkili olanları PROSEDUR satırlarıyla belirt (en fazla 5). Yalnızca listedeki doküman kodlarını kullan, kod UYDURMA. Hiçbiri gerçekten ilgili değilse tek satır ""PROSEDUR: YOK"" yaz.
    Bir işi zaten yürüten prosedürün süreç sahibi birimi, o işin sorumlu ekibi olması için güçlü bir kanıttır; ama ekip yine de yalnızca [E...] listesinden seçilir." : "")}
8. ÇIKTI FORMATI: Yanıtına KESİNLİKLE aşağıdaki satırlarla başla. JSON, tablo, markdown veya başka açıklama KULLANMA. Her ekip için ayrı bir EKIP satırı yaz; alanları | ile ayır.
EKIP: [E...] <listedeki ekip adı> | Sorumlu | <bu ekibin bu maddede ne yapacağı>
EKIP: [E...] <listedeki ekip adı> | Destek | <bu ekibin bu maddede ne yapacağı>
GUVEN: %<0-100>
GEREKCE: <kararı nasıl verdiğin, 2-4 cümle>{(prosedurVar ? @"
PROSEDUR: <listedeki doküman kodu> | <bu işle ilişkisi, 1 cümle>" : "")}");

        return sb.ToString();
    }
}

// =====================================================================
// MODEL YANITI AYRIŞTIRMA
// =====================================================================

static class CevapAyristirici
{
    static readonly JsonDocumentOptions Secenek = new() { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip };

    static string? Str(JsonNode? n, params string[] anahtarlar)
    {
        if (n is not JsonObject o) return null;
        foreach (var a in anahtarlar)
        {
            var bulunan = o.FirstOrDefault(kv => string.Equals(Metin.Normalize(kv.Key), Metin.Normalize(a), StringComparison.Ordinal)).Value;
            if (bulunan is JsonValue v) return v.ToString().Trim();
        }
        return null;
    }

    /// Düşünme bloklarını ve markdown süslemelerini temizler.
    static readonly Regex ProsedurSatiri = new(@"^\s*(?:[-*•#>]+\s*)?PROSED[ÜU]R\s*[:：=]\s*(?<deger>.*)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// PROSEDUR: kod | ilişki satırlarını okur; yalnızca kütüphanede olan kodları döner.
    public static List<ProsedurOnerisi> Prosedurler(string ham, ProsedurKutuphanesi? kutuphane)
    {
        var sonuc = new List<ProsedurOnerisi>();
        if (kutuphane is null) return sonuc;
        foreach (var satir in Temizle(ham).Split('\n'))
        {
            var m = ProsedurSatiri.Match(satir.TrimEnd('\r'));
            if (!m.Success) continue;
            var parcalar = m.Groups["deger"].Value.Split('|', 2).Select(x => x.Trim().Trim('[', ']', '"', '\'')).ToArray();
            // "PRS.KDY.10 RHEL ..." gibi kod + ad birlikte yazılmış olabilir: ilk kelimeyi de dene
            var p = kutuphane.Bul(parcalar[0]) ?? kutuphane.Bul(parcalar[0].Split(' ', 2)[0]);
            if (p is null || sonuc.Any(x => x.Kod == p.Kimlik)) continue;
            var ekip = kutuphane.Ekibi(p);
            sonuc.Add(new ProsedurOnerisi(p.Kimlik, p.Ad, ekip?.Ad ?? p.Birim, p.Link, parcalar.Length > 1 ? parcalar[1] : ""));
        }
        return sonuc;
    }

    public static string Temizle(string? ham)
    {
        var metin = ham ?? "";
        metin = Regex.Replace(metin, @"<think>.*?</think>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        // Açılışı sunucu tarafından yutulmuş ya da hiç kapanmamış düşünme bloğu
        var kapanis = metin.IndexOf("</think>", StringComparison.OrdinalIgnoreCase);
        if (kapanis >= 0) metin = metin[(kapanis + 8)..];
        var acilis = metin.IndexOf("<think>", StringComparison.OrdinalIgnoreCase);
        if (acilis >= 0) metin = metin[..acilis];
        metin = Regex.Replace(metin, @"```[a-zA-Z]*", "");
        return metin.Replace("**", "").Trim();
    }

    public static (List<EkipOnerisi> ekipler, string gerekce, int? guven)? Ayristir(string ham, Katalog katalog)
    {
        var metin = Temizle(ham);
        return SatirlardanAyristir(metin, katalog) ?? JsondanAyristir(metin, katalog);
    }

    static readonly Regex AnahtarSatiri = new(@"^\s*(?:[-*•#>]+\s*)?(?<anahtar>EK[İI]P(?:\s*\d+)?|G[ÜU]VEN|GEREK[ÇC]E|ANAL[İI]Z)\s*[:：=]\s*(?<deger>.*)$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// EKIP: ad | rol | görev / GUVEN: %80 / GEREKCE: ... biçimini okur.
    static (List<EkipOnerisi>, string, int?)? SatirlardanAyristir(string metin, Katalog katalog)
    {
        var ekipler = new List<EkipOnerisi>();
        var gerekce = new StringBuilder();
        int? guven = null;
        string? sonAnahtar = null;

        foreach (var satir in metin.Split('\n'))
        {
            if (ProsedurSatiri.IsMatch(satir)) { sonAnahtar = "PROSEDUR"; continue; }
            var m = AnahtarSatiri.Match(satir.TrimEnd('\r'));
            if (!m.Success)
            {
                // GEREKCE birden fazla satıra yayılabilir
                if (sonAnahtar is "GEREKCE" && satir.Trim().Length > 0) gerekce.Append(' ').Append(satir.Trim());
                continue;
            }
            var anahtar = Metin.Normalize(m.Groups["anahtar"].Value).ToUpperInvariant();
            var deger = m.Groups["deger"].Value.Trim();
            if (anahtar.StartsWith("EKIP"))
            {
                sonAnahtar = "EKIP";
                var parcalar = deger.Split('|', 3).Select(p => p.Trim().Trim('[', ']', '<', '>', '"', '\'')).ToArray();
                if (parcalar[0].Length == 0) continue;
                var rol = parcalar.Length > 1 && parcalar[1].Length > 0 ? parcalar[1] : (ekipler.Count == 0 ? "Sorumlu" : "Destek");
                var gorev = parcalar.Length > 2 ? parcalar[2] : "";
                // Model rolü yazmayıp ikinci alana doğrudan görevi yazmış olabilir
                if (parcalar.Length == 2 && Metin.Normalize(rol) is not ("sorumlu" or "destek")) { gorev = rol; rol = ekipler.Count == 0 ? "Sorumlu" : "Destek"; }
                EkipEkle(ekipler, katalog, parcalar[0], rol, gorev);
            }
            else if (anahtar.StartsWith("GUVEN"))
            {
                sonAnahtar = "GUVEN";
                if (Regex.Match(deger, @"\d+") is { Success: true } g) guven = Math.Clamp(int.Parse(g.Value), 0, 100);
            }
            else
            {
                sonAnahtar = "GEREKCE";
                gerekce.Append(gerekce.Length > 0 ? " " : "").Append(deger);
            }
        }
        return ekipler.Count == 0 ? null : (ekipler, gerekce.ToString().Trim(), guven);
    }

    static void EkipEkle(List<EkipOnerisi> ekipler, Katalog katalog, string ad, string rol, string gorev)
    {
        var eslesen = katalog.EkipBul(ad);
        var nihaiAd = eslesen?.Ad ?? ad.Trim();
        if (ekipler.Any(e => e.Ekip == nihaiAd)) return;
        var normRol = Metin.Normalize(rol);
        rol = normRol.StartsWith("sorumlu") ? "Sorumlu" : normRol.StartsWith("destek") ? "Destek" : rol;
        ekipler.Add(new EkipOnerisi(nihaiAd, rol, gorev, eslesen is not null));
    }

    /// Yedek: model yine de JSON döndürdüyse onu okur.
    static (List<EkipOnerisi>, string, int?)? JsondanAyristir(string metin, Katalog katalog)
    {
        int bas = metin.IndexOf('{'), son = metin.LastIndexOf('}');
        if (bas < 0 || son <= bas) return null;

        JsonNode? kok;
        try { kok = JsonNode.Parse(metin[bas..(son + 1)], documentOptions: Secenek); }
        catch (JsonException) { return null; }
        if (kok is not JsonObject obj) return null;

        var dizi = obj.FirstOrDefault(kv => Metin.Normalize(kv.Key) is "ekipler" or "ekip" or "teams").Value;
        var ekipler = new List<EkipOnerisi>();
        foreach (var item in dizi is JsonArray a ? a.ToList() : dizi is null ? [] : [dizi])
        {
            var ad = item is JsonValue v ? v.ToString() : Str(item, "ekip", "ekip_adi", "ad", "team");
            if (string.IsNullOrWhiteSpace(ad)) continue;
            var rol = Str(item, "rol", "role") is { Length: > 0 } r ? r : (ekipler.Count == 0 ? "Sorumlu" : "Destek");
            EkipEkle(ekipler, katalog, ad, rol, Str(item, "gorev", "görev", "gorevi", "task", "aciklama") ?? "");
        }

        var gerekce = Str(obj, "gerekce", "gerekçe", "analiz", "reason") ?? "";
        int? guven = null;
        var guvenMetni = Str(obj, "guven", "güven", "confidence");
        if (guvenMetni is not null && Regex.Match(guvenMetni, @"\d+") is { Success: true } m)
            guven = Math.Clamp(int.Parse(m.Value), 0, 100);

        return ekipler.Count == 0 ? null : (ekipler, gerekce, guven);
    }
}

// =====================================================================
// İŞ YÖNETİMİ (arka planda paralel model çağrıları)
// =====================================================================

class IsYoneticisi
{
    readonly YzAyarlari ayar;
    readonly ExcelAyarlari excelAyar;
    readonly IHttpClientFactory httpFactory;
    readonly ILogger<IsYoneticisi> log;
    readonly ConcurrentDictionary<string, AnalizIsi> isler = new();

    // Son yüklenen katalog diske de yazılır; uygulama yeniden başlasa da tekrar yüklemek gerekmez
    readonly string katalogYolu, katalogAdYolu;
    Katalog? katalog;
    string katalogAdi = "";
    DateTime katalogZamani;

    readonly string prosedurYolu, prosedurAdYolu;
    List<Prosedur>? prosedurler;
    string prosedurAdi = "";
    DateTime prosedurZamani;
    ProsedurKutuphanesi? kutuphane;   // katalog veya prosedür listesi değişince yeniden kurulur

    public IsYoneticisi(YzAyarlari ayar, ExcelAyarlari excelAyar, IHttpClientFactory httpFactory, ILogger<IsYoneticisi> log, IWebHostEnvironment env)
    {
        (this.ayar, this.excelAyar, this.httpFactory, this.log) = (ayar, excelAyar, httpFactory, log);
        var klasor = Path.Combine(env.ContentRootPath, "veri");
        katalogYolu = Path.Combine(klasor, "son_katalog.xlsx");
        katalogAdYolu = Path.Combine(klasor, "son_katalog_adi.txt");
        prosedurYolu = Path.Combine(klasor, "son_prosedur.xlsx");
        prosedurAdYolu = Path.Combine(klasor, "son_prosedur_adi.txt");
        if (File.Exists(prosedurYolu))
        {
            try
            {
                prosedurler = ExcelIslemleri.ProsedurOku(File.ReadAllBytes(prosedurYolu), excelAyar);
                prosedurAdi = File.Exists(prosedurAdYolu) ? File.ReadAllText(prosedurAdYolu).Trim() : "son_prosedur.xlsx";
                prosedurZamani = File.GetLastWriteTime(prosedurYolu);
            }
            catch (Exception ex) { log.LogWarning("Kayıtlı prosedür listesi okunamadı: {Hata}", ex.Message); }
        }
        if (File.Exists(katalogYolu))
        {
            try
            {
                katalog = ExcelIslemleri.KatalogOku(File.ReadAllBytes(katalogYolu), excelAyar);
                katalogAdi = File.Exists(katalogAdYolu) ? File.ReadAllText(katalogAdYolu).Trim() : "son_katalog.xlsx";
                katalogZamani = File.GetLastWriteTime(katalogYolu);
            }
            catch (Exception ex) { log.LogWarning("Kayıtlı katalog okunamadı: {Hata}", ex.Message); }
        }
    }

    public object KatalogBilgisi() => katalog is null
        ? new { var = false }
        : new
        {
            var = true,
            dosyaAdi = katalogAdi,
            yuklenme = katalogZamani,
            ekipSayisi = katalog.Ekipler.Count,
            servisSayisi = katalog.Ekipler.Sum(e => e.Servisler.Count),
            aciklamaliServis = katalog.Ekipler.Sum(e => e.Servisler.Count(sv => sv.Aciklama.Length > 0 && sv.Kod != "EK")),
            ekTanimliEkipler = katalog.Ekipler.Where(e => e.Servisler.Any(sv => sv.Kod == "EK")).Select(e => e.Ad)
        };

    public void KatalogYukle(byte[] bytes, string dosyaAdi)
    {
        var yeni = ExcelIslemleri.KatalogOku(bytes, excelAyar);
        (katalog, katalogAdi, katalogZamani, kutuphane) = (yeni, dosyaAdi, DateTime.Now, null);
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(katalogYolu)!);
            File.WriteAllBytes(katalogYolu, bytes);
            File.WriteAllText(katalogAdYolu, dosyaAdi);
        }
        catch (Exception ex) { log.LogWarning("Katalog diske yazılamadı: {Hata}", ex.Message); }
    }

    public object ProsedurBilgisi()
    {
        if (prosedurler is null) return new { var = false };
        var k = AktifKutuphane();
        return new
        {
            var = true,
            dosyaAdi = prosedurAdi,
            yuklenme = prosedurZamani,
            prosedurSayisi = prosedurler.Count,
            birimSayisi = prosedurler.Select(p => p.Birim).Distinct().Count(),
            eslesmeyenBirim = k?.EslesmeyenBirimSayisi
        };
    }

    public void ProsedurYukle(byte[] bytes, string dosyaAdi)
    {
        var yeni = ExcelIslemleri.ProsedurOku(bytes, excelAyar);
        (prosedurler, prosedurAdi, prosedurZamani, kutuphane) = (yeni, dosyaAdi, DateTime.Now, null);
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(prosedurYolu)!);
            File.WriteAllBytes(prosedurYolu, bytes);
            File.WriteAllText(prosedurAdYolu, dosyaAdi);
        }
        catch (Exception ex) { log.LogWarning("Prosedür listesi diske yazılamadı: {Hata}", ex.Message); }
    }

    public void ProsedurKaldir()
    {
        (prosedurler, kutuphane) = (null, null);
        try { File.Delete(prosedurYolu); File.Delete(prosedurAdYolu); } catch { }
    }

    ProsedurKutuphanesi? AktifKutuphane()
    {
        if (prosedurler is null || katalog is null) return null;
        return kutuphane ??= new ProsedurKutuphanesi(prosedurler, katalog);
    }

    Katalog AktifKatalog() => katalog ?? throw new ExcelOkumaHatasi("Önce servis kataloğu Excel dosyasını yükleyin.");

    public AnalizIsi ExcelIleBaslat(byte[] kontrolBytes, string dosyaAdi) => Baslat(new AnalizIsi
    {
        DosyaAdi = dosyaAdi,
        KontrolBytes = kontrolBytes,
        Kontrol = ExcelIslemleri.KontrolOku(kontrolBytes, excelAyar),
        Katalog = AktifKatalog(),
        Prosedurler = ayar.IlgiliProsedurSayisi > 0 ? AktifKutuphane() : null
    });

    public AnalizIsi MetinleBaslat(string metin, bool herSatirAyriMadde)
    {
        var maddeler = (herSatirAyriMadde
                ? metin.Split('\n').Select(m => Regex.Replace(m, @"^\s*(\d+[.)-]|[-*•])\s*", "").Trim())
                : [metin.Trim()])
            .Where(m => m.Length > 0).ToList();
        if (maddeler.Count == 0) throw new ExcelOkumaHatasi("Yapılacak işi anlatan bir metin yazın.");

        var satirlar = maddeler.Select((m, i) => new KontrolSatiri { Sira = i + 1, SatirNo = i + 1, Aksiyon = m, Serbest = true }).ToList();
        return Baslat(new AnalizIsi
        {
            DosyaAdi = maddeler.Count == 1 ? "Metinle soru" : $"Metinle soru ({maddeler.Count} madde)",
            Kontrol = new KontrolDosyasi { SayfaAdi = "Sonuç", Satirlar = satirlar },
            Katalog = AktifKatalog(),
            Prosedurler = ayar.IlgiliProsedurSayisi > 0 ? AktifKutuphane() : null
        });
    }

    AnalizIsi Baslat(AnalizIsi analiz)
    {
        foreach (var eski in isler.Values.Where(i => i.Baslangic < DateTime.Now.AddHours(-12)))
            isler.TryRemove(eski.Id, out _);
        isler[analiz.Id] = analiz;
        log.LogInformation("İş {Id} başladı: {Satir} madde, {Ekip} ekip", analiz.Id, analiz.Kontrol.Satirlar.Count, analiz.Katalog.Ekipler.Count);
        _ = Task.Run(() => CalistirAsync(analiz));
        return analiz;
    }

    public AnalizIsi? Getir(string id) => isler.GetValueOrDefault(id);

    public bool YenidenCalistir(string id, int sira)
    {
        if (Getir(id) is not { } analiz || analiz.Kontrol.Satirlar.FirstOrDefault(s => s.Sira == sira) is not { } satir)
            return false;
        if (satir.Durum == "İşleniyor") return true;
        satir.Durum = "Bekliyor";
        _ = Task.Run(() => SatirIsleAsync(analiz, satir, CancellationToken.None));
        return true;
    }

    async Task CalistirAsync(AnalizIsi analiz)
    {
        try
        {
            await Parallel.ForEachAsync(analiz.Kontrol.Satirlar,
                new ParallelOptions { MaxDegreeOfParallelism = Math.Max(1, ayar.Paralellik), CancellationToken = analiz.Iptal.Token },
                async (s, ct) => await SatirIsleAsync(analiz, s, ct));
            analiz.GenelDurum = "Tamamlandı";
        }
        catch (OperationCanceledException)
        {
            foreach (var s in analiz.Kontrol.Satirlar.Where(s => s.Durum is "Bekliyor" or "İşleniyor")) s.Durum = "İptal";
            analiz.GenelDurum = "İptal edildi";
        }
        catch (Exception ex)
        {
            log.LogError(ex, "İş {Id} beklenmeyen hata", analiz.Id);
            analiz.GenelDurum = "Hata";
        }
    }

    async Task SatirIsleAsync(AnalizIsi analiz, KontrolSatiri s, CancellationToken ct)
    {
        s.Durum = "İşleniyor";
        s.Hata = "";
        var katalog = analiz.Katalog;
        var adaylar = ayar.OnElemeYap ? katalog.AdaylariSec(s, ayar.AdayEkipSayisi) : katalog.Ekipler;
        s.Adaylar = adaylar.Select(a => a.Ad).ToList();

        var ilgili = analiz.Prosedurler?.IlgilileriSec(katalog.Sorgu(s), ayar.IlgiliProsedurSayisi) ?? [];
        s.GonderilenProsedurler = ilgili.Select(p => $"{p.Kimlik} {p.Ad}").ToList();
        var kullaniciMesaji = PromptOlusturucu.Olustur(s, katalog, adaylar, ayar, ilgili, analiz.Prosedurler);
        // Qwen3'ün "yumuşak" düşünme kapatma komutu; chat_template_kwargs'ı tanımayan sunucularda da işe yarar
        if (ayar.DusunmeyiKapat) kullaniciMesaji += "\n/no_think";
        s.PromptKarakter = kullaniciMesaji.Length + PromptOlusturucu.Sistem.Length;
        var mesajlar = new List<Mesaj> { new("system", PromptOlusturucu.Sistem), new("user", kullaniciMesaji) };
        string sonHata = "";

        for (int deneme = 1; deneme <= Math.Max(1, ayar.DenemeSayisi); deneme++)
        {
            try
            {
                var (ham, bitisNedeni) = await ModeliCagirAsync(mesajlar, ct);
                s.HamCevap = ham;
                if (CevapAyristirici.Ayristir(ham, katalog) is { } karar)
                {
                    var uydurma = karar.ekipler.Where(e => !e.KatalogdaVar).Select(e => e.Ekip).ToList();
                    bool sonDeneme = deneme >= Math.Max(1, ayar.DenemeSayisi);
                    if (uydurma.Count > 0 && !sonDeneme)
                    {
                        // Katalogda olmayan birim adı: kabul etme, modelden düzeltmesini iste
                        log.LogWarning("Satır {Satir} deneme {Deneme}: katalogda olmayan ekip(ler): {Ekipler}", s.SatirNo, deneme, string.Join(", ", uydurma));
                        sonHata = $"Model katalogda olmayan ekip önerdi: {string.Join(", ", uydurma)}";
                        mesajlar.Add(new("assistant", CevapAyristirici.Temizle(ham)));
                        mesajlar.Add(new("user",
                            $"Şu ekipler katalogda YOK, bunları kullanamazsın: {string.Join(", ", uydurma.Select(u => $"\"{u}\""))}. " +
                            "Bunların yerine YALNIZCA listede [E...] koduyla verilen ekiplerden görev alanı en uygun olanı seç. " +
                            "Yanıtını aynı biçimde (EKIP: [E...] ad | rol | görev, GUVEN, GEREKCE) baştan yaz."));
                        continue;
                    }

                    var gecerli = karar.ekipler.Where(e => e.KatalogdaVar).ToList();
                    if (gecerli.Count == 0)
                    {
                        sonHata = $"Model yalnızca katalogda olmayan ekipler önerdi: {string.Join(", ", uydurma)}";
                        break;
                    }
                    // Sorumlu çıkarıldıysa ilk kalan ekip sorumlu olur
                    if (!gecerli.Any(e => e.Rol == "Sorumlu")) gecerli[0] = gecerli[0] with { Rol = "Sorumlu" };
                    if (uydurma.Count > 0)
                        karar.gerekce += $" (Not: katalogda bulunmayan {string.Join(", ", uydurma.Select(u => $"'{u}'"))} önerisi çıkarıldı.)";

                    s.Prosedurler = CevapAyristirici.Prosedurler(ham, analiz.Prosedurler);
                    s.Ekipler = gecerli;
                    s.Gerekce = karar.gerekce.Trim();
                    s.Guven = karar.guven;
                    s.Uyum = UyumHesapla(katalog, s);
                    s.Durum = "Tamam";
                    return;
                }
                var temiz = CevapAyristirici.Temizle(ham);
                sonHata = bitisNedeni == "length"
                    ? $"Yanıt MaksToken ({ayar.MaksToken}) sınırında kesildi; appsettings.json'da MaksToken'ı artırın veya DusunmeyiKapat'ı açın."
                    : temiz.Length == 0
                        ? "Model boş yanıt döndü (yanıtın tamamı düşünme bloğu olabilir; DusunmeyiKapat ayarını deneyin)."
                        : $"Model yanıtı beklenen biçimde değil: {Metin.Kisalt(temiz.ReplaceLineEndings(" "), 200)}";
                log.LogWarning("Satır {Satir} deneme {Deneme}: {Hata}", s.SatirNo, deneme, sonHata);
                if (temiz.Length > 0)
                {
                    mesajlar.Add(new("assistant", temiz));
                    mesajlar.Add(new("user", "Yanıtın istenen biçimde değildi. Hiçbir açıklama eklemeden yalnızca şu satırları yaz:\nEKIP: [E...] <listedeki ekip adı> | Sorumlu | <ne yapacak>\nGUVEN: %<0-100>\nGEREKCE: <kısa gerekçe>"));
                }
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                s.Durum = "İptal";
                throw;
            }
            catch (Exception ex)
            {
                sonHata = ex is TaskCanceledException ? "YZ servisi zaman aşımına uğradı."
                    : ex.InnerException is null ? ex.Message : $"{ex.Message} ({ex.InnerException.Message})";
                log.LogWarning("Satır {Satir} deneme {Deneme} hata: {Hata}", s.SatirNo, deneme, sonHata);
                if (deneme < ayar.DenemeSayisi) await Task.Delay(2000 * deneme, ct);
            }
        }

        s.Hata = sonHata;
        s.Durum = "Hata";
    }

    async Task<(string icerik, string bitisNedeni)> ModeliCagirAsync(List<Mesaj> mesajlar, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(ayar.QwenApiUrl)) throw new InvalidOperationException("appsettings.json: YzAyarlari:QwenApiUrl tanımlı değil.");

        var payload = new Dictionary<string, object>
        {
            ["model"] = ayar.AiModel,
            ["messages"] = mesajlar,
            ["temperature"] = ayar.Temperature
        };
        if (ayar.MaksToken > 0) payload["max_tokens"] = ayar.MaksToken;
        if (ayar.JsonModu) payload["response_format"] = new { type = "json_object" };
        if (ayar.DusunmeyiKapat) payload["chat_template_kwargs"] = new { enable_thinking = false };

        using var req = new HttpRequestMessage(HttpMethod.Post, ayar.QwenApiUrl) { Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json") };
        if (!string.IsNullOrWhiteSpace(ayar.QwenApiKey)) req.Headers.Add("Authorization", $"Bearer {ayar.QwenApiKey}");

        var client = httpFactory.CreateClient("InsecureClient");
        using var res = await client.SendAsync(req, ct);
        var govde = await res.Content.ReadAsStringAsync(ct);
        if (!res.IsSuccessStatusCode)
            throw new HttpRequestException($"YZ servisi HTTP {(int)res.StatusCode}: {Metin.Kisalt(govde, 300)}");

        using var doc = JsonDocument.Parse(govde);
        var secim = doc.RootElement.GetProperty("choices")[0];
        var bitis = secim.TryGetProperty("finish_reason", out var f) && f.ValueKind == JsonValueKind.String ? f.GetString() ?? "" : "";
        var mesaj = secim.GetProperty("message");
        var icerik = mesaj.TryGetProperty("content", out var c) && c.ValueKind == JsonValueKind.String ? c.GetString() ?? "" : "";
        return (icerik, bitis);
    }

    /// Kullanıcının E kolonundaki ekiplerle modelin önerisini karşılaştırır.
    static string UyumHesapla(Katalog katalog, KontrolSatiri s)
    {
        if (s.KullaniciEkip.Length == 0) return "";
        var kullanici = katalog.MetindenEkipler(s.KullaniciEkip).Select(e => e.Ad).ToHashSet();
        var yz = s.Ekipler.Select(e => e.Ekip).ToHashSet();
        if (kullanici.Count == 0) return "Eşleşmedi";
        var ortak = kullanici.Intersect(yz).Count();
        return ortak == 0 ? "Farklı" : kullanici.SetEquals(yz) ? "Aynı" : "Kısmi";
    }
}
