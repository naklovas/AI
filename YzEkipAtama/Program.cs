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
}

class ExcelAyarlari
{
    public string KullaniciEkipKolonu { get; set; } = "E";
    public string KullaniciNotKolonu { get; set; } = "F";
    public string CiktiEkipBasligi { get; set; } = "YZ Önerilen Ekipler ve Görevleri";
    public string CiktiGerekceBasligi { get; set; } = "YZ Karar Gerekçesi";
    public string KatalogAciklamaKolonu { get; set; } = "F";
}

// =====================================================================
// MODELLER
// =====================================================================

record Servis(string Kod, string Ad, string Bolum, string Yonetici, string Kategori, string Aciklama);

record Mesaj(string role, string content);

record EkipOnerisi(string Ekip, string Rol, string Gorev, bool KatalogdaVar);

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

    // İşlem sonucu (arka plan görevleri tarafından güncellenir)
    public volatile string Durum = "Bekliyor";
    public List<EkipOnerisi> Ekipler = [];
    public string Gerekce = "";
    public int? Guven;
    public string Uyum = "";
    public string Hata = "";
    public string HamCevap = "";
    public List<string> Adaylar = [];
    public int PromptKarakter;

    public string EkipMetni() => string.Join("\n", Ekipler.Select(e =>
        $"• {e.Ekip} ({e.Rol}){(e.KatalogdaVar ? "" : " [katalogda bulunamadı]")}: {e.Gorev}"));

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
    public CancellationTokenSource Iptal { get; } = new();
    public volatile string GenelDurum = "Çalışıyor";

    public object Durum() => new
    {
        id = Id,
        tur = KontrolBytes is null ? "Metin" : "Excel",
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
        "gereklidir", "olup", "olacak", "var", "yok", "aksiyon", "kontrol", "the", "and", "for", "of", "to", "in"
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
// EXCEL OKUMA / YAZMA
// =====================================================================

class KontrolDosyasi
{
    public required string SayfaAdi { get; init; }
    public int BaslikSatiri { get; init; }
    public int CiktiEkipKolonu { get; init; }
    public int CiktiGerekceKolonu { get; init; }
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

        var katalog = new Katalog(servisler);
        if (katalog.Ekipler.Count == 0) throw new ExcelOkumaHatasi("Servis kataloğunda hiç ekip (Servis Sahibi Bölüm) bulunamadı.");
        return katalog;
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

        var basliklar = ws.Row(baslik).CellsUsed().ToDictionary(c => c.Address.ColumnNumber, c => Deger(c));
        int sonKolon = Math.Max(basliklar.Keys.DefaultIfEmpty(1).Max(), notKolon);

        // Dosya daha önce bu programdan çıkmışsa çıktı kolonlarını yeniden kullan
        int cikti1 = basliklar.FirstOrDefault(b => Metin.Normalize(b.Value) == cikti1Norm).Key;
        int cikti2 = basliklar.FirstOrDefault(b => Metin.Normalize(b.Value) == cikti2Norm).Key;
        if (cikti1 == 0) cikti1 = sonKolon + 1;
        if (cikti2 == 0) cikti2 = Math.Max(sonKolon, cikti1) + 1;

        var haric = new HashSet<int> { ekipKolon, notKolon, cikti1, cikti2 };
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
            satirlar.Add(new KontrolSatiri
            {
                Sira = satirlar.Count + 1,
                SatirNo = i,
                Aksiyon = aksiyon,
                AltKategori = altKategoriKolon == 0 ? "" : Deger(r.Cell(altKategoriKolon)),
                Alanlar = bilgiKolonlari
                    .Select(k => new KeyValuePair<string, string>(string.IsNullOrEmpty(basliklar[k]) ? $"Kolon {XLHelper.GetColumnLetterFromNumber(k)}" : basliklar[k], Deger(r.Cell(k))))
                    .Where(kv => kv.Value.Length > 0).ToList(),
                KullaniciEkip = Deger(r.Cell(ekipKolon)),
                KullaniciNot = Deger(r.Cell(notKolon))
            });
        }
        if (satirlar.Count == 0) throw new ExcelOkumaHatasi("Kontrol listesinde 'Aksiyon / Kontrol' kolonu dolu hiç satır bulunamadı.");

        return new KontrolDosyasi
        {
            SayfaAdi = ws.Name,
            BaslikSatiri = baslik,
            CiktiEkipKolonu = cikti1,
            CiktiGerekceKolonu = cikti2,
            Satirlar = satirlar
        };
    }

    /// Metinle sorulan işler için: İş / Ekipler / Gerekçe kolonlu yeni bir dosya üretir.
    static byte[] MetinSonucuYaz(AnalizIsi analiz, ExcelAyarlari ayar)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sonuç");
        var basliklar = new[] { "No", "İş", ayar.CiktiEkipBasligi, ayar.CiktiGerekceBasligi };
        for (int i = 0; i < basliklar.Length; i++)
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
            r++;
        }
        ws.Column(1).Width = 6; ws.Column(2).Width = 60; ws.Column(3).Width = 70; ws.Column(4).Width = 80;
        ws.Range(2, 1, Math.Max(2, r - 1), 4).Style.Alignment.WrapText = true;
        ws.Range(2, 1, Math.Max(2, r - 1), 4).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
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

        foreach (var (kolon, baslikMetni, genislik) in new[] { (k.CiktiEkipKolonu, ayar.CiktiEkipBasligi, 70.0), (k.CiktiGerekceKolonu, ayar.CiktiGerekceBasligi, 80.0) })
        {
            var hucre = ws.Cell(k.BaslikSatiri, kolon);
            if (kolon > 1) hucre.Style = ws.Cell(k.BaslikSatiri, kolon - 1).Style;
            hucre.Value = baslikMetni;
            hucre.Style.Font.Bold = true;
            hucre.Style.Fill.BackgroundColor = XLColor.FromHtml("#DDEBF7");
            hucre.Style.Alignment.WrapText = true;
            ws.Column(kolon).Width = genislik;
        }

        foreach (var s in k.Satirlar)
        {
            foreach (var (kolon, deger) in new[] { (k.CiktiEkipKolonu, s.EkipMetni()), (k.CiktiGerekceKolonu, s.GerekceMetni()) })
            {
                var hucre = ws.Cell(s.SatirNo, kolon);
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

    public static string Olustur(KontrolSatiri s, Katalog katalog, List<Ekip> adaylar, YzAyarlari ayar)
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
            sb.AppendLine($"- AKSİYON / KONTROL: {s.Aksiyon}");
        }
        foreach (var a in s.Alanlar) sb.AppendLine($"- {a.Key}: {a.Value}");
        sb.AppendLine();

        if (s.KullaniciEkip.Length > 0 || s.KullaniciNot.Length > 0)
        {
            sb.AppendLine("KULLANICININ ÖN GÖRÜŞÜ (yarım kalmış bir ön çalışmadır; doğruluğu garanti değildir, YALNIZCA İPUCU olarak kullan):");
            if (s.KullaniciEkip.Length > 0) sb.AppendLine($"- Kullanıcının düşündüğü ekipler: {s.KullaniciEkip}");
            if (s.KullaniciNot.Length > 0) sb.AppendLine($"- Mevcut süreçte var mı / nerede: {s.KullaniciNot}");
            sb.AppendLine();
        }

        sb.AppendLine($@"KURALLAR:
1. Bu maddeyi hayata geçirmek, uygulamak ve sürdürmek için işi FİİLEN yapması gereken ekip(ler)i seç. Sadece adı benzeyen ekibi değil, servis kataloğundaki görev alanı bu işe uyan ekibi seç.
2. YALNIZCA yukarıdaki listede [E...] koduyla verilen ekipleri seçebilirsin. Her ekibi KODU VE ADIYLA, listede yazıldığı gibi AYNEN yaz (örn: [E001] ...).
   Listede olmayan birim, müdürlük, komite, ofis veya ekip adı UYDURMA; kendi bildiğin genel birim adlarını (""Risk Yönetimi"", ""YZ Ofisi"", ""Hukuk"" vb.) listede yoksa ASLA yazma. Listede olmayan ad yazılan yanıt GEÇERSİZ sayılır.
   İşi birebir karşılayan ekip yoksa görev alanı en yakın olan listedeki ekibi seç ve GUVEN değerini düşür.
3. En az 1, en fazla {ayar.MaksEkipSayisi} ekip seç. İlk ekibin rolü ""Sorumlu"" (işin sahibi), diğerlerinin rolü ""Destek"" olsun. Gereksiz yere ekip ekleme.
4. Her ekip için bu maddede somut olarak NE YAPACAĞINI 1-2 cümleyle yaz (örn: ""Atama yazısını hazırlar ve süreç sahibini belirler"").
5. Kullanıcının ön görüşüne körü körüne uyma. Servis kataloğuyla uyumluysa destekleyici kanıt olarak kullan; uyumsuzsa gerekçede nedenini kısaca belirt.
6. Gerekçede kararı hangi servis(ler)e veya ekibin hangi görev alanına dayanarak verdiğini açıkça yaz (2-4 cümle).
7. GUVEN 0-100 arası bir sayıdır. Katalogda bu işi açıkça kapsayan bir servis yoksa 60'ın altında ver.
8. ÇIKTI FORMATI: Yanıtına KESİNLİKLE aşağıdaki satırlarla başla. JSON, tablo, markdown veya başka açıklama KULLANMA. Her ekip için ayrı bir EKIP satırı yaz; alanları | ile ayır.
EKIP: [E...] <listedeki ekip adı> | Sorumlu | <bu ekibin bu maddede ne yapacağı>
EKIP: [E...] <listedeki ekip adı> | Destek | <bu ekibin bu maddede ne yapacağı>
GUVEN: %<0-100>
GEREKCE: <kararı nasıl verdiğin, 2-4 cümle>");

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

    public IsYoneticisi(YzAyarlari ayar, ExcelAyarlari excelAyar, IHttpClientFactory httpFactory, ILogger<IsYoneticisi> log, IWebHostEnvironment env)
    {
        (this.ayar, this.excelAyar, this.httpFactory, this.log) = (ayar, excelAyar, httpFactory, log);
        var klasor = Path.Combine(env.ContentRootPath, "veri");
        katalogYolu = Path.Combine(klasor, "son_katalog.xlsx");
        katalogAdYolu = Path.Combine(klasor, "son_katalog_adi.txt");
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
            aciklamaliServis = katalog.Ekipler.Sum(e => e.Servisler.Count(sv => sv.Aciklama.Length > 0))
        };

    public void KatalogYukle(byte[] bytes, string dosyaAdi)
    {
        var yeni = ExcelIslemleri.KatalogOku(bytes, excelAyar);
        (katalog, katalogAdi, katalogZamani) = (yeni, dosyaAdi, DateTime.Now);
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(katalogYolu)!);
            File.WriteAllBytes(katalogYolu, bytes);
            File.WriteAllText(katalogAdYolu, dosyaAdi);
        }
        catch (Exception ex) { log.LogWarning("Katalog diske yazılamadı: {Hata}", ex.Message); }
    }

    Katalog AktifKatalog() => katalog ?? throw new ExcelOkumaHatasi("Önce servis kataloğu Excel dosyasını yükleyin.");

    public AnalizIsi ExcelIleBaslat(byte[] kontrolBytes, string dosyaAdi) => Baslat(new AnalizIsi
    {
        DosyaAdi = dosyaAdi,
        KontrolBytes = kontrolBytes,
        Kontrol = ExcelIslemleri.KontrolOku(kontrolBytes, excelAyar),
        Katalog = AktifKatalog()
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
            Katalog = AktifKatalog()
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

        var kullaniciMesaji = PromptOlusturucu.Olustur(s, katalog, adaylar, ayar);
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
