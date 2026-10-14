using System.Net.Http.Json;

namespace AppRel
{
    public static class AuthHelper
    {
        // Güvenlik anahtarı koda yazılmaz: appsettings.json -> AppSettings:DokuPanelKey
        // (DokuPanel tarafındakiyle birebir aynı olmalı).

        public static async Task<bool> YetkiKontrol(HttpContext context, IConfiguration config)
        {
            string DokuPanelApiUrl = config.GetValue<string>("AppSettings:DokuPanelApiUrl") ?? config.GetValue<string>("DokuPanelApiUrl") ?? "http://localhost/DokuPanel/api/config/get-groups-by-name?siteName=";
            string DokuPanelKey = config.GetValue<string>("AppSettings:DokuPanelKey") ?? "";
            // 1. Uygulama Kimliğini oku (appsettings.json içindeki AppSettings:CurrentSiteName)
            string siteName = config.GetValue<string>("AppSettings:CurrentSiteName") ?? "BilinmeyenSite";

            // Kullanıcı giriş yapmamışsa direkt reddet
            if (context.User?.Identity?.IsAuthenticated != true) return false;

            // 2. Windows Kimliği ile HttpClient oluştur (401.2 hatasını önlemek için)
            var handler = new HttpClientHandler { UseDefaultCredentials = true, ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true };

            using (var client = new HttpClient(handler))
            {
                try
                {
                    // API Key'i Header'a ekle
                    client.DefaultRequestHeaders.Add("X-DokuPanel-Key", DokuPanelKey);

                    // DokuPanel'den yetki listesini iste
                    var response = await client.GetAsync($"{DokuPanelApiUrl}{Uri.EscapeDataString(siteName)}");

                    // HATA DURUMU: API hata dönerse (404, 401, 500 vb.) detayı sakla
                    if (!response.IsSuccessStatusCode)
                    {
                        var errorDetail = await response.Content.ReadAsStringAsync();
                        context.Items["AuthHata"] = $"API Hatası ({response.StatusCode}): {errorDetail}";
                        return false;
                    }

                    // LİSTEYİ AL: ["ZT\Grup1", "visikhan", ...]
                    var izinliListe = await response.Content.ReadFromJsonAsync<List<string>>() ?? new List<string>();

                    if (izinliListe.Count == 0)
                    {
                        context.Items["AuthHata"] = $"DokuPanel'de '{siteName}' için hiç yetki (Grup/Kişi) tanımlanmamış!";
                        return false;
                    }

                    // KULLANICI NORMALİZASYONU (Küçük harf ve domain temizliği)
                    var tamGirenAd = (context.User.Identity.Name ?? "").ToLower().Trim();
                    var kisaGirenAd = tamGirenAd.Split('\\').Last();

                    // 3. EŞLEŞME KONTROLÜ
                    bool yetkili = false;
                    foreach (var item in izinliListe)
                    {
                        var temizDbItem = item.Trim().ToLower();
                        var temizDbItemKisa = temizDbItem.Split('\\').Last();

                        if (context.User.IsInRole(item.Trim()) || // AD Grup Kontrolü
                            temizDbItemKisa == kisaGirenAd ||     // Kısa isim (visikhan)
                            temizDbItem == tamGirenAd)            // Tam isim (ftroot\visikhan)
                        {
                            yetkili = true;
                            break;
                        }
                    }

                    if (!yetkili)
                    {
                        context.Items["AuthHata"] = $"İsminiz listede yok. Giren: {tamGirenAd} | Beklenenler: {string.Join(", ", izinliListe)}";
                    }

                    return yetkili;
                }
                catch (Exception ex)
                {
                    // Bağlantı kopukluğu veya DNS hatası gibi durumlar
                    context.Items["AuthHata"] = "Bağlantı Hatası: " + ex.Message;
                    return false;
                }
            }
        }

        // Yetkisiz Erişim Sayfası (Bootstrap tarzı şık tasarım)
        public static IResult YetkisizErisimSayfasi(string kullanici, string detay)
        {
            string html = $@"
        <!DOCTYPE html>
        <html lang='tr'>
        <head><meta charset='UTF-8'><title>Erişim Engellendi</title>
        <style>
            body {{ background: #f4f7f6; display: flex; align-items: center; justify-content: center; height: 100vh; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; }}
            .card {{ background: white; padding: 2.5rem; border-radius: 15px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); border-top: 6px solid #dc3545; width: 500px; text-align: center; }}
            h2 {{ color: #dc3545; margin-top: 0; }}
            .user-box {{ background: #e9ecef; padding: 10px; border-radius: 8px; margin: 15px 0; font-weight: bold; color: #495057; }}
            .debug-box {{ background: #fff3f3; border: 1px solid #f5c6cb; padding: 12px; border-radius: 8px; font-size: 0.85em; color: #721c24; text-align: left; margin-top: 20px; word-break: break-all; }}
            .footer {{ margin-top: 20px; font-size: 0.8em; color: #6c757d; }}
        </style>
        </head>
        <body>
            <div class='card'>
                <h2>❌ Erişim Reddedildi</h2>
                <p>Bu sayfayı görüntüleme yetkiniz bulunmamaktadır.</p>
                <div class='user-box'>{System.Net.WebUtility.HtmlEncode(kullanici)}</div>

            </div>
        </body></html>";

            return Results.Content(html, "text/html; charset=utf-8");
        }
    }

}
