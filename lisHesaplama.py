import pandas as pd

# 1. Dosya Yolu ve Sütun Adlarını Ayarlayın
# -----------------------------------------------------------------------
# Lütfen Excel dosyanızın adını buraya girin.
dosya_adi = "ardahanDevlet.xlsx"

# Tekrar eden test isimlerinin bulunduğu sütunun adını buraya girin.
TEST_ADI_SUTUNU = "S/H/M Adı"

# Toplanması gereken sayıların (test adetlerinin) bulunduğu sütunun adını buraya girin.
TEST_SAYISI_SUTUNU = "Miktar"
# -----------------------------------------------------------------------

try:
    # 2. Excel Dosyasını Oku
    df = pd.read_excel(dosya_adi)
    print("Veri Başarıyla Yüklendi. İlk 5 satır:")
    print(df.head())
    print("-" * 30)

    # 3. Verileri Grupla ve Topla (Pivot Table Mantığı)
    # Bu adım, belirtilen sütuna göre (Test Adı) gruplar ve diğer sütunu (Adet) toplar.
    toplam_testler = df.groupby(TEST_ADI_SUTUNU)[TEST_SAYISI_SUTUNU].sum().reset_index()

    # Sütun adlarını daha anlaşılır hale getirelim
    toplam_testler.columns = [TEST_ADI_SUTUNU, "Toplam " + TEST_SAYISI_SUTUNU]

    # 4. Sonucu Görüntüle
    print("Gruplanmış ve Toplanmış Sonuçlar:")
    print(toplam_testler)
    print("-" * 30)

    # 5. Sonucu Yeni Bir Excel Dosyasına Kaydet
    cikis_dosya_adi = "ardahanDevletLisToplami.xlsx"
    toplam_testler.to_excel(cikis_dosya_adi, index=False)
    print(f"Sonuçlar başarıyla '{cikis_dosya_adi}' adlı dosyaya kaydedildi.")

except FileNotFoundError:
    print(f"HATA: '{dosya_adi}' adlı dosya bulunamadı. Lütfen dosya adını ve yolunu kontrol edin.")
except KeyError as e:
    print(f"HATA: Belirtilen sütun adı bulunamadı: {e}. Lütfen sütun adlarını kontrol edin.")
except Exception as e:
    print(f"Beklenmeyen bir hata oluştu: {e}")