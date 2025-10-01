import pandas as pd
import datetime
import math
import re

# =========================================================================
# 1. AYARLAR VE EŞLEŞTİRME TABLOSU (DEĞİŞMEDİ)
# =========================================================================

# Tuple Yapısı (11 Elemanlı):
# (0:STOK KODU, 1:STOK ADI, 2:MİKTAR, 3:BİRİM KOD, 4:VADE TARİHİ, 5:KDV ORANI,
#  6:DEPO KOD, 7:OZELALAN1, 8:ID, 9:TEST_PER_KUTU, 10:SET_COUNT_PER_KUTU)
STOK_MAP = {
    # Glukoz: 1500 Test / Kutu | 5 Set / Kutu => (1500/5) = 300 Test/Set HESAPLANACAK
    "Glukoz (Serum/Plazma)": ("3L8222", "CC GLUKOZ (R*5)(20 ML)(1500 TEST)(ABBOTT)", "7500", "TEST", "9.09.2025", "10",
                              "LOST", "5K", "6464", 1500, 5),

    # Üre: 1400 Test / Kutu | 4 Set / Kutu => (1400/4) = 350 Test/Set HESAPLANACAK
    "Üre (Serum/Plazma)": ("4T1220", "CC ÜRE (R1*4+R2*4)(24,8 ML-10 ML)(1400 TEST)(ABBOTT)", "1400", "TEST",
                           "9.09.2025", "10", "LOST", "4K", "6465", 1400, 4),

    # Kutu/Set dönüşümü olmayanlar (None, None olarak kalır)
    "Kreatinin (Serum/Plazma)": ("4S9520", "CC KREATİNİN 2 (3600 TEST)(ABBOTT)", "3600", "TEST", "9.09.2025", "10",
                                 "LOST", "3K", "6466", 3600, 8),
    "Sodyum (Serum/Plazma)": ("2P3250", "CC ICT SAMPLE DILUENT (NA,K,CL) (R*10)(54 ML)(21000 TEST)(ABBOTT)", "21000",
                              "TEST", "9.09.2025", "10", "LOST", "2K", "6467", 21000, 10),
    # ... Diğer ürünler (kısaltıldı)
    "TSH": ("7K6230", "ARC.TSH (4*500 TEST)", "2000", "TEST", "9.09.2025", "10", "LOST",
            "2K", "6485", 2000, 4),
}

# XML Üst Bilgileri
FIS_NO = "Sİ-0489"
CARI_ADI = "ARDAHAN DEVLET HASTANESİ"
FISTAR_SEVKTAR = datetime.datetime(2025, 10, 1)

# Excel'deki İhtiyaç Sütunu Adı
IHTIYAC_SUTUNU = "2,5 AYLIK İHTİYAÇ MİKTARI (TEST)"
TEST_ADI_SUTUNU = "TEST ADI"


# =========================================================================

def excel_tarih_formatla(tarih):
    """Datetime nesnesini depo programının istediği seri tarih formatına çevirir (45930 gibi)"""
    delta = tarih - datetime.datetime(1899, 12, 30)
    return str(delta.days)


def excel_to_xml(excel_file_path, map_data):
    # XLSX dosyasını oku.
    try:
        df = pd.read_excel(excel_file_path, sheet_name=0, header=0)
    except Exception as e:
        print(f"HATA: Excel dosyasını okurken bir sorun oluştu. Hata: {e}")
        return

    df.columns = df.columns.str.strip().str.replace('"', '', regex=False)

    try:
        df_ihtiyac = df[[TEST_ADI_SUTUNU, IHTIYAC_SUTUNU]].copy()
    except KeyError as e:
        print(f"HATA: '{e.args[0]}' başlıklı sütun bulunamadı.")
        return

    # Veri Temizleme ve Yuvarlama
    df_ihtiyac[IHTIYAC_SUTUNU] = df_ihtiyac[IHTIYAC_SUTUNU].astype(str).str.replace(',', '.', regex=False)
    df_ihtiyac[IHTIYAC_SUTUNU] = pd.to_numeric(df_ihtiyac[IHTIYAC_SUTUNU], errors='coerce').fillna(0)
    df_ihtiyac = df_ihtiyac[df_ihtiyac[IHTIYAC_SUTUNU] > 0]
    df_ihtiyac[IHTIYAC_SUTUNU] = df_ihtiyac[IHTIYAC_SUTUNU].apply(lambda x: math.ceil(x))

    xml_satirlar = ""
    sira_no = 1

    for index, row in df_ihtiyac.iterrows():
        test_adi = str(row[TEST_ADI_SUTUNU]).strip()
        ihtiyac_miktari = int(row[IHTIYAC_SUTUNU])

        # **ÖNEMLİ:** Burada kullandığımız 'ihtiyac_miktari' değeri daha sonra XML'deki <MIKTAR> alanı olacak.

        if test_adi in map_data:
            # EŞLEŞTİRME VERİLERİNİ ÇEKME (11 eleman)
            (stok_kod, stok_adi, miktar, birim_kod, vade_tarihi, kdv_orani, depo_kod,
             ozel_alan1, _, test_per_kutu, set_count_per_kutu) = map_data[test_adi]

            # =========================================================================
            # KUTU/SET HESAPLAMA VE OZELALAN1 GÜNCELLEME MANTIĞI (GENEL VE OTOMATİK SET HESAPLI)
            # =========================================================================

            # Eğer ürün için Kutu Testi ve Kutu İçi Set Sayısı tanımlanmışsa bu bloğu çalıştır
            if test_per_kutu is not None and set_count_per_kutu is not None and test_per_kutu > 0 and set_count_per_kutu > 0:

                try:
                    # Dinamik olarak Set Test Miktarını Hesapla
                    TEST_PER_SET = test_per_kutu / set_count_per_kutu
                except ZeroDivisionError:
                    TEST_PER_SET = 0

                if TEST_PER_SET > 0:
                    # 1. Kutu Miktarını Hesapla (Tam Kısım)
                    kutu_miktari = ihtiyac_miktari // test_per_kutu
                    kalan_test = ihtiyac_miktari % test_per_kutu

                    # 2. Kalan Test Miktarını Sete Çevir ve Yukarı Yuvarla
                    set_miktari = math.ceil(kalan_test / TEST_PER_SET) if kalan_test > 0 else 0

                    # 3. YENİ MİKTARI HESAPLA: Yuvarlama sonucu oluşan tam test miktarı
                    # Bu yeni miktar, XML'deki <MIKTAR> etiketine yazılacak.
                    yeni_ihtiyac_miktari = (kutu_miktari * test_per_kutu) + (set_miktari * TEST_PER_SET)

                    # Gerekli düzenlemeyi yap
                    ihtiyac_miktari = int(yeni_ihtiyac_miktari)

                    # 4. Yeni OZELALAN1 değerini '3K+2SET' formatında oluşturma
                    ozel_alan_parts = []
                    if kutu_miktari > 0:
                        ozel_alan_parts.append(f"{kutu_miktari}K")
                    if set_miktari > 0:
                        ozel_alan_parts.append(f"{set_miktari}F")

                    if ozel_alan_parts:
                        ozel_alan1 = " + ".join(ozel_alan_parts)
                    else:
                        # Eğer her iki miktar da 0'sa (teorik olarak olmamalı) orijinali koru.
                        ozel_alan1 = map_data[test_adi][7]
                else:
                    # Eğer TEST_PER_SET hesaplanamazsa orijinal değerleri koru
                    ihtiyac_miktari = int(row[IHTIYAC_SUTUNU])  # Orijinal yuvarlanmış ihtiyacı geri al
                    ozel_alan1 = map_data[test_adi][7]

            # Diğer ürünler için (kutu/set bilgisi girilmeyenler) orijinal ozel_alan1 ve ihtiyac_miktari değeri korunur.

            # =========================================================================
            # XML Satırını Oluşturma
            # =========================================================================

            satir_xml = f"""
		<Satir>
			<SIRANO>{sira_no}</SIRANO>
			<KARTTIPI>S</KARTTIPI>
			<STOKKOD>{stok_kod}</STOKKOD>
			<STOKADI>{stok_adi}</STOKADI>
			<MIKTAR>{ihtiyac_miktari}</MIKTAR>
			<BIRIMKOD>{birim_kod}</BIRIMKOD>
			<SEMBOL>TL</SEMBOL>
			<KUR_YEREL>1</KUR_YEREL>
			<VADE_TARIHI>{vade_tarihi}</VADE_TARIHI>
			<KDVORANI>{kdv_orani}</KDVORANI>
			<DEPOKOD>{depo_kod}</DEPOKOD>
			<OZELALAN1>{ozel_alan1}</OZELALAN1> 
			<ID>{6464 + sira_no}</ID>
		</Satir>"""
            xml_satirlar += satir_xml
            sira_no += 1

        elif ihtiyac_miktari > 0:
            print(f"UYARI: '{test_adi}' ürünü eşleştirme tablosunda bulunamadı ve XML'e eklenmedi.")

    # Ana XML Yapısı (Kısaltıldı)
    tarih_seri = excel_tarih_formatla(FISTAR_SEVKTAR)
    tarih_saat = FISTAR_SEVKTAR.strftime("%H:%M:%S")

    xml_tamamlanmis = f"""<?xml version="1.0" encoding="UTF-8"?>
<Fis>
	<OWNERID>12600</OWNERID>
	<FISNO>{FIS_NO}</FISNO>
	<CARIID>ARDAHAN</CARIID>
	<CARIADI>{CARI_ADI}</CARIADI>
	<SEHIR>ARDAHAN</SEHIR>
	<ULKE>TÜRKİYE</ULKE>
	<Notlar>GOND: LOST MED / TRB
YURTİÇİ / P.Ö /  KOLİ
LABORATUVAR DIKKATINE</Notlar>
	<FISTAR>{tarih_seri}</FISTAR>
	<FISSAAT>{tarih_saat}</FISSAAT>
	<SEVKTAR>{tarih_seri}</SEVKTAR>
	<SEVKSAAT>{tarih_saat}</SEVKSAAT>
	<SevkPlakalari/>
	<Satirlar>{xml_satirlar}
	</Satirlar>
</Fis>"""

    # XML dosyasını kaydet
    output_filename = f"ARDAHAN_SEVK.xml"
    with open(output_filename, "w", encoding="utf-8") as f:
        f.write(xml_tamamlanmis)

    print("-" * 50)
    print(f"BAŞARILI! '{output_filename}' dosyası oluşturuldu.")
    print(f"Toplam {sira_no - 1} adet ürün XML'e eklendi.")
    print("-" * 50)
    print("NOT: Kutu/Set bilgileri tanımlanan ürünler için MİKTAR ve OZELALAN1 güncellenmiştir.")


if __name__ == "__main__":
    excel_dosya_adi = "Kitap1.xlsx"

    excel_to_xml(excel_dosya_adi, STOK_MAP)