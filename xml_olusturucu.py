import pandas as pd
import datetime
import math
import re
import random
import os
import sys

# =========================================================================
# 1. ORTAK AYARLAR VE SABİT EŞLEŞTİRME TABLOSU
# =========================================================================

# --- DİNAMİK VADE TARİHİ TANIMLAMA (BUGÜN) ---
BUGUNUN_TARIHI = datetime.date.today()
DINAMIK_VADE_TARIHI_STR = BUGUNUN_TARIHI.strftime("%d.%m.%Y")

TURKISH_MONTHS = {
    1: "OCAK", 2: "ŞUBAT", 3: "MART", 4: "NİSAN",
    5: "MAYIS", 6: "HAZİRAN", 7: "TEMMUZ", 8: "AĞUSTOS",
    9: "EYLÜL", 10: "EKİM", 11: "KASIM", 12: "ARALIK"
}


def get_dynamic_sheet_name(sheet_prefix, override_month_name=None):
    """
    Mevcut ayın adını veya override edilen ayı kullanarak dinamik sayfa adını oluşturur.
    Format: "{ÖN EK} {AY ADI} SAYIM"
    """
    current_month_number = datetime.datetime.now().month
    current_month_name = TURKISH_MONTHS[current_month_number]

    month_to_use = override_month_name if override_month_name else current_month_name

    dynamic_sheet_name = f"{sheet_prefix} {month_to_use} SAYIM"

    print(f"(Debug) Otomatik Sayfa Adı Tahmini: {dynamic_sheet_name}")
    return dynamic_sheet_name, current_month_name  # Hem dinamik adı hem de mevcut ay adını (hata mesajı için) döndür


# Tuple Yapısı (11 Elemanlı):
# (0:STOK KODU, 1:STOK ADI, 2:MİKTAR, 3:BİRİM KOD, 4:VADE TARİHİ (BUGÜN), 5:KDV ORANI,
#  6:DEPO KOD, 7:OZELALAN1, 8:ID, 9:TEST_PER_KUTU, 10:SET_COUNT_PER_KUTU)
STOK_MAP = {
    "Glukoz (Serum/Plazma)": ("3L8222", "CC GLUKOZ (R*5)(20 ML)(1500 TEST)(ABBOTT)", "7500", "TEST",
                              DINAMIK_VADE_TARIHI_STR, "10", "LOST", "5K", "6464", 1500, 5),
    "Üre (Serum/Plazma)": ("4T1220", "CC ÜRE (R1*4+R2*4)(24,8 ML-10 ML)(1400 TEST)(ABBOTT)", "1400", "TEST",
                           DINAMIK_VADE_TARIHI_STR, "10", "LOST", "4K", "6465", 1400, 4),
    "Kreatinin (Serum/Plazma)": ("4S9520", "CC KREATİNİN 2 (3600 TEST)(ABBOTT)", "3600", "TEST",
                                 DINAMIK_VADE_TARIHI_STR, "10", "LOST", "3K", "6466", 3600, 1),
    # Kutu tek sayıldığı için 1 Set olarak varsayalım
    "Sodyum (Serum/Plazma)": ("2P3250", "CC ICT SAMPLE DILUENT (NA,K,CL) (R*10)(54 ML)(21000 TEST)(ABBOTT)", "21000",
                              "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "2K", "6467", 21000, 1),
    "Potasyum (Serum/Plazma)": ("2P3250", "CC ICT SAMPLE DILUENT (NA,K,CL) (R*10)(54 ML)(21000 TEST)(ABBOTT)", "21000",
                                "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "2K", "6467", 21000, 1),
    "Ürik asit (Serum/Plazma)": ("4T1320", "CC ÜRİK ASİT 2 (4* )(640 TEST)(ABBOTT)", "640", "TEST",
                                 DINAMIK_VADE_TARIHI_STR, "10", "LOST", "2K", "6468", 640, 1),
    "Bilirubin, total (Serum/Plazma)": ("8G6322", "CC DİREKT BİLİRUBİN (R1*10+R2*10)(39 ML-13 ML)(2000 TEST)(ABBOTT)",
                                        "2000", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "2K", "6469", 2000, 1),
    "Bilirubin, direkt (Serum/Plazma)": ("8G6322", "CC DİREKT BİLİRUBİN (R1*10+R2*10)(39 ML-13 ML)(2000 TEST)(ABBOTT)",
                                         "2000", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "2K", "6469", 2000, 1),
    "Total Protein": ("4U4420", "CC TOTAL PROTEİN 2 (4*19,6 ML)(800 TEST)(ABBOTT)", "800", "TEST",
                      DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K", "6470", 800, 1),
    "Albümin (Serum/Plazma)": ("4T3420", "CC ALBUMİN BCG2 (4* ML)  (1044 TEST) (ABBOTT)", "1044", "TEST",
                               DINAMIK_VADE_TARIHI_STR, "10", "LOST", "2K+3F", "6471", 1044, 1),
    "Alanin aminotransferaz (ALT)": ("4S8830", "CC ALT 2 (4*990) (3960 TEST) (ABBOTT)", "3960", "TEST",
                                     DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K+3F", "6472", 3960, 1),
    "Aspartat aminotransferaz (AST)": ("4S9030", "CC AST 2 (4*990)(3960 TEST)(ABBOTT)", "3960", "TEST",
                                       DINAMIK_VADE_TARIHI_STR, "10", "LOST", "2K", "6473", 3960, 1),
    "Gamma glutamil transferaz (GGT)": ("4T0020", "CC GGT 2 (R1*4+R2*4) (600 TEST) (ABBOTT)", "600", "TEST",
                                        DINAMIK_VADE_TARIHI_STR, "10", "LOST", "2K+1F", "6474", 600, 1),
    "Laktat dehidrogenaz (LDH)": ("4T0320", "CC LDH 2 (R1*4+R2*4) (600 TEST) (ABBOTT)", "600", "TEST",
                                  DINAMIK_VADE_TARIHI_STR, "10", "LOST", "4K+2F", "6475", 600, 1),
    "Amilaz (Serum/Plazma)": ("4S8920", "CC AMİLAZ 2 (4*14,5ML+4*13,4 ML)(640 TEST) (ABBOTT)", "640", "TEST",
                              DINAMIK_VADE_TARIHI_STR, "10", "LOST", "3K+2F", "6476", 640, 1),
    "Kalsiyum (Serum/Plazma)": ("4S9120", "CC KALSİYUM 2 (1200 TEST)(ABBOTT)", "1200", "TEST", DINAMIK_VADE_TARIHI_STR,
                                "10", "LOST", "1K", "6477", 1200, 1),
    "Fosfor (Serum/Plazma)": ("4T0730", "CC FOSFOR 2 (4*700 LÜK) (2800 TEST)(ABBOTT)", "2800", "TEST",
                              DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K", "6478", 2800, 1),
    "Kolesterol (Serum/Plazma)": ("4S9220", "CC KOLESTEROL 2 (4*21,6 ML)(1000 TEST) (ABBOTT)", "1000", "TEST",
                                  DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K+1F", "6479", 1000, 1),
    "Alkalen fosfataz (Serum/Plazma)": ("4S8720", "CC ALP (R1*8+R2*8) (1600 TEST ) (ABBOTT)", "1600", "TEST",
                                        DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K+3F", "6480", 1600, 1),
    "Magnezyum (Serum/Plazma)": ("3P6822", "CC MAGNEZYUM (R1*5+R2*5)(39 ML-11 ML)(1000 TEST)(ABBOTT)", "1000", "TEST",
                                 DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K+1F", "6481", 1000, 1),
    "İdrar/BOS protein": ("7D7932", "CC URİNE/CSF PROTEİN (UPRO) (209 TEST) (ABBOTT)", "209", "TEST",
                          DINAMIK_VADE_TARIHI_STR, "10", "LOST", "2F", "6482", 209, 1),
    "HDL kolesterol": ("4T0120", "CC HDL CHOLESTEROL (R1*2+R2*2+R3*1)(140 TEST)(ABBOTT)", "140", "TEST",
                       DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K", "6483", 140, 1),
    "Demir (Serum/Plazma)": ("4T0220", "CC IRON/UIBC/TIBC REAGENT KIT (R1*4+R2*4)(1000 TEST)(ABBOTT)", "1000", "TEST",
                             DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K", "6484", 1000, 1),
    "Demir bağlama kapasitesi": ("4T0220", "CC IRON/UIBC/TIBC REAGENT KIT (R1*4+R2*4)(1000 TEST)(ABBOTT)", "1000",
                                 "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K", "6484", 1000, 1),
    "TSH": ("7K6230", "ARC.TSH (Tiroit Uyarıcı Hormon) (200 TEST)(ABBOTT)", "200", "TEST", DINAMIK_VADE_TARIHI_STR,
            "10", "LOST", "2K", "6485", 200, 1),
    "Serbest T3": ("7K6533", "ARC.FREE T3 (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST",
                   "1K", "6486", 100, 1),
    "Serbest T4": ("7K6534", "ARC.FREE T4 (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST",
                   "1K", "6487", 100, 1),
    "Follikül stimülan hormon (FSH)": ("7K6630", "ARC.FSH (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR,
                                       "10", "LOST", "1K", "6488", 100, 1),
    "Lüteinizan hormon (LH)": ("7K6730", "ARC.LH (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                               "LOST", "1K", "6489", 100, 1),
    "Prolaktin": ("7K6830", "ARC.PROLACTIN (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST",
                  "1K", "6490", 100, 1),
    "Total Testesteron": ("7K7130", "ARC.TESTOSTERONE (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                          "LOST", "1K", "6491", 100, 1),
    "Kortizol (Serum/Plazma)": ("7K7230", "ARC.CORTISOL (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR,
                                "10", "LOST", "1K", "6492", 100, 1),
    "Estradiol (E2) (Serum/Plazma)": ("7K7330", "ARC.ESTRADIOL (100 TEST)(ABBOTT)", "100", "TEST",
                                      DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K", "6493", 100, 1),
    "Progesteron": ("7K7430", "ARC.PROGESTERONE (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                    "LOST", "1K", "6494", 100, 1),
    "İPTH": ("7K7530", "ARC.INTACT PTH (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K",
             "6495", 100, 1),
    "TOTAL PSA": ("7K7025", "ARC.TOTAL PSA (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST",
                  "1K", "6496", 100, 1),
    "Alfa-Fetoprotein (AFP)": ("7K7630", "ARC.AFP (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                               "LOST", "1K", "6497", 100, 1),
    "CEA": ("7K7730", "ARC.CEA (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K", "6498",
            100, 1),
    "CA 19-9 (Serum/Plazma)": ("7K7830", "ARC.CA 19-9 (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                               "LOST", "1K", "6499", 100, 1),
    "CA 15-3 (Serum/Plazma)": ("7K7930", "ARC.CA 15-3 (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                               "LOST", "1K", "6500", 100, 1),
    "CA 125 (Serum/Plazma)": ("7K8030", "ARC.CA 125 (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                              "LOST", "1K", "6501", 100, 1),
    "Ferritin (Serum/Plazma)": ("7K8130", "ARC.FERRITIN (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR,
                                "10", "LOST", "1K", "6502", 100, 1),
    "Beta HCG (Serum/Plazma)": ("7K8230", "ARC.B-HCG (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                                "LOST", "1K", "6503", 100, 1),
    "Vitamin B12": ("7K8330", "ARC.VITAMIN B12 (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                    "LOST", "1K", "6504", 100, 1),
    "Folat (Serum/Plazma)": ("7K8430", "ARC.FOLATE (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                             "LOST", "1K", "6505", 100, 1),
    "İnsülin": ("7K8530", "ARC.INSULIN (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K",
                "6506", 100, 1),
    "Anti TG": ("7K8630", "ARC.ANTI-TG (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K",
                "6507", 100, 1),
    "Anti TPO": ("7K8730", "ARC.ANTI-TPO (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST",
                 "1K", "6508", 100, 1),
    "25-Hidroksi vitamin D": ("7K8830", "ARC.25-OH VITAMIN D (100 TEST)(ABBOTT)", "100", "TEST",
                              DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K", "6509", 100, 1),
    "Prokalsitonin (Serum/Plazma)": ("7K8930", "ARC.PROCALCITONIN (100 TEST)(ABBOTT)", "100", "TEST",
                                     DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K", "6510", 100, 1),
    "Troponin I/T*": ("7K9030", "ARC.TROPONIN-I (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                      "LOST", "1K", "6511", 100, 1),
    "Tiroglobulin": ("7K9130", "ARC.THYROGLOBULIN (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10",
                     "LOST", "1K", "6512", 100, 1),
    "NT PROBNP": ("7K9230", "ARC.NT-proBNP (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST",
                  "1K", "6513", 100, 1),
    "LİPAZ": ("4T0420", "CC LİPAZ (4*27,5 ML)(800 TEST)(ABBOTT)", "800", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST",
              "1K", "6514", 800, 1),
    "C reaktif protein (CRP)": ("7K9330", "ARC.C-REAKTİF PROTEİN (100 TEST)(ABBOTT)", "100", "TEST",
                                DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K", "6515", 100, 1),
    "Glike hemoglobin (Hb A1C)": ("7K9430", "ARC.HBA1C (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR,
                                  "10", "LOST", "1K", "6516", 100, 1),
    "CKMB": ("7K9530", "ARC.CK-MB (100 TEST)(ABBOTT)", "100", "TEST", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K",
             "6517", 100, 1),
    "Tam idrar analizi (Strip+Mikroskopi)": ("8J2101", "İDRAR ANALİZ STRİPİ 10 LU (100 TEST)", "100", "TEST",
                                             DINAMIK_VADE_TARIHI_STR, "10", "LOST", "1K", "6518", 100, 1),
}

# Excel'deki Test Adı Sütun Adı (Bu sabittir)
TEST_ADI_SUTUNU = "TEST ADI"
# =========================================================================


# =========================================================================
# 2. HASTANE KONFİGÜRASYONU
# =========================================================================

HASTANE_CONFIGS = [
    # Ardahan ve Posof için yuvarlama kuralı artık aktif.
    # Boş hücreler (NaN) atlanacak, 0 girilenler 1 Set/Kutuya yuvarlanacaktır.
    {
        "CARI_ADI": "ARDAHAN DEVLET HASTANESİ",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Ardahan Devlet\Eliza\23-ARDAHAN DEVLET HASTANESİ ELİZA 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "ardahanEliza_islenmis_veri",
        "cari_id": "ARDAHAN",
        "sheet_prefix": "ARDAHAN DH",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "2 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya Yuvarla
    },
    {
        "CARI_ADI": "POSOF İLÇE DEVLET HASTANESİ",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Posof İlçe\Biyokimya\23-POSOF İLÇE DEVLET HASTANESİ BİYOKİMYA 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "posofBiokimya_islenmis_veri",
        "cari_id": "POSOF",
        "sheet_prefix": "POSOF DH",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "3 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya Yuvarla
    },
]


# =========================================================================
# 3. İŞLEME FONKSİYONLARI
# =========================================================================

def generate_xml_content(lines, cari_id, cari_adi):
    """Verilen satırları kullanarak XML içeriğini oluşturur."""

    # Eşsiz bir fatura numarası oluşturma
    # Tarih + Rastgele 6 haneli sayı
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    random_suffix = str(random.randint(100000, 999999))
    fatura_no = f"INV{timestamp}{random_suffix}"

    xml_lines = [
        '<?xml version="1.0" encoding="utf-8"?>',
        '<RECEIVE_ORDER>',
        f'<HEADER FATURA_NO="{fatura_no}" CARI_ID="{cari_id}" CARI_ADI="{cari_adi}" FATURA_TARIHI="{DINAMIK_VADE_TARIHI_STR}"/>',
        '<LINES>'
    ]

    for item in lines:
        # Tuple: (STOK KODU, STOK ADI, MİKTAR, BİRİM KOD, VADE TARİHİ, KDV ORANI, DEPO KOD, OZELALAN1, ID)
        stok_kodu, stok_adi, miktar, birim_kod, vade_tarihi, kdv_orani, depo_kod, ozelalan1, id_kodu = item

        xml_lines.append(
            f'<LINE STOK_KOD="{stok_kodu}" STOK_ADI="{stok_adi}" MIKTAR="{miktar}" BIRIM_KOD="{birim_kod}" '
            f'VADE_TARIHI="{vade_tarihi}" KDV_ORANI="{kdv_orani}" DEPO_KOD="{depo_kod}" '
            f'OZELALAN1="{ozelalan1}" ID="{id_kodu}"/>'
        )

    xml_lines.append('</LINES>')
    xml_lines.append('</RECEIVE_ORDER>')
    return "\n".join(xml_lines)


def process_hospital_data(config):
    """Belirtilen Excel dosyasını okur, veriyi işler ve XML çıktısı oluşturur."""

    cari_adi = config["CARI_ADI"]
    input_path = config["input_path"]
    output_prefix = config["output_prefix"]
    cari_id = config["cari_id"]
    sheet_prefix = config["sheet_prefix"]
    override_month_name = config.get("override_month_name")
    ihtiyac_sutunu = config["ihtiyac_sutunu"]
    apply_min_roundup = config["apply_min_roundup"]  # 0 girilenleri 1 Set/Kutuya yuvarlama kuralı

    print(f"\n--- {cari_adi} ({cari_id}) İşleniyor ---")
    print(f"Excel Yolu: {input_path}")
    print(f"İhtiyaç Sütunu: {ihtiyac_sutunu}")
    print(f"Minimum Yuvarlama Kuralı (0 -> 1 Set/Kutu): {'Aktif' if apply_min_roundup else 'Pasif'}")

    try:
        # Dinamik sayfa adını bulmaya çalış
        dynamic_sheet_name, current_month_name = get_dynamic_sheet_name(sheet_prefix, override_month_name)

        # Excel dosyasını oku
        try:
            df = pd.read_excel(input_path, sheet_name=dynamic_sheet_name)
        except ValueError as e:
            if f"Worksheet named '{dynamic_sheet_name}' not found" in str(e):
                print(f"HATA: Excel dosyasında beklenen sayfa adı bulunamadı. "
                      f"Beklenen ad: '{dynamic_sheet_name}'. Lütfen sayfa adını kontrol edin.")
            else:
                print(f"HATA: Excel dosyası okunurken bilinmeyen bir hata oluştu: {e}")
            return

        print(f"Sayfa: '{dynamic_sheet_name}' başarıyla okundu.")

        xml_lines_data = []

        # Her satırı döngüye al
        for index, row in df.iterrows():
            test_adi = str(row.get(TEST_ADI_SUTUNU, "")).strip()

            # 1. Filtre: Test Adı, STOK_MAP'te tanımlı mı?
            if test_adi not in STOK_MAP:
                continue

            stok_bilgisi = STOK_MAP[test_adi]
            TEST_PER_KUTU = stok_bilgisi[9]
            SET_COUNT_PER_KUTU = stok_bilgisi[10]

            ihtiyac_miktari_raw = row.get(ihtiyac_sutunu)

            # --- YENİ MANTIK KONTROLÜ BAŞLANGIÇ ---

            # KONTROL A: Eğer hücre boşsa (NaN/None) veya anlamsız bir değerse, bu satırı atla (dahil etme).
            if pd.isna(ihtiyac_miktari_raw):
                # print(f"-> Atlandı (Boş Hücre/NaN): '{test_adi}'") # Hata ayıklama
                continue

            try:
                # KONTROL B: Değeri tamsayıya çevir. (Örn: 0.0 -> 0, 500 -> 500)
                ihtiyac_miktari = int(ihtiyac_miktari_raw)
            except (ValueError, TypeError):
                # KONTROL C: Sayısal olmayan (metin vb.) bir değer girildiyse, satırı atla.
                # print(f"-> Atlandı (Geçersiz Miktar Tipi): '{test_adi}'") # Hata ayıklama
                continue

            # --- YENİ MANTIK KONTROLÜ BİTİŞ ---

            # 2. Yuvarlama Mantığı: Eğer miktar 0 ise ve yuvarlama kuralı aktif ise (True)
            if ihtiyac_miktari == 0:
                if apply_min_roundup:
                    # Minimum sipariş 1 Kutu olarak kabul edilir.
                    istenilen_kutu_miktari = 1
                    # Yeni ihtiyaç miktarını 1 Kutu'nun test sayısına eşitle
                    ihtiyac_miktari = istenilen_kutu_miktari * TEST_PER_KUTU
                    print(
                        f"-> Yuvarlama Yapıldı: '{test_adi}' için **0** miktarı minimum **1 Kutu** ({TEST_PER_KUTU} Test) olarak sipariş edildi.")
                else:
                    # Kural pasifse ve miktar 0 ise, satırı atla.
                    continue

            # Eğer ihtiyaç miktarı > 0 ise (AHBS'den geldi, ya da 0'dan yuvarlandı)
            if ihtiyac_miktari > 0:

                # İstenen Test miktarını Kutu miktarına çevir ve yukarı yuvarla (ceil)
                kutu_miktari_float = ihtiyac_miktari / TEST_PER_KUTU
                istenilen_kutu_miktari = math.ceil(kutu_miktari_float)

                # XML Satırını Oluştur:
                miktar_str = str(istenilen_kutu_miktari)
                birim_kod_yeni = "KUTU"

                new_stok_bilgisi = (
                    stok_bilgisi[0],  # STOK KODU
                    stok_bilgisi[1],  # STOK ADI
                    miktar_str,  # MİKTAR (Kutu Cinsinden)
                    birim_kod_yeni,  # BİRİM KOD (KUTU)
                    stok_bilgisi[4],  # VADE TARİHİ (Sabit)
                    stok_bilgisi[5],  # KDV ORANI
                    stok_bilgisi[6],  # DEPO KOD
                    stok_bilgisi[7],  # OZELALAN1
                    stok_bilgisi[8]  # ID
                )
                xml_lines_data.append(new_stok_bilgisi)

                print(
                    f"-> Eklendi: '{test_adi}' - İstenen Test: {ihtiyac_miktari} -> Kutu: {miktar_str} {birim_kod_yeni}")

            # Eğer miktar < 0 ise, atla
            else:
                continue

        # XML Oluşturma ve Kaydetme
        if not xml_lines_data:
            print(f"UYARI: '{cari_adi}' için XML oluşturulmadı. Eklenecek satır bulunamadı.")
            return

        xml_output = generate_xml_content(xml_lines_data, cari_id, cari_adi)

        output_filename = f"{output_prefix}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xml"

        output_dir = os.path.join(os.path.dirname(input_path), "OUTPUT")
        os.makedirs(output_dir, exist_ok=True)

        output_path = os.path.join(output_dir, output_filename)

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(xml_output)

        print(f"BAŞARILI: XML dosyası şuraya kaydedildi: {output_path}")

    except Exception as e:
        print(f"KRİTİK HATA: {cari_adi} ({cari_id}) işlenirken beklenmedik bir hata oluştu: {e}")


def main():
    """Ana program akışını yönetir."""

    # Tüm hastane konfigürasyonlarını sırayla işle
    for config in HASTANE_CONFIGS:
        process_hospital_data(config)

    print("\n\n--- Tüm Hastane İşlemleri Tamamlandı ---")


if __name__ == "__main__":
    main()