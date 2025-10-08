import pandas as pd
import datetime
import math
import re
import random
import os
import sys

# =========================================================================
# 1. KLASÖR YÖNETİMİ VE ÇIKTI TANIMLARI
# =========================================================================

# Merkeze bir çıktı klasörü tanımlayın. Bu klasör, script'in çalıştığı yerde oluşur.
MERKEZI_KLASOR_ADI = "XML_Ciktilar"

# Her çalıştırmada benzersiz bir alt klasör oluşturmak için zaman damgası
CALISMA_ZAMANI = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
# Örn: XML_Ciktilar/20251007_120530
TUM_CIKTILAR_YOLU = os.path.join(MERKEZI_KLASOR_ADI, CALISMA_ZAMANI)


def klasor_olustur(yol):
    """
    Belirtilen yolu (path) kontrol eder, yoksa oluşturur.
    exist_ok=True sayesinde klasör zaten varsa hata vermez.
    Başarısız olursa (izinler vb.) programı sonlandırır.
    """
    try:
        # os.makedirs, iç içe klasörleri (ana klasör ve zaman damgalı alt klasörü) oluşturur.
        os.makedirs(yol, exist_ok=True)
        print(f"\n[INFO] Çıktı klasörü hazırlandı: {yol}")
        return True
    except Exception as e:
        print(f"KRİTİK HATA: Merkezin çıktı klasörü oluşturulamadı. Lütfen klasör izinlerini kontrol edin: {e}")
        return False


# =========================================================================
# 2. ORTAK AYARLAR VE SABİT EŞLEŞTİRME TABLOSU
# =========================================================================

# --- DİNAMİK TARİH TANIMLAMA ---
BUGUNUN_TARIHI = datetime.datetime.today().date()

# STOK_MAP içinde kullanılan (DD.MM.YYYY formatında) tarih stringi.
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
    # Hem dinamik adı hem de mevcut ay adını (hata mesajı için) döndür
    return dynamic_sheet_name, current_month_name


# YENİ YARDIMCI TARİH FONKSİYONLARI: Depo programının istediği özel tarih formatları için.

def date_to_ole_format(date_obj):
    """
    datetime.date objesini OLE Automation Date formatına (1899-12-30'dan itibaren gün sayısı) çevirir.
    Bu format, depo programınızdaki <FISTAR> ve <SEVKTAR> alanları için gereklidir.
    """
    ole_base_date = datetime.date(1899, 12, 30)
    delta = date_obj - ole_base_date
    return str(delta.days)


def format_date_d_m_yyyy_manual(date_str):
    """
    'DD.MM.YYYY' formatındaki tarihi, gün ve ayın başındaki sıfırları atarak 'D.M.YYYY' formatına çevirir.
    (Örn: '08.10.2025' -> '8.10.2025'). Bu, <VADE_TARIHI> alanı için gereklidir.
    """
    try:
        dt_obj = datetime.datetime.strptime(date_str, "%d.%m.%Y").date()
        day = str(dt_obj.day)
        month = str(dt_obj.month)
        year = str(dt_obj.year)
        return f"{day}.{month}.{year}"
    except:
        return date_str  # Hata durumunda orijinali döndür


# Tuple Yapısı (11 Elemanlı):
# (0:STOK KODU, 1:STOK ADI, 2:MİKTAR [DİNAMİK], 3:BİRİM KOD [DİNAMİK], 4:VADE TARİHİ (BUGÜN, DD.MM.YYYY), 5:KDV ORANI,
#  6:DEPO KOD, 7:OZELALAN1 [DİNAMİK], 8:ID, 9:TEST_PER_KUTU, 10:SET_COUNT_PER_KUTU)
STOK_MAP = {
    # MİKTAR, BİRİM KOD ve OZELALAN1 (indeks 2, 3 ve 7) dinamik olarak hesaplanacağı için boş bırakılmıştır.
    "Glukoz (Serum/Plazma)": ("3L8222", "CC GLUKOZ (R*5)(20 ML)(1500 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR,
                              "10", "LOST", "", "6464", 1500, 5),
    "Üre (Serum/Plazma)": ("4T1220", "CC ÜRE (R1*4+R2*4)(24,8 ML-10 ML)(1400 TEST)(ABBOTT)", "", "",
                           DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6465", 1400, 4),
    "Kreatinin (Serum/Plazma)": ("4S9520", "CC KREATİNİN 2 (3600 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10",
                                 "LOST", "", "6466", 3600, 8),
    "Sodyum (Serum/Plazma)": ("2P3250", "CC ICT SAMPLE DILUENT (NA,K,CL) (R*10)(54 ML)(21000 TEST)(ABBOTT)", "", "",
                              DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6467", 21000, 10),
    "Potasyum (Serum/Plazma)": ("2P3250", "CC ICT SAMPLE DILUENT (NA,K,CL) (R*10)(54 ML)(21000 TEST)(ABBOTT)", "", "",
                                DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6467", 21000, 10),
    "Ürik asit (Serum/Plazma)": ("4T1320", "CC ÜRİK ASİT 2 (4* )(640 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR,
                                 "10", "LOST", "", "6468", 640, 4),
    "Bilirubin, total (Serum/Plazma)": ("4T0930", "CC TOTAL BİLİRUBİN 2 (R1*8+R2*8) (3600 TEST)(ABBOTT)", "", "",
                                        DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6469", 3600, 8),
    "Bilirubin, direkt (Serum/Plazma)": ("8G6322", "CC DİREKT BİLİRUBİN (R1*10+R2*10)(39 ML-13 ML)(2000 TEST)(ABBOTT)",
                                         "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6469", 2000, 10),
    "Total Protein": ("4U4430", "CC TOTAL PROTEİN 2 (4*69 ML)(3200 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR,
                      "10", "LOST", "", "6470", 3200, 4),
    "Albümin (Serum/Plazma)": ("4T3420", "CC ALBUMİN BCG2 (4* ML)  (1044 TEST) (ABBOTT)", "", "",
                               DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6471", 1044, 4),
    "Alanin aminotransferaz (ALT)": ("4S8830", "CC ALT 2 (4*990) (3960 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR,
                                     "10", "LOST", "", "6472", 3960, 4),
    "Aspartat aminotransferaz (AST)": ("4S9030", "CC AST 2 (4*990)(3960 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR,
                                       "10", "LOST", "", "6473", 3960, 4),
    "Gamma glutamil transferaz (GGT)": ("4T0020", "CC GGT 2 (R1*4+R2*4) (600 TEST) (ABBOTT)", "", "",
                                        DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6474", 600, 4),
    "Laktat dehidrogenaz (LDH)": ("4T0320", "CC LDH 2 (R1*4+R2*4) (600 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR,
                                  "10", "LOST", "", "6475", 600, 4),
    "Amilaz (Serum/Plazma)": ("4S8920", "CC AMİLAZ 2 (4*14,5ML+4*13,4 ML)(640 TEST) (ABBOTT)", "", "",
                              DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6476", 640, 4),
    "Kalsiyum (Serum/Plazma)": ("4S9120", "CC KALSİYUM 2 (1200 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10",
                                "LOST", "", "6477", 1200, 4),
    "Fosfor (Serum/Plazma)": ("4T0730", "CC FOSFOR 2 (4*700 LÜK) (2800 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR,
                              "10", "LOST", "", "6478", 2800, 4),
    "Kolesterol (Serum/Plazma)": ("4S9220", "CC KOLESTEROL 2 (4*21,6 ML)(1000 TEST) (ABBOTT)", "", "",
                                  DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6479", 1000, 4),
    "Alkalen fosfataz (Serum/Plazma)": ("4S8720", "CC ALP (R1*8+R2*8) (1600 TEST ) (ABBOTT)", "", "",
                                        DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6480", 1600, 8),
    "Magnezyum (Serum/Plazma)": ("3P6822", "CC MAGNEZYUM (R1*5+R2*5)(39 ML-11 ML)(1000 TEST)(ABBOTT)", "", "",
                                 DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6481", 1000, 5),
    "UPRO (24 saatlik idrar)": ("7D7932", "CC URİNE/CSF PROTEİN (UPRO) (209 TEST) (ABBOTT)", "", "",
                                DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6482", 209, 3),
    "HDL kolesterol": ("2R0621", "HDL DİREKT (5*57 ML+5*19) (1532 TEST) (ARCHEM-ABBOTT)", "", "",
                       DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6483", 1532, 5),
    "Demir (Serum/Plazma)": ("6R5521", "İRON (1100 TEST) (ARCHEM-ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10",
                             "LOST", "", "6484", 1100, 5),
    "Demir bağlama kapasitesi": ("1R8521", "DEMİR BAĞLAMA KAPASİTESİ (5*40 ML+5*12ML) (1159 TEST) (ARCHEM-ABBOTT)", "",
                                 "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6484", 1159, 5),
    "TSH": ("7K6225", "ARC.TSH (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6485", 100, 1),
    "Serbest T3": ("7K6327", "ARC.FREE T3 (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6486", 100,
                   1),
    "Serbest T4": ("7K6529", "ARC.FREE T4 (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6487", 100,
                   1),
    "Follikül stimülan hormon (FSH)": ("7K7525", "ARC.FSH (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST",
                                       "", "6488", 100, 1),
    "Lüteinizan hormon (LH)": ("2P4025", "ARC.LH (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "",
                               "6489", 100, 1),
    "Prolaktin": ("7K7625", "ARC.PROLACTİN (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6490",
                  100, 1),
    "Total Testesteron": ("2P1328", "ARC.TESTOSTERONE (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "",
                          "6491", 100, 1),
    "Kortizol (Serum/Plazma)": ("8D1525", "ARC.CORTİSOL (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST",
                                "", "6492", 100, 1),
    "Estradiol (E2) (Serum/Plazma)": ("7K7225", "ARC.ESTRADİOL (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10",
                                      "LOST", "", "6493", 100, 1),
    "Progesteron": ("7K7725", "ARC.PROGESTERON (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6494",
                    100, 1),
    "İPTH": ("8K2528", "ARC.İPTH (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6495", 100, 1),
    "TOTAL PSA": ("7K7025", "ARC.TOTAL PSA (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6496",
                  100, 1),
    "Alfa-Fetoprotein (AFP)": ("3P3625", "ARC.AFP (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "",
                               "6497", 100, 1),
    "CEA": ("7K6827", "ARC.CEA (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6498", 100, 1),
    "CA 19-9 (Serum/Plazma)": ("2K9132", "ARC.CA 19-9 (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "",
                               "6499", 100, 1),
    "CA 15-3 (Serum/Plazma)": ("2K4427", "ARC.CA15-3 (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "",
                               "6500", 100, 1),
    "CA 125 (Serum/Plazma)": ("2K4529", "ARC.CA 125 (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "",
                              "6501", 100, 1),
    "Ferritin (Serum/Plazma)": ("7K5925", "ARC.FERRİTİN (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST",
                                "", "6502", 100, 1),
    "Beta HCG (Serum/Plazma)": ("7K7825", "ARC.TOTAL BHCG (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST",
                                "", "6503", 100, 1),
    "Vitamin B12": ("7K6125", "ARC.VİTAMİN B12 (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6504",
                    100, 1),
    "Folat (Serum/Plazma)": ("1P7425", "ARC.FOLATE (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "",
                             "6505", 100, 1),
    "İnsülin": ("8K4128", "ARC.İNSULİN (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6506", 100,
                1),
    "Anti TG": ("2K4625", "ARC.ANTİ TG (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6507", 100,
                1),
    "Anti TPO": ("2K4725", "ARC.ANTİ TPO (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6508", 100,
                 1),
    "25-Hidroksi vitamin D": ("5P0225", "ARC.VİTAMİN D (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "",
                              "6509", 100, 1),
    "Prokalsitonin (Serum/Plazma)": ("6P2225", "ARC.BRAHMS PCT (100 TEST)(PROKALSİTONİN)", "", "",
                                     DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6510", 100, 1),
    "Troponin I/T*": ("3P2527", "ARC.TROPONİN (100 TEST )", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6511",
                      100, 1),
    "Tiroglobulin": ("5P2025", "ARC.TG (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6512", 100,
                     1),
    "NT PROBNP": ("2R1025", "ARC.NT-proBNP (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6513",
                  100, 1),
    "LİPAZ": ("6R5523", "LİPAZ (1540 TEST) (ARCHEM-ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6514", 1540, 5),
    "C reaktif protein (CRP)": ("6T8163", "CRP TURBI WR (5*455) (2275 TEST) (ARCHEM-ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6515", 2275, 5),
    "Glike hemoglobin (Hb A1C)": ("1R8721", "HbA1c DIREKT (571 TEST) (ARCHEM-ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6516", 571, 1),
    "CKMB": ("1R8821", "CKMB (R1*5+R2*5)(54 ML-13,5 ML)(1795 TEST)(ARCHEM-ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 1795, 5),
    "Micro Albümin (MALB)": ("2K9821", "CC MİCROALBUMİN (500 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 500, 2),
    "Etanol ( Serum/ Plazma )": ("3L3620", "CC ETHANOL (200 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 200, 2),
    "Lityum (Serum/ Plazma)": ("8L2530", "CC LİTHİUM (2*R)(194 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 194, 2),
    "Valproik Asit (Serum/Plazma)": ("1E1320", "CC VALPROİC ACİD (180 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 180, 2),
    "Karbamazepin (Serum/Plazma)": ("5P0521", "CC CARBAMAZEPİNE (300 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 300, 2),
    "Amfetamin (İdrar)": ("3L3720", "CC AMPHETAMİNE (500 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 500, 2),
    "Benzodiyazepinler (İdrar)": ("3L3920", "CC BENZODİAZEPİNES (500 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 500, 2),
    "Kannabinoidler (THC) (İdrar)": ("3L4120", "CC CANNABİNOİDS (500 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 500, 2),
    "Kokain  (İdrar)": ("3L4020", "CC KOKAİN (500 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 500, 2),
    "Opiyatlar (İdrar)": ("3L3420", "CC OPİATES (500 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 500, 2),
    "D-dimer (Kantitatif)": ("6T0425", "D-DİMER 600 TEST (CC REAGENTS) (ARCHEM-ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 600, 3),
    "Total IgE": ("1R8921", "IGE (364 TEST) (ARCHEM-ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 364, 2),
    "Digoksin (Serum/Plazma)": ("1E0621", "CC DİGOXİN (450 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 450, 2),
    "LAKTİK ASİT": ("4T3020", "CC LACTİC ACİD (400 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 400, 2),
    "HBsAg": ("2G2225", "ARC.HBSAG (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti HCV": ("6C3727", "ARC.ANTİ HCV (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti HIV": ("6C3727", "ARC.ANTİ HİV AG/AB COMBO (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti HBs": ("7C1829", "ARC.ANTİ HBS (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti HAV IgM": ("6C3027", "ARC.ANTİ HAV IGM (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti HAV IgG (ANTİ HAV TOTAL)": ("6C2927", "ARC.ANTİ HAV TOTAL (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "HBeAg": ("6C3227", "ARC.HBEAG (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti HBe": ("6C3425", "ARC.ANTİ HBE (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "ANTİ HBC": ("8L4425", "ARC.ANTİ HBC (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti HBc IgM": ("6C3327", "ARC.ANTİ HBC IGM (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti toxoplazma IgM": ("6C2025", "ARC.TOXO IGM (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti toxoplazma IgG": ("6C1925", "ARC.TOXO IGG (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti rubella IgM": ("6C2025", "ARC.RUBELLA IGM (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti rubella IgG": ("6C1726", "ARC.RUBELLA IGG (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti CMV IgM": ("6C1625", "ARC.CMV IGM (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Anti CMV IgG": ("6C1525", "ARC.CMV IGG (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "ANTİ CCP": ("1P6525", "ARC.ANTİ CCP (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "Sifiliz  (VDRL)": ("8D0632", "ARC.SYPHİLİS (100 TEST) ", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "EBV EBNA IGG": ("3P6725", "ARC.EBV EBNA-1 IGG (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "EBNA VCA IGG": ("3P6525", "ARC.EBV VCA IGG (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),
    "EBNA VCA IGM": ("3P6625", "ARC.EBV VCA IGM (100 TEST)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 100, 1),



    # ********************************************************* ALNTY KITLER  *********************************************************
    "ALNTY_TSH": ("7P4820", "ALNTY. TSH (2*100 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 200, 2),
    "ALNTY_Serbest T3": ("7P6920", "ALNTY. FREE T3 (2*100 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 200, 2),
    "ALNTY_Serbest T4": ("7P7020", "ALNTY. FREE T4 (2*100 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 200, 2),
    "ALNTY_Estradiol (E2)": ("7P5020", "ALNTY. ESTRADİOL (2*100 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 200, 2),
    "ALNTY_Beta HCG": ("7P5130", "ALNTY. TOTAL HCG (2*600 TEST)(ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 1200, 2),
    "ALNTY_Prokalsitonin": ("1R1832", "ALNTY. PROKALSİTONİN (2*500 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 1000, 2),
    "ALNTY_Troponin I/T": ("8P1334", "ALNTY. TROPONİN (2*600 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 1200, 2),
    "ALNTY_Kütle CK-MB": ("4V3820", "ALNTY. CKMB (2*100 TEST) (ABBOTT) HORMON", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 200, 2),
    "ALNTY_BNP": ("8P2420", "ALNTY. BNP (2*100 TEST) (AABOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 200, 2),
    "ALNTY_MYOGLOBİN": ("9P3920", "ALNTY. MYOGLOBİN (2*100 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 200, 2),
    "ALNTY_SİKLOSPORİN": ("4V3720", "ALNTY. CYCLOSPORİNE (2*100 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 200, 2),
    "ALNTY_TAKROLİMUS": ("9P4220", "ALNTY. TACROLİMUS (2*100 TEST) (ABBOTT)", "", "", DINAMIK_VADE_TARIHI_STR, "10", "LOST", "", "6517", 200, 2),
}

# Excel'deki Test Adı Sütun Adı (Bu sabittir)
TEST_ADI_SUTUNU = "TEST ADI"
# =========================================================================


# =========================================================================
# 3. HASTANE KONFİGÜRASYONU (Değişiklik yapılmamıştır)
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
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
    {
        "CARI_ADI": "ARDAHAN DEVLET HASTANESİ",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Ardahan Devlet\Hormon-Biyokimya\24-ARDAHAN  DEVLET HASTANESİ BİYOKİMYA-HORMON 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "ardahanBiyokimyaHormon_islenmis_veri",
        "cari_id": "ARDAHAN",
        "sheet_prefix": "ARDAHAN DH",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "2,5 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
    {
        "CARI_ADI": "ARDAHAN MERKEZ HALK SAĞLIĞI LABORATUVARI",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Ardahan HSM\27-ARDAHAN HALK SAĞLIĞI LABORATUVARI 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "ardahanHSM_islenmis_veri",
        "cari_id": "ARDAHAN MERKEZ",
        "sheet_prefix": "ARDAHAN HALK",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "3 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
    {
        "CARI_ADI": "ATATÜRK ÜNİVERSİTESİ ARAŞTIRMA HASTANESİ MİKROBİYOLOJİ LABORATUVARI",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Atatürk Üniversitesi\29-ATATÜRK ÜNİVERSİTESİ ARAŞTIRMA HASTANESİ 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "ataturkUni_islenmis_veri",
        "cari_id": "ATATÜRK",
        "sheet_prefix": "ATATÜRK ÜNi.",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "2 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
    {
        "CARI_ADI": "GÖLE DEVLET HASTANESİ",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Göle Devlet\25-GÖLE DEVLET HASTANESİ 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "gole_islenmis_veri",
        "cari_id": "GÖLE",
        "sheet_prefix": "GÖLE DH",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "2 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
    {
        "CARI_ADI": "IĞDIR DR. NEVRUZ EREZ DEVLET HASTANESİ",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Iğdır Devlet\20-IĞDIR DEVLET HASTANESİ 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "igdirDH_islenmis_veri",
        "cari_id": "IĞDIR",
        "sheet_prefix": "IĞDIR DH",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "2,5 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
    {
        "CARI_ADI": "IĞDIR HALK SAĞLIĞI LABORATUVARI",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Iğdır Halk\28-IĞDIR HALK SAĞLIĞI LABORATUVARI 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "igdirHSM_islenmis_veri",
        "cari_id": "IĞDIR HALK",
        "sheet_prefix": "IĞDIR HSM",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "2,5 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
    {
        "CARI_ADI": "GİRESUN ÖZEL KENT HASTANESİ",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Kent Hastanesi\30-GİRESUN ÖZEL KENT HASTANESİ BİYOKİMYA-HORMON-ELİZA 2025 YILI SAYIMLAR....xlsx",
        "output_prefix": "giresunKent_islenmis_veri",
        "cari_id": "GİRESUN ÖZEL KENT",
        "sheet_prefix": "GİRESUN KENT",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "2,5 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
    {
        "CARI_ADI": "KTÜ TIP FAKÜLTESİ FARABİ HASTANESİ",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Ktü Hastanesi\31-KTÜ ACİL HORMON 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "ktu_islenmis_veri",
        "cari_id": "KTÜ TIP",
        "sheet_prefix": "KTÜ ACİL",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "2,5 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
    {
        "CARI_ADI": "POSOF İLÇE DEVLET HASTANESİ",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Posof\26-POSOF DEVLET HASTANESİ 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "posof_islenmis_veri",
        "cari_id": "POSOF",
        "sheet_prefix": "POSOF DH",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "3 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
    {
        "CARI_ADI": "RİZE DEVLET HASTANESİ",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Rize Devlet\22-RİZE DEVLET HASTANESİ 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "rizeDevlet_islenmis_veri",
        "cari_id": "RİZE",
        "sheet_prefix": "RİZE DEVLET",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "2 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
    {
        "CARI_ADI": "RİZE EĞİTİM VE ARAŞTIRMA HASTANESİ",
        "input_path": r"E:\Sayımlar\Lost Hastane Sayımları\Eylül\Rize Eğitim\21-RİZE EĞİTİM VE ARAŞTIRMA HASTANESİ 2025 YILI SAYIMLARI.XLSX",
        "output_prefix": "rizeEGT_islenmis_veri",
        "cari_id": "RİZE EĞİTİM",
        "sheet_prefix": "RİZE EĞİTİM",
        "override_month_name": "EYLÜL",
        "ihtiyac_sutunu": "2 AYLIK İHTİYAÇ MİKTARI (TEST)",
        "apply_min_roundup": True  # <-- 0 ise 1 Set/Kutuya yuvarla
    },
]


# =========================================================================
# 4. İŞLEME FONKSİYONLARI
# =========================================================================

# DEPO PROGRAMI FORMATINA UYGUN XML OLUŞTURMA FONKSİYONU
def generate_xml_content(lines, cari_id, cari_adi):
    """
    Verilen satırları kullanarak DEPO PROGRAMI formatına uygun (element tabanlı) XML içeriğini oluşturur.
    """

    # --- Sabit/Tahmini Başlık Değerleri ---
    OWNERID = "12600"
    FISNO = str(random.randint(90000000, 99999999))

    # SEHIR/ULKE: Cari ID'ye göre tahmin.
    SEHIR = "BİLİNMİYOR"
    if "ARDAHAN" in cari_id.upper() or "POSOF" in cari_id.upper():
        SEHIR = "ARDAHAN"
    elif "IĞDIR" in cari_id.upper():
        SEHIR = "IĞDIR"
    elif "RİZE" in cari_id.upper():
        SEHIR = "RİZE"
    elif "KTÜ" in cari_id.upper() or "ATATÜRK" in cari_id.upper():
        SEHIR = "TRABZON"
    elif "GİRESUN" in cari_id.upper():
        SEHIR = "GİRESUN"

    ULKE = "TÜRKİYE"

    # YENİ EKLENEN NOT ALANI
    NOTLAR = "GOND: LOST MED / TRB YURTİÇİ / P.Ö / KOLİ LABORATUVAR DIKKATINE"
    # -------------------------------------------------------------------

    # Tarih ve Saatler
    bugunun_tarihi_obj = datetime.datetime.today().date()
    xml_saat_str = datetime.datetime.now().strftime("%H:%M:%S")

    # Kritik: OLE Automation Date formatına çevir (Örn: 45938)
    fistar_ole = date_to_ole_format(bugunun_tarihi_obj)

    xml_lines = [
        '<?xml version="1.0" encoding="utf-8"?>',
        '<Fis>'
    ]

    # HEADER (Şimdi <Fis> altında alt elementler olarak)
    xml_lines.append(f'<OWNERID>{OWNERID}</OWNERID>')
    xml_lines.append(f'<FISNO>{FISNO}</FISNO>')
    xml_lines.append(f'<CARIID>{cari_id}</CARIID>')
    xml_lines.append(f'<CARIADI>{cari_adi}</CARIADI>')
    xml_lines.append(f'<SEHIR>{SEHIR}</SEHIR>')
    xml_lines.append(f'<ULKE>{ULKE}</ULKE>')
    xml_lines.append(f'<FISTAR>{fistar_ole}</FISTAR>')  # OLE Tarih formatı
    xml_lines.append(f'<FISSAAT>{xml_saat_str}</FISSAAT>')
    xml_lines.append(f'<SEVKTAR>{fistar_ole}</SEVKTAR>')  # OLE Tarih formatı
    xml_lines.append(f'<SEVKSAAT>{xml_saat_str}</SEVKSAAT>')
    xml_lines.append(f'<SevkPlakalari/>')  # Boş etiket

    # İSTENEN NOTLAR ETİKETİ BURAYA EKLENDİ
    xml_lines.append(f'<Notlar>{NOTLAR}</Notlar>')

    xml_lines.append('<Satirlar>')

    # LINE Loop
    for index, item in enumerate(lines):
        line_number = index + 1
        # Tuple Eşleşmesi:
        # 0:STOK KODU, 1:STOK ADI, 2:MİKTAR (TEST), 3:BİRİM KOD (TEST), 4:VADE TARİHİ, 5:KDV ORANI, 6:DEPO KOD, 7:OZELALAN1, 8:ID
        stok_kodu, stok_adi, miktar, birim_kod, vade_tarihi_str, kdv_orani, depo_kod, ozelalan1, id_kodu = item

        # VADE_TARIHI formatını depo programının istediği şekilde düzenle (D.M.YYYY)
        vade_tarihi_formatli = format_date_d_m_yyyy_manual(vade_tarihi_str)

        # Sabit Satır Bilgileri
        KARTTIPI = "S"
        SEMBOL = "TL"
        KUR_YEREL = "1"
        ERP_LOT_CIKIS_GIRIS_HARID = "-1"
        SIPHARID = "-1"

        # XML Satırı (Elementler olarak)
        xml_lines.append('<Satir>')
        xml_lines.append(f'<SIRANO>{line_number}</SIRANO>')
        xml_lines.append(f'<KARTTIPI>{KARTTIPI}</KARTTIPI>')
        xml_lines.append(f'<STOKKOD>{stok_kodu}</STOKKOD>')
        xml_lines.append(f'<STOKADI>{stok_adi}</STOKADI>')
        xml_lines.append(f'<MIKTAR>{miktar}</MIKTAR>')  # TEST Miktarı
        xml_lines.append(f'<BIRIMKOD>{birim_kod}</BIRIMKOD>')  # TEST Birimi
        xml_lines.append(f'<SEMBOL>{SEMBOL}</SEMBOL>')
        xml_lines.append(f'<KUR_YEREL>{KUR_YEREL}</KUR_YEREL>')
        xml_lines.append(f'<VADE_TARIHI>{vade_tarihi_formatli}</VADE_TARIHI>')  # D.M.YYYY formatında
        xml_lines.append(f'<KDVORANI>{kdv_orani}</KDVORANI>')
        xml_lines.append(f'<DEPOKOD>{depo_kod}</DEPOKOD>')
        xml_lines.append(f'<ID>{id_kodu}</ID>')
        # DİNAMİK OZELALAN1 (Kutu+Set Bilgisi)
        xml_lines.append(f'<OZELALAN1>{ozelalan1}</OZELALAN1>')
        xml_lines.append(f'<ERP_LOT_CIKIS_GIRIS_HARID>{ERP_LOT_CIKIS_GIRIS_HARID}</ERP_LOT_CIKIS_GIRIS_HARID>')
        xml_lines.append(f'<SIPHARID>{SIPHARID}</SIPHARID>')
        xml_lines.append('</Satir>')

    xml_lines.append('</Satirlar>')
    xml_lines.append('</Fis>')
    return "\n".join(xml_lines)


def process_hospital_data(config):
    """Belirtilen Excel dosyasını okur, veriyi işler ve XML çıktısı oluşturur."""
    global TUM_CIKTILAR_YOLU

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
            # NaN veya boşluk içeren değerleri atla.
            test_adi_raw = row.get(TEST_ADI_SUTUNU)
            if pd.isna(test_adi_raw):
                continue

            test_adi = str(test_adi_raw).strip()

            #print(f"DEBUG: Okunan Test Adı: '{test_adi}'")

            # 1. Filtre: Test Adı, STOK_MAP'te tanımlı mı?
            if test_adi not in STOK_MAP:
                continue

            stok_bilgisi = STOK_MAP[test_adi]

            # Set bazlı hesaplama için gerekli bilgileri al
            TEST_PER_KUTU = stok_bilgisi[9]
            SET_PER_KUTU = stok_bilgisi[10]

            # Set başına düşen Test miktarı
            if SET_PER_KUTU <= 0:
                TEST_PER_SET = TEST_PER_KUTU
                if TEST_PER_SET == 0:
                    TEST_PER_SET = 1
            else:
                TEST_PER_SET = TEST_PER_KUTU / SET_PER_KUTU

            ihtiyac_miktari_raw = row.get(ihtiyac_sutunu)

            # --- MİKTAR KONTROLÜ BAŞLANGIÇ ---

            if pd.isna(ihtiyac_miktari_raw):
                continue

            try:
                ihtiyac_miktari = int(ihtiyac_miktari_raw)
            except (ValueError, TypeError):
                continue

            # --- MİKTAR KONTROLÜ BİTİŞ ---

            # 2. Yuvarlama Mantığı (SET Bazlı)

            istenilen_set_miktari = 0

            # 1. Aşama: Sıfır ve Negatif Kontrolü
            if ihtiyac_miktari <= 0:
                if ihtiyac_miktari == 0 and apply_min_roundup:
                    # Minimum 1 Set sipariş et
                    istenilen_set_miktari = 1
                else:
                    continue

            # 2. Aşama: Pozitif Miktar Hesaplama ve Set Bazlı Yuvarlama
            elif ihtiyac_miktari > 0:
                # İstenen Test miktarını Set sayısına çevir
                set_miktari_float = ihtiyac_miktari / TEST_PER_SET

                # Yukarıya Yuvarlama (math.ceil) kullanarak en yakın TAM Set sayısını bul.
                istenilen_set_miktari = math.ceil(set_miktari_float)

            # --- XML ÇIKTI DEĞERLERİNİN HESAPLANMASI ---

            # Hesaplanan Set sayısını, XML'e yazılacak olan toplam TEST miktarına çevir.
            xml_test_miktari = istenilen_set_miktari * TEST_PER_SET

            miktar_str = str(int(xml_test_miktari))
            birim_kod_yeni = "TEST"

            # --- OZELALAN1 HESAPLAMASI (Kutu + Set Bilgisi) ---

            ozelalan1_str = ""

            if SET_PER_KUTU > 0 and istenilen_set_miktari > 0:
                tam_kutu_sayisi = istenilen_set_miktari // SET_PER_KUTU
                kalan_set_sayisi = istenilen_set_miktari % SET_PER_KUTU

                if tam_kutu_sayisi > 0:
                    ozelalan1_str += f"{tam_kutu_sayisi}K"

                if kalan_set_sayisi > 0:
                    if ozelalan1_str:
                        ozelalan1_str += " + "
                    ozelalan1_str += f"{kalan_set_sayisi}SET"

                if not ozelalan1_str and istenilen_set_miktari > 0:
                    # Sadece tam kutu sipariş edilmişse (Örn: 4 set = 1K)
                    ozelalan1_str = f"{istenilen_set_miktari // SET_PER_KUTU}K"

            elif istenilen_set_miktari > 0:
                # Kutu bilgisi yoksa sadece set sayısını yaz
                ozelalan1_str = f"{istenilen_set_miktari}SET"

            # --- XML ÇIKTI VERİSİNİN OLUŞTURULMASI ---

            # Yeni tuple'ı oluştur (miktar, birim kodu ve OZELALAN1 güncel değerlerle değiştir)
            # Tuple yapısı: 0, 1, 2(MİKTAR), 3(BİRİM KOD), 4, 5, 6, 7(OZELALAN1), 8(ID)
            new_stok_bilgisi = (
                stok_bilgisi[0],  # 0: STOK KODU
                stok_bilgisi[1],  # 1: STOK ADI
                miktar_str,  # 2: MİKTAR (TEST Cinsinden)
                birim_kod_yeni,  # 3: BİRİM KOD (TEST)
                stok_bilgisi[4],  # 4: VADE TARİHİ
                stok_bilgisi[5],  # 5: KDV ORANI
                stok_bilgisi[6],  # 6: DEPO KOD
                ozelalan1_str,  # 7: OZELALAN1 (Kutu+Set Bilgisi)
                stok_bilgisi[8]  # 8: ID
            )
            xml_lines_data.append(new_stok_bilgisi)

            print(
                f"-> XML: {test_adi} -> İhtiyaç: {ihtiyac_miktari} Test -> Sipariş: {ozelalan1_str} ({miktar_str} TEST)")

        # XML Oluşturma ve Kaydetme
        if not xml_lines_data:
            print(
                f"UYARI: '{cari_adi}' için XML oluşturulmadı. Eklenecek satır bulunamadı (Tüm satırlar boş, sıfır veya eşleşmeyen test adlarıydı).")
            return

        # generate_xml_content fonksiyonu burada çağrılır.
        xml_output = generate_xml_content(xml_lines_data, cari_id, cari_adi)

        output_filename = f"{output_prefix}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xml"

        # Merkezi klasör yolunu kullan
        output_path = os.path.join(TUM_CIKTILAR_YOLU, output_filename)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(xml_output)

        print(f"[BAŞARILI] XML dosyası kaydedildi: {output_path}")

    except Exception as e:
        print(f"KRİTİK HATA: '{cari_adi}' verisi işlenirken bir hata oluştu: {e}")
        import traceback
        print(traceback.format_exc())  # Hata detaylarını da yazdır


# =========================================================================
# 5. ANA ÇALIŞTIRMA BLOĞU
# =========================================================================

if __name__ == "__main__":
    if not klasor_olustur(TUM_CIKTILAR_YOLU):
        sys.exit(1)

    print("\n[BAŞLANGIÇ] Excel verileri okunuyor ve XML'ler oluşturuluyor...")
    for config in HASTANE_CONFIGS:
        process_hospital_data(config)

    print(f"\n[BİTİŞ] Tüm işlemler tamamlandı. Çıktılar: {TUM_CIKTILAR_YOLU}")