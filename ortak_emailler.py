import openpyxl

def oku_csv_email_set(dosya_adi, encodinglar=["utf-8", "cp1254", "iso-8859-9", "latin1"]):
    for encoding in encodinglar:
        try:
            with open(dosya_adi, encoding=encoding) as f:
                ilk_satir = f.readline().strip()
                if ";" in ilk_satir:
                    ayrac = ";"
                else:
                    ayrac = ","
                baslik = [b.strip() for b in ilk_satir.split(ayrac)]
                email_header = None
                for h in baslik:
                    if h.lower().replace(" ", "") in ["email", "eposta", "e-mail", "mail"]:
                        email_header = h
                        break
                if not email_header:
                    raise ValueError(f"Email sütunu bulunamadı! Başlıklar: {baslik}")
                email_index = baslik.index(email_header)
                email_set = set()
                for satir in f:
                    alanlar = [a.strip() for a in satir.strip().split(ayrac)]
                    if len(alanlar) > email_index:
                        email = alanlar[email_index].lower()
                        if email:
                            email_set.add(email)
                return email_set
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(f"{dosya_adi} dosyası şu encodinglerle açılamadı: {encodinglar}")

def oku_csv_email_satirli_s_sicil(dosya_adi, ortak_emailler, encodinglar=["utf-8", "cp1254", "iso-8859-9", "latin1"]):
    for encoding in encodinglar:
        try:
            with open(dosya_adi, encoding=encoding) as f:
                ilk_satir = f.readline().strip()
                if ";" in ilk_satir:
                    ayrac = ";"
                else:
                    ayrac = ","
                baslik = [b.strip() for b in ilk_satir.split(ayrac)]
                email_header = None
                sicil_header = None
                for h in baslik:
                    if h.lower().replace(" ", "") in ["email", "eposta", "e-mail", "mail"]:
                        email_header = h
                    if h.lower() == "sicil":
                        sicil_header = h
                if not email_header or not sicil_header:
                    raise ValueError(f"Email veya Sicil sütunu bulunamadı! Başlıklar: {baslik}")
                email_index = baslik.index(email_header)
                sicil_index = baslik.index(sicil_header)
                satirlar = []
                for satir in f:
                    alanlar = [a.strip() for a in satir.strip().split(ayrac)]
                    if len(alanlar) > max(email_index, sicil_index):
                        email = alanlar[email_index].lower()
                        sicil = alanlar[sicil_index]
                        if email in ortak_emailler and sicil.startswith("S"):
                            satirlar.append(alanlar)
                return baslik, satirlar
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(f"{dosya_adi} dosyası şu encodinglerle açılamadı: {encodinglar}")

def kullanmayan_stajyerler_excel_yaz(teammembers_csv, ortak_emailler, dosya_adi, encodinglar=["utf-8", "cp1254", "iso-8859-9", "latin1"]):
    for encoding in encodinglar:
        try:
            with open(teammembers_csv, encoding=encoding) as f:
                ilk_satir = f.readline().strip()
                if ";" in ilk_satir:
                    ayrac = ";"
                else:
                    ayrac = ","
                baslik = [b.strip() for b in ilk_satir.split(ayrac)]
                email_header = None
                usage_header = None
                premium_header = None
                for h in baslik:
                    if h.lower().replace(" ", "") in ["email", "eposta", "e-mail", "mail"]:
                        email_header = h
                    if h == "On-Demand Usage":
                        usage_header = h
                    if h == "Premium Requests":
                        premium_header = h
                if not email_header or not usage_header or not premium_header:
                    raise ValueError(f"Email, Premium Requests veya On-Demand Usage sütunu bulunamadı! Başlıklar: {baslik}")
                email_index = baslik.index(email_header)
                usage_index = baslik.index(usage_header)
                premium_index = baslik.index(premium_header)
                satirlar = []
                for satir in f:
                    alanlar = [a.strip() for a in satir.strip().split(ayrac)]
                    if len(alanlar) > max(email_index, usage_index, premium_index):
                        email = alanlar[email_index].lower()
                        usage = alanlar[usage_index].replace(",", ".")
                        premium = alanlar[premium_index].replace(",", ".")
                        try:
                            usage_float = float(usage)
                        except Exception:
                            usage_float = -1
                        try:
                            premium_float = float(premium)
                        except Exception:
                            premium_float = -1
                        if (
                            email in ortak_emailler and
                            usage_float == 0 and
                            premium_float == 0
                        ):
                            satirlar.append(alanlar)
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(baslik)
                for satir in satirlar:
                    ws.append(satir)
                wb.save(dosya_adi)
                return
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(f"{teammembers_csv} dosyası şu encodinglerle açılamadı: {encodinglar}")

# Dosya adları
kullanicilar_csv = "KullanıcıListesi.csv"
teammembers_csv = "team-members-8986257-2025-09-30.csv"

# 1. Ortak email setini bul
kullanici_emailler = oku_csv_email_set(kullanicilar_csv)
team_emailler = oku_csv_email_set(teammembers_csv)
ortak_emailler = kullanici_emailler & team_emailler

# 2. Sicili S ile başlayan ve ortak email sahibi olanları KullanıcıListesi.csv formatında al
baslik, satirlar = oku_csv_email_satirli_s_sicil(kullanicilar_csv, ortak_emailler)

# 3. stajyer_kullanıcılar.xlsx'e yaz
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Stajyerler"
ws.append(baslik)
for satir in satirlar:
    ws.append(satir)
wb.save("stajyer_kullanıcılar.xlsx")

# 4. kullanmayan_stajyerler.xlsx'e yaz
kullanmayan_stajyerler_excel_yaz(
    teammembers_csv,
    {satir[baslik.index([h for h in baslik if h.lower().replace(" ", "") in ["email", "eposta", "e-mail", "mail"]][0])] for satir in satirlar},
    "kullanmayan_stajyerler.xlsx"
)

print(f"Toplam {len(satirlar)} kişi stajyer_kullanıcılar.xlsx dosyasına kaydedildi.")
print("Premium Requests ve On-Demand Usage değerleri 0 olanlar kullanmayan_stajyerler.xlsx dosyasına kaydedildi.")