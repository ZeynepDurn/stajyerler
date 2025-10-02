import openpyxl
import json

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

def oku_json_email_lastactive(json_dosya):
    """JSON dosyasından email -> lastactive eşlemesi yapar"""
    with open(json_dosya, encoding="utf-8") as f:
        data = json.load(f)
        email2lastactive = {}
        email2user = {}
        for kayit in data:
            email = (kayit.get("email") or "").strip().lower()
            lastactive = kayit.get("lastactive", "")
            if email:
                email2lastactive[email] = lastactive
                email2user[email] = kayit
        return email2lastactive, email2user

def oku_json_email_set(json_dosya):
    with open(json_dosya, encoding="utf-8") as f:
        data = json.load(f)
        return { (k.get("email") or "").strip().lower() for k in data if k.get("email") }

def kullanmayan_stajyerler_excel_yaz_json(teammembers_json, ortak_emailler, dosya_adi):
    with open(teammembers_json, encoding="utf-8") as f:
        data = json.load(f)
    # Başlıklar
    basliklar = list(data[0].keys())
    # Kullanılmayanlar
    satirlar = []
    for kayit in data:
        email = (kayit.get("email") or "").strip().lower()
        usage = float(str(kayit.get("on_demand_usage", "0")).replace(",", "."))
        premium = float(str(kayit.get("premium_requests", "0")).replace(",", "."))
        if (email in ortak_emailler) and usage == 0 and premium == 0:
            satirlar.append([kayit.get(b, "") for b in basliklar])
    # Excel yaz
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(basliklar)
    for satir in satirlar:
        ws.append(satir)
    wb.save(dosya_adi)

# Dosya adları
kullanicilar_csv = "KullanıcıListesi.csv"
teammembers_json = "team_members_all.json"

# 1. Ortak email setini bul
kullanici_emailler = set()
for encoding in ["utf-8", "cp1254", "iso-8859-9", "latin1"]:
    try:
        with open(kullanicilar_csv, encoding=encoding) as f:
            ilk_satir = f.readline().strip()
            if ";" in ilk_satir:
                ayrac = ";"
            else:
                ayrac = ","
            baslik = [b.strip() for b in ilk_satir.split(ayrac)]
            email_index = [i for i, h in enumerate(baslik) if h.lower().replace(" ", "") in ["email", "eposta", "e-mail", "mail"]][0]
            for satir in f:
                alanlar = [a.strip() for a in satir.strip().split(ayrac)]
                if len(alanlar) > email_index:
                    email = alanlar[email_index].lower()
                    if email:
                        kullanici_emailler.add(email)
        break
    except UnicodeDecodeError:
        continue

team_emailler = oku_json_email_set(teammembers_json)
ortak_emailler = kullanici_emailler & team_emailler

# 2. Sicili S ile başlayan ve ortak email sahibi olanları KullanıcıListesi.csv formatında al
baslik, satirlar = oku_csv_email_satirli_s_sicil(kullanicilar_csv, ortak_emailler)
email2lastactive, _ = oku_json_email_lastactive(teammembers_json)

# 3. stajyer_kullanıcılar.xlsx'e yaz
# Başlıkları istenen sırada düzenle
yeni_baslik = ["Sicil", "Ad", "Soyad", "Ünvan", "Email", "Departman İsmi", "Son Aktif Olduğu Tarih"]
# Var olan başlıklardan eşleştirme için index bul
def baslik_index(basliklar, aranacak):
    for idx, b in enumerate(basliklar):
        if b.lower().replace("ü", "u").replace("ı", "i").replace("ş", "s").replace(" ", "") == aranacak.lower().replace("ü", "u").replace("ı", "i").replace("ş", "s").replace(" ", ""):
            return idx
    return -1

ind_sicil = baslik_index(baslik, "sicil")
ind_ad = baslik_index(baslik, "ad")
ind_soyad = baslik_index(baslik, "soyad")
ind_unvan = baslik_index(baslik, "ünvan")
if ind_unvan == -1: ind_unvan = baslik_index(baslik, "unvan") # olası karakter hatası için
ind_email = baslik_index(baslik, "email")
ind_departman = baslik_index(baslik, "departman ismi")
if ind_departman == -1: ind_departman = baslik_index(baslik, "departmanismi")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Stajyerler"
ws.append(yeni_baslik)
for satir in satirlar:
    email = satir[ind_email].lower()
    lastactive = email2lastactive.get(email, "")
    yeni_satir = [
        satir[ind_sicil] if ind_sicil != -1 else "",
        satir[ind_ad] if ind_ad != -1 else "",
        satir[ind_soyad] if ind_soyad != -1 else "",
        satir[ind_unvan] if ind_unvan != -1 else "",
        satir[ind_email] if ind_email != -1 else "",
        satir[ind_departman] if ind_departman != -1 else "",
        lastactive
    ]
    ws.append(yeni_satir)
wb.save("stajyer_kullanıcılar.xlsx")

# 4. kullanmayan_stajyerler.xlsx'e yaz (On-Demand Usage ve Premium Requests 0 olanlar)
kullanmayan_stajyerler_excel_yaz_json(
    teammembers_json,
    {satir[ind_email].lower() for satir in satirlar},
    "kullanmayan_stajyerler.xlsx"
)

print(f"Toplam {len(satirlar)} kişi stajyer_kullanıcılar.xlsx dosyasına kaydedildi.")
print("Premium Requests ve On-Demand Usage değerleri 0 olanlar kullanmayan_stajyerler.xlsx dosyasına kaydedildi.")