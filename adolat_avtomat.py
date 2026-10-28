"""
ADOLAT (cabinet.sud.uz) - da'vo arizalarini avtomatik tayyorlash va sudga yuborish.

Bitta ishga tushirishda BARCHA mijozlar navbat bilan ishlanadi (7 qadam):
  1 sud va da'vogar -> 2 asosiy ma'lumotlar -> 3 javobgar -> 4 sud xarajatlari ->
  5 hujjatlarni biriktirish -> 6 ko'rib chiqish (saqlash) -> 7 kvitansiyalar -> SUDGA YUBORISH

PAPKALAR (dastur turgan papkada):
  sozlamalar.ini      login, sud, da'vogar va boshqa doimiy sozlamalar
  umumiy_hujjatlar/   hamma ish uchun bir xil hujjatlar (ishonchnoma, oferta,
                      SSP a'zolik guvohnomasi, kompaniya guvohnomasi)
  mijozlar/<MIJOZ>/   har mijozning o'z hujjatlari (da'vo ariza, talabnoma,
                      dalolatnoma, 1C qarzdorlik, pochta kvitansiyasi, pasport)
  ishlar.xlsx         har mijoz uchun bitta qator. Ariza RAQAMI va SANASI shu
                      yerga yoziladi. Qolgani da'vo ariza PDFidan avtomatik olinadi.
  natijalar.csv       har ishning natijasi va havolasi.
  loglar/             log, har sahifaning skrinshoti va tahlili.

ISHLASH TARTIBI (tavsiya):
  1) adolat_avtomat --jadval    -> ishlar.xlsx yaratiladi, raqam/sanani to'ldiring
  2) adolat_avtomat             -> SINOV: hamma ishlar to'ldirilib saytda SAQLANADI,
                                   lekin sudga YUBORILMAYDI (CAPTCHA shu bosqichda)
  3) adolat_avtomat --yuborish  -> saqlangan ishlar "Sudga yuborish" bilan yuboriladi
                                   (CAPTCHA yo'q, har ish oldidan "ha" so'raladi)
  Boshqa: --tekshir (faqat tekshirish), --mijoz FAMILIYA (faqat shu mijoz), --qayta

PAROL: birinchi marta so'raladi, Windows DPAPI bilan shifrlab parol.dat ga saqlanadi va
keyin avtomatik kiritiladi. Almashtirish: --parol
CAPTCHA: brauzerda o'zingiz yechasiz - skript kutadi va o'zi davom etadi.

O'RNATISH:  pip install selenium pymupdf openpyxl opencv-python numpy
"""

import argparse
import configparser
import csv
import difflib
import getpass
import json
import os
import queue
import random
import re
import sys
import tempfile
import threading
import time
import traceback
import urllib.request
from dataclasses import dataclass, field
from datetime import date, datetime

import websocket   # Chrome'ni (chromedriver band bo'lsa ham) darhol yopish - To'xtatish
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchWindowException,
    InvalidSessionIdException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    NoAlertPresentException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

DASTUR_NOMI = "ADOLAT AVTOMAT"   # adolat_avtomat2.py (BUYRUQ tartibida) o'zgartiradi
VERSIYA = "2026-10-28-v33 (kvitansiya 'FAMILIYA - 262...pdf' ham taniladi; kvitansiyasiz/boshqa sudniki - ishlanmaydi)"

# ============================================================
# PAPKALAR VA SOZLAMALAR
# ============================================================
if getattr(sys, "frozen", False):  # PyInstaller exe
    PAPKA = os.path.dirname(os.path.abspath(sys.executable))
    # exe "AdolatAvtomat\" papkasida, ma'lumotlar esa bir pog'ona yuqorida bo'lsa - o'shani olamiz
    _yuqori = os.path.dirname(PAPKA)
    if not os.path.isdir(os.path.join(PAPKA, "mijozlar")) and os.path.isdir(os.path.join(_yuqori, "mijozlar")):
        PAPKA = _yuqori
else:
    PAPKA = os.path.dirname(os.path.abspath(__file__))
UMUMIY_PAPKA = os.path.join(PAPKA, "umumiy_hujjatlar")
MIJOZLAR_PAPKA = os.path.join(PAPKA, "mijozlar")
JADVAL = os.path.join(PAPKA, "ishlar.xlsx")
NATIJALAR = os.path.join(PAPKA, "natijalar.csv")
LOG_PAPKA = os.path.join(PAPKA, "loglar")
SOZLAMALAR_FAYL = os.path.join(PAPKA, "sozlamalar.ini")
PAROL_FAYL = os.path.join(PAPKA, "parol.dat")  # Windows DPAPI bilan shifrlangan

STANDART_SOZLAMALAR = {
    "sayt": "https://cabinet.sud.uz",
    "login": os.getenv("ADOLAT_LOGIN", ""),                      # ONE ID login (yoki sozlamalar.ini)
    "parol": "",
    "profil_stir": os.getenv("ADOLAT_PROFIL_STIR", ""),          # tashkilot STIRi (9 raqam)
    "viloyat": "Тошкент шаҳар",
    "sud": "Shayxontohur",
    "asosiy_davogar": os.getenv("ADOLAT_ASOSIY_DAVOGAR", ""),    # "Da'vogar nomi" ro'yxatidagi tashkilot
    "qoshimcha_davogar_stir": os.getenv("ADOLAT_QOSHIMCHA_DAVOGAR_STIR", ""),
    "qoshimcha_davogar_belgi": "savdo",
    "kichik_biznes": "Ha",
    "nizo_turi": "Фуқаролик ҳуқуқий",
    "ariza_turi": "Соддалаштирилган",
    "ish_turkumi": "Қарз ундиришга доир",
    "qoshimcha_ish_turkumi": "",
    "davlat_boji_imtiyozi": "8 - Давлат божи",   # "8 - Давлат божи тўғрисидаги қонуннинг 8-моддасига асосан озод этилган"
    "kutish_soniya": "40",
    "tezlik": "Tez",
    "pauza_min": "0.2",
    "pauza_max": "0.45",
    "harf_pauza": "0.03",
    "kechikish_koef": "0.5",
    "mijozlar_orasida": "30",
    "ocr_avto": "Ha",
    "ocr_ishonch": "0.90",
    "yuborishdan_oldin_sorash": "Yo'q",
    "xatoda_kutish": "0",
    "sud_hammaga_toshkent": "Yo'q",
}
SOZLAMA_IZOHLARI = """\
; ADOLAT AVTOMAT sozlamalari. Qiymatlar saytdagi matnning BIR QISMI bo'lishi kifoya.
; sayt                    - kabinet manzili
; login                   - ONE ID login
; parol                   - ONE ID paroli. Shu yerga yozsangiz, dastur birinchi ishga tushganda
;                           uni shifrlab parol.dat ga oladi va bu yerdan O'CHIRADI.
;                           Bo'sh qoldirsangiz, bir marta terminalda so'raladi.
; profil_stir             - kirgandan keyin tanlanadigan tashkilot STIRi
; sud_hammaga_toshkent    - Ha: HAMMA mijozlar (qaysi viloyatda yashashidan qat'i nazar) - Toshkent shahar,
;                           "sud" dagi sud (Shayxontohur); Yo'q: da'vo arizada yozilgan sud ("...sudiga"),
;                           arizada sud yozilmagan bo'lsa - javobgar manzili bo'yicha
; kvitansiya_*            - alohida dastur KvitansiyaYaratish.exe sozlamalari (billing.sud.uz kvitansiyalari)
; sud                     - Toshkent shahridagi javobgarlar (har qanday tumani) shu sudga: oferta 8.6 -
;                           sotuvchi yuridik manzili (Shayxontohur). Viloyatlar - javobgar manzili bo'yicha
; viloyat / sud           - sudlar jadvali (Fuqarolik_sudlari_yurisdiksiya.xlsx) yo'q bo'lsa shu tanlanadi
; asosiy_davogar          - "Da'vogar nomi" ro'yxatidan tanlanadigan tashkilot
; qoshimcha_davogar_stir  - STIR orqali qo'shiladigan da'vogar (Savdo-sanoat palatasi)
; qoshimcha_davogar_belgi - STIR qidiruvi natijasida chiqishi kerak bo'lgan so'z
; nizo_turi, ariza_turi, ish_turkumi, qoshimcha_ish_turkumi - 2-qadam ro'yxatlari
; davlat_boji_imtiyozi    - 4-qadam "To'lov imtiyozi" ro'yxatidan tanlanadigan variant
;                           (bo'sh qoldirilsa tanlanmaydi)
; kutish_soniya           - sayt sekin bo'lsa oshiring
; pauza_min / pauza_max   - har bir bosishdan keyingi tasodifiy pauza (soniya), odam kabi
; tezlik                  - Tez / O'rtacha / Sekin - quyidagi 4 ta qiymatni birdaniga o'rnatadi;
;                           Qo'lda - pauza_min, pauza_max, harf_pauza, kechikish_koef qanday yozilgan bo'lsa
; harf_pauza              - yozishda har bir harf orasidagi o'rtacha pauza (soniya)
; kechikish_koef          - sahifa "o'rnashishi" uchun qat'iy kutishlar ko'paytiruvchisi: 1 = eski (sekin),
;                           0.5 = 2 marta tezroq. Sayt ulgurmay xato bersa - 0.7 yoki 1 qiling
; mijozlar_orasida        - bir mijozdan keyingisiga o'tishdagi dam (soniya)
; ocr_avto                - Ha: qo'lda yozilgan raqam/sanani OCR ishonchli o'qisa va paket qoidasi/qo'shni
;                           arizalar ham shuni tasdiqlasa - tasdiqsiz ishlanadi; Yo'q: hammasini odam tasdiqlaydi
; ocr_ishonch             - OCR natijasi avtomatik qabul qilinadigan eng kam ishonch (0.90 = 90%)
; yuborishdan_oldin_sorash - Yo'q: BOSHLASH da har ariza so'ramasdan sudga yuboriladi; Ha: har biri oldidan so'raladi
; xatoda_kutish           - jarayon paytida biror qadam o'tmasa: 0 - hech narsa so'ralmaydi, shu mijoz XATO
;                           (sababi natijaga yoziladi) va darhol keyingisiga o'tiladi; 30 kabi son - shuncha soniya
;                           qo'lda tuzatish kutiladi. (Kirish/SMS, CAPTCHA misoli, E-IMZO - har doim odam kutiladi.)
"""


def sozlamalarni_yuklash():
    cp = configparser.ConfigParser(interpolation=None)
    if os.path.isfile(SOZLAMALAR_FAYL):
        cp.read(SOZLAMALAR_FAYL, encoding="utf-8-sig")
    if not cp.has_section("ADOLAT"):
        cp.add_section("ADOLAT")
    yangi = False
    for k, v in STANDART_SOZLAMALAR.items():
        if not cp.has_option("ADOLAT", k):
            cp.set("ADOLAT", k, v)
            yangi = True
    if yangi:
        with open(SOZLAMALAR_FAYL, "w", encoding="utf-8") as f:
            f.write(SOZLAMA_IZOHLARI)
            cp.write(f)
    return {k: cp.get("ADOLAT", k).strip() for k in STANDART_SOZLAMALAR}


S = dict(STANDART_SOZLAMALAR)
KUTISH = 40


def _son(kalit, standart):
    try:
        return max(0.0, float(str(S.get(kalit, standart)).replace(',', '.')))
    except ValueError:
        return standart


TOXTAT = threading.Event()   # "To'xtatish" tugmasi: jarayon DARHOL to'xtaydi (keyingi bosish/kutishda)


class Toxtatildi(BaseException):
    """'To'xtatish' bosildi. BaseException - oddiy 'except Exception' uni ushlab qolmasligi uchun."""


JORIY_BRAUZER = None   # ishlayotgan Brauzer - To'xtatish uni darhol yopadi


def chrome_darhol_yop(port):
    """Dasturning Chrome'ini debug porti orqali yopadi (Browser.close) - chromedriver buyruq kutayotgan bo'lsa ham."""
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=2) as r:
            ws_url = json.loads(r.read())["webSocketDebuggerUrl"]
        ws = websocket.create_connection(ws_url, timeout=3, suppress_origin=True)
        ws.send(json.dumps({"id": 1, "method": "Browser.close"}))
        ws.close()
        return True
    except Exception:
        return False


def toxtatish():
    """'To'xtatish' tugmasi: jarayon DARHOL to'xtaydi, Chrome yopiladi, hech narsa so'ralmaydi."""
    TOXTAT.set()

    def yop():
        b = JORIY_BRAUZER
        chrome_darhol_yop(CHROME_PORT)       # oynani yopadi (ONE ID sessiyasi profilda saqlanib qoladi)
        if b is not None:
            b.uzish()                        # chromedriver - kutayotgan buyruq ham darhol uziladi
    threading.Thread(target=yop, daemon=True).start()


def toxtat_tekshir():
    if TOXTAT.is_set():
        raise Toxtatildi()


def uxla(soniya):
    """time.sleep, lekin 'To'xtatish' bosilsa darhol uziladi."""
    if TOXTAT.wait(max(0.0, soniya)):
        raise Toxtatildi()


def tin(soniya):
    """Sahifa o'rnashishi uchun qat'iy kutish - sozlamalar: kechikish_koef (1 = eski tezlik, 0.5 = 2 barobar tez)."""
    uxla(soniya * min(2.0, max(0.1, _son("kechikish_koef", 0.5))))


def odamdek_pauza(kop=1.0):
    """Saytni bezovta qilmaslik uchun odam tezligidagi tasodifiy pauza."""
    a, b = _son('pauza_min', 0.7), _son('pauza_max', 1.8)
    uxla(random.uniform(min(a, b), max(a, b)) * kop)


def harf_pauza():
    h = _son('harf_pauza', 0.12)
    uxla(random.uniform(h * 0.5, h * 1.5))

# (kalit, nomi, fayl nomidagi so'zlar, saytdagi "Hujjat turi" variantlari)
# Tartib = 5-qadamdagi yuklash tartibi: SSP, ishonchnoma, oferta, shartnoma, talabnoma, pochta kvitansiya, 1C, keyin qolganlari.
# Fayl turini aniqlashda ham shu tartibda tekshiriladi (birinchi mos kelgani olinadi).
HUJJAT_TURLARI = [
    ("davo", "Da'vo arizasi", ["даъво", "davo", "da'vo", "da’vo", "даво"], []),
    ("ssp", "SSP a'zolik guvohnomasi va shartnomasi", ["ссп", "ssp", "аъзолик", "azolik", "a'zolik"], ["бошқа"]),
    ("ishonchnoma", "Ishonchnoma", ["ishonchnoma", "ишончнома", "доверен"], ["ишончнома", "бошқа"]),
    ("oferta", "Oferta", ["оферта", "oferta"], ["бошқа"]),
    ("dalolatnoma", "Dalolatnoma / shartnoma", ["шартнома", "shartnoma", "dalolatnoma", "далолатнома"], ["бошқа"]),
    ("talabnoma", "Talabnoma", ["talabnoma", "талабнома"], ["бошқа"]),
    ("pochta", "Pochta kvitansiyasi", ["почта", "pochta"], ["почта харажати", "бошқа"]),
    ("qarzdorlik", "1C qarzdorlik ma'lumoti", ["1с", "1c", r"^1\s+[сc]\s", "карздорлик", "қарздорлик", "qarzdorlik"], ["бошқа"]),
    ("guvohnoma", "Kompaniya guvohnomasi", ["guvohnoma", "гувоҳнома", "гувохнома"], ["бошқа"]),
    ("pasport", "Javobgar pasporti", ["pasport", "паспорт", "passport"], ["бошқа"]),
    # raqamli to'lov kvitansiyasi (262000000001.pdf) 5-qadamda YUKLANMAYDI - 4-qadamda '+' orqali raqami biriktiriladi
    # "262000000002.pdf" yoki KvitansiyaYaratish nomi "FAMILIYA_ISM - 262000000002.pdf" (oxirida 12 xonali raqam)
    ("tolov", "To'lov kvitansiyasi (raqami 4-qadamda)", [r"^\d{8,}$", r"^(?:.*\D)?\d{12}$", "kvitans", "квитан"], []),
]
YUKLANMAYDIGAN = ("davo", "tolov")   # 5-qadam o'ng tomoniga yuklanmaydi (da'vo - chapga, to'lov - 4-qadamda)
HUJJAT_KENGAYTMALARI = (".pdf", ".jpg", ".jpeg", ".png")
JADVAL_USTUNLAR = ["Papka", "Ariza raqami", "Ariza sanasi", "Javobgar", "JShShIR",
                   "Asosiy qarz", "Pochta xarajati", "Izoh", "Sud",   # "Sud" - bo'sh bo'lsa manzildan aniqlanadi
                   "Tekshirildi"]   # "ha" - ariza raqami/sanasi odam tomonidan tasdiqlangan (OCR o'qigani tasdiqlanmagan);
#                                     "ha!" - dastur boshqa qiymat kutgan bo'lsa ham odam aynan shuni tasdiqlagan


def tekshirilganmi(d):
    """ishlar.xlsx qatori: raqam/sana odam tomonidan tasdiqlanganmi ("ha" yoki "ha!").
    Eski jadvalda bu ustun yo'q - u yerdagi raqam/sana avval qabul qilingan, tasdiqlangan hisoblanadi."""
    if "Tekshirildi" in d:
        return str(d.get("Tekshirildi") or "").strip().lower() in ("ha", "ha!")
    return bool(str(d.get("Ariza raqami") or "").strip() and str(d.get("Ariza sanasi") or "").strip())


# ============================================================
# LOG
# ============================================================
class _Tee:
    def __init__(self, *streams):
        self.streams = streams

    def write(self, data):
        for s in self.streams:
            try:
                s.write(data)
                s.flush()
            except Exception:
                pass

    def flush(self):
        for s in self.streams:
            try:
                s.flush()
            except Exception:
                pass


SESSIYA = datetime.now().strftime("%Y%m%d_%H%M%S")


def log_boshlash(*qoshimcha):
    os.makedirs(LOG_PAPKA, exist_ok=True)
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    yol = os.path.join(LOG_PAPKA, f"{SESSIYA}.log")
    f = open(yol, "w", encoding="utf-8")
    sys.stdout = _Tee(sys.stdout, f, *qoshimcha)
    sys.stderr = _Tee(sys.stderr, f, *qoshimcha)
    return yol


# ============================================================
# ODAM BILAN ISHLASH (CAPTCHA, SMS, tasdiqlash)
# ============================================================
_kirish_navbati = queue.Queue()
_oquvchi_ishga_tushdi = False
_stdin_tugadi = False  # terminal yopiq / kiritish yo'q - Enter kutib qotib qolmaslik uchun


def _stdin_oqish():
    global _stdin_tugadi
    try:
        for line in sys.stdin:
            _kirish_navbati.put(line.rstrip("\r\n"))
    except Exception:
        pass
    _stdin_tugadi = True


def oquvchini_boshlash():
    global _oquvchi_ishga_tushdi
    if not _oquvchi_ishga_tushdi:
        threading.Thread(target=_stdin_oqish, daemon=True).start()
        _oquvchi_ishga_tushdi = True


def signal():
    try:
        import winsound
        winsound.Beep(1200, 250)
        winsound.Beep(900, 250)
    except Exception:
        print("\a", end="")


ODAM_ISHLOVCHI = None   # oyna rejimida gui_main o'rnatadi
PAROL_SORASH = None     # oyna rejimida gui_main o'rnatadi
ODAM_KUTISH = None      # mijozlar ishlanayotganda (ishga_tushirish) - sozlamalar: xatoda_kutish (0 - so'ramaslik)


class OdamKelmadi(Exception):
    """Jarayon paytida biror qadam o'tmadi (qo'lda tuzatish so'ralmaydi yoki javob bo'lmadi) - mijoz o'tkazib yuboriladi."""


def odam(xabar, shart=None, vaqt=None, cheksiz=False):
    """Odamni kutadi: Enter (yoki matn) yozilsa YOKI shart() rost bo'lsa qaytadi.
    Mijozlar ishlanayotganda (cheksiz=False) - ko'pi bilan ODAM_KUTISH soniya, keyin OdamKelmadi:
    odam yo'q bo'lsa navbat to'xtab qolmaydi, shu mijoz XATO bo'lib keyingisiga o'tiladi.
    cheksiz=True - odamsiz iloji yo'q joylar (kirish/SMS, E-IMZO imzo, CAPTCHA misoli)."""
    avto = vaqt is None and not cheksiz and ODAM_KUTISH is not None
    if avto and ODAM_KUTISH <= 0:      # so'ramasdan - shu mijoz XATO, darhol keyingisiga
        sabab = " ".join(q.strip() for q in xabar.strip().splitlines()[:2])
        print(f"    !!! {sabab}\n    -> so'ralmaydi: shu mijoz XATO, keyingi mijozga o'tiladi")
        raise OdamKelmadi(sabab)
    if avto:
        vaqt = ODAM_KUTISH
        xabar = xabar.rstrip() + f"\n(Javob bo'lmasa {vaqt:.0f} soniyadan keyin bu mijoz o'tkazib yuboriladi.)"
    j = _odam_kut(xabar, shart, vaqt)
    if j is None and avto:
        raise OdamKelmadi("qo'lda tuzatish kutildi, javob bo'lmadi: " + xabar.strip().splitlines()[0])
    return j


def _odam_kut(xabar, shart, vaqt):
    if ODAM_ISHLOVCHI:
        return ODAM_ISHLOVCHI(xabar, shart, vaqt)
    oquvchini_boshlash()
    while not _kirish_navbati.empty():
        _kirish_navbati.get_nowait()
    signal()
    print("\n    " + "=" * 62)
    for qator in xabar.strip().splitlines():
        print("    " + qator)
    print("    " + "=" * 62)
    boshlandi = time.time()
    while True:
        try:
            return _kirish_navbati.get(timeout=0.5)
        except queue.Empty:
            pass
        toxtat_tekshir()
        if shart:
            try:
                if shart():
                    print("    OK -> (skript o'zi sezdi, davom etilmoqda)")
                    return ""
            except Exception:
                pass
        elif _stdin_tugadi:
            print("    (terminaldan kiritish yo'q - javobsiz davom etilmoqda)")
            return ""
        if vaqt and time.time() - boshlandi > vaqt:
            return None


class Topilmadi(Exception):
    pass


class ParolXato(Exception):
    pass


class ShaxsTopilmadi(Exception):
    """Sayt JShShIR bo'yicha shaxsni topmadi (vafot etgan, JShShIR xato) - qayta urinilmaydi, keyingi mijozga o'tiladi."""


class KvitansiyaXato(Exception):
    """Sayt to'lov kvitansiyasini qabul qilmadi - shu mijoz XATO (kvitansiya), qayta urinilmaydi, keyingisiga o'tiladi."""


class CaptchaMisol(Exception):
    """Sayt matematik misolli CAPTCHA so'radi - shu mijoz sahifa yangidan ochilib qaytadan boshlanadi."""


class SaytBand(Exception):
    """Sayt "Тизим янгиланмади. Кутиб туринг." dedi - keyinroq urinish kerak."""


SAYT_BAND_XABAR = ("Sayt «Тизим янгиланмади. Кутиб туринг.» dedi - kabinet hozir kirishni qabul qilmayapti.\n"
                   "30 daqiqa - 1 soatdan keyin qayta urinib ko'ring.")


# ============================================================
# PAROL: bir marta so'raladi, Windows DPAPI bilan shifrlab saqlanadi
# (faqat shu Windows foydalanuvchisi ochib o'qiy oladi)
# ============================================================
def _dpapi(data, shifrlash):
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    buf = ctypes.create_string_buffer(data, len(data))
    kirish = BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    chiqish = BLOB()
    crypt32 = ctypes.windll.crypt32
    f = crypt32.CryptProtectData if shifrlash else crypt32.CryptUnprotectData
    if not f(ctypes.byref(kirish), None, None, None, None, 0x01, ctypes.byref(chiqish)):
        raise OSError("DPAPI xatosi")
    try:
        return ctypes.string_at(chiqish.pbData, chiqish.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(ctypes.cast(chiqish.pbData, ctypes.c_void_p))


def parol_saqlash(parol):
    try:
        with open(PAROL_FAYL, "wb") as f:
            f.write(_dpapi(parol.encode("utf-8"), True))
        print("    OK -> Parol shifrlangan holda saqlandi (keyingi safar so'ralmaydi)")
    except Exception as e:
        print(f"    (parolni saqlab bo'lmadi: {e})")


def parol_ochirish():
    if os.path.isfile(PAROL_FAYL):
        os.remove(PAROL_FAYL)
        print("    (saqlangan parol o'chirildi)")


def ini_paroldan_olish():
    """sozlamalar.ini dagi parolni shifrlangan parol.dat ga o'tkazadi va ini dan o'chiradi."""
    if not S.get("parol"):
        return
    parol_saqlash(S["parol"])
    S["parol"] = ""
    try:
        with open(SOZLAMALAR_FAYL, encoding="utf-8-sig") as f:
            matn = f.read()
        matn = re.sub(r"(?m)^(\s*parol\s*=).*$", lambda m: m.group(1) + " ", matn)
        with open(SOZLAMALAR_FAYL, "w", encoding="utf-8") as f:
            f.write(matn)
        print("    OK -> Parol sozlamalar.ini dan olib tashlandi (endi faqat shifrlangan holda saqlanadi)")
    except Exception as e:
        print(f"    !!! sozlamalar.ini dan parolni o'chirib bo'lmadi: {e} - uni qo'lda o'chiring!")


def parol_olish(yangi=False):
    """ADOLAT_PAROL muhit o'zgaruvchisi -> saqlangan parol.dat -> bir marta so'rash va saqlash."""
    if os.environ.get("ADOLAT_PAROL"):
        return os.environ["ADOLAT_PAROL"]
    if yangi:
        parol_ochirish()
    if os.path.isfile(PAROL_FAYL):
        try:
            with open(PAROL_FAYL, "rb") as f:
                p = _dpapi(f.read(), False).decode("utf-8")
            print("    OK -> Saqlangan parol ishlatiladi")
            return p
        except Exception:
            print("    (saqlangan parolni ochib bo'lmadi - qaytadan so'raladi)")
    if PAROL_SORASH:
        p = PAROL_SORASH()
    else:
        p = getpass.getpass(f"ONE ID paroli ({S['login']}), bir marta so'raladi (ekranda ko'rinmaydi): ")
    if p:
        parol_saqlash(p)
    return p


# ============================================================
# MA'LUMOTLAR: hujjatlar, PDF, jadval, natijalar
# ============================================================
@dataclass
class Ish:
    papka: str
    yol: str
    raqam: str = ""
    sana: date = None
    javobgar: str = ""
    jshshir: str = ""
    pasport: str = ""
    asosiy_qarz: str = ""
    pochta: str = ""
    tekshirildi: bool = False  # raqam/sana tasdiqlangan: odam ("ha") yoki OCR ishonchli o'qigan (ocr_avto)
    manzil: str = ""          # javobgar manzili (da'vo arizadan)
    sud: dict = None          # sudlar() dan: javobgar manzili bo'yicha (yoki ishlar.xlsx "Sud" ustunidan)
    hujjatlar: list = field(default_factory=list)  # [(kalit, nomi, yol)]
    ariza_pdf: str = None
    xatolar: list = field(default_factory=list)
    ogohlantirishlar: list = field(default_factory=list)
    qolyozma: dict = None       # qo'lda yozilgan raqam/sana nomzodlari (qolyozma_tahlil)
    ocr_ishonch: tuple = None   # (raqam, sana) ishonchi 0..1
    ocr_variantlar: dict = None  # {"raqam": [(qiymat, ehtimol)...], "sana": [...]} - oynada ko'rsatish uchun
    ocr_avto: bool = False      # raqam/sana OCR bilan avtomatik qabul qilindi
    ocr_izoh: str = ""
    tekshir_kerak: bool = False  # raqam/sanani odam ko'rishi kerak - ro'yxat boshida turadi
    raqam_xato: str = ""        # xatolar ichidagi raqam/sana xabari (qayta baholashda almashtiriladi)
    jadvaldagi: tuple = ("", None, False)   # ishlar.xlsx dagi (raqam, sana, tasdiqlangan)
    majburiy: bool = False      # "ha!" - odam dalillarga qaramay tasdiqlagan, qayta tekshirilmaydi
    ocr_tavsiya: tuple = None   # (raqam, sana, kuchli, raqam_zid, sana_zid) - dalillar bo'yicha eng ehtimolli qiymat;
    #                             zid - OCR o'zi boshqa qiymat o'qigan yoki jadvaldagi tasdiq shubhali
    jadval_shubha: tuple = None  # tasdiqlangan (raqam, sana) dalillarga zid - qayta tekshirishga qaytarilgan


def hujjat_turi(fayl_nomi):
    stem = os.path.splitext(os.path.basename(fayl_nomi))[0].lower()
    for kalit, _nomi, sozlar, _ in HUJJAT_TURLARI:
        for s in sozlar:
            if s.startswith("^"):
                if re.match(s, stem):
                    return kalit
            elif s in stem:
                return kalit
    return "boshqa"


def papka_fayllari(papka):
    """Mijoz papkasidagi fayllar [(nom, yo'l)] - ICHKI papkalar bilan: Windows "Извлечь всё" ZIP ni
    papka ichida papka qilib ochadi (mijoz/mijoz/fayllar). "_" / "." bilan boshlanadigan xizmat papkalari,
    Word vaqtinchalik fayllari (~$) va bir xil faylning takrori (nomi va hajmi bir xil) olinmaydi."""
    if not os.path.isdir(papka):
        return []
    natija, korildi = [], set()
    for ildiz, papkalar, fayllar in os.walk(papka):
        papkalar[:] = sorted(x for x in papkalar if not x.startswith(("_", ".")))
        for f in sorted(fayllar):
            yol = os.path.join(ildiz, f)
            try:
                kalit = (f.lower(), os.path.getsize(yol))
            except OSError:
                if len(yol) >= 255:
                    _uzun_yol_ogohlantirish(yol)
                continue
            if f.startswith("~$") or kalit in korildi:
                continue
            korildi.add(kalit)
            natija.append((f, yol))
    return natija


_UZUN_YOLLAR = set()


def _uzun_yol_ogohlantirish(yol):
    if yol not in _UZUN_YOLLAR:
        _UZUN_YOLLAR.add(yol)
        print(f"    !!! Fayl yo'li juda uzun ({len(yol)} belgi, Windows 260 dan uzunini o'qimaydi) - "
              f"papka nomini qisqartiring: {yol}")


def _uz(yol):
    """Windows uzun yo'l prefiksi (260 belgidan uzun yo'llar bilan ko'chirish uchun)."""
    if os.name != "nt" or yol.startswith("\\\\?\\"):
        return yol
    return "\\\\?\\" + os.path.abspath(yol)


def ichki_papkalarni_chiqarish():
    """Windows "Извлечь всё" ZIP ni mijoz/mijoz/fayllar qilib ochadi: mijoz papkasida fayl yo'q, faqat bitta
    ichki papka bo'lsa - uning ichidagilar bir pog'ona yuqoriga olinadi (yo'l ~55 belgiga qisqaradi;
    Windows 260 belgidan uzun yo'lni o'qimaydi). Hech narsa o'chirilmaydi; nomi bir xil fayl bo'lsa - tegilmaydi."""
    xizmat = {"desktop.ini", "thumbs.db"}
    for p in mijoz_papkalari():
        yol = os.path.join(MIJOZLAR_PAPKA, p)
        try:
            ichida = [x for x in os.listdir(_uz(yol)) if x.lower() not in xizmat]
        except OSError:
            continue
        papkalar = [x for x in ichida if os.path.isdir(_uz(os.path.join(yol, x)))]
        if len(papkalar) != 1 or len(ichida) != 1:
            continue
        ichki = os.path.join(yol, papkalar[0])
        try:
            elementlar = os.listdir(_uz(ichki))
            band = [x for x in elementlar if os.path.exists(_uz(os.path.join(yol, x)))]
            if band:
                print(f"    !!! {p}: ichki papka chiqarilmadi - bir xil nomli fayl bor: {band[0]}")
                continue
            for x in elementlar:
                os.rename(_uz(os.path.join(ichki, x)), _uz(os.path.join(yol, x)))
            os.rmdir(_uz(ichki))
            print(f"    OK -> {p}: hujjatlar ichki papkadan chiqarildi ({len(elementlar)} ta)")
        except OSError as e:
            print(f"    !!! {p}: ichki papka chiqarilmadi: {e}")


def papka_hujjatlari(papka):
    r = []
    for f, y in papka_fayllari(papka):
        if not f.lower().endswith(HUJJAT_KENGAYTMALARI):
            continue
        k = hujjat_turi(f)
        # nomidan bilinmagan PDF ichida billing.sud.uz kvitansiyasi ("ХОС РАҚАМ") bo'lsa - to'lov kvitansiyasi:
        # u 4-qadamda raqami bilan biriktiriladi (aks holda "boshqa" bo'lib biriktirilmay qolardi - sud rad etadi)
        if k == "boshqa" and f.lower().endswith(".pdf"):
            try:
                if kvitansiya_oqish(y):
                    k = "tolov"
            except Exception:
                pass
        r.append((k, y))
    return r


def mijoz_papkalari():
    """mijozlar ichidagi mijoz papkalari ("_" yoki "." bilan boshlanganlari - xizmat papkalari, olinmaydi)."""
    if not os.path.isdir(MIJOZLAR_PAPKA):
        return []
    return sorted(p for p in os.listdir(MIJOZLAR_PAPKA)
                  if os.path.isdir(os.path.join(MIJOZLAR_PAPKA, p)) and not p.startswith(("_", ".")))


def ziplarni_ochish():
    """mijozlar ga tashlangan har bir ZIP (bitta mijoz) ni shu nomdagi papkaga ochadi,
    ZIP ning o'zini mijozlar/_zip_arxiv ga oladi. Ochilgan papkalar ro'yxatini qaytaradi."""
    import shutil
    import zipfile
    ochildi = []
    if not os.path.isdir(MIJOZLAR_PAPKA):
        return ochildi
    for f in sorted(os.listdir(MIJOZLAR_PAPKA)):
        yol = os.path.join(MIJOZLAR_PAPKA, f)
        if not (f.lower().endswith(".zip") and os.path.isfile(yol)):
            continue
        nom = os.path.splitext(f)[0].strip()
        maqsad = os.path.join(MIJOZLAR_PAPKA, nom)
        try:
            with zipfile.ZipFile(yol) as z:
                soni = 0
                for i in z.infolist():
                    if i.is_dir():
                        continue
                    n = i.filename
                    if not (i.flag_bits & 0x800):          # eski ZIP: kirill nomlar cp866 da
                        try:
                            n = n.encode("cp437").decode("cp866")
                        except (UnicodeEncodeError, UnicodeDecodeError):
                            pass
                    fayl = os.path.basename(n.replace("\\", "/"))
                    if not fayl:
                        continue
                    os.makedirs(maqsad, exist_ok=True)
                    with z.open(i) as manba, open(os.path.join(maqsad, fayl), "wb") as chiq:
                        shutil.copyfileobj(manba, chiq)
                    soni += 1
            arxiv = os.path.join(MIJOZLAR_PAPKA, "_zip_arxiv")
            os.makedirs(arxiv, exist_ok=True)
            shutil.move(yol, os.path.join(arxiv, f))
            ochildi.append(nom)
            print(f"    OK -> ZIP ochildi: {f} -> mijozlar/{nom} ({soni} ta fayl)")
        except Exception as e:
            print(f"    !!! ZIP ochilmadi: {f}: {type(e).__name__}: {e}")
    return ochildi


FITZ_QULF = threading.RLock()   # PyMuPDF bir vaqtda bir nechta oqimda ishlamaydi (fon OCR + jadval)
_PDF_MATN = {}


def pdf_matn(yol):
    try:
        st = os.stat(yol)
        kalit = (yol, st.st_size, st.st_mtime_ns)
    except OSError:
        kalit = None
    if kalit in _PDF_MATN:          # 600 ta mijozda jadval har yangilanganda PDF qayta o'qilmasin
        return _PDF_MATN[kalit]
    try:
        import pymupdf as fitz
    except ImportError:
        try:
            import fitz
        except ImportError:
            print("    (PyMuPDF yo'q: pip install pymupdf)")
            return ""
    try:
        with FITZ_QULF, fitz.open(yol) as d:
            matn = "\n".join(p.get_text() for p in d)
    except Exception as e:
        print(f"    (PDF o'qishda xato {os.path.basename(yol)}: {e})")
        return ""
    matn = matn_tuzat(matn)
    if kalit:
        _PDF_MATN[kalit] = matn
    return matn


def matn_tuzat(t):
    """Shrift kodirovkasi noto'g'ri PDF: kirill (cp1251) latin-1 deb chiqqan bo'lsa ('Òîøêåíò' -> 'Тошкент')."""
    if t and sum(1 for c in t if "\u00c0" <= c <= "\u00ff") > len(t) * 0.15:
        try:
            u = t.encode("cp1252", errors="strict").decode("cp1251")
            if sum(1 for c in u if "\u0400" <= c <= "\u04ff") > len(u) * 0.2:
                return u
        except UnicodeError:
            pass
    return t


def matnsizmi(matn):
    """Da'vo ariza PDF matni yo'q yoki o'qib bo'lmaydi (skaner, noto'g'ri shrift) - unda OCR va boshqa hujjatlar."""
    t = (matn or "").strip()
    if len(t) < 50:
        return True
    harflar = [c for c in t if c.isalpha()]
    yaxshi = sum(1 for c in harflar if "a" <= c.lower() <= "z" or "\u0400" <= c <= "\u04ff")
    if not harflar or yaxshi / len(harflar) < 0.85:
        return True
    # matn bor, lekin da'vo arizaning birorta tanish so'zi yo'q - shrift buzuq
    return not re.search(r"(?i)javobgar|жавобгар|ответчик|da.?vo|даъво|даво|иск|ariza|ариза", t)


def docx_matn(yol):
    """Word (ДАЪВО АРИЗА.docx) matni - kompyuterda yozilgan, PDF skanidan ancha toza."""
    try:
        import docx
        d = docx.Document(yol)
    except Exception as e:
        print(f"    (DOCX o'qib bo'lmadi {os.path.basename(yol)}: {e})")
        return ""
    korildi, qismlar = set(), []
    for t in d.tables:
        for r in t.rows:
            for c in r.cells:
                x = c.text.strip()
                if x and x not in korildi:
                    korildi.add(x)
                    qismlar.append(x)
    for par in d.paragraphs:
        x = par.text.strip()
        if x and x not in korildi:
            korildi.add(x)
            qismlar.append(x)
    return "\n".join(qismlar)


def summa_tozalash(s):
    if s is None:
        return ""
    s = str(s).replace("\xa0", "").replace(" ", "").replace(",", ".").strip()
    try:
        return f"{float(s):.2f}" if s else ""
    except ValueError:
        return ""


SUMMA = r"(\d[\d \xa0]*\d(?:[.,]\d{1,2})?|\d)"   # 2 345 000 | 1 234 567,89 | 41 200,00


def matndan_malumot(m):
    """Da'vo ariza matnidan: javobgar FIO, JShShIR, pasport, asosiy qarz, pochta."""
    r = {}
    x = re.search(r"JSH\s*SH\s*IR\s*[:：]?\s*(\d{14})", m, re.I) or re.search(r"ЖШШИР\s*[:：]?\s*(\d{14})", m, re.I)
    if x:
        r["jshshir"] = x.group(1)
    x = re.search(r"(?:Javobgar|Жавобгар)\s*:\s*\n?\s*([^\n]+)", m, re.I)   # lotin yoki kirill ariza
    if x:
        fio = x.group(1).strip()
        qatorlar = [q.strip() for q in m[x.end():].splitlines() if q.strip()]
        # FIO ikki qatorga bo'lingan: "ФАМИЛИЯ ИСМ ОТАСИНИНГ_ИСМИ" + "ЎҒЛИ"
        if qatorlar and re.fullmatch(r"(?i)(o.?g.?li|qizi|ўғли|қизи|угли|кизи|оглы|кызы)\.?", qatorlar[0]):
            fio += " " + qatorlar.pop(0)
        r["javobgar"] = fio
        # manzil: FIO dan keyingi qatorlar - "Pasport seriyasi" / JShShIR / Tel gacha
        manzil = []
        for q in qatorlar:
            if re.match(r"(?i)(pasport|passport|паспорт|jsh|жшш|tel|тел|da.?vo ariza|даъво)", q) or len(manzil) >= 4:
                break
            manzil.append(q)
        if manzil:
            r["manzil"] = " ".join(manzil)
    # ariza kimga: "Fuqarolik ishlari bo'yicha Shayxontohur tumanlararo sudiga" (faqat boshidagi - sarlavha qismi)
    x = re.search(r"(?:Fuqarolik\s+ishlari\s+bo\S*yicha|Фуқаролик\s+ишлари\s+бўйича)\s+(.{3,80}?)\s+(?:sudiga|судига)\b",
                  m[:2500], re.I | re.S)
    if x:
        r["sud_nomi"] = re.sub(r"\s+", " ", x.group(1)).strip() + (" sudi" if re.search(r"[a-z]", x.group(1), re.I) else " суди")
    x = re.search(r"(?:Pasport|Паспорт)\s*(?:seriyasi|серияси)\s*:?\s*([A-ZА-Я]{2}\s?\d{7})", m, re.I)
    if x:
        r["pasport"] = x.group(1).replace(" ", "")
    som = r"\s*(?:so.?m|сўм|сум)\s+"
    x = (re.search(SUMMA + som + r"(?:asosiy\s+qarz|асосий\s+қарз)", m, re.I)
         or re.search(SUMMA + som + r"(?:miqdoridagi\s+qarzdorlik|миқдоридаги\s+қарздорлик)", m, re.I))
    if x:
        r["asosiy_qarz"] = summa_tozalash(x.group(1))
    x = re.search(SUMMA + som + r"(?:pochta|почта)", m, re.I)
    if x:
        r["pochta"] = summa_tozalash(x.group(1))
    return r


def davo_fayllari(papka):
    """Mijoz papkasidagi da'vo ariza: (PDF, DOCX) - qaysi biri bo'lmasa None."""
    pdf = docx_yol = None
    for f, yol in papka_fayllari(papka):
        if hujjat_turi(f) != "davo":
            continue
        if f.lower().endswith(".pdf") and not pdf:
            pdf = yol
        elif f.lower().endswith(".docx") and not docx_yol:
            docx_yol = yol
    return pdf, docx_yol


# ============================================================
# SUDNI ANIQLASH: javobgar manzili + Fuqarolik_sudlari_yurisdiksiya.xlsx (hudud -> sud -> tuman/shaharlar)
# ============================================================
# ц -> k: skaner "қ" ni "ц" deb o'qiydi (К,ашцадарё)
_KIRILL = {"а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo", "ж": "j", "з": "z", "и": "i",
           "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t",
           "у": "u", "ф": "f", "х": "h", "ц": "k", "ч": "ch", "ш": "sh", "щ": "sh", "ъ": "", "ы": "i", "ь": "",
           "э": "e", "ю": "yu", "я": "ya", "ў": "o", "қ": "k", "ғ": "g", "ҳ": "h"}
_SHAHAR_SOZ = {"shahar", "shahri", "shahr", "g", "gorod", "goroda"}
_TUMAN_SOZ = {"tumani", "tuman", "rayon", "rayoni", "rayona", "rayonu"}
_VILOYAT_SOZ = {"viloyati", "viloyat", "oblast", "oblasti", "obl", "vil", "respublikasi", "respublika", "resp"}
# rus va o'zbek yozilishi skeletda farq qiladiganlar: Каракалпакстан, Джизак
_SKELET_ALIAS = (("krklpkstn", "krklpgstn"), ("jzk", "jzh"))


def lotin(s):
    """Kirill (o'zbek/rus) va lotinni bitta oddiy lotinga: 'Шайхонтоҳур', 'Shayxontohur' -> 'shayhontohur'."""
    s = re.sub(r"[`'’‘ʻʼ\"«»]", "", str(s or "").lower())
    s = "".join(_KIRILL.get(ch, ch) for ch in s)
    return s.replace("x", "h").replace("q", "k").replace("w", "v").replace("dzh", "j").replace("dj", "j").replace("zh", "j")


def _sozlar(s):
    return re.findall(r"[a-z]+", lotin(s))


def _skelet(soz):
    """Unlilarsiz, qo'sh harflarsiz: 'Andijon' va 'Андижан' -> 'ndjn' (rus/o'zbek yozilishi farqi yo'qoladi)."""
    s = re.sub(r"(.)\1+", r"\1", re.sub(r"[aeiouy]", "", soz))
    for a, b in _SKELET_ALIAS:
        if s.startswith(a):
            return b + s[len(a):]
    return s


def _mos(w, p):
    """Manzildagi so'z skeleti w joy nomi skeleti p ga mosmi (rus qo'shimchasi: Андижанская, Шайхантахурский)."""
    return bool(p) and (w == p or (w.startswith(p) and (len(p) >= 4 or w[len(p):].startswith(("sk", "nsk")))))


def _joy_nomi(nom):
    """'Андижон шаҳар' -> (['ndjn'], 'shahar'); 'Мирзо-Улуғбек тумани' -> (['mrz', 'lgbk'], 'tuman')."""
    w = _sozlar(nom)
    tur = None
    if w and w[-1] == "sudi":
        w = w[:-1]
    if w and w[-1] in ("shahar", "shahri"):
        tur, w = "shahar", w[:-1]
    elif w and w[-1] in ("tumani", "tuman"):
        tur, w = "tuman", w[:-1]
    elif w and w[-1] == "tumanlararo":
        tur, w = "tumanlararo", w[:-1]
    elif w and w[-1] in ("viloyati", "respublikasi"):
        tur, w = "viloyat", w[:-1]
    return [_skelet(x) for x in w], tur


def _oraliqlar(sk, nomi_sk):
    """sk (so'zlar skeleti) ichida nomi_sk ketma-ketligi uchragan joylar: [(i, j)]."""
    n = len(nomi_sk)
    return [(i, i + n - 1) for i in range(len(sk) - n + 1) if n and all(_mos(sk[i + k], nomi_sk[k]) for k in range(n))]


def _soz_turi(w):
    return "shahar" if w in _SHAHAR_SOZ else "tuman" if w in _TUMAN_SOZ else "viloyat" if w in _VILOYAT_SOZ else None


SUDLAR_FAYL = os.path.join(PAPKA, "Fuqarolik_sudlari_yurisdiksiya.xlsx")
_SUDLAR = None


def sudlar():
    """Jadvaldagi sudlar: [{'hudud': 'Андижон вилояти', 'nomi': 'Андижон туманлараро суди', 'joylar': [...]}]."""
    global _SUDLAR
    if _SUDLAR is None:
        _SUDLAR = []
        if os.path.isfile(SUDLAR_FAYL):
            try:
                import openpyxl
                ws = openpyxl.load_workbook(SUDLAR_FAYL, data_only=True, read_only=True).active
                for r in ws.iter_rows(min_row=2, values_only=True):
                    if r and len(r) >= 5 and r[1] and r[2]:
                        _SUDLAR.append({"hudud": str(r[1]).strip(), "nomi": str(r[2]).strip(),
                                        "joylar": [j.strip() for j in str(r[4] or "").split(",") if j.strip()]})
            except Exception as e:
                print(f"    (sudlar jadvali o'qilmadi: {e})")
    return _SUDLAR


def _toshkent_shahri(hudud):
    return hudud.startswith("Тошкент") and _joy_nomi(hudud)[1] == "shahar"


def toshkent_sudi():
    """Toshkent shahridagi HAMMA javobgarlar uchun bitta sud - oferta 8.6: nizo sotuvchining yuridik manzili
    bo'yicha (Shayxontohur tumani). sozlamalar.ini: sud = Shayxontohur."""
    t = sud_nomidan(S.get("sud") or "Shayxontohur")
    return t if t and _toshkent_shahri(t["hudud"]) else sud_nomidan("Shayxontohur")


def hammaga_toshkent():
    """Sozlamalarda "hamma mijoz - Toshkent shahar, Shayxontohur" belgilanganmi."""
    return str(S.get("sud_hammaga_toshkent", "Yo'q")).strip().lower() in ("ha", "yes", "1", "true")


def avto_sud(manzil):
    """Avtomatik sud: belgi qo'yilgan bo'lsa - hamma mijoz toshkent_sudi(), aks holda javobgar manzili bo'yicha."""
    if hammaga_toshkent() and toshkent_sudi():
        return toshkent_sudi(), "sozlamalar: hamma mijoz - Toshkent shahar"
    return sud_aniqlash(manzil)


# shu so'zlardan OLDINGI nom - mahalla/ko'cha/qishloq nomi, tuman emas ("ДЕХҚОНОБОД МСГ" - Dehqonobod tumani emas)
MAHALLA_SOZLARI = {"mfy", "msg", "kfy", "qfy", "ssg", "sfy", "mahalla", "mahallasi", "mahallya", "kochasi", "kuchasi",
                   "kucha", "kuchas", "kishlogi", "qishlogi", "kishlok", "qishloq", "massivi", "mavzesi", "mavze",
                   "berk", "tor", "street", "ulitsa", "proezd", "tupik", "dom", "uy"}
OLD_SOZLARI = {"ul", "ulitsa", "mas", "massiv", "mavze", "kv", "proezd", "tupik", "mfy", "mahalla"}
MANZIL_TOLIQ_RE = re.compile(r"(?i)вилоят|viloyat|област|обл\.|шаҳр|шахр|shahr|туман|tuman|район|rayon|"
                             r"республик|respublik|\bг\.|\bш\.|\bsh\.")


def manzil_toliqmi(manzil):
    """Manzilda viloyat / tuman / shahar so'zi bormi (aks holda faqat mahalla-ko'cha - sud aniqlab bo'lmaydi)."""
    return bool(manzil and MANZIL_TOLIQ_RE.search(manzil))


def sud_aniqlash(manzil):
    """Javobgar manzilidan sud: (sud, izoh) yoki (None, sabab).
    Toshkent shahri (qaysi tumani bo'lmasin) - har doim toshkent_sudi(); viloyatlar - tuman/shahar bo'yicha."""
    sl = sudlar()
    if not sl:
        return None, "Fuqarolik_sudlari_yurisdiksiya.xlsx yo'q"
    if not manzil:
        return None, "javobgar manzili topilmadi"
    # skaner matni: "К,ашкадарё", "К^арши", "шах;ри" - so'z ichidagi belgilarni olib tashlaymiz
    matn = re.sub(r"(?<=[^\W\d_])[,^;:|](?=[^\W\d_])", "", manzil)
    tok = _sozlar(re.sub(r"(?i)\b(р-н|r-n)\b", " район ", matn))
    sk = [_skelet(w) for w in tok]

    def turi(i, j):
        """Nom yonidagi so'z: 'Андижан г.', 'Нукус шахри', 'Шайхонтоҳур тумани', 'г. Ташкент'."""
        t = _soz_turi(tok[j + 1]) if j + 1 < len(tok) else None
        if not t and i > 0 and tok[i - 1] in ("g", "gorod", "rayon"):
            t = _soz_turi(tok[i - 1])
        return t

    # 1) viloyat (bo'lsa) - nom yonida "вилояти / область / Республикаси" yoki Toshkent uchun "шаҳри"
    hududlar = set()
    for h in {s["hudud"] for s in sl}:
        nsk, htur = _joy_nomi(h)
        for i, j in _oraliqlar(sk, nsk):
            t = turi(i, j)
            if t == "viloyat" and htur == "viloyat" or (h.startswith("Тошкент") and t == htur == "shahar"):
                hududlar.add(h)
    if len(hududlar) == 1 and _toshkent_shahri(next(iter(hududlar))) and toshkent_sudi():
        return toshkent_sudi(), "Toshkent shahri - oferta 8.6 (sotuvchi yuridik manzili)"
    # 2) tuman / shahar nomlari
    nomzodlar = []
    for s in sl:
        if hududlar and s["hudud"] not in hududlar:
            continue
        for joy in s["joylar"]:
            nsk, jtur = _joy_nomi(joy)
            for i, j in _oraliqlar(sk, nsk):
                t = turi(i, j)
                if t == "viloyat":
                    continue                     # bu viloyat nomi (Андижанская область), tuman/shahar emas
                if not t and ((j + 1 < len(tok) and tok[j + 1] in MAHALLA_SOZLARI)
                              or (i > 0 and tok[i - 1] in OLD_SOZLARI)):
                    continue                     # "ДЕХҚОНОБОД МСГ", "ул. Навбахор" - mahalla/ko'cha nomi
                if not t and not hududlar and sum(len(x) for x in nsk) <= 3:
                    continue                     # qisqa nom (Сўх, Поп) "tumani/шаҳар" so'zisiz - tasodifiy moslik
                ball = 1 + (2 if t and t == jtur else -1 if t else 0) + (3 if hududlar else 0)
                if ball > 0:
                    nomzodlar.append((ball, s, joy, _yaqinlik(joy, tok, i, j)))
    if not nomzodlar:
        return None, "manzildagi tuman/shahar sudlar jadvalida topilmadi"
    eng = max(n[0] for n in nomzodlar)
    eng_sudlar = {}
    for n in sorted((n for n in nomzodlar if n[0] == eng), key=lambda n: -n[3]):
        eng_sudlar.setdefault(n[1]["nomi"], n)
    if len(eng_sudlar) > 1:
        # skeleti bir xil nomlar (Яшнобод / Сайхунобод -> "shnbd"): unlilari bilan yaqinroq yozilgani
        a, b = sorted(eng_sudlar.values(), key=lambda n: -n[3])[:2]
        if a[3] - b[3] < 0.15:
            return None, "bir nechta sud mos keldi: " + ", ".join(eng_sudlar)
        eng_sudlar = {a[1]["nomi"]: a}
    _b, s, joy, _y = next(iter(eng_sudlar.values()))
    if _toshkent_shahri(s["hudud"]) and toshkent_sudi():     # masalan Yashnobod -> Mirobod emas, Shayxontohur
        return toshkent_sudi(), f"{s['hudud']} / {joy} - oferta 8.6 (sotuvchi yuridik manzili)"
    return s, f"{s['hudud']} / {joy}"


def _yaqinlik(joy, tok, i, j):
    """Joy nomi manzildagi so'zlarga unlilari bilan qanchalik o'xshaydi (0..1): 'Яшнободский' ~ 'Яшнобод'."""
    nom = " ".join(w for w in _sozlar(joy) if _soz_turi(w) is None)
    return difflib.SequenceMatcher(None, nom, " ".join(tok[i:j + 1])[:len(nom) + 2]).ratio()


def sud_nomidan(nom):
    """Qo'lda yozilgan sud nomi (kirill yoki lotin) -> jadvaldagi sud."""
    ballar = [(sud_variant_balli(nom, s), s) for s in sudlar()]
    ballar = [x for x in ballar if x[0] > 0]
    return max(ballar, key=lambda x: x[0])[1] if ballar else None


def sud_variant_balli(variant, sud):
    """Saytdagi 'Sud nomi' varianti ('1003 - Fuqarolik ishlari bo'yicha Shayxontohur tumanlararo sudi') shu sudmi."""
    nsk, tur = _joy_nomi(sud["nomi"])
    tok = _sozlar(variant)
    if not _oraliqlar([_skelet(w) for w in tok], nsk):
        return 0
    ball = 2
    if tur in tok:
        ball += 2
    elif tur == "tuman" and "tumanlararo" in tok or tur == "tumanlararo" and "tuman" in tok:
        ball -= 1
    if "fuqarolik" in tok:
        ball += 1
    return max(ball, 1)


def viloyat_variant_balli(variant, hudud):
    """Saytdagi 'Viloyat' varianti ('10 - Тошкент шаҳар') shu hududmi. Toshkent shahri / viloyati turi bilan ajratiladi."""
    nsk, tur = _joy_nomi(hudud)
    tok = _sozlar(variant)
    if not _oraliqlar([_skelet(w) for w in tok], nsk):
        return 0
    turlar = {_soz_turi(w) for w in tok}
    if hudud.startswith("Тошкент"):
        return 3 if tur in turlar else 0
    return 3 if "viloyat" in turlar else 2


def davo_ariza_malumoti(pdf, docx_yol=None, ocr=True):
    """Avval PDF dan, keyin DOCX dan (DOCX toza matn - ustun turadi).
    PDF skaner (matnsiz) va DOCX yo'q bo'lsa: javobgar ma'lumotlari - dalolatnoma/talabnoma/pochtadan,
    summalar - arizaning o'zidan Windows OCR bilan (ocr=False - faqat avval o'qilgani, keshdan)."""
    matn = pdf_matn(pdf) if pdf else ""
    r = matndan_malumot(matn)
    if docx_yol:
        r.update({k: v for k, v in matndan_malumot(docx_matn(docx_yol)).items() if v})
    elif pdf and matnsizmi(matn):
        r = {}                                      # buzuq matndan olingani ishlatilmaydi
        r.update(hujjatlardan_malumot(os.path.dirname(pdf)))
        for k, v in skan_malumoti(pdf, ocr=ocr).items():
            r.setdefault(k, v)
    if pdf and not r.get("manzil"):    # ariza boshqa shablonda - manzil boshqa hujjatlardan (pochta "Қаерга:")
        for k, v in hujjatlardan_malumot(os.path.dirname(pdf)).items():
            r.setdefault(k, v)
    return r


# ---------- skaner qilingan (matnsiz) da'vo ariza ----------
def hujjatlardan_malumot(papka):
    """Javobgar FIO, pasport, JShShIR, manzil - mijozning matnli hujjatlaridan
    (dalolatnoma "shartnoma.pdf", zaxira: talabnoma, pochta kvitansiyasi)."""
    matnlar = {}
    for f, yol in papka_fayllari(papka):
        k = hujjat_turi(f)
        if k in ("dalolatnoma", "talabnoma", "pochta") and f.lower().endswith(".pdf") and k not in matnlar:
            matnlar[k] = re.sub(r"\s+", " ", pdf_matn(yol))
    r = {}
    t = matnlar.get("dalolatnoma", "")
    # "...fuqarosi, FIO (AA1234567 01.01.2020 yilda berilgan) MANZIL da doimiy yashovchi" (yoki ruscha qismi)
    x = (re.search(r"fuqarosi,?\s*(.+?)\s*\(\s*([A-Z]{2})\s?(\d{7})[^)]*\)\s*(.+?)\s*da doimiy yashovchi", t) or
         re.search(r"Узбекистан\s+(.+?),?\s*\(\s*([A-Z]{2})\s?(\d{7})[^)]*\)\s*проживающий по адресу\s*(.+?),?\s*именуемый", t))
    manzillar = []            # hamma hujjatdagi manzil - biridan sud topilmasa keyingisi (ish_tayyorlash)
    if x:
        r["javobgar"], r["pasport"] = x.group(1).strip(" ,"), x.group(2) + x.group(3)
        manzillar.append(x.group(4).strip(" ,"))
    x = re.search(r"JSH\s*SH\s*IR\s*[:：]?\s*(\d{14})", t, re.I)
    if x:
        r["jshshir"] = x.group(1)
    # talabnoma: "...-сонли FIOга MANZIL Тел: ..."
    x = re.search(r"-сонли\s+(.+?)га\s+(.+?)\s+Тел", matnlar.get("talabnoma", ""))
    if x:
        r.setdefault("javobgar", x.group(1).strip())
        manzillar.append(x.group(2).strip(" ,"))
    # pochta kvitansiyasi: "Қаерга: MANZIL , Етказиб ..."
    x = re.search(r"Қаерга:\s*(.+?)\s*,?\s*Етказиб", matnlar.get("pochta", ""))
    if x:
        manzillar.append(x.group(1).strip(" ,"))
    if manzillar:
        r["manzil"], r["manzillar"] = manzillar[0], manzillar
    return r


# Windows'ning o'z OCR'i (Windows.Media.Ocr) - qo'shimcha dastur o'rnatish shart emas.
# Argumentlar: rasm fayllari; har rasm matni "=====SAHIFA=====" bilan ajratiladi.
WIN_OCR_PS = r"""
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Kut($op, [Type]$tur) {
    $t = $asTask.MakeGenericMethod($tur).Invoke($null, @($op))
    $t.Wait(-1) | Out-Null
    $t.Result
}
$en = [Windows.Globalization.Language]::new("en-US")
if ([Windows.Media.Ocr.OcrEngine]::IsLanguageSupported($en)) { $eng = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($en) }
else { $eng = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
if ($eng -eq $null) { throw "Windows OCR tili topilmadi" }
foreach ($rasm in $args) {
    $fayl = Kut ([Windows.Storage.StorageFile]::GetFileFromPathAsync($rasm)) ([Windows.Storage.StorageFile])
    $oqim = Kut ($fayl.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $dek = Kut ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($oqim)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bmp = Kut ($dek.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $res = Kut ($eng.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
    "=====SAHIFA====="
    foreach ($q in $res.Lines) { $q.Text }
    $oqim.Dispose()
}
"""
SKAN_KESH_FAYL = os.path.join(PAPKA, "ocr_skan_kesh.json")   # skaner arizalar matni (har safar qayta o'qilmasin)
_KESHLAR = {}
_KESH_QULF = threading.Lock()


def _kesh_kalit(pdf):
    st = os.stat(pdf)
    return f"{os.path.relpath(pdf, PAPKA)}|{st.st_size}|{int(st.st_mtime)}"


def _kesh(fayl):
    """OCR natijalari keshi (fayl -> lug'at). Fon oqimi va oyna birga ishlatadi."""
    with _KESH_QULF:
        if fayl not in _KESHLAR:
            try:
                with open(fayl, encoding="utf-8") as f:
                    _KESHLAR[fayl] = json.load(f)
            except (OSError, ValueError):
                _KESHLAR[fayl] = {}
        return _KESHLAR[fayl]


def _kesh_yoz(fayl, kalit, qiymat):
    kesh = _kesh(fayl)
    with _KESH_QULF:
        kesh[kalit] = qiymat
        try:
            with open(fayl + ".tmp", "w", encoding="utf-8") as f:
                json.dump(kesh, f, ensure_ascii=False, indent=0)
            os.replace(fayl + ".tmp", fayl)
        except OSError:
            pass


def _windows_ocr(rasmlar):
    import subprocess
    with tempfile.TemporaryDirectory() as tmp:
        ps1 = os.path.join(tmp, "ocr.ps1")
        with open(ps1, "w", encoding="ascii") as f:
            f.write(WIN_OCR_PS)
        p = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", ps1, *rasmlar],
                           capture_output=True, timeout=180, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if p.returncode != 0:
        xato = (p.stderr or b"").decode("utf-8", "replace").strip().splitlines()
        raise RuntimeError(xato[0] if xato else f"Windows OCR xatosi ({p.returncode})")
    return p.stdout.decode("utf-8", "replace")


def skan_matn(pdf, ocr=True):
    """Matnsiz (skaner) PDF matni - Windows OCR bilan. Natija ocr_skan_kesh.json ga saqlanadi.
    ocr=False - faqat keshdan (hali o'qilmagan bo'lsa None)."""
    try:
        kalit = _kesh_kalit(pdf)
    except OSError:
        return ""
    kesh = _kesh(SKAN_KESH_FAYL)
    if kalit in kesh:
        return kesh[kalit]
    if not ocr:
        return None
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    print(f"    OCR -> skaner ariza o'qilmoqda: {os.path.relpath(pdf, MIJOZLAR_PAPKA)}")
    try:
        with tempfile.TemporaryDirectory() as tmp:
            rasmlar = []
            with FITZ_QULF, fitz.open(pdf) as d:
                for i, pg in enumerate(d):
                    rasmlar.append(os.path.join(tmp, f"{i}.png"))
                    pg.get_pixmap(matrix=fitz.Matrix(3, 3), colorspace=fitz.csGRAY).save(rasmlar[-1])
            matn = _windows_ocr(rasmlar)
    except Exception as e:
        print(f"    !!! Skaner arizani o'qib bo'lmadi ({os.path.basename(pdf)}): {e}")
        return ""
    _kesh_yoz(SKAN_KESH_FAYL, kalit, matn)
    return matn


def skan_malumoti(pdf, ocr=True):
    """Skaner arizadan (OCR): asosiy qarz, pochta, JShShIR. FIO/manzil olinmaydi - OCR ustunlarni aralashtiradi."""
    t = skan_matn(pdf, ocr=ocr)
    if not t:
        return {}
    t = re.sub(r"\bs[o0c][^\s\d]{0,2}(?:m|rn)\b", "so'm", t)       # OCR: soirn / scjm / so'm -> so'm
    t = re.sub(r"(\d{3}) (\d{2})(?=\s*so'm)", r"\1,\2", t)          # OCR: "41 200 00 so'm" -> "41 200,00 so'm"
    m = matndan_malumot(t)
    r = {k: m[k] for k in ("jshshir", "pochta") if m.get(k)}
    # asosiy qarz arizada 2 marta yozilgan: ikkalasi mos kelsa (yoki bittasi o'qilsa) olinadi
    qiymatlar = {summa_tozalash(x.group(1)) for x in (re.search(SUMMA + r"\s*so'm\s+asosiy\s+qarz", t, re.I),
                                                     re.search(SUMMA + r"\s*so'm\s+miqdoridagi\s+qarzdorlik", t, re.I))
                 if x} - {""}
    if len(qiymatlar) == 1:
        r["asosiy_qarz"] = qiymatlar.pop()
    elif qiymatlar:
        print(f"    !!! {os.path.basename(pdf)}: OCR ikki xil summa o'qidi ({', '.join(sorted(qiymatlar))}) - "
              "asosiy qarzni qo'lda kiriting")
    return r


def ariza_bosh_rasmi(pdf):
    """Da'vo arizaning qo'lda yozilgan sana va raqam joyining rasmi (PNG baytlari) - oynada ko'rsatish uchun."""
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        with FITZ_QULF, fitz.open(pdf) as d:
            pg = d[0]
            r = pg.rect
            clip = fitz.Rect(0, r.height * 0.14, r.width * 0.55, r.height * 0.30)
            return pg.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), clip=clip).tobytes("png")
    except Exception as e:
        print(f"    (arizaning rasmi olinmadi: {e})")
        return None


# ============================================================
# QO'LDA YOZILGAN SANA VA RAQAM (ariza tepasida, ko'k ruchka)
# 1) O'qish (qolyozma_tahlil, keshlanadi): ko'k siyoh bo'laklari topiladi - yuqori qator sana (kun va oy - alohida
#    bo'laklar), pastki qator raqam. Har biri ASL rasmdan (oq-qora qilinmaydi - model shunda yaxshi o'qiydi) 2 xil
#    ishlov x 2 xil chegara bilan kesiladi va PaddleOCR (rapidocr-onnxruntime) faqat raqamlar bo'yicha CTC bilan
#    baholanadi. Siyoh topilmasa (qora ruchka, oq-qora skaner) - RapidOCR qutilari ("2026yil", "/09-02" atrofi).
#    Keshda: nomzodlar log-ehtimoli + raqam rasmlarining CTC matritsalari (keyin yangi nomzodni ham baholash uchun).
# 2) Yechim (qolyozma_yechim): OCR x oldingi bilim -
#    - paket qoidasi: ariza raqami = papka raqami + doimiy farq (044 -> 100044, farq 100000). Farq SHU paketdagi
#      arizalarning OCR o'qishlaridan topiladi, raqam_tarixi.json dagi oldingi paketlar - qo'shimcha ovoz;
#    - sana: raqamlar vaqt bo'yicha o'sadi - raqami kattaroq arizaning sanasi oldinroq bo'lmaydi;
#    - bir raqam ikki arizada bo'lmaydi; avval yuborilgan arizaning raqami/sanasi - o'sha.
#    Avtomatik qabul: ishonch >= ocr_ishonch VA OCR o'zi ham aynan shu qiymatni o'qigan (ikki mustaqil dalil).
#    Odam tasdiqlagan qiymat dalillarga qattiq zid bo'lsa - qayta TEKSHIRING (shoshib Enter bosilgan bo'lishi mumkin).
# Yuborilgan/saqlangan har ariza raqam_tarixi.json ga yoziladi - keyingi yuklashda farq va sanalar undan ham olinadi.
# ============================================================
QOLYOZMA_KESH_FAYL = os.path.join(PAPKA, "ocr_qolyozma_kesh.json")
QOLYOZMA_VERSIYA = 4
RAQAM_TARIXI_FAYL = os.path.join(PAPKA, "raqam_tarixi.json")
SANA_OYNA = 60          # qo'lda yozilgan sana "... yil holatiga" sanasidan keyingi shuncha kun ichida
SANA_BUZISH = 1.5       # raqam/sana tartibiga zid har langar uchun jarima (log): e^-1.5 = 0.22 marta
_RAPID = None


def _rapid():
    """RapidOCR (bir marta yuklanadi): (dvigatel, tanib_oluvchi, '0'..'9' ustunlari). O'rnatilmagan bo'lsa None."""
    global _RAPID
    if _RAPID is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
            eng = RapidOCR()
            rec = eng.text_recognizer
            belgilar = rec.postprocess_op.character
            _RAPID = (eng, rec, [belgilar.index(d) for d in "0123456789"])
        except Exception as e:
            print(f"    !!! Qo'lyozma OCR ishlamaydi (pip install rapidocr-onnxruntime): {type(e).__name__}: {e}")
            _RAPID = False
    return _RAPID or None


def _bosh_bgr(pdf, zoom=3.0):
    """Ariza tepasi (sana va raqam joyi) - BGR rasm."""
    import cv2
    import numpy as np
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    with FITZ_QULF, fitz.open(pdf) as d:
        pg = d[0]
        r = pg.rect
        clip = fitz.Rect(0, r.height * 0.14, r.width * 0.55, r.height * 0.30)
        pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
        return cv2.cvtColor(np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3), cv2.COLOR_RGB2BGR)


def _siyoh_maska(img):
    """Ko'k/binafsha siyoh piksellari (1) - bosma matn, chiziqlar, kulrang fon kirmaydi."""
    import cv2
    import numpy as np
    h, s, v = cv2.split(cv2.cvtColor(img, cv2.COLOR_BGR2HSV))
    return ((h > 95) & (h < 150) & (s > 45) & (v > 40)).astype(np.uint8)


def _faqat_siyoh(img):
    """Ko'k/binafsha siyoh -> oq fonda qora yozuv (qutilar usulida RapidOCR uchun)."""
    import cv2
    import numpy as np
    m = cv2.dilate(_siyoh_maska(img), np.ones((2, 2), np.uint8))
    return np.where(m[..., None] > 0, 0, 255).astype(np.uint8).repeat(3, 2)


def _guruhla(qutilar, ox, oraliq):
    """Qutilar (x1, y1, x2, y2, piksel) ni o'q bo'yicha (0 - x, 1 - y) guruhlaydi: oldingisidan oraliq'dan ko'p
    uzoq bo'lsa - yangi guruh. Har guruh - umumiy quti."""
    gr = []
    for b in sorted(qutilar, key=lambda b: b[ox]):
        if gr and b[ox] - max(c[ox + 2] for c in gr[-1]) <= oraliq:
            gr[-1].append(b)
        else:
            gr.append([b])
    return [(min(b[0] for b in g), min(b[1] for b in g), max(b[2] for b in g), max(b[3] for b in g),
             sum(b[4] for b in g)) for g in gr]


def _siyoh_joylari(img):
    """Qo'lda yozilgan joylar: (sana qatori, (kun, oy) yoki None, raqam qatori) - har biri (x1, y1, x2, y2).
    Siyoh topilmasa yoki ikki qator bo'lmasa - None."""
    import cv2
    import numpy as np
    m = _siyoh_maska(img)
    n, lab, st, _c = cv2.connectedComponentsWithStats(cv2.dilate(m, np.ones((5, 5), np.uint8)), 8)
    px = np.bincount(lab[m > 0], minlength=n)
    bolak = [(st[i][0], st[i][1], st[i][0] + st[i][2], st[i][1] + st[i][3], int(px[i]))
             for i in range(1, n) if px[i] >= 40 and st[i][3] >= 12]
    qatorlar = [q for q in _guruhla(bolak, 1, -8) if q[4] >= 300 and q[3] - q[1] >= 25]   # ustma-ust tushmasa - boshqa qator
    if len(qatorlar) < 2:
        return None
    sana_q, raqam_q = qatorlar[0], qatorlar[1]
    if raqam_q[2] - raqam_q[0] < 1.5 * (raqam_q[3] - raqam_q[1]):
        return None
    qism = _guruhla([b for b in bolak if b[1] >= sana_q[1] - 1 and b[3] <= sana_q[3] + 1], 0, 18)
    qism = sorted(sorted(qism, key=lambda q: -q[4])[:2], key=lambda q: q[0])
    return sana_q[:4], (tuple(q[:4] for q in qism) if len(qism) == 2 else None), raqam_q[:4]


def _kes(src, q, pad, padx):
    """Quti atrofidan (balandlikka nisbatan pad/padx chegara bilan) kesim."""
    x1, y1, x2, y2 = q[:4]
    h = y2 - y1
    H, W = src.shape[:2]
    return src[max(0, int(y1 - h * pad)):min(H, int(y2 + h * pad)), max(0, int(x1 - h * padx)):min(W, int(x2 + h * padx))]


def _korinishlar(img):
    """OCR uchun 2 xil rasm: asl va kontrasti oshirilgani (CLAHE)."""
    import cv2
    g = cv2.createCLAHE(2.0, (8, 8)).apply(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
    return [img, cv2.cvtColor(g, cv2.COLOR_GRAY2BGR)]


def _raqam_logp(rec, ustun, kesim):
    """Kesim -> (T, 11) log-ehtimollar: [blank + raqam bo'lmagan hamma belgi, '0'..'9']."""
    import numpy as np
    h, w = kesim.shape[:2]
    norm = rec.resize_norm_img(kesim, max(1.0, w / float(h)))[None].astype(np.float32)
    d = rec.session(norm)[0][0][:, ustun]
    blank = np.clip(1.0 - d.sum(1, keepdims=True), 1e-6, 1.0)
    return np.log(np.clip(np.concatenate([blank, d], 1), 1e-12, 1.0))


def _ctc_baho(lp, yorliqlar):
    """CTC: bir xil uzunlikdagi raqam satrlarining har biri uchun log P(satr | rasm)."""
    import numpy as np
    lab = np.zeros((len(yorliqlar), 2 * len(yorliqlar[0]) + 1), dtype=int)
    for i, y in enumerate(yorliqlar):
        lab[i, 1::2] = [int(c) + 1 for c in y]
    g = lp[:, lab]                                           # (T, N, L)
    sakrash = np.zeros(lab.shape, dtype=bool)
    sakrash[:, 2:] = (lab[:, 2:] != 0) & (lab[:, 2:] != lab[:, :-2])
    a = np.full(lab.shape, -np.inf)
    a[:, 0], a[:, 1] = g[0, :, 0], g[0, :, 1]
    bosh = np.full((lab.shape[0], 1), -np.inf)
    for t in range(1, lp.shape[0]):
        a2 = np.concatenate([bosh, bosh, a[:, :-2]], 1)
        a2[~sakrash] = -np.inf
        a = np.logaddexp(np.logaddexp(a, np.concatenate([bosh, a[:, :-1]], 1)), a2) + g[t]
    return np.logaddexp(a[:, -1], a[:, -2])


def _beam(lp, kenglik=10):
    """CTC prefix beam search (faqat raqamlar): eng ehtimolli satrlar [(satr, log p)]."""
    import numpy as np
    nurlar = {"": (0.0, -np.inf)}                             # prefiks -> (log p, blank bilan; log p, raqam bilan)
    for t in range(lp.shape[0]):
        yangi = {}

        def qosh(k, pb, pn):
            ob, on = yangi.get(k, (-np.inf, -np.inf))
            yangi[k] = (np.logaddexp(ob, pb), np.logaddexp(on, pn))
        for pre, (pb, pn) in nurlar.items():
            qosh(pre, np.logaddexp(pb, pn) + lp[t, 0], -np.inf)
            for c in range(1, 11):
                d = str(c - 1)
                if pre and pre[-1] == d:
                    qosh(pre, -np.inf, pn + lp[t, c])               # takror - o'sha raqam davom etadi
                    qosh(pre + d, -np.inf, pb + lp[t, c])           # blank orqali - yangi raqam
                else:
                    qosh(pre + d, -np.inf, np.logaddexp(pb, pn) + lp[t, c])
        nurlar = dict(sorted(yangi.items(), key=lambda kv: -np.logaddexp(*kv[1]))[:kenglik])
    return [(k, float(np.logaddexp(*v))) for k, v in nurlar.items()]


def _ballar(korinish, satrlar):
    """Har satr uchun rasmlar (ko'rinishlar) bo'yicha o'rtacha log P(satr + qo'shimcha | rasm).
    korinish - [(lp, qo'shimcha)]: qo'shimcha - raqamga qo'shilib o'qilgan bosma "0902" ("/09-02")."""
    import numpy as np
    satrlar = [s for s in dict.fromkeys(satrlar) if s and s.isdigit()]
    if not satrlar or not korinish:
        return {}
    uz = {}
    for i, s in enumerate(satrlar):
        uz.setdefault(len(s), []).append(i)
    jami = np.zeros(len(satrlar))
    for lp, suf in korinish:
        for idx in uz.values():
            jami[idx] += _ctc_baho(lp, [satrlar[i] + suf for i in idx])
    return dict(zip(satrlar, np.maximum(jami / len(korinish), -1e4).round(3).tolist()))


def _lp_kodla(lp):
    import base64
    import numpy as np
    return base64.b64encode(np.maximum(lp, -60).astype(np.float16).tobytes()).decode("ascii")


def _lp_och(s):
    import base64
    import numpy as np
    return np.frombuffer(base64.b64decode(s), np.float16).astype(np.float64).reshape(-1, 11)


KUNLAR = [str(i) for i in range(1, 32)] + [f"{i:02d}" for i in range(1, 10)]
OYLAR = [str(i) for i in range(1, 13)] + [f"{i:02d}" for i in range(1, 10)]


def holatiga_sanasi(pdf, ocr=True):
    """Qo'lda yozilgan sana shundan keyin bo'ladi: arizadagi "27.07.2026 yil holatiga ... qarzdorlik" sanasi,
    topilmasa - talabnoma sanasi (talabnoma.pdf: "22.07.2026 йил")."""
    matn = pdf_matn(pdf)
    if matnsizmi(matn):
        matn = skan_matn(pdf, ocr=ocr) or ""
    # OCR xatolariga chidamli: "2707.2026 Yil hofatiga"
    x = re.search(r"(\d{1,2})[.,]?\s?(\d{2})[.,]\s?(20\d\d)\s*y[i1l]l\S*\s*ho\S{0,2}atiga", matn, re.I)
    if not x:
        for f, yol in papka_fayllari(os.path.dirname(pdf)):
            if hujjat_turi(f) == "talabnoma" and f.lower().endswith(".pdf"):
                x = re.search(r"(\d{2})\.(\d{2})\.(20\d\d)\s*йил", pdf_matn(yol))
                if x:
                    break
    try:
        return date(int(x.group(3)), int(x.group(2)), int(x.group(1))) if x else None
    except ValueError:
        return None


def _sana_nomzodlari(holatiga):
    """Qo'lda yozilgan sana bo'lishi mumkin bo'lgan kunlar."""
    from datetime import timedelta
    boshi = holatiga or (date.today() - timedelta(days=200))
    return [s for s in (boshi + timedelta(days=i) for i in range(SANA_OYNA if holatiga else 200)) if s <= date.today()]


def _qutilar_usuli(img, eng, rec, ustun):
    """Siyoh topilmaganda: RapidOCR topgan matn qutilari bo'yicha ("2026yil" qatori - sana, pastki - raqam).
    (raqam_rasmlari [(lp, qo'shimcha)], sana_rasmlari [lp])."""
    toza = _faqat_siyoh(img)
    qutilar = []
    for src in (img, toza):
        res, _ = eng(src)
        for b, t, _s in res or []:
            xs, ys = [p[0] for p in b], [p[1] for p in b]
            qutilar.append(((int(min(xs)), int(min(ys)), int(max(xs)), int(max(ys))), str(t)))
    yil = next((q for q, t in qutilar if "yil" in t.lower() or re.search(r"20\d\dy", t.replace(" ", "").lower())), None)
    # bosma "/09-02" (raqamdan o'ngda) - raqam qutisi topilmasa shu qatordan kesiladi
    kod_q = next((q for q, t in qutilar if (lambda r: 3 <= len(r) <= 5 and r.endswith("902"))(re.sub(r"\D", "", t))), None)
    sana_lp, raqam_lp = [], []
    for q, t in qutilar:
        rq = re.sub(r"\D", "", t)
        kod = len(rq) <= 5 and rq.endswith("902")             # bosma "/09-02"
        tur = None
        if yil:
            if abs((q[1] + q[3]) / 2 - (yil[1] + yil[3]) / 2) < (yil[3] - yil[1]) * 0.8:
                tur = "sana" if q[2] <= yil[0] + 5 else None     # sana qatori, "2026yil" dan chapda
            elif q[1] > yil[3] - 5 and len(rq) >= 5 and not kod:
                tur = "raqam"
        elif len(rq) >= 5 and not kod:
            tur = "raqam"
        for src in (img, toza) if tur else ():
            for pad in (0.05, 0.25):
                c = _kes(src, q, pad, pad)
                if c.size:
                    lp = _raqam_logp(rec, ustun, c)
                    if tur == "sana":
                        sana_lp.append(lp)
                    else:
                        g = "".join(s for s, _p in _beam(lp, 1))      # raqam "/09-02" bilan qo'shilib ketgan bo'lishi mumkin
                        raqam_lp.append((lp, "0902" if len(g) > 7 and g.endswith("902") else ""))
    if kod_q and not raqam_lp:                                    # raqam qutisi topilmadi - "/09-02" chapidagi butun qator
        x1, y1, x2, y2 = kod_q
        h = y2 - y1
        for src in (img, toza):
            for py in (0.6, 0.9):
                c = src[max(0, int(y1 - py * h)):int(y2 + 0.3 * h), int(x1 * 0.2):x1]
                if c.size:
                    raqam_lp.append((_raqam_logp(rec, ustun, c), ""))
    if yil and not sana_lp:                                       # sana qutisi topilmadi - "2026yil" chapidagi butun qator
        x1, y1, x2, y2 = yil
        h = y2 - y1
        for src in (img, toza):
            for py in (0.55, 0.75):
                c = src[max(0, int(y1 - py * h)):int(y2 + 0.3 * h), int(x1 * 0.25):x1]
                if c.size:
                    sana_lp.append(_raqam_logp(rec, ustun, c))
    return raqam_lp, sana_lp


def _qolyozma_oqish(pdf, eng, rec, ustun):
    holatiga = holatiga_sanasi(pdf)
    img = _bosh_bgr(pdf)
    joy = _siyoh_joylari(img)
    raqam_lp, kun_lp, oy_lp, sana_lp = [], [], [], []
    if joy:
        sana_q, kun_oy, raqam_q = joy
        for src in _korinishlar(img):
            for pad, padx in ((0.08, 0.15), (0.2, 0.3)):
                raqam_lp.append((_raqam_logp(rec, ustun, _kes(src, raqam_q, pad, padx)), ""))
            for pad, padx in ((0.2, 0.3), (0.35, 0.5)):
                if kun_oy:
                    kun_lp.append((_raqam_logp(rec, ustun, _kes(src, kun_oy[0], pad, padx + 0.15)), ""))
                    oy_lp.append((_raqam_logp(rec, ustun, _kes(src, kun_oy[1], pad, padx + 0.15)), ""))
                else:
                    sana_lp.append(_raqam_logp(rec, ustun, _kes(src, sana_q, pad, padx)))
    else:
        raqam_lp, sana_lp = _qutilar_usuli(img, eng, rec, ustun)
    natija = {"usul": "siyoh" if joy else "qutilar", "holatiga": holatiga.isoformat() if holatiga else "",
              "raqam": {}, "lp": [], "kun": {}, "oy": {}, "sana": {}}
    if raqam_lp:
        nomzod = set()
        for lp, suf in raqam_lp:
            for x, _p in _beam(lp):
                x = x[:-4] if suf and x.endswith("0902") else x
                nomzod.update(y for y in (x, x[:6], x[-6:]) if 5 <= len(y) <= 7)
        for x in list(nomzod):                                    # bitta raqami boshqacha variantlar ham
            for i in range(len(x)):
                nomzod.update(x[:i] + d + x[i + 1:] for d in "0123456789")
        ball = _ballar(raqam_lp, nomzod)
        natija["raqam"] = {k: ball[k] for k in sorted(ball, key=ball.get, reverse=True)[:80]}
        natija["lp"] = [[_lp_kodla(lp), suf] for lp, suf in raqam_lp]
    if kun_lp:
        natija["kun"] = _ballar(kun_lp, KUNLAR)
        natija["oy"] = _ballar(oy_lp, OYLAR)
    elif sana_lp:
        sanalar = _sana_nomzodlari(holatiga)
        yozuv = {}                                                # "0508", "508", "058", "58" -> sanalar
        for s in sanalar:
            for y in {f"{s.day:02d}{s.month:02d}", f"{s.day}{s.month:02d}", f"{s.day:02d}{s.month}", f"{s.day}{s.month}"}:
                yozuv.setdefault(y, []).append(s)
        ball = _ballar([(lp, "") for lp in sana_lp], yozuv)
        sb = {}
        for y, v in ball.items():
            for s in yozuv[y]:
                sb[s.isoformat()] = max(sb.get(s.isoformat(), -1e4), v)
        natija["sana"] = sb
    return natija


def qolyozma_tahlil(pdf, ocr=True):
    """Arizadagi qo'lda yozilgan raqam va sana bo'yicha OCR dalili (nomzodlar log-ehtimoli). Keshlanadi;
    ocr=False - faqat keshdan (hali o'qilmagan bo'lsa None)."""
    kalit = _kesh_kalit(pdf) + f"|v{QOLYOZMA_VERSIYA}"
    kesh = _kesh(QOLYOZMA_KESH_FAYL)
    if kalit in kesh:
        return kesh[kalit]
    if not ocr:
        return None
    r = _rapid()
    if not r:
        return None
    try:
        natija = _qolyozma_oqish(pdf, *r)
    except Exception as e:
        print(f"    !!! Qo'lda yozilgan sana/raqam o'qilmadi ({os.path.basename(pdf)}): {type(e).__name__}: {e}")
        return None
    _kesh_yoz(QOLYOZMA_KESH_FAYL, kalit, natija)
    return natija


class _Dalil:
    """Bitta arizaning OCR dalili (qolyozma_tahlil natijasi): istalgan raqam/sana uchun log-ehtimol."""

    def __init__(self, q):
        self.q = q
        self.raqam_ball = dict(q.get("raqam") or {})
        self._lp = None
        self.holatiga = date.fromisoformat(q["holatiga"]) if q.get("holatiga") else None

    def raqam(self, satrlar):
        """{satr: log P}; keshda yo'q nomzod - saqlangan CTC matritsalari bilan shu yerda baholanadi."""
        yangi = [s for s in satrlar if s not in self.raqam_ball]
        if yangi and self.q.get("lp"):
            if self._lp is None:
                self._lp = [(_lp_och(a), suf) for a, suf in self.q["lp"]]
            self.raqam_ball.update(_ballar(self._lp, yangi))
        return {s: self.raqam_ball.get(s, -1e4) for s in satrlar}

    def sana(self, sanalar):
        k, o, j = self.q.get("kun"), self.q.get("oy"), self.q.get("sana")
        if k and o:
            return {s: max(k.get(str(s.day), -1e4), k.get(f"{s.day:02d}", -1e4))
                    + max(o.get(str(s.month), -1e4), o.get(f"{s.month:02d}", -1e4)) for s in sanalar}
        return {s: j.get(s.isoformat(), -1e4) for s in sanalar} if j else {}


_DALILLAR = {}


def _dalil(q):
    e = _DALILLAR.get(id(q))
    if e is None or e[0] is not q:
        e = _DALILLAR[id(q)] = (q, _Dalil(q))
    return e[1]


def ocr_chegara():
    """Avtomatik qabul qilish uchun eng kam ishonch (sozlamalar: ocr_ishonch)."""
    return min(0.999, max(0.5, _son("ocr_ishonch", 0.90)))


def papka_raqami(papka):
    """Papka nomi boshidagi tartib raqami: "044 - RN-25-00000 - ..." -> 44 (yo'q bo'lsa None)."""
    m = re.match(r"\s*(\d{1,4})(?!\d)", papka or "")
    return int(m.group(1)) if m else None


# ---------- raqam tarixi: yuborilgan arizalar (keyingi yuklashlar uchun dalil) ----------
_RAQAM_TARIXI = None
_TARIX_QULF = threading.RLock()


def raqam_tarixi():
    """Sudga yuborilgan/saqlangan arizalar raqami va sanasi {papka: [raqam, 'dd.mm.yyyy']}.
    Fayl yo'q bo'lsa - loglar\\ dagi YUBORILDI/SAQLANDI bo'lganlardan bir marta yig'iladi."""
    global _RAQAM_TARIXI
    with _TARIX_QULF:
        if _RAQAM_TARIXI is None:
            try:
                with open(RAQAM_TARIXI_FAYL, encoding="utf-8") as f:
                    _RAQAM_TARIXI = json.load(f)
            except (OSError, ValueError):
                _RAQAM_TARIXI = loglardan_tarix()
                if _RAQAM_TARIXI:
                    _tarix_saqla()
        return _RAQAM_TARIXI


def loglardan_tarix():
    """Har log: ro'yxatdagi "Ariza: №... sana: ..." + shu logda YUBORILDI/SAQLANDI bo'lgan papkalar."""
    tarix = {}
    try:
        loglar = sorted(f for f in os.listdir(LOG_PAPKA) if f.endswith(".log"))
    except OSError:
        return tarix
    for f in loglar:
        try:
            with open(os.path.join(LOG_PAPKA, f), encoding="utf-8", errors="ignore") as fh:
                t = fh.read()
        except OSError:
            continue
        yuborilgan = set(re.findall(r">>> NATIJA: (.+?) = (?:YUBORILDI|SAQLANDI)\b", t))
        if not yuborilgan:
            continue
        for x in re.finditer(r"^\[\d+\] (.+?)(?:\s+\(oxirgi natija[^)]*\))?[ \t]*\n\s+Ariza: №(\d+)\s+sana: (\d\d\.\d\d\.\d{4})",
                             t, re.M):
            if x.group(1).strip() in yuborilgan:
                tarix[x.group(1).strip()] = [x.group(2), x.group(3)]
    return tarix


def _tarix_saqla():
    try:
        with open(RAQAM_TARIXI_FAYL + ".tmp", "w", encoding="utf-8") as f:
            json.dump(_RAQAM_TARIXI, f, ensure_ascii=False, indent=0)
        os.replace(RAQAM_TARIXI_FAYL + ".tmp", RAQAM_TARIXI_FAYL)
    except OSError as e:
        print(f"    (raqam_tarixi.json yozilmadi: {e})")


def raqam_tarixiga(ish):
    """Sudga yuborilgan/saqlangan ariza raqami va sanasi - keyingi paketlar uchun dalil."""
    if ish.raqam and ish.sana:
        raqam_tarixi()
        with _TARIX_QULF:
            _RAQAM_TARIXI[ish.papka] = [str(ish.raqam), ish.sana.strftime("%d.%m.%Y")]
            _tarix_saqla()


def _tarix_langarlari():
    """raqam_tarixi -> {papka: (raqam, sana)} (buzuq yozuvlar tashlanadi)."""
    with _TARIX_QULF:
        xom = dict(raqam_tarixi() or {})
    tarix = {}
    for p, v in xom.items():
        try:
            r, s = str(v[0]).strip(), datetime.strptime(str(v[1]).strip(), "%d.%m.%Y").date()
        except (ValueError, TypeError, IndexError):
            continue
        if r.isdigit():
            tarix[p] = (r, s)
    return tarix


def _paket_farqi(ishlar, ocr, aniq, tarix):
    """Paket qoidasi: ariza raqami = papka raqami + farq. Farq nomzodlari - arizalar OCR'ining eng yaxshi 5 nomzodidan
    va raqami aniqlaridan; har nomzod shu paketdagi arizalarning shu raqamga OCR ehtimollari yig'indisi (E) bilan
    baholanadi (oldingi paketlar - qo'shimcha yarim ovoz). -> (farq, qoida ishonchi, E) yoki (None, 0, E)."""
    import math
    from collections import Counter
    joriy = {i.papka for i in ishlar}
    hisob = Counter()
    royxat = []                                   # (papka raqami, aniq raqam yoki None, ocr yozuvi yoki None)
    for i in ishlar:
        n = papka_raqami(i.papka)
        if n is None or not (i.papka in aniq or id(i) in ocr):
            continue
        royxat.append((n, aniq.get(i.papka), ocr.get(id(i))))
        if i.papka in aniq:
            hisob[aniq[i.papka] - n] += 2
        else:
            Lr = ocr[id(i)][1]
            for c in sorted(Lr, key=Lr.get, reverse=True)[:5]:
                hisob[int(c) - n] += 1
    tarix_f = Counter(int(r) - papka_raqami(p) for p, (r, _s) in tarix.items()
                      if p not in joriy and papka_raqami(p) is not None)
    for f, k in tarix_f.items():
        hisob[f] += min(k, 2)
    nomzod = [f for f, k in hisob.items() if k >= 2 and f > 0]
    if not nomzod or not royxat:
        return None, 0.0, 0.0
    E = dict.fromkeys(nomzod, 0.0)
    for n, a, o in royxat:
        if a is not None:
            for f in nomzod:
                E[f] += 1.0 if a - n == f else 0.0
        else:
            dl, _Lr, z = o
            b = dl.raqam([str(n + f) for f in nomzod])
            for f in nomzod:
                E[f] += math.exp(min(0.0, b[str(n + f)] - z))
    def ovoz(f):
        return E[f] + 0.5 * min(tarix_f.get(f, 0), 10)
    # shu paketning kamida chorak qismi qoidaga mos o'qilgan bo'lishi shart (oldingi paket farqi o'zi yetmaydi)
    yaroqli = [f for f in nomzod if ovoz(f) >= 2.0 and E[f] >= 0.25 * len(royxat)]
    if not yaroqli:
        return None, 0.0, max(E.values())
    farq = max(yaroqli, key=ovoz)
    t = 0.5 * min(tarix_f.get(farq, 0), 10)
    return farq, min(0.95, 0.5 + ovoz(farq) / (len(royxat) + t)), E[farq]


def qolyozma_yechim(ishlar):
    """Qo'lda yozilgan raqam/sana: OCR dalili x oldingi bilim (paket qoidasi, raqam/sana tartibi, raqam_tarixi).
    Tasdiqlanmaganlarga qiymat, variantlar va ishonch beradi (ishonchlisini avtomatik qabul qiladi); odam
    tasdiqlagani ("ha", lekin "ha!" emas) dalillarga qattiq zid bo'lsa - qayta tekshirishga qaytaradi."""
    import numpy as np
    T = ocr_chegara()
    avto_ruxsat = S.get("ocr_avto", "Ha").strip().lower() in ("ha", "yes", "1", "true")
    tarix = _tarix_langarlari()
    tasdiq = [(i.papka, int(i.raqam), i.sana) for i in ishlar if i.tekshirildi and i.raqam.isdigit() and i.sana]
    tasdiq_p = {x[0] for x in tasdiq}
    tarix_l = [(p, int(r), s) for p, (r, s) in tarix.items() if p not in tasdiq_p]
    aniq = {p: r for p, r, _s in tasdiq}
    aniq.update({p: int(r) for p, (r, _s) in tarix.items() if p not in aniq and any(i.papka == p for i in ishlar)})
    baho = [i for i in ishlar if i.qolyozma and not i.majburiy]
    ocr = {}
    for i in baho:
        dl = _dalil(i.qolyozma)
        Lr = dict(dl.raqam_ball)
        if Lr:
            ocr[id(i)] = (dl, Lr, float(np.logaddexp.reduce(list(Lr.values()))))
    farq, qoida_w, _E = _paket_farqi(ishlar, ocr, aniq, tarix)
    natija = {}
    ishonchli = []                         # shu paketda OCR va qoida birga tasdiqlaganlar - keyingi aylanishda langar
    for _aylanish in range(2):
        langar = sorted(tasdiq + tarix_l + ishonchli, key=lambda x: x[1])
        L_r = np.array([x[1] for x in langar], dtype=np.int64)
        L_s = np.array([x[2].toordinal() for x in langar], dtype=np.int64)
        L_p = np.array([x[0] for x in langar], dtype=object)
        egasi = {}                                 # raqam -> shu raqamli papkalar
        for p, r, _s in langar:
            egasi.setdefault(r, set()).add(p)
        yangi = []
        for i in baho:
            dl = _dalil(i.qolyozma)
            band = {r for r, ps in egasi.items() if ps - {i.papka}}
            n = papka_raqami(i.papka)
            qoida_r = str(n + farq) if (farq is not None and n is not None) else None
            oz = tarix.get(i.papka)
            Lr0 = ocr[id(i)][1] if id(i) in ocr else {}
            Lr = dl.raqam(sorted(set(Lr0) | {x for x in (qoida_r, oz and oz[0]) if x}))

            def oldin_r(c, qoida_r=qoida_r, oz=oz, band=band):
                lp = 0.0
                if oz:                                   # shu papka avval yuborilgan
                    lp += np.log(0.95 if c == oz[0] else 0.05 / 300)
                elif qoida_r:
                    lp += np.log(qoida_w if c == qoida_r else (1 - qoida_w) / 300)
                if int(c) in band and not (oz and c == oz[0]):
                    lp += np.log(0.05)                   # bu raqam boshqa arizada bor
                return lp
            sanalar = _sana_nomzodlari(dl.holatiga)
            Ls = dl.sana(sanalar)
            top_r = sorted(Lr, key=lambda c: -(Lr[c] + oldin_r(c)))[:25] or [""]
            top_s = sorted(Ls, key=lambda s: -Ls[s])[:12]
            top_s += [s for s in {x[2] for x in langar if x[0] != i.papka} | ({oz[1]} if oz else set())
                      if s in Ls and s not in top_s]
            if not top_s:
                top_s = [None]
            # birgalikdagi ehtimol: raqam x sana; langarlar bilan tartibi buzilgani - jarima
            mr = np.array([Lr[c] + oldin_r(c) if c else 0.0 for c in top_r])
            ms = np.array([Ls[s] + (np.log(0.95 if s == oz[1] else 0.05 / SANA_OYNA) if oz else 0.0) if s else 0.0
                           for s in top_s])
            M = mr[:, None] + ms[None, :]
            if len(langar) and top_r[0] and top_s[0]:
                # har nomzod raqam uchun - raqami eng yaqin 10 tadan langar (pastda va yuqorida)
                br = np.array([int(c) for c in top_r])
                idx = np.searchsorted(L_r, br)[:, None] + np.arange(-10, 10)[None, :]          # (R, 20)
                bor = (idx >= 0) & (idx < len(L_r))
                idx = np.clip(idx, 0, len(L_r) - 1)
                bor &= L_p[idx] != i.papka
                lr, ls = L_r[idx][:, None, :], L_s[idx][:, None, :]
                bs = np.array([s.toordinal() for s in top_s])[None, :, None]
                buz = (((lr < br[:, None, None]) & (ls > bs)) | ((lr > br[:, None, None]) & (ls < bs))) & bor[:, None, :]
                M = M - SANA_BUZISH * buz.sum(2)
            M = np.exp(M - M.max())
            M /= M.sum()
            pr = dict(zip(top_r, M.sum(1).tolist()))
            ps = dict(zip(top_s, M.sum(0).tolist()))
            rb, sb = max(pr, key=pr.get), max(ps, key=ps.get)
            ocr_r = max(Lr0, key=Lr0.get) if Lr0 else ""
            ocr_s = max(Ls, key=Ls.get) if Ls else None
            kuchli = bool(rb and sb and pr[rb] >= T and ps[sb] >= T)
            avto = kuchli and rb == ocr_r and sb == ocr_s
            natija[id(i)] = (rb, pr[rb] if rb else 0.0, sb, ps[sb] if sb else 0.0, ocr_r, ocr_s, kuchli, avto, qoida_r, n,
                             [(k, p) for k, p in sorted(pr.items(), key=lambda kv: -kv[1])[:3] if k],
                             [(k, p) for k, p in sorted(ps.items(), key=lambda kv: -kv[1])[:3] if k])
            if avto and not i.tekshirildi:
                yangi.append((i.papka, int(rb), sb))
        ishonchli = yangi

    for i in baho:
        rb, rp, sb, sp, ocr_r, ocr_s, kuchli, avto, qoida_r, n, rv, sv = natija[id(i)]
        izoh = []
        if i.tekshirildi:                        # odam tasdiqlagan - dalillar qattiq zid bo'lsagina qaytariladi
            zid = ([f"raqam №{i.raqam}, dalillar - №{rb}"] if rb and rb != i.raqam and rp >= T else []) + \
                  ([f"sana {i.sana:%d.%m.%Y}, dalillar - {sb:%d.%m.%Y}"] if sb and sb != i.sana and sp >= T else [])
            if not zid:
                continue
            i.jadval_shubha = (i.raqam, i.sana)
            i.tekshirildi = False
            izoh.append("jadvalda tasdiqlangan qiymat shubhali (" + "; ".join(zid) + ") - rasm bilan solishtiring")
        r_zid = bool(ocr_r and rb and ocr_r != rb) or bool(i.jadval_shubha and rb != i.jadval_shubha[0])
        s_zid = bool(ocr_s and sb and ocr_s != sb) or bool(i.jadval_shubha and sb != i.jadval_shubha[1])
        i.ocr_variantlar = {"raqam": rv, "sana": sv}
        i.ocr_ishonch = (rp, sp)
        i.ocr_tavsiya = (rb, sb, kuchli, r_zid, s_zid)
        if rb:
            i.raqam = rb
        if sb:
            i.sana = sb
        if ocr_r and rb and ocr_r != rb:
            izoh.append(f"OCR №{ocr_r} o'qidi" + (f", paket qoidasi (papka {n:03d} + {farq}) - №{qoida_r}"
                                                   if rb == qoida_r else f", dalillar bo'yicha №{rb}"))
        if ocr_s and sb and ocr_s != sb:
            izoh.append(f"OCR sanani {ocr_s:%d.%m.%Y} o'qidi, qo'shni arizalar bo'yicha {sb:%d.%m.%Y}")
        i.ocr_izoh = "; ".join(izoh)
        if avto and avto_ruxsat and not i.jadval_shubha:
            i.tekshirildi = True
            i.ocr_avto = True


def sana_oqish(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def jadval_oqish():
    if not os.path.isfile(JADVAL):
        return {}
    import openpyxl
    wb = openpyxl.load_workbook(JADVAL, data_only=True)
    ws = wb.active
    sarlavha = [str(c.value or "").strip() for c in ws[1]]
    qatorlar = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = {sarlavha[i]: row[i] for i in range(min(len(sarlavha), len(row)))}
        p = str(d.get("Papka") or "").strip()
        if p:
            qatorlar[p] = d
    return qatorlar


def jadval_yangilash():
    """mijozlar/ ichidagi har papka uchun ishlar.xlsx da qator bo'lishini ta'minlaydi."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    os.makedirs(MIJOZLAR_PAPKA, exist_ok=True)
    os.makedirs(UMUMIY_PAPKA, exist_ok=True)
    eski = jadval_oqish()
    ziplarni_ochish()
    ichki_papkalarni_chiqarish()
    papkalar = mijoz_papkalari()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ishlar"
    ws.append(JADVAL_USTUNLAR)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    yangi = 0
    for p in papkalar:
        pdf, docx_yol = davo_fayllari(os.path.join(MIJOZLAR_PAPKA, p))
        if p in eski:
            d = eski[p]
            sana = sana_oqish(d.get("Ariza sanasi"))
            qator = [p, str(d.get("Ariza raqami") or "").strip(), sana.strftime("%d.%m.%Y") if sana else d.get("Ariza sanasi"),
                     d.get("Javobgar"), d.get("JShShIR"), d.get("Asosiy qarz"), d.get("Pochta xarajati"), d.get("Izoh"),
                     d.get("Sud"), ("ha!" if str(d.get("Tekshirildi") or "").strip() == "ha!" else "ha")
                     if tekshirilganmi(d) else ""]
        else:
            yangi += 1
            # qo'lda yozilgan raqam/sana bu yerga yozilmaydi - OCR (fonda) o'qiydi, ishonchsizini odam tasdiqlaydi
            m = davo_ariza_malumoti(pdf, docx_yol, ocr=False) if (pdf or docx_yol) else {}
            qator = [p, "", "", m.get("javobgar", ""), m.get("jshshir", ""), m.get("asosiy_qarz", ""),
                     m.get("pochta", ""), "" if pdf else "DA'VO ARIZA PDF YO'Q", "", ""]
        ws.append(["" if v is None else str(v) for v in qator])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.number_format = "@"
    for i, w in enumerate([34, 14, 14, 42, 18, 14, 14, 30, 34, 12], start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "B2"
    try:
        wb.save(JADVAL)
    except PermissionError:
        print(f"!!! {JADVAL} saqlanmadi - Excel'da ochiq bo'lsa yoping va qayta urinib ko'ring.")
        return
    print(f"OK -> ishlar.xlsx: {len(papkalar)} ta mijoz ({yangi} ta yangi qo'shildi).")


def natijalar_oqish():
    """Har papka uchun OXIRGI natija: {papka: {'holat':..., 'havola':..., 'izoh':...}}."""
    holat = {}
    if os.path.isfile(NATIJALAR):
        with open(NATIJALAR, encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f, delimiter=";"):
                holat[r.get("papka", "")] = {"holat": r.get("holat", ""), "havola": r.get("havola", ""),
                                             "izoh": r.get("izoh", "") or ""}
    return holat


def natija_yozish(ish, holat, izoh="", havola="", rejim=""):
    yangi = not os.path.isfile(NATIJALAR)
    with open(NATIJALAR, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        if yangi:
            w.writerow(["vaqt", "papka", "holat", "izoh", "havola", "rejim"])
        w.writerow([datetime.now().strftime("%d.%m.%Y %H:%M:%S"), ish.papka, holat, izoh, havola, rejim])


def ishlarni_tayyorlash(filtr=None, ocr=True):
    """Mijozlar ro'yxati (hujjatlar, ma'lumotlar, kamchiliklar). ocr=False - OCR ishga tushmaydi, faqat avval
    o'qilgani (keshdan) olinadi: oyna tez ochilishi uchun; o'qilmaganlarini fon oqimi o'qiydi (ocr_oqilmaganlar)."""
    ziplarni_ochish()
    ichki_papkalarni_chiqarish()
    jadval = jadval_oqish()
    if any(p not in jadval for p in mijoz_papkalari()):   # yangi mijoz papkasi -> jadvalga o'zi qo'shiladi
        try:
            jadval_yangilash()
            jadval = jadval_oqish()
        except PermissionError:
            print("!!! ishlar.xlsx Excel'da ochiq - yangi mijozlar jadvalga qo'shilmadi (Excel'ni yoping).")
    if not os.path.isdir(MIJOZLAR_PAPKA):
        print(f"!!! {MIJOZLAR_PAPKA} papkasi yo'q.")
        return []
    umumiy = papka_hujjatlari(UMUMIY_PAPKA)
    ishlar = [ish_tayyorlash(p, jadval.get(p), umumiy, ocr) for p in mijoz_papkalari()
              if not filtr or filtr.lower() in p.lower()]
    raqam_sana_baholash(ishlar)
    return ishlar


def ish_tayyorlash(p, d, umumiy, ocr=True):
    """Bitta mijoz: hujjatlar, javobgar ma'lumotlari, kamchiliklar. d - ishlar.xlsx qatori (yo'q bo'lsa None).
    Qo'lda yozilgan raqam/sana bu yerda faqat o'qiladi - yakuniy qiymat raqam_sana_baholash() da."""
    tartib = {k: i for i, (k, *_r) in enumerate(HUJJAT_TURLARI)}
    nomlar = {k: n for k, n, *_r in HUJJAT_TURLARI}
    yol = os.path.join(MIJOZLAR_PAPKA, p)
    ish = Ish(papka=p, yol=yol)
    oziniki = papka_hujjatlari(yol)
    bor_turlar = {k for k, _y in oziniki}
    # mijoz papkasida o'zi bor bo'lgan turdagi umumiy hujjat qayta qo'shilmaydi (ikki marta yuklanmasin)
    hujjatlar = sorted(oziniki + [(k, y) for k, y in umumiy if k not in bor_turlar],
                       key=lambda kv: tartib.get(kv[0], 99))
    ish.hujjatlar = [(k, nomlar.get(k, "Boshqa hujjat"), y) for k, y in hujjatlar]

    davo = [y for k, y in hujjatlar if k == "davo"]
    pdf, docx_yol = davo_fayllari(yol)
    m = davo_ariza_malumoti(pdf, docx_yol, ocr=ocr) if (pdf or docx_yol) else {}
    ish.ariza_pdf = pdf
    d_bor, d = d, d or {}

    def ol(ustun, kalit):
        v = d.get(ustun)
        return str(v).strip() if v not in (None, "") else m.get(kalit, "")

    x = re.match(r"\s*(\d+)", str(d.get("Ariza raqami") or ""))
    ish.raqam = x.group(1) if x else ""
    ish.sana = sana_oqish(d.get("Ariza sanasi"))
    ish.tekshirildi = tekshirilganmi(d)
    ish.majburiy = ish.tekshirildi and str(d.get("Tekshirildi") or "").strip() == "ha!"
    ish.jadvaldagi = (ish.raqam, ish.sana, ish.tekshirildi)   # OCR qayta baholanganda shundan boshlanadi
    if pdf:   # tasdiqlanganlari ham - odam tasdiqlagani dalillarga zid emasligi tekshiriladi
        ish.qolyozma = qolyozma_tahlil(pdf, ocr=ocr)
    ish.javobgar = ol("Javobgar", "javobgar")
    ish.jshshir = re.sub(r"\D", "", ol("JShShIR", "jshshir"))
    ish.pasport = m.get("pasport", "")
    ish.asosiy_qarz = summa_tozalash(ol("Asosiy qarz", "asosiy_qarz"))
    ish.pochta = summa_tozalash(ol("Pochta xarajati", "pochta"))
    # Sud: ishlar.xlsx "Sud" ustunida qo'lda yozilgan bo'lsa - o'sha, aks holda javobgar manzilidan
    ish.manzil = m.get("manzil", "")
    if not manzil_toliqmi(ish.manzil):   # arizada faqat mahalla/ko'cha - shartnoma/talabnoma/pochtadagi to'liq manzil
        ish.manzil = next((z for z in hujjatlardan_malumot(yol).get("manzillar", []) if manzil_toliqmi(z)), ish.manzil)
    if sudlar():
        qolda = str(d.get("Sud") or "").strip()
        if qolda:
            ish.sud = sud_nomidan(qolda)
            if not ish.sud:
                ish.xatolar.append(f"ishlar.xlsx dagi sud nomi tanilmadi: '{qolda}'")
        else:
            ish.sud, sabab = None, ""
            # belgi qo'yilmagan bo'lsa: da'vo arizada yozilgan sud (kvitansiya ham shu sudga olingan) -
            # sud yozilmagan bo'lsa - javobgar manzili bo'yicha
            if not hammaga_toshkent() and m.get("sud_nomi"):
                ish.sud = sud_nomidan(m["sud_nomi"])
            if not ish.sud:
                ish.sud, sabab = avto_sud(ish.manzil)
            if not ish.sud:   # arizadagi manzilda tuman/shahar yo'q - shartnoma/talabnoma/pochtadagi to'liq manzil
                for zaxira in hujjatlardan_malumot(yol).get("manzillar", []):
                    sud, _s = avto_sud(zaxira)
                    if sud:
                        print(f"    (sud ariza manzilidan emas, hujjatdagi to'liq manzildan: {zaxira})")
                        ish.sud, ish.manzil = sud, zaxira
                        break
            if not ish.sud:
                ish.xatolar.append(f"sud aniqlanmadi ({sabab}; manzil: '{ish.manzil or '?'}') - "
                                   "qatorni 2 marta bosib 'Sud' ni yozing")

    if not davo:
        ish.xatolar.append("da'vo ariza PDF yo'q")
    if d_bor is None:
        ish.xatolar.append("ishlar.xlsx da qator yo'q (--jadval bilan yarating)")
    kv = [kvitansiya_oqish(y) for k, _n, y in ish.hujjatlar if k == "tolov" and y.lower().endswith(".pdf")]
    kv = [k for k in kv if k]
    # kvitansiyasiz yoki boshqa sudning kvitansiyasi bilan yuborilsa sud rad etadi - bunday mijoz ishlanmaydi
    sud_nomi = ish.sud["nomi"] if ish.sud else "mijoz sudi"
    if not any(k == "tolov" for k, _n, _y in ish.hujjatlar):
        ish.xatolar.append(f"to'lov kvitansiyasi yo'q - KvitansiyaYaratish.exe bilan {sud_nomi} ga yarating va papkaga qo'ying")
    elif ish.sud and kv and not any(kvitansiya_sudi_mos(k, ish.sud) for k in kv):
        ish.xatolar.append(f"kvitansiya boshqa sudga ({kv[0]['sud'].replace('Фуқаролик ишлари бўйича ', '')}) - "
                           f"KvitansiyaYaratish.exe bilan {sud_nomi} ga yarating")
    if len(ish.jshshir) != 14:
        ish.xatolar.append(f"JShShIR 14 raqam emas: '{ish.jshshir}'")
    if not ish.asosiy_qarz:
        ish.xatolar.append("asosiy qarz summasi yo'q")
    bor = {k for k, *_r in ish.hujjatlar}
    for k, nomi, *_r in HUJJAT_TURLARI:
        if k not in bor:
            ish.ogohlantirishlar.append(f"hujjat yo'q: {nomi}")
    for _k, _n, y in ish.hujjatlar:
        mb = os.path.getsize(y) / 1024 / 1024
        if mb > 5:
            ish.ogohlantirishlar.append(f"katta fayl ({mb:.1f} MB) - sayt qabul qilmasligi mumkin: {os.path.basename(y)}")
    return ish


def raqam_sana_baholash(ishlar):
    """Raqam/sana: ishlar.xlsx da odam tasdiqlagani, bo'lmasa OCR + qo'shnilar (qolyozma_yechim).
    Qayta chaqirsa bo'ladi (odam bitta mijozni tasdiqlagach qo'shnilari qayta baholanadi)."""
    for ish in ishlar:
        if ish.raqam_xato in ish.xatolar:
            ish.xatolar.remove(ish.raqam_xato)
        ish.raqam, ish.sana, ish.tekshirildi = ish.jadvaldagi
        ish.ocr_ishonch, ish.ocr_variantlar, ish.ocr_avto, ish.ocr_izoh = None, None, False, ""
        ish.tekshir_kerak, ish.raqam_xato, ish.ocr_tavsiya, ish.jadval_shubha = False, "", None, None
    qolyozma_yechim(ishlar)
    T = ocr_chegara()
    for ish in ishlar:
        if ish.raqam and ish.sana and ish.tekshirildi:
            continue
        ish.tekshir_kerak = True
        if ish.qolyozma is None and ish.ariza_pdf and not (ish.raqam and ish.sana):
            xato = "qo'lda yozilgan raqam/sana hali o'qilmagan (OCR navbatda)"
        elif not ish.raqam and not ish.sana:
            xato = "ariza RAQAMI va SANASI o'qilmadi - qatorni 2 marta bosib kiriting"
        else:
            rp, sp = ish.ocr_ishonch or (0.0, 0.0)
            qism = [f"raqam {ish.raqam or '?'}" + (f" ({rp:.0%})" if ish.ocr_ishonch and rp < T else ""),
                    f"sana {ish.sana.strftime('%d.%m.%Y') if ish.sana else '?'}"
                    + (f" ({sp:.0%})" if ish.ocr_ishonch and sp < T else "")]
            xato = "raqam/sana TEKSHIRING: " + ", ".join(qism) + (f" - {ish.ocr_izoh}" if ish.ocr_izoh else "")
        ish.raqam_xato = xato
        ish.xatolar.insert(0, xato)


def ocr_oqilmaganlar():
    """Hali OCR qilinmagan arizalar [(papka, pdf, docx)]: skaner matni yoki qo'lda yozilgan raqam/sana keshda yo'q.
    Avval tasdiqlanmaganlari (ro'yxatda kutib turganlar), keyin tasdiqlanganlari (tasdiq dalillarga zid emasmi)."""
    jadval = jadval_oqish()
    kerak, keyin = [], []
    for p in mijoz_papkalari():
        pdf, docx = davo_fayllari(os.path.join(MIJOZLAR_PAPKA, p))
        if not pdf:
            continue
        d = jadval.get(p) or {}
        tasdiqlangan = (str(d.get("Ariza raqami") or "").strip() and sana_oqish(d.get("Ariza sanasi"))
                        and tekshirilganmi(d))
        skaner = not docx and matnsizmi(pdf_matn(pdf))
        if (skaner and skan_matn(pdf, ocr=False) is None) or qolyozma_tahlil(pdf, ocr=False) is None:
            (keyin if tasdiqlangan else kerak).append((p, pdf, docx))
    return kerak + keyin


def ocr_bitta(pdf, docx=None):
    """Bitta arizani OCR qiladi (natija keshga): skaner matni (summalar uchun) va qo'lda yozilgan raqam/sana."""
    if not docx and matnsizmi(pdf_matn(pdf)):
        skan_matn(pdf)
    qolyozma_tahlil(pdf)


def ishlarni_korsatish(ishlar, oldingi):
    print("\n" + "=" * 70)
    print(f"MIJOZLAR: {len(ishlar)} ta")
    print("=" * 70)
    for i, ish in enumerate(ishlar, 1):
        h = oldingi.get(ish.papka, {}).get("holat", "")
        print(f"\n[{i}] {ish.papka}" + (f"   (oxirgi natija: {h})" if h else ""))
        print(f"    Ariza: №{ish.raqam or '?'}  sana: {ish.sana.strftime('%d.%m.%Y') if ish.sana else '?'}")
        print(f"    Javobgar: {ish.javobgar or '?'}   JShShIR: {ish.jshshir or '?'}")
        print(f"    Asosiy qarz: {ish.asosiy_qarz or '?'}   Pochta: {ish.pochta or '-'}")
        print(f"    Manzil: {ish.manzil or '?'}")
        print(f"    Sud: {ish.sud['hudud'] + ' / ' + ish.sud['nomi'] if ish.sud else '?'}")
        print(f"    Hujjatlar ({len(ish.hujjatlar)}):")
        for _k, nomi, y in ish.hujjatlar:
            print(f"       - {nomi:<38} {os.path.basename(y)}")
        for o in ish.ogohlantirishlar:
            print(f"    ! {o}")
        for x in ish.xatolar:
            print(f"    !!! XATO: {x}")
    print()


# ============================================================
# BRAUZER YORDAMCHILARI (JavaScript)
# ============================================================
JS = r"""
const norm = s => (s || '').toString().toLowerCase()
  .replace(/[`'’‘ʻʼ"«»]/g, '').replace(/қ/g, 'к').replace(/ҳ/g, 'х').replace(/ў/g, 'у').replace(/ғ/g, 'г')
  .replace(/\s+/g, ' ').trim();
const vis = e => {
  if (!e || !e.getBoundingClientRect) return false;
  const r = e.getBoundingClientRect();
  if (r.width < 1 || r.height < 1) return false;
  const s = getComputedStyle(e);
  return s.visibility !== 'hidden' && s.display !== 'none' && parseFloat(s.opacity || '1') > 0.05;
};
const txt = e => norm(e.innerText || '');
const OVERLAY = '.toast,.toast-container,.swal2-container,[role=alert],.notification,.ant-message,.snackbar,mat-snack-bar-container,.p-toast';
const DIALOG = '[role=dialog],.modal.show,.modal-dialog,.swal2-popup,mat-dialog-container,.p-dialog,.ant-modal';
const VARIANT = 'mat-option,[role=option],.ng-option,.dropdown-item,.p-dropdown-item,.ant-select-item-option';
function topish(root, text, sel, exact, last) {
  const t = norm(text); let best = null, bl = 1e12;
  for (const e of (root || document).querySelectorAll(sel || '*')) {
    if (!vis(e)) continue;
    const it = txt(e);
    if (exact ? it !== t : !it.includes(t)) continue;
    if (last) { best = e; continue; }
    if (it.length < bl) { best = e; bl = it.length; }
  }
  return best;
}
function inp(ph) {
  const t = norm(ph);
  for (const e of document.querySelectorAll('input,textarea')) {
    if (vis(e) && norm(e.getAttribute('placeholder')).includes(t)) return e;
  }
  return null;
}
function imzo() {
  const p = [location.pathname];
  for (const e of document.querySelectorAll('input,textarea,select,mat-select,legend,h1,h2,h3,h4,h5,button')) {
    if (!vis(e) || e.closest(OVERLAY)) continue;
    p.push(e.tagName + ':' + norm(e.getAttribute('placeholder') || e.innerText || '').slice(0, 30));
  }
  return p.join('|');
}
function xatolar() {
  const r = [];
  for (const e of document.querySelectorAll('mat-error,.mat-error,.mat-mdc-form-field-error,.error,.invalid-feedback,.text-danger,.toast,.toast-message,.swal2-html-container,.swal2-title,.alert,[role=alert],mat-snack-bar-container,.ant-message,.notification,.p-toast-message')) {
    if (!vis(e)) continue;
    const t = (e.innerText || '').trim();
    if (t && !r.includes(t)) r.push(t.slice(0, 200));
  }
  for (const e of document.querySelectorAll('input.ng-invalid,textarea.ng-invalid,mat-select.ng-invalid')) {
    if (vis(e)) r.push('Noto\'g\'ri/bo\'sh maydon: ' + (e.getAttribute('placeholder') || e.name || e.id || e.tagName));
  }
  return r;
}
function yorliq(e) {
  if (e.getAttribute('aria-label')) return e.getAttribute('aria-label');
  if (e.id) { const l = document.querySelector('label[for="' + e.id + '"]'); if (l) return l.innerText.trim(); }
  let c = e;
  for (let i = 0; i < 4 && c.parentElement; i++) {
    c = c.parentElement;
    const t = (c.innerText || '').trim();
    if (t.length > 1 && t.length < 160) return t.replace(/\s+/g, ' ');
  }
  return '';
}
function maydon_yorliq(label) {
  // "Hujjat turi", "To'lov imtiyozi" kabi yozuv yonidagi ro'yxat/maydon
  const t = norm(label);
  const SEL = 'mat-select,[role=combobox],ng-select,.ng-select,select,.p-dropdown,.ant-select,[aria-haspopup],[class*=select],[class*=dropdown],input:not([type=file]):not([type=checkbox]):not([type=radio])';
  for (const l of document.querySelectorAll('label,span,div,p,td,th,legend,b,strong,mat-label')) {
    if (!vis(l) || txt(l) !== t) continue;
    let c = l;
    for (let i = 0; i < 4 && c; i++) {
      for (const s of c.querySelectorAll(SEL)) if (vis(s)) return s;
      c = c.parentElement;
    }
  }
  return null;
}
function variantlar() {
  const r = [];
  for (const e of document.querySelectorAll(VARIANT + ',option')) {
    if (e.tagName !== 'OPTION' && !vis(e)) continue;
    const s = (e.innerText || e.textContent || '').trim().replace(/\s+/g, ' ');
    if (s && !r.includes(s)) r.push(s);
  }
  return r;
}
function variant(text) {
  const t = norm(text); let best = null, bl = 1e9;
  for (const e of document.querySelectorAll(VARIANT)) {
    if (!vis(e)) continue;
    const s = txt(e);
    if (s.includes(t) && s.length < bl) { best = e; bl = s.length; }
  }
  return best;
}
function fayl_inputlar() {
  // Sahifadagi fayl maydonlari (yashirinlari ham - saytda ular odatda ko'rinmaydi)
  const r = [];
  for (const e of document.querySelectorAll('input[type=file]')) {
    if (e.hasAttribute('data-adolat-fayl')) continue;    // "Yuklash" bosilganda ushlangan - alohida olinadi
    let c = e, ctx = '', tomon = '';
    for (let i = 0; i < 12 && c.parentElement; i++) {
      c = c.parentElement;
      const t = txt(c);
      if (!ctx && t.length > 3) ctx = (c.innerText || '').trim();
      const il = t.includes('hujjat turi'), ar = t.includes('arizani biriktirish');
      if (il || ar) { tomon = il && ar ? '' : (il ? 'ilova' : 'ariza'); break; }
    }
    r.push({el: e, multiple: !!e.multiple, accept: e.accept || '', ctx: ctx.slice(0, 300), tomon: tomon});
  }
  return r;
}
function dialogsiz() {
  // Windows fayl tanlash oynasi ochilmasin: sayt "Yuklash" bosilganda yaratadigan fayl maydonini
  // ushlab qolamiz (window.__adolat_oxirgi) - fayl unga Selenium orqali beriladi
  for (const i of document.querySelectorAll('input[type=file]')) { i.click = function () {}; i.showPicker = function () {}; }
  if (!window.__adolat_tarmoq) {
    // fayl yuborayotgan so'rovlar (FormData/Blob): nechtasi ketyapti, nechtasi tugadi, nechtasi xato -
    // fayl sahifada ko'rinishi yetmaydi, server javob berguncha kutamiz
    const T = window.__adolat_tarmoq = {yuk_faol: 0, yuk_boshlandi: 0, yuk_tugadi: 0, yuk_xato: 0};
    const fayl = b => (typeof FormData !== 'undefined' && b instanceof FormData) || (typeof Blob !== 'undefined' && b instanceof Blob);
    const tugadi = s => { T.yuk_faol = Math.max(0, T.yuk_faol - 1); if (s >= 200 && s < 300) T.yuk_tugadi++; else T.yuk_xato++; };
    const send = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.send = function (body) {
      if (fayl(body)) { T.yuk_faol++; T.yuk_boshlandi++; this.addEventListener('loadend', () => tugadi(this.status)); }
      return send.apply(this, arguments);
    };
    const f = window.fetch;
    if (f) window.fetch = function (url, init) {
      if (!init || !fayl(init.body)) return f.apply(this, arguments);
      T.yuk_faol++; T.yuk_boshlandi++;
      return f.apply(this, arguments).then(r => { tugadi(r.status); return r; }, e => { tugadi(0); throw e; });
    };
  }
  if (!window.__adolat_dialogsiz) {
    window.__adolat_dialogsiz = true;
    const ushla = el => {
      if (!el.isConnected) { el.style.display = 'none'; document.body.appendChild(el); }
      el.setAttribute('data-adolat-fayl', '1');
      window.__adolat_oxirgi = el;
    };
    const click = HTMLInputElement.prototype.click;
    HTMLInputElement.prototype.click = function () {
      if ((this.type || '').toLowerCase() === 'file') return ushla(this);
      return click.apply(this, arguments);
    };
    const picker = HTMLInputElement.prototype.showPicker;
    if (picker) HTMLInputElement.prototype.showPicker = function () {
      if ((this.type || '').toLowerCase() === 'file') return ushla(this);
      return picker.apply(this, arguments);
    };
  }
  return true;
}
function tomon_ildiz(tomon) {
  // Chap ("Arizani biriktirish") yoki o'ng ("Hujjat turi") qism - ichida "Yuklash" bor eng kichik blok
  const kalit = tomon === 'ariza' ? 'arizani biriktirish' : 'hujjat turi';
  const boshqa = tomon === 'ariza' ? 'hujjat turi' : 'arizani biriktirish';
  let best = null, bl = 1e12;
  for (const e of document.querySelectorAll('fieldset,section,form,mat-card,div')) {
    if (!vis(e)) continue;
    const t = txt(e);
    if (!t.includes(kalit) || t.includes(boshqa) || !t.includes('yuklash')) continue;
    if (t.length < bl) { best = e; bl = t.length; }
  }
  return best;
}
function yuklash_tugmasi(tomon) {
  const r = tomon_ildiz(tomon);
  if (!r) return null;
  let t = null;
  for (const x of r.querySelectorAll('button,[role=button]')) if (vis(x) && !x.disabled && txt(x).includes('yuklash')) t = x;
  return t || topish(r, 'fayl yuklash', null, false, false) || topish(r, 'yuklash', null, false, false);
}
function fayl_maydoni(tomon) {
  // bitta chaqiruvda: dialogsiz + shu tomonning tayyor fayl maydoni yoki uning "Yuklash" tugmasi
  dialogsiz();
  window.__adolat_oxirgi = null;
  for (const f of fayl_inputlar()) if (f.tomon === tomon) return {el: f.el};
  return {tugma: yuklash_tugmasi(tomon)};
}
const BAND = 'mat-spinner,mat-progress-spinner,.mat-progress-spinner,.mat-mdc-progress-spinner,mat-progress-bar,.mat-progress-bar,.mat-mdc-progress-bar,.spinner-border,.spinner-grow,ngx-spinner,[role=progressbar]';
function band() {
  // sayt "yuklanmoqda" (spinner / progress) ko'rsatyaptimi
  for (const s of document.querySelectorAll(BAND)) if (vis(s)) return true;
  return false;
}
function ochiq_variantlar() {
  let n = 0;
  for (const e of document.querySelectorAll(VARIANT)) if (vis(e)) n++;
  return n;
}
function royxat_ochiqmi(el) {
  // ro'yxat (mat-select) paneli ochiqmi; el = null - faqat panel bo'yicha
  if (el && el.isConnected && el.getAttribute('aria-expanded') === 'true') return true;
  for (const p of document.querySelectorAll('.mat-select-panel,.mat-mdc-select-panel,.ng-dropdown-panel,.p-dropdown-panel,.ant-select-dropdown')) {
    if (vis(p)) return true;
  }
  return ochiq_variantlar() > 0;
}
function royxat_trigger(el) {
  // mat-select o'zi JS click'da ochilmaydi - ichidagi trigger bosilishi kerak
  return (el && el.querySelector && el.querySelector('.mat-select-trigger,.mat-mdc-select-trigger,.ng-select-container,.p-dropdown-trigger')) || el;
}
function yuklash_holati(stem) {
  // fayl nomi sahifada necha marta ko'rinadi + yuklangan fayllar qatorlari ("0.27 mb") soni + spinner
  const m = norm(document.body.innerText);
  let nom = 0;
  if (stem) {
    for (let i = m.indexOf(stem); i >= 0; i = m.indexOf(stem, i + 1)) nom++;
    for (const i of document.querySelectorAll('input')) if (i.type !== 'file' && norm(i.value).includes(stem)) nom++;
  }
  const hajm = (m.match(/\d+(?:[.,]\d+)?\s?(?:kb|mb|gb|кб|мб|гб)(?![a-zа-яё])/g) || []).length;
  return Object.assign({nom: nom, hajm: hajm, band: band()}, window.__adolat_tarmoq || {});
}
function kvitansiya_plus() {
  // «+» saytda <button> emas (kursori ham oddiy) - ekranda "...tugmasi orqali biriktiring." matnining
  // darhol O'NG tomonida turgan elementni olamiz. Ro'yxatlar (mat-select) hech qachon olinmaydi.
  const kalit = topish(null, 'tugmasi orqali', null, false, false) ? 'tugmasi orqali' : 'kvitantsiya mavjud';
  let t = topish(null, kalit, null, false, false);
  if (!t) return null;
  // matnning o'zi (eng ichki element) - «+» bilan birga o'ralgan blok emas
  for (let o = true; o;) {
    o = false;
    for (const ch of t.children) if (txt(ch).includes(kalit)) { t = ch; o = true; break; }
  }
  const SELECT = 'mat-select,.mat-select,.mat-mdc-select,.mat-select-trigger,.mat-mdc-select-trigger,[role=combobox],[role=listbox]';
  const TAQIQ = /keyingi|orqaga|qoralama|chiqish/;
  t.scrollIntoView({block: 'center'});
  const r = t.getBoundingClientRect(), markaz = r.top + r.height / 2;
  const yonida = x => { const q = x.getBoundingClientRect(); return Math.abs(q.top + q.height / 2 - markaz) < 60; };
  for (const y of [r.top + r.height / 2, r.top + r.height * 0.25, r.top + r.height * 0.75]) {
    for (let dx = 6; dx <= 140; dx += 6) {
      let e = document.elementFromPoint(r.right + dx, y);
      while (e && !(e instanceof HTMLElement)) e = e.parentElement;   // svg/path -> HTML ota
      if (!e || e === document.body || e === document.documentElement || e.contains(t) || e.closest(SELECT) || TAQIQ.test(txt(e))) continue;
      // 80px gacha bo'lgan eng tashqi ota - butun tugma
      let c = e;
      while (c.parentElement && !c.parentElement.contains(t)) {
        const pr = c.parentElement.getBoundingClientRect();
        if (pr.width > 80 || pr.height > 80) break;
        c = c.parentElement;
      }
      return c;
    }
  }
  // zaxira: matn atrofidagi kichik bosiladigan elementlar (eng o'ngdagisi)
  let c = t;
  for (let i = 0; i < 6 && c; i++) {
    const n = [...c.querySelectorAll('button,[role=button],a,mat-icon,i,img,span,div')].filter(x => {
      if (!vis(x) || x.contains(t) || t.contains(x) || x.closest(SELECT) || TAQIQ.test(txt(x)) || !yonida(x)) return false;
      const r = x.getBoundingClientRect();
      if (r.width > 80 || r.height > 80) return false;
      const s = (x.innerText || '').trim();
      return x.tagName === 'BUTTON' || s === '+' || s === 'add' || getComputedStyle(x).cursor === 'pointer';
    });
    if (n.length) {
      const R = x => x.getBoundingClientRect();
      n.sort((a, b) => (R(b).right - R(a).right) || (R(b).width * R(b).height - R(a).width * R(a).height));
      return n[0];
    }
    c = c.parentElement;
  }
  return null;
}
function robot_quti() {
  // Faqat "Robot emasman" ning o'z qutisi: yuqoriga boshqa katak uchramaguncha chiqamiz,
  // aks holda yonidagi "Men JShShIR ni bilmayman" / "Xorijiy tashkilot" katagi ham olinib qoladi
  const box = topish(document, 'robot emasman', null, false, false);
  if (!box) return null;
  const KATAK = 'input[type=checkbox],[role=checkbox],mat-checkbox,[aria-checked]';
  const begona = k => {   // o'z yozuvi (bir necha so'z) bor boshqa katak; ikonka nomi (check_box) hisobga olinmaydi
    const t = txt(k.closest('mat-checkbox,label,[role=checkbox]') || k);
    return /\S\s+\S/.test(t) && !t.includes('robot emasman');
  };
  let c = box;
  while (c.parentElement) {
    const p = c.parentElement, t = txt(p);
    if (!t.includes('robot emasman') || t.length > 60) break;
    if ([...p.querySelectorAll(KATAK)].some(k => !c.contains(k) && begona(k))) break;
    c = p;
  }
  return c;
}
function robot_imzo() {
  // qutining ko'rinishi: katak belgilanganda (ikonka/klass) o'zgaradi
  const c = robot_quti();
  return c ? c.outerHTML.replace(/\s+/g, ' ').slice(0, 3000) : '';
}
function robot_holati() {
  // 'yoq' - CAPTCHA yo'q | 'misol' - matematik misol ochiq | 'yuklanmoqda' | 'kutilmoqda' - hali belgilanmagan
  // 'belgilandi' - "Robot emasman" katagi belgilangan (yechilganini Python vaqt bo'yicha tasdiqlaydi)
  for (const e of document.querySelectorAll('input,textarea')) {
    const p = norm(e.getAttribute('placeholder'));
    if (vis(e) && p.includes('javob') && !p.includes('javobgar')) return 'misol';
  }
  const c = robot_quti();
  if (!c) return 'yoq';
  for (const s of document.querySelectorAll('mat-spinner,mat-progress-spinner,.mat-progress-spinner,.mat-mdc-progress-spinner,.spinner-border,[role=progressbar]')) {
    if (vis(s)) return 'yuklanmoqda';
  }
  const cb = c.querySelector('input[type=checkbox]');
  if (cb) return cb.checked ? 'belgilandi' : 'kutilmoqda';
  const rc = c.hasAttribute('aria-checked') ? c : c.querySelector('[aria-checked]');
  if (rc) return rc.getAttribute('aria-checked') === 'true' ? 'belgilandi' : 'kutilmoqda';
  const h = (c.outerHTML || '').toLowerCase().replace(/unchecked|aria-checked="false"/g, '');
  return /(checked|verified|success|tasdiqlandi)/.test(h) ? 'belgilandi' : 'kutilmoqda';
}
function asosiy_tugma() {
  // Sahifa pastidagi o'ngdagi asosiy tugma (Orqaga / Qoralamalarga saqlash emas)
  let best = null, bb = -1, br = -1;
  for (const b of document.querySelectorAll('button')) {
    if (!vis(b) || b.disabled || b.closest(OVERLAY) || b.closest('nav,aside,header,.sidebar,mat-sidenav')) continue;
    const t = txt(b);
    if (t.includes('orqaga') || t.includes('qoralama') || t.includes('yangilash') || t.includes('chiqish')) continue;
    const r = b.getBoundingClientRect();
    if (r.bottom > bb + 5 || (Math.abs(r.bottom - bb) <= 5 && r.right > br)) { best = b; bb = r.bottom; br = r.right; }
  }
  return best;
}
function tahlil() {
  const o = {url: location.href, title: document.title, inputs: [], buttons: [], files: [], sarlavhalar: [], xatolar: xatolar(), variantlar: variantlar(), matn: ''};
  for (const e of document.querySelectorAll('input,textarea,select,mat-select,[role=combobox]')) {
    if ((e.type || '').toLowerCase() === 'file' || !vis(e)) continue;
    o.inputs.push({tag: e.tagName, type: e.type || '', ph: e.getAttribute('placeholder') || '', name: e.name || '',
                   id: e.id || '', value: (e.value !== undefined ? e.value : (e.innerText || '')).toString().slice(0, 80),
                   dis: !!e.disabled, ro: !!e.readOnly, checked: !!e.checked, label: yorliq(e).slice(0, 120)});
  }
  for (const e of document.querySelectorAll('button,[role=button],a.btn,input[type=submit]')) {
    if (!vis(e)) continue;
    const t = (e.innerText || e.value || e.getAttribute('aria-label') || '').trim().replace(/\s+/g, ' ');
    o.buttons.push((e.disabled ? '[o\'chiq] ' : '') + (t || '(matnsiz:' + (e.className || '').toString().slice(0, 40) + ')'));
  }
  for (const f of fayl_inputlar()) o.files.push({multiple: f.multiple, accept: f.accept, tomon: f.tomon, ctx: f.ctx.replace(/\s+/g, ' ')});
  for (const e of document.querySelectorAll('legend,h1,h2,h3,h4,h5,label,.mat-step-label,.title')) {
    if (!vis(e)) continue;
    const t = (e.innerText || '').trim().replace(/\s+/g, ' ');
    if (t && t.length < 120 && !o.sarlavhalar.includes(t)) o.sarlavhalar.push(t);
  }
  o.matn = (document.body.innerText || '').slice(0, 20000);
  return o;
}
"""


def norm_py(s):
    """JS norm() bilan bir xil: kichik harf, tutuq belgisiz, қ->к ҳ->х ў->у ғ->г."""
    s = re.sub(r"[`'’‘ʻʼ\"«»]", "", str(s or "").lower())
    s = s.replace("қ", "к").replace("ҳ", "х").replace("ў", "у").replace("ғ", "г")
    return re.sub(r"\s+", " ", s).strip()


def xpath_literal(s):
    if "'" not in s:
        return f"'{s}'"
    if '"' not in s:
        return f'"{s}"'
    return "concat(" + ", \"'\", ".join(f"'{p}'" for p in s.split("'")) + ")"


def xavfsiz_nom(s):
    return re.sub(r"[^\w\-]+", "_", s)[:60]


# Oddiy Chrome: dastur chrome.exe ni o'zi ochadi (avtomatlashtirish bayroqlarisiz), Selenium keyin ulanadi.
# Profil doimiy - cookie, ONE ID sessiyasi, o'rnatilgan kengaytmalar (E-IMZO) saqlanib qoladi.
CHROME_PORT = 9241
CHROME_PROFIL = os.path.join(os.environ.get("LOCALAPPDATA") or PAPKA, "AdolatAvtomat", "chrome_profil")


def chrome_yoli():
    """O'rnatilgan Google Chrome (chrome.exe) yo'li yoki None."""
    nomzodlar = []
    try:
        import winreg
        for ildiz in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            try:
                with winreg.OpenKey(ildiz, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe") as k:
                    nomzodlar.append(winreg.QueryValue(k, None))
            except OSError:
                pass
    except ImportError:
        pass
    for o in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        if os.environ.get(o):
            nomzodlar.append(os.path.join(os.environ[o], "Google", "Chrome", "Application", "chrome.exe"))
    return next((y for y in nomzodlar if y and os.path.isfile(y)), None)


def chrome_javob_beradimi():
    """CHROME_PORT da ochiq Chrome bormi (proksisiz tekshiriladi)."""
    import urllib.request
    try:
        ochuvchi = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with ochuvchi.open(f"http://127.0.0.1:{CHROME_PORT}/json/version", timeout=1) as r:
            return r.status == 200
    except Exception:
        return False


class Kutish(WebDriverWait):
    """WebDriverWait, lekin 'To'xtatish' bosilsa darhol uziladi."""
    def until(self, method, message=""):
        def m(d):
            toxtat_tekshir()
            return method(d)
        return super().until(m, message)


class Brauzer:
    def __init__(self, fon=False):
        self.chrome = None      # biz ochgan chrome.exe jarayoni
        self.ulangan = False    # oddiy Chrome'ga ulanganmi (yopishda kerak)
        self.d = self._oddiy_chrome(fon) or self._selenium_chrome(fon)
        if not fon:
            try:
                self.d.maximize_window()
            except WebDriverException:
                pass
        self.wait = Kutish(self.d, KUTISH)
        self.prefix = SESSIYA
        self.n = 0
        self.robot_toza = {}    # CAPTCHA: belgilanmagan paytdagi qutining ko'rinishi (qidiruv nomi bo'yicha)
        self.uzildi = False     # To'xtatish bosildi - Chrome va chromedriver o'ldirilgan
        global JORIY_BRAUZER
        JORIY_BRAUZER = self
        if TOXTAT.is_set():     # Chrome ochilayotganda To'xtatish bosilgan
            self.uzish()
            raise Toxtatildi()

    def uzish(self):
        """To'xtatish: chromedriver va Chrome jarayonlarini darhol o'ldiradi."""
        import subprocess
        self.uzildi = True
        time.sleep(1)                        # Browser.close o'zi yopishga ulgursin
        pidlar = []
        try:
            pr = self.d.service.process
            if pr and pr.poll() is None:
                pidlar.append(pr.pid)
        except Exception:
            pass
        if self.chrome is not None and self.chrome.poll() is None:
            pidlar.append(self.chrome.pid)
        for pid in pidlar:
            try:
                subprocess.run(["taskkill", "/T", "/F", "/PID", str(pid)], capture_output=True, timeout=10,
                               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            except Exception:
                pass

    def _oddiy_chrome(self, fon):
        """Oddiy Chrome'ni ochib unga ulanadi. Bo'lmasa None (unda Selenium o'zi ochadi)."""
        if chrome_javob_beradimi():
            print("    (avvalgi Chrome oynasi ochiq - shunga ulanilmoqda)")
        else:
            yol = chrome_yoli()
            if not yol:
                print("    (Google Chrome topilmadi - Selenium Chrome ishlatiladi)")
                return None
            import subprocess
            os.makedirs(CHROME_PROFIL, exist_ok=True)
            arg = [yol, f"--remote-debugging-port={CHROME_PORT}", f"--user-data-dir={CHROME_PROFIL}",
                   "--no-first-run", "--no-default-browser-check", "--disable-notifications"]
            arg += ["--headless=new", "--window-size=1600,1000"] if fon else ["--start-maximized"]
            arg.append("about:blank")
            try:
                self.chrome = subprocess.Popen(arg, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except OSError as e:
                print(f"    (Chrome ochilmadi: {e} - Selenium Chrome ishlatiladi)")
                return None
            oxiri = time.time() + 30
            while not chrome_javob_beradimi():
                if time.time() > oxiri or self.chrome.poll() is not None:
                    print("    (Chrome ulanishga tayyor bo'lmadi - Selenium Chrome ishlatiladi)")
                    self._chrome_toxtat()
                    return None
                if TOXTAT.is_set():
                    self._chrome_toxtat()
                    raise Toxtatildi()
                time.sleep(0.5)
        o = webdriver.ChromeOptions()
        o.debugger_address = f"127.0.0.1:{CHROME_PORT}"
        try:
            d = webdriver.Chrome(options=o)
        except WebDriverException as e:
            print(f"    (Chrome'ga ulanib bo'lmadi: {str(e).splitlines()[0]} - Selenium Chrome ishlatiladi)")
            self._chrome_toxtat()
            return None
        self.ulangan = True
        print("    OK -> Oddiy Chrome ochildi (doimiy profil, avtomatlashtirish belgilarisiz)")
        return d

    def _selenium_chrome(self, fon):
        """Zaxira: Selenium o'zi ochadigan Chrome (vaqtinchalik profil)."""
        o = webdriver.ChromeOptions()
        if fon:
            o.add_argument("--headless=new")
            o.add_argument("--window-size=1600,1000")
        o.add_argument("--disable-notifications")
        o.add_experimental_option("excludeSwitches", ["enable-logging"])
        return webdriver.Chrome(options=o)

    def _chrome_toxtat(self):
        if self.chrome is not None and self.chrome.poll() is None:
            try:
                self.chrome.terminate()
                self.chrome.wait(10)
            except Exception:
                pass
        self.chrome = None

    def ajrat(self):
        """Chrome'ni YOPMASDAN dasturni undan uzadi (TEST: brauzerda ko'rib chiqish uchun)."""
        if self.uzildi:
            return
        self.uzildi = True
        try:
            self.d.quit()                    # ulanish rejimida Chrome yopilmaydi
        except Exception:
            pass

    def yop(self):
        """Brauzerni yopadi. Ulanish rejimida driver.quit() Chrome'ni yopmaydi - shuning uchun Browser.close."""
        if self.uzildi:          # To'xtatish allaqachon yopgan
            return
        if self.ulangan:
            try:
                self.d.execute_cdp_cmd("Browser.close", {})
            except Exception:
                pass
        try:
            self.d.quit()
        except Exception:
            pass
        if self.chrome is not None:
            try:
                self.chrome.wait(10)
            except Exception:
                pass
            self._chrome_toxtat()

    def js(self, code, *args):
        toxtat_tekshir()
        try:
            return self.d.execute_script(JS + "\n" + code, *args)
        except NoSuchWindowException:
            if not self.oyna_tikla():
                raise
            return self.d.execute_script(JS + "\n" + code, *args)

    def oyna_tikla(self):
        """Joriy oyna/tab yopilgan (sayt yangi tabda ochgan) - qolgan oxirgi tabga o'tadi."""
        try:
            h = self.d.window_handles
            if h:
                self.d.switch_to.window(h[-1])
                return True
        except WebDriverException:
            pass
        return False

    def kut(self, fn, vaqt=None, oraliq=0.4):
        oxiri = time.time() + (KUTISH if vaqt is None else vaqt)
        while True:
            try:
                r = fn()
                if r:
                    return r
            except (WebDriverException, StaleElementReferenceException) as e:
                if brauzer_uzildimi(e):        # Chrome yopilgan/uzilgan - 60 soniya behuda kutilmaydi
                    raise
                if isinstance(e, NoSuchWindowException):
                    self.oyna_tikla()
            if time.time() > oxiri:
                return None
            uxla(oraliq)

    def top(self, text, sel=None, exact=False, last=False):
        return self.js("return topish(null, arguments[0], arguments[1], arguments[2], arguments[3]);",
                       text, sel, exact, last)

    def bos_el(self, el, pauza=1.0):
        """pauza - odamdek pauzalar ko'paytiruvchisi (natijasi alohida tekshiriladigan joylarda 0.5)."""
        try:
            self.d.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            odamdek_pauza(0.4 * pauza)
            el.click()
        except (ElementClickInterceptedException, ElementNotInteractableException):
            # ustini toast/oyna yopib turibdi - JS bilan; mat-select o'zi emas, ichidagi trigger bosiladi
            self.d.execute_script(
                "const e = arguments[0]; (e.querySelector('.mat-select-trigger,.mat-mdc-select-trigger') || e).click();", el)
        odamdek_pauza(pauza)

    def band_kut(self, vaqt=10):
        """Sayt spinneri ("yuklanmoqda") ko'rinsa - yo'qolguncha (ko'pi bilan `vaqt` soniya) kutadi."""
        return self.kut(lambda: not self.js("return band();"), vaqt, 0.15)

    def royxat_och(self, el):
        """Ro'yxatni (mat-select) ochadi va panel HAQIQATAN ochilganini tekshiradi. Oddiy bosish o'tmasa
        (sayt band / ustini toast yopgan) - spinner tugashini kutib trigger'ni JS bilan, so'ng klaviatura bilan."""
        ochiq = lambda: self.js("return royxat_ochiqmi(arguments[0]);", el)
        try:
            teg = el.tag_name.lower()
            for usul in range(3):
                if usul:
                    if ochiq():          # oldingi bosish kech bo'lsa ham ochgan
                        return True
                    self.band_kut()
                    if ochiq():
                        return True
                if usul == 0:
                    self.bos_el(el, 0.5)
                elif usul == 1:
                    self.d.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    self.js("royxat_trigger(arguments[0]).click(); return true;", el)
                else:
                    el.send_keys(Keys.ARROW_DOWN if teg in ("input", "textarea") else Keys.SPACE)
                if self.kut(ochiq, 1.5, 0.1):
                    return True
        except StaleElementReferenceException:
            pass                         # ro'yxat qayta chizildi - tanla_yorliq uni qaytadan topadi
        except WebDriverException as e:
            if brauzer_uzildimi(e):
                raise
        return False

    def bos(self, text, nom=None, vaqt=None, majburiy=True, sel=None, exact=False, last=False):
        nom = nom or text
        for _ in range(3):
            el = self.kut(lambda: self.top(text, sel, exact, last), vaqt)
            if not el:
                break
            try:
                self.bos_el(el)
                print(f"    OK -> {nom}")
                return el
            except StaleElementReferenceException:
                tin(0.5)
        if majburiy:
            print(f"    !!! TOPILMADI -> {nom}")
            raise Topilmadi(nom)
        return None

    def tugma(self, text, nom=None, vaqt=None, majburiy=True):
        """Ko'rinib turgan, matnida `text` bor OXIRGI tugmani bosadi."""
        return self.bos(text, nom, vaqt, majburiy, sel="button,[role=button],a.btn", last=True)

    def inp(self, ph, vaqt=None):
        return self.kut(lambda: self.js("return inp(arguments[0]);", ph), vaqt)

    def tozala(self, el):
        """Maydonni bo'shatadi (fokus maydonda bo'lishi kerak)."""
        el.send_keys(Keys.CONTROL, "a")
        el.send_keys(Keys.DELETE)
        tin(0.15)
        if el.get_attribute("value"):   # Ctrl+A ishlamay "a" yozilib qolishi mumkin - boshqa usulda tozalaymiz
            try:
                el.clear()
            except WebDriverException:
                pass
            if el.get_attribute("value"):
                self.js_yoz(el, "")

    def yoz(self, el, qiymat):
        self.bos_el(el)
        self.tozala(el)
        for ch in str(qiymat):
            el.send_keys(ch)
            harf_pauza()
        odamdek_pauza(0.5)
        return (el.get_attribute("value") or "").strip()

    def js_yoz(self, el, qiymat):
        self.d.execute_script(
            """
            const i = arguments[0];
            const proto = i.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
            Object.getOwnPropertyDescriptor(proto, 'value').set.call(i, arguments[1]);
            for (const t of ['focus', 'input', 'change', 'blur']) i.dispatchEvent(new Event(t, {bubbles: true}));
            """, el, str(qiymat))
        tin(0.2)
        return (el.get_attribute("value") or "").strip()

    def imzo(self):
        try:
            return self.js("return imzo();")
        except WebDriverException:
            return ""

    def xatolar(self):
        try:
            return self.js("return xatolar();") or []
        except WebDriverException:
            return []

    def matn(self):
        try:
            return self.js("return norm(document.body.innerText);") or ""
        except WebDriverException:
            return ""

    def esc(self):
        try:
            self.d.switch_to.active_element.send_keys(Keys.ESCAPE)
        except Exception:
            pass

    def alert_yop(self):
        try:
            self.d.switch_to.alert.accept()
        except (NoAlertPresentException, WebDriverException):
            pass

    def tahlil(self, nom):
        """Joriy sahifa: skrinshot + barcha maydon/tugmalar ro'yxati (loglar/ ga)."""
        self.n += 1
        asos = os.path.join(LOG_PAPKA, f"{self.prefix}_{self.n:02d}_{xavfsiz_nom(nom)}")
        try:
            self.d.save_screenshot(asos + ".png")
        except Exception:
            pass
        try:
            o = self.js("return tahlil();")
        except Exception as e:
            print(f"    (sahifa tahlili xatosi: {e})")
            return
        with open(asos + ".txt", "w", encoding="utf-8") as f:
            f.write(f"URL: {o['url']}\nTITLE: {o['title']}\n\n== XATOLAR ==\n")
            f.writelines(f"  {x}\n" for x in o["xatolar"])
            f.write("\n== SARLAVHALAR / LABELLAR ==\n")
            f.writelines(f"  {x}\n" for x in o["sarlavhalar"])
            f.write("\n== MAYDONLAR ==\n")
            for i in o["inputs"]:
                f.write(f"  <{i['tag']} type={i['type']}> ph='{i['ph']}' name='{i['name']}' id='{i['id']}' "
                        f"value='{i['value']}' dis={i['dis']} ro={i['ro']} checked={i['checked']} | {i['label']}\n")
            f.write("\n== TUGMALAR ==\n")
            f.writelines(f"  {x}\n" for x in o["buttons"])
            f.write("\n== OCHIQ RO'YXAT VARIANTLARI ==\n")
            f.writelines(f"  {x}\n" for x in o["variantlar"])
            f.write("\n== FAYL INPUTLAR ==\n")
            f.writelines(f"  tomon={x['tomon']} multiple={x['multiple']} accept='{x['accept']}' | {x['ctx']}\n"
                         for x in o["files"])
            f.write("\n== SAHIFA MATNI ==\n" + o["matn"])
        print(f"    (sahifa saqlandi: loglar\\{os.path.basename(asos)}.png/.txt)")

    # --- avvalgi skriptlarda saytda ISHLAGANI isbotlangan usullar ---
    def xp_bos(self, text, nom=None, exact=True, vaqt=None):
        nom = nom or text
        lit = xpath_literal(text)
        xp = f"(//*[normalize-space(text())={lit}])[1]" if exact else f"(//*[contains(., {lit})])[last()]"
        oxiri = time.time() + (KUTISH if vaqt is None else vaqt)
        while time.time() < oxiri:
            try:
                el = Kutish(self.d, 3).until(EC.element_to_be_clickable((By.XPATH, xp)))
                self.bos_el(el)
                print(f"    OK -> {nom}")
                return el
            except (StaleElementReferenceException, TimeoutException):
                tin(0.4)
        print(f"    !!! TOPILMADI -> {nom}")
        raise Topilmadi(nom)

    def ochish(self, ph, nom=None):
        """Dropdown/inputni placeholder bo'yicha ochadi."""
        nom = nom or ph
        el = self.inp(ph, vaqt=8)
        if not el:
            lit = xpath_literal(ph)
            el = self.wait.until(EC.element_to_be_clickable((By.XPATH, f"(//*[contains(., {lit})])[last()]")))
        self.bos_el(el)
        print(f"    OK -> {nom} (ochildi)")
        return el

    def tanla_yorliq(self, yorliq, matnlar, nom):
        """Yozuv (label) yonidagi ro'yxatni ochib, `matnlar`dan birinchi topilganini tanlaydi.
        Qat'iy pauzalar yo'q: ro'yxat ochilgani, variantlar chiqqani va panel yopilgani tekshiriladi.
        Ro'yxat ochilmasa (sayt band edi) - 3 martagacha qayta urinadi."""
        for urinish in range(3):
            el = self.kut(lambda: self.js("return maydon_yorliq(arguments[0]);", yorliq), 15 if urinish == 0 else 5, 0.2)
            if not el:
                print(f"    !!! '{yorliq}' ro'yxati topilmadi")
                return False
            if el.tag_name.lower() == "select":
                s = Select(el)
                for m in matnlar:
                    for o in s.options:
                        if norm_py(m) in norm_py(o.text):
                            s.select_by_visible_text(o.text)
                            print(f"    OK -> {nom}: {o.text.strip()}")
                            return True
                print(f"    !!! {nom}: mos variant yo'q. Variantlar: {[o.text.strip() for o in s.options]}")
                return False
            if not self.royxat_och(el):
                print(f"    ({nom}: ro'yxat ochilmadi - qayta urinish)")
                self.esc()
                continue
            if not self.kut(lambda: self.js("return ochiq_variantlar();"), 5, 0.1):   # variantlar serverdan kelyapti
                print(f"    ({nom}: ro'yxat bo'sh - qayta urinish)")
                self.esc()
                continue
            # variantlar birga chiziladi - ustuvorlik tartibida birinchi mos kelgani
            v = self.kut(lambda: next(filter(None, (self.js("return variant(arguments[0]);", m) for m in matnlar)), None),
                         1.0, 0.15)
            if not v:
                print(f"    !!! {nom}: mos variant yo'q. Variantlar: {self.js('return variantlar();')}")
                self.esc()
                return False
            matn = v.text.strip()
            self.bos_el(v, 0.5)
            if not self.kut(lambda: not self.js("return royxat_ochiqmi(null);"), 1.0, 0.1):
                self.esc()       # ko'p tanlovli ro'yxat o'zi yopilmaydi
            print(f"    OK -> {nom}: {matn}")
            return True
        print(f"    !!! {nom}: ro'yxat 3 marta ochilmadi. Xabarlar: {self.xatolar()}")
        return False


# ============================================================
# QADAMLAR
# ============================================================
def kirish(b, parol):
    print("\n>>> Kirish...")
    b.d.get(S["sayt"] + "/sign-in")
    # Doimiy profilda avvalgi sessiya saqlangan bo'lishi mumkin - unda kabinet o'zi ochiladi
    sahifa = b.kut(lambda: "kabinet" if b.top("Murojaatlar (ADOLAT)")
                   else ("kirish" if b.top("Kirish", exact=True) else None), KUTISH)
    if sahifa == "kabinet":
        print("    OK -> Kabinet ochiq (avvalgi sessiya saqlangan)")
        return
    b.xp_bos("Kirish", "Kirish (sign-in sahifasi)")
    login_xp = "//input[contains(@placeholder,'Loginni')]"

    def oneid_holati():
        """ONE ID: login formasi (element) | 'rol' - saqlangan sessiyada darhol rol tanlash | 'kabinet'."""
        if b.top("Murojaatlar (ADOLAT)"):
            return "kabinet"
        if any(s in b.matn() for s in KUTIB_TURING):
            raise SaytBand(SAYT_BAND_XABAR)
        if b.top("rolda davom") or (S["profil_stir"] and b.top(S["profil_stir"])):
            return "rol"
        for el in b.d.find_elements(By.XPATH, login_xp):
            if el.is_displayed() and el.is_enabled():
                return el
        return None

    login_el = None
    for urinish in range(2):
        login_el = b.kut(oneid_holati, max(KUTISH, 40))
        if login_el or urinish:
            break
        print("    (ONE ID sahifasi ochilmadi - bir marta qayta yuklanmoqda)")
        b.tahlil("oneid_bosh_sahifa")
        b.d.refresh()
    if login_el in ("kabinet", "rol"):
        print("    OK -> ONE ID sessiyasi saqlangan (login/parol kiritish shart emas)")
        kirish_yakuni(b)
        return
    if login_el is None:
        b.tahlil("oneid_ochilmadi")
        odam("ONE ID kirish sahifasi ochilmadi (sahifa bo'sh qoldi).\n"
             "Brauzerda O'ZINGIZ kiring (login, parol, kerak bo'lsa tasdiqlash).\n"
             "Kabinet ochilganda dastur o'zi davom etadi.",
             shart=lambda: b.top("Murojaatlar (ADOLAT)"), cheksiz=True)
        if not b.kut(lambda: b.top("Murojaatlar (ADOLAT)"), 30):
            raise Topilmadi("ONE ID orqali kirish bajarilmadi")
        print("    OK -> Kabinet ochildi (kirish qo'lda bajarildi)")
        return
    login_el.clear()
    login_el.send_keys(S["login"])
    print(f"    OK -> Login: {S['login']}")
    for _ in range(4):
        try:
            p = Kutish(b.d, 8).until(
                EC.element_to_be_clickable((By.XPATH, "//input[contains(@placeholder,'Parolni')]")))
            p.send_keys(parol)
            break
        except (TimeoutException, StaleElementReferenceException):
            tin(1.5)
    else:
        raise Topilmadi("Parol maydoni")
    print("    OK -> Parol kiritildi")
    b.tugma("Kirish", "Kirish (ONE ID forma)")
    kirish_yakuni(b)


KUTIB_TURING = ("янгиланмади", "кутиб туринг", "yangilanmadi", "kutib turing")   # sayt: "Тизим янгиланмади. Кутиб туринг."
KUTIB_TURING_SONIYA = 30


def kirish_yakuni(b):
    """ONE ID dan keyin: rol (tashkilot STIRi) tanlash, SMS - kabinet ochilguncha.
    Rol faqat ONE ID sahifasida tanlanadi; 'Davom etish'dan keyin sahifa almashguncha qayta bosilmaydi."""
    from urllib.parse import urlparse
    kabinet_host = urlparse(S["sayt"]).netloc

    def oneid_da():
        return urlparse(b.d.current_url or "").netloc not in ("", kabinet_host)

    def kabinet():
        return b.top("Murojaatlar (ADOLAT)")

    def kutib_turing():
        m = b.matn()
        return any(s in m for s in KUTIB_TURING)

    def rol_sahifasi():
        return oneid_da() and ((S["profil_stir"] and b.top(S["profil_stir"])) or b.top("Yuridik shaxs"))

    def rol_tanla():
        # sozlamadagi STIR bo'lsa - o'sha; bo'lmasa (boshqa akkaunt) - "Yuridik shaxs" kartasi
        if S["profil_stir"] and b.top(S["profil_stir"]):
            b.bos(S["profil_stir"], f"Profil: STIR {S['profil_stir']}")
        else:
            b.bos("Yuridik shaxs", "Profil: Yuridik shaxs")
        # "Davom etish" rol tanlangandan keyingina yoqiladi
        davom = b.kut(lambda: b.js("""
            for (const x of document.querySelectorAll('button'))
                if (vis(x) && !x.disabled && txt(x).includes('davom etish')) return x;
            return null;"""), 8)
        if davom:
            b.bos_el(davom)
            print("    OK -> Davom etish (kabinet ochilishi kutilmoqda)")
        # Sayt sekin bo'lishi mumkin - qayta bosmasdan kutamiz
        b.kut(lambda: kabinet() or not oneid_da() or kutib_turing(), 60)

    rol_urinish = kutish_soni = 0
    rol_kutish_boshi = None
    oxiri = time.time() + 600
    while time.time() < oxiri:
        if kabinet():
            print("    OK -> Kabinet ochildi")
            return
        if kutib_turing():
            raise SaytBand(SAYT_BAND_XABAR)
        if rol_sahifasi():
            if rol_urinish >= 3:
                b.tahlil("rol_tanlanmadi")
                odam("ONE ID: rolni tanlab bo'lmadi. Brauzerda o'zingiz tanlang va 'Davom etish'ni bosing.\n"
                     "Kabinet ochilganda dastur o'zi davom etadi.", shart=kabinet, cheksiz=True)
                continue
            rol_urinish += 1
            rol_tanla()
            continue
        if oneid_da() and b.top("rolda davom"):
            # rollar hali yuklanmoqda (STIR ko'rinmayapti) - biroz kutamiz, keyin odamdan so'raymiz
            rol_kutish_boshi = rol_kutish_boshi or time.time()
            if time.time() - rol_kutish_boshi > 30:
                b.tahlil("rol_sahifasi")
                odam("ONE ID: qaysi rolda davom etishni brauzerda tanlang va 'Davom etish'ni bosing.\n"
                     "Kabinet ochilganda dastur o'zi davom etadi.", shart=kabinet, cheksiz=True)
                rol_kutish_boshi = None
            tin(1)
            continue
        m = b.matn()
        if "parol" in m and ("notogri" in m or "xato" in m or "неверн" in m):
            raise ParolXato("Login yoki parol noto'g'ri")
        if "sms" in m or "tasdiqlash kodi" in m or b.js("return inp('kod');"):
            odam("SMS / tasdiqlash kodini brauzerda kiriting.\nKabinet ochilganda skript o'zi davom etadi (yoki Enter).",
                 shart=lambda: b.top("Murojaatlar (ADOLAT)"), cheksiz=True)
            continue
        tin(1)
    b.tahlil("kabinet_ochilmadi")
    raise Topilmadi("Kabinet (Murojaatlar (ADOLAT)) 10 daqiqada ochilmadi")


def sessiya_tekshir(b, parol):
    """Sessiya tugagan bo'lsa (sign-in sahifasiga qaytgan) qayta kiradi."""
    url = b.d.current_url or ""
    if "sign-in" in url or "login" in url or "oneid" in url:
        print("    (sessiya tugagan - qayta kirilmoqda)")
        kirish(b, parol)       # o'xshamasa - mijoz XATO -> brauzer qayta ochilib, shu mijoz yana bir marta


def murojaat_ochish(b, parol):
    print("\n>>> Murojaat yaratish sahifasi...")
    for _ in range(2):
        b.alert_yop()
        b.d.get(S["sayt"] + "/cases/create")
        b.alert_yop()
        if b.kut(lambda: b.top("Fuqarolik sudi", exact=True), KUTISH):
            print("    OK -> Murojaat yaratish")
            return
        sessiya_tekshir(b, parol)
    # zaxira: menyu orqali
    if not b.top("Murojaat yaratish", exact=True):
        b.bos("Murojaatlar (ADOLAT)", "Murojaatlar (ADOLAT) menyusi")
        tin(1)
    b.bos("Murojaat yaratish", "Murojaat yaratish (menyu)", exact=True)
    if not b.kut(lambda: b.top("Fuqarolik sudi", exact=True), KUTISH):
        raise Topilmadi("Murojaat yaratish sahifasi")


def boshlash(b):
    print(">>> Sud turi, murojaat turi, instantsiya...")
    b.xp_bos("Fuqarolik sudi", "Fuqarolik sudi")
    b.xp_bos("arizasi", "Murojaat turi dropdown", exact=False)
    b.xp_bos("vo tartibida", "Da'vo tartibida", exact=False)
    b.xp_bos("Birinchi instantsiya", "Birinchi instantsiya")
    b.xp_bos("Oldinga", "Oldinga")


ROBOT_BOS_JS = """
    const labels = Array.from(document.querySelectorAll('span'));
    const robotLabel = labels.find(el => {
        const t = el.textContent.trim().toLowerCase();
        return t.includes('robot emasman') || t.includes('робот эмасман');
    });
    if (robotLabel) {
        const container = robotLabel.parentElement.querySelector('.mat-checkbox-inner-container') || robotLabel.parentElement;
        if (container) container.click();
    }
"""


def captcha_holati(b):
    """CAPTCHA holati - FAQAT tekshiradi, bosmaydi: 'yoq' | 'misol' | 'yuklanmoqda' | 'kutilmoqda' | 'belgilandi'.
    (Avval bu funksiya har chaqirilganda katakni bosardi: natija kutilayotganda belgi har 0.4 soniyada
    qo'yilib-olinib, qidiruv o'tmay qolardi - 40 soniyadan 2-3 marta.)"""
    return b.js("return typeof robot_holati === 'function' ? robot_holati() : 'yoq';")




def robot_belgila(b, nom):
    """"Robot emasman" katagini bosadi. Misolli CAPTCHA chiqsa - CaptchaMisol: mijoz sahifa yangidan
    ochilib qaytadan boshlanadi (ishga_tushirish, 3 martagacha + navbat oxirida yana bir marta)."""
    b.kut(lambda: captcha_holati(b) != "yuklanmoqda", 10, 0.3)
    h = captcha_holati(b)
    if h == "misol":
        raise CaptchaMisol(f"CAPTCHA misoli chiqdi ({nom})")
    if h != "kutilmoqda":
        return h
    imzo = b.js("return robot_imzo();") or ""
    toza = b.robot_toza.setdefault(nom, imzo)
    if imzo != toza:
        return "belgilandi"
    b.js(ROBOT_BOS_JS)
    # bosilgandan keyin belgilanishini yoki misol chiqishini kutamiz
    b.kut(lambda: captcha_holati(b) in ("belgilandi", "misol") or (b.js("return robot_imzo();") or "") != toza, 4, 0.2)
    h = captcha_holati(b)
    if h == "misol":
        raise CaptchaMisol(f"CAPTCHA misoli chiqdi ({nom})")
    uxla(1)
    print(f"    OK -> CAPTCHA: 'Robot emasman' belgilandi ({nom})")
    return h


def captcha_kutish(b, nom, natija_shart):
    """CAPTCHA: "Robot emasman" belgilanadi; misol chiqsa - CaptchaMisol (mijoz qaytadan boshlanadi)."""
    robot_belgila(b, nom)

def qidirish_bos(b, nom):
    """Yoqilgan 'Qidirish' tugmasini bosadi."""
    el = b.kut(lambda: b.js("""
        let t = null;
        for (const x of document.querySelectorAll('button'))
            if (vis(x) && !x.disabled && !x.closest(OVERLAY) && txt(x).includes('qidirish')) t = x;
        return t;"""), 15)
    if not el:
        print(f"    !!! 'Qidirish' tugmasi yoqilmadi ({nom}). Xabarlar: {b.xatolar()}")
        return False
    b.bos_el(el)
    print(f"    OK -> Qidirish ({nom})")
    return True

def maydonni_tikla(b, ph, qiymat, nom):
    """Qidiruv maydoni (STIR / JShShIR) bo'shab qolsa qayta yozadi - sayt muvaffaqiyatsiz qidiruvdan keyin tozalaydi."""
    def f():
        el = b.inp(ph, vaqt=3)
        if el and re.sub(r"\D", "", el.get_attribute("value") or "") != re.sub(r"\D", "", qiymat):
            print(f"    ({nom} maydoni bo'shab qolgan - qayta yozilmoqda)")
            b.yoz(el, qiymat)
    return f


# saytning "topilmadi / vafot etgan" xabarlari (norm_py dan o'tgan matnda: kichik harf, tutuq belgisiz)
TOPILMADI_RE = re.compile(
    r"topilmad|topilmag|mavjud emas|malumot(?:lar)? (?:yoq|topilmad)|vafot|olganligi|tirik emas|"
    r"топилмад|мавжуд эмас|маълумот(?:лар)? (?:йук|топилмад)|вафот|улганлиг|"
    r"не найден|не существ|нет данных|нет сведений|не обнаружен|умер|сконча|смерт|not found|no data|deceased")


VAFOT_RE = re.compile(r"vafot|вафот|tirik emas|olganlig|улганлиг|умер|сконча|смерт|deceased")
XATO_RE = re.compile(r"xato|хато|ошибк|error|muvaffaqiyatsiz|не удалось|imkoni yoq|мумкин эмас|невозможно")
CAPTCHA_RE = re.compile(r"robot|captcha|kapcha|капча|робот|misol|rasmdagi")


def sayt_xatosi(b, oldin_xatolar=(), oldin_matn=()):
    """Qidiruvdan keyin saytda YANGI chiqqan xabar: 'topilmadi / vafot etgan' yoki boshqa xato (pastdagi qizil
    yozuv, toast). CAPTCHA va bo'sh maydon belgilari hisobga olinmaydi. Topilmasa ''."""
    for x in b.xatolar():
        n = norm_py(x)
        if x in oldin_xatolar or n.startswith("notogri/bosh maydon") or CAPTCHA_RE.search(n):
            continue
        return x
    return _yangi_qator(b, oldin_matn, (TOPILMADI_RE, XATO_RE))


def sahifa_qatorlari(b):
    try:
        return [q.strip() for q in (b.js("return document.body.innerText;") or "").splitlines() if q.strip()]
    except WebDriverException:
        return []


def _yangi_qator(b, oldin, regexlar):
    """Qidiruvdan keyin sahifada YANGI paydo bo'lgan qator (pastdagi "Ушбу шахс вафот этган" kabi)."""
    oldin = set(oldin or ())
    for rx in regexlar:
        for q in sahifa_qatorlari(b):
            n = norm_py(q)
            if q not in oldin and len(q) < 300 and rx.search(n) and not CAPTCHA_RE.search(n):
                return q
    return ""


def topilmadi_xabari(b, oldin=()):
    """Sayt 'topilmadi / vafot etgan' dedimi - xabar matni, aks holda ''. oldin - avvalgi sahifa qatorlari."""
    for x in b.xatolar():
        if TOPILMADI_RE.search(norm_py(x)):
            return x
    return _yangi_qator(b, oldin, (TOPILMADI_RE,))


def qidiruv_captcha(b, natija_shart, nom, tayyorla=None):
    """CAPTCHA bo'lsa odam yechadi, yo'q bo'lsa to'xtamaydi -> Qidirish -> natija. 3 marta urinadi.
    CAPTCHA Qidirishdan KEYIN chiqsa ham kutiladi va qayta qidiriladi."""
    for urinish in range(1, 4):
        captcha_kutish(b, nom, natija_shart)
        if natija_shart():
            return True
        if tayyorla:
            tayyorla()
        oldin, oldin_xat = sahifa_qatorlari(b), b.xatolar()
        if qidirish_bos(b, nom):
            t0 = time.time()

            def holat():                          # katak bu yerda BOSILMAYDI - faqat kuzatiladi
                if natija_shart():
                    return "natija"
                if time.time() - t0 > 1.5 and sayt_xatosi(b, oldin_xat, oldin):
                    return "yoq"
                if time.time() - t0 > 3:
                    h = captcha_holati(b)
                    if h == "misol":
                        return "captcha"          # sayt misol so'radi
                    # belgi olib tashlangan va 12 soniyada javob yo'q - qayta belgilab qidiramiz (40 soniya kutmasdan)
                    if (time.time() - t0 > 12 and h == "kutilmoqda"
                            and (b.js("return robot_imzo();") or "") == b.robot_toza.get(nom)):
                        return "captcha"
                return None
            r = b.kut(holat, min(KUTISH, 25), 0.3)
            if r == "natija":
                return True
            if r == "yoq":
                tin(1.5)
                if natija_shart():                # natija kechikib chiqdi
                    return True
                xabar = sayt_xatosi(b, oldin_xat, oldin) or "sayt shaxsni topmadi"
                print(f"    !!! {nom}: sayt xabari - {xabar}")
                b.tahlil(f"qidiruv_xato_{nom}")
                raise ShaxsTopilmadi(xabar)
            if r == "captcha":
                print("    (Qidirishdan keyin CAPTCHA qayta so'raldi - belgilab qayta qidiriladi)")
                continue
        print(f"    !!! Qidiruv natijasi chiqmadi (urinish {urinish}). Sahifadagi xabarlar: {b.xatolar()}")
        b.tahlil(f"qidiruv_natijasiz_{nom}")
    raise RuntimeError(f"Qidiruv muvaffaqiyatsiz tugadi: {nom}")



def keyingi(b, nom="Keyingi"):
    """'Keyingi'ni bosadi va sahifa haqiqatan almashganini tekshiradi."""
    oldin = b.imzo()
    b.tugma("Keyingi", nom)
    if b.kut(lambda: b.imzo() != oldin, KUTISH, 0.25):
        tin(1.5)
        return True
    print(f"    !!! Sahifa almashmadi. Xabarlar: {b.xatolar()}")
    b.tahlil("keyingi_otmadi")
    return False


def keyingi_majburiy(b, nom="Keyingi"):
    oldin_matn = sahifa_qatorlari(b)
    if keyingi(b, nom):
        return
    xabar = topilmadi_xabari(b, oldin_matn)
    if xabar:
        raise ShaxsTopilmadi(xabar)
    sayt = "; ".join(x for x in b.xatolar() if not norm_py(x).startswith("notogri/bosh maydon"))[:200]
    oldin = b.imzo()
    odam(f"'Keyingi' o'tmadi. Sayt: {sayt or '-'}\n"
         "Brauzerda kamchilikni to'g'irlab 'Keyingi'ni O'ZINGIZ bosing - skript sezadi (yoki Enter).",
         shart=lambda: b.imzo() != oldin)
    tin(1.5)


# ---------------- 1-qadam ----------------
def royxatdan_tanla(b, ph, nom, balli):
    """Ro'yxatni (placeholder ph) ochib, variantlardan balli(matn) eng kattasini tanlaydi."""
    b.ochish(ph, nom)

    def eng_yaxshi():
        v = [(balli(x), x) for x in (b.js("return variantlar();") or [])]
        v = [x for x in v if x[0] > 0]
        return max(v, key=lambda x: x[0])[1] if v else None

    matn = b.kut(eng_yaxshi, 10)
    el = b.js("return variant(arguments[0]);", matn) if matn else None
    if not el:
        print(f"    !!! {nom}: mos variant yo'q. Variantlar: {(b.js('return variantlar();') or [])[:40]}")
        b.esc()
        return False
    b.bos_el(el)
    print(f"    OK -> {nom}: {matn}")
    return True


def qadam1_sud_davogar(b, ish):
    print("\n>>> 1-qadam: sud va da'vogar")
    if ish.sud:
        # javobgar manzili bo'yicha (Fuqarolik_sudlari_yurisdiksiya.xlsx); Toshkent shahri yoki sozlamalardagi
        # "hamma mijoz - Toshkent shahar" belgisi - toshkent_sudi()
        print(f"    (sud: {ish.sud['hudud']} / {ish.sud['nomi']}"
              + (" - sozlamalar: hamma mijoz Toshkent shahar)" if hammaga_toshkent() else f"; javobgar manzili: {ish.manzil})"))
        if not royxatdan_tanla(b, "Viloyat", "Viloyat", lambda v: viloyat_variant_balli(v, ish.sud["hudud"])):
            b.tahlil("viloyat_topilmadi")
            odam(f"Viloyatni QO'LDA tanlang: {ish.sud['hudud']}\nTanlagach Enter bosing.")
        tin(1)
        if not royxatdan_tanla(b, "Sud nomi", "Sud nomi", lambda v: sud_variant_balli(v, ish.sud)):
            b.tahlil("sud_topilmadi")
            odam(f"Sudni QO'LDA tanlang: {ish.sud['nomi']}\nTanlagach Enter bosing.")
    else:
        # sudlar jadvali yo'q - sozlamalar.ini dagi viloyat/sud
        b.ochish("Viloyat", "Viloyat")
        b.xp_bos(S["viloyat"], f"Viloyat: {S['viloyat']}", exact=False)
        b.ochish("Sud nomi", "Sud nomi")
        b.xp_bos(S["sud"], f"Sud: {S['sud']}", exact=False)

    el = b.ochish("vogarni tanlang", "Da'vogar nomi")
    try:
        el.send_keys(S["asosiy_davogar"][:11])
    except Exception:
        pass
    if not b.bos(S["asosiy_davogar"], f"Da'vogar: {S['asosiy_davogar']}", vaqt=8, majburiy=False):
        # sozlamadagi nom ro'yxatda yo'q (boshqa akkaunt) - jismoniy shaxs emas, tashkilot tanlanadi
        v = b.kut(lambda: b.js("return variant('jamiyat') || variant('mchj') || variant('masuliyati');"), 5)
        if not v:
            b.tahlil("davogar_topilmadi")
            raise Topilmadi(f"Da'vogar ro'yxatida '{S['asosiy_davogar']}' ham, tashkilot ham yo'q")
        nomi = (v.text or "").strip()
        b.bos_el(v)
        print(f"    OK -> Da'vogar (tashkilot): {nomi}")
    b.xp_bos(S["kichik_biznes"], f"Kichik biznes subyekti: {S['kichik_biznes']}")
    tin(1.5)

    if not S["qoshimcha_davogar_stir"]:
        keyingi_majburiy(b)
        return
    print(">>> Qo'shimcha da'vogar (STIR orqali)...")
    try:
        plus = Kutish(b.d, 6).until(EC.element_to_be_clickable(
            (By.XPATH, f"(//*[contains(., {xpath_literal('shimcha')})]/following::button[1])")))
        b.bos_el(plus)
        print("    OK -> Qo'shimcha da'vogar '+'")
    except TimeoutException:
        b.xp_bos("shimcha", "Qo'shimcha da'vogar +", exact=False)
    tin(1)
    b.xp_bos("Yuridik shaxs", "Yuridik shaxs", exact=False)
    stir = b.inp("12345678910")
    if not stir:
        raise Topilmadi("STIR maydoni")
    print(f"    OK -> STIR: {b.yoz(stir, S['qoshimcha_davogar_stir'])}")

    belgi = norm_py(S["qoshimcha_davogar_belgi"])

    def natija():
        return b.js("""
            for (const e of document.querySelectorAll('input,textarea')) {
                if (vis(e) && norm(e.value).includes(arguments[0])) return true;
            }
            return false;""", belgi)

    qidiruv_captcha(b, natija, "da'vogar STIR",
                    maydonni_tikla(b, "12345678910", S["qoshimcha_davogar_stir"], "STIR"))
    print("    OK -> Tashkilot topildi")
    b.tugma("shish", "Qo'shish (qo'shimcha da'vogar)")
    tin(1.5)
    b.tahlil("1_davogar")
    keyingi_majburiy(b)


# ---------------- 2-qadam ----------------
def sana_mos(qiymat, s):
    for fmt in ("%m/%d/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            if datetime.strptime(qiymat.strip(), fmt).date() == s:
                return True
        except ValueError:
            pass
    return False


def kalendar_orqali(b, el, s):
    """Angular Material kalendari: sarlavha -> yil -> oy (tartib raqami bo'yicha) -> kun."""
    try:
        b.bos_el(el)
        tin(0.6)
        ochiq = "return !!document.querySelector('mat-calendar,.mat-calendar,.mat-datepicker-content');"
        if not b.js(ochiq):
            t = b.js("const c = arguments[0].closest('mat-form-field,.mat-form-field,.mat-mdc-form-field');"
                     "return c ? c.querySelector('mat-datepicker-toggle button, button') : null;", el)
            if t:
                b.bos_el(t)
                tin(0.6)
        if not b.js(ochiq):
            return False
        b.bos_el(b.js("return document.querySelector('.mat-calendar-period-button');"))
        tin(0.5)
        yil = b.js("return [...document.querySelectorAll('.mat-calendar-body-cell')]"
                   ".find(c => txt(c) === arguments[0]) || null;", str(s.year))
        if not yil:
            return False
        b.bos_el(yil)
        tin(0.5)
        oylar = b.js("return [...document.querySelectorAll('.mat-calendar-body-cell')];")
        if len(oylar) != 12:
            return False
        b.bos_el(oylar[s.month - 1])
        tin(0.6)
        kun = b.js("return [...document.querySelectorAll('.mat-calendar-body-cell')]"
                   ".find(c => txt(c) === arguments[0] && !c.className.includes('disabled')) || null;", str(s.day))
        if not kun:
            return False
        b.bos_el(kun)
        tin(0.4)
        return True
    except Exception as e:
        print(f"    (kalendar: {type(e).__name__}: {e})")
        return False


def sana_kiritish(b, el, s):
    maqsad = f"{s.month}/{s.day}/{s.year}"   # sayt formati: 4/14/2026
    korinish = s.strftime("%d.%m.%Y")

    def tekshir():
        return sana_mos(el.get_attribute("value") or "", s)

    # Saytda sana maydoni readonly - faqat kalendar orqali tanlanadi
    if not el.get_attribute("readonly"):
        try:
            b.yoz(el, maqsad)
            el.send_keys(Keys.TAB)
            tin(0.4)
        except Exception:
            pass
        b.esc()
        if tekshir():
            print(f"    OK -> Ariza sanasi: {korinish}")
            return
    for _ in range(2):
        if kalendar_orqali(b, el, s) and tekshir():
            print(f"    OK -> Ariza sanasi: {korinish} (kalendar orqali)")
            return
        b.esc()
        tin(0.8)
    b.tahlil("sana_otmadi")
    odam(f"Ariza sanasini QO'LDA tanlang: {korinish}\nTanlangach skript o'zi sezadi (yoki Enter).", shart=tekshir)


def summa_mos(qiymat, kerak):
    """Maydondagi summa kerakli summaga TENGmi (faqat raqamlar emas - kasr qismi ham)."""
    t = str(qiymat or "").replace("\xa0", "").replace(" ", "").strip()
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".") else t.replace(",", "")
    else:
        t = t.replace(",", ".")
    try:
        return abs(float(t) - float(kerak)) < 0.005
    except ValueError:
        return False


def butun_som(summa):
    """Sayt tiyinni (nuqta/vergul) qabul qilmaydi - faqat butun so'm: 1234567.89 -> "1234567" (yaxlitlanmaydi)."""
    return str(int(float(summa)))


def summa_yoz(b, el, summa, nom):
    """Summani faqat BUTUN so'mda klaviaturada yozadi va maydonda shu son turganini tekshiradi."""
    kerak = butun_som(summa)
    for urinish in range(3):
        try:
            b.bos_el(el)
            b.tozala(el)
            for ch in kerak:
                el.send_keys(ch)
                harf_pauza()
            el.send_keys(Keys.TAB)
            tin(0.5)
        except (StaleElementReferenceException, ElementNotInteractableException):
            tin(0.8)
            continue
        v = (el.get_attribute("value") or "").strip()
        if summa_mos(v, kerak):
            izoh = f" (tiyin tashlandi: {summa})" if float(summa) != float(kerak) else ""
            print(f"    OK -> {nom}: {v}{izoh}")
            return True
        print(f"    (urinish {urinish + 1}: maydonda '{v}', kerak '{kerak}')")
    return False


def klaviatura_yoz(b, el, qiymat, nom):
    """Odam kabi klaviaturada yozadi (Angular modeli qiymatni olishi uchun), so'ng tekshiradi."""
    kerak = re.sub(r"\D", "", str(qiymat))
    for urinish in range(3):
        try:
            b.bos_el(el)
            b.tozala(el)
            for ch in str(qiymat):
                el.send_keys(ch)
                harf_pauza()
            el.send_keys(Keys.TAB)
            tin(0.4)
        except (StaleElementReferenceException, ElementNotInteractableException):
            tin(0.8)
            continue
        v = (el.get_attribute("value") or "").strip()
        if re.sub(r"\D", "", v) == kerak:
            print(f"    OK -> {nom}: {v}")
            return True
        print(f"    (urinish {urinish + 1}: {nom} maydonida '{v}', kerak '{qiymat}')")
    return False


def mat_tanla(b, yorliq, qiymat, nom):
    """Angular Material ro'yxatidan (mat-select) yozuvi bo'yicha tanlaydi va tekshiradi."""
    if not qiymat:
        return True
    korindi = lambda: (lambda e: e is None or norm_py(qiymat) in norm_py(e.text))(
        b.js("return maydon_yorliq(arguments[0]);", yorliq))
    for _ in range(2):
        if b.tanla_yorliq(yorliq, [qiymat], nom):
            if b.kut(korindi, 1.5, 0.15):
                return True
            el = b.js("return maydon_yorliq(arguments[0]);", yorliq)
            print(f"    (tanlov ko'rinmadi: '{(el.text if el else '').strip()}')")
        b.esc()
        tin(1)
    b.tahlil(f"tanlanmadi_{yorliq}")
    odam(f"'{yorliq}' ro'yxatidan QO'LDA tanlang: {qiymat}\nTanlagach Enter bosing.")
    return False


def qadam2_asosiy(b, ish):
    print("\n>>> 2-qadam: asosiy ma'lumotlar")
    raqam = b.inp("Raqam")
    if not raqam:
        b.tahlil("raqam_maydoni_yoq")
        raise Topilmadi("Ariza raqami maydoni (2-qadam ochilmadi)")
    klaviatura_yoz(b, raqam, ish.raqam, "Ariza raqami")

    sana_el = b.wait.until(EC.element_to_be_clickable(
        (By.XPATH, f"(//*[contains(., {xpath_literal('ariza sanasi')})]/following::input[1])")))
    sana_kiritish(b, sana_el, ish.sana)

    # Bular mat-select (input emas) - yozuvi (label) bo'yicha tanlanadi
    mat_tanla(b, "Nizo turi", S["nizo_turi"], "Nizo turi")
    mat_tanla(b, "Ariza turi", S["ariza_turi"], "Ariza turi")
    mat_tanla(b, "Asosiy ish turkumi", S["ish_turkumi"], "Asosiy ish turkumi")
    tin(1)
    mat_tanla(b, "Qo'shimcha ish turkumi", S["qoshimcha_ish_turkumi"], "Qo'shimcha ish turkumi")
    # Ish turkumi jadvalga qo'shilishi kerak ("Qo'shish" - ish turkumi yonidagi tugma)
    turkum = norm_py(S["ish_turkumi"])
    if b.matn().count(turkum) < 2:
        xp = (f"(//*[contains(., {xpath_literal('ish turkumi')})]/following::button[contains(translate(., "
              "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'QO')])[1]")
        try:
            b.bos_el(Kutish(b.d, 6).until(EC.element_to_be_clickable((By.XPATH, xp))))
            print("    OK -> Ish turkumi -> Qo'shish (jadvalga)")
            tin(1)
        except TimeoutException:
            print("    !!! Ish turkumi 'Qo'shish' tugmasi topilmadi")

    xp = (f"(//*[contains(., {xpath_literal('vo summasi')})]/following::button[contains(translate(., "
          "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'QO')])[1]")
    if not b.inp("Asosiy qarz", vaqt=2):
        try:
            b.bos_el(Kutish(b.d, 8).until(EC.element_to_be_clickable((By.XPATH, xp))))
            print("    OK -> Da'vo summasi -> Qo'shish")
        except TimeoutException:
            b.tugma("shish", "Qo'shish (da'vo summasi)")
    qarz = b.inp("Asosiy qarz")
    if not qarz:
        raise Topilmadi("Asosiy qarz maydoni")
    summa_yoz(b, qarz, ish.asosiy_qarz, "Asosiy qarz")
    tin(1)

    # --- YAKUNIY TEKSHIRUV: hamma maydon to'lganmi? ---
    holat = b.js("""
        const r = {};
        const i = inp('Raqam'); r.raqam = i ? i.value : '';
        const q = inp('Asosiy qarz'); r.qarz = q ? q.value : '';
        const s = inp('vo summasi'); r.summa = s ? s.value : '';
        for (const k of ['Nizo turi', 'Ariza turi']) { const m = maydon_yorliq(k); r[k] = m ? (m.innerText || '').trim() : ''; }
        r.turkum = norm(document.body.innerText).split(arguments[0]).length - 1;
        return r;""", turkum) or {}
    print(f"    Tekshiruv: raqam='{holat.get('raqam')}' nizo='{holat.get('Nizo turi')}' ariza='{holat.get('Ariza turi')}' "
          f"qarz='{holat.get('qarz')}' da'vo summasi='{holat.get('summa')}' turkum_jadvalda={holat.get('turkum', 0) >= 1}")
    kamchilik = []
    if re.sub(r"\D", "", holat.get("raqam", "")) != ish.raqam:
        kamchilik.append(f"Ariza raqami ({ish.raqam})")
    if S["nizo_turi"] and norm_py(S["nizo_turi"]) not in norm_py(holat.get("Nizo turi", "")):
        kamchilik.append(f"Nizo turi ({S['nizo_turi']})")
    if S["ariza_turi"] and norm_py(S["ariza_turi"]) not in norm_py(holat.get("Ariza turi", "")):
        kamchilik.append(f"Ariza turi ({S['ariza_turi']})")
    if not summa_mos(holat.get("qarz", ""), butun_som(ish.asosiy_qarz)):
        kamchilik.append(f"Asosiy qarz ({butun_som(ish.asosiy_qarz)})")
    if not re.sub(r"\D", "", holat.get("summa", "")):
        kamchilik.append("Da'vo summasi bo'sh - sayt asosiy qarzni qabul qilmagan")
    b.tahlil("2_asosiy")
    if kamchilik:
        odam("2-qadamda TO'LMAGAN maydonlar:\n" + "\n".join(f"  - {k}" for k in kamchilik) +
             "\nBrauzerda QO'LDA to'ldiring, so'ng Enter bosing.")
    keyingi_majburiy(b)


# ---------------- 3-qadam ----------------
def fio_oddiy(s):
    jadval = str.maketrans({"А": "A", "Б": "B", "В": "V", "Г": "G", "Д": "D", "Е": "E", "Ё": "YO", "Ж": "J",
                            "З": "Z", "И": "I", "Й": "Y", "К": "K", "Л": "L", "М": "M", "Н": "N", "О": "O",
                            "П": "P", "Р": "R", "С": "S", "Т": "T", "У": "U", "Ф": "F", "Х": "X", "Ц": "TS",
                            "Ч": "CH", "Ш": "SH", "Щ": "SH", "Ъ": "", "Ы": "I", "Ь": "", "Э": "E", "Ю": "YU",
                            "Я": "YA", "Ў": "O", "Қ": "Q", "Ғ": "G", "Ҳ": "H"})
    return re.sub(r"[^A-Z]", "", s.upper().translate(jadval))


def qadam3_javobgar(b, ish):
    print("\n>>> 3-qadam: javobgar")
    if not b.kut(lambda: b.top("javobgar tomon"), KUTISH * 2):
        b.tahlil("javobgar_sahifasi_yoq")
        raise Topilmadi("Javobgar sahifasi")
    b.bos("Jismoniy shaxs", "Jismoniy shaxs", vaqt=3, majburiy=False)
    jsh = b.inp("12345678910")
    if not jsh:
        b.tahlil("jshshir_maydoni_yoq")
        raise Topilmadi("JShShIR maydoni")
    print(f"    OK -> JShShIR: {b.yoz(jsh, ish.jshshir)}")

    fio_js = """
        const r = {};
        for (const e of document.querySelectorAll('input')) {
            if (!vis(e)) continue;
            const p = norm(e.getAttribute('placeholder'));
            if (p.includes('famil')) r.f = e.value;
            else if (p.startsWith('ism')) r.i = e.value;
            else if (p.includes('sharif')) r.s = e.value;
        }
        return r;"""
    qidiruv_captcha(b, lambda: (b.js(fio_js) or {}).get("f", "").strip(), "javobgar JShShIR",
                    maydonni_tikla(b, "12345678910", ish.jshshir, "JShShIR"))
    fio = b.js(fio_js) or {}
    topilgan = " ".join(x for x in (fio.get("f"), fio.get("i"), fio.get("s")) if x)
    print(f"    OK -> Saytdan topildi: {topilgan}")

    # familiya solishtirilmaydi (AL-FAMILIYA / ALFAMILIYA kabi yozilish farqi to'xtatmasin) - JShShIR bo'yicha topilgan
    b.tugma("shish", "Qo'shish (javobgar)")
    tin(1.5)
    b.tahlil("3_javobgar")
    keyingi_majburiy(b)


# ---------------- 4-qadam ----------------
def kvitansiya_oqish(pdf):
    """billing.sud.uz kvitansiyasi (PDF): {raqam, sud, summa, muddat, pochta} yoki None (kvitansiya emas)."""
    t = pdf_matn(pdf)
    x = re.search(r"ХОС\s+РАҚАМ\s*(\d{8,})", t, re.I)
    if not x:
        return None
    sud = next((q.strip() for q in t.splitlines() if re.search(r"суди\s*$", q.strip(), re.I)), "")
    sm = re.search(r"Пул миқдори\s*\(сўм\)\s*([\d \xa0]+(?:,\d{1,2})?)", t, re.I)
    m = re.search(r"(\d{2})\.(\d{2})\.(20\d\d)\s*йилга\s*қадар", t)
    try:
        muddat = date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None
    except ValueError:
        muddat = None
    stirlar = re.findall(r"СТИР\s*(\d{9})", t[:x.start()])     # oxirgisi - to'lovchiniki (birinchisi - sudniki)
    return {"raqam": x.group(1), "sud": sud, "summa": float(summa_tozalash(sm.group(1)) or 0) if sm else 0.0,
            "muddat": muddat, "pochta": "почта" in t.lower(), "tolovchi_stir": stirlar[-1] if stirlar else ""}


def kvitansiya_sudi_mos(k, sud):
    """Kvitansiyadagi sud shu sudmi (turi bilan: "Чуст туманлараро" - "Наманган вилоят" emas)."""
    return bool(k and sud and sud_variant_balli(k["sud"], sud) >= 4)


def kvitansiya_raqami(ish):
    """4-qadam: mijoz papkasidagi to'lov kvitansiyasi. Bir nechta bo'lsa - SUDI shu ishniki bilan mos (yangisi);
    mos yo'q bo'lsa - birinchisi (eski tartib). Kvitansiyani alohida dastur KvitansiyaYaratish.exe yaratadi."""
    tolov = [y for k, _n, y in ish.hujjatlar if k == "tolov"]
    mos = []
    for y in tolov:
        if y.lower().endswith(".pdf"):
            k = kvitansiya_oqish(y)
            if kvitansiya_sudi_mos(k, ish.sud):
                mos.append((os.path.getmtime(y), k["raqam"]))
    if mos:
        return max(mos)[1]
    for y in tolov:
        x = re.search(r"\d{8,}", os.path.basename(y))
        if x:
            return x.group(0)
    return ""


def kvitansiya_xatosi(ish, raqam, xabarlar):
    """Sayt kvitansiyani qabul qilmadi: izoh (natijalar.csv / jadval) - sayt xabari va PDF dan bilinadigan sabab."""
    q = [f"KVITANSIYA NOTO'G'RI: {raqam} saytda qabul qilinmadi"]
    if xabarlar:
        q.append("sayt: " + "; ".join(xabarlar)[:200])
    k = next((k for k in (kvitansiya_oqish(y) for kk, _n, y in ish.hujjatlar
                          if kk == "tolov" and y.lower().endswith(".pdf")) if k and k["raqam"] == raqam), None)
    if k and ish.sud and not kvitansiya_sudi_mos(k, ish.sud):
        q.append(f"kvitansiya boshqa sud uchun: {k['sud']} (ish sudi: {ish.sud['nomi']})")
    elif k and k["muddat"] and k["muddat"] < date.today():
        q.append(f"kvitansiya muddati o'tgan ({k['muddat']:%d.%m.%Y})")
    elif not xabarlar:
        q.append("ehtimol avvalgi urinishda boshqa murojaatga (qoralamaga) biriktirilgan yoki to'lanmagan - billing.sud.uz da tekshiring")
    return " - ".join(q)


def qadam4_xarajatlar(b, ish):
    print("\n>>> 4-qadam: sud xarajatlari")
    if S["davlat_boji_imtiyozi"]:
        # mat_tanla tanlovni ro'yxatda ko'rinishi bo'yicha tekshiradi, bo'lmasa qayta urinadi / odamdan so'raydi
        mat_tanla(b, "To'lov imtiyozi", S["davlat_boji_imtiyozi"], "To'lov imtiyozi")

    raqam = kvitansiya_raqami(ish)
    if raqam:
        print(f">>> Mavjud kvitansiya biriktirilmoqda: {raqam}")
        sahifada = lambda: raqam in (b.js("return document.body.innerText;") or "")
        ok = False
        bosilgan = []      # Qidirish -> Biriktirish kabi ketma-ket tugmalar; bir xilini qayta bosmaymiz
        xabarlar = []      # tugma bosilgach saytda chiqqan xabarlar (toast tez yo'qoladi - yig'ib boramiz)
        oldin_xat = set(b.xatolar())

        def kuzat(soniya):
            """`soniya` davomida saytning yangi xabarlarini yig'adi; raqam ro'yxatga chiqsa - darhol True."""
            oxiri = time.time() + soniya
            while True:
                for x in b.xatolar():
                    if x not in oldin_xat and x not in xabarlar and not norm_py(x).startswith("notogri/bosh maydon"):
                        xabarlar.append(x)
                if sahifada():
                    return True
                if time.time() > oxiri:
                    return False
                uxla(0.25)

        b.esc()                  # ochiq qolgan ro'yxat (To'lov imtiyozi) yopilsin
        tin(0.4)
        plus = b.kut(lambda: b.js("return kvitansiya_plus();"), 10)
        if not plus:
            print("    !!! '+' (kvitansiya) topilmadi")
        else:
            # '+' dan keyin paydo bo'lgan yangi maydon - kvitansiya raqami (dialogda yoki shu joyda)
            b.js("window.__adolat_inp = new Set(document.querySelectorAll('input')); return true;")
            b.bos_el(plus)
            print("    OK -> '+' (kvitansiya biriktirish)")
            el = b.kut(lambda: b.js("""
                const yangi = i => vis(i) && !i.disabled && !i.readOnly;
                for (const d of document.querySelectorAll(DIALOG))
                    if (vis(d)) for (const i of d.querySelectorAll('input:not([type=checkbox]):not([type=radio]):not([type=file])'))
                        if (yangi(i)) return i;
                for (const i of document.querySelectorAll('input:not([type=checkbox]):not([type=radio]):not([type=file])'))
                    if (yangi(i) && !window.__adolat_inp.has(i)) return i;
                return null;"""), 10)
            if not el:
                print("    !!! Kvitansiya raqami maydoni chiqmadi")
            else:
                print(f"    OK -> Kvitansiya raqami: {b.yoz(el, raqam)}")
                if captcha_holati(b) != "yoq":
                    captcha_kutish(b, "kvitansiya", sahifada)
                # Maydon yonidagi tasdiqlash tugmasi (Qidirish / Qo'shish / Biriktirish / Saqlash ...)
                TASDIQ_JS = """
                        const i = arguments[0], bosilgan = arguments[1];
                        const MOS = /qidir|qoshish|biriktir|saqla|tasdiq|^ok$|izla|добав|сохран|qabul/;
                        const YOQ = /qoralama|orqaga|keyingi|bekor|yopish/;
                        let c = (i && i.isConnected && (i.closest(DIALOG) || i)) || document.querySelector(DIALOG);
                        for (let k = 0; k < 7 && c; k++) {
                            for (const x of c.querySelectorAll('button,[role=button],a,div,span')) {
                                const s = txt(x);
                                if (!vis(x) || x.disabled || s.length > 30 || !MOS.test(s) || YOQ.test(s) || bosilgan.includes(s)) continue;
                                if (x.tagName !== 'BUTTON' && x.getAttribute('role') !== 'button' && getComputedStyle(x).cursor !== 'pointer') continue;
                                return x;
                            }
                            if (c.matches && c.matches(DIALOG)) break;
                            c = c.parentElement;
                        }
                        return null;"""
                for _ in range(3):
                    try:
                        t = b.js(TASDIQ_JS, el, bosilgan)
                    except StaleElementReferenceException:   # maydon qayta chizildi
                        t = b.js(TASDIQ_JS, None, bosilgan)
                    if not t:
                        break
                    nomi = (t.text or "").strip() or "tasdiqlash"
                    bosilgan.append(norm_py(nomi))
                    b.bos_el(t)
                    print(f"    OK -> {nomi}")
                    if kuzat(1.0):       # kvitansiya ro'yxatga chiqdi - boshqa tugma kerak emas
                        break
                ok = sahifada() or kuzat(8)
        if not ok:
            b.tahlil("4_kvitansiya")
            if bosilgan:   # raqam yozildi va tasdiqlandi, lekin sayt kvitansiyani ro'yxatga qo'shmadi
                xato = kvitansiya_xatosi(ish, raqam, xabarlar)
                print(f"    !!! {xato}")
                raise KvitansiyaXato(xato)
            odam(f"Kvitansiyani QO'LDA biriktiring: '+' tugmasi -> raqam {raqam}.\nTugagach Enter bosing.")
        else:
            print(f"    OK -> Kvitansiya {raqam} biriktirildi")
    b.tahlil("4_xarajatlar")
    keyingi_majburiy(b)


# ---------------- 5-qadam ----------------
def yuklash_holati(b, yol):
    """Sahifada shu fayl nomi necha marta ko'rinadi, yuklangan fayllar qatorlari ("0.27 mb") soni, spinner."""
    stem = norm_py(os.path.splitext(os.path.basename(yol))[0])[:14]
    try:
        return b.js("return yuklash_holati(arguments[0]);", stem) or {}
    except WebDriverException as e:
        if brauzer_uzildimi(e):
            raise
        return {}


def yuklandimi(h, boshi):
    """Fayl YANGI paydo bo'ldimi. Nom borligi emas, SONI tekshiriladi: masalan 'шартнома' so'zi
    ССП_аъзолик_гувоҳномаси_ва_шартномаси faylida oldindan bor - eski usul shartnoma.pdf ni yuklanmasdan
    "yuklandi" deb hisoblab, keyingi faylni sayt band paytida boshlardi."""
    return h.get("nom", 0) > boshi.get("nom", 0) or h.get("hajm", 0) > boshi.get("hajm", 0)


def yuklanishini_kut(b, yol, boshi):
    """send_keys dan keyin: fayl sahifada paydo bo'lishi va sayt spinneri tugashini kutadi (0.15 s oraliq).
    Sayt hech narsa qilmasa (spinner ham, yangi fayl ham yo'q) - fayl hajmiga qarab 10-40 soniyada False."""
    try:
        mb = os.path.getsize(yol) / 1e6
    except OSError:
        mb = 1.0
    jim = min(KUTISH, 10 + 4 * mb)          # spinnersiz va o'zgarishsiz kutish chegarasi
    boshlandi = harakat = time.time()
    chiqdi = None                            # fayl sahifada paydo bo'lgan payt
    while True:
        h = yuklash_holati(b, yol)
        band = bool(h.get("band")) and not boshi.get("band")   # oldindan turgan spinner hisobga olinmaydi
        faol = h.get("yuk_faol", 0) > 0                         # fayl hali serverga ketyapti
        hozir = time.time()
        if h.get("yuk_xato", 0) > boshi.get("yuk_xato", 0) and not faol:
            print(f"    (server faylni qabul qilmadi: {os.path.basename(yol)})")
            return False
        if yuklandimi(h, boshi):
            chiqdi = chiqdi or hozir
            serverda = h.get("yuk_tugadi", 0) > boshi.get("yuk_tugadi", 0)
            sorovsiz = h.get("yuk_boshlandi", 0) == boshi.get("yuk_boshlandi", 0)   # sayt so'rovni ko'rsatmadi
            # nom sahifada darhol chiqadi, lekin fayl hali yuborilayotgan bo'lishi mumkin - server javobini kutamiz
            if not faol and (serverda or (sorovsiz and hozir - chiqdi > 3)) and (not band or hozir - chiqdi > 8):
                return True
        if band or faol:
            harakat = hozir                  # sayt ishlayapti - kutishda davom etamiz
        if hozir - harakat > jim or hozir - boshlandi > KUTISH * 3 + 20 * mb:
            return False
        uxla(0.15)


def fayl_inputi(b, tomon):
    """Chap ('ariza') yoki o'ng ('ilova') tomonning fayl maydoni. Saytda u oldindan yo'q -
    "Yuklash" bosilganda yaratiladi: shuni ushlab olamiz (Windows oynasi ochilmaydi)."""
    x = b.js("return fayl_maydoni(arguments[0]);", tomon) or {}
    if x.get("el"):
        return x["el"]
    t = x.get("tugma")
    if not t:
        print(f"    !!! '{tomon}' tomonida 'Yuklash' topilmadi")
        return None
    b.bos_el(t, 0.5)
    return b.kut(lambda: b.js("return window.__adolat_oxirgi || null;") or
                 next((i["el"] for i in (b.js("return fayl_inputlar();") or []) if i["tomon"] == tomon), None), 6, 0.1)


def fayl_yuklash(b, tomon, yol, turlar=None, nomi=""):
    """Bitta faylni yuklaydi (o'ng tomonda avval 'Hujjat turi' tanlanadi) va sayt uni qabul qilganini
    tekshiradi. O'tmasa - 3 martagacha qayta urinadi; fayl aslida chiqqan bo'lsa qayta yuklanmaydi (dublikat yo'q)."""
    boshi = yuklash_holati(b, yol)
    for urinish in range(3):
        if urinish:
            if yuklandimi(yuklash_holati(b, yol), boshi):
                return True
            print(f"    (qayta urinish {urinish + 1}/3: {os.path.basename(yol)})")
            b.esc()
            b.band_kut()
        # turi tanlanmasa ham oxirgi urinishda baribir yuklanadi (eski xatti-harakat: tur avvalgidek qoladi)
        if turlar and not b.tanla_yorliq("Hujjat turi", turlar, f"Hujjat turi ({nomi})") and urinish < 2:
            continue
        el = fayl_inputi(b, tomon)
        if not el:
            continue
        try:
            el.send_keys(yol)
        except WebDriverException as e:
            if brauzer_uzildimi(e):
                raise
            print(f"    (fayl berilmadi: {str(e).splitlines()[0][:150]})")
            continue
        if yuklanishini_kut(b, yol, boshi):
            return True
    return yuklandimi(yuklash_holati(b, yol), boshi)


def yetishmayotganlar(b, davo, ilovalar):
    """'Keyingi'dan oldin: sahifada yo'q fayllar. Da'vo arizasi - chap blokda nomi; ilovalar - har birining nomi
    sahifada va "... mb" qatorlari soni ilovalar sonidan kam emas (aks holda qaysi biri yo'qligini nomdan bilamiz)."""
    try:
        # avval hamma yuborilayotgan fayllar serverga yetib borsin ('Keyingi' yarim yuklangan faylni uzib qo'yadi)
        b.kut(lambda: not b.js("return (window.__adolat_tarmoq || {}).yuk_faol || 0;"), KUTISH * 3, 0.2)
        stem = lambda y: norm_py(os.path.splitext(os.path.basename(y))[0])[:14]
        chap = b.js("const r = tomon_ildiz('ariza'); return norm((r || document.body).innerText);") or ""
        sahifa = b.matn()
        yoq = [h for h in davo[:1] if stem(h[2]) not in chap]
        yoq += [h for h in ilovalar if stem(h[2]) not in sahifa]
        qator = (b.js("return yuklash_holati('');") or {}).get("hajm", 0)
        if not yoq and qator < len(ilovalar):
            print(f"    ! sahifada {qator} ta fayl qatori bor, kerak {len(ilovalar)} ta")
        return yoq
    except WebDriverException as e:
        if brauzer_uzildimi(e):
            raise
        return []


def qadam5_hujjatlar(b, ish, oxirgi=False):
    """oxirgi=True ("5-qadamgacha TEST") - yuklab bo'lgach shu yerda to'xtaydi va natijani qaytaradi."""
    print("\n>>> 5-qadam: hujjatlarni biriktirish")
    try:   # zaxira: sayt fayl oynasini baribir ochmoqchi bo'lsa - Chrome uni ko'rsatmaydi
        b.d.execute_cdp_cmd("Page.setInterceptFileChooserDialog", {"enabled": True})
    except Exception:
        pass
    b.kut(lambda: b.js("return !!tomon_ildiz('ariza') || !!tomon_ildiz('ilova');"), KUTISH)
    yuklanmadi = []
    davo = [h for h in ish.hujjatlar if h[0] == "davo"]
    ilovalar = [h for h in ish.hujjatlar if h[0] not in YUKLANMAYDIGAN]

    # Da'vo arizasi - chap tomon ("Arizani biriktirish")
    boshlandi = time.time()
    if davo:
        if fayl_yuklash(b, "ariza", davo[0][2]):
            print(f"    OK -> Ariza (chap): {os.path.basename(davo[0][2])}")
        else:
            print(f"    !!! YUKLANMADI -> Ariza: {os.path.basename(davo[0][2])}  {b.xatolar()}")
            yuklanmadi.append(davo[0])

    # Qolgan hamma hujjat - o'ng tomon ("Hujjat turi" + "Yuklash"). Qat'iy pauza yo'q: keyingi fayl
    # oldingisi sahifada paydo bo'lib, spinner tugagach darhol boshlanadi
    turlar = {k: t for k, _n, _f, t in HUJJAT_TURLARI}
    for kalit, nomi, yol in ilovalar:
        if fayl_yuklash(b, "ilova", yol, turlar.get(kalit) or ["бошқа"], nomi):
            print(f"    OK -> {nomi} (o'ng): {os.path.basename(yol)}")
        else:
            print(f"    !!! YUKLANMADI -> {nomi}: {os.path.basename(yol)}  {b.xatolar()}")
            yuklanmadi.append((kalit, nomi, yol))
    # kafolat: 'Keyingi'dan oldin hamma fayl serverga yetgani va sahifada turgani tekshiriladi; yo'q bo'lsa -
    # bir marta qayta yuklanadi, baribir yo'q bo'lsa 'Keyingi' bosilmaydi (mijoz qaytadan boshlanadi)
    for qayta in range(2):
        yoq = [h for h in yetishmayotganlar(b, davo, ilovalar) if h not in yuklanmadi]
        if not yoq:
            break
        if qayta:
            yuklanmadi += yoq
            break
        for kalit, nomi, yol in yoq:
            print(f"    ! sahifada ko'rinmadi - qayta yuklanmoqda: {os.path.basename(yol)}")
            tomon = "ariza" if kalit == "davo" else "ilova"
            fayl_yuklash(b, tomon, yol, None if tomon == "ariza" else (turlar.get(kalit) or ["бошқа"]), nomi)
    print(f"    ({len(davo[:1]) + len(ilovalar)} ta fayl {time.time() - boshlandi:.0f} soniyada)")

    b.tahlil("5_hujjatlar")
    jami = len(davo[:1]) + len(ilovalar)
    if oxirgi:   # "5-qadamgacha TEST": 'Keyingi' bosilmaydi, murojaat saqlanmaydi
        print("    TEST: 5-qadamda to'xtatildi - 'Keyingi' BOSILMADI, murojaat saqlanmadi va yuborilmadi")
        if yuklanmadi:
            return ("XATO", f"{len(yuklanmadi)}/{jami} ta fayl yuklanmadi: " +
                    ", ".join(os.path.basename(y) for _k, _n, y in yuklanmadi), b.d.current_url)
        print(f"    OK -> Barcha {jami} ta hujjat yuklandi")
        return "YUKLANDI", f"{jami} ta hujjat yuklandi (5-qadamgacha test)", b.d.current_url
    if yuklanmadi:
        odam("Quyidagi fayllar yuklanmadi - brauzerda QO'LDA yuklang, so'ng Enter bosing:\n" +
             "\n".join(f"  [{n}] {y}" for _k, n, y in yuklanmadi))
    else:
        print(f"    OK -> Barcha {jami} ta hujjat yuklandi")
    keyingi_majburiy(b)


# ---------------- 6-qadam ----------------
def qadam6_korib_chiqish(b, ish):
    print("\n>>> 6-qadam: ko'rib chiqish va saqlash")
    b.tahlil("6_korib_chiqish")
    t = b.js("return asosiy_tugma();")
    if not t:
        raise Topilmadi("Ko'rib chiqish sahifasidagi saqlash tugmasi")
    print(f"    -> '{(t.text or '').strip() or 'saqlash'}' bosilmoqda (murojaat saytda saqlanadi, sudga YUBORILMAYDI)")
    b.bos_el(t)
    if not b.kut(lambda: "/receipt/" in b.d.current_url or "tolov kvitantsiya" in b.matn(), KUTISH * 3, 0.3):
        print(f"    !!! Kvitansiya sahifasi ochilmadi. Xabarlar: {b.xatolar()}")
        b.tahlil("6_saqlanmadi")
        odam("Murojaat saqlanmadi. Brauzerda tekshirib, saqlash tugmasini O'ZINGIZ bosing.\n"
             "'To'lov kvitantsiyalari' sahifasi ochilganda skript o'zi davom etadi (yoki Enter).",
             shart=lambda: "/receipt/" in b.d.current_url)
    print(f"    OK -> Murojaat saqlandi: {b.d.current_url}")


# ---------------- 7-qadam ----------------
def qadam7_yuborish(b, ish, rejim):
    print("\n>>> 7-qadam: to'lov kvitansiyalari va sudga yuborish")
    if not b.kut(lambda: b.top("Sudga yuborish", sel="button"), KUTISH * 2):
        b.tahlil("7_tugma_yoq")
        return "NOANIQ", "'Sudga yuborish' tugmasi yo'q (balki avval yuborilgan)", b.d.current_url
    havola = b.d.current_url
    b.bos("Holatni yangilash", "Holatni yangilash", vaqt=3, majburiy=False)
    tin(2)
    holat = b.js("""
        const t = topish(null, 'tolov kvitantsiyalari', null, false, false);
        let c = t; for (let i = 0; i < 3 && c && c.parentElement; i++) c = c.parentElement;
        return ((c || document.body).innerText || '').slice(0, 900);""") or ""
    print("    --- Kvitansiyalar holati ---")
    for q in holat.splitlines():
        if q.strip():
            print("    | " + q.strip())
    if re.search(r"тўланмаган|to.?lanmagan|не оплач", holat, re.I):
        print("    ! DIQQAT: to'lanmagan kvitansiya bor (sayt yuborishga ruxsat bermasligi mumkin)")
    b.tahlil("7_kvitansiyalar")

    if rejim != "yuborish":
        print("    SINOV: murojaat saqlandi, 'Sudga yuborish' BOSILMADI.")
        print("           Yuborish uchun keyin: --yuborish (shu saqlangan murojaat ochiladi)")
        return "SAQLANDI", "sudga yuborilmagan", havola

    if S.get("yuborishdan_oldin_sorash", "Yo'q").strip().lower() in ("ha", "yes", "1", "true"):
        j = odam(f"'{ish.papka}' arizasi SUDGA YUBORILSINMI?\n"
                 f"Brauzerda tekshiring. Yuborish uchun 'ha' yozib Enter bosing (boshqa javob = yubormaslik).", cheksiz=True)
        if (j or "").strip().lower() != "ha":
            return "SAQLANDI", "foydalanuvchi yubormadi", havola
    else:
        print("    -> Sudga yuborilmoqda (so'ramasdan - BOSHLASH da bir marta tasdiqlangan)")

    def muvaffaqiyat():
        m = b.matn()
        return "muvaffaqiyatli yuborildi" in m or "yuborildi" in m or "/cases/lists" in b.d.current_url

    b.bos("Sudga yuborish", "SUDGA YUBORISH", sel="button")
    oxiri = time.time() + KUTISH * 3
    tasdiq = False
    while time.time() < oxiri and not muvaffaqiyat():
        tin(1)
        # sayt "Sud harajatlari to`lanmagan" dedi - yuborilmaydi: holat TOLANMAGAN (to'langach qayta yuboriladi)
        x = next((x for x in b.xatolar() if re.search(r"tolanmag|не оплач", norm_py(x))), None)
        if x or "harajatlari tolanmagan" in b.matn():
            x = x or "Sud harajatlari to'lanmagan"
            b.tahlil("7_tolanmagan")
            print(f"    !!! Sayt: {x} - sudga YUBORILMADI")
            return "TOLANMAGAN", f"sayt: {x} - kvitansiyani to'lab, keyin qayta yuboring", havola
        if not tasdiq:
            d = b.js("""
                for (const d of document.querySelectorAll(DIALOG)) {
                    if (!vis(d)) continue;
                    for (const t of d.querySelectorAll('button')) {
                        const s = txt(t);
                        if (vis(t) && ['ha', 'tasdiqlash', 'yuborish', 'sudga yuborish', 'ok', 'davom etish', 'да'].includes(s)) return t;
                    }
                }
                return null;""")
            if d:
                b.bos_el(d)
                tasdiq = True
                print("    OK -> Tasdiqlash oynasi")
        if re.search(r"e-imzo|eimzo|eri kalit|kalitni tanlang", b.matn()):
            b.tahlil("7_imzo")
            odam("ERI (E-IMZO) bilan imzolash so'ralmoqda: kalitni tanlab imzolang.\n"
                 "Tugagach skript o'zi sezadi (yoki Enter).", shart=muvaffaqiyat, cheksiz=True)
    b.tahlil("7_yuborilgandan_keyin")
    if muvaffaqiyat():
        print("    OK -> SUDGA YUBORILDI")
        return "YUBORILDI", "", havola
    print("    !!! Yuborilgani tasdiqlanmadi - 'Mening murojaatlarim'da tekshiring!")
    return "NOANIQ", f"yuborish bosildi, natija noaniq: {b.xatolar()}", havola


# ---------------- qadamlar ketma-ketligi ----------------
def sahifa_turi(b):
    m = b.matn()
    url = b.d.current_url or ""
    if "/receipt/" in url or "tolov kvitantsiyalari" in m:
        return "kvitansiya"
    if "umumiy malumot" in m and "ishga doir hujjatlar" in m:
        return "korib_chiqish"
    if "hujjatlarni biriktirish" in m or "arizani biriktirish" in m:
        return "hujjatlar"
    if "sud harajatlari" in m or "sud xarajatlari" in m or "tolov imtiyozi" in m:
        return "xarajatlar"
    return "nomalum"


def qolgan_qadamlar(b, ish, rejim):
    bajarildi = set()
    for _ in range(10):
        tin(1)
        tur = b.kut(lambda: (lambda t: t if t != "nomalum" else None)(sahifa_turi(b)), KUTISH) or "nomalum"
        if rejim == "qadam5" and tur in ("korib_chiqish", "kvitansiya"):   # 5-qadam o'tib ketilgan - saqlamaymiz
            b.tahlil("5_qadam_topilmadi")
            return "NOANIQ", "5-qadam (hujjatlar) sahifasi chiqmadi - to'xtatildi, saqlanmadi", b.d.current_url
        if tur == "kvitansiya":
            return qadam7_yuborish(b, ish, rejim)
        if tur in bajarildi:
            print(f"    ! '{tur}' sahifasi qayta ochildi")
        bajarildi.add(tur)
        if tur == "xarajatlar":
            qadam4_xarajatlar(b, ish)
        elif tur == "hujjatlar":
            if rejim == "qadam5":
                return qadam5_hujjatlar(b, ish, oxirgi=True)
            qadam5_hujjatlar(b, ish)
        elif tur == "korib_chiqish":
            qadam6_korib_chiqish(b, ish)
        else:
            print("\n>>> Noma'lum sahifa")
            b.tahlil("nomalum_sahifa")
            keyingi_majburiy(b)
    raise RuntimeError("Kvitansiya sahifasiga yetib bo'lmadi")


UZILISH_BELGILARI = ("invalid session id", "not connected to devtools", "browser has closed", "chrome not reachable",
                     "disconnected", "session deleted", "target closed", "connection refused", "failed to establish")


def brauzer_uzildimi(e):
    """Xato - Chrome yopilgani / ulanish uzilgani (qayta ochish kerak)."""
    return isinstance(e, InvalidSessionIdException) or any(x in str(e).lower() for x in UZILISH_BELGILARI)


def kirish_ishonchli(b, parol, fon=False):
    """Kirish. Chrome uzilib/yopilib qolsa yoki kirish o'tmasa - brauzerni qayta ochib 3 marta urinadi.
    Ishlayotgan brauzerni qaytaradi (yangi bo'lishi mumkin)."""
    for urinish in range(1, 4):
        try:
            if b is None:
                if urinish > 1:
                    print("\n>>> Chrome qayta ochilmoqda...")
                b = Brauzer(fon=fon)
            kirish(b, parol)
            return b
        except (ParolXato, SaytBand):
            raise
        except Exception as e:
            if TOXTAT.is_set():
                raise Toxtatildi()
            if urinish == 3:
                raise
            print(f"\n!!! Kirishda xato ({urinish}-urinish): {type(e).__name__}: {(str(e).strip().splitlines() or [''])[0]}")
            if b is not None and (brauzer_uzildimi(e) or not brauzer_tirikmi(b)):
                print(">>> Brauzer uzilib qoldi - yopilib, qaytadan ochiladi")
                b.yop()
                b = None
            uxla(5)


def brauzer_tirikmi(b):
    """Chrome oynasi hali ochiq va javob beryaptimi (foydalanuvchi yopgan / qulagan bo'lishi mumkin)."""
    try:
        return bool(b.d.window_handles) and b.d.current_url is not None
    except Exception:
        return False


def bitta_ish(b, ish, rejim, parol, oldingi):
    b.prefix = f"{datetime.now():%Y%m%d_%H%M%S}_{xavfsiz_nom(ish.papka)}"
    b.n = 0
    b.robot_toza = {}
    if ish.ocr_avto:
        rp, sp = ish.ocr_ishonch
        print(f"    (raqam/sana OCR bilan avtomatik: {ish.raqam}, {ish.sana:%d.%m.%Y} - ishonch {rp:.0%} / {sp:.0%})")
    old = oldingi.get(ish.papka, {})
    # murojaat saytda allaqachon saqlangan (oldin SAQLANDI yoki saqlangandan keyin XATO) - yangisi yaratilmaydi,
    # aks holda sudda bir xil ariza ikki marta paydo bo'ladi
    if (rejim != "qadam5" and "/receipt/" in old.get("havola", "")
            and (old.get("holat") in ("XATO", "TOLANMAGAN") or (rejim == "yuborish" and old.get("holat") == "SAQLANDI"))):
        print(f"\n>>> Avval saqlangan murojaat ochilmoqda (yangisi yaratilmaydi): {old['havola']}")
        b.d.get(old["havola"])
        tin(2)
        sessiya_tekshir(b, parol)
        if "/receipt/" not in b.d.current_url:
            b.d.get(old["havola"])
        return qadam7_yuborish(b, ish, rejim)
    murojaat_ochish(b, parol)
    boshlash(b)
    qadam1_sud_davogar(b, ish)
    qadam2_asosiy(b, ish)
    qadam3_javobgar(b, ish)
    return qolgan_qadamlar(b, ish, rejim)


# ============================================================
# MAIN
# ============================================================
TEZLIKLAR = {   # sayt ulgurmasa - O'rtacha yoki Sekin
    "Tez": {"pauza_min": "0.2", "pauza_max": "0.45", "harf_pauza": "0.03", "kechikish_koef": "0.5",
            "mijozlar_orasida": "5"},
    "O'rtacha": {"pauza_min": "0.4", "pauza_max": "0.8", "harf_pauza": "0.06", "kechikish_koef": "0.75",
                 "mijozlar_orasida": "15"},
    "Sekin": {"pauza_min": "0.7", "pauza_max": "1.5", "harf_pauza": "0.12", "kechikish_koef": "1",
              "mijozlar_orasida": "30"},
}


def sozlamalarni_qollash():
    global S, KUTISH
    S = sozlamalarni_yuklash()
    S.update(TEZLIKLAR.get(S.get("tezlik", ""), {}))      # tanlangan tezlik pauzalarni belgilaydi
    try:
        KUTISH = max(10, int(S["kutish_soniya"]))
    except ValueError:
        KUTISH = 40


URINISHLAR = 3   # xato bo'lsa bitta mijoz ketma-ket necha marta urinadi; o'tmasa - navbat oxirida yana bir marta


def qayta_urinsa_boladimi(e):
    """Qayta urinish foyda beradimi: kvitansiya noto'g'ri va dastur xatosi (NameError) - yo'q;
    qolgani (CAPTCHA misoli, sayt sekin, fayl yuklanmadi, 'Keyingi' o'tmadi, ...) - ha.
    (Vafot etgan / topilmadi - ishga_tushirish da alohida, u ham qayta urinilmaydi.)"""
    return not isinstance(e, (KvitansiyaXato, NameError, ImportError))


def holat_matni(holat, izoh=""):
    """Jadvaldagi holat yozuvi: XATO/TOPILMADI sababi qisqa ko'rsatiladi."""
    iz = norm_py(izoh)
    if holat == "TOPILMADI":
        return "VAFOT ETGAN" if VAFOT_RE.search(iz) else "TOPILMADI (vafot?)"
    if holat == "XATO" and iz.startswith("kvitansiya notogri"):
        return "XATO: KVITANSIYA"
    if holat == "TOLANMAGAN":
        return "TO'LANMAGAN"
    return holat


def ishlansinmi(ish, oldingi, rejim, qayta):
    h = oldingi.get(ish.papka, {}).get("holat", "")
    if ish.xatolar:
        return False
    if rejim == "qadam5":   # hech narsa saqlanmaydi/yuborilmaydi - istalgan mijozda sinab ko'rish mumkin
        return True
    if h in ("YUBORILDI", "NOANIQ", "TOPILMADI"):   # TOPILMADI - vafot etgan: qayta urinilmaydi
        return qayta
    if h == "SAQLANDI" and rejim == "sinov":
        return qayta
    return True


def ishga_tushirish(rejim, filtr=None, tanlangan=None, qayta=False, fon=False, terminal_tasdiq=False):
    """Asosiy jarayon (terminal va oyna uchun umumiy). Natija: [(papka, holat, izoh)]."""
    sozlamalarni_qollash()
    TOXTAT.clear()
    ishlar = ishlarni_tayyorlash(filtr)
    if tanlangan:
        ishlar = [i for i in ishlar if i.papka in tanlangan]
    oldingi = natijalar_oqish()
    ishlarni_korsatish(ishlar, oldingi)
    tayyor = [i for i in ishlar if ishlansinmi(i, oldingi, rejim, qayta)]
    otkazildi = [i for i in ishlar if i not in tayyor]
    if otkazildi:
        print("O'tkazib yuboriladi:")
        for i in otkazildi:
            print(f"   - {i.papka}: {'; '.join(i.xatolar) or oldingi.get(i.papka, {}).get('holat', '')}")
    if not tayyor:
        print("Ishlanadigan mijoz yo'q.")
        return []

    print("\nREJIM: " + {"yuborish": "!!! SUDGA YUBORISH !!!",
                         "qadam5": "TEST 5-QADAMGACHA (hujjatlar yuklanadi, keyin TO'XTAYDI - saqlanmaydi, yuborilmaydi)"}.get(
        rejim, "TAYYORLASH (saytda saqlanadi, sudga yuborilmaydi)"))
    print(f"Ishlanadi: {'1' if rejim == 'qadam5' else len(tayyor)} ta mijoz")
    if terminal_tasdiq:
        try:
            javob = input("Boshlaymizmi? (Enter = ha, 'yoq' = chiqish): ")
        except EOFError:
            javob = ""
        if javob.strip().lower() in ("yoq", "yo'q", "n", "no"):
            return []
    ini_paroldan_olish()
    parol = parol_olish()
    if not parol:
        print("!!! Parol kiritilmadi.")
        return []

    print("\n>>> Chrome ochilmoqda...")
    b = None
    hisobot = []
    boshlandi = time.time()
    joriy = None   # hozir ishlanayotgan mijoz (to'xtatilganda hisobotga yoziladi)
    toxtatildi = False
    global ODAM_KUTISH
    ODAM_KUTISH = _son("xatoda_kutish", 0)   # 0 - xato bo'lsa so'ramasdan keyingi mijoz (navbat to'xtamaydi)

    def mijoz(ish, urinishlar):
        """Bitta mijoz. Xato bo'lsa sahifa yangidan ochilib shu mijoz qaytadan - `urinishlar` martagacha.
        Vafot etgan / topilmadi va kvitansiya noto'g'ri - qayta urinilmaydi. Natija: (holat, izoh, havola, qayta_mumkin)."""
        nonlocal b, oldingi
        for urinish in range(1, urinishlar + 1):
            try:
                holat, izoh, havola = bitta_ish(b, ish, rejim, parol, oldingi)
                return holat, izoh, havola, False
            except (ParolXato, SaytBand):
                raise
            except Exception as e:
                if TOXTAT.is_set():             # To'xtatish - Chrome yopildi, qayta urinilmaydi
                    raise Toxtatildi()
                xabar = (str(e).strip().splitlines() or [""])[0]
                if isinstance(e, ShaxsTopilmadi) and TOPILMADI_RE.search(norm_py(str(e))):
                    sayt = str(e).strip()[:200]
                    izoh = (f"VAFOT ETGAN - sayt: {sayt}" if VAFOT_RE.search(norm_py(sayt))
                            else f"javobgar saytda topilmadi (JShShIR xato?) - sayt: {sayt}")
                    print(f"\n!!! {izoh}\n>>> Keyingi mijozga o'tilmoqda (qayta urinilmaydi)")
                    return "TOPILMADI", izoh, "", False
                if isinstance(e, KvitansiyaXato):
                    print(f"\n!!! {xabar}\n>>> Keyingi mijozga o'tilmoqda (kvitansiyani tuzating - qayta urinilmaydi)")
                    return "XATO", xabar, "", False
                if isinstance(e, ShaxsTopilmadi):
                    izoh = f"sayt xatosi: {str(e).strip()[:200]}"
                elif isinstance(e, (OdamKelmadi, CaptchaMisol)):
                    izoh = xabar
                else:
                    izoh = f"{type(e).__name__}: {xabar}"
                havola = ""
                tirik = brauzer_tirikmi(b)
                if tirik:
                    havola = b.d.current_url
                print(f"\n!!! XATO ({urinish}/{urinishlar}-urinish): {izoh}")
                if not isinstance(e, (OdamKelmadi, CaptchaMisol)):
                    traceback.print_exc()
                if tirik:
                    b.tahlil("XATO")
                else:                                   # Chrome yopilib/qulab qoldi - navbat to'xtamasin
                    print("\n>>> Brauzer yopilib qoldi - qayta ochilib, kabinetga kirilmoqda...")
                    b.yop()
                    b = kirish_ishonchli(None, parol, fon)
                if "/receipt/" in (havola or ""):        # saqlangan - qayta ochiladi, yangisi yaratilmaydi
                    oldingi = {**oldingi, ish.papka: {"holat": "XATO", "havola": havola}}
                qayta = qayta_urinsa_boladimi(e)
                if not qayta or urinish == urinishlar:
                    return "XATO", izoh, havola, qayta
                print(f"\n>>> Shu mijoz QAYTADAN boshlanmoqda ({urinish + 1}/{urinishlar}-urinish, sahifa yangidan ochiladi)")
                b.alert_yop()
                b.esc()
                uxla(3)

    def yakunla(ish, holat, izoh, havola, t0):
        # 5-qadamgacha test natijasi yozilmaydi (avvalgi holat/havola buzilmasin); TOPILMADI - yoziladi
        if rejim != "qadam5" or holat == "TOPILMADI":
            natija_yozish(ish, holat, izoh, havola, rejim)
        if rejim != "qadam5" and holat in ("YUBORILDI", "SAQLANDI", "TOLANMAGAN"):
            raqam_tarixiga(ish)     # keyingi yuklashlarda paket qoidasi va sanalar uchun dalil
        for i, (p, _h, _iz) in enumerate(hisobot):
            if p == ish.papka:      # yakuniy aylanish - birinchi natija almashtiriladi
                hisobot[i] = (ish.papka, holat, izoh)
                break
        else:
            hisobot.append((ish.papka, holat, izoh))
        belgi = holat_matni(holat, izoh)
        print(f"\n>>> NATIJA: {ish.papka} = {holat}" + (f" [{belgi}]" if belgi != holat else "")
              + f" ({(time.time() - t0) / 60:.1f} daqiqa)")

    def navbat(royxat, urinishlar):
        """Mijozlarni ketma-ket ishlaydi. Qayta urinsa bo'ladigan XATO lar ro'yxatini qaytaradi."""
        nonlocal joriy
        qolgan = []
        for n, ish in enumerate(royxat, 1):
            toxtat_tekshir()
            if n > 1 or hisobot:
                dam = _son("mijozlar_orasida", 30)
                print(f"\n... keyingi mijozdan oldin ~{dam:.0f} soniya dam (sayt bloklamasligi uchun)")
                uxla(random.uniform(dam * 0.7, dam * 1.3))
            joriy = ish
            print("\n" + "#" * 70)
            print(f"# [{n}/{len(royxat)}] {ish.papka}")
            print("#" * 70)
            t0 = time.time()
            holat, izoh, havola, qayta = mijoz(ish, urinishlar)
            yakunla(ish, holat, izoh, havola, t0)
            joriy = None
            if holat == "XATO" and qayta:
                qolgan.append(ish)
            if rejim == "qadam5" and holat != "TOPILMADI":   # test: bitta mijoz (topilmasa - keyingisi)
                return []
        return qolgan

    try:
        try:
            b = kirish_ishonchli(b, parol, fon)
            qayta_navbat = navbat(tayyor, URINISHLAR)
            if qayta_navbat:
                print("\n" + "=" * 70)
                print(f"YAKUNIY AYLANISH: {URINISHLAR} urinishda o'tmagan {len(qayta_navbat)} ta mijoz yana bir marta ishlanadi")
                print("=" * 70)
                navbat(qayta_navbat, 1)
        except SaytBand as e:
            print(f"\n!!! {e}")
            b and b.tahlil("tizim_yangilanmadi")
            hisobot.append(("(kirish)", "XATO", str(e)))
        except ParolXato as e:
            print(f"\n!!! {e}. Saqlangan parol o'chirildi - keyingi safar yangi parol so'raladi.")
            parol_ochirish()
            b and b.tahlil("parol_xato")
            hisobot.append(("(kirish)", "XATO", str(e)))
        except Exception as e:
            if TOXTAT.is_set():
                raise Toxtatildi()
            print(f"\n!!! To'xtadi: {type(e).__name__}: {e}")
            traceback.print_exc()
            b and b.tahlil("umumiy_XATO")
            hisobot.append(("(umumiy)", "XATO", f"{type(e).__name__}: {e}"))
    except Toxtatildi:
        toxtatildi = True
        print("\n>>> TO'XTATILDI (foydalanuvchi 'To'xtatish' tugmasini bosdi) - jarayon to'xtatildi, Chrome yopildi")
        if joriy:
            hisobot.append((joriy.papka, "TOXTATILDI", "foydalanuvchi to'xtatdi"))
            print(f"\n>>> NATIJA: {joriy.papka} = TOXTATILDI")
    TOXTAT.clear()   # yakuniy xabar va brauzerni yopish to'xtatilmasin
    ODAM_KUTISH = None

    print("\n" + "=" * 70)
    print(f"YAKUN ({(time.time() - boshlandi) / 60:.0f} daqiqa):")
    for p, h, iz in hisobot:
        print(f"  {h:<10} {p}  {iz}")
    for h in ("YUBORILDI", "SAQLANDI", "TOLANMAGAN", "YUKLANDI", "XATO", "NOANIQ", "TOPILMADI", "TOXTATILDI"):
        k = sum(1 for _p, x, _i in hisobot if x == h)
        if k:
            print(f"  {h}: {k} ta")
    print("(natijalar.csv va loglar\\ papkasida saqlandi)")
    print("=" * 70)
    # hech narsa so'ralmaydi: ish tugadi - Chrome yopiladi; TEST 5-qadam - Chrome ko'rib chiqish uchun ochiq qoladi
    if b and rejim == "qadam5" and not toxtatildi and any(h == "YUKLANDI" for _p, h, _i in hisobot):
        print("\n>>> TEST 5-qadamda to'xtadi: murojaat SAQLANMADI va YUBORILMADI.\n"
              "    Chrome ochiq qoldirildi - yuklangan hujjatlarni ko'rib chiqing, keyin o'zingiz yoping.")
        b.ajrat()
    elif b:
        b.yop()
    global JORIY_BRAUZER
    JORIY_BRAUZER = None
    return hisobot


def main():
    ap = argparse.ArgumentParser(description="ADOLAT da'vo arizalarini avtomatik yuborish")
    ap.add_argument("--yuborish", action="store_true", help="haqiqiy yuborish (aks holda faqat tayyorlash va saqlash)")
    ap.add_argument("--mijoz", help="faqat nomida shu so'z bor mijoz papkalari")
    ap.add_argument("--jadval", action="store_true", help="ishlar.xlsx ni yaratish/yangilash")
    ap.add_argument("--tekshir", action="store_true", help="faqat ma'lumotlarni tekshirish (brauzersiz)")
    ap.add_argument("--qayta", action="store_true", help="avval saqlangan/yuborilganlarni ham qaytadan to'ldirish")
    ap.add_argument("--parol", action="store_true", help="saqlangan parolni almashtirish (qayta so'raladi)")
    ap.add_argument("--brauzer-tekshir", action="store_true", help="Chrome ishga tushishini tekshirish")
    ap.add_argument("--terminal", action="store_true", help="oynasiz (terminal) rejim")
    ap.add_argument("--fon", action="store_true", help=argparse.SUPPRESS)
    args = ap.parse_args()

    log = log_boshlash()
    sozlamalarni_qollash()
    print("=" * 70)
    print(f"ADOLAT AVTOMAT  versiya {VERSIYA}")
    print(f"Papka: {PAPKA}")
    print(f"Log:   {log}")
    print("=" * 70)

    if args.jadval:
        jadval_yangilash()
        return
    if args.brauzer_tekshir:
        b = Brauzer(fon=True)
        print(f"OK -> Chrome {b.d.capabilities.get('browserVersion')} ishga tushdi, chromedriver tayyor"
              f" ({'oddiy Chrome' if b.ulangan else 'Selenium Chrome'}).")
        b.yop()
        return
    if args.tekshir:
        ishlar = ishlarni_tayyorlash(args.mijoz)
        oldingi = natijalar_oqish()
        ishlarni_korsatish(ishlar, oldingi)
        tayyor = [i for i in ishlar if ishlansinmi(i, oldingi, "yuborish" if args.yuborish else "sinov", args.qayta)]
        print(f"\nTekshiruv tugadi: {len(tayyor)} ta ish ishlashga tayyor.")
        return
    if args.parol:
        parol_ochirish()
    ishga_tushirish("yuborish" if args.yuborish else "sinov", filtr=args.mijoz, qayta=args.qayta,
                    fon=args.fon, terminal_tasdiq=True)


# ============================================================
# OYNA (dastur interfeysi)
# ============================================================
def jadval_qator_saqlash(papka, qiymatlar):
    """ishlar.xlsx dagi bitta mijoz qatorini yangilaydi (yo'q bo'lsa qo'shadi)."""
    import openpyxl
    if not os.path.isfile(JADVAL):
        jadval_yangilash()
    wb = openpyxl.load_workbook(JADVAL)
    ws = wb.active
    sarlavha = [str(c.value or "").strip() for c in ws[1]]
    qator = None
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value or "").strip() == papka:
            qator = row
            break
    if qator is None:
        ws.append([papka] + [""] * (len(sarlavha) - 1))
        qator = ws[ws.max_row]
    for k, v in qiymatlar.items():
        if k not in sarlavha and k in JADVAL_USTUNLAR:     # eski jadvalda yangi ustun ("Sud") yo'q - qo'shamiz
            ws.cell(row=1, column=len(sarlavha) + 1, value=k)
            sarlavha.append(k)
            if k == "Tekshirildi":   # ustun yo'q paytda jadvaldagi raqam/sana qabul qilingan edi - tasdiqlangan
                for r in ws.iter_rows(min_row=2):
                    d = {sarlavha[i]: r[i].value for i in range(min(len(sarlavha) - 1, len(r)))}
                    if r[0].row != qator[0].row and tekshirilganmi(d):
                        ws.cell(row=r[0].row, column=len(sarlavha), value="ha")
        if k in sarlavha:
            c = ws.cell(row=qator[0].row, column=sarlavha.index(k) + 1)
            c.value = "" if v is None else str(v)
            c.number_format = "@"
    wb.save(JADVAL)


def sozlamalarni_saqlash(qiymatlar):
    cp = configparser.ConfigParser(interpolation=None)
    if os.path.isfile(SOZLAMALAR_FAYL):          # KvitansiyaYaratish.exe ning kvitansiya_* kalitlari saqlanib qolsin
        cp.read(SOZLAMALAR_FAYL, encoding="utf-8-sig")
    if not cp.has_section("ADOLAT"):
        cp.add_section("ADOLAT")
    for k in STANDART_SOZLAMALAR:
        cp.set("ADOLAT", k, str(qiymatlar.get(k, S.get(k, STANDART_SOZLAMALAR[k]))).strip())
    cp.set("ADOLAT", "parol", "")
    with open(SOZLAMALAR_FAYL, "w", encoding="utf-8") as f:
        f.write(SOZLAMA_IZOHLARI)
        cp.write(f)


def gui_main():
    global ODAM_ISHLOVCHI, PAROL_SORASH
    import tkinter as tk
    from tkinter import messagebox, scrolledtext, simpledialog, ttk

    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    uiq = queue.Queue()

    class OynaYozuvchi:
        def write(self, s):
            if s:
                uiq.put(("log", s))

        def flush(self):
            pass

    log = log_boshlash(OynaYozuvchi())
    os.makedirs(MIJOZLAR_PAPKA, exist_ok=True)
    os.makedirs(UMUMIY_PAPKA, exist_ok=True)
    sozlamalarni_qollash()

    root = tk.Tk()
    root.title(f"{DASTUR_NOMI} - da'vo arizalarini sudga avtomatik yuborish")
    root.geometry("1280x820")
    root.minsize(1000, 650)
    try:
        root.state("zoomed")
    except tk.TclError:
        pass
    st = ttk.Style()
    try:
        st.theme_use("vista")
    except tk.TclError:
        pass
    SHRIFT = ("Segoe UI", 10)
    st.configure("Treeview", font=SHRIFT, rowheight=26)
    st.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    st.configure("Katta.TButton", font=("Segoe UI", 11, "bold"), padding=8)
    st.configure("Toxtat.TButton", font=("Segoe UI", 11, "bold"), padding=8, foreground="#C62828")
    st.configure("Belgi.TCheckbutton", font=("Segoe UI", 10, "bold"))
    st.configure("TButton", font=SHRIFT, padding=5)

    holat = {"ishlayapti": False, "odam": None, "ocr": False, "yuklanmoqda": False, "qayta_yuklash": False,
             "oxirgi_sana": None}
    avto = os.environ.get("ADOLAT_AVTO_TEST", "")   # faqat ichki avtomatik test uchun
    if avto:
        messagebox.askyesno = lambda *a, **k: True
        messagebox.showinfo = lambda *a, **k: print("    [oyna xabari] " + " | ".join(str(x) for x in a))

    def skrin(nom):
        if not os.environ.get("ADOLAT_SKRIN"):
            return
        try:
            from PIL import ImageGrab
            root.attributes("-topmost", True)
            root.lift()
            root.update()
            time.sleep(0.4)
            x, y = root.winfo_rootx(), root.winfo_rooty()
            ImageGrab.grab((x, y, x + root.winfo_width(), y + root.winfo_height())).save(
                os.path.join(os.environ["ADOLAT_SKRIN"], f"{nom}.png"))
            root.attributes("-topmost", False)
        except Exception as e:
            print(f"(skrin xatosi: {e})")

    # ---------- yuqori panel ----------
    tepa = ttk.Frame(root, padding=(10, 8))
    tepa.pack(fill="x")
    ttk.Label(tepa, text=DASTUR_NOMI, font=("Segoe UI", 16, "bold")).pack(side="left")
    ttk.Label(tepa, text=f"   versiya {VERSIYA}", foreground="#777").pack(side="left")

    tugmalar = ttk.Frame(root, padding=(10, 0))
    tugmalar.pack(fill="x")
    tugmalar2 = ttk.Frame(root, padding=(10, 0))
    tugmalar2.pack(fill="x")

    # ---------- odam kerak bo'lganda chiqadigan sariq panel ----------
    banner = tk.Frame(root, bg="#FFE58F", bd=0, highlightthickness=2, highlightbackground="#E0A800")
    banner_matn = tk.Label(banner, text="", bg="#FFE58F", font=("Segoe UI", 12, "bold"), justify="left", anchor="w")
    banner_matn.pack(side="left", fill="x", expand=True, padx=12, pady=10)
    banner_tugmalar = tk.Frame(banner, bg="#FFE58F")
    banner_tugmalar.pack(side="right", padx=10)

    # ---------- holat qatori ----------
    holat_qator = ttk.Frame(root, padding=(10, 4))
    holat_qator.pack(fill="x")
    holat_yozuv = ttk.Label(holat_qator, text="Tayyor. Mijozlarni tanlang yoki hammasini ishga tushiring.",
                            font=("Segoe UI", 10, "bold"))
    holat_yozuv.pack(side="left")
    progress = ttk.Progressbar(holat_qator, length=300, mode="determinate")
    progress.pack(side="right")
    ocr_yozuv = ttk.Label(holat_qator, text="", foreground="#1565C0")
    ocr_yozuv.pack(side="right", padx=8)

    # ---------- mijozlar jadvali + log ----------
    bolinma = ttk.PanedWindow(root, orient="vertical")
    bolinma.pack(fill="both", expand=True, padx=10, pady=(4, 10))
    jadval_ramka = ttk.Frame(bolinma)
    ustunlar = [("n", "№", 40), ("papka", "Mijoz (papka)", 230), ("raqam", "Ariza №", 80), ("sana", "Sana", 90),
                ("javobgar", "Javobgar", 260), ("jshshir", "JShShIR", 130), ("qarz", "Asosiy qarz", 110),
                ("hujjat", "Hujjatlar", 75), ("holat", "Holat", 110), ("izoh", "Izoh / kamchilik", 360),
                ("sud", "Sud (da'vo ariza / manzil)", 260)]
    tv = ttk.Treeview(jadval_ramka, columns=[u[0] for u in ustunlar], show="headings", selectmode="extended")
    for k, sarlavha, en in ustunlar:
        tv.heading(k, text=sarlavha)
        tv.column(k, width=en, anchor="w", stretch=k in ("javobgar", "izoh"))
    tv.tag_configure("YUBORILDI", background="#C8F7C5")
    tv.tag_configure("SAQLANDI", background="#CFE8FF")
    tv.tag_configure("XATO", background="#FFC9C9")
    tv.tag_configure("NOANIQ", background="#FFE0B2")
    tv.tag_configure("TOLANMAGAN", background="#F8BBD0")   # saqlangan, lekin sud xarajati to'lanmagan
    tv.tag_configure("KAMCHILIK", background="#FFF3C4")
    tv.tag_configure("ISHLANMOQDA", background="#E1D5FF")
    tv.tag_configure("YUKLANDI", background="#D5F5E3")
    tv.tag_configure("TOXTATILDI", background="#E0E0E0")
    tv.tag_configure("TOPILMADI", background="#D7CCC8")   # javobgar saytda yo'q (vafot etgan)
    tv.tag_configure("TEKSHIR", background="#FFD8A8")     # raqam/sana ishonchsiz - odam ko'rishi kerak
    tv.tag_configure("OCR_KUT", background="#E3F2FD")     # OCR hali o'qimagan
    tv.tag_configure("AVTO", background="#E8F5E9")        # raqam/sana OCR bilan avtomatik
    sb = ttk.Scrollbar(jadval_ramka, orient="vertical", command=tv.yview)
    tv.configure(yscrollcommand=sb.set)
    tv.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    bolinma.add(jadval_ramka, weight=3)

    log_ramka = ttk.Frame(bolinma)
    ttk.Label(log_ramka, text="Jarayon (nima bo'layotgani):", font=("Segoe UI", 10, "bold")).pack(anchor="w")
    log_oyna = scrolledtext.ScrolledText(log_ramka, height=12, font=("Consolas", 10), bg="#1E1E1E", fg="#E8E8E8",
                                         insertbackground="white", wrap="word")
    log_oyna.pack(fill="both", expand=True)
    log_oyna.tag_configure("ok", foreground="#7EE787")
    log_oyna.tag_configure("xato", foreground="#FF7B72")
    log_oyna.tag_configure("bosh", foreground="#79C0FF")
    log_oyna.tag_configure("odam", foreground="#FFD866")
    bolinma.add(log_ramka, weight=2)

    ishlar_xotira = {}

    def jadvalni_yuklash(keyin=None):
        """Ro'yxatni fonda tayyorlaydi (600 ta mijozda ham oyna qotmasin), tayyor bo'lgach jadvalni chizadi."""
        if holat["yuklanmoqda"]:
            holat["qayta_yuklash"] = True
            return
        holat["yuklanmoqda"] = True
        if not ishlar_xotira:
            holat_yozuv.config(text="Mijozlar ro'yxati tayyorlanmoqda...")

        def ishchi():
            try:
                ishlar = ishlarni_tayyorlash(ocr=False)
            except Exception as e:
                print(f"!!! Ro'yxat tayyorlanmadi: {type(e).__name__}: {e}")
                traceback.print_exc()
                ishlar = None
            uiq.put(("ishlar", ishlar, keyin))

        threading.Thread(target=ishchi, daemon=True).start()

    def qator_guruhi(ish, h):
        """Ro'yxatdagi o'rni: 0 - raqam/sanani tekshirish kerak (eng tepada), 1 - OCR navbatda,
        2 - boshqa kamchilik, 3 - tayyor, 4 - bajarilgan."""
        if h in ("YUBORILDI", "SAQLANDI", "TOPILMADI"):
            return 4
        if ish.tekshir_kerak:
            return 0 if (ish.qolyozma is not None or not ish.ariza_pdf) else 1
        return 2 if ish.xatolar else 3

    def jadvalni_chizish(ishlar):
        tanlangan = set(tanlanganlar())
        for i in tv.get_children():
            tv.delete(i)
        oldingi = natijalar_oqish()
        h = {i.papka: oldingi.get(i.papka, {}).get("holat", "") for i in ishlar}
        # bir guruh ichida sana va raqam bo'yicha: bir kunda yozilgan arizalar yonma-yon turadi
        tartib = sorted(ishlar, key=lambda i: (qator_guruhi(i, h[i.papka]), i.sana or date.max, i.raqam or "", i.papka))
        ishlar_xotira.clear()
        for n, ish in enumerate(tartib, 1):
            ishlar_xotira[ish.papka] = ish
            g = qator_guruhi(ish, h[ish.papka])
            if h[ish.papka]:
                teg = h[ish.papka]
                matn = holat_matni(teg, oldingi.get(ish.papka, {}).get("izoh", ""))
            else:
                teg, matn = {0: ("TEKSHIR", "TEKSHIRING"), 1: ("OCR_KUT", "OCR..."), 2: ("KAMCHILIK", "TO'LDIRING")}.get(
                    g, ("AVTO", "tayyor (OCR)") if ish.ocr_avto else ("", "tayyor"))
            kam = ish.xatolar + [o for o in ish.ogohlantirishlar if "katta fayl" not in o]
            iid = tv.insert("", "end", values=(n, ish.papka, ish.raqam, ish.sana.strftime("%d.%m.%Y") if ish.sana else "",
                                               ish.javobgar, ish.jshshir, ish.asosiy_qarz, len(ish.hujjatlar),
                                               matn, "; ".join(kam), ish.sud["nomi"] if ish.sud else "?"), tags=(teg,))
            if ish.papka in tanlangan:
                tv.selection_add(iid)
        tayyor = sum(1 for i in ishlar if not i.xatolar)
        avto_soni = sum(1 for i in ishlar if i.ocr_avto and not i.xatolar)
        tek = sum(1 for i in ishlar if qator_guruhi(i, "") == 0)
        kut = sum(1 for i in ishlar if qator_guruhi(i, "") == 1)
        holat_yozuv.config(text=f"Mijozlar: {len(ishlar)} ta, ishlashga tayyor: {tayyor} ta"
                                + (f" (raqam/sanasi OCR bilan avtomatik: {avto_soni})" if avto_soni else "") + "."
                                + (f"   TEKSHIRISH KERAK: {tek} ta (ro'yxat boshida) -> '✍ Raqam va sanani kiritish'"
                                   if tek else "")
                                + (f"   OCR navbatda: {kut} ta" if kut else ""))

    def bitta_yangilash(papka):
        """Bitta mijoz o'zgargach: uni qayta tayyorlab, hamma raqam/sanani qayta baholaydi (qo'shnilari ham)."""
        try:
            ish = ish_tayyorlash(papka, jadval_oqish().get(papka), papka_hujjatlari(UMUMIY_PAPKA), ocr=False)
        except Exception as e:
            print(f"    (mijoz qayta o'qilmadi: {type(e).__name__}: {e})")
            jadvalni_yuklash()
            return
        ishlar_xotira[papka] = ish
        ishlar = list(ishlar_xotira.values())
        raqam_sana_baholash(ishlar)
        jadvalni_chizish(ishlar)

    def ocr_fonda():
        """O'qilmagan arizalarni fonda OCR qiladi: ro'yxat darhol ko'rinadi, natijalar kelgan sari to'ladi."""
        if holat["ocr"]:
            return
        holat["ocr"] = True
        tugmalar_holati(holat["ishlayapti"])

        def ishchi():
            soni = 0
            try:
                kerak = ocr_oqilmaganlar()
                soni = len(kerak)
                for n, (papka, pdf, docx) in enumerate(kerak, 1):
                    uiq.put(("ocr_holat", n, soni, papka))
                    try:
                        ocr_bitta(pdf, docx)
                    except Exception as e:
                        print(f"    !!! OCR xatosi ({papka}): {type(e).__name__}: {e}")
                    if n % 10 == 0:
                        uiq.put(("ocr_jadval",))
            except Exception as e:
                print(f"!!! OCR to'xtadi: {type(e).__name__}: {e}")
                traceback.print_exc()
            uiq.put(("ocr_tugadi", soni))

        threading.Thread(target=ishchi, daemon=True).start()

    def tanlanganlar():
        return [tv.item(i, "values")[1] for i in tv.selection()]

    def qator_ozgartir(_e=None, papka=None, navbat_qolgan=None):
        """Mijoz oynasi: arizaning qo'lda yozilgan sana/raqam rasmi + maydonlar. OCR ishonchsiz o'qigan maydon -
        sariq, eng ehtimolli variant oldindan yozilgan. navbat_qolgan berilsa - saqlagach keyingi tekshiriladigan
        mijoz o'zi ochiladi (ketma-ket kiritish)."""
        if papka is None:
            sel = tanlanganlar()
            if not sel:
                return
            papka = sel[0]
        ish = ishlar_xotira.get(papka)
        T = ocr_chegara()
        rp, sp = ish.ocr_ishonch if ish and ish.ocr_ishonch else (None, None)
        tavsiya = ish.ocr_tavsiya if ish and ish.ocr_tavsiya else ("", None, False, False, False)
        r_shubha = bool(ish and ish.tekshir_kerak and (rp is None or rp < T or tavsiya[3]))
        s_shubha = bool(ish and ish.tekshir_kerak and (sp is None or sp < T or tavsiya[4]))
        d = tk.Toplevel(root)
        d.title(f"Mijoz: {papka}" + (f"   (yana {len(navbat_qolgan)} ta qoldi)" if navbat_qolgan is not None else ""))
        d.transient(root)
        d.grab_set()
        qator = 0
        png = ariza_bosh_rasmi(ish.ariza_pdf) if ish and ish.ariza_pdf else None
        if png:
            import base64
            rasm = tk.PhotoImage(data=base64.b64encode(png).decode("ascii"))
            if ish and ish.jadval_shubha:
                sarlavha = "Jadvalda tasdiqlangan qiymat dalillarga ZID - RASM BILAN SOLISHTIRING. To'g'ri bo'lsa Enter:"
            elif ish and ish.tekshir_kerak and (ish.raqam or ish.sana):
                sarlavha = "Dastur o'qidi, lekin SARIQ maydonlarda ishonchi past - RASM BILAN SOLISHTIRING. To'g'ri bo'lsa Enter:"
            elif ish and ish.ocr_avto:
                sarlavha = "Dastur (OCR) avtomatik o'qidi - rasm bilan solishtirib ko'ring:"
            else:
                sarlavha = "Arizadagi qo'lda yozilgan sana va raqam:"
            ttk.Label(d, text=sarlavha, font=("Segoe UI", 10, "bold"),
                      foreground="#C62828" if ish and ish.tekshir_kerak else "").grid(
                row=qator, column=0, columnspan=4, sticky="w", padx=10, pady=(10, 2))
            qator += 1
            rasm_yozuv = tk.Label(d, image=rasm, bd=1, relief="solid")
            rasm_yozuv.image = rasm
            rasm_yozuv.grid(row=qator, column=0, columnspan=4, padx=10, pady=(0, 8))
            qator += 1
        katta = ("Segoe UI", 14)
        s = ish.sana if ish and ish.sana else None
        if s is None and holat["oxirgi_sana"]:     # arizalar odatda bir kunda yozilgan - oxirgi kiritilgan sana
            s, s_shubha = holat["oxirgi_sana"], True

        def kiritish(ota, en, shubha, **kw):
            return tk.Entry(ota, width=en, font=katta, relief="solid", bd=1, bg="#FFD8A8" if shubha else "white", **kw)

        ttk.Label(d, text="Sana:", font=katta).grid(row=qator, column=0, sticky="e", padx=(10, 4))
        sana_f = ttk.Frame(d)
        sana_f.grid(row=qator, column=1, columnspan=3, sticky="w")
        e_kun = kiritish(sana_f, 4, s_shubha, justify="center")
        e_oy = kiritish(sana_f, 4, s_shubha, justify="center")
        e_yil = kiritish(sana_f, 6, False, justify="center")
        e_kun.insert(0, f"{s.day:02d}" if s else "")
        e_oy.insert(0, f"{s.month:02d}" if s else "")
        e_yil.insert(0, str(s.year if s else date.today().year))
        for w, matn in ((e_kun, "kun"), (e_oy, "oy"), (e_yil, "yil")):
            w.pack(side="left")
            ttk.Label(sana_f, text=f" {matn}   ", foreground="#777").pack(side="left")
        qator += 1
        ttk.Label(d, text="Raqam №:", font=katta).grid(row=qator, column=0, sticky="e", padx=(10, 4), pady=6)
        e_raqam = kiritish(d, 14, r_shubha)
        e_raqam.insert(0, ish.raqam if ish else "")
        e_raqam.grid(row=qator, column=1, sticky="w", pady=6)
        ttk.Label(d, text="/ 09-02  (faqat raqam)", foreground="#777").grid(row=qator, column=2, sticky="w")
        qator += 1
        if ish and ish.ocr_variantlar:
            # variantlar - tugma: bosilsa maydonga yoziladi
            def raqam_qoy(k):
                e_raqam.delete(0, "end")
                e_raqam.insert(0, k)
                e_raqam.focus_set()

            def sana_qoy(k):
                for w, v in ((e_kun, f"{k.day:02d}"), (e_oy, f"{k.month:02d}"), (e_yil, str(k.year))):
                    w.delete(0, "end")
                    w.insert(0, v)
                e_raqam.focus_set()
            vf = ttk.Frame(d)
            vf.grid(row=qator, column=0, columnspan=4, sticky="w", padx=10)
            ttk.Label(vf, text="Variantlar (bosing):", foreground="#C62828" if (r_shubha or s_shubha) else "#2E7D32"
                      ).pack(side="left", padx=(0, 6))
            for k, p in ish.ocr_variantlar["raqam"]:
                ttk.Button(vf, text=f"№{k} ({p:.0%})", command=lambda k=k: raqam_qoy(k)).pack(side="left", padx=2)
            ttk.Label(vf, text="   ").pack(side="left")
            for k, p in ish.ocr_variantlar["sana"]:
                ttk.Button(vf, text=f"{k:%d.%m.%Y} ({p:.0%})", command=lambda k=k: sana_qoy(k)).pack(side="left", padx=2)
            qator += 1
        if ish and ish.jadval_shubha:
            jr, js = ish.jadval_shubha
            ttk.Label(d, text=f"Jadvalda edi: №{jr or '?'}, {js.strftime('%d.%m.%Y') if js else '?'}  "
                      "(maydonlarga dalillar bo'yicha qiymat yozildi)", foreground="#777").grid(
                row=qator, column=0, columnspan=4, sticky="w", padx=10)
            qator += 1
        if ish and ish.ocr_izoh and ish.tekshir_kerak:
            ttk.Label(d, text="Diqqat: " + ish.ocr_izoh, foreground="#C62828", wraplength=760, justify="left").grid(
                row=qator, column=0, columnspan=4, sticky="w", padx=10, pady=(2, 0))
            qator += 1
        ttk.Separator(d).grid(row=qator, column=0, columnspan=4, sticky="ew", padx=10, pady=6)
        qator += 1
        boshqa = [("Javobgar", ish.javobgar if ish else ""), ("JShShIR", ish.jshshir if ish else ""),
                  ("Asosiy qarz", ish.asosiy_qarz if ish else ""), ("Pochta xarajati", ish.pochta if ish else ""),
                  ("Sud", ish.sud["nomi"] if ish and ish.sud else "")]
        kirishlar = {}
        for nom, qiymat in boshqa:
            ttk.Label(d, text=nom + ":", font=SHRIFT).grid(row=qator, column=0, sticky="e", padx=(10, 4), pady=3)
            e = ttk.Entry(d, width=48, font=SHRIFT)
            e.insert(0, qiymat or "")
            e.grid(row=qator, column=1, columnspan=3, sticky="w", pady=3)
            kirishlar[nom] = e
            qator += 1
        ttk.Label(d, text="(bular arizadan avtomatik olingan - odatda o'zgartirish shart emas)",
                  foreground="#777").grid(row=qator, column=0, columnspan=4, sticky="w", padx=10)
        qator += 1

        def oldinga(w, keyingi, uzunlik):
            """2 ta raqam yozilgach keyingi maydonga o'zi o'tadi."""
            def f(e):
                if e.char.isdigit() and len(re.sub(r"\D", "", w.get())) >= uzunlik:
                    keyingi.focus_set()
                    keyingi.select_range(0, "end")
            w.bind("<KeyRelease>", f)
        oldinga(e_kun, e_oy, 2)
        oldinga(e_oy, e_raqam, 2)

        def keyingisi():
            if navbat_qolgan is None:
                return
            # qo'shnisi tasdiqlangach OCR o'zi ishonch hosil qilganlari o'tkazib yuboriladi
            qolgan = [p for p in navbat_qolgan if p in ishlar_xotira and ishlar_xotira[p].tekshir_kerak]
            if qolgan:
                qator_ozgartir(papka=qolgan[0], navbat_qolgan=qolgan[1:])
            else:
                messagebox.showinfo("Tayyor", "Tekshiriladigan mijoz qolmadi.")

        def saqla(_e=None):
            kun, oy, yil = (re.sub(r"\D", "", w.get()) for w in (e_kun, e_oy, e_yil))
            raqam = re.sub(r"\D", "", e_raqam.get())
            sana = ""
            if kun or oy:
                try:
                    sana = date(int(yil), int(oy), int(kun)).strftime("%d.%m.%Y")
                except (ValueError, TypeError):
                    messagebox.showerror("Xato", "Sana noto'g'ri (kun, oy, yil).", parent=d)
                    return
            # dastur ishonch bilan boshqa qiymat kutgan bo'lsa - shoshib xato kiritilmadimi, so'raladi
            tr, ts, kuchli = tavsiya[:3]
            farq = []
            if kuchli and tr and raqam and raqam != tr:
                farq.append(f"raqam: siz №{raqam}, dastur №{tr}")
            if kuchli and ts and sana and sana != ts.strftime("%d.%m.%Y"):
                farq.append(f"sana: siz {sana}, dastur {ts:%d.%m.%Y}")
            if farq and not messagebox.askyesno(
                    "Tekshiring", "Dastur (OCR + paket qoidasi/qo'shni arizalar) boshqa qiymat kutgan:\n\n"
                    + "\n".join(farq) + "\n\nRasmga qarang. Siz kiritgan qiymat to'g'rimi?", parent=d):
                return
            q = {k: e.get().strip() for k, e in kirishlar.items()}
            avto = avto_sud(ish.manzil)[0] if ish else None
            if avto and q.get("Sud") == avto["nomi"]:
                q["Sud"] = ""          # o'zgartirilmagan - manzildan avtomatik aniqlanaveradi
            q["Ariza raqami"] = raqam
            q["Ariza sanasi"] = sana
            # odam ko'rib tasdiqladi; "ha!" - dalillarga zid bo'lsa ham shu qiymat (qayta so'ralmaydi)
            q["Tekshirildi"] = ("ha!" if farq else "ha") if raqam and sana else ""
            try:
                jadval_qator_saqlash(papka, q)
            except PermissionError:
                messagebox.showerror("Xato", "ishlar.xlsx Excel'da ochiq - yopib qayta urinib ko'ring.", parent=d)
                return
            if sana:
                holat["oxirgi_sana"] = datetime.strptime(sana, "%d.%m.%Y").date()
            d.destroy()
            bitta_yangilash(papka)
            keyingisi()

        def otkaz():
            d.destroy()
            keyingisi()

        f = ttk.Frame(d)
        f.grid(row=qator, column=0, columnspan=4, pady=12)
        ttk.Button(f, text="Saqlash" + (" va keyingisi" if navbat_qolgan else "") + "  (Enter)",
                   command=saqla).pack(side="left", padx=5)
        if navbat_qolgan is not None:
            ttk.Button(f, text="O'tkazib yuborish", command=otkaz).pack(side="left", padx=5)
        ttk.Button(f, text="Hujjatlar papkasi",
                   command=lambda: os.startfile(os.path.join(MIJOZLAR_PAPKA, papka))).pack(side="left", padx=5)
        ttk.Button(f, text="Yopish", command=d.destroy).pack(side="left", padx=5)
        d.bind("<Return>", saqla)
        d.bind("<Escape>", lambda _e: d.destroy())
        birinchi = e_raqam if (r_shubha and not s_shubha) else (e_kun if (s_shubha or not e_kun.get()) else e_raqam)
        birinchi.focus_set()
        birinchi.select_range(0, "end")

    def raqam_sana_kiritish():
        """Raqam/sanasini odam ko'rishi kerak bo'lgan mijozlarni (ro'yxat boshidagilarni) birin-ketin ochadi."""
        kerak = [p for p, i in ishlar_xotira.items() if i.tekshir_kerak]
        if not kerak:
            messagebox.showinfo("Tayyor", "Hamma mijozda ariza raqami va sanasi bor va tekshirilgan."
                                + ("\n(OCR hali ishlayapti - yangilari keyin chiqadi)" if holat["ocr"] else ""))
            return
        qator_ozgartir(papka=kerak[0], navbat_qolgan=kerak[1:])

    tv.bind("<Double-1>", qator_ozgartir)

    # ---------- odam kerak (CAPTCHA, tasdiq) ----------
    def banner_korsat(req):
        holat["odam"] = req
        for w in banner_tugmalar.winfo_children():
            w.destroy()
        x = req["xabar"].strip()
        banner_matn.config(text=x)
        if "YUBORILSINMI" in x:
            variantlar = [("HA, SUDGA YUBORISH", "ha", "#2E7D32"), ("Yo'q, yubormaslik", "", "#666")]
        elif "'yoq'" in x:
            variantlar = [("Davom etish", "", "#1565C0"), ("Bu mijozni o'tkazib yuborish", "yoq", "#C62828")]
        else:
            variantlar = [("Bajardim - davom etish", "", "#1565C0")]
        for matn, qiymat, rang in variantlar:
            tk.Button(banner_tugmalar, text=matn, bg=rang, fg="white", font=("Segoe UI", 11, "bold"), padx=14, pady=6,
                      relief="flat", command=lambda v=qiymat: banner_javob(v)).pack(side="left", padx=4)
        banner.pack(fill="x", padx=10, pady=6, after=tugmalar2)
        root.deiconify()
        root.lift()
        root.attributes("-topmost", True)
        root.after(1500, lambda: root.attributes("-topmost", False))
        root.bell()
        if avto:
            root.after(700, lambda: skrin("banner_" + str(int(time.time()))))
            if "CAPTCHA" not in x:
                root.after(1500, lambda: banner_javob("ha" if "YUBORILSINMI" in x else ""))

    def banner_javob(qiymat):
        req = holat["odam"]
        if req:
            req["javob"] = qiymat
            req["event"].set()
        banner_yashir(req)

    def banner_yashir(req):
        if holat["odam"] is req:
            holat["odam"] = None
            banner.pack_forget()

    def gui_odam(xabar, shart=None, vaqt=None):
        toxtat_tekshir()
        req = {"xabar": xabar, "event": threading.Event(), "javob": ""}
        print("\n    [SIZ KERAKSIZ] " + " | ".join(q.strip() for q in xabar.strip().splitlines()))
        uiq.put(("odam", req))
        signal()
        boshlandi = time.time()
        try:
            while not req["event"].wait(0.4):
                toxtat_tekshir()
                if shart:
                    try:
                        if shart():
                            print("    OK -> (skript o'zi sezdi, davom etilmoqda)")
                            return ""
                    except Exception:
                        pass
                if vaqt and time.time() - boshlandi > vaqt:
                    return None
            toxtat_tekshir()   # panel 'To'xtatish' bilan yopilgan bo'lsa - davom etmaymiz
            return req["javob"]
        finally:
            uiq.put(("odam_yop", req))

    def ui_chaqir(fn):
        """Ishchi oqimdan oyna oqimida funksiya bajarish (dialog va h.k.)."""
        javob, ev = [None], threading.Event()
        uiq.put(("chaqir", fn, javob, ev))
        ev.wait()
        return javob[0]

    def gui_parol():
        return ui_chaqir(lambda: simpledialog.askstring(
            "ONE ID paroli", f"ONE ID paroli ({S['login']}):\n(bir marta so'raladi, shifrlab saqlanadi)",
            show="*", parent=root)) or ""

    ODAM_ISHLOVCHI = gui_odam
    PAROL_SORASH = gui_parol

    # ---------- ishga tushirish ----------
    def holatni_qator(papka, teg, matn):
        for iid in tv.get_children():
            if tv.item(iid, "values")[1] == papka:
                v = list(tv.item(iid, "values"))
                v[8] = matn
                tv.item(iid, values=v, tags=(teg,))
                tv.see(iid)

    def tugmalar_holati(ishlayapti):
        holat["ishlayapti"] = ishlayapti
        band = ishlayapti or holat["ocr"]      # OCR fonda ishlayotganda ham boshlanmaydi
        for t in (t_boshla, t_test, t_test5):
            t.config(state="disabled" if band else "normal")
        for t in (t_yangila, t_sozlama):
            t.config(state="disabled" if ishlayapti else "normal")
        t_toxtat.config(state="normal" if ishlayapti else "disabled")

    def boshlash(rejim):
        if holat["ishlayapti"]:
            return
        sel = tanlanganlar()
        if rejim != "qadam5" and len(sel) < 2:   # bitta qator (ko'rish uchun bosilgan) tanlangan - baribir HAMMA mijoz
            sel = []
        qayta = qayta_var.get()
        ishlar = [i for i in ishlar_xotira.values() if (not sel or i.papka in sel)]
        oldingi = natijalar_oqish()
        tayyor = [i for i in ishlar if ishlansinmi(i, oldingi, rejim, qayta)]
        if not tayyor:
            messagebox.showinfo("Ishlanadigan mijoz yo'q",
                                "Tanlangan mijozlarda ishlanadigan ish yo'q.\n\n"
                                "- 'TO'LDIRING' bo'lsa: '✍ Raqam va sanani kiritish' tugmasini bosing.\n"
                                "- Avval bajarilganlarni qayta ishlash uchun 'Qayta ishlash' ni belgilang.")
            return
        kut = _son('xatoda_kutish', 0)
        captcha = ("CAPTCHA 'Robot emasman' ni dastur o'zi belgilaydi (misol chiqsa yoki ERI imzo - sariq panel va signal).\n"
                   + ("Biror qadam o'tmasa yoki sayt xato desa - SO'RAMASDAN shu mijoz XATO/TOPILMADI bo'ladi\n"
                      "(sababi natijaga yoziladi) va dastur o'zi keyingi mijozga o'tadi.\n\n" if kut <= 0 else
                      f"Biror qadam o'tmasa - {kut:.0f} soniya qo'lda tuzatish kutiladi, keyin keyingi mijoz.\n\n"))
        if rejim == "qadam5":
            tayyor = tayyor[:1]
            sel = [tayyor[0].papka]
            savol = (f"TEST 5-QADAMGACHA - mijoz: {tayyor[0].papka}\n\n"
                     f"1-5 qadamlar to'ldiriladi va hamma hujjatlar yuklanadi, keyin TO'XTAYDI.\n"
                     f"'Keyingi' bosilmaydi: murojaat SAQLANMAYDI va sudga YUBORILMAYDI.\n"
                     f"(Bir nechta mijoz tanlangan bo'lsa - faqat birinchisi.)\n\n" + captcha +
                     f"Testni boshlaymizmi?")
        elif rejim == "sinov":
            savol = (f"TEST: {len(tayyor)} ta mijoz ishlanadi - 1-6 qadamlar: saytda to'ldiriladi, hujjatlar yuklanadi\n"
                     f"va murojaat kabinetda SAQLANADI. Sudga YUBORILMAYDI (7-qadam bajarilmaydi).\n\n" + captcha +
                     f"Keyin '▶ BOSHLASH' bosilsa - shu saqlangan murojaat ochilib, sudga yuboriladi.\n\n"
                     f"Testni boshlaymizmi?")
        else:
            sorash = S.get("yuborishdan_oldin_sorash", "Yo'q").strip().lower() in ("ha", "yes", "1", "true")
            savol = (f"{len(tayyor)} ta mijoz ishlanadi: saytda to'ldiriladi, hujjatlar yuklanadi va SUDGA YUBORILADI.\n\n"
                     + captcha +
                     ("Har bir ariza yuborilishidan oldin 'HA, SUDGA YUBORISH' bilan tasdiqlaysiz.\n\n" if sorash else
                      "Har bir ariza SO'RALMASDAN yuboriladi va dastur o'zi keyingi mijozga o'tadi.\n\n")
                     + "Boshlaymizmi?")
        if not messagebox.askyesno("Boshlash" if rejim == "yuborish" else "Test", savol):
            return
        tugmalar_holati(True)
        progress.config(maximum=len(tayyor), value=0)
        holat_yozuv.config(text="Ishga tushmoqda...")

        def ishchi():
            natija = None
            try:
                natija = ishga_tushirish(rejim, tanlangan=set(sel) if sel else None, qayta=qayta,
                                         fon=bool(os.environ.get("ADOLAT_FON")))
            except Toxtatildi:   # brauzer ochilishidan oldin to'xtatildi
                print("\n>>> TO'XTATILDI")
                TOXTAT.clear()
                natija = []
            except Exception as e:
                print(f"\n!!! {type(e).__name__}: {e}")
                traceback.print_exc()
                natija = [("(umumiy)", "XATO", f"{type(e).__name__}: {e}")]
            uiq.put(("tugadi", natija))

        threading.Thread(target=ishchi, daemon=True).start()

    def toxtat():
        toxtatish()
        holat_yozuv.config(text="TO'XTATILDI - Chrome yopilmoqda...")
        if holat["odam"]:
            banner_javob("")

    def sozlamalar_oyna():
        sozlamalarni_qollash()
        d = tk.Toplevel(root)
        d.title("Sozlamalar")
        d.transient(root)
        d.grab_set()
        nomlar = [("login", "ONE ID login"), ("parol", "ONE ID paroli (o'zgartirish uchun yozing)"),
                  ("sayt", "Sayt"), ("profil_stir", "Profil STIR"), ("viloyat", "Viloyat"), ("sud", "Sud"),
                  ("asosiy_davogar", "Asosiy da'vogar"), ("qoshimcha_davogar_stir", "Qo'shimcha da'vogar STIR"),
                  ("davlat_boji_imtiyozi", "Davlat boji imtiyozi (matn qismi)"), ("kutish_soniya", "Kutish (soniya)"),
                  ("pauza_min", "Pauza min (soniya)"), ("pauza_max", "Pauza max (soniya)"),
                  ("harf_pauza", "Harf orasidagi pauza (soniya)"),
                  ("kechikish_koef", "Kutishlar koeffitsienti (1 = sekin, 0.5 = tez)"),
                  ("mijozlar_orasida", "Mijozlar orasida dam (soniya)"),
                  ("ocr_avto", "OCR ishonchli o'qisa - avtomatik (Ha/Yo'q)"), ("ocr_ishonch", "OCR ishonch chegarasi (0.90)"),
                  ("yuborishdan_oldin_sorash", "Har arizani yuborishdan oldin so'rash (Ha/Yo'q)"),
                  ("xatoda_kutish", "Xatoda kutish, soniya (0 = so'ramay keyingi mijoz)")]
        kirishlar = {}
        yarim = (len(nomlar) + 1) // 2     # ikki ustun - kichik ekranda ham "Saqlash" ko'rinsin
        for i, (k, nom) in enumerate(nomlar):
            r, c = i % yarim, 2 * (i // yarim)
            ttk.Label(d, text=nom + ":", font=SHRIFT).grid(row=r, column=c, sticky="w", padx=(10, 4), pady=3)
            e = ttk.Entry(d, width=34, font=SHRIFT, show="*" if k == "parol" else "")
            if k != "parol":
                e.insert(0, S.get(k, ""))
            e.grid(row=r, column=c + 1, padx=(0, 12), pady=3)
            kirishlar[k] = e
        tez_f = ttk.Frame(d)
        tez_f.grid(row=yarim, column=0, columnspan=4, sticky="w", padx=10, pady=(10, 2))
        ttk.Label(tez_f, text="TEZLIK:", font=("Segoe UI", 11, "bold")).pack(side="left")
        tezlik = tk.StringVar(value=S.get("tezlik") if S.get("tezlik") in TEZLIKLAR else "Qo'lda")
        for nom in list(TEZLIKLAR) + ["Qo'lda"]:
            ttk.Radiobutton(tez_f, text=nom, value=nom, variable=tezlik,
                            command=lambda: tezlik_tanlandi()).pack(side="left", padx=8)
        ttk.Label(tez_f, text="(sayt ulgurmasa - O'rtacha / Sekin)", foreground="#777").pack(side="left", padx=6)

        def tezlik_tanlandi():
            for k, v in TEZLIKLAR.get(tezlik.get(), {}).items():
                kirishlar[k].delete(0, "end")
                kirishlar[k].insert(0, v)

        hammaga = tk.BooleanVar(value=hammaga_toshkent())
        ttk.Checkbutton(d, variable=hammaga, style="Belgi.TCheckbutton",
                        text="HAMMA mijozlar - Toshkent shahar, Shayxontohur sudi\n"
                             "(belgilanmasa - da'vo arizada yozilgan sud, yozilmagan bo'lsa javobgar manzili)").grid(
            row=yarim + 1, column=0, columnspan=4, sticky="w", padx=10, pady=(8, 4))
        ttk.Label(d, text=("Parol saqlangan (shifrlangan)." if os.path.isfile(PAROL_FAYL)
                           else "Parol hali saqlanmagan - yozing yoki ishga tushganda so'raladi."),
                  foreground="#2E7D32" if os.path.isfile(PAROL_FAYL) else "#C62828").grid(
            row=yarim + 2, column=0, columnspan=4, sticky="w", padx=10)

        def saqla():
            q = {k: e.get() for k, e in kirishlar.items()}
            t = tezlik.get()          # pauzalar qo'lda o'zgartirilgan bo'lsa - "Qo'lda"
            q["tezlik"] = t if all(q[k].strip() == v for k, v in TEZLIKLAR.get(t, {}).items()) else "Qo'lda"
            if t not in TEZLIKLAR:
                q["tezlik"] = "Qo'lda"
            q["sud_hammaga_toshkent"] = "Ha" if hammaga.get() else "Yo'q"
            yangi_parol = q.pop("parol", "").strip()
            sozlamalarni_saqlash(q)
            if yangi_parol:
                parol_saqlash(yangi_parol)
            sozlamalarni_qollash()
            d.destroy()
            print("    OK -> Sozlamalar saqlandi")
            jadvalni_yuklash()      # OCR chegarasi / sud belgisi o'zgargan bo'lishi mumkin

        f = ttk.Frame(d)
        f.grid(row=yarim + 3, column=0, columnspan=4, pady=10)
        ttk.Button(f, text="Saqlash", command=saqla).pack(side="left", padx=5)
        ttk.Button(f, text="Bekor", command=d.destroy).pack(side="left", padx=5)

    def yangilash():
        try:
            jadval_yangilash()
        except PermissionError:
            messagebox.showerror("Xato", "ishlar.xlsx Excel'da ochiq - yopib qayta urinib ko'ring.")
        jadvalni_yuklash(keyin=ocr_fonda)       # yangi papkalar - fonda OCR

    def och(yol):
        if not os.path.splitext(yol)[1]:
            os.makedirs(yol, exist_ok=True)
        elif not os.path.exists(yol):
            messagebox.showinfo("Hali yo'q", f"{os.path.basename(yol)} hali yaratilmagan (birinchi ishdan keyin paydo bo'ladi).")
            return
        try:
            os.startfile(yol)
        except OSError as e:
            messagebox.showerror("Xato", str(e))

    t_boshla = ttk.Button(tugmalar, text="▶  BOSHLASH", style="Katta.TButton",
                          command=lambda: boshlash("yuborish"))
    t_test = ttk.Button(tugmalar, text="🧪  TEST (yubormasdan)", style="Katta.TButton",
                        command=lambda: boshlash("sinov"))
    t_test5 = ttk.Button(tugmalar, text="🧪  TEST: 5-qadamgacha", style="Katta.TButton",
                         command=lambda: boshlash("qadam5"))
    t_toxtat = ttk.Button(tugmalar, text="⏹  To'xtatish", style="Toxtat.TButton", command=toxtat, state="disabled")
    t_yangila = ttk.Button(tugmalar2, text="🔄 Yangilash (yangi papkalar)", command=yangilash)
    t_sozlama = ttk.Button(tugmalar2, text="⚙ Sozlamalar (login, parol)", command=sozlamalar_oyna)
    for t in (t_boshla, t_test, t_test5, t_toxtat):
        t.pack(side="left", padx=(0, 8), pady=6)
    qayta_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(tugmalar, text="Qayta ishlash", variable=qayta_var).pack(side="left", padx=6)
    ttk.Button(tugmalar2, text="✍ Raqam va sanani kiritish", command=raqam_sana_kiritish).pack(
        side="left", padx=(0, 6), pady=(0, 4))
    t_yangila.pack(side="left", padx=6, pady=(0, 4))
    t_sozlama.pack(side="left", padx=6, pady=(0, 4))
    for matn, buyruq in (("📁 Mijozlar papkasi", lambda: och(MIJOZLAR_PAPKA)),
                         ("📊 Natijalar", lambda: och(NATIJALAR)), ("📄 Loglar", lambda: och(LOG_PAPKA)),
                         ("❓ Qo'llanma", lambda: och(os.path.join(PAPKA, "QOLLANMA.txt")))):
        ttk.Button(tugmalar2, text=matn, command=buyruq).pack(side="left", padx=6, pady=(0, 4))

    # ---------- navbatni o'qish (oyna oqimi) ----------
    joriy = {"papka": None}

    def log_qosh(s):
        for qator in s.splitlines(keepends=True):
            teg = None
            if "!!!" in qator or "XATO" in qator or "Traceback" in qator:
                teg = "xato"
            elif "OK ->" in qator:
                teg = "ok"
            elif qator.startswith((">>>", "#", "\n>>>")) or ">>> " in qator:
                teg = "bosh"
            elif "[SIZ KERAKSIZ]" in qator:
                teg = "odam"
            log_oyna.insert("end", qator, teg)
            m = re.search(r"# \[(\d+)/(\d+)\] (.+)", qator)
            if m:
                progress.config(maximum=int(m.group(2)), value=int(m.group(1)) - 1)
                joriy["papka"] = m.group(3).strip()
                holatni_qator(joriy["papka"], "ISHLANMOQDA", "ishlanmoqda...")
                holat_yozuv.config(text=f"Mijoz {m.group(1)}/{m.group(2)}: {joriy['papka']}")
            m = re.search(r">>> (\d)-qadam: (.+)", qator)
            if m and joriy["papka"]:
                holat_yozuv.config(text=f"{holat_yozuv.cget('text').split('  |')[0]}  |  {m.group(1)}/7 qadam: {m.group(2)}")
            m = re.search(r">>> NATIJA: (.+?) = (\w+)(?: \[([^\]]+)\])?", qator)
            if m:
                holatni_qator(m.group(1), m.group(2), m.group(3) or holat_matni(m.group(2)))
                progress.step(1)
        log_oyna.see("end")

    def navbat():
        try:
            while True:
                x = uiq.get_nowait()
                tur = x[0]
                if tur == "log":
                    log_qosh(x[1])
                elif tur == "odam":
                    banner_korsat(x[1])
                elif tur == "odam_yop":
                    banner_yashir(x[1])
                elif tur == "ishlar":
                    holat["yuklanmoqda"] = False
                    if x[1] is not None:
                        jadvalni_chizish(x[1])
                    if holat["qayta_yuklash"]:
                        holat["qayta_yuklash"] = False
                        jadvalni_yuklash(keyin=x[2])
                    elif x[2]:
                        x[2]()
                elif tur == "ocr_holat":
                    _t, n, jami, papka = x
                    ocr_yozuv.config(text=f"OCR: {n}/{jami}  {papka[:45]}")
                    progress.config(maximum=jami, value=n - 1)
                elif tur == "ocr_jadval":
                    jadvalni_yuklash()
                elif tur == "ocr_tugadi":
                    holat["ocr"] = False
                    ocr_yozuv.config(text="")
                    progress.config(value=0)
                    tugmalar_holati(holat["ishlayapti"])
                    if x[1]:
                        print(f"OK -> OCR tugadi: {x[1]} ta ariza o'qildi")
                        jadvalni_yuklash(keyin=ocr_fonda)    # shu orada yangi papka qo'shilgan bo'lsa
                elif tur == "chaqir":
                    try:
                        x[2][0] = x[1]()
                    finally:
                        x[3].set()
                elif tur == "tugadi":
                    tugmalar_holati(False)
                    joriy["papka"] = None
                    natija = x[1] or []
                    sanoq = {}
                    for _p, h, _i in natija:
                        sanoq[h] = sanoq.get(h, 0) + 1
                    matn = ", ".join(f"{k}: {v}" for k, v in sanoq.items()) or "hech narsa bajarilmadi"
                    holat_yozuv.config(text=f"Tugadi. {matn}")
                    jadvalni_yuklash(keyin=lambda m=matn: holat_yozuv.config(text=f"Tugadi. {m}"))
                    if any(SAYT_BAND_XABAR in iz for _p, _h, iz in natija):
                        holat_yozuv.config(text="Sayt hozir ishlamayapti - 30 daqiqa / 1 soatdan keyin qayta urinib ko'ring")
                        messagebox.showwarning("Sayt hozir ishlamayapti", SAYT_BAND_XABAR)
                    else:
                        messagebox.showinfo("Tugadi", f"Ish tugadi.\n\n{matn}\n\nTafsilot: jadvalda va 'Natijalar' faylida.")
        except queue.Empty:
            pass
        root.after(150, navbat)

    def yopish():
        if holat["ishlayapti"] and not messagebox.askyesno("Chiqish", "Jarayon ketyapti. Baribir chiqasizmi?"):
            return
        root.destroy()
        os._exit(0)

    root.protocol("WM_DELETE_WINDOW", yopish)
    print(f"ADOLAT AVTOMAT {VERSIYA}\nPapka: {PAPKA}\nLog: {log}\n")
    jadvalni_yuklash(keyin=ocr_fonda)     # ro'yxat darhol, o'qilmagan arizalar - fonda OCR
    root.after(150, navbat)
    if avto == "korinish":
        root.after(2500, lambda: (skrin("asosiy"), os._exit(0)))
    elif avto == "dialog":
        root.after(2000, raqam_sana_kiritish)
        def dialog_skrin():
            try:
                from PIL import ImageGrab
                for w in root.winfo_children():
                    if isinstance(w, tk.Toplevel):
                        w.attributes("-topmost", True)
                        w.lift()
                        w.update()
                        time.sleep(0.5)
                        x, y = w.winfo_rootx(), w.winfo_rooty()
                        ImageGrab.grab((x, y, x + w.winfo_width(), y + w.winfo_height())).save(
                            os.path.join(os.environ["ADOLAT_SKRIN"], "dialog.png"))
            except Exception as e:
                print(f"(dialog skrin xatosi: {e})")
            os._exit(0)
        root.after(4500, dialog_skrin)
    elif avto in ("sinov", "qadam5", "yuborish"):
        root.after(1500, lambda: boshlash(avto))
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) == 1 or "--oyna" in sys.argv:
        gui_main()
    else:
        main()
