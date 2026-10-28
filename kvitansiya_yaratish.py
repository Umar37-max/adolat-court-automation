"""
KVITANSIYA YARATISH - billing.sud.uz da har mijozga pochta xarajati kvitansiyasi (alohida dastur).

Mijoz papkasidagi da'vo arizadan (Word / PDF / shartnoma) javobgar manzili -> sud aniqlanadi (AdolatAvtomat
bilan bir xil qoidalar; sozlamalar.ini dagi "HAMMA mijozlar - Toshkent shahar, Shayxontohur" belgisi ham,
ishlar.xlsx dagi qo'lda yozilgan "Sud" ham). Papkadagi kvitansiya shu sudga, summasi va muddati mos bo'lsa -
yangisi yaratilmaydi. Yaratilgan kvitansiya PDF i mijoz papkasiga <raqam>.pdf bo'lib saqlanadi -
AdolatAvtomat uni papkadan olib 4-qadamda biriktiradi (sudi mos kelganini).

"Robot emasman" dan keyin chiqadigan MATEMATIK MISOLNI dastur yechmaydi - odam yechadi (sariq panel + signal),
dastur 'Yaratish' ni o'zi bosadi.

Butunlay ALOHIDA dastur: o'z papkasi (KvitansiyaYaratish\\): mijozlar\\ (yoki "Papkani tanlash"),
kvitansiya_sozlamalar.ini, kvitansiyalar.csv, loglar\\. AdolatAvtomat papkalariga tegmaydi.
Ishga tushirish: KvitansiyaYaratish.exe
yoki: python kvitansiya_yaratish.py   (--tekshir - faqat ro'yxat, brauzersiz)
"""
import csv
import os
import queue
import re
import sys
import threading
import traceback
from datetime import date, datetime

import adolat_avtomat as A
from selenium.common.exceptions import WebDriverException

VERSIYA = "2026-10-13-v10"

A.STANDART_SOZLAMALAR.update({
    "kvitansiya_sayt": "https://billing.sud.uz",
    "kvitansiya_summa": "44000",
    "kvitansiya_tolov_turi": "Почта харажатлари",
    "kvitansiya_sud_turi": "Фуқаролик ишлари бўйича суд",
    "kvitansiya_tashkilot": os.getenv("KVITANSIYA_TASHKILOT", ""),   # to'lovchi tashkilot nomi
    "kvitansiya_stir": os.getenv("KVITANSIYA_STIR", ""),             # to'lovchi STIRi
    "kvitansiya_viloyat": "Тошкент шаҳар",
    "kvitansiya_tuman": "Шайхантохур тумани",
    "kvitansiya_manzil": os.getenv("KVITANSIYA_MANZIL", ""),         # to'lovchi manzili (ko'cha, uy)
})
# ---------- O'Z papkasi: AdolatAvtomat dan butunlay alohida (o'z mijozlari, sozlamalari, jurnali, loglari) ----------
if getattr(sys, "frozen", False):
    KV_PAPKA = os.path.dirname(os.path.abspath(sys.executable))          # KvitansiyaYaratish.exe yonida
else:
    KV_PAPKA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "KvitansiyaYaratish")
os.makedirs(KV_PAPKA, exist_ok=True)
A.PAPKA = KV_PAPKA
A.SOZLAMALAR_FAYL = os.path.join(KV_PAPKA, "kvitansiya_sozlamalar.ini")
A.LOG_PAPKA = os.path.join(KV_PAPKA, "loglar")
A.JADVAL = os.path.join(KV_PAPKA, "_jadval_yoq.xlsx")                    # AdolatAvtomat jadvali ishlatilmaydi
A.UMUMIY_PAPKA = os.path.join(KV_PAPKA, "_umumiy_yoq")
A.SKAN_KESH_FAYL = os.path.join(KV_PAPKA, "ocr_skan_kesh.json")
A.QOLYOZMA_KESH_FAYL = os.path.join(KV_PAPKA, "ocr_qolyozma_kesh.json")
if os.path.isfile(os.path.join(KV_PAPKA, "Fuqarolik_sudlari_yurisdiksiya.xlsx")):
    A.SUDLAR_FAYL = os.path.join(KV_PAPKA, "Fuqarolik_sudlari_yurisdiksiya.xlsx")
# AdolatAvtomat2 (buyruq) mijozlari ham: asosiy ariza nomida "даъво" emas "суд буйруғи" bo'ladi - aks holda
# ariza topilmaydi, sud ("...sudiga") aniqlanmaydi va kvitansiya mijozning o'z sudiga yaratilmaydi
A.HUJJAT_TURLARI[0][2].extend(["буйру", "буюру", "buyru", "buyuru"])
KIRISH_PAPKA = os.path.join(KV_PAPKA, "KVITANSIYALAR")            # mijoz papkalari shu yerga tashlanadi
TAYYOR_PAPKA = os.path.join(KV_PAPKA, "TAYYOR KVITANSIYALAR")     # hamma tayyor kvitansiya bir joyda
# o'z sozlamalari - faqat kerakligi (AdolatAvtomat ning login/da'vogar sozlamalari bu yerda yo'q)
A.STANDART_SOZLAMALAR = {k: v for k, v in A.STANDART_SOZLAMALAR.items()
                         if k.startswith("kvitansiya_") or k in ("sud", "sud_hammaga_toshkent", "tezlik", "kutish_soniya", "pauza_min", "pauza_max",
                                  "harf_pauza", "kechikish_koef")}
A.STANDART_SOZLAMALAR["kvitansiya_mijozlar"] = KIRISH_PAPKA
A.SOZLAMA_IZOHLARI = """\
; KVITANSIYA YARATISH sozlamalari (dastur oynasidan ham o'zgartiriladi)
; kvitansiya_mijozlar     - mijoz papkalari turgan papka (standart: KvitansiyaYaratish\\KVITANSIYALAR)
; sud_hammaga_toshkent    - Ha: hamma mijozga Toshkent shahar / "sud" dagi sud (Shayxontohur); Yo'q: da'vo arizadagi sud / manzil
; kvitansiya_summa        - summa (so'm); kvitansiya_tashkilot / _stir / _viloyat / _tuman / _manzil - to'lovchi
"""
A.CHROME_PORT = 9242     # AdolatAvtomat bilan bir vaqtda ishlasa ham Chrome'lari aralashmasin
A.CHROME_PROFIL = os.path.join(os.environ.get("LOCALAPPDATA") or KV_PAPKA, "AdolatAvtomat", "chrome_kvitansiya")
A.SESSIYA = "kvitansiya_" + A.SESSIYA
JURNAL = os.path.join(KV_PAPKA, "kvitansiyalar.csv")   # yaratilgan kvitansiyalar (PDF yuklanmasa ham raqami)


def sozlamalar():
    """Sozlamalarni o'qiydi va mijozlar papkasini (sozlamalardagi yoki KvitansiyaYaratish\\mijozlar) qo'llaydi."""
    import configparser
    cp = configparser.ConfigParser(interpolation=None)
    cp.read(A.SOZLAMALAR_FAYL, encoding="utf-8-sig")
    if cp.has_section("ADOLAT") and any(k not in A.STANDART_SOZLAMALAR for k in cp["ADOLAT"]):
        toza = configparser.ConfigParser(interpolation=None)      # eski versiya yozgan begona kalitlar olib tashlanadi
        toza.add_section("ADOLAT")
        for k in A.STANDART_SOZLAMALAR:
            if cp.has_option("ADOLAT", k):
                toza.set("ADOLAT", k, cp.get("ADOLAT", k))
        with open(A.SOZLAMALAR_FAYL, "w", encoding="utf-8") as f:
            f.write(A.SOZLAMA_IZOHLARI)
            toza.write(f)
    A.sozlamalarni_qollash()
    eski = os.path.join(KV_PAPKA, "mijozlar")             # oldingi versiyadagi nomi
    if os.path.normcase(A.S.get("kvitansiya_mijozlar", "")) == os.path.normcase(eski):
        if os.path.isdir(eski) and not os.path.exists(KIRISH_PAPKA):
            os.rename(eski, KIRISH_PAPKA)
        A.sozlamalarni_saqlash({"kvitansiya_mijozlar": KIRISH_PAPKA})
        A.sozlamalarni_qollash()
    A.MIJOZLAR_PAPKA = A.S.get("kvitansiya_mijozlar") or KIRISH_PAPKA
    if not os.path.isdir(os.path.dirname(A.MIJOZLAR_PAPKA.rstrip("\\/"))):   # boshqa kompyuterdagi yo'l
        A.MIJOZLAR_PAPKA = KIRISH_PAPKA
        A.sozlamalarni_saqlash({"kvitansiya_mijozlar": KIRISH_PAPKA})
    os.makedirs(A.MIJOZLAR_PAPKA, exist_ok=True)
    os.makedirs(TAYYOR_PAPKA, exist_ok=True)


def tayyorga(ish, raqam, yol):
    """Kvitansiya nusxasi - TAYYOR KVITANSIYALAR\\<mijoz> - <raqam>.pdf (hammasi bir joyda)."""
    import shutil
    maqsad = os.path.join(TAYYOR_PAPKA, f"{A.xavfsiz_nom(ish.papka)[:80]} - {raqam}.pdf")
    if os.path.isfile(yol) and not os.path.isfile(maqsad):
        shutil.copy2(yol, maqsad)

A.JS += r"""
function b_maydon(label) {
  // billing.sud.uz: mat-form-field ANIQ yozuvi bo'yicha ('Sud' - 'Sud turi' / 'Sud hududi' emas; '*' hisobga olinmaydi)
  const t = norm(label);
  for (const l of document.querySelectorAll('mat-label')) {
    if (norm((l.innerText || '').replace(/\*/g, '')) !== t) continue;
    const f = l.closest('mat-form-field');
    if (f && vis(f)) return f;
  }
  return null;
}
function b_variantlar(f) {
  // shu maydon OCHGAN ro'yxat variantlari (yopilgan eski ro'yxatlar sahifada qolib ketadi)
  const s = f && f.querySelector('mat-select');
  const id = s && s.getAttribute('aria-controls');
  const p = id && document.getElementById(id);
  return p ? [...p.querySelectorAll('mat-option')] : [];
}
function b_tugma(t) {
  const d = document.querySelector('mat-dialog-container');
  for (const x of (d || document).querySelectorAll('button'))
    if (vis(x) && !x.disabled && txt(x).includes(t)) return x;
  return null;
}
"""


class OtkazibYuborildi(Exception):
    """Odam misolni yechmay "O'tkazib yuborish" ni bosdi."""


def summa():
    return A._son("kvitansiya_summa", 44000)


def mosmi(k, ish):
    """Papkadagi kvitansiya shu ishga yaraydimi: sudi (turi bilan), summasi, muddati, maqsadi (pochta), to'lovchi STIR i."""
    stir = re.sub(r"\D", "", A.S.get("kvitansiya_stir", ""))
    return bool(k and k["pochta"] and A.kvitansiya_sudi_mos(k, ish.sud) and abs(k["summa"] - summa()) < 1
                and (k["muddat"] is None or k["muddat"] >= date.today())
                and (not stir or not k.get("tolovchi_stir") or k["tolovchi_stir"] == stir))


def papkadagi(ish):
    """Papkadagi mos kvitansiya raqami yoki ''."""
    mos = []
    for k, _n, y in ish.hujjatlar:
        if k == "tolov" and y.lower().endswith(".pdf"):
            m = A.kvitansiya_oqish(y)
            if mosmi(m, ish):
                mos.append((os.path.getmtime(y), m["raqam"]))
    return max(mos)[1] if mos else ""


JURNAL_USTUNLAR = ["vaqt", "papka", "raqam", "sud", "summa", "stir"]


def _jurnal():
    if not os.path.isfile(JURNAL):
        return []
    with open(JURNAL, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


def jurnalga(ish, raqam):
    qatorlar = _jurnal()
    eski = qatorlar and "stir" not in qatorlar[0]          # eski jurnal (stir ustunisiz) - yangi ustun qo'shiladi
    with open(JURNAL, "w" if eski else "a", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=JURNAL_USTUNLAR, delimiter=";", extrasaction="ignore")
        if eski or not qatorlar:
            w.writeheader()
        for r in qatorlar if eski else []:
            w.writerow(r)
        w.writerow({"vaqt": datetime.now().strftime("%d.%m.%Y %H:%M:%S"), "papka": ish.papka, "raqam": raqam,
                    "sud": ish.sud["nomi"], "summa": f"{summa():.0f}", "stir": A.S.get("kvitansiya_stir", "")})


def jurnaldagi(ish):
    """Shu mijozga yaqinda (25 kun) shu sud, summa va to'lovchi STIR iga yaratilgan, lekin PDF i papkada yo'q
    kvitansiya raqami (qayta yaratmaslik uchun)."""
    if not ish.sud:
        return ""
    for r in reversed(_jurnal()):
        try:
            yangi = (date.today() - datetime.strptime(r["vaqt"][:10], "%d.%m.%Y").date()).days <= 25
        except (KeyError, ValueError):
            continue
        if (yangi and r.get("papka") == ish.papka and r.get("sud") == ish.sud["nomi"]
                and abs(float(r.get("summa") or 0) - summa()) < 1 and (r.get("stir") or "") == A.S.get("kvitansiya_stir", "")):
            return r["raqam"]
    return ""


def holati(ish):
    """('bor', raqam) | ('pdf', raqam) - yaratilgan, PDF yuklanmagan | ('kerak', '') | ('sudsiz', '')."""
    if not ish.sud:
        return "sudsiz", ""
    r = papkadagi(ish)
    if r:
        return "bor", r
    r = jurnaldagi(ish)
    return ("pdf", r) if r else ("kerak", "")


def mijozlar():
    """Mijozlar (AdolatAvtomat qoidalari bilan): manzil -> sud, papkadagi hujjatlar. OCR ishlatilmaydi."""
    A.ziplarni_ochish()
    A.ichki_papkalarni_chiqarish()
    jadval = A.jadval_oqish()
    umumiy = A.papka_hujjatlari(A.UMUMIY_PAPKA)
    return [A.ish_tayyorlash(p, jadval.get(p), umumiy, ocr=False) for p in A.mijoz_papkalari()]


# ---------------- billing.sud.uz ----------------
def _matn_balli(kerak):
    return lambda v: 2 if A.norm_py(v) == A.norm_py(kerak) else (1 if A.norm_py(kerak) in A.norm_py(v) else 0)


def _maydon(b, yorliq, vaqt=15):
    f = b.kut(lambda: b.js("return b_maydon(arguments[0]);", yorliq), vaqt)
    if not f:
        b.tahlil(f"billing_{yorliq}")
        raise A.Topilmadi(f"billing.sud.uz: '{yorliq}' maydoni")
    return f


def _yoz(b, yorliq, qiymat):
    b.js_yoz(b.js("return arguments[0].querySelector('input,textarea');", _maydon(b, yorliq)), qiymat)
    print(f"    OK -> {yorliq}: {qiymat}")


def _tanla(b, yorliq, balli, nom=None):
    """Ro'yxatni (mat-select) ochib, balli(matn) eng kattasini tanlaydi va tanlanganini tekshiradi."""
    nom = nom or yorliq
    variantlar = []
    for _ in range(3):
        b.bos_el(b.js("return arguments[0].querySelector('mat-select') || arguments[0];", _maydon(b, yorliq)))
        variantlar = b.kut(lambda: b.js(
            "return b_variantlar(b_maydon(arguments[0])).map(o => o.innerText.trim());", yorliq), 10) or []
        baho = sorted(((balli(v), v) for v in variantlar), reverse=True)
        if baho and baho[0][0] > 0:
            eng = baho[0][1]
            el = b.js("return b_variantlar(b_maydon(arguments[0])).find(o => o.innerText.trim() === arguments[1]) || null;",
                      yorliq, eng)
            if el:
                b.bos_el(el)
                if b.kut(lambda: A.norm_py(eng) in A.norm_py(
                        b.js("return (b_maydon(arguments[0]) || {}).innerText || '';", yorliq)), 5):
                    print(f"    OK -> {nom}: {eng}")
                    return eng
        b.esc()
        A.tin(1)
    b.tahlil(f"billing_tanlanmadi_{yorliq}")
    raise A.Topilmadi(f"billing.sud.uz: '{nom}' tanlanmadi (variantlar: {variantlar[:12]})")


def pdf_yuklash(b, ish, raqam):
    """Kvitansiya PDF ini sayt API sidan (shu brauzer sessiyasi bilan) mijoz papkasiga <raqam>.pdf qilib saqlaydi."""
    import base64
    sayt = (A.S.get("kvitansiya_sayt") or "https://billing.sud.uz").rstrip("/")
    if not (b.d.current_url or "").startswith(sayt):
        b.d.get(sayt + "/")
    b.d.set_script_timeout(60)
    xato = ""
    for _ in range(3):
        b64 = b.d.execute_async_script("""
            const done = arguments[arguments.length - 1];
            fetch(arguments[0], {credentials: 'include'}).then(r => r.arrayBuffer()).then(buf => {
                const a = new Uint8Array(buf); let s = '';
                for (let i = 0; i < a.length; i += 0x8000) s += String.fromCharCode.apply(null, a.subarray(i, i + 0x8000));
                done(btoa(s));
            }).catch(e => done('XATO:' + e));""", f"{sayt}/api/invoice/asDocument?invoice={raqam}")
        data = base64.b64decode(b64) if b64 and not b64.startswith("XATO:") else b""
        if data.startswith(b"%PDF"):
            yol = os.path.join(ish.yol, f"{raqam}.pdf")
            with open(yol, "wb") as f:
                f.write(data)
            k = A.kvitansiya_oqish(yol)
            if k and not mosmi(k, ish):
                raise A.Topilmadi(f"kvitansiya {raqam} sudi/summasi mos emas: {k['sud']}, {k['summa']:.0f}")
            ish.hujjatlar.append(("tolov", "To'lov kvitansiyasi", yol))
            tayyorga(ish, raqam, yol)
            print(f"    OK -> PDF saqlandi: {os.path.relpath(yol, A.MIJOZLAR_PAPKA)}  (+ TAYYOR KVITANSIYALAR)")
            return yol
        xato = (b64 or "")[:120]
        A.uxla(3)
    raise A.Topilmadi(f"kvitansiya {raqam} PDF i yuklanmadi ({xato}) - qayta ishga tushiring (yangisi yaratilmaydi)")


def yaratish(b, ish):
    """billing.sud.uz da shu ish SUDIGA pochta xarajati kvitansiyasi; PDF mijoz papkasiga. Raqamini qaytaradi."""
    sayt = (A.S.get("kvitansiya_sayt") or "https://billing.sud.uz").rstrip("/")
    b.d.get(sayt + "/create-receipt")
    b.bos("Yuridik shaxs", "Yuridik shaxs", vaqt=A.KUTISH)
    _yoz(b, "Tashkilot nomi", A.S["kvitansiya_tashkilot"])
    _yoz(b, "STIR", A.S["kvitansiya_stir"])
    b.bos_el(_maydon(b, "Tashkilot manzili"))
    _maydon(b, "Viloyat", 10)
    _tanla(b, "Viloyat", _matn_balli(A.S["kvitansiya_viloyat"]), "To'lovchi viloyati")
    _tanla(b, "Tuman", _matn_balli(A.S["kvitansiya_tuman"]), "To'lovchi tumani")
    _yoz(b, "Manzil", A.S["kvitansiya_manzil"])
    saqla = b.kut(lambda: b.js("return b_tugma('saqlash');"), 10)
    if not saqla:
        b.tahlil("billing_manzil")
        raise A.Topilmadi("billing.sud.uz: manzil oynasidagi 'Saqlash'")
    b.bos_el(saqla)
    b.kut(lambda: not b.js("return !!document.querySelector('mat-dialog-container');"), 10)
    _tanla(b, "Sud turi", _matn_balli(A.S["kvitansiya_sud_turi"]))
    _tanla(b, "Sud hududi", lambda v: A.viloyat_variant_balli(v, ish.sud["hudud"]))
    _tanla(b, "Sud", lambda v: A.sud_variant_balli(v, ish.sud))
    _tanla(b, "To'lov turi", _matn_balli(A.S["kvitansiya_tolov_turi"]))
    _yoz(b, "Kvitansiya summasi", f"{summa():.0f}")
    b.tahlil("billing_toldirildi")

    tugma = lambda: b.js("return b_tugma('yaratish');")
    b.js(A.ROBOT_BOS_JS)
    if not b.kut(tugma, 3):
        j = A.odam(f"{ish.papka}:\nbrauzerda MATEMATIK MISOLNI yeching va 'Yuborish' ni bosing - 'Yaratish' ni dastur o'zi bosadi.",
                   shart=tugma, cheksiz=True)
        if j == "otkaz":
            raise OtkazibYuborildi()
    t = tugma()
    if not t:
        b.tahlil("billing_yaratish_yoqilmadi")
        raise A.Topilmadi("billing.sud.uz: 'Yaratish' yoqilmadi (misol yechilmadi)")
    b.bos_el(t)
    if not b.kut(lambda: re.search(r"/invoice/(\d+)", b.d.current_url or ""), A.KUTISH):
        b.tahlil("billing_invoice_yoq")
        raise A.Topilmadi(f"billing.sud.uz: kvitansiya yaratilmadi. Xabarlar: {b.xatolar()}")
    raqam = re.search(r"/invoice/(\d+)", b.d.current_url).group(1)
    print(f"    OK -> Kvitansiya yaratildi: {raqam} ({ish.sud['nomi']}, {summa():.0f} so'm)")
    jurnalga(ish, raqam)
    pdf_yuklash(b, ish, raqam)
    return raqam


def sozlamalar_matni():
    return (f"tezlik={A.S.get('tezlik')} (pauza {A.S.get('pauza_min')}-{A.S.get('pauza_max')} s, "
            f"koef {A.S.get('kechikish_koef')}), summa={A.S.get('kvitansiya_summa')}, STIR={A.S.get('kvitansiya_stir')}, "
            f"hamma Toshkent={A.S.get('sud_hammaga_toshkent')}")


def jarayon(ishlar, xabar=print):
    """Kvitansiyasi yo'q / mos emas mijozlarga kvitansiya. Natija: [(papka, holat, izoh)]."""
    A.TOXTAT.clear()
    for i in ishlar:                         # bori ham TAYYOR KVITANSIYALAR ga yig'iladi
        h, r = holati(i)
        if h == "bor":
            y = next((y for k, _n, y in i.hujjatlar if k == "tolov" and r in os.path.basename(y)), None)
            if y:
                tayyorga(i, r, y)
    navbat = [(i, *holati(i)) for i in ishlar]
    navbat = [(i, h, r) for i, h, r in navbat if h in ("kerak", "pdf")]
    natija = []
    if not navbat:
        print("Hamma mijozda mos kvitansiya bor - yaratiladigani yo'q.")
        return natija
    print(f"\n>>> {len(navbat)} ta mijozga kvitansiya ({summa():.0f} so'm, o'z sudi bo'yicha)")
    print(f"    Sozlamalar: {sozlamalar_matni()}")
    b = None
    try:
        b = A.Brauzer()
        for n, (ish, h, raqam) in enumerate(navbat, 1):
            A.toxtat_tekshir()
            print(f"\n{'#' * 60}\n# [{n}/{len(navbat)}] {ish.papka} -> {ish.sud['hudud']} / {ish.sud['nomi']}\n{'#' * 60}")
            xabar(ish.papka, "yaratilmoqda...", "ISHLANMOQDA")
            for urinish in (1, 2):
                try:
                    if h == "pdf":                 # avval yaratilgan, faqat PDF i yo'q - yangisi yaratilmaydi
                        pdf_yuklash(b, ish, raqam)
                    else:
                        raqam = yaratish(b, ish)
                    natija.append((ish.papka, "YARATILDI", raqam))
                    xabar(ish.papka, f"YARATILDI {raqam}", "OK")
                    break
                except OtkazibYuborildi:
                    natija.append((ish.papka, "O'TKAZILDI", "misol yechilmadi"))
                    xabar(ish.papka, "o'tkazib yuborildi", "XATO")
                    break
                except (A.Topilmadi, WebDriverException) as e:
                    A.toxtat_tekshir()             # To'xtatish bosilgan bo'lsa - qayta urinilmaydi
                    izoh = f"{type(e).__name__}: {(str(e).strip().splitlines() or [''])[0]}"
                    print(f"    !!! {izoh}")
                    raqam_bor = jurnaldagi(ish)
                    if urinish == 2 or raqam_bor:       # yaratilgan bo'lsa - qayta yaratilmaydi
                        natija.append((ish.papka, "XATO", izoh))
                        xabar(ish.papka, "XATO: " + izoh[:80], "XATO")
                        break
                    if not A.brauzer_tirikmi(b):
                        b.yop()
                        b = A.Brauzer()
                    A.uxla(3)
    except A.Toxtatildi:
        print("\n>>> TO'XTATILDI - Chrome yopildi")
    except Exception:
        if not A.TOXTAT.is_set():
            raise
        print("\n>>> TO'XTATILDI - Chrome yopildi")
    finally:
        A.TOXTAT.clear()
        if b:
            b.yop()
    ok = sum(1 for _p, h, _i in natija if h == "YARATILDI")
    print(f"\n{'=' * 60}\nYAKUN: {ok} ta kvitansiya yaratildi, {len(natija) - ok} ta yaratilmadi. Jurnal: {JURNAL}\n{'=' * 60}")
    return natija


# ---------------- oyna ----------------
def gui():
    import tkinter as tk
    from tkinter import messagebox, scrolledtext, ttk
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    uiq = queue.Queue()

    class Yozuvchi:
        def write(self, s):
            if s:
                uiq.put(("log", s))

        def flush(self):
            pass

    log = A.log_boshlash(Yozuvchi())
    sozlamalar()
    root = tk.Tk()
    root.title("KVITANSIYA YARATISH - billing.sud.uz (pochta xarajati)")
    root.geometry("1400x800")
    st = ttk.Style()
    try:
        st.theme_use("vista")
    except tk.TclError:
        pass
    SHRIFT = ("Segoe UI", 10)
    st.configure("Treeview", font=SHRIFT, rowheight=25)
    st.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    st.configure("Katta.TButton", font=("Segoe UI", 11, "bold"), padding=8)
    st.configure("Toxtat.TButton", font=("Segoe UI", 11, "bold"), padding=8, foreground="#C62828")
    st.configure("Belgi.TCheckbutton", font=("Segoe UI", 10, "bold"))
    holat = {"ishlayapti": False, "odam": None}

    tepa = ttk.Frame(root, padding=(10, 8))
    tepa.pack(fill="x")
    ttk.Label(tepa, text="KVITANSIYA YARATISH", font=("Segoe UI", 15, "bold")).pack(side="left")
    ttk.Label(tepa, text=f"   versiya {VERSIYA}", foreground="#777").pack(side="left")
    papka_yozuv = ttk.Label(root, text="", font=("Segoe UI", 10), padding=(10, 0))
    papka_yozuv.pack(fill="x")

    soz = ttk.LabelFrame(root, text=" Sozlamalar ", padding=(10, 6))
    soz.pack(fill="x", padx=10)
    hammaga = tk.BooleanVar(value=A.hammaga_toshkent())
    ttk.Checkbutton(soz, variable=hammaga, style="Belgi.TCheckbutton",
                    text="HAMMA mijozlar - Toshkent shahar, Shayxontohur sudi (belgilanmasa - da'vo arizada yozilgan sud, bo'lmasa manzil)"
                    ).grid(row=0, column=0, columnspan=8, sticky="w", pady=(0, 6))
    maydonlar = [("kvitansiya_summa", "Summa (so'm)", 10), ("kvitansiya_tashkilot", "To'lovchi", 24),
                 ("kvitansiya_stir", "STIR", 12), ("kvitansiya_viloyat", "Viloyat", 16),
                 ("kvitansiya_tuman", "Tuman", 20), ("kvitansiya_manzil", "Manzil", 22)]
    kirish = {}
    for i, (k, nom, en) in enumerate(maydonlar):
        r, c = 1 + i // 3, 2 * (i % 3)
        ttk.Label(soz, text=nom + ":").grid(row=r, column=c, sticky="e", padx=(8, 4), pady=2)
        e = ttk.Entry(soz, width=en + 8, font=SHRIFT)
        e.insert(0, A.S.get(k, ""))
        e.grid(row=r, column=c + 1, sticky="w", pady=2)
        kirish[k] = e

    def saqla(jim=False):
        """Oynadagi sozlamalarni faylga yozadi va DARHOL qo'llaydi (jim=True - ro'yxatni qayta o'qimasdan)."""
        q = {k: e.get().strip() for k, e in kirish.items()}
        q["sud_hammaga_toshkent"] = "Ha" if hammaga.get() else "Yo'q"
        t = tezlik.get()      # soniyalar qo'lda o'zgartirilgan bo'lsa - "Qo'lda"
        if t not in A.TEZLIKLAR or any(q[k] != v for k, v in A.TEZLIKLAR[t].items() if k in q):
            t = "Qo'lda"
            tezlik.set(t)
        q["tezlik"] = t
        A.sozlamalarni_saqlash(q)
        sozlamalar()
        print(f"    OK -> Sozlamalar saqlandi: {sozlamalar_matni()}")
        if not jim:
            yukla()
    tez_f = ttk.Frame(soz)
    tez_f.grid(row=3, column=0, columnspan=8, sticky="w", pady=(6, 0))
    ttk.Label(tez_f, text="Tezlik:", font=("Segoe UI", 10, "bold")).pack(side="left")
    tezlik = tk.StringVar(value=A.S.get("tezlik") if A.S.get("tezlik") in A.TEZLIKLAR else "Qo'lda")

    def tezlik_tanlandi():
        for k, v in A.TEZLIKLAR.get(tezlik.get(), {}).items():   # tayyor tezlik - soniyalar o'zi yoziladi
            if k in kirish:
                kirish[k].delete(0, "end")
                kirish[k].insert(0, v)
        saqla(jim=True)                                            # darhol yoziladi
    for nom in list(A.TEZLIKLAR) + ["Qo'lda"]:
        ttk.Radiobutton(tez_f, text=nom, value=nom, variable=tezlik, command=tezlik_tanlandi).pack(side="left", padx=8)
    # o'z soniyalaringiz: yozing -> Saqlash (yoki Enter) - tezlik o'zi "Qo'lda" bo'ladi
    for k, nom in (("pauza_min", "Pauza min (s)"), ("pauza_max", "Pauza max (s)"), ("kechikish_koef", "Kutish koef")):
        ttk.Label(tez_f, text=nom + ":").pack(side="left", padx=(12, 3))
        e = ttk.Entry(tez_f, width=6, font=SHRIFT, justify="center")
        e.insert(0, A.S.get(k, ""))
        e.pack(side="left")
        e.bind("<Return>", lambda _e: saqla(jim=True))
        kirish[k] = e
    ttk.Label(tez_f, text="(pauza - har bosishdan keyin; koef 1 = sekin, 0.5 = tez)", foreground="#777").pack(
        side="left", padx=8)
    ttk.Button(soz, text="Saqlash va yangilash", command=saqla).grid(row=1, column=6, rowspan=2, padx=12)

    tugmalar = ttk.Frame(root, padding=(10, 6))
    tugmalar.pack(fill="x")
    banner = tk.Frame(root, bg="#FFE58F", highlightthickness=2, highlightbackground="#E0A800")
    banner_matn = tk.Label(banner, text="", bg="#FFE58F", font=("Segoe UI", 12, "bold"), justify="left", anchor="w")
    banner_matn.pack(side="left", fill="x", expand=True, padx=12, pady=10)

    def banner_javob(qiymat):
        req = holat["odam"]
        if req:
            req["javob"] = qiymat
            req["event"].set()
        holat["odam"] = None
        banner.pack_forget()
    tk.Button(banner, text="O'tkazib yuborish (bu mijoz)", bg="#C62828", fg="white", font=("Segoe UI", 10, "bold"),
              relief="flat", padx=12, pady=6, command=lambda: banner_javob("otkaz")).pack(side="right", padx=10)

    holat_yozuv = ttk.Label(root, text="", font=("Segoe UI", 10, "bold"), padding=(10, 2))
    holat_yozuv.pack(fill="x")
    bolinma = ttk.PanedWindow(root, orient="vertical")
    bolinma.pack(fill="both", expand=True, padx=10, pady=(4, 10))
    ramka = ttk.Frame(bolinma)
    ustunlar = [("n", "№", 40), ("papka", "Mijoz (papka)", 330), ("sud", "Sud (ariza / manzil bo'yicha)", 300),
                ("kv", "Papkadagi kvitansiya", 260), ("holat", "Holat", 260)]
    tv = ttk.Treeview(ramka, columns=[u[0] for u in ustunlar], show="headings")
    for k, nom, en in ustunlar:
        tv.heading(k, text=nom)
        tv.column(k, width=en, anchor="w", stretch=k in ("papka", "holat"))
    for teg, rang in (("BOR", "#E8F5E9"), ("KERAK", "#FFD8A8"), ("SUDSIZ", "#FFC9C9"), ("OK", "#C8F7C5"),
                      ("XATO", "#FFC9C9"), ("ISHLANMOQDA", "#E1D5FF")):
        tv.tag_configure(teg, background=rang)
    sb = ttk.Scrollbar(ramka, orient="vertical", command=tv.yview)
    tv.configure(yscrollcommand=sb.set)
    tv.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    bolinma.add(ramka, weight=3)
    log_oyna = scrolledtext.ScrolledText(bolinma, height=10, font=("Consolas", 10), bg="#1E1E1E", fg="#E8E8E8", wrap="word")
    bolinma.add(log_oyna, weight=2)
    ishlar = []

    def chiz(royxat):
        ishlar[:] = royxat
        for i in tv.get_children():
            tv.delete(i)
        sanoq = {"bor": 0, "kerak": 0, "pdf": 0, "sudsiz": 0}
        for n, ish in enumerate(royxat, 1):
            h, r = holati(ish)
            sanoq[h] += 1
            kv = [A.kvitansiya_oqish(y) for k, _n, y in ish.hujjatlar if k == "tolov" and y.lower().endswith(".pdf")]
            kv = sorted((k for k in kv if k), key=lambda k: not mosmi(k, ish))      # mos kelgani birinchi
            kv = [f"{k['raqam']} ({k['sud'].replace('Фуқаролик ишлари бўйича ', '')[:24]}, {k['summa']:.0f}, "
                  f"STIR {k.get('tolovchi_stir') or '?'})" for k in kv]
            matn, teg = {"bor": (f"BOR - mos: {r}", "BOR"), "pdf": (f"yaratilgan {r}, PDF yuklanadi", "KERAK"),
                         "kerak": ("YARATILADI", "KERAK"),
                         "sudsiz": ("sud aniqlanmadi - AdolatAvtomat da 'Sud' ni yozing", "SUDSIZ")}[h]
            tv.insert("", "end", iid=ish.papka, values=(n, ish.papka, ish.sud["nomi"] if ish.sud else "?",
                                                        "; ".join(kv) or "-", matn), tags=(teg,))
        holat_yozuv.config(text=f"Mijozlar: {len(royxat)} ta.  Mos kvitansiya bor: {sanoq['bor']}.  "
                                f"Yaratiladi: {sanoq['kerak'] + sanoq['pdf']}.  Sud aniqlanmagan: {sanoq['sudsiz']}.")

    def yukla():
        papka_yozuv.config(text=f"Mijoz papkalarini shu yerga tashlang: {A.MIJOZLAR_PAPKA}   "
                                f"(har mijoz - alohida papka, ichida da'vo ariza)  ->  🔄 Yangilash")
        holat_yozuv.config(text="Mijozlar o'qilmoqda...")

        def ish():
            try:
                uiq.put(("royxat", mijozlar()))
            except Exception as e:
                print(f"!!! Ro'yxat o'qilmadi: {e}")
                traceback.print_exc()
        threading.Thread(target=ish, daemon=True).start()

    def gui_odam(xabar, shart=None, vaqt=None):
        A.toxtat_tekshir()
        req = {"xabar": xabar, "event": threading.Event(), "javob": ""}
        print("\n    [SIZ KERAKSIZ] " + " | ".join(q.strip() for q in xabar.strip().splitlines()))
        uiq.put(("odam", req))
        A.signal()
        try:
            while not req["event"].wait(0.4):
                A.toxtat_tekshir()
                if shart:
                    try:
                        if shart():
                            return ""
                    except Exception:
                        pass
            A.toxtat_tekshir()
            return req["javob"]
        finally:
            uiq.put(("odam_yop", req))
    A.ODAM_ISHLOVCHI = gui_odam

    def boshlash():
        if holat["ishlayapti"]:
            return
        saqla(jim=True)               # oynada o'zgartirilgan, lekin saqlanmagan sozlamalar ham qo'llansin
        kerak = sum(1 for i in ishlar if holati(i)[0] in ("kerak", "pdf"))
        if not kerak:
            messagebox.showinfo("Kvitansiya", "Hamma mijozda mos kvitansiya bor - yaratiladigani yo'q.")
            return
        if not messagebox.askyesno("Kvitansiya yaratish",
                                   f"{kerak} ta mijozga billing.sud.uz da kvitansiya yaratiladi "
                                   f"({summa():.0f} so'm, har biri o'z sudiga).\n\n"
                                   "Har birida brauzerda MATEMATIK MISOLNI O'ZINGIZ yechasiz (sariq panel + signal).\n"
                                   "PDF lar mijoz papkalariga saqlanadi.\n\nBoshlaymizmi?"):
            return
        holat["ishlayapti"] = True
        t_bosh.config(state="disabled")
        t_stop.config(state="normal")

        def ishchi():
            try:
                jarayon(list(ishlar), xabar=lambda p, m, t: uiq.put(("qator", p, m, t)))
            except Exception as e:
                print(f"!!! {type(e).__name__}: {e}")
                traceback.print_exc()
            uiq.put(("tugadi",))
        threading.Thread(target=ishchi, daemon=True).start()

    def toxtat():
        A.toxtatish()                  # darhol to'xtaydi, Chrome yopiladi
        if holat["odam"]:
            banner_javob("")

    def och(yol):
        try:
            os.startfile(yol)
        except OSError as e:
            messagebox.showerror("Xato", str(e))

    t_bosh = ttk.Button(tugmalar, text="▶  Kvitansiyalarni yaratish", style="Katta.TButton", command=boshlash)
    t_stop = ttk.Button(tugmalar, text="⏹  To'xtatish", style="Toxtat.TButton", command=toxtat, state="disabled")
    for t in (t_bosh, t_stop):
        t.pack(side="left", padx=(0, 8))
    def papka_tanla():
        from tkinter import filedialog
        yol = filedialog.askdirectory(title="Mijozlar papkasi (ichida har mijozning papkasi)", initialdir=A.MIJOZLAR_PAPKA)
        if yol:
            A.sozlamalarni_saqlash({"kvitansiya_mijozlar": os.path.normpath(yol)})
            sozlamalar()
            yukla()

    for matn, buyruq in (("📂 Papkani tanlash", papka_tanla), ("🔄 Yangilash", yukla),
                         ("📁 KVITANSIYALAR (mijozlar)", lambda: och(A.MIJOZLAR_PAPKA)),
                         ("📁 TAYYOR KVITANSIYALAR", lambda: och(TAYYOR_PAPKA)),
                         ("📊 kvitansiyalar.csv", lambda: och(JURNAL) if os.path.isfile(JURNAL) else None),
                         ("📄 Loglar", lambda: och(A.LOG_PAPKA))):
        ttk.Button(tugmalar, text=matn, command=buyruq).pack(side="left", padx=4)

    def navbat():
        try:
            while True:
                x = uiq.get_nowait()
                if x[0] == "log":
                    log_oyna.insert("end", x[1])
                    log_oyna.see("end")
                elif x[0] == "royxat":
                    chiz(x[1])
                    if os.environ.get("KV_SKRIN"):
                        root.after(800, skrin)
                elif x[0] == "qator" and tv.exists(x[1]):
                    v = list(tv.item(x[1], "values"))
                    v[4] = x[2]
                    tv.item(x[1], values=v, tags=(x[3],))
                    tv.see(x[1])
                elif x[0] == "odam":
                    holat["odam"] = x[1]
                    banner_matn.config(text=x[1]["xabar"].strip())
                    banner.pack(fill="x", padx=10, pady=4, after=tugmalar)
                    root.lift()
                    root.attributes("-topmost", True)
                    root.after(1500, lambda: root.attributes("-topmost", False))
                elif x[0] == "odam_yop":
                    if holat["odam"] is x[1]:
                        holat["odam"] = None
                        banner.pack_forget()
                elif x[0] == "tugadi":
                    holat["ishlayapti"] = False
                    t_bosh.config(state="normal")
                    t_stop.config(state="disabled")
                    yukla()
        except queue.Empty:
            pass
        root.after(150, navbat)

    def skrin():          # faqat ichki sinov uchun
        from PIL import ImageGrab
        root.update()
        x, y = root.winfo_rootx(), root.winfo_rooty()
        ImageGrab.grab((x, y, x + root.winfo_width(), y + root.winfo_height())).save(os.environ["KV_SKRIN"])
        os._exit(0)

    def yopish():
        if holat["ishlayapti"] and not messagebox.askyesno("Chiqish", "Kvitansiyalar yaratilmoqda. Baribir chiqasizmi?"):
            return
        root.destroy()
        os._exit(0)
    root.protocol("WM_DELETE_WINDOW", yopish)
    print(f"KVITANSIYA YARATISH {VERSIYA}\nPapka: {A.PAPKA}\nLog: {log}\n")
    yukla()
    root.after(150, navbat)
    root.mainloop()


def main():
    if "--tekshir" in sys.argv:          # brauzersiz: kim uchun kvitansiya yaratilishi
        A.log_boshlash()
        sozlamalar()
        print(f"KVITANSIYA YARATISH {VERSIYA}  |  {A.MIJOZLAR_PAPKA}  |  summa {summa():.0f}"
              f"  |  hamma Toshkent: {A.hammaga_toshkent()}")
        for ish in mijozlar():
            h, r = holati(ish)
            print(f"  {h:7} {r:13} {ish.sud['nomi'] if ish.sud else '?':30} {ish.papka}")
        return
    gui()


if __name__ == "__main__":
    main()
