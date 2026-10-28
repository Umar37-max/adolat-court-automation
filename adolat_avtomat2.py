"""
ADOLAT AVTOMAT 2 - BUYRUQ tartibidagi arizalar (cabinet.sud.uz/cases/create/civil/decree/first).

Hamma narsa adolat_avtomat.py dagidek (kod o'sha yerdan olinadi - u yerdagi tuzatishlar bu dasturga ham o'tadi).
Farqi faqat:
  1) murojaat yaratish sahifasida murojaat turi - "Buyruq tartibida" ("Da'vo tartibida" emas);
  2) 2-qadam: "Nizo turi" / "Ariza turi" yo'q, ish turkumi - "111 - ёзма битимга асосланган ..." va
     qo'shimcha "111.2 - майда ва маиший кредит ..." (sozlamalar.ini da o'zgartirsa bo'ladi).
Ma'lumotlar (mijozlar, ishlar.xlsx, natijalar.csv, sozlamalar.ini, loglar) - O'Z papkasida, 1-dastur bilan aralashmaydi.
Chrome profili umumiy (ONE ID, E-IMZO), lekin port boshqa - ikkala dasturni bir vaqtda ishlatmang.
"""
import os
import sys

if getattr(sys, "frozen", False):          # AdolatAvtomat2.exe - o'z papkasi
    PAPKA2 = os.path.dirname(os.path.abspath(sys.executable))
else:                                      # manbadan - loyiha ichidagi AdolatAvtomat2\ papkasi
    PAPKA2 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "AdolatAvtomat2")
# 1-dastur "mijozlar yo'q bo'lsa - yuqori papka" qoidasi bu dasturni 1-dastur ma'lumotlariga ulab qo'ymasin
for _p in ("mijozlar", "umumiy_hujjatlar"):
    os.makedirs(os.path.join(PAPKA2, _p), exist_ok=True)

import adolat_avtomat as A  # noqa: E402

# manbadan ishga tushirilganda adolat_avtomat o'z papkasini olgan - hamma yo'llarni AdolatAvtomat2\ ga ko'chiramiz
if os.path.normcase(A.PAPKA) != os.path.normcase(PAPKA2):
    eski = A.PAPKA
    for nom in dir(A):
        qiymat = getattr(A, nom)
        if nom.isupper() and isinstance(qiymat, str) and qiymat.startswith(eski + os.sep):
            yangi = PAPKA2 + qiymat[len(eski):]
            if nom == "SUDLAR_FAYL" and not os.path.isfile(yangi):
                continue                   # sudlar jadvali nusxasi yo'q - 1-dasturnikidan o'qiladi
            setattr(A, nom, yangi)
    A.PAPKA = PAPKA2

A.DASTUR_NOMI = "ADOLAT AVTOMAT 2 (BUYRUQ)"
A.VERSIYA = "BUYRUQ tartibida | " + A.VERSIYA
A.CHROME_PORT = 9242                       # 1-dastur Chrome'iga ulanib qolmasin
A.STANDART_SOZLAMALAR.update({
    "nizo_turi": "",                       # buyruq arizasida bu ro'yxatlar yo'q - tanlanmaydi va tekshirilmaydi
    "ariza_turi": "",
    "ish_turkumi": "111 - ёзма битимга асосланган",
    "qoshimcha_ish_turkumi": "111.2 - майда ва маиший кредит",
})
A.S.update(A.STANDART_SOZLAMALAR)
# asosiy ariza - FAQAT nomida "буйруқ/buyruq" bor fayl ("суд буйруғи ариза.pdf"); "даъво" fayllari asosiy emas
# (kalit "davo" qoladi - qolgan kod shu kalit bo'yicha ishlaydi)
A.HUJJAT_TURLARI[0] = ("davo", "Buyruq arizasi", ["буйру", "буюру", "buyru", "buyuru"], [])


def boshlash(b):
    print(">>> Sud turi, murojaat turi (BUYRUQ tartibida), instantsiya...")
    b.xp_bos("Fuqarolik sudi", "Fuqarolik sudi")
    b.xp_bos("arizasi", "Murojaat turi dropdown", exact=False)
    b.xp_bos("Buyruq tartibida", "Buyruq tartibida", exact=False)
    b.xp_bos("Birinchi instantsiya", "Birinchi instantsiya")
    b.xp_bos("Oldinga", "Oldinga")


A.boshlash = boshlash

if __name__ == "__main__":
    if len(sys.argv) == 1 or "--oyna" in sys.argv:
        A.gui_main()
    else:
        A.main()
