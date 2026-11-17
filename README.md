# Murojaatlar Telegram boti

Python 3.11+, python-telegram-bot 22.x, APScheduler 3.x va SQLite.
Bot matnlari o'zbek tilida. Taxminan 20 foydalanuvchi/kun uchun bitta jarayon yetarli.

## Ishga tushirish (Windows PowerShell)

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py --check
.\.venv\Scripts\python.exe main.py
```

Git Bash orqali ham ishga tushirish mumkin:

```bash
source .venv/Scripts/activate
python main.py
```

To'xtatish: `Ctrl+C`. "Botning boshqa nusxasi allaqachon ishlayapti" chiqsa,
avval oldingi terminaldagi yoki fondagi botni to'xtating. `data/bot.lock` faylini
o'chirish shart emas: jarayon tugashi bilan operatsion tizim qulfni bo'shatadi.

Mavjud `.env` faylingizni saqlang; tokenni kodga yozmang. Kerakli sozlamalar:

```dotenv
BOT_TOKEN=BotFather_bergan_token
ADMIN_IDS=123456789,987654321
TIMEZONE=Asia/Tashkent
ADMIN_REMINDER_HOUR=9
SMS_API_KEY=
SMS_API_URL=
SMS_TARGET=both
```

`ADMIN_IDS` — username emas, adminlarning raqamli Telegram ID lari. Telegram hisobingizning raqamli ID sini kiriting. Bo'sh bo'lsa, hech kim admin huquqini olmaydi.
`.env` o'zgargach botni qayta ishga tushiring. `.env.example` faqat namuna.
Token, kontaktlar va murojaat mazmuni logga yozilmaydi; `.env`, `data/`, `logs/`
va shaxsiy ma'lumotli Excel/Word fayllari `.gitignore` da.

## Tugmali ish tartibi

Ochiq buyruqlar: `/start`, `/menu`, `/bekor`. Qolgan amallar tugmalarda.

1. **Xodimlar -> Excel yuklash** orqali xodimlarni kiriting yoki **Qo'shish** ni tanlang.
2. Xodim **Ulanish** ni bosib, Telegram kontakt tugmasida o'z telefonini yuboradi.
3. **Xodimlar → Ulanish so'rovlari (N)** da **✅ Tasdiqlash** ni bosing.
4. **Murojaatlar -> Excel yuklash** orqali jadvalni yuboring.
5. **Yuborish -> Darhol** yoki **Soat tanlash** ni bir marta tanlang.
   Rejim tanlanmagan bo'lsa, import natijasining o'zida Darhol / 09:00 / Boshqa vaqt chiqadi.
6. **Holat** da Telegram va SMS natijalarini, e'tibor kerak murojaatlarni ko'ring.

Admin bosh menyusi beshta bo'limdan iborat: Murojaatlar, Xodimlar, Yuborish,
Holat, Qo'shimcha. Admin bo'lmagan hisobga bu bo'limlar hamda callback amallari yopiq.
Xodim **Murojaatlarim** orqali faqat o'ziga biriktirilgan, vaqti kelgan murojaatlarni
matn sifatida o'qiydi. Oddiy foydalanuvchi **Murojaat yuborish** da ism, telefon va matn kiritadi.

Murojaatlar 5 tadan sahifalanadi. Sahifalar tugmasi sahifa almashtirishni ochadi.
Murojaatni mijoz ismi bo'yicha tanlang: ⏳ kutmoqda, ✅ yuborildi, ⚠️ e'tibor kerak.
Kartochkada holatga mos amallar chiqadi: hozir yuborish, kun va soat tanlash,
xodimni ro'yxatdan biriktirish, xato kanalni qayta urinish yoki noma'lum natijani belgilash.
Rad etish, qayta urinish va noma'lum natijani tasdiqlashdan oldin qo'shimcha tasdiq so'raladi.
ID yoki xodim ismini qayta yozish talab qilinmaydi.

Matn kerak bo'lsa, bot bitta savol va namuna beradi. Xato javobda shu qadamda qoladi.
Har qadamda **❌ Bekor** yoki `/bekor` bilan chiqish mumkin. Ichki ekranlarda
**Orqaga** va **Menyu** bor. Bosh menyu yozish maydoni ostidagi doimiy klaviaturada turadi.
Ichki tanlovlar xabar ichidagi tugmalarda ochiladi. Matn kiritilgach navbatdagi savol
chat oxirida yangi xabar bo‘lib chiqadi; fayllar ham alohida yuboriladi.
Kontakt ulashish Telegram talabi sababli alohida telefon tugmasidan foydalanadi.
Bot qayta ishga tushsa, tugallanmagan suhbatni `/menu` bilan boshlang;
eski tugma bosilganda bot xabar beradi.

### Avtomatik yuborish va Excel tahrirlash

- Darhol rejimi vaqtsiz yangi murojaatlarni yuklangach navbatga oladi.
- Soat rejimi eng yaqin shu soatni belgilaydi; o'tgan bo'lsa ertangi kun tanlanadi.
  Har murojaat bir marta yuboriladi. Rejim SQLite'da saqlanadi.
- Excelda alohida vaqt bo'lsa o'sha vaqt ustun. Umumiy rejim o'zgarganda faqat
  avtomatik belgilangan va hali yuborish boshlanmagan kelajak vazifalari yangilanadi.
- O'chirish keyingi vaqtsiz murojaatlar uchun umumiy jadvalni o'chiradi;
  oldin rejalashtirilgan va Excelda belgilangan vaqtlar saqlanadi.
- Xodim ismi bo'sh bo'lsa, tasdiqlangan va smenasi mos xodim avtomatik tanlanadi.
  Noma'lum ism boshqa xodimga almashtirilmaydi.
- **Qo'shimcha** da bo'sh shablonlar, xodimlar va murojaatlarni Excel olish bor.
  Tahrirlangan faylni tegishli bo'limning **Excel yuklash** tugmasida qayta yuboring.
  Murojaat ID si o'zgarmasin; bo'sh vaqt/xodim avvalgi qiymatni saqlaydi.
  Yuborish boshlangan qatorni o'zgartirish butun importni rad etadi.
- Eski jadvalning asl ustunlari `asl:` prefiksi bilan saqlanadi.
- Tasdiqlanmagan xodimga Telegram ham, SMS ham yuborilmaydi.

### Yangilanishdan keyin

Ishlayotgan terminalda **Ctrl+C**, keyin **python main.py**, Telegram'da **/start**.
Yangi menyu eski klaviaturani almashtiradi. Baza migratsiyasi talab qilinmaydi;
saqlangan murojaatlar, yuborish holatlari va rejalashtirish davom etadi.

## Excel formatlari

Sarlavhalar birinchi qatorda, ma'lumot birinchi varaqda bo'lsin. 5 MB, 5000 qator
va 60 ustungacha. Formulalar qabul qilinmaydi. Xatoda qator raqami va sabab
ko'rsatiladi, butun import bekor qilinadi. Qisman import bo'lmaydi.

**murojaatlar.xlsx:** `id`, `mijoz_ismi`, `telefon`, `murojaat_matni`,
`yuborish_vaqti`, ixtiyoriy `xodim`.

- `id`: 1–64 ta harf/raqam, `_`, `-`, `.`; yagona bo'lsin.
- `telefon`: xalqaro format, masalan `+998901234567`; 9 xonali O'zbekiston
  raqamlari avtomatik `+998` bilan to'ldiriladi. Excelda matn sifatida saqlang.
- `mijoz_ismi`, `murojaat_matni`: Excel katagi chegarasigacha (32767 belgi).
  Botning oddiy foydalanuvchi anketasida qisqa kiritish cheklovi saqlanadi.
- `yuborish_vaqti`: Excel sana-vaqti, `YYYY-MM-DD HH:MM` yoki `DD.MM.YYYY HH:MM`.
  Bo'sh katak uchun umumiy **Yuborish** rejimi ishlatiladi; jadval hali tanlanmagan bo'lsa,
  bot uni bir marta tanlashni taklif qiladi. O'tgan vaqt darhol navbatga tushadi.
- `xodim`: ro'yxatdagi to'liq ism, Telegram ID yoki `@username`. Noma'lum xodim
  standart formatda xato hisoblanadi. Bo'sh bo'lsa, shu yuborish vaqtida smenada
  bo'lgan, admin tasdiqlagan xodimlar
  orasidan eng kam murojaat biriktirilgan xodim tanlanadi; tenglikda bazadagi
  tartib ishlatiladi. Mos xodim bo'lmasa, murojaat kutadi va admin eslatma oladi.
- Bir xil ID va mazmun qayta yuklansa, yangi yuborish yaratilmaydi. Shu ID ostidagi
  ma'lumot faqat yuborish boshlanmagan bo'lsa yangilanadi.

**xodimlar.xlsx:** `ism`, `telefon`, ixtiyoriy `telegram_id`, `smena_vaqti`,
`qoshimcha_telefon`.

`qoshimcha_telefon` ustuni yoki katagi bo'sh bo'lishi mumkin. Eski fayllar ham
qabul qilinadi; ularda qo'shimcha raqam bo'sh deb olinadi. Xodimning asosiy va
qo'shimcha raqami bir xil bo'lsa, qo'shimchasi saqlanmaydi. Bir raqam boshqa
xodimning istalgan telefon ustunida takrorlansa, butun import bekor qilinadi;
xato xabarida qator va sabab ko'rsatiladi.

Raqamlar mavjud `phone_number` funksiyasi orqali `+998XXXXXXXXX` shakliga keladi:
`90 123 45 67`, `901234567`, `998901234567`, `+998 (90) 123-45-67` bir xil
`+998901234567` bo'ladi. Boshqa davlat kodi va noto'g'ri uzunlik rad etiladi.
Xodimning o'z kontakti asosiy **yoki qo'shimcha** raqamga mos kelsa ulanish
so'rovi ochiladi; keyin rahbar tasdiqlaydi.

Rahbar yangi xodimni Excel ishlatmasdan **Xodimlar -> Qo'shish** orqali
ism va telefon bilan kiritishi mumkin. Telegram ID hamda smena ixtiyoriy.
Oxirida qo'shimcha telefon ham so'raladi; `yo'q` yoki `-` bilan o'tkazib yuboring.
Xodim o'z telefonini ulashgach, rahbar tasdiqlaydi.

- `telegram_id`: musbat raqamli ID, `@username` yoki bo'sh. Maydon faqat
  ma'lumot uchun; ulanish tasdig'i har doim Telegram yuborgan o'z kontakt
  telefonini Excel'dagi telefon bilan solishtirib beriladi.
- Telegram bot foydalanuvchiga birinchi bo'lib shaxsiy chat ocholmaydi: `/start` shart.
  Telefon mos kelsa ulanish so'rovi ochiladi, admin tasdiqlagach ID bog'lanadi; tasdiqlangan ID
  username o'zgargani sababli boshqa odamga avtomatik o'tkazilmaydi.
- `smena_vaqti`: `09:00-18:00`, `22:00-06:00`; bo'sh yoki teng vaqt — cheklovsiz.
  Smena avtomatik biriktirishda ishlatiladi. Adminning aniq tanlovi ustun.
- Ismlar yagona bo'lsin. Bir xil ism bilan qayta import kontaktlarni yangilaydi;
  boshqa xodimga tegishli telefon yoki Telegram identifikatori rad etiladi.
  Asosiy telefon, qo'shimcha telefon yoki Telegram manzili o'zgarsa, qayta ulanish va admin tasdig'i talab
  qilinadi. Avvalgi tasdiqlash tugmalari yaroqsizlanadi. Yuborilgan holatlar o'zgarmaydi.

### Papkadagi eski murojaatlar jadvali

`Кириш рақами`, `Аризачи`, `Ҳужжат мазмуни`, `Кимга юборилган` sarlavhalari
dastlabki 20 qatordan topilsa, eski format ham qabul qilinadi. Kirish raqamidan
barqaror `L-...` bot ID si yaratiladi, asl raqam murojaat matnida saqlanadi.
`Аризачи` uzun bo'lsa ham to'liq saqlanadi; bo'sh bo'lsa `Ko'rsatilmagan` yoziladi.
Hujjat mazmuni bo'sh bo'lsa, shu haqda belgi saqlanadi. Barcha asl ustunlar saqlanadi.
`Кимга юборилган` ro'yxatdagi ismga mos bo'lsa biriktiriladi. Mos kelmasa import
to'xtamaydi: ism saqlanadi, murojaat biriktirishni kutadi va tasodifiy boshqa
xodimga yuborilmaydi. Keyin mos xodimni yuklang yoki **Xodim biriktirish**ni bosing.
Telefon bo'lmasa, `Ko'rsatilmagan` saqlanadi; SMS xodim ro'yxatidagi raqamga boradi.
`Муддати` — hujjat muddati; yuborish vaqtiga aylantirilmaydi. **Yuborish**
tugmasida bir marta jadval tanlash yetarli; alohida ID kiritish shart emas.
Asl fayl o'zgartirilmaydi. Word xodimlar ro'yxati import qilinmaydi; ma'lumotni
**Qo'shimcha -> Bo'sh shablonlar** dagi xodimlar Excel formatiga ko'chirish kerak.

## Eslatmalar va ishonchlilik

- `data/bot.sqlite3`: xodimlar, murojaatlar, vaqtlar, kanal holatlari va eslatmalar.
- `data/jobs.sqlite3`: APScheduler'ning doimiy SQLite job store'i.
- Worker har 15 soniyada muddati kelgan yuborishlarni tekshiradi. Bot o'chib turgan
  paytda muddati o'tgan murojaatlar qayta ishga tushgach bajariladi.
- Telegram va SMS mustaqil: birining xatosi ikkinchisini takroran yubortirmaydi.
  SMS hali ulanmagan bo'lsa, Telegram ishlaydi, SMS `API kutilmoqda` bo'ladi.
- Aniq rad etilgan yuborish: 1 asosiy urinish + 3 qayta urinish, 30/60/120 soniyali
  kutish. Telegram rate-limit kutish vaqti ham hurmat qilinadi. Natija bazada,
  xatoning turi `logs/bot.log` da. kartochkadagi **Qayta urinish**
  orqali faqat xato yoki SMS API kutilayotgan kanal yana navbatga qo'yiladi.
- Telegram API idempotency kaliti bermaydi: aloqa uzilib, server qabul qilgan-qilmagani
  noma'lum bo'lsa, mutlaq "exactly once" kafolati mumkin emas. Dublikat xavfini
  kamaytirish uchun bunday holat `natija noma'lum` bo'ladi va avtomatik qaytarilmaydi.
  Yuborish paytida jarayon uzilishi ham shunday qayd etiladi. Qabul qiluvchidan
  tekshirgach kartochkadagi **Natijani belgilash -> Yuborildi / Yuborilmadi**
  tugmalarini ishlating; ikkinchisi tasdiqdan keyin qayta navbatga qo'yadi.
- Har kuni 09:00 dan boshlab bugungi reja bir marta yuboriladi; vaqt belgilanmagan
  yoki muddati o'tgan tugallanmagan murojaatlar haqida soatiga bir eslatma bor.
  Eslatmalar admin ID va vaqt oralig'i bo'yicha bazada takrorlanishdan himoyalangan.
  **Murojaatlar -> Bugun** istalgan paytda joriy ro'yxatni ko'rsatadi.
- Bitta ishchi nusxa ishlating. Mahalliy OS lock ikkinchi nusxani to'xtatadi.
  Boshqa kompyuterda ayni token bilan polling ishga tushirmang.

Kompyuter o'chsa yoki uyquga ketsa bot ishlamaydi. Doimiy ish uchun keyin server
yoki Windows Task Scheduler orqali avtomatik ishga tushirishni sozlash mumkin.
Baza zaxirasi uchun botni to'xtatib butun `data/` papkasini nusxalang.

## SMS adapter

`SMS_TARGET=both` (standart) — asosiy va mavjud qo'shimcha raqamga;
`SMS_TARGET=primary` — faqat asosiy raqamga. Noto'g'ri qiymat botning ishga
tushishini to'xtatadi. Sozlama o'zgargach botni qayta ishga tushiring.
Qabul qiluvchi raqamlar murojaat muddati kelib, xodim tasdiqlangach navbatga
olishda qayd etiladi. Hali birorta haqiqiy SMS urinishi boshlanmagan
`pending`/`API kutilmoqda` yozuvlar kontakt yoki rejimga mos yangilanadi.
Yuborish boshlangan raqamlar tarixi saqlanadi; sozlama ularni qayta yaratmaydi. Eski `sent` yoki `unknown`
SMS holati asosiy raqamga ko'chiriladi; tarixiy xabarlar qo'shimcha raqamga
avtomatik qayta yuborilmaydi.

`sms_deliveries` jadvalida har bir `(murojaat, telefon)` uchun holat, urinishlar,
keyingi urinish va xato alohida saqlanadi. Kartochkadagi **Urinishlar** ikkala raqam natijasini
ko'rsatadi; **Holat** SMS'ni murojaat bo'yicha umumlashtiradi. Barcha raqamlarda
API qabul qilgandagina umumiy holat `yuborildi` bo'ladi. **Qayta urinish (SMS)** faqat
`xato` yoki `API kutilmoqda` raqamlarni qayta navbatga qo'yadi; muvaffaqiyatli
raqamga xabar takrorlanmaydi. **SMS natijasini belgilash** faqat noma'lum natijali
raqamlarga qo'llanadi. Telegram holatiga ta'sir qilmaydi.
Provayder ulanmaguncha har ikkala SMS raqamida `API kutilmoqda` saqlanadi;
bu kutish urinishlar limitini sarflamaydi.

### Qo'shimcha telefon uchun migratsiya va qayta ishga tushirish

1. Ishlayotgan botni `Ctrl+C` bilan to'xtating; `data/` papkasini zaxiralang.
2. `.env` da ixtiyoriy `SMS_TARGET=both` yoki `SMS_TARGET=primary` belgilang.
   Bu qator bo'lmasa `both` ishlatiladi.
3. `python main.py` bilan ishga tushiring. `Database` eski bazani ochganda
   `employees.phone_extra` ustunini `NULL` qiymat bilan qo'shadi va
   `sms_deliveries` jadvalini yaratadi. Eski xodimlar, murojaatlar, Telegram/SMS
   holatlari va rejalashtirilgan ishlar o'chirilmaydi. SQL'ni qo'lda ishlatish kerak emas.
4. Excel'da qo'shimcha raqamlarni to'ldirib qayta yuklang. Qo'shimcha raqamni
   qo'shish, almashtirish yoki bo'shatish ham telefon o'zgarishi hisoblanadi:
   xodim yana o'z kontaktini yuborib, rahbar tasdig'ini oladi.

`sms.py` dagi `async send_sms(phone, text)` hozir `SMSNotConfigured` qaytaradi;
hech qanday soxta muvaffaqiyat yoki haqiqiy HTTP so'rovi yo'q. `SMS_API_KEY` va
`SMS_API_URL` `.env` dan olinadi, lekin ularni to'ldirishning o'zi provayderni ulamaydi.
Provayder API va alohida SMS shablonlarini berganingizdan keyin shu adapter ulanadi.
Kelajak adapteri tasdiqlangan javobda provayder message ID sini qaytarishi,
aniq rad etilishda `SMSRetryableError`, noma'lum natijada `SMSDeliveryUnknown`
berishi kerak. `sent` provayder/Telegram API qabul qilganini bildiradi, o'qilganini
yoki telefon apparatiga yetganini bildirmaydi; SMS delivery receipt keyin ulanadi.

## Tekshirish

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe main.py --check
```

Testlar sun'iy ma'lumotlar va vaqtinchalik bazalardan foydalanadi, hech kimga
haqiqiy xabar yubormaydi. `--check` faqat Telegram avtorizatsiyasini tekshiradi.

Texnik manbalar: [python-telegram-bot ApplicationBuilder](https://docs.python-telegram-bot.org/en/stable/telegram.ext.applicationbuilder.html),
[APScheduler SQLAlchemyJobStore](https://apscheduler.readthedocs.io/en/3.x/modules/jobstores/sqlalchemy.html).
# m_bot
