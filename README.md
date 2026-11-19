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
ADMIN_REMINDER_HOUR=9
TEXTUP_EMAIL=
TEXTUP_PASSWORD=
SMS_TARGET=both
```

`ADMIN_IDS` — username emas, adminlarning raqamli Telegram ID lari. Faqat loyiha
`.env` faylidagi qiymat ishlatiladi; tizim muhitidagi `ADMIN_IDS` va `${...}` orqali
almashtirish admin huquqini bermaydi. Bo'sh bo'lsa, hech kim admin bo'lmaydi.
Barcha vaqtlar, import, ekranlar, jadval va loglar qat'iy `Asia/Tashkent` (+05:00).
Tizimdagi yoki eski `.env` dagi boshqa `TIMEZONE` qiymati ishlatilmaydi.
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
6. **Holat** da yuborilgan, yuborilmagan va e'tibor kerak murojaatlarni ko'ring.
   SMS natijalari **Sozlamalar → SMS sozlamalari → SMSda bormaganlar** da.

Admin bosh menyusi beshta bo'limdan iborat: Murojaatlar, Xodimlar, Yuborish,
Holat, Sozlamalar. Eski Qo‘shimcha tugmasi ham Sozlamalar bo‘limini ochadi.
Admin bo'lmagan hisobga bu bo'limlar hamda callback amallari yopiq.
Xodim **Murojaatlarim** orqali faqat o'ziga biriktirilgan, vaqti kelgan murojaatlarni
o'qiydi va **✅ Bajarildi** tugmasi bilan ijroni belgilaydi. Vaqt va xodim bazada
saqlanadi; adminga xabar navbatga qo'yiladi. Takroriy bosish yangi xabar yaratmaydi.
Admin kartochkada ijroni, **Holat → Bajarilgan** da ro'yxatni ko'radi.
`Yuborildi` — yetkazish holati, `Bajarildi` — alohida ijro holati.
Oddiy foydalanuvchi **Murojaat yuborish** da ism, telefon va matn kiritadi.

Murojaatlar 5 tadan sahifalanadi. Sahifalar tugmasi sahifa almashtirishni ochadi.
Murojaatni mijoz ismi bo'yicha tanlang: ⏳ kutmoqda, ✅ yuborildi, ⚠️ e'tibor kerak.
**Yuborilgan** Telegramda yuborilganlarni, **Yuborilmagan** esa hali Telegramda
yuborilmagan va bajarilmaganlarni ko'rsatadi. **Hammasi** arxivni ham saqlaydi.
Telegram yuborilgach, SMS natijasidan qat'i nazar **✅ Yuborildi** chiqadi.
Oddiy kartochkada SMS holati, urinishlar soni va SMS qayta urinish tugmasi chiqmaydi.
Kartochkada holatga mos amallar chiqadi: hozir yuborish, kun va soat tanlash,
xodimni ro'yxatdan biriktirish, Telegram xatosini qayta urinish yoki noma'lum natijani belgilash.
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

- **Sozlamalar → SMS sozlamalari → SMS o‘chirish / SMS yoqish** orqali SMS kanalini boshqaring.
  **SMSda bormaganlar** ro'yxati ham shu yerda. SMS tugmasi Sozlamalar boshida ham bor.
  Bu tugma **Yuborish** bo‘limida ham bor. Yangilanishdan keyin `/menu` bosing.
  Sozlama bazada saqlanadi va qayta ishga tushirish talab qilinmaydi. Dastlab SMS yoqilgan.
  SMS o‘chirilganda Telegram ishlaydi; hali yuborilmagan SMS navbati va urinishlar tarixi
  saqlanadi. Qayta yoqish paytida vaqti o‘tgan SMS’lar avtomatik yuborilmaydi.
  Kerakli SMS’ni **SMSda bormaganlar → SMSni navbatga qo‘yish** orqali alohida yuboring.
  Hali vaqti kelmagan murojaatlar jadval bo‘yicha yuboriladi.
  Allaqachon boshlangan tarmoq so‘rovi yakunlanishi mumkin.
- Eski Excelda arizachi ismi bo‘sh bo‘lsa, ro‘yxat, kartochka va Telegram xabarida
  `Юборган ижрочи` ustunidagi qiymat ko‘rsatiladi. Ikkalasi ham bo‘sh bo‘lsa,
  ro‘yxat tugmasida **Murojaat #ID** chiqadi.
- **Holat** ekranida e’tibor talab qiladigan murojaatlar sabablari va soni ko‘rsatiladi.
  Muammo bo‘lmasa **E’tibor kerak** tugmasi chiqmaydi. Murojaat kartochkasida ham
  aniq sabab yoziladi. SMS natijasi bu hisobga kirmaydi. Telegramda yuborilgan
  murojaat xodim keyin uzilgani yoki eski SMS xatosi uchun yana e'tibor ro'yxatiga tushmaydi.
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
- **Holat → Mos xodim topilmagan** va **Xodimi ulanmagan** alohida ro'yxatlardir.
  Muammo aniqlanganda adminga navbatdagi worker tekshiruvida ogohlantirish boradi;
  bir xil murojaat va sabab uchun soatiga bir martadan ko'p takrorlanmaydi.
- Xodim kartochkasidagi **Ulanishni uzish** tasdiqlangach, kirish va yuborish huquqi
  bekor qilinadi, kutayotgan ulanish so'rovlari eskiradi, xodimga xabar beriladi.
  Qayta ulanish uchun o'z kontaktini yana yuborishi va admin tasdiqlashi kerak.
  Ulanish tasdiqlanganda ham xodimga alohida xabar yuboriladi.

### Yangilanishdan keyin

Ishlayotgan terminalda **Ctrl+C**, keyin **python main.py**, Telegram'da **/start**.
Yangi menyu eski klaviaturani almashtiradi. Baza o'zgarishlari avtomatik va qo'shimcha
ustunlar orqali bajariladi; mavjud murojaatlar va yuborish tarixi saqlanadi.
Muddati o'tgan, hali yuborilmagan vazifalar ishga tushgandagi birinchi worker
tekshiruvida yuboriladi. Oldingi `sent` va `unknown` holatlar avtomatik takrorlanmaydi.

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
  Bajarilgan murojaat ham tahrirlanmaydi. Bir faylda bir ID ikki marta kelishi xato;
  xabarda qator va ID ko'rsatiladi. O'zgartirish taqiqlangan mavjud ID uchun
  "Yangi murojaatga yangi ID bering" deyiladi va butun import bekor qilinadi.

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
Bu ustun mavjud bo'lsa, import natijasida alohida **⚠️ Муддати** ogohlantirishi chiqadi.
Asl fayl o'zgartirilmaydi. Word xodimlar ro'yxati import qilinmaydi; ma'lumotni
**Qo'shimcha -> Bo'sh shablonlar** dagi xodimlar Excel formatiga ko'chirish kerak.

## Eslatmalar va ishonchlilik

- `data/bot.sqlite3`: xodimlar, murojaatlar, vaqtlar, kanal holatlari va eslatmalar.
- `data/jobs.sqlite3`: APScheduler'ning doimiy SQLite job store'i.
- Worker har 15 soniyada muddati kelgan yuborishlarni tekshiradi. Bot o'chib turgan
  paytda muddati o'tgan murojaatlar qayta ishga tushgach bajariladi.
- Telegram va SMS mustaqil: birining xatosi ikkinchisini takroran yubortirmaydi.
  TextUP sozlanmagan yoki kirish rad etilgan bo'lsa, Telegram ishlaydi, SMS `API kutilmoqda` bo'ladi.
- Aniq rad etilgan yuborish: 1 asosiy urinish + 3 qayta urinish, 30/60/120 soniyali
  kutish. Telegram rate-limit kutish vaqti ham hurmat qilinadi. Natija bazada,
  xatoning turi `logs/bot.log` da. Oddiy kartochkadagi **Qayta urinish** Telegram uchun.
  SMS xatosi va API kutish holatini **SMSda bormaganlar** ichidagi **SMS qayta urinish**
  orqali qayta navbatga qo'ying.
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

### Timeout, avtomatik backup va loglar

Har bir Telegram/SMS tarmoq yuborishiga 60 soniya timeout bor. Javob olinmasa,
natija `noma'lum` bo'ladi va avtomatik qayta yuborilmaydi. Qo'shimcha tekshiruv
120 soniyadan ortiq qolgan `yuborilmoqda` yozuvlarini ham `noma'lum` ga o'tkazadi.
Bot qayta ishga tushganda uzilib qolgan yuborishlar darhol shu holatga o'tadi.

Worker har kuni Toshkent sanasi bo'yicha `data/backups/` ichiga asosiy baza va
mavjud scheduler bazasidan alohida nusxa oladi. SQLite backup API ishlatiladi:
WAL'dagi tasdiqlangan yozuvlar ham nusxaga kiradi. `PRAGMA quick_check` muvaffaqiyatli
bo'lgach `.partial` fayl yakuniy `.sqlite3` nomiga o'tadi. Shu kunning nusxalari
takror yaratilmaydi; eski nusxalar avtomatik o'chirilmaydi. Backup xatosi yuborishni
to'xtatmaydi, adminga ogohlantirish tayyorlanadi. Diskda joy yetarli bo'lsin.

Log fayli va konsol formatlagichi token, parol, Bearer/JWT hamda exception
matnidagi sirlarni `[REDACTED]` bilan yashiradi. HTTP kutubxona loglari cheklangan.
Sozlamalar obyektining matnli ko'rinishida bot tokeni chiqmaydi.

## SMS adapter

SMS xodimga qisqa bildirishnoma yuboradi; katta murojaat matni Telegramda qoladi.
Shablon (`sms.py` dagi `TASK_NOTIFICATION_TEMPLATE`):

```text
Sizga yangi vazifa biriktirildi. Batafsil: {TELEGRAM_HAVOLA}
```

`{TELEGRAM_HAVOLA}` avtomatik `https://t.me/<bot_username>?start=task_<raqam>`
havolasiga almashtiriladi. Bot username'i ishlayotgan Telegram botdan olinadi;
alohida sozlama kerak emas. Havoladagi raqam murojaatning ichki raqami.
Havola botdagi aynan shu murojaatni ochadi; uzun matnni **Davomi** tugmasida
o'qish mumkin. Faqat o'sha murojaat biriktirilgan, ulanishi tasdiqlangan xodimga
va yuborish vaqti kelgandan keyin ko'rsatiladi. Havola boshqa kishiga yuborilsa,
murojaat mazmuni ochilmaydi. SMS ichiga murojaat matni yoki mijoz telefoni qo'shilmaydi.

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
keyingi urinish va xato alohida saqlanadi. **SMSda bormaganlar** ichidagi SMS kartochkasi
ikkala raqam natijasini ko'rsatadi. Barcha raqamlarda
API qabul qilgandagina SMS umumiy holati `yuborildi` bo'ladi. **SMS qayta urinish** faqat
`xato` yoki `API kutilmoqda` raqamlarni qayta navbatga qo'yadi; muvaffaqiyatli
raqamga xabar takrorlanmaydi. **SMS natijasini belgilash** faqat noma'lum natijali
raqamlarga qo'llanadi. Telegram holatiga ta'sir qilmaydi.
TextUP sozlamalari yetishmasa yoki kirish rad etilsa, SMS `API kutilmoqda` holatida
qoladi; sababi SMS kartochkasida ko'rsatiladi. Bu kutish urinishlar limitini sarflamaydi.

### Botning batafsil ishlash mantiqi

1. **Kirish va rollar.** Bot shaxsiy Telegram chatida ishlaydi. `/start` yoki `/menu`
   foydalanuvchini ro'yxatga oladi va uning roliga mos menyuni chiqaradi. Adminlar
   raqamli Telegram ID orqali aniqlanadi. Oddiy foydalanuvchi murojaat yuboradi;
   tasdiqlangan xodim o'z vazifalarini ko'radi. Admin callback amallari ham tekshiriladi.
2. **Xodimlarni tayyorlash.** Admin ism va asosiy telefon bilan xodim qo'shadi yoki
   Excel yuklaydi. Qo'shimcha telefon, Telegram manzili va smena ixtiyoriy.
   Telefonlar +998 formatiga keltiriladi; boshqa xodimning raqami takrorlanmaydi.
3. **Ulanishni tasdiqlash.** Xodim Telegramning kontakt yuborish tugmasi orqali
   o'z raqamini ulashadi. Asosiy yoki qo'shimcha telefon mos kelsa so'rov ochiladi.
   Admin tasdiqlamaguncha Telegram va SMS yuborish boshlanmaydi. Telefon yoki
   Telegram manzili o'zgarsa qayta ulanish kerak. Ulanishni uzish kirishni bekor qiladi.
4. **Murojaatni qabul qilish.** Admin Excel yuklaydi; oddiy foydalanuvchi ism,
   telefon va matnni ketma-ket kiritadi. Har murojaatda yagona ID bo'ladi.
   Excel xatosida butun import bekor qilinadi. Bir xil ID va mazmun yangi yuborish
   yaratmaydi. Eski jadvalda arizachi bo‘sh bo‘lsa, bot `Юборган ижрочи` qiymatini,
   u ham bo‘sh bo‘lsa, ro‘yxat tugmasida Murojaat #ID ni ko‘rsatadi.
5. **Vaqt tanlash.** Excelda yuborish vaqti bo'lsa o'sha ishlaydi. Bo'sh vaqtga
   umumiy Darhol yoki HH:MM rejimi qo'llanadi. O'tib ketgan soat ertangi kunga
   qo'yiladi. Har murojaat bir marta yuboriladi; HH:MM bir murojaatni har kuni takrorlamaydi.
   Umumiy rejim o'chirilsa oldin belgilangan vaqtlar saqlanadi. Barcha vaqtlar Toshkent vaqti.
6. **Xodim tanlash.** Excelda yoki kartochkada aniq xodim ko'rsatilsa o'sha saqlanadi.
   Xodim bo'sh bo'lsa, yuborish vaqtida smenasi mos tasdiqlangan xodimlar orasidan
   eng kam murojaat biriktirilgani tanlanadi. Hisob jami biriktirilgan murojaatlar
   bo'yicha; tenglikda bazadagi tartib ishlaydi. Mos xodim bo'lmasa murojaat kutadi.
   Eski jadvaldagi noma'lum ism tasodifiy boshqa xodimga almashtirilmaydi.
7. **Navbatni ishlash.** Har 15 soniyada ishchi jarayon vaqti kelgan murojaatlarni
   tekshiradi. Xodim tasdiqlangan bo'lsa Telegramga to'liq matn yuboriladi; uzun matn
   qismlarga bo'linadi. Qabul qilingan har bir qism bazaga yoziladi. Telegram muvaffaqiyati
   umumiy ro'yxatda Yuborildi degani; o'qilganlik tasdig'i olinmaydi.
8. **Ro'yxatlar va holat.** Yuborilgan ro'yxati Telegram natijasiga qaraydi.
   Yuborilmagan ichida kelajakdagi vazifalar va hali hal qilinmaganlari ham bor.
   Hammasi yuborilgan va bajarilganlarni arxiv sifatida saqlaydi. Oddiy kartochkada
   mijoz, telefon, xodim, vaqt, murojaat matni va yuborish holati ko'rinadi.
   Bajarilgan bo'lsa ijro vaqti ham chiqadi. SMS urinishlari bu ekranda chiqmaydi.
9. **E'tibor kerak.** Bu ro'yxat hali yuborilmagan va bajarilmagan murojaat uchun:
   xodim yo'q, xodim ulanmagan, vaqt belgilanmagan, Telegram xatosi yoki noma'lum
   natija bo'lsa chiqadi. Sabablar Holat va kartochkada yoziladi. Muammo nol bo'lsa
   tugma yashiriladi. SMS xatosi umumiy murojaat holatini o'zgartirmaydi.
10. **SMS alohida kanal.** SMS yoqilgan bo'lsa vaqti kelganda xodimga qisqa
    bildirishnoma va murojaat havolasi yuboriladi. Mijozga SMS yuborilmaydi.
    Rejimga qarab asosiy yoki asosiy va qo'shimcha raqam ishlatiladi. Havolani faqat
    tegishli, tasdiqlangan xodim vaqti kelgach ochadi. SMS o'chirilsa navbat saqlanadi;
    qayta yoqilganda vaqti o‘tgan SMS’lar faqat qo‘lda navbatga qo‘yiladi. Hali vaqti kelmagan
    SMS’lar jadval bo‘yicha yuboriladi. Boshlangan so‘rov yakunlanishi mumkin.
11. **SMSda bormaganlar.** Sozlamalar → SMS sozlamalari ichida vaqti kelgan va
    SMS API qabul qilgani hali tasdiqlanmagan murojaatlar ko'rinadi. Kelajakdagi,
    vaqtsiz va barcha raqamlariga SMS qabul qilingan murojaatlar kirmaydi. Pending,
    qayta urinish, API kutish, xato va noma'lum natijalar ko'rinadi. Har bir raqamning
    urinishlari va sababi SMS kartochkasida turadi. Natija noma'lum bo'lsa SMS bormagan
    deb xulosa qilinmaydi: avval qabul qiluvchidan tekshiriladi. Bajarilgan murojaatning
    eski SMS xatosi tarix sifatida ko'rinadi, ammo qayta urinish tugmasi berilmaydi.
12. **Xatolar va takror yuborish.** Qayta urinish mumkin bo'lgan xatolarda bir asosiy
    va uch qo'shimcha urinish bor; kutish 30, 60, 120 soniya. Aniq rad etilgan SMS
    to'g'ridan-to'g'ri xatoga tushishi mumkin. Javob noma'lum bo'lsa avtomatik takrorlanmaydi.
    Admin tekshirgach Yuborildi yoki Yuborilmadi deb belgilaydi. SMS qayta urinish
    faqat xato/API kutish raqamlariga ta'sir qiladi; muvaffaqiyatli raqam va Telegram takrorlanmaydi.
13. **Ijroni yakunlash.** Xodim Murojaatlarim ichida o'zining vaqti kelgan
    vazifasini ochib Bajarildi bosadi. Kim va qachon bajargani saqlanadi, admin
    xabardor qilinadi. Takroriy bosish yangi xabar yaratmaydi. Yetkazish va ijro
    alohida: Yuborildi xabar yuborilganini, Bajarildi ish bajarilganini bildiradi.
14. **Qayta ishga tushirish va zaxira.** Murojaatlar, jadval, SMS tanlovi va yuborish
    tarixi SQLite'da saqlanadi. Bot yoqilgach muddati o'tgan vazifalar tekshiriladi;
    yuborilganlar qayta yuborilmaydi. Jarayon yuborish paytida uzilgan bo'lsa natija
    noma'lumga o'tadi. Har kuni bazalarning zaxirasi olinadi. Tugallanmagan suhbat
    xotirada bo'lgani uchun qayta ishga tushgach `/menu` kerak. Kompyuter uyquda yoki
    o'chirilganida bot ishlamaydi; ayni token bilan bitta ishchi nusxa ishlatiladi.

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

`sms.py` dagi `async send_sms(phone, text)` TextUP bilan ishlaydi. Bot ishga tushganda
`TEXTUP_EMAIL` va `TEXTUP_PASSWORD` `.env` dan yuklanadi. Parolni kodga yoki logga yozmang.
Login: `POST https://api-auth.textup.uz/v1/login`; javobdagi `accessToken` va `user.id`
xotirada saqlanadi. SMS: `POST https://sms-api.textup.uz/v1/send`, Bearer token va
`recipients: [telefon]` bilan. `groupId` yuborilmaydi: xabar faqat belgilangan xodimga boradi.

`TEXTUP_TEMPLATE_ID` va `TEXTUP_NICKNAME_ID` ixtiyoriy. Berilsa tegishli `templateId`
va `nicknameId` yuboriladi; bo'sh bo'lsa so'rovga qo'shilmaydi. Nickname berilmasa,
TextUP hujjatiga ko'ra qisqa raqam ishlatiladi. Shablon ishlatsangiz, TextUP'da
dinamik Telegram havolasi bilan yuboriladigan matnga mos shablon tanlang.
Eski `SMS_API_KEY` va `SMS_API_URL` endi ishlatilmaydi.

Token sababli SMS so'rovi `401` olsa, bot bir marta qayta login qilib urinadi.
Berilgan API hujjatida refresh endpoint yo'q, shu sababli taxminiy endpoint ishlatilmaydi.
`smsId` kelgandagina SMS `yuborildi` bo'ladi va shu ID bazada saqlanadi.
`400` kabi aniq rad etilish `xato` bo'ladi; avtomatik takrorlanmaydi.
Ulanish o'rnatilmagan bo'lsa yoki `429` qaytsa, mavjud qayta urinish jadvali ishlaydi.
SMS so'rovidan keyin aloqa uzilishi, `408`/`5xx`, noto'g'ri javob yoki `smsId` yo'qligi
`natija noma'lum` deb saqlanadi va avtomatik qayta yuborilmaydi.
API javob matni, token va parol loglarga yoki xato xabarlariga kiritilmaydi.
`sent` API qabul qilganini bildiradi; telefon apparatiga yetganini yoki o'qilganini
bildirmaydi. SMS delivery receipt hozir ulanmagan.

## Linux serverda ishga tushirish

`deploy/mj-bot.service` Linux va systemd uchun tayyorlangan. Yo‘riqnomada
bot `/opt/mj_bot` ichida va `mjbot` tizim foydalanuvchisi nomidan ishlaydi.
Serverda Python 3.11 yoki undan yuqori versiya hamda `venv` moduli kerak.
Serverdan Telegram API va, SMS ishlatilsa, TextUP API’ga chiqish kerak.
Bot Telegramni polling orqali tekshiradi. Alohida kiruvchi port ochish talab qilinmaydi.

Serverda bir marta papkalarni tayyorlang:

```bash
sudo useradd --system --user-group --no-create-home --shell /usr/sbin/nologin mjbot
sudo install -d -m 0755 /opt/mj_bot
sudo install -d -m 0700 -o mjbot -g mjbot /opt/mj_bot/data /opt/mj_bot/logs
```

Loyiha kodining amaldagi nusxasini `/opt/mj_bot` ga joylang. Ishga tushirish
uchun `*.py`, `handlers/`, `requirements.txt` va `deploy/` kerak.
Windows’dagi virtual muhitni serverga ko‘chirmang.

```bash
cd /opt/mj_bot
sudo python3 -m venv .venv
sudo .venv/bin/python -m pip install -r requirements.txt
sudo touch /opt/mj_bot/.env
sudo chown root:mjbot /opt/mj_bot/.env
sudo chmod 0640 /opt/mj_bot/.env
sudoedit /opt/mj_bot/.env
```

`.env` ichiga yuqoridagi namuna bo‘yicha bot tokeni, adminlarning raqamli
Telegram ID’lari va kerak bo‘lsa TextUP sozlamalarini yozing. Yangidan
o‘rnatishda xodim va murojaatlar bazasi bo‘sh ochiladi.

Mavjud yozuvlar kerak bo‘lsa, avval eski botni to‘xtating va zaxira oling.
`data/bot.sqlite3` faylini serverdagi `data/` ga nusxalang va
`sudo chown mjbot:mjbot /opt/mj_bot/data/bot.sqlite3` buyrug‘ini bajaring.
Agar to‘xtagandan keyin SQLite WAL fayli qolgan bo‘lsa, ko‘chirish uchun
tekshirilgan SQLite zaxirasidan foydalaning. Bot ishchi jadvalni qayta yaratadi.
Muddati kelgan navbat serverda davom etishi mumkin. Ko‘chirishdan oldin navbatni tekshiring.

```bash
sudo -u mjbot /opt/mj_bot/.venv/bin/python /opt/mj_bot/main.py --check
sudo install -m 0644 deploy/mj-bot.service /etc/systemd/system/mj-bot.service
sudo systemd-analyze verify /etc/systemd/system/mj-bot.service
sudo systemctl daemon-reload
sudo systemctl enable --now mj-bot
sudo systemctl status mj-bot --no-pager
sudo journalctl -u mj-bot -n 50 --no-pager
```

Bir token bilan faqat bitta polling jarayonini ishlating. Kod yoki sozlama
yangilangach `sudo systemctl restart mj-bot` buyrug‘ini bering.

## Mijozga namoyish

Namoyish uchun alohida Telegram bot tokeni va namoyishni olib boruvchi adminning raqamli
Telegram ID sini ishlating. Quyidagi buyruq birinchi ishga tushishda `data/demo/` ichida
bo‘sh baza yaratadi. Keyingi ishga tushirishlarda namoyish yozuvlari shu bazada saqlanadi.
Namoyish rejimida TextUP o‘rniga SMS holati «API kutilmoqda» bo‘ladi.

```powershell
$env:DEMO_BOT_TOKEN = Read-Host -Prompt 'Namoyish botining tokeni' -MaskInput
$env:DEMO_ADMIN_ID = Read-Host -Prompt 'Admin Telegram ID'
.\.venv\Scripts\python.exe main.py --demo
```

Windows PowerShell 5.1 da `-MaskInput` o‘rniga
`$env:DEMO_BOT_TOKEN = [System.Net.NetworkCredential]::new('', (Read-Host 'Namoyish botining tokeni' -AsSecureString)).Password`
buyrug‘idan foydalaning.

Namoyishni admin chatida `/start` bilan boshlang. Xodimni qo‘shing va alohida
sinov akkauntidan ulanishni tasdiqlang. Uchinchi akkauntdan sinov murojaati
yuboring yoki Excel importini ko‘rsating. **Yuborish → Darhol** ni tanlang.
Xodim murojaatni ochib **Bajarildi** bosgach, admin **Holat** bo‘limida natijani ko‘radi.

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
