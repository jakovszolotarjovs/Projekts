# Projekts
DIP225 projekts

# Problēma
Kādreiz es stradāju nelielā viesnīcā. Kad man bija nakts maiņas, viens no maniem pienākumiem bija izsūtīšana e-pastu klientiem ar dažādām piedavājumiem. Informāciju par klientiem man bija dota Excel datu bāzē. Man bija aizliegts nosūtīt vēstules uzreiz visiem cilvēkiem, jo tas izplatija personisku informāciju par mūsu klientiem. Vēstulei bija jāsatur klienta personisko informāciju tādu kā vārdu un E-pastu. Tas uzdevums prasīja ļoti daudz laika. Es pamēģināju automatizēt šo procesu.

# Projekta uzdevums
Uzrakstīt skriptu, izmantojot Python programmēšanas valodu, kurš lasa informāciju no exel datu bāzes(vārds uzvārds, epasts) un automatiski katrām cilvēkam no datu bāzes(pastāvīgām klientam) ģenerē speciālo(randomu) atlaides kodu. Tas atlaides kods tiek ierakstīts exel tabulā lai administrators varētu parbaudīt vai tas tieši ir mūsu pastāvīgais klients, un vai piedavajums nebija izmantots vairāk par vienu reizi. Pēc tām izmantojot loginu un paroli programma ienāk inbox.lv kontā. Skripts izmantojot cikla operatoru for raksta katram klientam personalizēto e-vēstuli kura satur klienta vārdu, uzvārdu, atlaides kodu un piedavājumu.

# Izmantotas bibliotēkas
Selenium- selenium bibliotēka ļauj man automatiski strādāt ar google chrom vietni. Lasīt, rakstīt un apstrādāt informāciju no tā. Interaktīvi nodarboties ar pogām un t.t

Time - time bibliotēka ļauj man apturēt programmu uz kādu laiku lai dot ieladēties visiem lapas elementiem.

Random - random bibliotēka ļauj man ģenerēt piedavājuma kuponu ar nejaušiem skaitliem un burtiem.

String - string bibliotēka darbojas kopā ar random bibliotēku.

Openpyxl - openpyxl ļauj man strādāt ar exel dokumentu. Lasīt, rakstīt un apkopot informāciju no tā.

# Izmantotas metodes
cikla for operators.

''.join(random.choices(string.ascii_uppercase + string.digits, k=8))-ģenerē piedavājuma kodu.

sheet.cell()-pieraksta kodu exel tabulā.

if operators.

append() - pievieno informāciju mainīgajā "list".

wb.save()- saglābā exel dokumentu.

get() - atver inbox.lv vietni.

find() - ļauj man meklēt informāciju no vietnes, rakstīt informāciju un klikšķināt pogas.

time.sleep()- aptura programmas darbību uz kādu laiku.

driver.switch_to.frame() - parslēdza savu darbību uz Iframe lai es varētu rakstīt un rediģēt tekstu vēstūlē.

# piezīmes
Programmas darbs var atšķirties atkarībā no interneta ātruma, jo dažreiz reklāma lādējās atrāk nēkā programma nosūta vēstuli. Vislabāk programma darbojās jā ir uzstādīts "Adblock", gadijumā ja programma nepabeidza savu darbu korekti lūdzu atkartoti palaist to. E-pasta loginu un paroli jūs varat atrast programmas kodā, tur viss ir nokomentēts. Pēc programma pabeidza savu darbību ir nepieciešams atjaunoti atvert exel dokumentu un parbaudīt jaunos piedavājuma kodus. Saiti uz vidio es pielikšu ortusā. 






