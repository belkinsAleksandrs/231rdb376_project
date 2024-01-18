# Dzīvokļu datu iegūšanas automatizācija vietnē www.ss.lv
Šis skripts izmanto Selenium un OpenPyXL, lai automatizētu dzīvokļu datu iegūšanas procesu vietnē www.ss.lv, kā arī nosūta rezultātus pa e-pastu.

## Atkarību instalēšana
Instalējiet nepieciešamās bibliotēkas, izpildot šo komandu:

```
pip install selenium openpyxl
```

Pārliecinieties, vai Jums ir instalēta jaunākā Chrome versija. [Chrome lejupielādes saite](https://www.google.com/chrome/?brand=FKPE&gclid=Cj0KCQiAtaOtBhCwARIsAN_x-3JTKE3L7aAKFPHmwO8KK4ExPKFP9WapLClz0bDg1Ueu4-WBZdibfdwaAqiEEALw_wcB&gclsrc=aw.ds#:~:text=the%20Chrome%20installer%3F-,Download,-here)

## Kā izmantot
1. Nomainiet mainīgo vērtības `my_email`, `my_password`, `url`, `file_path`, `max_price`, `number_of_rooms`, `recipient_email` uz saviem.
2. Palaidiet skriptu:
```
python main.py
```
   
## Koda apraksts
- `initialize_driver()`
Funkcija izveido Chrome tīmekļa draiveri, izmantojot Selenium.

- `scrape_data(driver, url, max_price, number_of_rooms)`
Funkcija izmanto Selenium, lai mijiedarbotos ar www.ss.lv vietni un iegūtu dzīvokļu datus. Rezultāts tiek saglabāts kā saraksta saraksts.

- `save_to_excel(data, file_path)`
Funkcija izmanto OpenPyXL, lai saglabātu datus Excel failā.

- `send_email(file_path, my_email, my_password, recipient_email)`
Funkcija nosūta e-pastu ar pievienotu Excel failu, kurā ir apkopoti dati.

## Lietošana
1. Tīmekļa draivera inicializācija.
2. Datu vākšana par dzīvokļiem.
3. Datu saglabāšana Excel failā.
4. E-pasta sūtīšana ar pielikumu - Excel failu.

## Piezīmes
- Ir svarīgi pārliecināties, lai būtu instalēta pareizā jaunākā Chrome versija.
- Ieteicams uzmanīgi izpētīt kodu un pielāgot to savām vajadzībām.
- Drošības nolūkos ieteicams nepaturēt paroli kodā. Izmantojiet vides mainīgos vai konfigurācijas failu.
  
Autors: Aleksandrs Belkins, 231RDB376
Datums: 18.01.2024
