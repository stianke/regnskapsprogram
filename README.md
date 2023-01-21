# Regnskapsprogram til Filter

Python-program som importerer en fil med transaksjoner fra Sparebanken Sør, og lager et regnskapsdokument med transaksjonene fra et bestemt år. Eventuelt kan man velge et eksisterende regnskap og legge nye transaksjoner til i dette.

## Hvordan kjøre
Før programmet kan kjøres må man laste med [Python](https://www.python.org/downloads). Deretter laster man ned fire Python-biblioteker ved å åpne terminalen og kjøre følgende kommandoer:
* pip install easygui
* pip install pandas
* pip install openpyxl
* pip install pyautogui

En måte å kjøre programmet på er å velge Python som applikasjon for `.py` filer. Dette gjøres ved å høyretrykke på `main.py` -> Properties -> Opens With -> Change -> More Apps -> Python.

Programmet kan nå kjøres ved å dobbeltklikke på `main.py`.
* Når programmet kjører blir du bedt om å velge csv-fil eksportert fra Sparebanken Sør. Se seksjon under for hvordan ekspoertere.
* Etter dette kan du velge et eksisterende regnskap for å flette transaksjonene fra Sparebanken Sør inn i. Dette MÅ være et regnskap som tidligere er generert av `main.py` for å være kompitabelt. Dersom du ikke allerede har generert et regnskap for det aktuelle året med `main.py`, trykker du `Cancel`.
* Hvis du trykket `Cancel` i forrige vindu vil det opprettes et helt nytt regnskap. Da får du opp en prompt for å velge hvilket årstall regnskapet skal gjelde for, og en prompt for å velge hvor det skal lagres.


## Eksportere csv-fil fra Sparebanken Sør
Gå til 'Søk i Transaksjoner' og søk etter perioden du ønsker. Deretter trykker du på ![Eksporter til regneark](https://nettbedriften.evry.com/cpsnbg2/bank/2844/images/excel.gif)-knappen øverst til høyre i søkeresultatene for å laste ned et regneark med transaksjoner.


## Endre python-programmet
Om man skal gjøre endringer på programmet anbefales det å bruke PyCharm.
