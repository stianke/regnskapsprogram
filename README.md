# Regnskapsprogram til Filter

Dette er et Python-program som lager et regnskapsdokument med transaksjonene fra et bestemt år, basert på en fil med transaksjoner som er lastet ned fra Sparebanken Sør. Eventuelt kan man velge et eksisterende regnskap, slik at de nye transaksjoner legges til i dette.

## Installasjon
### Last ned Python
For å kjøre programmet må man først laste med Python (lastes ned [herfra](https://www.python.org/downloads)). Hvis du velger "Customize Installation", så huk av for å installere "_pip_" og "_td/tk and IDLE_" i installasjonen.

Last deretter ned fire Python-biblioteker som brukes av programmet:
* Åpne kommandolinjen (**Windows tast** → skriv **cmd*** → **Enter**)
* Kopier inn følgende kommando, og trykk enter `pip install easygui pandas openpyxl pyautogui`

### Last ned selve programmet
Trykk **Code** → **Download Zip**. Høyretrykk **regnskapsprogram.zip** → **Pakk ut alle** → Velg lokasjon hvor du vil lagre programmet.

### Sett opp Python til å åpne programmet
Den enkleste måten å kjøre programmet på er å velge Python som applikasjon for `.py` filer. Dette gjøres ved å høyretrykke på `regnskapsprogram.py` -> Egenskaper -> Åpne med -> Endre -> Flere apper -> Python. Programmet kan nå kjøres ved å dobbeltklikke på `regnskapsprogram.py` eller en snarvei til `regnskapsprogram.py`.

## Hvordan bruker programmet
* Når programmet kjører blir du bedt om å velge en transaksjonsoversikt eksportert fra Sparebanken Sør. Se seksjon under for hvordan ekspoertere.
* Etter dette kan du velge et eksisterende regnskap for å flette transaksjonene fra Sparebanken Sør inn i. Dette MÅ være et regnskap som tidligere er generert av `regnskapsprogram.py` for å være kompitabelt. Dersom du ikke allerede har generert et regnskap for det aktuelle året med `regnskapsprogram.py`, trykker du `Cancel`.
* Hvis du trykket `Cancel` i forrige vindu vil det opprettes et helt nytt regnskap. Da får du opp en prompt for å velge hvilket årstall regnskapet skal gjelde for, og en prompt for å velge hvor det skal lagres.


## Eksportere transaksjonsoversikt fra Sparebanken Sør
Gå til 'Transaksjonsoversikt' og søk etter perioden du ønsker. Deretter trykker du på ![Eksporter til regneark](https://nettbedriften.evry.com/cpsnbg2/bank/2844/images/excel.gif)-knappen øverst til høyre i søkeresultatene for å laste ned et regneark med transaksjoner.

