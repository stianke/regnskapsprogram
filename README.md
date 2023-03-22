# Regnskapsprogram til Filter

Dette er et Python-program som lager et regnskapsdokument med transaksjonene fra et bestemt år, basert på en fil med transaksjoner som er lastet ned fra Sparebanken Sør. Eventuelt kan man velge et eksisterende regnskap, slik at de nye transaksjoner legges til i dette.

## Installasjon
### Last ned Python
For å kjøre programmet må man først laste med Python (lastes ned [herfra](https://www.python.org/downloads)). Hvis du velger "Customize Installation", så huk av for å installere "_pip_" i installasjonen.

Last deretter ned fire Python-biblioteker som brukes av programmet:
* Åpne kommandolinjen (**Windows tast** → skriv **cmd** → **Enter**)
* Kopier inn følgende kommando, og trykk enter `pip install pandas openpyxl pyqt5`

### Last ned selve programmet
Trykk **Code** → **Download Zip**. Høyretrykk **regnskapsprogram.zip** → **Pakk ut alle** → Velg lokasjon hvor du vil lagre programmet.

### Sett opp Python til å åpne programmet
Den letteste måten å kjøre programmet på er å velge Python som applikasjon for `.py` filer. Dette gjøres ved å høyretrykke på `regnskapsprogram` -> Egenskaper -> Åpne med -> Endre -> Flere apper -> Python. Programmet kan nå kjøres ved å dobbeltklikke på `regnskapsprogram` eller en snarvei til `regnskapsprogram`.

## Hvordan bruker programmet
Når programmet kjører må du velge en transaksjonsoversikt eksportert fra Sparebanken Sør. Se avsnittet under for hvordan dette ekspoerteres.
Første gang et regnskap lages hvert år, huker man av for "Opprett nytt regnskap". Da må man oppgi årstall for regnskapet, og velge hvor man skal lagre resultatet. For å legge inn nye transaksjoner i et eksisterende regnskap, huker man av for "Utvid eksisterende regnskap", og velger det eksisterende regnskapet. Dette MÅ være et regnskap som tidligere er generert av `regnskapsprogram` for å være kompitabelt.

## Eksportere transaksjonsoversikt fra Sparebanken Sør
1. Fra hovedsiden i Sparebanken Sør, trykk på 'Transaksjonsoversikt' for den aktuelle bankkontoen. 
2. Søk etter perioden du ønsker å legge inn i regnskapet, f.eks. 01. januar til dags dato.
3. Trykk på ![Eksporter til regneark](https://nettbedriften.evry.com/cpsnbg2/bank/2844/images/excel.gif)-knappen øverst til høyre i søkeresultatene for å laste ned et regneark med transaksjoner.

