import win32com.client
import pandas as pd
import re




#Podłączny excel
Testowy_Dla_Pythone = win32com.client.Dispatch("Excel.Application")
Testowy_Dla_Pythone.Visible = 1

#Ladujemy Excel z komponentami
baza_komponentow = pd.read_excel(r"Baza_wiedzy\Komponenty.xlsx")
# print(baza_komponentow['Ilosc padow'])

##Definiujemy nazwę projektu
#Pobieramy nazwe projektu
Linijka_poczatkowa = str(Testowy_Dla_Pythone.ActiveCell.Row)
# print(Linijka_poczatkowa)
Nazwa_projektu = Testowy_Dla_Pythone.Range('C' + str(Linijka_poczatkowa)).Value
# print(Nazwa_projektu)

#Sprawdzamy granice projektu
#Odliczamy górną granice
licznik = int(Linijka_poczatkowa)
while Testowy_Dla_Pythone.Range('C' + str(licznik)).Value == Nazwa_projektu:
    licznik = licznik - 1
gorna_granica = licznik + 1
#Odliczamy dolną granice
licznik = int(Linijka_poczatkowa)
while Testowy_Dla_Pythone.Range('C' + str(licznik)).Value == Nazwa_projektu:
    licznik = licznik + 1
dolna_granica = licznik - 1
print(f'Granice projektu to {gorna_granica} - {dolna_granica}')

#Zmienne poczatkowe
komponenty_spisok = []
mistakes_spisok = []
mistakes = 0
SMT_statistic = 0
THT_statistic = 0
SMT_pads_statistic = 0
THT_pads_statistic = 0

#Lista obudow standartowych
obudowy_standartowe = []
odnalezione_obudowy = baza_komponentow[baza_komponentow['Ilosc padow'] == 0].index.tolist()
for eleme in odnalezione_obudowy:
    obudowy_standartowe.append(str(baza_komponentow.loc[eleme, 'Obudowa']).replace('-',''))



# obudowy_standartowe = ['SMT', 'THT', 'BGA', 'PGA', 'LGA', 'CSP', 'LCC', 'SON', 'DFN', 'QFN', 'QFP', 'SOP', 'TSSOP' 'SOL', 'SOJ', 'SOM']
#Odpalamy głowny cykl
for stroka in range (gorna_granica, dolna_granica):
    try:
        if Testowy_Dla_Pythone.Range('H' + str(stroka)).Value not in ['DNI', 'dni', 'Dni'] and Testowy_Dla_Pythone.Range('T' + str(stroka)).Value not in ['-', None]: 

            #Zerowanie wartości na poczatku dla biezpieczenstwa
            obudowa_koncowa = '-'
            ilosc_padow = '-'
            ilosc_komponentow = '-'
            typ_komponentu = '-'

            #Pobieranie informacji z bazy danych
            obudowa = Testowy_Dla_Pythone.Range('T' + str(stroka)).Value
            ilosc_komponentow = Testowy_Dla_Pythone.Range('F' + str(stroka)).Value
            obudowa_wzorowa = obudowa

            #Konwertacja dla obudow standartowych
            for el in obudowy_standartowe:
                if el in obudowa: 
                    obudowa = re.sub(r'\d', '', obudowa)
                    if obudowa[-1] != '-': obudowa = obudowa + '-'
            
            #Sprawdzenie w bazie
            odnaleziony_wiersz = baza_komponentow[baza_komponentow['Obudowa'] == obudowa].index.tolist()
            if odnaleziony_wiersz:
                ilosc_padow = str(baza_komponentow.loc[odnaleziony_wiersz, 'Ilosc padow'].values[0])
            typ_komponentu = baza_komponentow.loc[odnaleziony_wiersz, 'Typ'].values[0]
            
            #Dodajemy komponent do wspolnego spisku
            if ilosc_padow == '0':
                # obudowa_koncowa = str(obudowa) + str(obudowa_wzorowa.split('-')[1])
                # ilosc_padow = str(obudowa_wzorowa.split('-')[1])
                obudowa_koncowa = str(obudowa) + str(re.sub(r'\D', '', obudowa_wzorowa))
                ilosc_padow = str(re.sub(r'\D', '', obudowa_wzorowa))
            else: obudowa_koncowa = obudowa

            komponenty_mini_spisok = [[odnaleziony_wiersz[0]], obudowa_koncowa, ilosc_padow, str(int(ilosc_komponentow)), typ_komponentu]
            komponenty_spisok.append(komponenty_mini_spisok)

            #Obliczenie punktow lutownicznych
            if komponenty_mini_spisok[4] == 'SMT': 
                SMT_statistic = SMT_statistic + 1
                SMT_pads_statistic = str(int(SMT_pads_statistic) + int(ilosc_padow) * int(ilosc_komponentow))
            elif komponenty_mini_spisok[4] == 'THT': 
                THT_statistic = THT_statistic + 1
                THT_pads_statistic = str(int(THT_pads_statistic) + int(ilosc_padow) * int(ilosc_komponentow))
    except: 
        #Na wypadek gdy jest błąd
        mistakes = mistakes + 1
        mistakes_spisok.append([[str(stroka)], obudowa_wzorowa])
        
#Podsumowanie
print(f'Ilosc komponentow SMT: {SMT_statistic}')
print(f'Ilosc padow dla elementow SMT: {SMT_pads_statistic}')
print(f'Ilosc komponentow THT: {THT_statistic}')
print(f'Ilosc padow dla elementow THT: {THT_pads_statistic}')
print('--------')
print(f'Lista bledow: {mistakes_spisok}')