
#Bloczki KivyMD UI
from kivymd.app import MDApp
from kivy.lang import Builder
from kivy.properties import ObjectProperty, StringProperty
from kivy.uix.screenmanager import ScreenManager, Screen
from kivymd.uix.dialog import MDDialog
from kivymd.uix.button import MDFlatButton
from kivymd.uix.card import MDCard
from kivy.clock import Clock
from kivy.core.window import Window

#Bloczki logiki Backendu

#Biblioteki standartowe
import os
import socket
import re
import win32com.client
import pandas as pd
import time
from datetime import datetime
import threading

#POCZĄTKOWE OKNO
class StartWindow(Screen):

    ####Początkowa konfiguracja
    #Rozmiary
    Window.size = (1400, 950)
    print(Window.width)
    print(Window.height)
    Window.left = (Window.width - Window.size[0]/1.2) / 2
    Window.top = (Window.height - Window.size[1]/1.2) / 2



    ####ZMIENNE
    ####Zmienne globalne
    global LZ_type
    LZ_type = 'New'


    ####FUNKCJE 
    #Config changes
    def change_LZ(self, LZ_Type):
        global LZ_type, Przedzial_stat, Licznik_stat, errors_stat
        LZ_type = LZ_Type
        if LZ_type == 'Old':
            Przedzial_stat = 'BE'
            Licznik_stat = 'BD'
            errors_stat = 'BF'
        else:
            Przedzial_stat = 'BN'
            Licznik_stat = 'BM'
            errors_stat = 'BO'

    def Pokaz_instarukcje(self):
        pass

    def Pokaz_liste_montowni(self):
        pass



#OKONO DLA PRACY Z KOMPONENTAMI
class WorkWithComponentsWindow(Screen):

    ####Zmienne z pliku .kv
    edytujbutton = ObjectProperty(None)
    addkomentarz = ObjectProperty(None)
    errorstroki = ObjectProperty(None)
    goodstroki = ObjectProperty(None)
    loadingspinner = ObjectProperty(None)
    ogolneklient = ObjectProperty(None)
    ogolnesztuk = ObjectProperty(None)
    ogolnenazwa = ObjectProperty(None)
    statSMT = ObjectProperty(None)
    statSMTpady = ObjectProperty(None)
    statTHT = ObjectProperty(None)
    statTHTpady = ObjectProperty(None)   
    iloscwarstw1 = ObjectProperty(None) 
    iloscwarstw2 = ObjectProperty(None) 
    uwagimycie = ObjectProperty(None) 
    uwagilakierowanie = ObjectProperty(None) 
    uwagiaoi = ObjectProperty(None) 
    uwagirtg = ObjectProperty(None) 
    uwagiprogramowanie = ObjectProperty(None) 
    uwagiszlifowanie = ObjectProperty(None) 
    uwagiipc = ObjectProperty(None) 
    uwagitestowanie = ObjectProperty(None) 
    

    # V 1.0
    # nazwaprojektu = StringProperty()
    # graniceprojektu = StringProperty()
    # smtkomponenty = StringProperty()
    # smtkomponentypady = StringProperty()
    # thtkomponenty = StringProperty()
    # thtkomponentypady = StringProperty()
    # rezystoryilosc = StringProperty()
    # kondensatoryilosc = StringProperty()
    # rezystorykondensatoryrazem = StringProperty()
    # obudowysmtilosc = StringProperty()
    # obudowythtilosc = StringProperty()
    # obudowyrazemilosc = StringProperty()
    # ilosckomponentowbazowa = StringProperty()
    # dniilosc = StringProperty()
    # errorsilosc = StringProperty()
    # uwagizprojektu = StringProperty()
    # peregorodka = StringProperty()

    # V 2.0
    pliklz = StringProperty()
    datawypelnienia = StringProperty()
    nazwaprojektu = StringProperty()
    graniceprojektu = StringProperty()
    smtkomponenty = StringProperty()
    smtkomponentypady = StringProperty()
    thtkomponenty = StringProperty()
    thtkomponentypady = StringProperty()
    rezystoryilosc = StringProperty()
    kondensatoryilosc = StringProperty()
    rezystorykondensatoryrazem = StringProperty()
    obudowysmtilosc = StringProperty()
    obudowythtilosc = StringProperty()
    obudowyrazemilosc = StringProperty()
    ilosckomponentowbazowa = StringProperty()
    dniilosc = StringProperty()
    statystykablendy = StringProperty()
    errorsilosc = StringProperty()
    uwagizprojektu = StringProperty()
    peregorodka = StringProperty()
    


    ####Zmienne globalne
    block = True
    ilosc_warstw = '0'
    czy_recznie = 'Automatycznie'
    UwagiCheckBox = []
    UwagiStatystyka = []
    dialog = None


    ####FUNKCJE 

    ###Wyciągnij komponenty
    ####Funkcje pomocnicze do obsługi UI
    def Wyciagnij_komponenty_potok(self):
        thread1 = threading.Thread(target=self.Wyciagnij_komponenty)
        thread1.start()
        
    def update_loading_spinner(self,wartosc):
        def update_spinner(dt):
            self.loadingspinner.opacity = wartosc
        Clock.schedule_once(update_spinner, 0)
    
    def change_ilosc_warstw(self, ile):
        self.ilosc_warstw = ile
    
    def change_klient_MDTekstFrield_on_focus(self):
        if self.ogolnenazwa.text == '<Wpisz>':
            self.ogolnenazwa.text = ''
    
    def add_UwagiCheckBox(self, element):
        #Dodajemy/Usuwamy element z listy
        if element not in self.UwagiCheckBox:
            self.UwagiCheckBox.append(element)
        elif element in self.UwagiCheckBox:
            self.UwagiCheckBox.remove(element)


    def create_good_stroka(self, part_good):
        dobry_komponent = StrokaGood(info=part_good)
        self.goodstroki.add_widget(dobry_komponent)
    
    def create_error_stroka(self, part_error):
        error_komponent = StrokaError(info=part_error)
        self.errorstroki.add_widget(error_komponent)
    
    def clear_good_stroki(self, *args):
        self.goodstroki.clear_widgets()
    
    def clear_error_stroki(self, *args):
        self.errorstroki.clear_widgets()
    
    def update_parametrs(self, *args):
        # ObjectProperties()
        self.ogolneklient.text = self.ogolneklient1
        self.ogolnenazwa.text = self.ogolnenazwa1
        self.ogolnesztuk.text = self.sztuki1
        self.statSMT.text = self.statSMT1
        self.statSMTpady.text = self.statSMTpady1
        self.statTHT.text = self.statTHT1
        self.statTHTpady.text = self.statTHTpady1
        self.iloscwarstw1.active = False
        self.iloscwarstw2.active = False
        self.uwagimycie.active = False
        self.uwagilakierowanie.active = False
        self.uwagiaoi.active = False
        self.uwagirtg.active = False
        self.uwagiprogramowanie.active = False
        self.uwagiszlifowanie.active = False
        self.uwagiipc.active = False
        self.uwagitestowanie.active = False
        self.UwagiCheckBox = []
        self.UwagiStatystyka = []
        # StringProperties()
        self.nazwaprojektu = self.Kod_projektu
        self.graniceprojektu = self.graniceprojektu1
        self.pliklz = self.pliklz1
        self.datawypelnienia = self.data1
        self.rezystoryilosc = self.rezystoryilosc1
        self.kondensatoryilosc = self.kondensatoryilosc1
        self.rezystorykondensatoryrazem = self.rezystorykondensatoryrazem1
        self.obudowysmtilosc = self.obudowysmtilosc1
        self.obudowythtilosc = self.obudowythtilosc1
        self.obudowyrazemilosc = self.obudowyrazemilosc1
        self.ilosckomponentowbazowa = self.ilosckomponentowbazowa1
        self.dniilosc = self.dniilosc1
        self.peregorodka = '-----------------------'
        self.errorsilosc = self.errorsilosc1
        self.statystykablendy = self.statystykablendy1

    ####Funkcja głowna
    def Wyciagnij_komponenty(self):

        #Przygotowanie
        Excel = win32com.client.Dispatch("Excel.Application")
        Excel.Visible = 1
        Adres_linijki = Excel.ActiveCell.Row
        print(Adres_linijki)
        
        #Włączamy loading spinner
        self.update_loading_spinner(1)

        #Sprzątamy po ostatnim użyciu
        Clock.schedule_once(self.clear_good_stroki, 0)
        Clock.schedule_once(self.clear_error_stroki, 0)
        self.UwagiCheckBox = []
        self.UwagiStatystyka = []
        self.nazwaprojektu = f''
        self.pliklz = f''
        self.data = f''
        self.datawypelnienia = f''
        self.graniceprojektu = f''
        self.smtkomponenty = f''
        self.smtkomponentypady = f''
        self.thtkomponenty = f''
        self.thtkomponentypady = f''
        self.rezystoryilosc = f''
        self.kondensatoryilosc = f''
        self.rezystorykondensatoryrazem = f''
        self.obudowysmtilosc = f''
        self.obudowythtilosc = f''
        self.obudowyrazemilosc = f''
        self.ilosckomponentowbazowa = f''
        self.dniilosc = f''
        self.errorsilosc = f''
        self.statystykablendy = f''
        self.uwagizprojektu = f''
        self.peregorodka = ''
        self.czy_recznie = 'Automatycznie'

        #Zerujemy


        

        #Definicja Zmiennych w zalezności od Exelu
        if LZ_type == 'Old':
            self.status_kol = 'A'
            self.part_ID_kol = 'B'
            self.Project_kol = 'C'
            self.Qty_of_PCB_kol = 'D'
            self.String_nr_kol = 'E'
            self.Qty_per_PCB_kol = 'F'
            self.Designators_kol = 'G'
            self.DNP_kol = 'H'
            self.Input_Value_kol = 'I'
            self.Input_Description_kol = 'J'
            self.Input_Footprint_kol = 'K'
            self.Input_MFG_kol = 'L'
            self.Input_SKU_kol = 'M'
            self.Input_comments_kol = 'N'
            self.Qty_to_mint_kol = 'O'
            self.Tech_zapas_kol = 'P'
            self.Qty_to_buy_kol = 'Q'
            self.Output_value_kol = 'R'
            self.Output_description_kol = 'S'
            self.Output_footprints_kol = 'T'
            self.Part_type_kol = 'U'
            self.SMD_THT_kol = 'V'
            self.Vendor_A_kol = 'W'
            self.MFG_A_kol = 'X'
            self.MFG_A_part_N_kol = 'Y'
            self.Vendor_A_SKU_kol = 'Z'
            self.Status_kol = 'AA'
            self.A_MQ_1_kol = 'AB'
            self.A_Uprice_1_PLN_kol = 'AC'
            self.A_Uprice_1_USD_kol = 'AD'
            self.A_Uprice_1_EUR_kol = 'AE'
            self.A_Uprice_1_RMB_kol = 'AF'
            self.A_MQ_2_kol = 'AG'
            self.A_Uprice_2_PLN_kol = 'AH'
            self.A_Uprice_2_USD_kol = 'AI'
            self.A_Uprice_2_EUR_kol = 'AJ'
            self.A_Uprice_2_RMB_kol = 'AK'
            self.Sub_A1_PLN_kol = 'AL'
            self.Sub_A2_PLN_kol = 'AM'
            self.B_MQ_1_kol = 'AN'
            self.B_Uprice_1_PLN_kol = 'AO'
            self.B_Uprice_1_USD_kol = 'AP'
            self.B_Uprice_1_EUR_kol = 'AQ'
            self.B_Uprice_1_RMB_kol = 'AR'
            self.B_MQ_2_kol = 'AS'
            self.B_Uprice_2_PLN_kol = 'AT'
            self.B_Uprice_2_USD_kol = 'AU'
            self.B_Uprice_2_EUR_kol = 'AV'
            self.B_Uprice_2_RMB_kol = 'AW'
            self.Sub_B1_PLN_kol = 'AX'
            self.Sub_B2_PLN_kol = 'AY'
            self.Vendor_B_kol = 'AZ'
            self.MFG_B_kol = 'BA'
            self.MFG_Part_N_kol = 'BB'
            self.Vendor_B_SKU_kol = 'BC'
            self.Przedzial_stat = 'BE'
            self.Licznik_stat = 'BD'
            self.errors_stat = 'BF'

        if LZ_type == 'New':
            self.status_kol = 'A'
            self.part_ID_kol = 'B'
            self.Project_kol = 'C'
            self.Qty_of_PCB_kol = 'D'
            self.String_nr_kol = 'E'
            self.Qty_per_PCB_kol = 'F'
            self.Designators_kol = 'G'
            self.DNP_kol = 'H'
            self.Input_Value_kol = 'I'
            self.Input_Description_kol = 'J'
            self.Input_Footprint_kol = 'K'
            self.Input_MFG_kol = 'L'
            self.Input_SKU_kol = 'M'
            self.URL_kol = 'N'
            self.Input_comments_kol = 'O'
            self.Notes_for_assembly_kol = 'P'
            self.Alternative_parts_confirmed_kol = 'Q'
            self.Alternative_parts_comments_kol = 'R'
            self.Qty_to_mint_kol = 'S'
            self.Tech_zapas_kol = 'T'
            self.Qty_to_buy_kol = 'U'
            self.Output_value_kol = 'V'
            self.Output_description_kol = 'W'
            self.Output_footprints_kol = 'X'
            self.Part_type_kol = 'Y'
            self.SMD_THT_kol = 'Z'
            self.Vendor_A_kol = 'AA'
            self.MFG_A_kol = 'AB'
            self.MFG_A_part_N_kol = 'AC'
            self.Vendor_A_SKU_kol = 'AD'
            self.Status_kol = 'AE'
            self.Order_number = 'AF'
            self.A_MQ_1_kol = 'AG'
            self.A_Uprice_1_PLN_kol = 'AH'
            self.A_Uprice_1_USD_kol = 'AI'
            self.A_Uprice_1_EUR_kol = 'AJ'
            self.A_Uprice_1_RMB_kol = 'AK'
            self.A_MQ_2_kol = 'AL'
            self.A_Uprice_2_PLN_kol = 'AM'
            self.A_Uprice_2_USD_kol = 'AN'
            self.A_Uprice_2_EUR_kol = 'AO'
            self.A_Uprice_2_RMB_kol = 'AP'
            self.Sub_A1_PLN_kol = 'AQ'
            self.Sub_A2_PLN_kol = 'AR'
            self.B_MQ_1_kol = 'AS'
            self.B_Uprice_1_PLN_kol = 'AT'
            self.B_Uprice_1_USD_kol = 'AU'
            self.B_Uprice_1_EUR_kol = 'AV'
            self.B_Uprice_1_RMB_kol = 'AW'
            self.B_MQ_2_kol = 'AX'
            self.B_Uprice_2_PLN_kol = 'AY'
            self.B_Uprice_2_USD_kol = 'AZ'
            self.B_Uprice_2_EUR_kol = 'BA'
            self.B_Uprice_2_RMB_kol = 'BB'
            self.Sub_B1_PLN_kol = 'BC'
            self.Sub_B2_PLN_kol = 'BD'
            self.Vendor_B_kol = 'BE'
            self.MFG_B_kol = 'BF'
            self.MFG_Part_N_kol = 'BG'
            self.Vendor_B_SKU_kol = 'BH'
            self.Przedzial_stat = 'BN'
            self.Licznik_stat = 'BM'
            self.errors_stat = 'BO'

        ##Sprawdzamy adres projektu
        #Ladujemy Excel z komponentami
        # baza_komponentow = pd.read_excel(r"Baza_wiedzy\Komponenty.xlsx")
        baza_komponentow = pd.read_excel(r"\\gcl-ne-fs-1\Shared Files Office\Temp\Baza_wiedzy\Komponenty.xlsx")
        # print(baza_komponentow['Ilosc padow'])

        ##Definiujemy nazwę projektu
        #Pobieramy nazwe projektu
        Linijka_poczatkowa = str(Excel.ActiveCell.Row)
        # print(Linijka_poczatkowa)
        Kod_projektu = Excel.Range(self.Project_kol + str(Linijka_poczatkowa)).Value
        self.Kod_projektu = Kod_projektu
        # print(Kod_projektu)

        #Sprawdzamy granice projektu
        #Odliczamy górną granice
        licznik = int(Linijka_poczatkowa)
        czy_juz_status = ''
        while Excel.Range(self.Project_kol + str(licznik)).Value == Kod_projektu and 'STAT' not in str(czy_juz_status):
            licznik = licznik - 1
            czy_juz_status = Excel.Range(self.status_kol + str(licznik)).Value
        gorna_granica = licznik + 1
        #Odliczamy dolną granice
        licznik = int(Linijka_poczatkowa)
        czy_juz_status = ''
        while Excel.Range(self.Project_kol + str(licznik)).Value == Kod_projektu and 'STAT' not in str(czy_juz_status):
            licznik = licznik + 1
            czy_juz_status = Excel.Range(self.status_kol + str(licznik)).Value
        dolna_granica = licznik
        self.graniceprojektu1 = f'{gorna_granica} - {dolna_granica}'
        print(f'Granice projektu to {gorna_granica} - {dolna_granica}')



        #Zmienne poczatkowe
        komponenty_spisok = []
        mistakes_spisok = []
        mistakes = 0
        mistakes_ilosc_komponentow = 0
        wskaznik_poprawnosci = 'git'
        SMT_statistic = 0
        THT_statistic = 0
        SMT_pads_statistic = 0
        THT_pads_statistic = 0
        rezystory_kondensatory = 0
        rezystory_ilosc = 0
        kondensatory_ilosc = 0
        DNI = 0
        SMT_statistic_pojedynczy = 0
        THT_statistic_pojedynczy = 0
        SMT_THT_statistic_pojedynczy = 0
        ilosc_komponentow_bazowych = 0

        #Lista obudow standartowych (pandas)
        obudowy_standartowe = []
        odnalezione_obudowy = baza_komponentow[baza_komponentow['Ilosc padow'] == 0].index.tolist()
        for eleme in odnalezione_obudowy:
            obudowy_standartowe.append(str(baza_komponentow.loc[eleme, 'Obudowa']).replace('-',''))



        # obudowy_standartowe = ['SMT', 'THT', 'BGA', 'PGA', 'LGA', 'CSP', 'LCC', 'SON', 'DFN', 'QFN', 'QFP', 'SOP', 'TSSOP' 'SOL', 'SOJ', 'SOM']
        ##Odpalamy głowny cykl
        for stroka in range (gorna_granica, dolna_granica):
            try:
                ilosc_komponentow_bazowych = ilosc_komponentow_bazowych + int(Excel.Range(self.Qty_per_PCB_kol + str(stroka)).Value)
                if Excel.Range(self.DNP_kol + str(stroka)).Value not in ['DNI', 'dni', 'Dni'] and Excel.Range(self.Output_footprints_kol + str(stroka)).Value not in ['-', None]: 

                    #Zerowanie wartości na poczatku dla biezpieczenstwa
                    obudowa_koncowa = '-'
                    ilosc_padow = '-'
                    ilosc_komponentow = '-'
                    typ_komponentu = '-'
                    typ_obudowy = '-'
                    THT_SMT = '-'


                    #Pobieranie informacji z bazy danych
                    obudowa = Excel.Range(self.Output_footprints_kol + str(stroka)).Value
                    typ_obudowy = Excel.Range(self.Part_type_kol + str(stroka)).Value
                    THT_SMT = Excel.Range(self.SMD_THT_kol + str(stroka)).Value
                    ilosc_komponentow = Excel.Range(self.Qty_per_PCB_kol + str(stroka)).Value
                    obudowa_wzorowa = obudowa

                    #Konwertacja dla obudow standartowych
                    for el in obudowy_standartowe:
                        if el in obudowa: 
                            obudowa = re.sub(r'\d', '', obudowa)
                            if obudowa[-1] != '-': obudowa = obudowa + '-'
                    
                    #Sprawdzenie w bazie (pandas)
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

                    komponenty_mini_spisok = [[stroka, odnaleziony_wiersz[0]], obudowa_koncowa, ilosc_padow, str(int(ilosc_komponentow)), typ_komponentu, typ_obudowy]
                    komponenty_spisok.append(komponenty_mini_spisok)

                    #Obliczenie ilości rezystorów/kondensatorów
                    if typ_obudowy in ['Resistor', 'Capacitor'] and THT_SMT in ['SMT', 'SMD']:
                        if typ_obudowy == 'Resistor': rezystory_ilosc = str(int(rezystory_ilosc) + 1 * int(ilosc_komponentow))
                        if typ_obudowy == 'Capacitor': kondensatory_ilosc = str(int(kondensatory_ilosc) + 1 * int(ilosc_komponentow))
                        

                    #Obliczenie punktow lutownicznych
                    if komponenty_mini_spisok[4] == 'SMT': 
                        SMT_statistic_pojedynczy = SMT_statistic_pojedynczy + 1
                        SMT_statistic = SMT_statistic + 1 * int(ilosc_komponentow)
                        SMT_pads_statistic = str(int(SMT_pads_statistic) + int(ilosc_padow) * int(ilosc_komponentow))
                    elif komponenty_mini_spisok[4] == 'THT': 
                        THT_statistic_pojedynczy = THT_statistic_pojedynczy + 1
                        THT_statistic = THT_statistic + 1 * int(ilosc_komponentow)
                        THT_pads_statistic = str(int(THT_pads_statistic) + int(ilosc_padow) * int(ilosc_komponentow))
                    
                    #Bloczek ze statystyką do uwag technicznych
                    def Dodaj_do_spisku(poczatek_obudowy, sztuk):

                        wartosc = [poczatek_obudowy, sztuk]
                        indikator = 0
                        #Bloczek dodania 1 sztuki
                        for ele in self.UwagiStatystyka:
                            if ele[0] == wartosc[0]: indikator = 1
                        if indikator == 0: self.UwagiStatystyka.append(wartosc)
                        
                        #Bloczek dodania do juz istniejacych
                        else:
                            a = 0
                            for el in self.UwagiStatystyka:
                                if el[0] == poczatek_obudowy:
                                    self.UwagiStatystyka[a][1] = self.UwagiStatystyka[a][1] + 1 * sztuk
                                a = a+1

                    
                    #Obudowy
                    poczatek_obudowy = obudowa_wzorowa.split('-')[0]
                    if poczatek_obudowy in ['BGA','QFH','PGA','LGA','QFA','TQFA','LQFA', 'QFN']:
                        Dodaj_do_spisku(poczatek_obudowy, int(ilosc_komponentow))
                    
                    #Rezystory/Kondensatory
                    if str(obudowa_wzorowa) in ['0402','0201','01005'] and typ_obudowy == 'Resistor':
                        Dodaj_do_spisku(f'R{obudowa_wzorowa}', int(ilosc_komponentow))
                    if str(obudowa_wzorowa) in ['0402','0201','01005'] and typ_obudowy == 'Capacitor':
                        Dodaj_do_spisku(f'C{obudowa_wzorowa}', int(ilosc_komponentow))

                    
                else: 
                    ilosc_komponentow_bazowych = ilosc_komponentow_bazowych - int(Excel.Range(self.Qty_per_PCB_kol + str(stroka)).Value)
                    DNI = DNI + 1

            except: 
                #Na wypadek gdy jest błąd
                mistakes = mistakes + 1
                mistakes_spisok.append([[str(stroka)], obudowa_wzorowa, str(int(ilosc_komponentow))])
                try: mistakes_ilosc_komponentow = mistakes_ilosc_komponentow + 1 * int(ilosc_komponentow)
                except: wskaznik_poprawnosci = 'zle'


        #Podsumowanie ilości
        rezystory_kondensatory = int(rezystory_ilosc) + int(kondensatory_ilosc)
        SMT_THT_statistic_pojedynczy = SMT_statistic_pojedynczy + THT_statistic_pojedynczy

        #Podsumowanie przed wyświetleniem wartości do UI
        
        

        print('Uwagi: ', self.UwagiStatystyka)
        print(f'Kondensatory: {kondensatory_ilosc}')
        print(f'Rezystory: {rezystory_ilosc}')
        print(f'Rezystory i kondensatory: {rezystory_kondensatory}')

        #Podsumowanie
        print(f'Ilosc komponentow SMT: {SMT_statistic}')
        print(f'Ilosc padow dla elementow SMT: {SMT_pads_statistic}')
        print(f'Ilosc komponentow THT: {THT_statistic}')
        print(f'Ilosc padow dla elementow THT: {THT_pads_statistic}')
        print('--------')
        print(f'Lista bledow: {mistakes_spisok}')

        #Pokazujemy użytkownikowi statystykę
        # V 1.0
        # self.nazwaprojektu = f'Nazwa projektu: {Kod_projektu}'
        # self.graniceprojektu = f'Granice projektu: {gorna_granica} - {dolna_granica}'
        # self.smtkomponenty = f'Suma komponentów SMT: {SMT_statistic}'
        # self.smtkomponentypady = f'Suma padów dla elementów SMT: {SMT_pads_statistic}'
        # self.thtkomponenty = f'Suma komponentów THT: {THT_statistic}'
        # self.thtkomponentypady = f'Suma padów dla elementów THT: {THT_pads_statistic}'
        # self.rezystoryilosc = f'Suma SMT rezystorów: {rezystory_ilosc}'
        # self.kondensatoryilosc = f'Suma SMT kondensatorów: {kondensatory_ilosc}'
        # self.rezystorykondensatoryrazem = f'Suma pasywnych SMT: {rezystory_kondensatory}'
        # self.obudowysmtilosc = f'Ilość rodzaji obudów SMT: {SMT_statistic_pojedynczy}'
        # self.obudowythtilosc = f'Ilość rodzaji obudów THT: {THT_statistic_pojedynczy}' 
        # self.obudowyrazemilosc = f'Ilość różnych obudów: {SMT_THT_statistic_pojedynczy}'
        # self.ilosckomponentowbazowa = f'Domyślna liczba komponentów: {ilosc_komponentow_bazowych}'
        # self.dniilosc = f'Ilość komponentów DNI: {DNI}'
        # self.uwagizprojektu = '???'
        # self.peregorodka = '-----------------------'
        # if wskaznik_poprawnosci == 'git': self.errorsilosc = f'Ilość błędow: {len(mistakes_spisok)} ({mistakes_ilosc_komponentow} w sumie)'
        # else: self.errorsilosc = f'Ilość błędow: {len(mistakes_spisok)}'

        # V 2.0
        self.data1 = str(datetime.now().strftime("%d.%m.%Y"))
        self.imieuzytkownika1 = socket.gethostname()
        self.pliklz1 = str(Excel.ActiveWorkbook.Name)
        self.ogolneklient1 = str(Excel.ActiveWorkbook.Name).replace('_','-').split('-LZ')[0]
        self.ogolnenazwa1 = '<Wpisz>'
        try:self.sztuki1 = str(int(Excel.Range(self.Qty_of_PCB_kol + str(dolna_granica)).Value))
        except: self.sztuki1 = str(Excel.Range(self.Qty_of_PCB_kol + str(dolna_granica)).Value).replace('/0','')
        self.statSMT1 = str(SMT_statistic)
        self.statSMTpady1 = str(SMT_pads_statistic)
        self.statTHT1 = str(THT_statistic)
        self.statTHTpady1 = str(THT_pads_statistic)
        self.rezystoryilosc1 = f'Suma SMT rezystorów: {rezystory_ilosc}'
        self.kondensatoryilosc1 = f'Suma SMT kondensatorów: {kondensatory_ilosc}'
        self.rezystorykondensatoryrazem1 = f'Suma pasywnych SMT: {rezystory_kondensatory}'
        self.obudowysmtilosc1 = f'Ilość rodzaji obudów SMT: {SMT_statistic_pojedynczy}'
        self.obudowythtilosc1 = f'Ilość rodzaji obudów THT: {THT_statistic_pojedynczy}' 
        self.obudowyrazemilosc1 = f'Ilość różnych obudów: {SMT_THT_statistic_pojedynczy}'
        self.ilosckomponentowbazowa1 = f'Domyślna liczba komponentów: {ilosc_komponentow_bazowych}'
        self.dniilosc1 = f'Ilość komponentów DNI: {DNI}'
        self.peregorodka = '-----------------------'
        if wskaznik_poprawnosci == 'git': self.errorsilosc1 = f'Ilość błędow: {len(mistakes_spisok)} ({mistakes_ilosc_komponentow} w sumie)'
        else: self.errorsilosc1 = f'Ilość błędow: {len(mistakes_spisok)}'
        self.statystykablendy1 = ''
        self.UwagiStatystyka.sort()
        for elmen in self.UwagiStatystyka:
            self.statystykablendy1 = self.statystykablendy1 + f'{elmen[0]} ({elmen[1]}); '
        self.ilosc_warstw = '0'
        
        #Updatujemy UI
        Clock.schedule_once(self.update_parametrs, 0)
        
        #Generujemy karteczki do UI
        for part_good in komponenty_spisok:
            Clock.schedule_once(lambda dt, part_good=part_good: self.create_good_stroka(part_good), 0)

        for part_error in mistakes_spisok:
            Clock.schedule_once(lambda dt, part_error=part_error: self.create_error_stroka(part_error), 0)
        
        #Wyłączamy loading spinner
        self.update_loading_spinner(0)
        
    

    ###Edytuj_wyniki
    def Edytuj_wyniki(self):
        # Zmieniamy readonly w polach tekstowych na false
        Clock.schedule_once(self.Zmien_stan_pol_tekstowych, 0)
        
    
    def Zmien_stan_pol_tekstowych(self, dt):
        self.czy_recznie = 'Ręcznie'

        if self.block == True:
            #change text button color
            self.edytujbutton.text_color = (0,1,0,1)
            self.edytujbutton.line_color = (0,1,0,1)
            #readonly
            self.ogolnesztuk.readonly = False
            self.statSMT.readonly = False
            self.statSMTpady.readonly = False
            self.statTHT.readonly = False
            self.statTHTpady.readonly = False
            #color
            self.ogolnesztuk.text_color_normal = (0,1,0,1)
            self.statSMT.text_color_normal = (0,1,0,1)
            self.statSMTpady.text_color_normal = (0,1,0,1)
            self.statTHT.text_color_normal = (0,1,0,1)
            self.statTHTpady.text_color_normal = (0,1,0,1)

            self.block = False

        elif self.block == False:
            #change text button color
            self.edytujbutton.text_color = (1,1,1,1)
            self.edytujbutton.line_color = (1,1,1,0.5)
            #readonly
            self.ogolnesztuk.readonly = True
            self.statSMT.readonly = True
            self.statSMTpady.readonly = True
            self.statTHT.readonly = True
            self.statTHTpady.readonly = True
            #color
            self.ogolnesztuk.text_color_normal = (1,1,1,0.5)
            self.statSMT.text_color_normal = (1,1,1,0.5)
            self.statSMTpady.text_color_normal = (1,1,1,0.5)
            self.statTHT.text_color_normal = (1,1,1,0.5)
            self.statTHTpady.text_color_normal = (1,1,1,0.5)

            self.block = True
        

        ###Exportuj do Excela
    def Exportuj_do_Excela(self):
        # Zmieniamy readonly w polach tekstowych na false
        Clock.schedule_once(self.Exportuj_do_Excela_funk, 0)

    
    def Exportuj_do_Excela_funk(self, dt):
        ##Obsługa błędów
        self.dialog = None
        if self.statSMT.text == '':
            self.Okienko_informacyjne('Nie pobrałeś dane o płycie, pobierz je i spróbuj jeszcze raz.')
            return
        elif self.ogolnenazwa.text in ['<Wpisz>', '']:
            self.Okienko_informacyjne('Nie wpisałeś nazwę klienta. Wpisz ją i spróbuj jeszcze raz.')
            return
        elif self.ilosc_warstw == '0':
            self.Okienko_informacyjne('Nie wybrałeś ilości warstw montażu, postaw ptaszkę w odpowiednim miejscu i spróbuj jeszcze raz.')
            return
        
        ##Sprawdzamy nazwę Exceli
        Excel = win32com.client.Dispatch("Excel.Application")
        Excel.Visible = 1
        NewName = str(Excel.ActiveWorkbook.Name)
        if NewName == self.pliklz1:
            self.Okienko_informacyjne('Jesteś w złym Excelu. Kliknij w Excelu "MontazIN_PL" i spróbuj jeszcze raz.')
            return
        
        ##Wpisujemy Dane
        last_row = Excel.ActiveWorkbook.ActiveSheet.Cells(Excel.ActiveWorkbook.ActiveSheet.Rows.Count, 1).End(-4162).Row
        Stroka_dla_danych = str(int(last_row) + 1)

        ##Sprawdzamy czy ten kod juz jest w bazie
        znalezione_linijki = []
        poszukiwany_kod = self.Kod_projektu.replace('_rev1','')
        for i in range(1, last_row + 1):
            try:
                if Excel.ActiveWorkbook.ActiveSheet.Cells(i, 3).Value.replace('_rev1','') == poszukiwany_kod: 
                    znalezione_linijki.append(i)
            except: pass
        #list to string
        linijki = ''
        for elik in znalezione_linijki: 
            linijki = linijki + str(elik) + '; '
        try:
            if linijki[-2:] == '; ':linijki = linijki[:-2]
        except: pass

        #Wyswietlamy komunikat że kod juz jest
        if znalezione_linijki != []: 

            # zmienna = f'W linijkach: {linijki} już jest ten projekt. \n\nCo chcesz zrobić dalej?'
            # wybor_klienta = self.Okienko_informacyjne_rozszerzone(zmienna)

            zmienna = f'W linijkach: {linijki} już jest ten projekt. Usuń go i działaj dalej.'
            self.Okienko_informacyjne(zmienna)
            return

        #V1.0 Zmienne
        Data = 'A'
        Klient = 'B'
        Kod_Projektu = 'C'
        Nazwa_Projektu = 'D'
        Ilosc_do_wyceny = 'E'
        Elementy_SMD = 'H'
        Pady_SMD = 'I'
        Elementy_THT = 'J'
        Pady_THT = 'K'
        Ilosci_stron_montazu = 'L'
        Uwagi_techniczne = 'M'
        Notatki = 'N'
        Odpowiedzialna_osoba = 'O'
        Automatyczne_lub_Recznie = 'P'

        #Kształtowanie danyc do wpisu
        Data_wpis = self.data1
        Klient_wpis = self.ogolneklient.text
        Kod_Projektu_wpis = self.Kod_projektu
        Nazwa_Projektu_wpis = self.ogolnenazwa.text
        Ilosc_do_wyceny_wpis = self.ogolnesztuk.text
        Elementy_SMD_wpis = self.statSMT.text
        Pady_SMD_wpis = self.statSMTpady.text
        Elementy_THT_wpis = self.statTHT.text
        Pady_THT_wpis = self.statTHTpady.text
        Ilosci_stron_montazu_wpis = self.ilosc_warstw
        Notatki_wpis = self.addkomentarz.text
        Odpowiedzialna_osoba_wpis = self.imieuzytkownika1
        Automatyczne_lub_Recznie_wpis = self.czy_recznie
        Uwagi_techniczne_wpis = ''
        print('UwagiCheckBox: ', self.UwagiCheckBox)
        print('UwagiStatystyka: ', self.UwagiStatystyka)
        self.UwagiCheckBox.sort()
        for el in self.UwagiCheckBox:
            Uwagi_techniczne_wpis = Uwagi_techniczne_wpis + el + '; '
        Uwagi_techniczne_wpis = Uwagi_techniczne_wpis + self.statystykablendy
        
        

        #Bezpośrednie wpisanie do tablicy
        Excel.Range(Data + Stroka_dla_danych).Value = Data_wpis
        Excel.Range(Klient + Stroka_dla_danych).Value = Klient_wpis
        Excel.Range(Kod_Projektu + Stroka_dla_danych).Value = Kod_Projektu_wpis
        Excel.Range(Nazwa_Projektu + Stroka_dla_danych).Value = Nazwa_Projektu_wpis
        Excel.Range(Ilosc_do_wyceny + Stroka_dla_danych).Value = Ilosc_do_wyceny_wpis
        Excel.Range(Elementy_SMD + Stroka_dla_danych).Value = Elementy_SMD_wpis
        Excel.Range(Pady_SMD + Stroka_dla_danych).Value = Pady_SMD_wpis
        Excel.Range(Elementy_THT + Stroka_dla_danych).Value = Elementy_THT_wpis
        Excel.Range(Pady_THT + Stroka_dla_danych).Value = Pady_THT_wpis
        Excel.Range(Ilosci_stron_montazu + Stroka_dla_danych).Value = Ilosci_stron_montazu_wpis
        Excel.Range(Uwagi_techniczne + Stroka_dla_danych).Value = Uwagi_techniczne_wpis
        Excel.Range(Notatki + Stroka_dla_danych).Value = Notatki_wpis
        Excel.Range(Odpowiedzialna_osoba + Stroka_dla_danych).Value = Odpowiedzialna_osoba_wpis
        Excel.Range(Automatyczne_lub_Recznie + Stroka_dla_danych).Value = Automatyczne_lub_Recznie_wpis



        


    
    def Okienko_informacyjne(self, info):
        if not self.dialog:
            self.dialog = MDDialog(
                text=info,
                buttons=[
                    MDFlatButton(
                        text="Ok",
                        on_release = self.close_dialog
                    ),
                ],
            )
        self.dialog.open()
        return 'ok'

    def Okienko_informacyjne_rozszerzone(self, info):
        # self.result = None
        if not self.dialog:
            self.dialog = MDDialog(
                text=info,
                buttons=[
                    MDFlatButton(
                        text="Zastąp",
                        on_release = lambda x: self.close_dialog_and_give_feedback('Zastąp')
                    ),
                    MDFlatButton(
                        text="Nowy",
                        on_release = lambda x: self.close_dialog_and_give_feedback('Nowy')
                    ),
                    MDFlatButton(
                        text="Nic",
                        on_release = lambda x: self.close_dialog_and_give_feedback('Nic')
                    ),
                ],
            )
        self.dialog.open()
        return 'ok'
    
    # def close_dialog(self, obj):
    def close_dialog(self, obj):
        self.dialog.dismiss()

    def close_dialog_and_give_feedback(self, button_text):
        self.dialog.dismiss()
        # self.result = button_text
        return button_text










#######ODATKOWE MODULY
class StrokaGood(MDCard):
    def __init__(self, info = None, **kwargs):
        super().__init__(**kwargs)
        self.info = info
        self.linijka = str(info[0][0])
        self.obudowa = str(info[1])
        self.ilosc_padow = f'{str(info[3])} x {str(info[2])} = {str(int(info[3]) * int(info[2]))}'

    
class StrokaError(MDCard):
    def __init__(self, info = None, **kwargs):
        super().__init__(**kwargs)
        self.info = info
        self.linijka = str(info[0][0])
        self.obudowa = str(info[1])
        self.ilosc_szt = str(info[2])


######ODPALENIE APLIKACJI
class WindowManager(ScreenManager):
    pass


class NanotechApp(MDApp):
    def build(self):
        self.theme_cls.theme_style = 'Dark'
        self.theme_cls.primary_palette = 'BlueGray'
        return Builder.load_file("main.kv")


if __name__ == "__main__":
    NanotechApp().run()
