# -*- coding: cp1250 -*-
# verze 2021-12
# vytvoril Milan Hutera
# doplněno otvírání CSV s názvem sloupců
# doplněno kodování u ukládání
# Head2 bere celou hlavičku u všech prvků ne jen u prvního

import os
import mmap
import codecs
from shutil import copyfile
from glob import glob

class Soubor:
    Pomoc = 'pp'
    def PorovnejSoubory(self, Prvni,Druhy):
        """Vrátí pole s pozicema kde je rozdíl, k dalšímu zpracování je určená funkce VyrezStringu."""
        Rozdil = []
        if len(Prvni)<=len(Druhy):
            S1 = Prvni
            S2 = Druhy
        else:
            S1=Druhy
            S2=Prvni
        Pozice = 0
       
        for i in S1:
            if i!=S2[Pozice]:
                Rozdil.append(Pozice)
            Pozice=Pozice+1
        return Rozdil

    def VyrezStringu(self,Zdroj,Vyrez):
        """ Zdroj je soubor, Vyrez je pole z funkce PorovnejSoubory. Vysledek je pole s pozicí a slovem kde byl stringem."""
        Poradi = 0
        Pamet = Vyrez[0]
        Slovo = ''
        Vysledek = []
        Cast = [Vyrez[0]]
        for i in Vyrez:
            if i == Pamet:
                Slovo = Slovo + Zdroj[i]
                Pamet = Vyrez[Poradi] + 1
            else:
                Cast.append(Slovo)
                Vysledek.append(Cast)
                Cast = [Vyrez[Poradi]]
                Pamet = Vyrez[Poradi] + 1
                Slovo = Zdroj[i]
            Poradi = Poradi + 1
        if len(Vysledek)>0:
            #if Vysledek[-1][1]<>Slovo:
            Cast.append(Slovo)
            Vysledek.append(Cast)
        else:
            Cast.append(Slovo)
            Vysledek.append(Cast)
        return Vysledek

    def VyznaceniRozdilu(self,Zdroj,Rozdil,Znak='|'):
        """ Zdroj=Soubor,Rozdil=pole z PorovnejSoubory. Vloží před a za rozdíl Znak."""
        Vyrez = self.VyrezStringu(Zdroj,Rozdil)
        PredchoziPozice = 0
        Vysledek = ''
        
        for i in range(len(Vyrez)):
            #print Zdroj[PredchoziPozice:Vyrez[i][0]]
            Vysledek = Vysledek + Zdroj[PredchoziPozice:Vyrez[i][0]] + Znak + Vyrez[i][1]
            PredchoziPozice = Vyrez[i][0] + len(Vyrez[i][1])
        Vysledek = Vysledek + Zdroj[PredchoziPozice:PredchoziPozice+len(Vyrez[-1][1])]+Znak+Zdroj[PredchoziPozice+len(Vyrez[-1][1]):]
        return Vysledek
            
            
        
    

    def VycistiAdresarFiles(self, Zdroj, Pripona=''):
        SmazaneSoubory = list()
        if Zdroj[-1]!='\\':
            Zdroj = Zdroj + '\\'
        for dirpath, dirnames, filenames in os.walk(Zdroj):
            for NazevSouboru in filenames:
                Smaz = os.path.join(dirpath, NazevSouboru)
                if Pripona=='' or NazevSouboru[-len(Pripona):]==Pripona:
                    #print 'Mazu Soubor:' , os.path.join(dirpath, NazevSouboru)
                    SmazaneSoubory.append(Smaz)
                    os.remove(Smaz)
        return SmazaneSoubory

    def VycistiAdresarDirs(self, Zdroj, Pripona=''):
        SmazaneAdresare = list()
        if Zdroj[-1]!='\\':
            Zdroj = Zdroj + '\\'
        for dirpath, dirnames, filenames in os.walk(Zdroj):
            for NazevAdresare in dirnames:
                Smaz = os.path.join(dirpath, NazevAdresare)
                SmazaneAdresare.append(Smaz)
        SmazaneAdresare.reverse()
        for Smaz in SmazaneAdresare:
            print('Mazu Adresář: ', Smaz)
            if (os.path.isdir(Smaz)):
                os.removedirs(Smaz)
        if not(os.path.isdir(Zdroj)):
            os.makedirs(Zdroj)
        return SmazaneAdresare

    def VycistiAdresarAll(self,Zdroj):
        self.VycistiAdresarFiles(Zdroj)
        self.VycistiAdresarDirs(Zdroj)
        
    def VytvorAdresar( self, Adresar=''):
        if Adresar =='' or Adresar[0]== '\\':
            Adresar = (os.getcwd() + Adresar)        
        if os.path.isdir(Adresar):
            return False
        os.mkdir(Adresar)
        return True        
        
    def CopyAll(self, Zdroj, Cil, Pripona=''):
        Pripona = Pripona[0].lstrip('.*').upper()
        KopirovaneSoubory = list()
        if Zdroj[-1]!='\\':
            Zdroj = Zdroj + '\\'
        if Cil[-1]!='\\':
            Cil = Cil + '\\'        
        for dirpath, dirnames, filenames in os.walk(Zdroj):
            for name in dirnames:
                NazevAdresare = os.path.join(dirpath, name)
                if not(os.path.isdir(Cil + NazevAdresare.replace(Zdroj,''))):
                    print(Cil + NazevAdresare.replace(Zdroj,''))
                    os.makedirs(Cil + NazevAdresare.replace(Zdroj,''))
            for name in filenames:
                NazevSouboru = os.path.join(dirpath, name)
                if not(os.path.isfile(Cil + NazevAdresare.replace(Zdroj,''))) and ( Pripona=='' or NazevSouboru[-len(Pripona):].upper()==Pripona.upper()):
                    KopirovaneSoubory.append(Cil + NazevSouboru.replace(Zdroj,''))
                    copyfile(NazevSouboru,Cil + NazevSouboru.replace(Zdroj,''))
        return KopirovaneSoubory

    
    def Copy(self,Zdroj, Cil, SeznamSouboru=['*.*']):
        """Kopíruje soubory neptá se na přepsání rovnou přepíše."""
        KopirovaneSoubory = list()
        
        if Zdroj=='' or Zdroj[-1]!='\\':
            Zdroj = Zdroj + '\\'
        if Cil=='' or Cil[-1]!='\\':            
            Cil = Cil + '\\'
            
        if Zdroj =='' or Zdroj[0]== '\\':
            Zdroj = (os.getcwd() + Zdroj)
        if Cil =='' or Cil[0]== '\\':
            Cil = (os.getcwd() + Cil)
        if not(os.path.isdir(Cil)):
            print('Vytvářím adresář: ',Cil)
            os.makedirs(Cil)
            
        for i in SeznamSouboru:
            if i == '*':
                KopirovaneSoubory.append(self.CopyAll(Zdroj,Cil))
            else:
                if i.find('*')<0:
                    copyfile(Zdroj+i,Cil+i)
                    KopirovaneSoubory.append(Cil+i)
                else:
                    SeznamSouboru2 = self.ListFiles(Zdroj+i)
                    print('Zkopírováno z adresáře ',Zdroj,':')
                    for f in SeznamSouboru2:
                        KopirovaneSoubory.append(Cil+f)
                        print(Cil+f)
                        copyfile(Zdroj+f,Cil+f)
        return KopirovaneSoubory

        
    def Dir(self, Adresar = ''):
        if Adresar =='' or Adresar[0]== '\\':
            Adresar = (os.getcwd() + Adresar)
        List = os.listdir(Adresar)
        return List

    def ListDir(self, Adresar=''):
        if Adresar =='' or Adresar[0]== '\\':
            Adresar = (os.getcwd() + Adresar)
            
        if Adresar.find('*')<0:
            SeznamAdresaru = [s for s in os.listdir(Adresar) if os.path.isdir(s)]
        else:
            SeznamAdresaru = [s for s in glob(Adresar) if os.path.isdir(s)]
            s = []
            for i in SeznamAdresaru:
                s.append(i[i.rfind('\\')+1:])
            SeznamAdresaru = s
        return SeznamAdresaru
    
    def ListFiles(self, Adresar='',CelouCestu = False):
        """Vratí seznam souboru. Chcešli všechny dej ....\\* """
        if Adresar =='' or Adresar[0]== '\\':
            Adresar = (os.getcwd() + Adresar)
        if Adresar.find('*')<0:
            SeznamSouboru = [s for s in os.listdir(Adresar) if os.path.isfile(s)]
        else:
            SeznamSouboru = [s for s in glob(Adresar) if os.path.isfile(s)]
            if not(CelouCestu):
                s = []
                for i in SeznamSouboru:
                    s.append(i[i.rfind('\\')+1:])
                SeznamSouboru = s                         
        return SeznamSouboru
    
    def Nahrad(self,Zdroj,Old, New):
        """Nahradí old za new, nezaleží však na velikosti písmen jako u replace."""
        ZdrojUp = Zdroj.upper()
        OldUp = Old.upper()
        Vyskytu = ZdrojUp.count(OldUp)
        Start = 0
        Delka = len(Old)
        for i in range(Vyskytu):
            Start = ZdrojUp.find(OldUp,Start)
            Zdroj = Zdroj[:Start]+New+Zdroj[Start+Delka:]
            Start = Start + len(New)
            ZdrojUp = Zdroj.upper()
        return Zdroj
            
        
        
    
    def Serad(self,Co,PodleCeho=0):
        "Metoda na seřazení dat podle velkosti."
        def Porovnej(Seznam):
            return str.lower(Seznam[PodleCeho])  #lower proto aby se neovlivňovalo pořadí velikostí písmen

        Co.sort(key=Porovnej)
        return Co
    def Nahraj(self, Co = 'Seznam.txt'):      
        "Metoda pro nahrání dat ze souboru."
        if Co =='' or Co[0]== '\\':
            Co = os.getcwd() + Co
            
        Nazev = Co
        Data = []
        if self.FileExists(Nazev)==0:
            return -1
        Soubor = open(Nazev,'r+')  # 'r+' Diky tomuto mohu do souboru i zapisovat a to na pozici kam chci
        DelkaSouboru = os.path.getsize(Nazev)
        DataSouboru = mmap.mmap(Soubor.fileno(),DelkaSouboru)
        Data = DataSouboru.read(DelkaSouboru)
        DataSouboru.close()
        Soubor.close()
        return Data

    def NahrajRadky(self, Co = 'Seznam.txt'):      
        "Metoda pro nahrání dat ze souboru.Rozdělí na jednotlivý řádky."
        if Co =='' or Co[0]== '\\':
            Co = os.getcwd() + Co
        Nahravka = []
        Nazev = Co
        if self.FileExists(Nazev)==0:
            return
        Soubor = file(Nazev,'r')  
        for Line in Soubor:
            Line = Line.strip('\r\n')           # zbaví řádku znaků "\r\n"
            
            Nahravka.append(Line)
        return Nahravka

    
    
    def UdelejCSV(self, Co ,Separator=None):      
        "Metoda pro vytvoření pole z dat. Rozdělí na jednotlivý řádky a slova."
        Nahravka = []
        Soubor = Co.split('\r\n')
        for Line in Soubor:
            Nahravka.append(Line.split(Separator))
        return Nahravka

    def UdelejData(self, Co ,Separator=None):      
        "Opak od Udelj CSV. Spoji CSV v jeden celek."
        Result = ''
        for i in CO:
            for ii in i:
                Result = Result + ii + ' '
            Result = Result + '\r\n'
        return Result   

    def NahrajRadkySlov(self, Co = 'Seznam.txt',Separator=None):      
        "Metoda pro nahrání dat ze souboru. Rozdělí na jednotlivý řádky a slova."
        if Co =='' or Co[0]== '\\':
            Co = os.getcwd() + Co
        Nahravka = []
        Nazev = Co
        if self.FileExists(Nazev)==0:
            return
        Soubor = file(Nazev,'r')  
        for Line in Soubor:
            Line = Line.strip('\r\n')           # zbaví řádku znaků "\r\n"
            
            Nahravka.append(Line.split(Separator))
        return Nahravka
    
    def NahrajCSV(self, CsvSoubor='IO.CSV',Separator=';'):
        """Metoda pro nahrání dat z CSV souboru, oddělovač středník"""
        Vysledek = self.NahrajRadkySlov(CsvSoubor,Separator)
        return Vysledek

    def NahrajCSVHead(self,CsvSoubor='IO.CSV',Separator=';'):
        """Metoda pro nahrání dat z CSV souboru, oddělovač středník možnost změnit parametrem Separator."""
        "Doplněno o rozdělení do sloupců dle názvů. "
        def Head(Source):
            
            Index = 0
            Result={}
            for i in Source:
                Result[i]=Index
                Index = Index + 1
            return Result
        Vysledek = self.NahrajRadkySlov(CsvSoubor,Separator)
        
        HVysledek = Head(Vysledek[0])
        Result = []
        AuxPoradi = 0
        for i in Vysledek[1:]:
            r = {}
            if AuxPoradi == 0:
                r['XXHead']= Vysledek[0]
                AuxPoradi = 1
            for ii in HVysledek:
                if len(i)>HVysledek[ii]:
                    r[ii] = i[HVysledek[ii]]  # 
                else:
                    r[ii] = ""      # pokud nastane chyba exists je 0
            Result.append(r)
        return Result
    def Row(self,Length):
        "vrátí array vámi dané délky"
        Result = []
        for i in range(Length):
            Result.append('')
        return Result
    
    def UlozCSVHead(self,Co,Kam = 'IO.csv',Separator=';',ParametrSouboru='w',Code='cp1250',Head = []):
        "Metoda na uložení dat ve formátu Head."
##        Data = self.Row(len(Co[0]))
        if 'XXHead' in Co[0] and Head == []:
            Head = Co[0].pop('XXHead')
            
        Data = [Head]
        for i in list(Co[0].keys()):            
            if Data[0].count(i)==0:
                Data[0].append(i)
        
        for i in Co:
            Data.append([])
            for ii in Data[0]:
                
                if ii in i:
                    Data[-1].append(i[ii])
                else:
                    Data[-1].append('')
        Co[0]['XXHead'] = Data[0]
        self.UlozCSV3(Data,Kam,Separator,ParametrSouboru)                

    def UlozCSVHead2(self,Co,Kam = 'IO.csv',Separator=';',ParametrSouboru='w',Head = []):
        "Metoda na uložení dat ve formátu Head."
##        Data = self.Row(len(Co[0]))
        Data = [Head]
        Aux = {}            
        for i in Co:
            for ii in i:
                if Data[0].count(ii)==0:
                    Aux[ii]=1
        for i in Aux:
            Data[0].append(i)            
        for i in Co:
            Data.append([])
            for ii in Data[0]:
                if ii in i:
                    Data[-1].append(i[ii])
                else:
                    Data[-1].append('')
##        self.UlozCSV(Data,Kam,Separator,ParametrSouboru)
        return Data

    
    def FileExists(self,Soubor):
        "Metoda určující existenci souboru"
        
        try:
            file = open(Soubor)  # pokusí se otevřít soubor
        except IOError:
            exists = 0      # pokud nastane chyba exists je 0
        else:
            exists = 1      
        return exists
    
    def VytvorAdresar(self,Kam):
        if not os.path.isdir(Kam[:Kam.rfind('\\')]) and Kam.rfind('\\')>-1:
            os.makedirs(Kam[:Kam.rfind('\\')])

    def Uloz(self,Co,Kam = 'Soubor.txt'):
        "Metoda pro ukládání dat do souboru."
        "Pokud chci unicode použij Uloz(Co.encode('UTF-8'))"
        if Kam =='' or Kam[0]== '\\':
            Kam = os.getcwd() + Kam
        self.VytvorAdresar(Kam)
        Soubor=file(Kam,'w')
        Soubor.write(' ')
        Soubor.close()
        Soubor=file(Kam,'r+')      # 'a' append , 'w' write prepise stary, 'r' read
        DataSouboru = mmap.mmap(Soubor.fileno(),len(Co))
        DataSouboru.write(Co)
        DataSouboru.close()
        Soubor.close()
        
    def UlozCSV2(self,Co,Kam = 'IO.csv',Separator=';',ParametrSouboru='w',Code = 'cp1250'):
        if Kam =='' or Kam[0]== '\\':
            Kam = os.getcwd() + Kam
        if not os.path.isdir(Kam[:Kam.rfind('\\')]):
            os.makedirs(Kam[:Kam.rfind('\\')])
        self.Soubor=file(Kam,ParametrSouboru)      # 'a' append , 'w' write prepise stary, 'r' read
        for Poradi in Co:
##            print Poradi
            for i in Poradi[:-1]:
                
                self.Soubor.write(str(i) + Separator)     # e.
            
            self.Soubor.write(Poradi[-1]+'\n')      # nov v souboru 
        self.Soubor.close()                    


    def UlozCSV(self,Co,Kam = 'IO.csv',Separator=';',ParametrSouboru='w'):
        """Metoda pro ukládání dat array do souboru CSV oddělovač je středník."""
        if Kam =='' or Kam[0]== '\\':
            Kam = os.getcwd() + Kam
        self.VytvorAdresar(Kam)
        self.Soubor=file(Kam,ParametrSouboru)      # 'a' append , 'w' write prepise stary, 'r' read
        for Poradi in Co:
            Row = ''
            for i in Poradi:
##                print i
                if type(i) is int:
                    Row = Row + str(i) + Separator
                else:
                    Row = Row + i + Separator
             
            if len(Row)>0:
                Row = Row[:-1]
            self.Soubor.write(Row)     # zapíše data do souboru v unicode.
            self.Soubor.write('\n')      # nová řádka v souboru 
        self.Soubor.close() 

    def UlozCSV3(self,Co,Kam = 'IO.csv',Separator=';',ParametrSouboru='w',Code = 'cp1250'):
        """Metoda pro ukládání dat array do souboru CSV oddělovač je středník."""
        if Kam =='' or Kam[0]== '\\':
            Kam = os.getcwd() + Kam
        self.VytvorAdresar(Kam)
        self.Soubor=file(Kam,ParametrSouboru)      # 'a' append , 'w' write prepise stary, 'r' read
        for Poradi in Co:
            Row = ''
            
            for i in Poradi:
##                print i
                if type(i) is int and i!='':
                    Row = Row + str(i) + Separator
                else:
                    if type(i) is str:
                        Row = Row + i.decode(Code) + Separator
                    else:
                        print("problem s typem")
             
            if len(Row)>0:
                Row = Row[:-1]
            self.Soubor.write(Row.encode(Code))     # zapíše data do souboru v unicode.
            self.Soubor.write('\n')      # nová řádka v souboru 
        self.Soubor.close()        
