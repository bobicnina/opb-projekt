import sqlite3
import bottle
import hashlib
from xlrd import *
import urllib.request
import json


bottle.debug(True)

static_dir = "./Static"

baza_datoteka="prevoznistvo.sqlite3"
baza = sqlite3.connect(baza_datoteka, isolation_level=None)
secret="123kajf98453jf"



##Pomožne funkcije

def uvoz_p(exc):
    '''Uvozi podatke iz excela v bazo'''
    napaka=None
    stran=exc.sheet_by_name('List1')
    ime=stran.cell_value(0, 9) #prvi je indeks vrstice, drugi stolpca
    c=baza.cursor()
    c.execute('''SELECT registrska FROM tovornjak WHERE ime=?''', [ime])
    a=tuple(c)
    if len(a)==0:
        napaka="Te registrske številke ni v bazi."
        return napaka
    else: registrska=a[0][0]
    
    #pregleda, če so te prevozi že vnešeni v bazo
    prvi_datum=xldate_as_tuple(stran.cell_value(4, 2), 0)
    datum=str(prvi_datum[0])+"-"+str(prvi_datum[1])+"-"+str(prvi_datum[2])
    c.execute("SELECT * FROM prevoz WHERE datum=? AND registrska=?",
                  [datum, registrska])
    if c.fetchone() is not None:
        napaka="Prevozi so že vnešeni v bazo."
        return napaka

    
    i=4 #v peti vrstici se začnejo datumi prevozov
    while i<35:
        if len(stran.cell_value(i,1))==0:
            #če je prazen ali ima napisano dopust, ga ne vpiše v bazo
            i+=1
            continue
        if stran.cell_value(i,1)=="DOPUST":
            i+=1
            continue
        #mesto1 je začetek, mesto2 konec prevoza
        mesto1, mesto2 = stran.cell_value(i,1).split(' - ')
        mesto1.strip()
        mesto2.strip()
        if mesto1=="LJ" or mesto1=="lj":
            mesto1="Ljubljana"
        if mesto2=="LJ" or mesto2=="lj":
            mesto2="Ljubljana"

    #Ostali podatki o prevozu
        kolicina=stran.cell_value(i, 7)
        cena=stran.cell_value(i, 6)
        datum=xldate_as_tuple(stran.cell_value(i, 2), 0)
        datum=str(datum[0])+"-"+str(datum[1])+"-"+str(datum[2])
    
        
    #Na internet pogleda razdaljo od Škocjana do začetka/konca prevoza, če mesta
    #še ni v bazi, ga vpiše
        getdata = {"origins": "skocjan", "destinations": mesto1, "mode": "driving", "language": "en-EN", "sensor": "false"}
        url1="http://maps.googleapis.com/maps/api/distancematrix/json?" + urllib.parse.urlencode(getdata)
        f1 = urllib.request.urlopen(url1).read().decode("UTF-8")
        result1=json.loads(f1)
        razdalja1 = int((result1['rows'][0]['elements'][0]['distance']['value'])/1000)
        c.execute("""SELECT * FROM mesta WHERE ime=?""", [mesto1])
        if c.fetchone() is None:
            baza.execute("""INSERT INTO mesta(ime, razdalja) VALUES (?, ?)""", (mesto1, razdalja1))
        else: pass

        getdata2 = {"origins": "skocjan", "destinations": mesto2, "mode": "driving", "language": "en-EN", "sensor": "false"}
        url2="http://maps.googleapis.com/maps/api/distancematrix/json?" + urllib.parse.urlencode(getdata2)
        f2 = urllib.request.urlopen(url2).read().decode('iso-8859-1')
        result2=json.loads(f2)
        razdalja2 = int((result2['rows'][0]['elements'][0]['distance']['value'])/1000)
        c.execute("""SELECT * FROM mesta WHERE ime=?""", [mesto2])
        if c.fetchone() is None:
            baza.execute("""INSERT INTO mesta(ime, razdalja) VALUES (?, ?)""", (mesto2, razdalja2))
        else: pass
        i+=1
     #VNOS PREVOZA
        c.execute("""INSERT INTO prevoz(datum, kolicina, cena_tone, zacetek, konec, registrska) VALUES (?, ?, ?, ?, ?, ?)""", (datum, kolicina, cena, mesto1, mesto2, registrska))

    #LETO-MESEC
    mesec=str(prvi_datum[0])+'-'+str(prvi_datum[1])   #ven da leto-mesec v številkah

    #CESTNINA
    slovenija=stran.cell_value(53, 2)
    hrvaska=stran.cell_value(54, 2)
    c.execute("""INSERT INTO cestnina(registrska, mesec, slovenija, hrvaska)
                  VALUES (?, ?, ?, ?)""", [registrska, mesec, slovenija, hrvaska])

    #MESEČNI STROŠKI
    gorivo=stran.cell_value(45, 6)
    kilometri=stran.cell_value(46,5)
    c.execute("""INSERT INTO mesecni_stroski(registrska, mesec, gorivo, razdalja) VALUES (?, ?, ?, ?)""",
              [registrska, mesec, gorivo, kilometri])

    baza.commit()
    napaka="Uspešno ste vnesli podatke o mesečnih prevozih."
    return napaka


def mesec_beseda(mesec, leto):
    #napiše ime meseca in leto ter še enega pred njim
    meseci=['Januar', 'Februar', 'Marec', 'April', 'Maj', 'Junij', 'Julij',
            'Avgust', 'September', 'November', 'December']
    meseca=[]
    if mesec=="1":
        leto2=leto-1
        meseca.append("Januar"+" "+leto)
        meseca.append("December"+" "+leto2)
    else:
        meseca.append(meseci[int(mesec)-1]+" "+leto)
        meseca.append(meseci[int(mesec)-2]+" "+leto)
    return meseca

def en_mesec(mesec, leto):
    meseci=['Januar', 'Februar', 'Marec', 'April', 'Maj', 'Junij', 'Julij',
            'Avgust', 'September', 'November', 'December']
    return meseci[int(mesec)-1]+' '+leto

def dobi_uporabnika(auto_login = True):
    #preveri, če je uporabnik vpisan, če ni, ga vrže na prijavno stran
    uporabnik = bottle.request.get_cookie('uporabnik', secret=secret)
    if uporabnik is not None:
        c = baza.cursor()
        c.execute("SELECT uporabnik FROM uporabnik WHERE uporabnik=?", [uporabnik])
        r = c.fetchone()
        c.close ()
        if r is not None:
            # uporabnik obstaja, vrnemo njegove podatke
            return tuple(r)[0]
    # Če pridemo do sem, uporabnik ni prijavljen, naredimo redirect
    if auto_login:
        bottle.redirect('/prijava/')
    else:
        return None


def password_md5(s):
    """Vrni MD5 hash danega UTF-8 niza. Gesla vedno spravimo v bazo
       kodirana s to funkcijo."""
    h = hashlib.md5()
    h.update(s.encode('utf-8'))
    return h.hexdigest()





##BOTTLE strani

@bottle.route("/static/<filename:path>")
def static(filename):
    """Splošna funkcija, ki servira vse statične datoteke iz naslova
       /static/..."""
    return bottle.static_file(filename, root=static_dir)


@bottle.route("/")
def main():
    """Glavna stran."""
    # Iz cookieja dobimo uporabnika (ali ga preusmerimo na login, če
    # nima cookija)
    uporabnik=dobi_uporabnika()
    return bottle.template("main.html", uporabnik=uporabnik)


@bottle.get("/prijava/")
def prijava():
    return bottle.template("prijava.html", napaka=None, uporabnik=None)


@bottle.post("/prijava/")
def login_post():
    uporabnik = bottle.request.forms.uporabnik
    geslo = password_md5(bottle.request.forms.geslo)
    c = baza.cursor()
    c.execute("SELECT 1 FROM uporabnik WHERE uporabnik=? AND geslo=?",
              [uporabnik, geslo])
    if c.fetchone() is None:
        # Username in geslo se ne ujemata
        return bottle.template("prijava.html",
                               napaka="Nepravilna prijava",
                               uporabnik=uporabnik)
    else:
        # Vse je v redu, nastavimo cookie in preusmerimo na glavno stran
        bottle.response.set_cookie('uporabnik', uporabnik, path='/', secret=secret)
        bottle.redirect("/")


@bottle.get("/odjava/")
def odjava():
    bottle.response.delete_cookie('uporabnik')
    return bottle.template("odjava.html")

@bottle.get("/registracija/")
def register():
    return bottle.template("registracija.html", napaka=None, uporabnik=None)

    
@bottle.post("/registracija/")
def register_post():
    """Registriraj novega uporabnika."""
    uporabnik = bottle.request.forms.uporabnik
    ime = bottle.request.forms.ime
    geslo1 = bottle.request.forms.geslo1
    geslo2 = bottle.request.forms.geslo2
    # Ali uporabnik že obstaja?
    c = baza.cursor()
    c.execute("SELECT 1 FROM uporabnik WHERE uporabnik=?", [uporabnik])
    if c.fetchone():
        # Uporabnik že obstaja
        return bottle.template("registracija.html",
                               uporabnik=uporabnik,
                               ime=ime,
                               napaka='To uporabniško ime je že zavzeto')
    elif not geslo1 == geslo2:
        # Gesli se ne ujemata
        return bottle.template("registracija.html",
                               uporabnik=uporabnik,
                               ime=ime,
                               napaka='Gesli se ne ujemata')
    else:
        # Vse je v redu, vstavi novega uporabnika v bazo
        geslo = password_md5(geslo1)
        c.execute("INSERT INTO uporabnik (uporabnik, geslo) VALUES (?, ?)",
                  (uporabnik, geslo))
        # Daj uporabniku cookie
        bottle.response.set_cookie('uporabnik', uporabnik, path='/', secret=secret)
        bottle.redirect("/")


@bottle.route("/pregled/")
def pregled():
    #Izpiše imena in priimke voznikov
    c=baza.cursor()
    c.execute('''SELECT ime, priimek, registrska FROM tovornjak''')
    seznam=tuple(c)

    #Izpiše zadnja dva meseca
    b=baza.cursor()
    b.execute('''SELECT datum FROM prevoz ORDER BY datum DESC LIMIT 1''')
    b=tuple(b)
    if len(b)==0: #če v bazi ni nobenega prevoza
        meseci=[]
        leto=[]
    else:
        datum=b[0][0].split("-")
        leto=datum[0]
        meseci=mesec_beseda(datum[1], leto)
        meseci.append(datum[1]) #da se naredi link
        meseci.append(int(datum[1])-1)
        #izgleda tako, npr.: meseci=['Januar 2014', 'Februar 2014', 1, 2]
    return bottle.template("pregled.html", seznam=seznam, meseci=meseci, leto=leto)

@bottle.route("/pregled/<registrska>/")
def pregled(registrska):
    #Izpiše vse prevoze enega voznika in njegove podatke
    a=baza.cursor()
    a.execute("SELECT * FROM tovornjak WHERE registrska=?", [registrska])
    a=tuple(a)[0] #seznam vseh njegovih podatkov
    c=baza.cursor()
    c.execute("SELECT * FROM prevoz WHERE registrska=?", [registrska])
    c=tuple(c)
    return bottle.template("voznik.html", podatkivoznika=a, podatki=c)

@bottle.route("/pregled/<leto>/<mesec>/")
def pregled1(leto, mesec):
    napaka=None
    datum=str(leto)+'-'+str(mesec)
    c=baza.cursor()
    c.execute("SELECT registrska, gorivo, razdalja FROM mesecni_stroski WHERE mesec=?", [datum])
    m=tuple(c)
    if len(m)==0:
        napaka="Za ta mesec ni nobenega podatka." 
    mesec=en_mesec(mesec, leto) 
    podatki=[]
    for seznam in m:
        registrska=seznam[0]
        c.execute("SELECT ime, priimek FROM tovornjak WHERE registrska=?", [registrska])
        ime=tuple(c)
        voznik=[]
        voznik.extend((ime[0][0], ime[0][1], seznam[0], seznam[1], seznam[2]))
        podatki.append(voznik)
    return bottle.template("mesec.html", datum=mesec, podatki=podatki, napaka=napaka)

@bottle.get("/uvoz/")
def uvoz():
#prikaži formo za vnos datoteke
    return bottle.template("uvoz.html", napaka=None)

@bottle.post("/uvoz1/")
def uvoz():
#doda novega voznika
    napaka=None
    ime=bottle.request.forms.ime
    priimek=bottle.request.forms.priimek
    datum=bottle.request.forms.datum
    registrska=bottle.request.forms.registrska
    nosilnost=bottle.request.forms.nosilnost
    c=baza.cursor()
    c.execute('''INSERT INTO tovornjak(registrska, nosilnost,
               datum_rojstva, ime, priimek) VALUES (?, ?, ?, ?, ?)''',
              [registrska, nosilnost, datum, ime, priimek])
    baza.commit()
    return bottle.template("uvoz.html", napaka="Uspešno ste vnesli novega voznika.")
        
@bottle.post("/uvoz/")
#uploada datoteko
def uvoz():
    data=bottle.request.files.data.file
    exc=open_workbook(file_contents=data.read())
    napaka=uvoz_p(exc)
    return bottle.template("uvoz.html", napaka=napaka)


@bottle.route("/sprememba/")
def sprememba():
    return bottle.template("sprememba.html", obvestilo=None)

@bottle.post("/sprememba/")
def sprememba_voznik():
    ime_new = bottle.request.forms.ime
    priimek_new = bottle.request.forms.priimek
    datum_rojstva_new = bottle.request.forms.datum
    registrska = bottle.request.forms.registrska
    c=baza.cursor()
    c.execute("SELECT * FROM tovornjak WHERE registrska=?", [registrska])
    if c.fetchone() is None:
        return bottle.template("sprememba.html", obvestilo="Te registrske številke ni v bazi.")
    c.execute("UPDATE  tovornjak SET ime=?, priimek=?, datum_rojstva=? WHERE registrska=?",
              [ime_new, priimek_new, datum_rojstva_new, registrska])
    return bottle.template("sprememba.html", obvestilo="Uspešno ste spremenili podatke o vozniku.")

@bottle.post("/sprememba1/")
def sprememba():
    stara=bottle.request.forms.stara
    nova=bottle.request.forms.nova
    nova_n=bottle.request.forms.nova_n
    c=baza.cursor()
    
    c.execute("SELECT * FROM tovornjak WHERE registrska=?", [stara])
    if c.fetchone() is None:
        return bottle.template("sprememba.html", obvestilo="Registrske številke "+stara+" ni v bazi.")

    c.execute("SELECT * FROM tovornjak WHERE registrska=?", [nova])
    if c.fetchone() is not None:
        return bottle.template("sprememba.html", obvestilo="Registrska številka "+nova+" je že v bazi.")

    c.execute("UPDATE  tovornjak SET registrska=?, nosilnost=?  WHERE registrska=?",
              [nova, nova_n, stara])
    return bottle.template("sprememba.html", obvestilo="Uspešno ste spremenili registrsko številko.")

    
bottle.run(host='localhost', port=8080)

