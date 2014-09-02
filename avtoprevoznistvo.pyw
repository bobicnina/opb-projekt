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
secret="sara"



##Pomožne funkcije

def uvoz_p(exc):
    '''Uvozi podatke iz excela v bazo'''
    stran=exc.sheet_by_name('List1')
    ime=stran.cell_value(0, 9)
    c=baza.cursor()
    c.execute('''SELECT registrska FROM tovornjak WHERE ime=?''', [ime])
    if c.fetchone() is None:
        napaka="Te registrske številke ni v bazi."
        return bottle.template("uvoz.html", napaka=napaka)
    else: registrska=tuple(c)[0][0]
    stran=exc.sheet_by_name('List1')
    i=4
    while i<35:
        if len(stran.cell_value(i,1))==0:
            i+=1
            continue
        if stran.cell_value(i,1)=="DOPUST":
            i+=1
            continue
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
        kilometri=stran.cell_value(46,5)
        datum=xldate_as_tuple(stran.cell_value(i, 2), 0)
        datum=str(datum[0])+"-"+str(datum[1])+"-"+str(datum[2])
        
        c=baza.cursor()

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
        print(i, mesto1, razdalja1)

        getdata2 = {"origins": "skocjan", "destinations": mesto2, "mode": "driving", "language": "en-EN", "sensor": "false"}
        url2="http://maps.googleapis.com/maps/api/distancematrix/json?" + urllib.parse.urlencode(getdata2)
        f2 = urllib.request.urlopen(url2).read().decode('iso-8859-1')
        result2=json.loads(f2)
        razdalja2 = int((result2['rows'][0]['elements'][0]['distance']['value'])/1000)
        c.execute("""SELECT * FROM mesta WHERE ime=?""", [mesto2])
        if c.fetchone() is None:
            baza.execute("""INSERT INTO mesta(ime, razdalja) VALUES (?, ?)""", (mesto2, razdalja2))
        else: pass
        print(i, mesto2, razdalja2)
        i+=1
     #VNOS PREVOZA
        c.execute("""INSERT INTO prevoz(datum, kolicina, cena_tone, zacetek, konec, registrska) VALUES (?, ?, ?, ?, ?, ?)""", (datum, kolicina, cena, mesto1, mesto2, registrska))
    baza.commit()


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
        # Geslo se ne ujemata
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
    c.execute('''SELECT * FROM tovornjak''')
    imen=tuple(c)
    imena={}
    for one in imen:
        imena[one[3]]=one[4]

    #Izpiše zadnja dva meseca
    b=baza.cursor()
    b.execute('''SELECT datum FROM prevoz ORDER BY datum DESC''')
    datum=str(tuple(b)[0][0]).split("-")
    leto=datum[0]
    meseci=mesec_beseda(datum[1], leto)
    meseci.append(datum[1]) #da se naredi link
    meseci.append(int(datum[1])-1)
    return bottle.template("pregled.html", imena=imena, meseci=meseci, leto=leto)

@bottle.route("/pregled/<voznik>/")
def pregled(voznik):
    #Izpiše vse prevoze enega voznika in njegove podatk
    a=baza.cursor()
    a.execute("SELECT * FROM tovornjak WHERE ime=?", [voznik])
    a=tuple(a)[0]
    b=baza.cursor() #za voznika poišče njegovo registrsko
    b.execute("SELECT registrska FROM tovornjak WHERE ime=?", [voznik])
    b=b.fetchone()
    registrska=tuple(b)[0]
    c=baza.cursor()
    c.execute("SELECT * FROM prevoz WHERE registrska=?", [registrska])
    c=tuple(c)
    return bottle.template("voznik.html", voznik=voznik, podatkivoznika=a, podatki=c)

@bottle.route("/pregled1/<leto>/<mesec>")
def pregled1(leto, mesec):
    print("kja")
    napaka=None
    c=baza.cursor()
    c.execute("SELECT * FROM mesecni_stroski WHERE mesec=? AND leto=?", [mesec, leto])
    if c.fetchone() is None:
        napaka="Za ta mesec ni nobenega podatka."
    else: c=tuple(c)
    mesec=mesec_beseda(1, leto)
    return bottle.template("mesec.html", datum=mesec, podatki=c, napaka=napaka)

@bottle.get("/uvoz/")
def uvoz1():
#prikaži formo za vnos datoteke
    return bottle.template("uvoz.html")

@bottle.post("/uvoz1/")
def uvoz():
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
    bottle.redirect("/")
        
@bottle.post("/uvoz/")
#uploada datoteko
def uvoz():
    data=bottle.request.files.data.file
    exc=open_workbook(file_contents=data.read())
    uvoz_p(exc)
    bottle.redirect("/")


bottle.run(host='localhost', port=8080)

