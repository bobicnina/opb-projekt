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

def uvoz_p(exc):
    stran=exc.sheet_by_name('List1')
    ime=stran.cell_value(0, 9)
    #c=baza.cursor()
    #c.execute('''SELECT registrska FROM tovornjak WHERE ime=?''', [ime])
    #registrska=tuple(c)[0][0]
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

        kolicina=stran.cell_value(i, 7)
        cena=stran.cell_value(i, 6)
        kilometri=stran.cell_value(46,5)
        datum=xldate_as_tuple(stran.cell_value(i, 2), 0)
        datum=str(datum[2])+"."+str(datum[1])+"."+str(datum[0])
        
        c=baza.cursor()

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
        c.execute("""INSERT INTO prevoz(datum, kolicina, cena_tone, zacetek, konec, registrska) VALUES (?, ?, ?, ?, ?, ?)""", (datum, kolicina, cena, mesto1, mesto2, ime))
    baza.commit()





def dobi_uporabnika(auto_login = True):
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
    # Vrnemo predlogo za glavno stran
    return bottle.template("main.html", uporabnik=uporabnik)


@bottle.get("/prijava/")
def prijava():
    return bottle.template("prijava.html", napaka=None, username=None)


@bottle.post("/prijava/")
def login_post():
    uporabnik = bottle.request.forms.username
    password = password_md5(bottle.request.forms.password)
    c = baza.cursor()
    c.execute("SELECT 1 FROM uporabnik WHERE uporabnik=? AND geslo=?",
              [uporabnik, password])
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
    return bottle.template("registracija.html", napaka=None, username=None)

    
@bottle.post("/registracija/")
def register_post():
    """Registriraj novega uporabnika."""
    username = bottle.request.forms.username
    ime = bottle.request.forms.ime
    password1 = bottle.request.forms.password1
    password2 = bottle.request.forms.password2
    # Ali uporabnik že obstaja?
    c = baza.cursor()
    c.execute("SELECT 1 FROM uporabnik WHERE uporabnik=?", [username])
    if c.fetchone():
        # Uporabnik že obstaja
        return bottle.template("registracija.html",
                               username=username,
                               ime=ime,
                               napaka='To uporabniško ime je že zavzeto')
    elif not password1 == password2:
        # Geslo se ne ujemata
        return bottle.template("registracija.html",
                               username=username,
                               ime=ime,
                               napaka='Gesli se ne ujemata')
    else:
        # Vse je v redu, vstavi novega uporabnika v bazo
        password = password_md5(password1)
        c.execute("INSERT INTO uporabnik (uporabnik, geslo) VALUES (?, ?)",
                  (username, password))
        # Daj uporabniku cookie
        #bottle.response.set_cookie('username', username, path='/', secret=secret)
        bottle.redirect("/")


@bottle.route("/pregled/")
def pregled():
    c=baza.cursor()
    c.execute('''SELECT * FROM tovornjak''')
    imen=tuple(c)
    imena={}
    for one in imen:
        imena[one[3]]=one[4]
    b=baza.cursor()
    b.execute('''SELECT strftime('%m', datum) as MESEC FROM prevoz''')
    print(tuple(b))
    return bottle.template("pregled.html", imena=imena)

@bottle.route("/pregled/<voznik>/")
def pregled(voznik):
    c=baza.cursor()
    c.execute("SELECT * FROM prevoz WHERE registrska=?", [voznik])
    c=tuple(c)
    print(c)
    return bottle.template("voznik.html", voznik=voznik, podatki=c)

@bottle.get("/uvoz/")
def uvoz1():
#prikaži formo za vnos datoteke
    return bottle.template("uvoz.html", akcija=None)

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

