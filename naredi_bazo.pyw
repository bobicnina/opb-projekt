import sqlite3

# Datoteka, v kateri je baza
BAZA = "prevoznistvo.sqlite3"

# Naredimo povezavo z bazo. Funkcija sqlite3.connect vrne objekt,
# ki hrani podatke o povezavi z bazo.
baza = sqlite3.connect(BAZA)

baza.execute('''CREATE TABLE IF NOT EXISTS tovornjak (
  registrska TEXT PRIMARY KEY  collate nocase,
  nosilnost INTEGER NOT NULL,
  datum_rojstva DATE NOT NULL,
  ime TEXT NOT NULL collate nocase,
  priimek TEXT NOT NULL collate nocase)''')

baza.execute('''CREATE TABLE IF NOT EXISTS mesta (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ime TEXT collate nocase,
  razdalja INTEGER)''')

baza.execute('''CREATE TABLE IF NOT EXISTS prevoz (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  datum DATE NOT NULL,
  registrska TEXT REFERENCES tovornjak(registrska),
  kolicina INTEGER,
  cena_tone INTEGER,
  zacetek INTEGER REFERENCES mesta(id),
  konec INTEGER REFERENCES mesta(id))''')

baza.execute('''CREATE TABLE IF NOT EXISTS mesecni_stroski (
  registrska TEXT REFERENCES tovornjak(registrska),
  mesec TEXT,
  leto TEXT,
  gorivo INTEGER,
  popravilo INTEGER,
  razdalja INTEGER,
  PRIMARY KEY(registrska, mesec))''')

baza.execute('''CREATE TABLE IF NOT EXISTS cestnina (
  registrska TEXT REFERENCES tovornjak(registrska),
  mesec TEXT NOT NULL,
  italija INTEGER CHECK(italija >= 0) DEFAULT 0,
  slovenija INTEGER CHECK(slovenija >= 0) DEFAULT 0,
  madzarska INTEGER CHECK(madzarska >= 0) DEFAULT 0,
  hrvaska INTEGER CHECK(hrvaska >= 0) DEFAULT 0,
  avstrija INTEGER CHECK(avstrija >= 0) DEFAULT 0,
  PRIMARY KEY(registrska, mesec))''')

baza.execute('''CREATE VIEW IF NOT EXISTS cestnina_na_mesec
AS SELECT
  SUM(italija) AS italjanska,
  SUM(slovenija) AS slovenska,
  SUM(madzarska) AS madzarska,
  SUM(hrvaska) AS hrvaska,
  SUM(avstrija) AS avstrijska
  FROM cestnina GROUP BY mesec''')

baza.execute('''CREATE VIEW IF NOT EXISTS cestnina_na_tovornjak
AS SELECT registrska,
  SUM(italija) AS italjanska,
  SUM(slovenija) AS slovenska,
  SUM(madzarska) AS madzarska,
  SUM(hrvaska) AS hrvaska,
  SUM(avstrija) AS avstrijska,
  SUM(italija) + SUM(slovenija) + SUM(madzarska) + SUM(hrvaska) + SUM(avstrija) AS skupaj
  FROM cestnina GROUP BY mesec, registrska''')

baza.execute('''CREATE VIEW IF NOT EXISTS stroski_na_tovornjak
AS SELECT
  SUM(gorivo) AS gorivo_na_mesec,
  SUM(skupaj) AS cestnina_na_mesec
  FROM mesecni_stroski JOIN cestnina_na_tovornjak ON mesecni_stroski.registrska = cestnina_na_tovornjak.registrska
  GROUP BY mesecni_stroski.registrska''')
  
baza.execute('''CREATE TABLE IF NOT EXISTS uporabnik (
  uporabnik TEXT PRIMARY KEY,
  geslo TEXT NOT NULL)''')

baza.commit()
