-- ================================================================
-- MPMTransport - Tabele za upravljanje zalihama
-- Access SQL (Jet 4.0 sintaksa)
-- ================================================================

-- Tabela: zalihe (artikli u magacinu)
CREATE TABLE zalihe (
    id_artikla         AUTOINCREMENT PRIMARY KEY,
    naziv              TEXT(100) NOT NULL,
    kategorija         TEXT(50),
    jedinica_mere      TEXT(20),
    kolicina_na_stanju INTEGER DEFAULT 0,
    min_kolicina       INTEGER DEFAULT 0,
    lokacija_u_magacinu TEXT(50),
    opis               MEMO,
    datum_azuriranja   DATETIME
);

-- Tabela: dobavljaci
CREATE TABLE dobavljaci (
    id_dobavljaca  AUTOINCREMENT PRIMARY KEY,
    naziv_firme    TEXT(100) NOT NULL,
    kontakt        TEXT(100),
    adresa         TEXT(200),
    pib            TEXT(20),
    email          TEXT(100),
    telefon        TEXT(30)
);

-- Tabela: nalog_nabavka (narudzbenice prema dobavljacima)
CREATE TABLE nalog_nabavka (
    id_naloga       AUTOINCREMENT PRIMARY KEY,
    id_dobavljaca   INTEGER NOT NULL,
    id_artikla      INTEGER NOT NULL,
    kolicina        INTEGER NOT NULL,
    cena_po_komadu  CURRENCY,
    ukupna_vrednost CURRENCY,
    datum_naloga    DATETIME,
    datum_isporuke  DATETIME,
    status          TEXT(30),   -- 'U obradi', 'Primljeno', 'Otkazano'
    napomena        MEMO
);

-- Tabela: nalog_izdavanje (izdavanje robe vozilima)
CREATE TABLE nalog_izdavanje (
    id_izdavanja    AUTOINCREMENT PRIMARY KEY,
    id_vozila       INTEGER NOT NULL,
    id_artikla      INTEGER NOT NULL,
    kolicina        INTEGER NOT NULL,
    datum_izdavanja DATETIME,
    status          TEXT(30),   -- 'Izdato', 'Ceka preuzimanje', 'Otkazano'
    napomena        MEMO
);

-- ================================================================
-- Primer pocetnih podataka
-- ================================================================

INSERT INTO dobavljaci (naziv_firme, kontakt, adresa, pib, email, telefon)
VALUES
  ('Petro DOO', 'Marko Petrovic', 'Beograd, Vojvode Misica 12', '101234567', 'info@petro.rs', '011-123-4567'),
  ('AutoParts SRB', 'Nikola Jovic', 'Novi Sad, Bulevar Oslobodenja 45', '202345678', 'prodaja@autoparts.rs', '021-234-5678'),
  ('GumaTrans', 'Ana Milic', 'Kragujevac, Sumadijska 8', '303456789', 'kontakt@gumatrans.rs', '034-345-6789');

INSERT INTO zalihe (naziv, kategorija, jedinica_mere, kolicina_na_stanju, min_kolicina, lokacija_u_magacinu)
VALUES
  ('Filter ulja Mann W940',    'Rezervni delovi', 'kom.', 42, 20, 'Reg. A-12'),
  ('Motorno ulje 10W-40 5L',   'Maziva i goriva', 'lit.', 28, 50, 'Reg. B-03'),
  ('Guma 315/80 R22.5',        'Gume i felne',    'kom.',  4,  8, 'Reg. C-01'),
  ('Kocione plocice Brembo',   'Rezervni delovi', 'kom.', 16, 10, 'Reg. A-07'),
  ('Rashladna tecnost 1L',     'Maziva i goriva', 'lit.', 12, 20, 'Reg. B-05'),
  ('Sijalica H7 12V 55W',      'Rezervni delovi', 'kom.', 30,  8, 'Reg. A-15'),
  ('Remacnik upravljaca',      'Rezervni delovi', 'kom.',  6,  4, 'Reg. A-22'),
  ('AdBlue 10L',               'Maziva i goriva', 'lit.', 80, 40, 'Reg. B-10');
