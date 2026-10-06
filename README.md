# Idegenrendészeti ügyintézés — megtérülési modell

Üres, kitölthető Excel-modell és a hozzá tartozó szöveges anyagok ahhoz, hogy kimutatható legyen: megéri-e pénzügyileg házon belül végezni a külföldi munkavállalók idegenrendészeti ügyintézését, a korábbi külső szolgáltatói modellhez képest.

**A repó nem tartalmaz valós adatot.** Az Excel minden bemeneti cellája üres; a számokat a felhasználó tölti ki helyben.

## Tartalom

| Fájl | Leírás |
|---|---|
| `Megterulesi-kimutatas-idegenrendeszet.xlsx` | A modell. Sárga cella = kitöltendő, szürke = számított, zöld = végeredmény. |
| `Vezetoi-osszefoglalo.md` | 1 oldalas vezetői anyag, helykitöltőkkel + cellahivatkozási táblával. |
| `Prezentacio-vazlat.md` | 5 dia tartalma, diagramjavaslatokkal és a várható kérdésekre adható válaszokkal. |
| `build.py` | A munkafüzetet generáló script (openpyxl). |
| `selfcheck.py` | Képletellenőrzés Excel COM-on át; próbaszámokkal számol, majd mentés nélkül zár. |

## Az Excel szerkezete

- **Útmutató** — használat, módszertani döntések, színkód.
- **Adatok** — soronként egy ügytípus: érintett fő, hosszabbítási ciklus, külső szolgáltatói díj / ügy. Az évesített ügyszám és az elkerült díj számított.
- **Összegzés** — teljes munkáltatói költség, évesített eredmény (nettó megtakarítás, megtérülés, fedezet, fedezeti ügyszám), tényleges 12 hónapos kontrollnézet, hároméves kumulált hatás diagrammal.

## Módszertan

- **Évesítés** — a 2–4 évente hosszabbítandó ügyek nem a tényleges évükben számolódnak el, hanem elosztva a ciklus hosszával (9 fő / 3 éves ciklus = 3 ügy/év). Így a kimutatás nem egy kedvezően kiválasztott évet mutat.
- **Az eljárási illeték ki van véve** a megtakarításból, mert belső és külső ügyintézés esetén egyaránt fizetendő. Csak tájékoztató oszlopként szerepel.
- **Költségoldal: teljes munkáltatói költség** — bruttó bér + szociális hozzájárulási adó + egyéb munkáltatói ráfordítás, nem csak a bruttó bér.
- **Csak pénzbeli tételek** — elkerült bírság, felszabaduló HR-idő és rövidebb átfutási idő nincs beszámítva, csak érvként említve.

## Újragenerálás

```sh
pip install openpyxl
python build.py        # legenerálja az üres munkafüzetet
python selfcheck.py    # Windows + Excel szükséges; ellenőrzi a képleteket
```

## Licenc

MIT
