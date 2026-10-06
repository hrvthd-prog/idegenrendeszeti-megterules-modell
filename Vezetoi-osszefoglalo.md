# Idegenrendészeti ügyintézés házon belül — megtérülési kimutatás

**Készítette:** ⟨név⟩ · **Dátum:** ⟨dátum⟩ · **Adatok forrása:** ⟨szolgáltatói árlista / számlák⟩ + ⟨ügyintézési napló⟩

---

## A kérdés

Megtérül-e tisztán pénzügyileg, hogy az idegenrendészeti ügyintézést házon belül végezzük, szemben a korábbi külső szolgáltatói modellel?

## A válasz

| | Évesített (1 év) | Tény (elmúlt 12 hónap) |
|---|---|---|
| Elkerült külső szolgáltatási díj | ⟨… Ft⟩ | ⟨… Ft⟩ |
| Teljes munkáltatói költség | −⟨… Ft⟩ | −⟨… Ft⟩ |
| **Nettó megtakarítás** | **⟨… Ft⟩** | **⟨… Ft⟩** |
| Fedezet (elkerült díj / bérköltség) | ⟨…%⟩ | ⟨…%⟩ |

**Hároméves kumulált nettó megtakarítás: ⟨… Ft⟩** (⟨…%⟩ éves bér- és díjindexálással)

> Egy mondatban: a pozíció a saját teljes munkáltatói költségét ⟨…⟩-szorosan fedezi, és évi ⟨… Ft⟩ nettó megtakarítást hoz.

## Honnan jön a szám

| Ügytípus | Évesített ügyszám | Külső díj / ügy | Elkerült díj / év |
|---|---|---|---|
| Tartózkodási engedély — új kérelem | ⟨…⟩ | ⟨… Ft⟩ | ⟨… Ft⟩ |
| Tartózkodási engedély — hosszabbítás | ⟨…⟩ | ⟨… Ft⟩ | ⟨… Ft⟩ |
| Lakcím / szálláshely bejelentés | ⟨…⟩ | ⟨… Ft⟩ | ⟨… Ft⟩ |
| Adószám / TAJ igénylés | ⟨…⟩ | ⟨… Ft⟩ | ⟨… Ft⟩ |
| Egyéb hatósági ügyintézés | ⟨…⟩ | ⟨… Ft⟩ | ⟨… Ft⟩ |
| **Összesen** | **⟨…⟩** | — | **⟨… Ft⟩** |

Összesen ⟨…⟩ külföldi munkavállaló ügyeit kezeljük; közülük ⟨…⟩ fő évente, ⟨…⟩ fő 2–4 évente hosszabbít.

## Módszertan — mit NEM számoltunk bele

Ez a kimutatás szándékosan konzervatív. A következő tényezők nem szerepelnek benne, pedig mind a házon belüli modell mellett szólnak:

- **Eljárási illeték.** Az államnak fizetett illeték ki van véve, mert külső szolgáltató esetén is fizetni kellene. Évesítve ⟨… Ft⟩ — ez semleges tétel, nem megtakarítás.
- **Elkerült bírság és jogszabálysértési kockázat.** Lejárt engedéllyel történő foglalkoztatás szankciója nincs beszámítva.
- **Felszabadult HR- és vezetői idő.** A külső szolgáltató koordinálása korábban belső munkaórákat vitt el.
- **Rövidebb átfutási idő.** A gyorsabb ügyintézés korábbi munkakezdést jelent, ez nincs pénzre fordítva.

A költségoldalon viszont a **teljes** munkáltatói költség szerepel: bruttó bér + szociális hozzájárulási adó + egyéb munkáltatói ráfordítás. Nem csak a bruttó bér.

**Évesítési szabály:** a 2–4 évente hosszabbítandó ügyeket nem a tényleges évükben számoljuk el, hanem elosztjuk a ciklus hosszával (pl. 3 éves ciklusnál az érintett létszám egyharmada esik egy évre). Így a kimutatás nem egy kedvezően kiválasztott évet mutat.

## Javaslat

A házon belüli ügyintézés fenntartása. A modell a mellékelt Excel-munkafüzetben bármikor újraszámolható friss adatokkal; a bemeneti adatok a szolgáltatói árlistával, az ügyintézési naplóval és a bérszámfejtési adatokkal alátámasztottak.

---

### Kitöltési segédlet (ezt a részt törölheted a leadás előtt)

| Helykitöltő | Honnan vedd az Excelből |
|---|---|
| Elkerült díj, évesített | `Összegzés` B12 |
| Teljes munkáltatói költség | `Összegzés` B9 |
| Nettó megtakarítás, évesített | `Összegzés` B14 |
| Fedezet % | `Összegzés` B16 |
| Nettó megtakarítás, tény 12 hó | `Összegzés` B22 |
| 3 éves kumulált nettó | `Összegzés` D32 |
| Évesített ügyszám összesen | `Adatok` D20 |
| Illeték összesen (tájékoztató) | `Adatok` I20 |
| „…-szorosan fedezi” | `Összegzés` B16 értéke (pl. 180% → „1,8-szorosan”) |
