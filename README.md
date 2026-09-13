# Execuție bugetară — machetă CAS

Aplicație statică (fără server) pentru completarea machetei lunare de raportare a plăților din FNUASS
pe unități sanitare cu paturi, cu validare în timp real și export în formatul original `.xlsx`.

Rulează integral în browser. Se publică pe GitHub Pages.

## Ce face

- **Completare asistată** a celor 121 de rânduri (8–128) ale foii `BC`. Rândurile de sinteză se calculează automat după arborele de agregare din machetă; se completează doar frunzele.
- **Mod „delta"** — după ce încarci luna precedentă, completezi doar plățile lunii (col. 7) și delta cheltuielilor efective; coloanele cumulate (3, 5, 6, 8) se calculează singure.
- **Lista de control** rulează la fiecare tastă:
  1. corelația cu fișa BC39 (r.8 col.3 × 1000 = Bugetar, col.4 × 1000 = Angajament);
  2. r.8: col.3 = col.5 = col.6;
  3. cumulat monoton față de luna precedentă;
  4. plăți cumulate ≤ credite de angajament (col.6 ≤ col.4);
  5. cheltuieli efective ≥ plăți (col.8 ≥ col.6), cu decalajul accrual–cash;
  6. integritatea șablonului (formula `I117`);
  7. formatul fișei BC39 (numărul coloanelor de valori diferă între luni).
- **Export `.xlsx` identic cu macheta CAS.** Aplicația nu generează un fișier nou: deschide șablonul original, scrie doar valorile și lasă neatinse formatarea, formulele, înălțimile de rând și setările de tipărire. Verificat prin recalculare: totalurile exportate coincid la a treia zecimală cu fișierele originale.
- **Tipărire / PDF** direct din browser (A4 landscape), cu antetul și rubricile de semnătură.
- **Persistență**: ciornă automată în browser, export/import JSON și, opțional, un depozit GitHub ca bază de date.

## Baza de date în GitHub

GitHub Pages servește doar fișiere statice — nu există server și nici bază de date clasică.
Varianta folosită aici: **depozitul este baza de date**. Fiecare lună se salvează ca `data/AAAA-LL.json`
prin GitHub Contents API, iar istoricul commit-urilor devine jurnal de modificări (cine, când, ce s-a schimbat).

Configurare recomandată — **două depozite**:

| Depozit | Vizibilitate | Conținut |
|---|---|---|
| `executie-bugetara` | public | aplicația, publicată prin GitHub Pages |
| `executie-bugetara-date` | **privat** | doar `data/*.json` — cifrele de execuție |

Datele nu ajung niciodată în depozitul public. Aplicația citește și scrie în cel privat, prin API, cu un token pe care îl introduci tu în ecranul „Bază de date (GitHub)".

### Token

Setări GitHub → Developer settings → Personal access tokens → **Fine-grained tokens**:

- **Repository access**: doar `executie-bugetara-date`
- **Permissions → Repository permissions → Contents**: `Read and write`
- **Expiration**: cât mai scurtă (30–90 zile)

Tokenul se păstrează în `localStorage`, în browserul tău. Oricine are acces la acel profil de browser îl poate extrage — de aceea trebuie limitat la un singur depozit și reînnoit periodic. Fără token, aplicația funcționează complet local (import/export de fișiere); citirea fără token merge doar dintr-un depozit public.

Dacă preferi să nu ții niciun token în browser: lasă secțiunea GitHub necompletată și folosește **Export JSON** + commit manual în depozitul de date. Rezultatul în istoric este identic.

## Publicare pe GitHub Pages

```bash
git init
git add .
git commit -m "Aplicație execuție bugetară"
git branch -M main
git remote add origin git@github.com:<user>/executie-bugetara.git
git push -u origin main
```

Apoi: **Settings → Pages → Source: Deploy from a branch → `main` / `(root)`**.

O singură dată, după clonare: `git config core.hooksPath .githooks`. Hook-ul pune la fiecare commit o versiune
nouă în adresele fișierelor JS și CSS, altfel browserul continuă să ruleze codul vechi din cache după publicare.
Versiunea încărcată apare în meniul ⋯.
Adresa va fi `https://<user>.github.io/executie-bugetara/`. Toate căile din aplicație sunt relative, deci funcționează și sub un subdirector.

## Structura

```
index.html               interfața
assets/app.js            logica de completare, agregare, validare
assets/xlsx.js           citire/scriere .xlsx prin editarea directă a XML-ului din pachet
assets/store.js          GitHub Contents API + localStorage
assets/rows.js           modelul celor 121 de rânduri, generat din macheta CAS
assets/num.js            parsare/formatare numerică românească
assets/jszip.min.js      JSZip 3.10.1 (local, fără CDN)
template/macheta.xlsx    șablonul CAS, cu valorile golite
```

## Șablonul

`template/macheta.xlsx` este macheta CAS originală, cu trei corecții față de fișierul primit pe iulie 2026:

- `I117` avea `=H117` (coloana „Luna curentă" copia cumulatul) — restaurat la `=+I119+I118`, forma corectă din macheta de aprilie;
- `J50`, `J63` trimiteau la coloana `K`, care nu există în machetă — transformate în celule de valoare;
- `H54`, `H63` aveau formule acolo unde toate celelalte rânduri au valori statice — uniformizate.

Valorile au fost golite, iar cache-ul formulelor șters (`fullCalcOnLoad`), astfel încât Excel recalculează la deschidere.

**Când CAS trimite o machetă nouă**, înlocuiește `template/macheta.xlsx`. Dacă se schimbă și structura rândurilor, trebuie regenerat `assets/rows.js` — modelul este derivat din formulele coloanei E, rândurile 8–128.

## Limite

- Nu semnează digital. Semnarea rămâne pas separat, pe PDF-ul exportat.
- Nu generează fișa BC39; o importă pentru verificare și afișează valorile calculate în lei, pentru comparație.
- Nu validează încadrarea pe articole de clasificație — verifică doar coerența internă și corelațiile descrise mai sus.
