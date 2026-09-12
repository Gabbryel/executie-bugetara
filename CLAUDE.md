# CLAUDE.md — Execuție bugetară (machetă CAS)

Aplicație web statică, fără build și fără server, pentru completarea machetei lunare
„Plățile efectuate din FNUASS pe unitățile sanitare cu paturi" (raportare la CAS Bacău)
pentru SC Spinal Care Dobreci SRL. Publicată pe GitHub Pages. Limba proiectului: **română**
(UI, comentarii, commit-uri).

## Contextul de domeniu (citește înainte de orice modificare)

Trei documente circulă lunar:

| Fișier | Unitate | Rol |
|---|---|---|
| `Macheta Executie spital <LUNA> <AN>.xlsx`, foaia `BC` | **mii lei**, 3 zecimale | macheta propriu-zisă, 121 rânduri de indicatori (rândurile 8–128), 6 coloane de valori (E–J = coloanele 3–8 din antet) |
| același fișier ca PDF | mii lei | export 1:1, semnat digital de manager + director economic, trimis la CAS |
| `BC39-<luna>.xlsx` | **lei** | fișă sintetică cu 2 rânduri: tip credit `Bugetar` și `Angajament` |

Coloanele machetei: **3** credite bugetare aprobate · **4** credite de angajament · **5** credite trimestriale cumulate · **6** plăți nete de casă cumulat · **7** luna curentă · **8** cheltuieli efective (accrual).
Coloanele 3, 5, 6, 8 sunt **cumulate de la 1 ianuarie**; coloana 7 e doar luna.

Corelații verificate pe două luni reale (aprilie și iulie 2026):

- `BC39 Bugetar (lei) = rând 8 col. 3 × 1000`; `BC39 Angajament (lei) = rând 8 col. 4 × 1000`. Exact, fără rotunjiri.
- Rândul 8 col. 3 = col. 5 = col. 6 (creditele se consumă integral).
- Col. 8 − col. 6 = facturi înregistrate și neplătite; decalajul e stabil (~43–44 mii lei).
- Col. 4 poate fi **și sub, și peste** col. 3 — nu presupune un semn fix.
- Rândul 83 (programe naționale de sănătate) e 0: unitatea nu derulează PNS.
- Formatul BC39 variază între luni (o coloană `Spital` sau două: `Total` + `Spital (inclusiv PNS)`).

Detalii complete: `README.md`.

## Arhitectura

```
index.html            interfața (fără framework)
assets/app.js         stare, agregare bottom-up, modurile cumulat/delta, lista de control
assets/rows.js        GENERAT din formulele coloanei E ale machetei — nu se editează manual
assets/xlsx.js        citire/scriere .xlsx prin editarea XML-ului din pachet (JSZip)
assets/store.js       GitHub Contents API (depozit = bază de date) + localStorage
assets/num.js         parsare/formatare numerică românească (virgulă zecimală)
assets/jszip.min.js   JSZip 3.10.1 vendorizat — fără CDN
template/macheta.xlsx șablonul CAS cu valorile golite și 3 formule corectate
tests/smoke.js        test end-to-end cu Playwright
```

Modelul rândurilor: fiecare rând e frunză (se completează) sau sinteză (`ch` = lista rândurilor-copil).
Toți copiii au numere de rând mai mari decât părintele, deci agregarea se face parcurgând rândurile
descrescător. Rândul 38 („Contribuții, din care:") e static în șablon, dar în model e sinteza rândului 39.

## Invarianți — nu le încălca

1. **Exportul .xlsx editează șablonul, nu construiește un fișier nou.** `xlsx.js` scrie `<v>` doar în
   celulele fără `<f>`; formulele, stilurile, `printerSettings` și înălțimile de rând rămân bit-identice.
   Nu înlocui cu SheetJS/ExcelJS — pierd formatarea și au fost respinse exact din acest motiv.
2. **`rows.js` se regenerează, nu se editează.** Când CAS trimite o machetă nouă: înlocuiești
   `template/macheta.xlsx` (după golirea valorilor și corectarea formulelor, vezi README „Șablonul"),
   apoi regenerezi modelul din formulele coloanei E, rândurile 8–128.
3. **Modul delta depinde de luna precedentă.** Ciorna din localStorage salvează și `prev`; la restaurare
   fără `prev`, aplicația trece pe `cumulat` și păstrează valorile. Nu schimba această cădere.
4. **Numerele sunt mii lei cu 3 zecimale.** Rotunjire `r3()` la fiecare agregare; comparațiile cu BC39
   se fac pe întregi (`Math.round(x*1000)`).
5. **Nimic identificabil în depozitul public în afara numelui unității.** Datele reale
   (`data/AAAA-LL.json`) merg într-un depozit privat separat. Fișierele-exemplu (machete reale) **nu se
   commit-uiesc** — testele le primesc prin variabile de mediu.
6. **Fără dependențe externe la runtime.** Tot ce se încarcă vine din depozit; aplicația merge offline.

## Cum rulezi și testezi

```bash
python3 -m http.server 8899            # serverul static; modulele ES nu merg din file://
node tests/smoke.js                    # necesită: npm i -D playwright (Chromium)
```

Testul fără variabile de mediu verifică: pagina se încarcă fără erori, șablonul trece verificarea de
integritate, exportul produce un .xlsx valid. Cu fișiere reale (ținute în afara depozitului):

```bash
MACHETA_CUR=~/priv/Macheta_IULIE.xlsx MACHETA_PREV=~/priv/Macheta_APRILIE.xlsx BC39=~/priv/BC39-iulie.xlsx node tests/smoke.js
```

atunci testul confirmă și că importul reproduce rândul 8 exact, că BC39 se corelează, și că toate
verificările trec. Pentru validarea exportului față de original, recalculează cu LibreOffice:
`soffice --headless --convert-to xlsx --outdir /tmp/recalc export.xlsx` și compară totalurile.

## Lista de control (implementată în `runChecks`)

1. corelația BC39 · 2. r.8 col.3 = col.5 = col.6 · 3. cumulat monoton față de luna precedentă ·
4. col.6 ≤ col.4 pe fiecare rând · 5. col.8 ≥ col.6, cu decalajul accrual–cash ·
6. integritatea șablonului (`I117` = `=+I119+I118`, nu `=H117`) · 7. formatul BC39.

Când adaugi o verificare: `chk(ok, titlu, detaliu, warnIfFail)`; `ok === null` = informativ.

## Stil

- Vanilla JS, module ES, fără bundler, fără framework. Un singur fișier CSS cu variabile de temă
  (light/dark prin `data-theme` și `prefers-color-scheme`).
- Șirurile din UI în română cu diacritice. Numerele afișate cu `toLocaleString('ro-RO')`.
- Commit-uri scurte, în română, la timpul prezent („Adaugă filtrul pe rânduri").
