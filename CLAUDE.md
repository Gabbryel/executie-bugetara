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
Coloanele 3, 5, 6, 8 sunt **cumulate de la 1 ianuarie**; coloana 7 e doar luna. În ianuarie baza e „început de an"
(zero), nu decembrie: `incarcaPrev` înlocuiește o lună precedentă din anul trecut cu `bazaInceputAn` (luna 0).

Corelații verificate pe două luni reale (aprilie și iulie 2026):

- `BC39 Bugetar (lei) = rând 8 col. 3 × 1000`; `BC39 Angajament (lei) = rând 8 col. 4 × 1000`. Exact, fără rotunjiri.
- Rândul 8 col. 3 = col. 5 = col. 6 (creditele se consumă integral).
- Col. 8 − col. 6 = facturi înregistrate și neplătite; decalajul e stabil de la o lună la alta.
- Col. 7 = col. 6 − col. 6 din luna precedentă, exact, pe toate rândurile (iunie → iulie → august 2026).
- Col. 4 poate fi **și sub, și peste** col. 3 — nu presupune un semn fix. Pe rânduri, contabilul o **reface
  lunar din col. 3**, cu 3–4 rânduri rotunjite în sus la mii întregi și unul care primește restul (verificat
  iunie/iulie 2026); rotunjirile nu se moștenesc de la o lună la alta. „repartizează" din rândul-țintă face la fel
  (`repartizeazaCol4`: rândurile mari întâi, restul pe cel mai mare). De aceea în modul delta col. 4 urmează col. 3 pe fiecare
  rând (`din.f == null`, celulă cu chenar punctat) până tastezi o valoare; golirea celulei o readuce la automat.
  Ciornele cu `versiune < 2` copiau col. 4 din luna precedentă; `restore` le migrează.
- Col. 8 (cheltuieli efective) = col. 6 pe rândurile de personal, hrană, capital; diferă doar pe bunuri și servicii
  (r.44, 45, 52, 53, 61, 62, 64, 112, 116 în iulie 2026), unde sunt facturi înregistrate și neplătite. În delta,
  Δ col. 8 urmează col. 7 (`din.dj == null`) până e tastată; ciornele cu `versiune < 3` aveau `dj = 0` = netastat.
- Rândul 83 (programe naționale de sănătate) e 0: unitatea nu derulează PNS.
- Formatul BC39 variază între luni (o coloană `Spital` sau două: `Total` + `Spital (inclusiv PNS)`).
  Fișa are și `An`, `Luna (cumulat)`, `Denumire` (ex. „SPINAL CARE SRL"), `Cod`; rândul `Angajament` vine
  înaintea celui `Bugetar`. Referința pentru rândul 8 este coloana `Total` când există (r.8 include PNS).
- Macheta de aprilie 2026 are o inconsecvență în fișierul sursă: r.32 col. 4 diferă de suma copiilor, iar BC39
  urmează totalul scris în fișier (r.8), nu suma rândurilor. Aplicația agregă de jos în sus, deci afișează alt
  total și semnalează diferența — corect, nu e o eroare a ei (`tests/reale.test.js` o tratează ca excepție cunoscută).
- În documentele publice (acest fișier, README, commit-uri) nu se scriu sume reale — doar reguli.

Detalii complete: `README.md`.

## Arhitectura

```
index.html                 interfața (fără framework)
assets/app.js              interfața: tabel, ghidare, lista de control, importuri, salvare
assets/calc.js             calculul fără DOM: agregare, delta, luna precedentă, repartizarea col. 4 (testat)
assets/rows.js             GENERAT din formulele coloanei E ale machetei — nu se editează manual
assets/xlsx.js             citire/scriere .xlsx prin editarea XML-ului din pachet (JSZip)
assets/store.js            GitHub Contents API (depozit = bază de date) + localStorage + folder pe disc
assets/num.js              parsare/formatare numerică românească (virgulă zecimală)
assets/jszip.min.js        JSZip 3.10.1 vendorizat — fără CDN
template/macheta.xlsx      formularul CAS original (decembrie 2025, `#,##0.000`), golit, fără autor
tools/genereaza-rows.mjs   rows.js din șablon (fără argumente: verifică; --scrie: rescrie)
tools/pregateste-sablon.mjs șablonul dintr-o machetă CAS nouă (golește, corectează, fullCalcOnLoad)
tools/jszip-node.mjs       JSZip vendorizat încărcat în Node, fără npm
tests/*.test.js            teste `node --test`, fără instalare; reale.test.js folosește MACHETE_DIR
tests/smoke.cjs            test end-to-end cu Playwright (opțional)
```

Modelul rândurilor: fiecare rând e frunză (se completează) sau sinteză (`ch` = lista rândurilor-copil).
Toți copiii au numere de rând mai mari decât părintele, deci agregarea se face parcurgând rândurile
descrescător. Rândul 38 („Contribuții, din care:") e static în șablon, dar în model e sinteza rândului 39.

## Invarianți — nu le încălca

1. **Exportul .xlsx editează șablonul, nu construiește un fișier nou.** `xlsx.js` scrie valorile în celulele
   fără `<f>`; în celulele cu formulă adaugă doar valoarea calculată `<v>` imediat după `<f>` (ca Quick Look și
   previzualizările din e-mail să arate totalurile), formula rămâne neatinsă, iar Excel recalculează oricum
   (`fullCalcOnLoad`). Stilurile, `printerSettings` și înălțimile de rând rămân bit-identice.
   Nu înlocui cu SheetJS/ExcelJS — pierd formatarea și au fost respinse exact din acest motiv.
2. **`rows.js` se regenerează, nu se editează.** Când CAS trimite o machetă nouă:
   `node tools/pregateste-sablon.mjs "<macheta CAS>.xlsx" template/macheta.xlsx` (golește valorile, scoate
   formulele de pe rândurile-frunză, reface formulele de sinteză abătute de la coloana E, ex. I117), apoi
   `node tools/genereaza-rows.mjs` (și `--scrie` dacă s-a schimbat structura) și `npm test`.
   Șablonul actual = formularul CAS gol din decembrie 2025 trecut prin unealtă (are 0 abateri; cele trei
   corectate — I117, H54/H63, J50/J63 — au apărut în machetele contabilului din 2026). Unealta scoate și autorul
   din `docProps/core.xml`: șablonul e public, iar fiecare export îl moștenea.
3. **Modul delta depinde de luna precedentă.** Ciorna, arhiva, „Export JSON" și GitHub salvează și `prev`
   (cu `an`, `luna`); `restore` golește luna precedentă încărcată înainte și ia doar ce e în fișier. La restaurare
   fără `prev`, aplicația trece pe `cumulat` și păstrează valorile. Nu schimba această cădere.
4. **Numerele sunt mii lei cu 3 zecimale.** Rotunjire `r3()` la fiecare agregare; comparațiile cu BC39
   se fac pe întregi (`Math.round(x*1000)`).
5. **Nimic identificabil în depozitul public în afara numelui unității.** Datele reale
   (`data/AAAA-LL.json`) merg într-un depozit privat separat. Fișierele-exemplu (machete reale) **nu se
   commit-uiesc** — testele le primesc prin variabile de mediu.
6. **Fără dependențe externe la runtime.** Tot ce se încarcă vine din depozit; aplicația merge offline.
7. **Originea contează.** Pe `gabbryel.github.io` toate site-urile Pages ale contului împart localStorage și
   IndexedDB (alte site-uri ale contului încarcă scripturi de pe CDN-uri). Tokenul GitHub și arhiva sunt în
   siguranță doar pe un domeniu propriu; dialogul GitHub avertizează când rulează pe `*.github.io`.

## Salvarea lunilor

O singură acțiune, „Salvează luna" (sau Ctrl/Cmd+S): pune snapshot-ul (cu luna precedentă și referința BC39)
în arhiva din browser (`store.luni`, cheie `AAAA-LL` în `localStorage['execbug.luni']`) și, dacă GitHub e
configurat cu token, îl trimite și în depozit (`saveMonth`). Butonul arată un punct galben când starea diferă
de versiunea salvată a lunii curente (`amprenta()` vs `amprentaSalvata`). Meniul ⋯ → „Luni salvate…" listează
arhiva: Deschide (restaurează pentru modificări; salvarea suprascrie acea lună), Ca lună precedentă, JSON,
Șterge. Ciorna (`execbug.draft`) rămâne separată: e starea de lucru, arhiva e ce ai salvat explicit. Ciorna ține și
`cheie` = luna din arhivă din care provine formularul (`cheieDeschisa`).

Protecții la salvare și deschidere (dialog în pagină, `intreaba()`, nu `confirm()`):
- deschiderea altei luni, importul ca lună curentă, GitHub „Deschide" și „Formular nou" întreabă când există
  modificări nesalvate (`inainteDeInlocuire`): salvează și continuă / continuă fără / anulează;
- „Salvează luna" peste o lună existentă din care formularul nu provine (ex. după schimbarea lunii din antet)
  cere confirmare;
- la deschiderea unei luni în delta, dacă luna ei precedentă a fost salvată din nou cu alte valori
  (`verificaBazaActualizata`), propune recalcularea pe versiunea salvată;
- „Descarcă toate lunile" / „Importă lunile din fișier…" (dialogul Luni salvate) mută arhiva între browsere sau
  adrese; la import, o lună existentă se înlocuiește doar cu o versiune salvată mai târziu.
- O ciornă deteriorată nu blochează pornirea: `curataStare` aduce tipurile la valori valide, iar dacă tot nu se
  poate citi, se mută în `execbug.draft.deteriorata` și pornește un formular gol.

**Folder de salvare pe disc** (`store.folder`, File System Access API, doar Chrome/Edge): meniul ⋯ →
„Folder de salvare…" deschide selectorul nativ o singură dată; handle-ul se ține în IndexedDB
(`execbug/kv/folder`). „Salvează luna" scrie apoi `AAAA-LL.json` și în folder, fără niciun token; permisiunea
de scriere se cere la primul clic după repornirea browserului (necesită gest al utilizatorului, de aceea nu se
cere la încărcare). La listare și deschidere, fișierul de pe disc are prioritate față de arhiva din browser,
pentru că poate fi modificat și din altă parte (folder sincronizat cu OneDrive/Drive). În browsere fără suport,
opțiunea e ascunsă și totul rămâne în arhiva din browser. GitHub cu token rămâne opțional, în plus.
„Export .xlsx" scrie macheta și în același folder (`scrieInFolder`), pe lângă descărcare. PDF-ul nu poate fi scris
programatic (îl produce browserul din „Tipărire / PDF"): la tipărire titlul paginii devine numele machetei, ca
„Salvează ca PDF" să-l propună; folderul îl alege utilizatorul în dialog, iar Chrome îl ține minte pe ultimul.

## Versionarea fișierelor (cache)

GitHub Pages și browserul țin în cache `app.js`, modulele și CSS-ul. Toate adresele lor poartă `?v=<data-ora>`
(`index.html` și importurile din `app.js`/`xlsx.js`). `tools/versiune.sh` pune o valoare nouă peste tot;
`.githooks/pre-commit` îl rulează la fiecare commit. **La o clonă nouă**: `git config core.hooksPath .githooks`.
Versiunea încărcată se vede în meniul ⋯, jos („versiune 20260913-101500"), ca să confirmi că ai codul nou.
Nu adăuga importuri fără `?v=0` (scriptul înlocuiește orice valoare, dar trebuie să existe sufixul).

## Cum rulezi și testezi

```bash
npm test                               # = node --test tests/*.test.js — Node 22+, fără instalare
MACHETE_DIR=~/priv/machete npm test    # și pe machetele reale (în afara depozitului)
python3 -m http.server 8899            # serverul static; modulele ES nu merg din file://
```

`npm test` acoperă: parsarea numerelor, arborele de agregare, `rows.js` = generat din șablon, modul delta
(inclusiv comutarea dus-întors), luna precedentă și ianuarie, col. 7, repartizarea col. 4, export → import cu
formulele neatinse și totalurile scrise, citirea BC39 (antet mutat, sume ca text), idempotența pregătirii
șablonului. Cu `MACHETE_DIR` (fișiere „Macheta Executie spital <LUNA> <AN>.xlsx" și „BC39-<luna>.xlsx"):
agregarea = totalul din fișier, col. 7 = Δ col. 6 pe luni consecutive, BC39 = r.8 × 1000, export dus-întors,
șablonul = macheta din iulie pregătită.

Testul de interfață cu Playwright e opțional (`npm i -D playwright`, apoi `npm run smoke`); acceptă
`MACHETA_CUR`, `MACHETA_PREV` (luna dinainte!) și `BC39` ca fișiere reale.

## Lista de control (implementată în `runChecks`)

1. corelația BC39 · 2. r.8 col.3 = col.5 = col.6 · 3. cumulat monoton față de luna precedentă ·
4. col.6 ≤ col.4 pe fiecare rând · 5. col.8 ≥ col.6, cu decalajul accrual–cash ·
6. integritatea șablonului (`I117` = `=+I119+I118`, nu `=H117`) · 7. formatul BC39 ·
8. luna precedentă = luna dinainte (`starePrecedenta`; în ianuarie „început de an"), prima în listă pentru că
o bază greșită strică toate cumulatele; are acțiunea „Pornește de la zero" (`act:'zero'`) ·
9. în cumulat, col. 7 = col. 6 − col. 6 din luna precedentă (`col7Inconsecvent`; în ianuarie față de zero).

Când adaugi o verificare: `chk(ok, titlu, detaliu, warnIfFail, scurt)`; `ok === null` = informativ; un câmp `act`
pune în linia de ghidare un buton (cheile din `ACTIUNI`). Cu `scurt`,
verificarea trecută nu apare ca element propriu, ci în linia „Trec fără observații" (2, 3, 4, 6, 7 sunt tautologice
în situația obișnuită: coloane legate, delta, șablon corect). Verificarea 5 e informativă când nicio Δ nu e
tastată (col. 8 urmează col. 7) și listează rândurile cu facturi neplătite din luna precedentă. Verificările 4 și 5 sunt
avertismente, nu erori: machetele reale le-au încălcat pe rânduri (aprilie 2026, r.52 și r.53/r.79).

## Layout și ghidare

Principiu: **la vedere stau doar tabelul și exportul**; restul e ascuns până e nevoie.

- Antet pe o linie: unitatea, luna și anul ca text editabil (`.plain`), căutare, butonul „Verificări" cu
  numărul de probleme, „Export .xlsx" și meniul `⋯` (importuri, export JSON/PDF, GitHub, opțiuni, temă,
  formular nou). Comutatorul cumulat/delta apare doar după încărcarea lunii precedente.
- **Linia de ghidare** (`#guide`, calculată de `guidance()` la fiecare randare) spune mereu ce urmează:
  de unde începi, ce țintă are rândul 8 față de BC39 și cât lipsește, prima verificare picată, sau că
  totul se corelează și poți exporta. Este singurul loc unde utilizatorul primește feedback fără să
  deschidă ceva; orice verificare nouă trebuie să aibă un `det` formulat ca acțiune, nu ca stare.
- Lista completă de control și sinteza stau în panoul `.side`, sertar suprapus din dreapta, închis implicit
  (butonul „Verificări" sau linkul din ghidare). Escape / clic în afară îl închide.
- Tabelul e singura zonă care derulează; antetul și rândul 8 sunt lipicioase. Celulele col. 3/4 ale
  rândului 8 se colorează roșu/verde față de BC39.
- **Rândul „Țintă BC39"** (`renderTinta`, al doilea rând din `thead`, lipicios sub antet): după import, valorile
  din fișă apar în tabel exact în coloanele lor (col. 3 = Bugetar, col. 4 = Angajament, col. 5 și 6 = Bugetar,
  col. 7 = țintă − luna precedentă), cu „rămân X" / „✓ atins" sub fiecare. Utilizatorul se așteaptă să vadă
  cifrele BC39 în primele rânduri ale machetei, nu doar într-o propoziție.
- **La importul BC39** (`setBc39` + verificarea 1): se citesc luna, anul și denumirea; formularul gol preia
  luna/anul din fișă; dacă luna diferă de a formularului, verificarea e eroare (nu se compară cifre din luni
  diferite); denumirea se potrivește pe cuvinte-cheie (fără SC/SRL); cu lună precedentă încărcată, ținta
  arată și „cu X peste <luna precedentă>". Referința BC39 se salvează în ciornă/JSON (`bc39ref`).
- **La importul unei machete** (`verificaConsecventa`): totalurile statice din fișier se compară cu suma
  rândurilor-copil; diferențele apar ca verificare informativă (`importNote`) până la următorul import.
- Opțiunea „leagă col. 3 și 5 de col. 6" se aplică în ambele moduri; la importul unei machete în care
  col. 3/5 diferă de col. 6 pe rânduri, se dezleagă automat ca să nu suprascrie valorile.
- „repartizează" apare în rândul-țintă sub col. 4 când lipsesc credite de angajament; în ianuarie fără lună
  precedentă, linia de ghidare oferă „Pornește de la zero".
- Lipirea din Excel a mai multor celule (`onPaste`) le așază pe rândurile machetei de la celula selectată în jos
  și pe coloanele din dreapta; rândurile de total și celulele calculate se sar, iar mesajul spune câte.
- Căutarea ignoră diacriticele și caută și după cod (sub-capitol, titlu, alineat).
- Sub 760px pagina derulează normal. Printarea ascunde tot în afara tabelului și arată **toate** rândurile,
  indiferent de căutare sau „ascunde zero".

Rândurile cu text lung în coloanele de cod (ex. r.110, unde macheta ține denumirea în coloana „Alin.")
se afișează cu textul mutat în coloana de denumire — `buildTable` face asta, `rows.js` rămâne neatins.

## Stil

- Vanilla JS, module ES, fără bundler, fără framework. Un singur fișier CSS cu variabile de temă
  (light/dark prin `data-theme` și `prefers-color-scheme`).
- Șirurile din UI în română cu diacritice. Numerele afișate cu `toLocaleString('ro-RO')`.
- Commit-uri scurte, în română, la timpul prezent („Adaugă filtrul pe rânduri").
