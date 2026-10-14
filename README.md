# Karma Diagnose

ICD-10-søk for karkirurgi (PWA). Selvstendig program, ikke integrert i Karma Kode.

## Filer
- `index.html` – selve appen (data er innebygd, fungerer offline)
- `data/kar-icd10-synonymer.json` – **kilden til alle koder, synonymer og søkeregler**. Det er bare denne som redigeres.
- `build.py` – limer JSON inn i `index.html` og øker cache-versjonen i `sw.js`
- `manifest.json`, `sw.js`, `icon-180.png` – PWA (installerbar, offline)

## Arbeidsflyt
1. Endre `data/kar-icd10-synonymer.json`
2. Kjør `python3 build.py`
3. Commit på `dev`, test, åpne pull request mot `main`
4. Merge → GitHub Pages publiserer

Søkemotoren (normalisering, termgrupper, sokeregler, prioriteringsregler) ligger i `index.html`.
