# open-accountability-ledger

The data and page for the **Open Accountability Ledger**, one public, permanent, machine-readable register of what government was told, what it promised, and what it did. Prepared for the office of Bevan Eatts MLA.

**Status: build-phase, not for publication.** The page is meant to sit behind a password on bevaneatts.com.au until go-live. Five entries are on HOLD (unverified or live) and render greyed and flagged, on purpose, until confirmed to primary source.

## How it fits together

The site reads this repo the same way the rest of Bevan's site reads `in-parliament-data`: the page fetches `data/manifest.json` first (with `cache:'no-store'`), reads its `version`, then fetches the register as `open_accountability_ledger.json?v=<version>`. **`manifest.version` is the cache-buster, bump it by one on every data change** or browsers keep serving the old file.

```
data/
  manifest.json                        version + file list (the cache-buster)
  open_accountability_ledger.json      the register, the single source of truth
  open_accountability_ledger.schema.json   the JSON Schema contract
page/
  open-accountability-ledger.html      the page engine, paste this into Squarespace
views/
  open_accountability_ledger.csv       flat mirror (generated, do not hand-edit)
  open_accountability_ledger.xlsx      editing workbook (generated)
tools/
  build_register.py                    (re)builds the JSON, applies house style
  generate_views.py                    rebuilds csv + xlsx from the JSON
  validate.py                          the gate, must end RESULT: all checks passed
```

The page raw URL it reads from:
`https://raw.githubusercontent.com/bevaneattsmla/open-accountability-ledger/main/data/`

## The page

`page/open-accountability-ledger.html` is a single self-contained file (no build step, no dependencies). Paste it into a Squarespace code block on the password-gated page. It renders the four lens tabs, the self-submission panel, the stat tiles and every entry, greying anything on HOLD. To preview locally, run a static server from the repo root and open `page/open-accountability-ledger.html?data=local` (the `?data=local` flag points it at the local `data/` folder instead of the GitHub raw URL).

## Editing the register

1. Edit `data/open_accountability_ledger.json` (or the authored builder `tools/build_register.py`, then run it).
2. Run `python3 tools/generate_views.py` to refresh the csv + xlsx.
3. Run `python3 tools/validate.py`, it must pass (schema, HOLD rule, counts, no banned punctuation).
4. **Bump `data/manifest.json` `version` by one.**
5. Commit and push. The live page picks up the change on next load.

## House rule

No em-dashes, arrows or middots anywhere in the data (the site's anti-AI-tell rule). `tools/validate.py` enforces it. Use commas or "to".

## The discipline this encodes

Transparency pointed upward at power, not down at the public. Holds recommendations, not people, no personal data, no algorithm, no automated decision, the opposite of Robodebt. Every figure carries a source and a confidence flag, and figures are never summed. Entries with status `UNVERIFIED` or confidence `LOW` are withheld from public use automatically.
