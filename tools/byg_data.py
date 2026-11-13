#!/usr/bin/env python3
"""Bygger data.json til stjerneberegning.dk.

Input er tekst-udtrækkene af Excel-filerne, som Microsoft 365-forbindelsen
returnerer (read_resource): faneblade som "## Sheet: <navn>" og rækker
adskilt med tabulator.

Brug:
  python3 -I tools/byg_data.py TILBUD.txt UDBUD.txt DATO > data.json
    TILBUD.txt = Tilbudssager.xlsx (fanerne "Opgaveliste" og "Kommende udbud")
    UDBUD.txt  = Udbudsjaeger.xlsx (fanen "Alle udbud")
    DATO       = ÅÅÅÅ-MM-DD

Advarsler skrives på stderr. Exit-kode 2 ved fejl, der skal stoppe push.
"""
import json
import re
import sys

FORBUDTE_NAVNE = ["novo", "zealand", "ppv", "højvang"]


def sheets(text):
    out, name, rows = {}, None, []
    for line in text.splitlines():
        m = re.match(r"^## Sheet: (.+?) — ", line)
        if m:
            if name:
                out[name] = rows
            name, rows = m.group(1).strip(), []
            continue
        if name is None:
            continue
        if line.startswith("Formulas:"):
            out[name] = rows
            name, rows = None, []
            continue
        rows.append(line.split("\t"))
    if name:
        out[name] = rows
    return out


def table(rows, first_header):
    """Find overskriftsrækken (første celle = first_header) og returnér dicts."""
    for i, r in enumerate(rows):
        if r and r[0].strip().lower().startswith(first_header.lower()):
            head = [h.strip() for h in r]
            res = []
            for r2 in rows[i + 1:]:
                if not any(c.strip() for c in r2):
                    continue
                r2 = r2 + [""] * (len(head) - len(r2))
                res.append({h: c.strip() for h, c in zip(head, r2)})
            return head, res
    return None, []


def col(head, *names):
    for n in names:
        for h in head:
            if h.lower().startswith(n.lower()):
                return h
    raise KeyError("Kolonne mangler: " + names[0])


def iso(s):
    m = re.match(r"^(\d{1,2})[-./](\d{1,2})[-./](\d{4})", s or "")
    if m:
        return "%s-%02d-%02d" % (m.group(3), int(m.group(2)), int(m.group(1)))
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s or "")
    return m.group(0) if m else ""


def ja(s):
    return (s or "").strip().lower() in ("ja", "yes", "x")


def fag_fra_note(note):
    n = (note or "").lower()
    if "hvac" in n:
        return ["HVAC"]
    fag = ["Ventilation"]
    if "brand" in n:
        fag.append("Brand")
    if "div" in n:
        fag.append("Øvrige fag")
    return fag


def main():
    tilbud_txt = open(sys.argv[1], encoding="utf-8").read()
    udbud_txt = open(sys.argv[2], encoding="utf-8").read()
    dato = sys.argv[3]
    warn, fejl = [], []

    t = sheets(tilbud_txt)
    if "Opgaveliste" not in t:
        sys.exit("FEJL: fanen Opgaveliste blev ikke fundet i Tilbudssager")
    head, rows = table(t["Opgaveliste"], "År-måned")
    if not head:
        sys.exit("FEJL: overskriftsrækken i Opgaveliste blev ikke fundet")
    c_navn, c_start, c_lic = head[0], col(head, "Startdato"), col(head, "Licitation")
    c_note, c_web, c_yd = col(head, "Noter"), col(head, "Webtitel"), col(head, "Ydelse")
    c_vis, c_en = col(head, "Vis på web"), col(head, "Webtitel EN")

    sager = []
    for r in rows:
        if not ja(r[c_vis]):
            continue
        m = re.match(r"^(\d{4}-\d{1,2}-\d{1,2})", r[c_navn])
        start, lic = iso(r[c_start]), iso(r[c_lic])
        if not m or not start or not lic or not r[c_web]:
            warn.append("Sprunget over (mangler nr/dato/webtitel): " + r[c_navn])
            continue
        sager.append({
            "nr": m.group(1),
            "titel": r[c_web],
            "titel_en": r[c_en] or r[c_web],
            "type": r[c_yd] or "Tilbudsberegning",
            "fag": fag_fra_note(r[c_note]),
            "start": start,
            "lic": lic,
        })

    kommende = []
    # 1) Fanen "Kommende udbud" i Tilbudssager.xlsx
    if "Kommende udbud" in t:
        kh, kr = table(t["Kommende udbud"], "Udbud")
        if kh:
            for r in kr:
                if not ja(r.get(col(kh, "Godkendt"), "")):
                    continue
                frist = iso(r.get(col(kh, "Tilbudsfrist"), ""))
                if not frist:
                    warn.append("Kommende udbud uden frist: " + r[kh[0]])
                    continue
                fag = [f.strip() for f in re.split(r"[,/+]", r.get(col(kh, "Fag"), "")) if f.strip()]
                kommende.append({"titel": r[kh[0]], "omraade": r.get(col(kh, "Område"), ""),
                                 "entreprise": r.get(col(kh, "Entrepriseform"), ""),
                                 "fag": fag or ["Ventilation"], "frist": frist})
    # 2) Udbudsjaeger.xlsx, fanen "Alle udbud", Regner på = Ja
    u = sheets(udbud_txt)
    if "Alle udbud" not in u:
        sys.exit("FEJL: fanen Alle udbud blev ikke fundet i Udbudsjaeger")
    uh, ur = table(u["Alle udbud"], "ID")
    if not uh:
        sys.exit("FEJL: overskriftsrækken i Alle udbud blev ikke fundet")
    c_reg = col(uh, "Regner på")
    for r in ur:
        if not ja(r[c_reg]):
            continue
        titel = r[col(uh, "Webtitel (")] or r[col(uh, "Titel")]
        if not r[col(uh, "Webtitel (")]:
            warn.append("Udbud uden webtitel – bruger Titel: " + titel)
        tf, pq = iso(r[col(uh, "Tilbudsfrist")]), iso(r[col(uh, "PQ-frist")])
        frist = iso(r[col(uh, "Næste frist")]) or min([d for d in (tf, pq) if d] or [""])
        if not frist:
            warn.append("Udbud uden frist – ikke vist: " + titel)
            continue
        ent = r[col(uh, "Udbudsform")]
        k = {"titel": titel, "omraade": r[col(uh, "Landsdel")],
             "entreprise": "" if ent.lower() == "ikke oplyst" else ent,
             "fag": ["Ventilation"], "frist": frist, "pq": bool(pq) and frist == pq}
        en = r.get(col(uh, "Webtitel EN"), "")
        if en:
            k["titel_en"] = en
        kommende.append(k)

    data = {"opdateret": dato, "sager": sager, "kommende": kommende}
    blob = json.dumps(data, ensure_ascii=False).lower()
    for navn in FORBUDTE_NAVNE:
        if navn in blob:
            hits = [x["titel"] for x in sager + kommende if navn in json.dumps(x, ensure_ascii=False).lower()]
            fejl.append("Internt navn '%s' i titel: %s" % (navn, "; ".join(hits)))
    for w in warn:
        print("ADVARSEL: " + w, file=sys.stderr)
    for f in fejl:
        print("FEJL: " + f, file=sys.stderr)
    print("OPTALT: sager %d | kommende %d" % (len(sager), len(kommende)), file=sys.stderr)
    sys.stdout.write(json.dumps(data, ensure_ascii=False, indent=1))
    sys.exit(2 if fejl else 0)


if __name__ == "__main__":
    main()
