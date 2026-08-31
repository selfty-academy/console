#!/usr/bin/env python3
"""Console Selfty Academy : xlsx « Candidature Webi » -> console.html.

Usage : python3 build.py [chemin/vers/candidature-webi.xlsx]
Le xlsx = export du Sheet « Candidature Webi »
(1mKA765MImL3103Foil5kYu1IR14Ea55nOIv4n7UbHCc), onglets Inscriptions,
Visites, « Mail a contacter webi  ».
"""
import sys, re, json, base64, datetime, io
from collections import Counter, OrderedDict
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
XLSX = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "candidature-webi.xlsx"
TEST_EMAILS = {"alexyoucompte99@gmail.com", "alexandre.majorel@tsm-education.fr",
               "anaisbrault86@gmail.com"}  # Anaïs teste son propre funnel

def norm_phone(v):
    if v is None:
        return ""
    if isinstance(v, float):
        v = int(v)
    s = str(v).strip()
    if not s or s.startswith("#"):
        return ""
    had_plus = s.startswith("+")
    d = re.sub(r"\D", "", s)
    if not d:
        return ""
    if had_plus:
        return d
    if d.startswith("00"):
        return d[2:]
    if d.startswith("0") and len(d) == 10:
        return "33" + d[1:]
    if len(d) == 9 and d[0] in "67":
        return "33" + d
    return d

def fmt_date(dt):
    return dt.strftime("%d/%m %H:%M") if isinstance(dt, datetime.datetime) else ""

def src_label(s):
    s = (s or "").lower()
    if "utm_source=email" in s or "utm_medium=email" in s:
        return "E-mail"
    if "link_in_bio" in s:
        return "Bio Insta"
    if "story" in s:
        return "Story Insta"
    if "utm_source=ig" in s or "instagram" in s:
        return "Insta autre"
    return "Direct / autre"

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb["Inscriptions"]
hdr = [c.value for c in ws[1]]
rows = [dict(zip(hdr, [c.value for c in r])) for r in ws.iter_rows(min_row=2) if any(c.value for c in r)]
real = [r for r in rows
        if (str(r.get("E-mail") or "").strip().lower() not in TEST_EMAILS)
        and not str(r.get("Prénom") or "").lower().startswith("alex test")]

inscrits, cands, la = [], [], []
for r in real:
    tel = norm_phone(r.get("Téléphone"))
    base = {
        "n": str(r.get("Prénom") or "?").strip(),
        "tel": tel,
        "mail": str(r.get("E-mail") or "").strip(),
        "mode": r.get("Mode") or "",
        "date": fmt_date(r.get("Date")),
        "ts": r.get("Date").isoformat() if isinstance(r.get("Date"), datetime.datetime) else "",
        "src": src_label(r.get("Source")),
        "statut": str(r.get("Statut ") or "").strip(),
        "etape": r.get("Dernière étape") or "",
    }
    det = []
    for col, lab in [("Situation", "Sa situation"), ("Sujet du coaching", "Sujet du coaching"),
                     ("Déjà essayé", "Déjà essayé"), ("Accord coaching live", "Accord coaching en direct"),
                     ("LA1 · Parcours", "Parcours"), ("LA2 · Résultat idéal", "Résultat idéal"),
                     ("LA3 · Envie d'apprendre", "Envie d'apprendre"), ("LA4 · Différence", "Ce qui ferait la différence")]:
        v = str(r.get(col) or "").strip()
        if v and v.lower() != "non":
            det.append([lab, v])
    base["det"] = det
    inscrits.append(base)
    if r.get("Mode") == "coaching":
        cands.append({**base,
                      "situation": str(r.get("Situation") or "").strip(),
                      "sujet": str(r.get("Sujet du coaching") or "").strip(),
                      "deja": str(r.get("Déjà essayé") or "").strip(),
                      "accord": str(r.get("Accord coaching live") or "").strip()})
    if str(r.get("Liste d'attente") or "").strip().lower() == "oui":
        la.append({**base,
                   "la1": str(r.get("LA1 · Parcours") or "").strip(),
                   "la2": str(r.get("LA2 · Résultat idéal") or "").strip(),
                   "la3": str(r.get("LA3 · Envie d'apprendre") or "").strip(),
                   "la4": str(r.get("LA4 · Différence") or "").strip()})

la_deja = sum(1 for r in real if r.get("Dernière étape") == "liste_attente_deja_inscrite")

# ---- Liste d'attente école (Sheet « École de coaching  (réponses) ») ----
ewb = openpyxl.load_workbook(HERE / "liste-attente.xlsx", data_only=True)
ews = ewb.active
ehdr = [str(c.value or "").strip() for c in ews[1]]
def short_label(h):
    return h.split("\n")[0].strip().rstrip('?" ').strip() or h[:40]
FORM_COLS = [
    ("Es-tu déjà coach", "Parcours"),
    ("Quel serait ton résultat", "Résultat idéal"),
    ("Qu'est ce que tu as le plus envie d'apprendre", "Envie d'apprendre"),
    ("Qu'est ce qui selon toi ferait la différence", "Ce qui ferait la différence"),
]
ecole = []
for r in ews.iter_rows(min_row=2):
    row = dict(zip(ehdr, [c.value for c in r]))
    if not any(row.values()):
        continue
    nom = str(row.get("Nom prénom et age") or "").strip()
    mail = str(row.get("ton e-mail") or "").strip().lower()
    if not nom and not mail:
        continue
    detail = []
    for pref, lab in FORM_COLS:
        for h in ehdr:
            if h.startswith(pref) and row.get(h):
                detail.append([lab, str(row[h]).strip()])
                break
    notes = []
    idx_start = ehdr.index("ton e-mail") + 5 if "ton e-mail" in ehdr else 12
    for h in ehdr[idx_start:]:
        v = row.get(h)
        if v and str(v).strip() and h not in ("Enregistrement appel",):
            notes.append([short_label(h)[:60], str(v).strip()])
    dt = row.get("Horodateur")
    ecole.append({
        "n": nom or mail,
        "mail": mail,
        "tel": norm_phone(row.get("Numéro de téléphone")),
        "date": fmt_date(dt),
        "ts": dt.isoformat() if isinstance(dt, datetime.datetime) else "",
        "statut": str(row.get("Statut") or "").strip(),
        "chaud": str(row.get("Chaud pour closing ?") or "").strip(),
        "qui": str(row.get("Qui prend ?") or "").strip(),
        "comm": str(row.get("Commentaire") or "").strip(),
        "rec": str(row.get("Enregistrement appel") or "").strip(),
        "detail": detail,
        "notes": notes,
        "src": "Formulaire école",
    })
# + candidatures liste d'attente venues de la page du live, absentes du Sheet
ecole_mails = {e["mail"] for e in ecole if e["mail"]}
for c in la:
    if c["mail"].lower() not in ecole_mails:
        ecole.append({
            "n": c["n"], "mail": c["mail"], "tel": c["tel"], "date": c["date"],
            "ts": c["ts"], "statut": "", "chaud": "", "qui": "", "comm": "",
            "rec": "", "src": "Page du live",
            "detail": [x for x in [["Parcours", c["la1"]], ["Résultat idéal", c["la2"]],
                                   ["Envie d'apprendre", c["la3"]], ["Ce qui ferait la différence", c["la4"]]] if x[1]],
            "notes": [],
        })
ecole.sort(key=lambda e: e["ts"], reverse=True)
ecole_uniques = len({e["mail"] or e["n"] for e in ecole if "doublon" not in e["statut"].lower()})
# recoupement : inscrites au webi (n'importe quel mode)
webi_mails = {i["mail"].lower() for i in inscrits if i["mail"]}
for e in ecole:
    e["webi"] = bool(e["mail"] and e["mail"] in webi_mails)
# segments des inscrits webi : candidat coaching / intéressé école
eco_mails = {e["mail"] for e in ecole if e["mail"]}
ecole_by_mail_all = {e["mail"]: e for e in reversed(ecole) if e["mail"]}
for i in inscrits:
    i["coach"] = i["mode"] == "coaching"
    i["eco"] = bool(i["mail"] and i["mail"].lower() in eco_mails)
    if i["eco"]:
        fiche = ecole_by_mail_all.get(i["mail"].lower())
        if fiche:
            deja = {d[0] for d in i["det"]}
            i["det"] += [d for d in fiche["detail"] if d[0] not in deja]

# ---- Compta : onglets « Paiements » et « Charges » du Sheet École (optionnels) ----
def read_tab(wb_, name):
    for ws_ in wb_.worksheets:
        if ws_.title.strip().lower() == name:
            h = [str(c.value or "").strip() for c in ws_[1]]
            return [dict(zip(h, [c.value for c in r])) for r in ws_.iter_rows(min_row=2) if any(c.value for c in r)]
    return None

def eur(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v)).replace(",", ".")
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0

paiements_rows = read_tab(ewb, "paiements")
charges_rows = read_tab(ewb, "charges")
compta_ok = paiements_rows is not None or charges_rows is not None
paiements, charges = [], []
for r in (paiements_rows or []):
    if not (r.get("Client") or r.get("E-mail")):
        continue
    dt = r.get("Date")
    paiements.append({
        "date": fmt_date(dt) if isinstance(dt, datetime.datetime) else str(dt or "").strip(),
        "ts": dt.isoformat() if isinstance(dt, datetime.datetime) else "",
        "client": str(r.get("Client") or "").strip(),
        "mail": str(r.get("E-mail") or "").strip().lower(),
        "montant": eur(r.get("Montant")),
        "total": eur(r.get("Prix total")),
        "note": str(r.get("Note") or "").strip(),
    })
for r in (charges_rows or []):
    if not (r.get("Poste") or r.get("Montant")):
        continue
    dt = r.get("Date")
    charges.append({
        "date": fmt_date(dt) if isinstance(dt, datetime.datetime) else str(dt or "").strip(),
        "ts": dt.isoformat() if isinstance(dt, datetime.datetime) else "",
        "poste": str(r.get("Poste") or "?").strip(),
        "montant": eur(r.get("Montant")),
        "note": str(r.get("Note") or "").strip(),
    })
# lien automatique paiements -> suivi des appels (fiche école par e-mail)
ecole_by_mail = {e["mail"]: e for e in ecole if e["mail"]}
clients = OrderedDict()
for p in sorted(paiements, key=lambda x: x["ts"]):
    k = p["mail"] or p["client"].lower()
    c = clients.setdefault(k, {"nom": p["client"] or p["mail"], "mail": p["mail"],
                              "paiements": [], "recu": 0.0, "total": 0.0})
    c["paiements"].append({"date": p["date"], "montant": p["montant"], "note": p["note"]})
    c["recu"] += p["montant"]
    c["total"] = max(c["total"], p["total"])
    suivi = ecole_by_mail.get(p["mail"])
    if suivi:
        c["appel"] = {"qui": suivi["qui"], "statut": suivi["statut"], "chaud": suivi["chaud"]}

# Visites
vws = wb["Visites"]
visites = [[c.value for c in r] for r in vws.iter_rows(min_row=2) if any(c.value for c in r)]
v_mobile = sum(1 for v in visites if str(v[1]).strip().lower() == "mobile")

# Campagne mail
mws = wb["Mail a contacter webi "]
mhdr = [str(c.value or "").strip() for c in mws[1]]
i_statut = mhdr.index("Statut envoi") if "Statut envoi" in mhdr else None
mails_total, mails_envoyes = 0, 0
for r in mws.iter_rows(min_row=2):
    vals = [c.value for c in r]
    if not any(vals):
        continue
    mails_total += 1
    if i_statut is not None and vals[i_statut] and "envoyé" in str(vals[i_statut]).lower():
        mails_envoyes += 1

# Inscriptions par jour
par_jour = OrderedDict()
for i in sorted(inscrits, key=lambda x: x["ts"]):
    if i["ts"]:
        d = i["ts"][5:10]
        key = d[3:5] + "/" + d[0:2]
        par_jour[key] = par_jour.get(key, 0) + 1

sources = Counter(i["src"] for i in inscrits)

data = {
    "maj": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
    "webi": {
        "label": "Live du lundi 31 août, 18h",
        "meet": "https://meet.google.com/oxf-vzjg-bhr",
        "groupe": "https://chat.whatsapp.com/JRCRXWKVg8qBUS0uUUTkCA?mode=gi_t",
        "lp": "https://selfty-academy.github.io/live-31-aout/",
    },
    "inscrits": sorted(inscrits, key=lambda x: x["ts"], reverse=True),
    "cands": sorted(cands, key=lambda c: (c["accord"] != "oui", c["ts"])),
    "la": la,
    "laDeja": la_deja,
    "ecole": ecole,
    "ecoleUniques": ecole_uniques,
    "compta": {
        "ok": compta_ok,
        "sheetUrl": "https://docs.google.com/spreadsheets/d/1CUiT962_dGEAWhydaboYmC23ir8gA-CtZyUXB4gErIc/edit",
        "clients": list(clients.values()),
        "paiements": paiements,
        "charges": sorted(charges, key=lambda x: x["ts"], reverse=True),
    },
    "stats": {
        "visites": len(visites),
        "vMobile": v_mobile,
        "inscrits": len(inscrits),
        "coaching": len(cands),
        "mailsTotal": mails_total,
        "mailsEnvoyes": mails_envoyes,
        "parJour": list(par_jour.items()),
        "sources": dict(sources.most_common()),
    },
}

# Logo en data URI (réduit)
try:
    from PIL import Image
    im = Image.open(HERE / "logo-selfty-encre.png")
    im.thumbnail((360, 360))
    buf = io.BytesIO()
    im.save(buf, "PNG", optimize=True)
    logo = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
except Exception:
    logo = "data:image/png;base64," + base64.b64encode((HERE / "logo-selfty-encre.png").read_bytes()).decode()

tpl = (HERE / "template.html").read_text()
out = tpl.replace("__DATA__", json.dumps(data, ensure_ascii=False)).replace("__LOGO__", logo)
(HERE / "console.html").write_text(out)
print(f"console.html : {len(inscrits)} inscrits, {len(cands)} candidatures live, {len(ecole)} lignes école ({ecole_uniques} personnes), {len(visites)} visites")
