#!/usr/bin/env python3
"""Console Selfty Academy : xlsx « Candidature Webi » -> console.html.

Usage : python3 build.py [chemin/vers/candidature-webi.xlsx]
Le xlsx = export du Sheet « Candidature Webi »
(1mKA765MImL3103Foil5kYu1IR14Ea55nOIv4n7UbHCc), onglets Inscriptions,
Visites, « Mail a contacter webi  ».
"""
import sys, re, json, base64, datetime, io, os
from collections import Counter, OrderedDict
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
XLSX = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "candidature-webi.xlsx"
TEST_EMAILS = {"alexyoucompte99@gmail.com", "alexandre.majorel@tsm-education.fr",
               "anaisbrault86@gmail.com"}  # Anaïs teste son propre funnel
# SHOW_TEST=1 (build local uniquement, jamais en CI) : garde les e-mails de test dans Clientes/Contrats
# et ajoute les faux calls de test-calls.json (gitignoré) pour tester la console de bout en bout
SHOW_TEST = os.environ.get("SHOW_TEST") == "1"

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
    if "source=ads" in s or "utm_source=fb" in s or "utm_source=meta" in s:
        return "Ads Meta"
    if "utm_source=email" in s or "utm_medium=email" in s or "source=mail" in s:
        return "E-mail"
    if "link_in_bio" in s:
        return "Bio Insta"
    if "story" in s:
        return "Story Insta"
    if "utm_source=ig" in s or "instagram" in s:
        return "Insta autre"
    return "Direct / autre"

def load_wb(path):
    """Charge un export xlsx Google en remettant les entiers en int : depuis mi-septembre 2026 l'export
    écrit « 2509944.0 » au lieu de « 2509944 », ce qui cassait la correspondance Call ID iClosed <-> Suivi Calls."""
    wb_ = openpyxl.load_workbook(path, data_only=True)
    for ws_ in wb_.worksheets:
        for row_ in ws_.iter_rows():
            for c_ in row_:
                if isinstance(c_.value, float) and c_.value.is_integer():
                    c_.value = int(c_.value)
    return wb_

wb = load_wb(XLSX)
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
        "live": "21sept" if "21-sept" in str(r.get("Source") or "") else "9sept" if "9-sept" in str(r.get("Source") or "") else "31aout",
        # déjà inscrite au live 1 / 2 : trace écrite par le webhook dans la Source lors de la réinscription
        # (« · déjà inscrite au live 1 », « au live 2 », « au live 1 et 2 »)
        "deja1": "déjà inscrite au live 1" in str(r.get("Source") or ""),
        "deja2": bool(re.search(r"déjà inscrite au live (2|1 et 2)", str(r.get("Source") or ""))),
        # LP du live 3 : « lp-webi-21-sept-v1 · source=ads utm_content=l3-img1 »
        "srcp": (re.search(r"(?:^|\s)source=(\S+)", str(r.get("Source") or "")) or [None, ""])[1],
        "ad": (re.search(r"utm_content=(\S+)", str(r.get("Source") or "")) or [None, ""])[1],
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
ewb = load_wb(HERE / "liste-attente.xlsx")
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

def montants(prix, np, ac=0):
    """Montants par défaut des versements, comme dans le contrat (script Contrats, même règle) :
    parts égales, l'arrondi sur le DERNIER (5 000 € en 3x = 1 666 + 1 666 + 1 668).
    Avec un acompte (Amandine : 1 600 € d'abord) : le complément du 1er versement s'ajoute au 2e (1 600 + 1 732 + 1 668)."""
    np = max(1, int(np or 1)); prix = int(prix or 0); ac = int(ac or 0)
    base = prix // np
    t = [base] * (np - 1) + [prix - base * (np - 1)]
    if not ac or ac >= prix or np < 2:
        return t
    return [ac, t[1] + t[0] - ac] + t[2:]

def montants_fixes(raw, prix, np):
    """Colonne « Montants » d'un contrat (« 1600|1732|1668 », fixée à l'envoi, modifiable dans la console) ; None si absente ou incohérente."""
    try:
        m = [float(x.strip().replace(" ", "").replace(",", ".")) for x in str(raw or "").split("|") if x.strip()]
    except ValueError:
        return None
    if len(m) != max(1, int(np or 1)) or any(x <= 0 for x in m) or abs(sum(m) - float(prix or 0)) > 0.5:
        return None
    return [int(x) if x.is_integer() else x for x in m]

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
    # lignes « Impayé versement k/n » (0 €) : marque « pas reçu » posée depuis la console, pas un paiement
    if str(r.get("Note") or "").strip().startswith("Impay"):
        paiements.append({
            "date": fmt_date(dt) if isinstance(dt, datetime.datetime) else str(dt or "").strip(),
            "ts": dt.isoformat() if isinstance(dt, datetime.datetime) else "",
            "client": str(r.get("Client") or "").strip(),
            "mail": str(r.get("E-mail") or "").strip().lower(),
            "montant": 0.0, "total": 0.0,
            "note": str(r.get("Note") or "").strip(),
        })
        continue
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
# Charges mensuelles : une ligne dont la note contient « [mensuel] » est reportée
# automatiquement chaque mois (même jour) depuis sa date jusqu'à aujourd'hui.
def _mois_suivant(d):
    y, m = (d.year + 1, 1) if d.month == 12 else (d.year, d.month + 1)
    return d.replace(year=y, month=m)
_rec = []
for c in charges:
    if "[mensuel]" not in c["note"].lower() or not c["ts"]:
        continue
    d = _mois_suivant(datetime.datetime.fromisoformat(c["ts"]))
    while d <= datetime.datetime.now():
        _rec.append(dict(c, date=fmt_date(d), ts=d.isoformat(), note=c["note"].replace("[mensuel]", "(mensuel, reporté auto)")))
        d = _mois_suivant(d)
charges += _rec
# ---- Suivi calls (onglet « Suivi Calls », écrit par la console via le pont) ----
track_rows = read_tab(ewb, "suivi calls")
track = {}
for r in (track_rows or []):
    cid = re.sub(r"\.0$", "", str(r.get("Call ID") or "").strip())
    if not cid or (cid.upper().startswith("TEST") and not SHOW_TEST):
        continue
    track[cid] = {
        "s": str(r.get("Show-up") or "").strip(),
        "r": str(r.get("Résultat") or "").strip(),
        "prix": eur(r.get("Prix")),
        "np": int(eur(r.get("Nb paiements")) or 0),
        "ac": int(eur(r.get("Acompte")) or 0),
        "q": int(eur(r.get("Qualif /10")) or 0),
        "retrans": str(r.get("Retranscription") or "").strip(),
        "comment": str(r.get("Commentaire") or "").strip(),
        "prep": str(r.get("Prépa Alex") or "").strip(),
        "prepMaj": str(r.get("Prépa MAJ") or "").strip(),
        "dc": str(r.get("Date close") or "").strip(),   # date réelle du closing saisie dans la console
    }

def dstr_(v):
    return v.strftime("%d/%m/%Y %H:%M") if isinstance(v, datetime.datetime) else str(v or "").strip()

# ---- Clientes signées (onglet « Clients », créé par le pont à la 1re vente) ----
clients_rows = read_tab(ewb, "clients")
clientes = []
for r in (clients_rows or []):
    mail = str(r.get("E-mail") or "").strip().lower()
    nom = str(r.get("Nom") or "").strip()
    if (not mail and not nom) or (mail in TEST_EMAILS and not SHOW_TEST):
        continue
    dt = r.get("Date signature")
    clientes.append({
        "n": nom or mail,
        "mail": mail,
        "tel": norm_phone(r.get("Téléphone")),
        "date": dt.strftime("%d/%m/%Y") if isinstance(dt, datetime.datetime) else str(dt or "").strip(),
        "ts": dt.isoformat() if isinstance(dt, datetime.datetime) else "",
        "offre": str(r.get("Offre") or "").strip(),
        "prix": eur(r.get("Prix")),
        "statut": str(r.get("Statut") or "").strip(),
        "notes": str(r.get("Notes") or "").strip(),
        # mail de bienvenue envoyé depuis la console (date) + PDF « Ta vision » (Drive)
        "bienvenue": dstr_(r.get("Mail bienvenue")),
        "vision": str(r.get("Vision PDF") or "").strip(),
    })
clientes.sort(key=lambda c: c["ts"], reverse=True)

# ---- Présences aux calls de groupe (onglet « Présences », écrit par la console via presence_set) ----
pres_rows = read_tab(ewb, "présences")
presences = []
for r in (pres_rows or []):
    mail = str(r.get("E-mail") or "").strip().lower()
    d = r.get("Date session")
    date = d.strftime("%d/%m/%Y") if isinstance(d, datetime.datetime) else str(d or "").strip()
    if not mail or not date or (mail in TEST_EMAILS and not SHOW_TEST):
        continue
    presences.append({"date": date, "mail": mail, "nom": str(r.get("Nom") or "").strip(),
                      "present": str(r.get("Présent") or "").strip().lower() == "oui",
                      "src": str(r.get("Source") or "").strip()})

# ---- Contrats envoyés / signés (onglet « Contrats », écrit par le pont) ----
# dossiers Drive (compte selfty.academy) où le script Contrats range les PDF signés et les factures
DRIVE_CONTRATS = "https://drive.google.com/drive/folders/1x9Evn3ZP_j5JQ9MT9sNAoqXHnfC7lB_j"
DRIVE_FACTURES = "https://drive.google.com/drive/folders/1M7YKwkC2FV1czGkMZGOTAnUy985uSQ2D"
def dstr(v):
    return v.strftime("%d/%m/%Y %H:%M") if isinstance(v, datetime.datetime) else str(v or "").strip()

contrats_rows = read_tab(ewb, "contrats")
contrats = {}
for r in (contrats_rows or []):
    mail = str(r.get("E-mail") or "").strip().lower()
    if not mail or (mail in TEST_EMAILS and not SHOW_TEST):
        continue
    st = str(r.get("Statut") or "").strip()
    rec = {
        "token": str(r.get("Token") or "").strip(),
        "mail": mail,
        "prenom": str(r.get("Prénom") or "").strip(),
        "nom": str(r.get("Nom") or "").strip(),
        "prix": eur(r.get("Prix")),
        "np": int(eur(r.get("Nb paiements")) or 1),
        "ac": int(eur(r.get("Acompte")) or 0),
        "statut": st,
        "sent": dstr(r.get("Date envoi")),
        "signedAt": dstr(r.get("Date signature")),
        "pdf": str(r.get("PDF") or "").strip(),
        "caseA": str(r.get("Case A") or "").strip() == "Oui",
        "caseB": str(r.get("Case B") or "").strip() == "Oui",
        # contrat au nom de la société de la cliente (Type = Société) + échéancier daté + rappels J-2 envoyés
        "type": str(r.get("Type") or "Particulier").strip() or "Particulier",
        "soc": ({"nom": str(r.get("Société") or "").strip(), "forme": str(r.get("Forme") or "").strip(),
                 "siren": str(r.get("SIREN") or "").strip(), "tva": str(r.get("TVA intra") or "").strip(),
                 "siege": str(r.get("Siège") or "").strip(), "fonction": str(r.get("Fonction") or "").strip()}
                if str(r.get("Type") or "").strip() == "Société" else None),
    }
    _m = montants_fixes(r.get("Montants"), rec["prix"], rec["np"]) or montants(rec["prix"], rec["np"], rec["ac"])
    rec["mt"] = _m
    _dates = [d.strip() for d in str(r.get("Échéances") or "").split("|") if d.strip()][:len(_m)]
    _rap = dict(x.split(":", 1) for x in str(r.get("Rappels") or "").split(";") if ":" in x)
    rec["ech"] = [{"k": i + 1, "n": len(_m), "date": d, "montant": _m[i] if i < len(_m) else 0,
                   "paye": False, "rappel": _rap.get(str(i + 1), "")} for i, d in enumerate(_dates)]
    prev = contrats.get(mail)
    # un contrat signé prime ; sinon le plus récent (dernière ligne) l'emporte
    if not prev or st == "Signé" or prev["statut"] != "Signé":
        contrats[mail] = rec

# lien automatique paiements -> suivi des appels (fiche école par e-mail)
ecole_by_mail = {e["mail"]: e for e in ecole if e["mail"]}
clients = OrderedDict()
for p in sorted(paiements, key=lambda x: x["ts"]):
    if p["note"].startswith("Impay"):
        continue
    k = p["mail"] or p["client"].lower()
    c = clients.setdefault(k, {"nom": p["client"] or p["mail"], "mail": p["mail"],
                              "paiements": [], "recu": 0.0, "total": 0.0})
    c["paiements"].append({"date": p["date"], "montant": p["montant"], "note": p["note"]})
    c["recu"] += p["montant"]
    c["total"] = max(c["total"], p["total"])
    suivi = ecole_by_mail.get(p["mail"])
    if suivi:
        c["appel"] = {"qui": suivi["qui"], "statut": suivi["statut"], "chaud": suivi["chaud"]}
# paiements -> fiches clientes (encaissé / contracté)
for c in clientes:
    p = clients.get(c["mail"])
    c["recu"] = p["recu"] if p else 0.0
    c["total"] = max(c["prix"], p["total"] if p else 0.0, c["recu"])

# Visites
vws = wb["Visites"]
visites = [[c.value for c in r] for r in vws.iter_rows(min_row=2) if any(c.value for c in r)]
v_mobile = sum(1 for v in visites if str(v[1]).strip().lower() == "mobile")

# Live 3 (lundi 28/09, reporté du 21 puis du 24 : le tag « 21-sept » et la clé « 21sept » restent) : visites de la LP par source / ad (colonne « Campagne (UTM) » = « lp=21-sept source=ads utm_content=l3-img1 »)
LIVE3_TESTS_AVANT = "2026-09-15T01:05"   # visites de test à la mise en ligne de la LP (captures, vérifications)
live3_vis = Counter()
live3_vis_src = Counter()
for v in visites:
    camp = str(v[3] or "") if len(v) > 3 else ""
    if "lp=21-sept" not in camp:
        continue
    if isinstance(v[0], datetime.datetime) and v[0].strftime("%Y-%m-%dT%H:%M") < LIVE3_TESTS_AVANT:
        continue
    ad = (re.search(r"utm_content=(\S+)", camp) or [None, ""])[1]
    sp = (re.search(r"(?:^|\s)source=(\S+)", camp) or [None, ""])[1]
    live3_vis[ad or "(sans utm_content)"] += 1
    live3_vis_src[sp or "direct"] += 1

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

# ---- iClosed : tous les calls bookés (passés, à venir, annulés) ----
import os
import urllib.request

ic_key = os.environ.get("ICLOSED_KEY", "")
if not ic_key and (HERE / "iclosed-key.txt").exists():
    ic_key = (HERE / "iclosed-key.txt").read_text().strip()
icalls, ic_ok = [], False
if ic_key:
    try:
        raw = []
        for page in range(0, 20):  # pages 0-indexées, 100 par page, garde-fou 2000 calls
            req = urllib.request.Request(
                f"https://public.api.iclosed.io/v1/eventCalls?limit=100&page={page}",
                headers={"Authorization": "Bearer " + ic_key})
            batch = json.load(urllib.request.urlopen(req, timeout=30)).get("data", {}).get("eventCalls", [])
            raw += batch
            if len(batch) < 100:
                break
        seen_ids = set()
        for c in raw:
            if c.get("id") in seen_ids:
                continue
            seen_ids.add(c.get("id"))
            quest = []
            for q in c.get("secondaryAnswers") or []:
                ans = " / ".join(str(a.get("answer") or "") for a in (q.get("answer") or []) if a.get("answer"))
                if ans:
                    quest.append([str(q.get("statement") or "?").strip(), ans])
            task = (c.get("task") or [{}])[0]
            icalls.append({
                "id": c.get("id"),
                "cid": c.get("contactId") or "",
                "n": str(c.get("inviteeName") or "?").strip(),
                "mail": str(c.get("inviteeEmail") or "").strip().lower(),
                "tel": norm_phone(c.get("phoneNumber")),
                "utc": c.get("dateTimeUTC") or "",
                "link": c.get("locationLinkInvitee") or "",
                "event": str((c.get("event") or {}).get("name") or "").strip(),
                "closer": str((c.get("user") or {}).get("firstName") or "").strip(),
                "cancel": bool(c.get("cancelReason")) or c.get("eventType") == "CANCELLED",
                "cancelWhy": str(c.get("cancelReason") or "").strip(),
                "outcome": str(task.get("outcome") or "").strip(),
                "notes": str(task.get("notes") or c.get("notes") or "").strip(),
                "quest": quest,
            })
        ic_ok = True
        print(f"iClosed : {len(icalls)} calls")
    except Exception as ex:
        print("iClosed fetch KO (on garde la console sans) :", ex)
# faux calls : env TEST_CALLS (secret GitHub, injecté AUSSI en CI : call de démo pour qu'Anaïs teste
# le remplissage / contrat / facture sur la vraie console) ou fichier local test-calls.json (SHOW_TEST=1 seulement).
# Un faux call se retire comme un vrai : pill « Call test 🧪 » -> exclu au build suivant.
raw_tc = os.environ.get("TEST_CALLS", "")
if not raw_tc and SHOW_TEST and (HERE / "test-calls.json").exists():
    raw_tc = (HERE / "test-calls.json").read_text()
if raw_tc.strip():
    fake = json.loads(raw_tc)
    icalls += fake
    print(f"faux calls ajoutés : {len(fake)}")
# ---- Leads iClosed sans call : contacts iClosed (formulaire commencé) sans aucun call booké ----
# suivi (statut de relance + notes) dans l'onglet « Leads iClosed » du Sheet École, écrit par la console (lead_update)
leads, leads_ok, ic_contacts = [], False, []
if ic_key:
    try:
        req = urllib.request.Request("https://public.api.iclosed.io/v1/contacts?limit=100&page=0",
                                     headers={"Authorization": "Bearer " + ic_key})
        with urllib.request.urlopen(req, timeout=30) as resp:
            cdata = json.loads(resp.read().decode("utf-8")).get("data", {})
        contacts = cdata.get("contacts") or []
        ic_contacts = contacts
        booked_ids = {str(c.get("cid") or "") for c in icalls}
        booked_mails = {c["mail"] for c in icalls if c.get("mail")}
        booked_tels = {c["tel"] for c in icalls if c.get("tel")}
        lead_rows = read_tab(ewb, "leads iclosed") or []
        lead_suivi = {str(r.get("Contact ID") or "").strip(): r for r in lead_rows}
        for ct in contacts:
            cid = str(ct.get("id") or "")
            mail = str(ct.get("email") or "").strip().lower()
            if "@" not in mail:
                mail = ""   # iClosed recopie le numéro dans l'e-mail quand le formulaire s'arrête au téléphone
            tel = norm_phone(ct.get("phoneNumber"))
            if tel.startswith("330") and len(tel) == 12:
                tel = "33" + tel[3:]   # « +33 06… » saisi avec le 0
            nom = (str(ct.get("firstName") or "") + " " + str(ct.get("lastName") or "")).strip()
            if cid in booked_ids or (mail and mail in booked_mails) or (tel and tel in booked_tels):
                continue
            if mail in TEST_EMAILS or nom.lower().startswith("alex") and "test" in nom.lower():
                continue
            sv = lead_suivi.get(cid, {})
            leads.append({
                "id": cid, "n": nom or mail or ("+" + tel if tel else "?"), "mail": mail, "tel": tel,
                "cree": str(ct.get("createdAt") or ""),
                "st": str(ct.get("status") or ""),
                "statut": str(sv.get("Statut") or "").strip(),
                "notes": str(sv.get("Notes") or "").strip(),
                "maj": dstr_(sv.get("MAJ")) if sv else "",
            })
        leads_ok = True
        print(f"Leads iClosed sans call : {len(leads)} / {len(contacts)} contacts")
    except Exception as ex:
        print("iClosed contacts KO (onglet leads vide) :", ex)

# suivi closing du Sheet accroché à chaque call ; « Call test » = exclu de partout
for c in icalls:
    c["trk"] = track.get(str(c["id"]))
# garde-fou : Suivi Calls rempli mais aucun call iClosed ne s'y accroche = format d'ID cassé
# (bug du 16/09 : export « 2509944.0 ») -> alerte Telegram au lieu de tout renvoyer « à remplir » en silence
import urllib.parse
_ids_reels = [k for k in track if k.isdigit()]
if ic_ok and len(_ids_reels) >= 5 and not any(c["trk"] for c in icalls):
    print("⚠ ALERTE : aucun call iClosed ne correspond à l'onglet Suivi Calls (format des Call ID ?)", _ids_reels[:3])
    if os.environ.get("TG_TOKEN") and os.environ.get("TG_CHAT"):
        try:
            urllib.request.urlopen(urllib.request.Request(
                f"https://api.telegram.org/bot{os.environ['TG_TOKEN']}/sendMessage",
                data=urllib.parse.urlencode({"chat_id": os.environ["TG_CHAT"], "text":
                    "⚠ Console Selfty : aucun call iClosed ne correspond au Sheet Suivi Calls, "
                    "tous les calls vont ressortir « à remplir ». Format des Call ID à vérifier."}).encode()), timeout=10)
        except Exception as ex:
            print("alerte Telegram KO :", ex)
icalls =[c for c in icalls if not (c["trk"] and c["trk"]["s"].lower() == "call test")]
if ic_ok:
    # dump minimal pour notify_calls.py (notif Telegram des nouveaux bookings)
    (HERE / "icalls.json").write_text(json.dumps(
        [{"id": c["id"], "n": c["n"], "utc": c["utc"], "event": c["event"], "cancel": c["cancel"]}
         for c in icalls], ensure_ascii=False))

# ---- Scholarship : candidatures Tally (form Np1Gy0, compte perso Alex) ----
try:
    from zoneinfo import ZoneInfo
    TZ_PARIS = ZoneInfo("Europe/Paris")
except Exception:
    TZ_PARIS = None

def iso_paris(at):
    try:
        d = datetime.datetime.fromisoformat(at.replace("Z", "+00:00"))
        if TZ_PARIS:
            d = d.astimezone(TZ_PARIS)
        return d.strftime("%d/%m %H:%M")
    except Exception:
        return ""

ty_key = os.environ.get("TALLY_API_KEY", "")
if not ty_key and (HERE / "tally-key.txt").exists():
    ty_key = (HERE / "tally-key.txt").read_text().strip()
schol_subs, schol_ok, schol_stats = [], False, {}
SCHOL_SKIP = TEST_EMAILS | {"test@test.fr"}
ID_PRENOM, ID_NOM, ID_MAIL, ID_TEL, ID_INSTA, ID_HIDDEN = "yE0EXd", "XMRM5z", "8PJPNr", "0JbJVA", "zr0rEg", "g7Q7bJ"

def ty_txt(a):
    if a is None:
        return ""
    if isinstance(a, list):
        return " · ".join(str(x) for x in a if not isinstance(x, dict))
    return str(a).strip()

if ty_key:
    try:
        qlabels, raw_subs, page = {}, [], 1
        while True:
            req = urllib.request.Request(
                f"https://api.tally.so/forms/Np1Gy0/submissions?filter=all&page={page}",
                headers={"Authorization": "Bearer " + ty_key,
                         "User-Agent": "curl/8.4.0"})  # Cloudflare bloque l'UA Python
            d = json.load(urllib.request.urlopen(req, timeout=30))
            for q in d.get("questions") or []:
                qlabels[q["id"]] = (str(q.get("title") or "?").strip(), str(q.get("type") or ""))
            raw_subs += d.get("submissions") or []
            schol_stats = d.get("totalNumberOfSubmissionsPerFilter") or {}
            if not d.get("hasMore"):
                break
            page += 1
        for s in raw_subs:
            a = {r.get("questionId"): r.get("answer") for r in s.get("responses") or []}
            mail = ty_txt(a.get(ID_MAIL)).lower()
            if mail in SCHOL_SKIP:
                continue
            # personne n'a rien rempli d'identifiable : du bruit, on saute
            if not mail and not ty_txt(a.get(ID_PRENOM)) and not ty_txt(a.get(ID_TEL)):
                continue
            det, files = [], []
            for r in s.get("responses") or []:
                qid = r.get("questionId")
                if qid in (ID_PRENOM, ID_NOM, ID_MAIL, ID_TEL, ID_INSTA, ID_HIDDEN):
                    continue
                lab, qtype = qlabels.get(qid, ("?", ""))
                ans = r.get("answer")
                if qtype == "FILE_UPLOAD" and isinstance(ans, list):
                    files += [{"n": str(f.get("name") or "fichier"), "u": str(f.get("url") or "")}
                              for f in ans if isinstance(f, dict)]
                    continue
                v = ty_txt(ans)
                if v:
                    det.append([lab, v])
            hid = a.get(ID_HIDDEN) if isinstance(a.get(ID_HIDDEN), dict) else {}
            at = str(s.get("submittedAt") or "")
            schol_subs.append({
                "id": s.get("id"),
                "n": (ty_txt(a.get(ID_PRENOM)) + " " + ty_txt(a.get(ID_NOM))).strip() or mail or "?",
                "mail": mail,
                "tel": norm_phone(ty_txt(a.get(ID_TEL))),
                "insta": ty_txt(a.get(ID_INSTA)).lstrip("@"),
                "at": at,
                "date": iso_paris(at),
                "done": bool(s.get("isCompleted")),
                "src": str(hid.get("source") or "").strip(),
                "det": det,
                "files": files,
            })
        schol_subs.sort(key=lambda x: x["at"], reverse=True)
        schol_ok = True
        print(f"Scholarship Tally : {len(schol_subs)} candidature(s)")
    except Exception as ex:
        print("Tally scholarship KO (on garde la console sans) :", ex)

# ---- Candidatures coaching individuel d'Anaïs (form Tally WOLYdJ, même clé ; script tally-coaching-anais/build_form.py) ----
COACH_FORM = "WOLYdJ"
# lu par LIBELLÉ de question (changer ici si un titre change dans build_form.py)
CQ_PRENOM, CQ_NOM, CQ_TEL, CQ_INSTA = "Ton prénom", "Ton nom", "Ton téléphone", "Ton compte Instagram"
CQ_IDEAL, CQ_FREIN, CQ_REVENU = "Quelle est ta situation idéale", "Te connaissant", "Tes revenus"
coach_subs, coach_ok, coach_stats = [], False, {}
if ty_key:
    try:
        qlabels, raw_subs, page = {}, [], 1
        while True:
            req = urllib.request.Request(
                f"https://api.tally.so/forms/{COACH_FORM}/submissions?filter=all&page={page}",
                headers={"Authorization": "Bearer " + ty_key, "User-Agent": "curl/8.4.0"})
            d = json.load(urllib.request.urlopen(req, timeout=30))
            for q in d.get("questions") or []:
                qlabels[q["id"]] = (str(q.get("title") or "").strip(), str(q.get("type") or ""))
            raw_subs += d.get("submissions") or []
            coach_stats = d.get("totalNumberOfSubmissionsPerFilter") or {}
            if not d.get("hasMore"):
                break
            page += 1
        for s in raw_subs:
            v, hid = {}, {}
            for r in s.get("responses") or []:
                lab, qtype = qlabels.get(r.get("questionId"), ("", ""))
                ans = r.get("answer")
                if qtype == "HIDDEN_FIELDS" and isinstance(ans, dict):
                    hid = ans
                    continue
                for pre in (CQ_PRENOM, CQ_NOM, CQ_TEL, CQ_INSTA, CQ_IDEAL, CQ_FREIN, CQ_REVENU):
                    if lab.startswith(pre):
                        v[pre] = ty_txt(ans)
            src = str(hid.get("source") or "").strip()
            if src == "test" and not SHOW_TEST:
                continue
            if not v.get(CQ_PRENOM) and not v.get(CQ_TEL):
                continue
            at = str(s.get("submittedAt") or s.get("createdAt") or "")
            rev = v.get(CQ_REVENU, "")
            coach_subs.append({
                "id": s.get("id"),
                "n": (v.get(CQ_PRENOM, "") + " " + v.get(CQ_NOM, "")).strip() or "?",
                "prenom": v.get(CQ_PRENOM, ""),
                "tel": norm_phone(v.get(CQ_TEL, "")),
                "insta": v.get(CQ_INSTA, "").strip().lstrip("@").split("instagram.com/")[-1].strip("/ "),
                "ideal": v.get(CQ_IDEAL, ""),
                "frein": v.get(CQ_FREIN, ""),
                "revenu": rev,
                "tranche": "100" if rev.startswith("100") else "10" if "10 et 100" in rev else "0" if rev else "",
                "at": at,
                "date": iso_paris(at),
                "done": bool(s.get("isCompleted")),
                "src": src,
            })
        coach_subs.sort(key=lambda x: x["at"], reverse=True)
        coach_ok = True
        print(f"Coaching Tally : {len(coach_subs)} candidature(s)")
    except Exception as ex:
        print("Tally coaching KO (on garde la console sans) :", ex)

# ---- Sondage « Le prochain live, c'est toi qui choisis » (form Tally VLKopa, même clé ; script selfty-marketing-sept/sondage_form.py) ----
SOND_FORM = "VLKopa"
# lu par LIBELLÉ de question (changer ici si un titre change dans sondage_form.py)
SQ_MASTER, SQ_MASTER_SUJET, SQ_MASTER_LIBRE = "Une masterclass d", "Si oui, sur quel sujet", "Un autre sujet"
SQ_SOMA, SQ_SUJET, SQ_PRENOM = "Une expérience somatique", "Le sujet que tu veux", "Ton prénom"
# sujets de masterclass proposés -> libellé court pour le graphique (préfixe du choix)
SOND_SUJETS = [("Ta première cliente", "Première cliente payante"), ("Annoncer ton prix", "Annoncer son prix, rapport à l'argent"),
               ("La structure d", "Structure d'une séance (5 niveaux)"), ("Te montrer", "Visibilité, se montrer"), ("Un autre sujet", "Autre sujet")]
SOND_SOURCES = ["mail", "whatsapp", "story", "ecole"]
sond_subs, sond_ok, sond_stats = [], False, {}


def sond_court(lab):
    for pre, court in SOND_SUJETS:
        if lab.startswith(pre):
            return court
    return lab[:40] if lab else ""


def sond_oui(v):
    v = v.strip().lower()
    return None if not v else v.startswith("oui")


if ty_key:
    try:
        qlabels, raw_subs, page = {}, [], 1
        while True:
            req = urllib.request.Request(
                f"https://api.tally.so/forms/{SOND_FORM}/submissions?filter=all&page={page}",
                headers={"Authorization": "Bearer " + ty_key, "User-Agent": "curl/8.4.0"})
            d = json.load(urllib.request.urlopen(req, timeout=30))
            for q in d.get("questions") or []:
                qlabels[q["id"]] = (str(q.get("title") or "").strip(), str(q.get("type") or ""))
            raw_subs += d.get("submissions") or []
            sond_stats = d.get("totalNumberOfSubmissionsPerFilter") or {}
            if not d.get("hasMore"):
                break
            page += 1
        # numéro / nom retrouvés par e-mail dans les inscrites, la liste d'attente, les calls et les contacts iClosed
        sond_who = {}
        for lst in (leads, icalls, ecole, inscrits):
            for x in lst:
                m = str(x.get("mail") or "").strip().lower()
                if m and (m not in sond_who or (x.get("tel") and not sond_who[m]["tel"])):
                    sond_who[m] = {"tel": x.get("tel") or "", "n": str(x.get("n") or "").strip()}
        for s in raw_subs:
            v, hid = {}, {}
            for r in s.get("responses") or []:
                lab, qtype = qlabels.get(r.get("questionId"), ("", ""))
                ans = r.get("answer")
                if qtype == "HIDDEN_FIELDS" and isinstance(ans, dict):
                    hid = ans
                    continue
                for pre in (SQ_MASTER, SQ_MASTER_SUJET, SQ_MASTER_LIBRE, SQ_SOMA, SQ_SUJET, SQ_PRENOM):
                    if lab.startswith(pre):
                        v[pre] = ty_txt(ans)
            src = str(hid.get("source") or "").strip().lower()
            mail = str(hid.get("email") or "").strip().lower()
            if (src == "test" or mail in TEST_EMAILS) and not SHOW_TEST:
                continue
            if not v.get(SQ_SUJET) and not v.get(SQ_MASTER) and not v.get(SQ_SOMA):
                continue  # rien de répondu
            who = sond_who.get(mail, {})
            prenom = (v.get(SQ_PRENOM) or str(hid.get("prenom") or "")).strip()
            at = str(s.get("submittedAt") or s.get("createdAt") or "")
            sond_subs.append({
                "id": s.get("id"),
                "prenom": prenom or (who.get("n") or "").split(" ")[0],
                "mail": mail,
                "tel": who.get("tel") or "",
                "src": src if src in SOND_SOURCES else ("autre" if src else ""),
                "master": sond_oui(v.get(SQ_MASTER, "")),
                "sujet": v.get(SQ_MASTER_SUJET, ""),
                "sujetCourt": sond_court(v.get(SQ_MASTER_SUJET, "")),
                "sujetLibre": v.get(SQ_MASTER_LIBRE, ""),
                "soma": sond_oui(v.get(SQ_SOMA, "")),
                "texte": v.get(SQ_SUJET, ""),
                "at": at,
                "date": iso_paris(at),
                "done": bool(s.get("isCompleted")),
            })
        sond_subs.sort(key=lambda x: x["at"], reverse=True)
        sond_ok = True
        print(f"Sondage Tally : {len(sond_subs)} réponse(s)")
    except Exception as ex:
        print("Tally sondage KO (on garde la console sans) :", ex)

# ---- Bilans hebdo des clientes (« EOW », form Tally 1AekPO « Mon bilan de la semaine », même clé) ----
EOW_FORM = "1AekPO"
ECOLE_DEBUT = "2026-10-10"

def iso_week(at):
    try:
        d = datetime.datetime.fromisoformat(at.replace("Z", "+00:00"))
        if TZ_PARIS:
            d = d.astimezone(TZ_PARIS)
        y, w, _ = d.isocalendar()
        return f"{y}-W{w:02d}"
    except Exception:
        return ""

# attribution manuelle (onglet « Bilans attribution », écrit par la console) : Submission ID -> e-mail ('__ignore__' = ignoré)
attrib = {}
for r in (read_tab(ewb, "bilans attribution") or []):
    sid = str(r.get("Submission") or "").strip()
    if sid:
        attrib[sid] = str(r.get("E-mail") or "").strip().lower()

eow_subs, eow_ok = [], False
if ty_key:
    try:
        qlabels, raw_subs, page = {}, [], 1
        while True:
            req = urllib.request.Request(
                f"https://api.tally.so/forms/{EOW_FORM}/submissions?filter=completed&page={page}",
                headers={"Authorization": "Bearer " + ty_key, "User-Agent": "curl/8.4.0"})
            d = json.load(urllib.request.urlopen(req, timeout=30))
            for q in d.get("questions") or []:
                qlabels[q["id"]] = (str(q.get("title") or "?").strip(), str(q.get("type") or ""))
            raw_subs += d.get("submissions") or []
            if not d.get("hasMore"):
                break
            page += 1
        for s in raw_subs:
            hid, mail, prenom, det, scores = {}, "", "", [], {}
            for r in s.get("responses") or []:
                lab, qtype = qlabels.get(r.get("questionId"), ("?", ""))
                ans = r.get("answer")
                if qtype == "HIDDEN_FIELDS" and isinstance(ans, dict):
                    hid = ans
                    continue
                if lab == "E-mail":
                    mail = mail or ty_txt(ans).lower()
                    continue
                if lab == "Prénom":
                    prenom = prenom or ty_txt(ans)
                    continue
                v = ty_txt(ans)
                if v == "":
                    continue
                if qtype == "LINEAR_SCALE":
                    try:
                        scores[lab] = int(float(v))
                    except ValueError:
                        pass
                det.append([lab, v])
            mail_decl = mail
            mail = str(hid.get("email") or "").strip().lower() or mail
            sid = str(s.get("id") or "")
            if attrib.get(sid):
                mail = attrib[sid]
            prenom = str(hid.get("prenom") or "").strip() or prenom
            # sans e-mail : gardé quand même, il sortira en « non attribué » dans la console
            if mail in TEST_EMAILS and not SHOW_TEST:
                continue
            at = str(s.get("submittedAt") or "")
            pres = next((v for l, v in det if l.startswith("Cette semaine, aux calls de groupe")), "")
            eow_subs.append({
                "id": sid, "mail": mail, "prenom": prenom, "mailDecl": mail_decl, "attrib": bool(attrib.get(sid)), "at": at, "date": iso_paris(at),
                "week": str(hid.get("semaine") or "").strip() or iso_week(at),
                "feel": scores.get("Comment tu te sens, là, en cette fin de semaine ?"),
                "energie": scores.get("Ton niveau d’énergie sur la semaine"),
                "confiance": scores.get("Ta confiance dans ta posture de coach cette semaine", scores.get("Ta confiance à l’idée de donner ta première séance")),
                "ressentiSeances": scores.get("Comment tu t’es sentie pendant tes séances ?"),
                "seances": next((v for l, v in det if l.startswith("Est-ce que tu donnes déjà des séances")), ""),
                "note": scores.get("Ta note globale de la semaine"),
                "presence": pres,
                # croisement avec les engagements : champs cachés posés par le mail du vendredi / le portail
                "serie": str(hid.get("serie") or "").strip(), "jours": str(hid.get("jours") or "").strip(),
                "hebdoLien": str(hid.get("hebdo") or "").strip(),
                "hebdoDecl": next((v for l, v in det if l.startswith("Ton action hebdo")), ""),
                "det": det,
            })
        eow_subs.sort(key=lambda x: x["at"], reverse=True)
        eow_ok = True
        print(f"Bilans hebdo Tally : {len(eow_subs)} bilan(s)")
    except Exception as ex:
        print("Tally bilans KO (on garde la console sans) :", ex)


# ---- Objectif 8 : liste priorisée des personnes à rappeler pour booker un call avec Anaïs ----
# Sources croisées : liste d'attente école, inscrites aux lives, calls iClosed uniquement annulés,
# contacts iClosed sans call, candidatures coaching (Tally WOLYdJ) et bourse (Tally Np1Gy0).
# Exclues : clientes (onglet Clients + ventes de Suivi Calls + contrats), tests, toute personne avec un call iClosed non annulé
# (passé ou à venir, 15/09 ; sauf si cochée « Contactée » AVANT ce call : elle reste en « Call booké » pour le suivi),
# + les exclusions / messages perso / prénoms du fichier LOCAL objectif8.json (secret GitHub OBJECTIF8 en CI :
# aucune donnée perso ne vit dans le repo public). Coche partagée = onglet Sheet « Rappels objectif 8 » (pont rappel_set).
import unicodedata
ICLOSED_LINK = "https://app.iclosed.io/e/SelftyAcademy/clarity-call"
OBJ8_DEBUT = "2026-09-13"   # lancement de l'objectif : les ventes comptent à partir de ce jour
OBJ8_CIBLE = 8
raw_o8 = os.environ.get("OBJECTIF8", "")
if not raw_o8 and (HERE / "objectif8.json").exists():
    raw_o8 = (HERE / "objectif8.json").read_text()
if raw_o8.strip().startswith("gz:"):
    import gzip
    raw_o8 = gzip.decompress(base64.b64decode(raw_o8.strip()[3:])).decode()
try:
    o8cfg = json.loads(raw_o8) if raw_o8.strip() else {}
except Exception as ex:
    print("objectif8.json illisible :", ex)
    o8cfg = {}
o8_msgs = {str(k).lower(): v for k, v in (o8cfg.get("messages") or {}).items()}
o8_notes = {str(k).lower(): v for k, v in (o8cfg.get("notes") or {}).items()}
o8_prenoms = {str(k).lower(): v for k, v in (o8cfg.get("prenoms") or {}).items()}
o8_alias = {str(k).lower(): str(v).lower() for k, v in (o8cfg.get("alias") or {}).items()}   # e-mail mal tapé -> e-mail de la même personne
O8_SEG = OrderedDict([
    ("cand", "Candidature coaching"), ("chaud", "École chaud / à rappeler"), ("annule", "Call annulé, à rebooker"),
    ("bourse", "Candidature bourse"), ("ecole", "Liste d'attente non traitée"), ("webi2", "Aux 2 lives"),
    ("webi1", "Un seul live"), ("lead", "iClosed sans call"), ("froid", "École déjà contactée"),
])
O8_LVL = {s: i + 1 for i, s in enumerate(O8_SEG)}
# messages de repli par segment (les messages perso de objectif8.json priment) : pas d'emoji (le lien wa.me les casse), pas de tiret cadratin.
# 15/09 : chaque message annonce que l'agenda d'Anaïs est complet cette semaine (O8_DISPO) avant le lien.
O8_DISPO = "Cette semaine, l'agenda d'Anaïs est complet. De nouveaux créneaux s'ouvrent dans quelques jours : réserve dès maintenant pour bloquer le tien.\n\n{link}"
O8_SIGN = "Alex, l'associé d'Anaïs Brault"
O8_TPL = {k: "Hello {p},\n\n" + v + "\n\n" + O8_DISPO + "\n\n" + O8_SIGN for k, v in {
    "cand": "Merci pour ta candidature au coaching avec Anaïs. Elle l'a lue et veut en parler avec toi de vive voix : sa Selfty Academy ouvre le 10 octobre.",
    "chaud": "On avait échangé sur l'école de coaching d'Anaïs : elle ouvre le 10 octobre et la première promo se remplit. Le plus simple pour savoir si c'est le bon moment, c'est un appel avec elle.",
    "annule": "Ton appel avec Anaïs du {date} a été annulé, et l'école ouvre le 10 octobre. Si l'envie est toujours là, reprends un créneau avec elle.",
    "bourse": "On a bien reçu ta candidature pour la bourse et Anaïs lit tout elle-même. Quelle que soit sa décision, elle veut t'entendre de vive voix : l'école ouvre le 10 octobre.",
    "ecole": "Tu avais répondu au questionnaire de l'école de coaching d'Anaïs : elle ouvre le 10 octobre et je reprends contact avec toute la liste d'attente. Le plus simple, c'est un appel avec Anaïs.",
    "webi2": "Tu as suivi deux lives d'Anaïs, donc le sujet te parle. Sa Selfty Academy ouvre le 10 octobre : un appel avec elle te dira si c'est fait pour toi.",
    "webi1": "Tu avais pris ta place au live d'Anaïs du {live}. Sa Selfty Academy ouvre le 10 octobre : si devenir coach (ou aller plus loin dans ta pratique) te parle, un appel avec elle te dira vite si c'est fait pour toi.",
    "lead": "Tu avais commencé ta candidature pour la Selfty Academy sans réserver ton appel avec Anaïs. L'école ouvre le 10 octobre : 45 minutes avec elle, sans engagement, pour voir si c'est fait pour toi.",
    "froid": "On avait échangé au sujet de l'école de coaching d'Anaïs. Elle ouvre le 10 octobre : si c'est toujours d'actualité pour toi, un appel avec elle te dira si c'est le bon moment.",
}.items()}

def o8_tel(v):
    t = norm_phone(v)
    if t.startswith("330") and len(t) == 12:
        t = "33" + t[3:]   # « +33 06… » saisi avec le 0
    return t

def o8_nk(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return frozenset(t for t in re.findall(r"[a-z]{2,}", s) if t not in ("ans", "an", "de", "la", "le", "du", "des", "et", "test"))

def o8_clean_name(s):
    s = str(s or "").split("\n")[0]
    s = re.sub(r"[\s,\-]+\d{2,3}(\s*ans?)?\b.*$", "", s)   # « Bastien Demoy 28 ans », « Albert Neo 25 », « Zoé X 26ans » -> sans l'âge
    s = re.sub(r"[^\w\s'’\-]", "", s, flags=re.UNICODE)  # emoji, etc.
    return re.sub(r"\s+", " ", s).strip()

def o8_dd(iso):
    try:
        d = datetime.datetime.fromisoformat(str(iso).replace("Z", "+00:00"))
        if TZ_PARIS and d.tzinfo:
            d = d.astimezone(TZ_PARIS)
        return d.strftime("%d/%m")
    except Exception:
        return ""

def o8_is_test(mail, nom):
    m = (mail or "").lower(); n = (nom or "").lower()
    return m in TEST_EMAILS or "+test" in m or m.startswith("test@") or (n.startswith("alex") and "test" in n)

o8, o8_by_mail, o8_by_tel, o8_by_name = [], {}, {}, []
def o8_find(mail, tel, nom):
    if mail and mail in o8_by_mail:
        return o8_by_mail[mail]
    if tel and tel in o8_by_tel:
        return o8_by_tel[tel]
    nk = o8_nk(nom)
    if len(nk) >= 2:
        for k, p in o8_by_name:
            if nk <= k or k <= nk:
                return p
    return None

def o8_add(mail, tel, nom, prenom=""):
    mail = str(mail or "").strip().lower()
    if "@" not in mail:
        mail = ""
    mail = o8_alias.get(mail, mail)
    tel = o8_tel(tel)
    nom = o8_clean_name(nom)
    p = o8_find(mail, tel, nom)
    if not p:
        p = {"names": [], "prenom": "", "mail": "", "tel": "", "sig": [], "info": [], "lives": set(),
             "client": "", "excl": "", "non": "", "upcoming": "", "calls": []}
        o8.append(p)
    if mail and not p["mail"]:
        p["mail"] = mail
    if tel and not p["tel"]:
        p["tel"] = tel
    if mail:
        o8_by_mail[mail] = p
    if tel:
        o8_by_tel[tel] = p
    nk = o8_nk(nom)
    if len(nk) >= 2 and not any(p is q and k == nk for k, q in o8_by_name):
        o8_by_name.append((nk, p))
    if nom and nom not in p["names"]:
        p["names"].append(nom)
    if prenom and not p["prenom"]:
        p["prenom"] = prenom.strip()
    return p

def o8_sig(p, seg, ts, why):
    p["sig"].append({"seg": seg, "lvl": O8_LVL[seg], "ts": ts or "", "why": why})

def o8_info(p, pairs):
    seen = {a for a, _ in p["info"]}
    for a, b in pairs:
        if b and a not in seen:
            p["info"].append([a, str(b).strip()[:600]])
            seen.add(a)

# 1) inscrites aux lives (candidature coaching live = Mode coaching)
for i in inscrits:
    p = o8_add(i["mail"], i["tel"], i["n"], i["n"])
    p["lives"].add(i["live"])
    if i.get("deja1"):
        p["lives"].add("31aout")
    if i.get("deja2"):
        p["lives"].add("9sept")
    o8_info(p, i["det"])
    if i["coach"]:
        accord = any(a == "Accord coaching en direct" for a, _ in i["det"])
        o8_sig(p, "cand", i["ts"], f"Candidature au coaching live du {o8_dd(i['ts'])}" + (" (accord pour le coaching en direct)" if accord else ""))
# 2) liste d'attente école (formulaire + page du live)
for e in ecole:
    if "doublon" in e["statut"].lower() and e["mail"] and e["mail"] in o8_by_mail:
        o8_info(o8_by_mail[e["mail"]], e["detail"]); continue
    p = o8_add(e["mail"], e["tel"], e["n"])
    o8_info(p, e["detail"] + ([["Commentaire (Sheet)", e["comm"]]] if e["comm"] else []) + [[f"Notes d'appel · {t}", v] for t, v in e["notes"]])
    ch, st = e["chaud"].strip(), e["statut"].strip()
    qui = f", pris par {e['qui']}" if e["qui"] else ""
    if re.match(r"^(chaud|a rappeler|à rappeler|tiède|tiede)", ch, re.I):
        o8_sig(p, "chaud", e["ts"], f"Liste d'attente école : {ch}{qui}" + (f" · statut : {st[:80]}" if st else ""))
    elif re.search(r"hors cible|ne veut pas|autre projet|pas dans la cible", ch + " " + st, re.I):
        p["non"] = (st or ch)[:120]
        o8_sig(p, "froid", e["ts"], f"École : {(st or ch)[:120]}")
    elif not st and not ch:
        o8_sig(p, "ecole", e["ts"], f"Questionnaire école du {o8_dd(e['ts'])} ({e['src']}), jamais traité")
    else:
        o8_sig(p, "froid", e["ts"], f"École{qui} : {st[:120]}")
# 3) calls iClosed : vente -> cliente ; call non annulé (passé ou à venir) -> exclue ; seulement annulé -> à rebooker
_now_iso = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
for c in icalls:
    if str(c["id"]).upper().startswith(("TEST", "DEMO")) or o8_is_test(c["mail"], c["n"]):
        continue
    p = o8_add(c["mail"], c["tel"], c["n"], c["n"].split(" ")[0])
    o8_info(p, c["quest"])
    t = c.get("trk") or {}
    if str(t.get("r", "")).lower().startswith("vente"):
        p["client"] = "vente enregistrée dans Suivi Calls"; continue
    if c["cancel"]:
        o8_sig(p, "annule", c["utc"], f"Call iClosed du {o8_dd(c['utc'])} annulé" + (" (par elle)" if "CONTACT" in c["cancelWhy"].upper() else "") + ", à rebooker")
    elif c["utc"]:
        p["calls"].append(c["utc"])
        if c["utc"][:19] > _now_iso and (not p["upcoming"] or c["utc"] < p["upcoming"]):
            p["upcoming"] = c["utc"]
# ventes de Suivi Calls dont le call n'est plus dans iClosed
for r in (track_rows or []):
    if str(r.get("Résultat") or "").lower().startswith("vente") and not str(r.get("Call ID") or "").upper().startswith(("TEST", "DEMO")):
        o8_add(r.get("E-mail"), r.get("Téléphone"), r.get("Nom"))["client"] = "vente enregistrée dans Suivi Calls"
# 4) contacts iClosed sans call
for l in leads:
    p = o8_add(l["mail"], l["tel"], l["n"], l["n"].split(" ")[0])
    o8_sig(p, "lead", l["cree"], f"Formulaire iClosed commencé le {o8_dd(l['cree'])}, pas de call réservé" + (f" · suivi : {l['statut']}" if l["statut"] else ""))
# 5) bourse (Tally)
for s in schol_subs:
    p = o8_add(s["mail"], s["tel"], s["n"], s["n"].split(" ")[0])
    o8_info(p, [[f"Bourse · {q[:60]}", a] for q, a in s["det"]])
    sit = next((a for q, a in s["det"] if q.startswith("Aujourd")), "")
    o8_sig(p, "bourse", s["at"], f"Candidature bourse du {o8_dd(s['at'])} ({'complète' if s['done'] else 'pas finie'})" + (f" · « {sit} »" if sit else "") + " · a dit ne pas pouvoir financer l'Academy")
# 6) coaching individuel (Tally)
for s in coach_subs:
    p = o8_add("", s["tel"], s["n"], s["prenom"])
    o8_info(p, [["Situation idéale", s["ideal"]], ["Ce qui pourrait l'en empêcher", s["frein"]], ["Revenus", s["revenu"]]])
    o8_sig(p, "cand", s["at"], f"Candidature au coaching individuel du {o8_dd(s['at'])}" + ("" if s["done"] else " (pas finie)"))
# 7) clientes (onglet Clients + contrats envoyés)
for c in clientes:
    o8_add(c["mail"], c["tel"], c["n"])["client"] = "fiche Clientes"
for k, c in contrats.items():
    o8_add(k, "", (c["prenom"] + " " + c["nom"]).strip())["client"] = f"contrat {c['statut'].lower()}"
# 8) exclusions du fichier local / secret (anciennes clientes hors console, mineur, hors cible…)
for x in (o8cfg.get("exclure") or []):
    p = o8_find(str(x.get("mail") or "").lower(), o8_tel(x.get("tel")), o8_clean_name(x.get("nom")))
    if p:
        p["excl"] = str(x.get("pourquoi") or "exclue")
# 9) coche partagée (onglet « Rappels objectif 8 », écrit par la console via rappel_set)
rappel_rows = read_tab(ewb, "rappels objectif 8")
rappels = {}
for r in (rappel_rows or []):
    rec = {"c": str(r.get("Contactée") or "").strip().lower() == "oui", "cd": dstr_(r.get("Contactée le")),
           "b": str(r.get("Call booké") or "").strip().lower() == "oui", "bd": dstr_(r.get("Booké le")),
           "note": str(r.get("Note") or "").strip(), "maj": dstr_(r.get("MAJ"))}
    for k in (str(r.get("E-mail") or "").strip().lower(), o8_tel(r.get("Téléphone")), str(r.get("Clé") or "").strip().lower()):
        if k:
            rappels[k] = rec

def o8_prenom(p):
    for k in (p["mail"], p["tel"]):
        if k and k in o8_prenoms:
            return o8_prenoms[k]
    pre = (p["prenom"] or (p["names"][0].split(" ")[0] if p["names"] else "") or "").strip()
    if pre and (pre.isupper() or pre.islower()):
        pre = "-".join(w.capitalize() for w in pre.split("-"))
    return pre

obj8_list, obj8_exclues = [], []
for p in o8:
    nom = max(p["names"], key=lambda s: (min(len(s.split()), 4), -p["names"].index(s))) if p["names"] else (p["mail"] or ("+" + p["tel"] if p["tel"] else "?"))
    if o8_is_test(p["mail"], nom):
        continue
    rap = rappels.get(p["mail"]) or rappels.get(p["tel"]) or {}
    if p["client"] or p["excl"]:
        obj8_exclues.append({"n": nom, "pourquoi": p["excl"] or p["client"]}); continue
    if p["calls"]:
        # call iClosed non annulé (passé ou à venir) = exclue ; exception : cochée « Contactée » avant ce call
        # (booké grâce à la relance Objectif 8) -> reste visible en « Call booké » pour le suivi
        cd = str(rap.get("cd") or "")[:10] if rap.get("c") else ""
        if not (cd and all(u[:10] >= cd for u in p["calls"])):
            passes = sorted(u for u in p["calls"] if u[:19] <= _now_iso)
            obj8_exclues.append({"n": nom, "motif": "call", "pourquoi": f"a déjà eu un call le {o8_dd(passes[-1])}" if passes else f"call déjà booké le {o8_dd(p['upcoming'])}"}); continue
    LIVE_NOMS = {"31aout": "31 août", "9sept": "9 septembre", "21sept": "28 septembre"}
    if len(p["lives"]) >= 2:
        o8_sig(p, "webi2", "", f"A pris sa place à {len(p['lives'])} lives (" + " + ".join(LIVE_NOMS[x] for x in ("31aout", "9sept", "21sept") if x in p["lives"]) + ")")
    elif len(p["lives"]) == 1:
        o8_sig(p, "webi1", "", "A pris sa place au live du " + LIVE_NOMS.get(next(iter(p["lives"])), "31 août"))
    if not p["sig"]:
        continue
    if p["non"] and min(s["lvl"] for s in p["sig"]) > O8_LVL["bourse"]:
        for s in p["sig"]:
            s["lvl"] = O8_LVL["froid"]   # a dit non / hors cible, sans signal fort depuis : tout en bas
    sigs = sorted(p["sig"], key=lambda s: (s["lvl"], s["ts"] == "", s["ts"]))
    best = sigs[0]
    seg = "froid" if best["lvl"] >= O8_LVL["froid"] else best["seg"]
    last = max((s["ts"] for s in p["sig"] if s["ts"]), default="")
    pre = o8_prenom(p)
    toks = [("-".join(w.capitalize() for w in t.split("-")) if (t.isupper() or t.islower()) and len(t) > 2 else t) for t in nom.split()]
    if pre and toks and toks[0].lower() != pre.lower() and pre.lower() in [t.lower() for t in toks]:
        toks = [pre] + [t for t in toks if t.lower() != pre.lower()]
    nom = " ".join(toks)
    why = list(OrderedDict.fromkeys(s["why"] for s in sigs))
    live = "28 septembre" if "21sept" in p["lives"] else "9 septembre" if "9sept" in p["lives"] else "31 août"
    date_call = next((o8_dd(s["ts"]) for s in sigs if s["seg"] == "annule"), "")
    msg = o8_msgs.get(p["mail"]) or o8_msgs.get(p["tel"]) or O8_TPL[seg]
    msg = msg.replace("{p}", pre or "toi").replace("{link}", ICLOSED_LINK).replace("{live}", live).replace("{date}", date_call or "quelques jours")
    obj8_list.append({
        "k": p["mail"] or p["tel"], "n": nom, "prenom": pre, "mail": p["mail"], "tel": p["tel"],
        "lvl": O8_LVL[seg], "seg": seg, "segLab": O8_SEG[seg],
        "why": why, "note": o8_notes.get(p["mail"]) or o8_notes.get(p["tel"]) or "",
        "last": last, "lastLab": o8_dd(last) if last else "",
        "msg": msg, "perso": bool(o8_msgs.get(p["mail"]) or o8_msgs.get(p["tel"])), "info": p["info"][:14],
        "c": bool(rap.get("c")), "cd": rap.get("cd", ""),
        "b": bool(rap.get("b")) or bool(p["calls"]), "bd": rap.get("bd", "") or (f"iClosed, call le {o8_dd(p['upcoming'] or max(p['calls']))}" if p["calls"] else ""),
        "rnote": rap.get("note", ""), "maj": rap.get("maj", ""),
    })
# du plus chaud au plus froid ; à niveau égal, la plus récente d'abord
obj8_list.sort(key=lambda x: x["last"], reverse=True)
obj8_list.sort(key=lambda x: x["lvl"])
# ventes du mois (Suivi Calls, Résultat = Vente, date close sinon date du call), tests exclus
def o8_vente_date(r):
    dc = r.get("Date close")
    if isinstance(dc, datetime.datetime):
        return dc.strftime("%Y-%m-%d")
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})", str(dc or "").strip())
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return str(r.get("Date call") or "")[:10]
ventes_mois, ventes_depuis = 0, 0
_mois = datetime.date.today().strftime("%Y-%m")
for r in (track_rows or []):
    cid = str(r.get("Call ID") or "").strip().upper()
    if not str(r.get("Résultat") or "").lower().startswith("vente") or cid.startswith(("TEST", "DEMO")) or o8_is_test(str(r.get("E-mail") or ""), str(r.get("Nom") or "")):
        continue
    d = o8_vente_date(r)
    ventes_mois += d.startswith(_mois)
    ventes_depuis += d >= OBJ8_DEBUT
obj8 = {"ok": True, "debut": OBJ8_DEBUT, "cible": OBJ8_CIBLE, "link": ICLOSED_LINK, "rappelsOk": rappel_rows is not None,
        "ventesMois": ventes_mois, "ventesDepuis": ventes_depuis, "segments": list(O8_SEG.items()),
        "list": obj8_list, "exclues": len(obj8_exclues), "exclCall": sum(1 for x in obj8_exclues if x.get("motif") == "call"), "persoMsgs": len(o8_msgs)}
print(f"Objectif 8 : {len(obj8_list)} personnes à rappeler ({len(obj8_exclues)} exclues dont {obj8['exclCall']} pour un call iClosed), {ventes_depuis} vente(s) depuis le {OBJ8_DEBUT}, {ventes_mois} sur le mois")

# ---- Mail du sondage : destinataires (calculés ici, embarqués dans la page chiffrée, envoyés par le script Contrats `mail_liste`) ----
# Toutes les personnes uniques par e-mail : inscrites live 1 + live 2, liste d'attente école, contacts + calls iClosed
# (les candidatures coaching Tally n'ont pas d'e-mail : injoignables par mail). Moins : clientes signées, contrats envoyés,
# ventes de Suivi Calls, tests, exclusions objectif8.json (anciennes clientes, mineur, hors cible) et désinscrites (« stop » / « désinscri » dans un statut).
def sd_desinscrite(*vals):
    return any(re.search(r"\bstop\b|d[ée]sinscri|unsubscribe|ne veut plus", str(v or ""), re.I) for v in vals)

sd_excl = set(TEST_EMAILS) | {"selfty.academy@gmail.com"} | {c["mail"] for c in clientes if c["mail"]} | set(contrats.keys())   # l'expéditeur ne s'écrit pas à lui-même
for r in (track_rows or []):
    if str(r.get("Résultat") or "").lower().startswith("vente"):
        sd_excl.add(str(r.get("E-mail") or "").strip().lower())
for x in (o8cfg.get("exclure") or []):
    m = str(x.get("mail") or "").strip().lower()
    if m and not x.get("seulement_objectif8"):   # « seulement_objectif8 » : retirée des relances, garde le mail du sondage
        sd_excl.add(m)
sd_dest, sd_src = OrderedDict(), Counter()
def sd_add(mail, prenom, src, nom=""):
    mail = o8_alias.get(str(mail or "").strip().lower(), str(mail or "").strip().lower())
    if "@" not in mail or "." not in mail.split("@")[-1] or mail in sd_excl or o8_is_test(mail, nom or prenom):
        return
    prenom = str(prenom or "").strip().split(" ")[0]
    if mail not in sd_dest:
        sd_dest[mail] = {"email": mail, "prenom": "", "src": src}
        sd_src[src] += 1
    if prenom and not sd_dest[mail]["prenom"]:
        sd_dest[mail]["prenom"] = prenom
for i in inscrits:
    if sd_desinscrite(i["statut"], i["etape"]):
        sd_excl.add(i["mail"].lower()); continue
    sd_add(i["mail"], i["n"], {"31aout": "live1", "9sept": "live2"}.get(i["live"], "live3"), i["n"])
for e in ecole:
    if sd_desinscrite(e["statut"], e["chaud"]):
        sd_excl.add(e["mail"]); continue
    sd_add(e["mail"], "", "ecole", e["n"])
for ct in ic_contacts:
    sd_add(ct.get("email"), ct.get("firstName"), "iclosed", (str(ct.get("firstName") or "") + " " + str(ct.get("lastName") or "")).strip())
for c in icalls:
    if not str(c["id"]).upper().startswith(("TEST", "DEMO")):
        sd_add(c["mail"], c["n"].split(" ")[0], "iclosed", c["n"])
for m, d in sd_dest.items():
    p = o8_by_mail.get(m)
    pre = o8_prenom(p) if p else ""
    if not pre and not d["prenom"] and p and p["names"]:
        pre = o8_clean_name(p["names"][0]).split(" ")[0]
    if pre:
        d["prenom"] = pre
    if d["prenom"] and (d["prenom"].isupper() or d["prenom"].islower()):
        d["prenom"] = "-".join(w.capitalize() for w in d["prenom"].split("-"))
sd_dest = [d for d in sd_dest.values() if d["email"] not in sd_excl]
print(f"Mail sondage : {len(sd_dest)} destinataires ({dict(sd_src)}), {len(sd_excl)} e-mails exclus")

# ---- Pointage automatique des présences aux calls de groupe via l'API Zoom (inactif tant que les secrets manquent) ----
# Secrets GitHub à créer quand Anaïs a créé l'app Zoom « Server-to-Server OAuth » (scope report:read:admin ou meeting:read:past_participant) :
# ZOOM_ACCOUNT_ID, ZOOM_CLIENT_ID, ZOOM_CLIENT_SECRET, ZOOM_MEETING_ID (id de la réunion récurrente des Selfty Calls).
# Pour chaque occurrence passée depuis le début de l'école qui n'a encore AUCUN pointage : chaque cliente active = Présente si son
# e-mail ou son nom apparaît dans les participants Zoom, sinon Absente -> pont principal presence_set (source zoom). Un pointage
# manuel existant n'est jamais écrasé. Test local : SHOW_TEST=1 + test-zoom.json ({"instances":[{"start_time","participants":[…]}]}).
def zoom_instances():
    zk = {k: os.environ.get("ZOOM_" + k, "") for k in ("ACCOUNT_ID", "CLIENT_ID", "CLIENT_SECRET", "MEETING_ID")}
    if SHOW_TEST and (HERE / "test-zoom.json").exists():
        return json.loads((HERE / "test-zoom.json").read_text()).get("instances", [])
    if not all(zk.values()):
        return None
    auth = base64.b64encode(f"{zk['CLIENT_ID']}:{zk['CLIENT_SECRET']}".encode()).decode()
    req = urllib.request.Request(f"https://zoom.us/oauth/token?grant_type=account_credentials&account_id={zk['ACCOUNT_ID']}",
                                 data=b"", headers={"Authorization": "Basic " + auth})
    token = json.load(urllib.request.urlopen(req, timeout=30))["access_token"]
    H = {"Authorization": "Bearer " + token}
    api = lambda path: json.load(urllib.request.urlopen(urllib.request.Request("https://api.zoom.us/v2" + path, headers=H), timeout=30))
    out = []
    for m in api(f"/past_meetings/{zk['MEETING_ID']}/instances").get("meetings") or []:
        uid = m["uuid"]
        uid = urllib.parse.quote(urllib.parse.quote(uid, safe=""), safe="") if uid.startswith("/") or "//" in uid else urllib.parse.quote(uid, safe="")
        parts, tok = [], ""
        while True:
            d = api(f"/past_meetings/{uid}/participants?page_size=300" + (f"&next_page_token={tok}" if tok else ""))
            parts += d.get("participants") or []
            tok = d.get("next_page_token") or ""
            if not tok:
                break
        out.append({"start_time": m.get("start_time"), "participants": parts})
    return out


import urllib.parse
zoom_pointes = 0
try:
    zi = zoom_instances()
    if zi is not None:
        pont_z = {"url": os.environ.get("PONT_URL", ""), "key": os.environ.get("PONT_KEY", "")}
        if (HERE / "pont.json").exists() and not pont_z["url"]:
            pont_z.update(json.loads((HERE / "pont.json").read_text()))
        deja_dates = {p_["date"] for p_ in presences}
        actives_z = [c for c in clientes if c["mail"] and c["statut"] not in ("Terminée", "En pause")]
        norm_z = lambda t: re.sub(r"[^a-z]", "", unicodedata.normalize("NFD", str(t or "").lower()))
        for inst in zi:
            st = str(inst.get("start_time") or "")
            try:
                dt = datetime.datetime.fromisoformat(st.replace("Z", "+00:00")).astimezone(ZoneInfo("Europe/Paris"))
            except ValueError:
                continue
            if dt.strftime("%Y-%m-%d") < ECOLE_DEBUT or dt.strftime("%d/%m/%Y") in deja_dates:
                continue
            date_fr = dt.strftime("%d/%m/%Y")
            mails_z = {str(x.get("user_email") or "").strip().lower() for x in inst["participants"]}
            noms_z = [norm_z(x.get("name")) for x in inst["participants"]]
            for c in actives_z:
                nk = norm_z(c["n"])
                present = c["mail"].lower() in mails_z or (len(nk) >= 5 and any(nk in n_ or (n_ and len(n_) >= 5 and n_ in nk) for n_ in noms_z))
                if not SHOW_TEST and pont_z["url"]:
                    body = json.dumps({"what": "presence_set", "key": pont_z["key"], "date": date_fr, "email": c["mail"], "nom": c["n"], "present": present, "source": "zoom"}).encode()
                    urllib.request.urlopen(urllib.request.Request(pont_z["url"], data=body, headers={"Content-Type": "text/plain"}), timeout=60)
                presences.append({"date": date_fr, "mail": c["mail"].lower(), "nom": c["n"], "present": present, "src": "zoom"})
                zoom_pointes += 1
            deja_dates.add(date_fr)
        print(f"Zoom : {len(zi)} occurrence(s), {zoom_pointes} pointage(s) ajouté(s)")
except Exception as ex:
    print("Zoom : pointage impossible", ex)

# ---- Exercices rendus (pont « Exercices Selfty ») + Engagements des élèves (pont « Engagements Selfty ») ----
# Les 2 Sheets sont privés (Drive de selfty.academy) : on passe par les ponts Apps Script. Échec = onglet vide + bandeau.
# Clés : env EXOS_KEY / ENG_KEY / ENG_CKEY (secrets CI) ou fichiers locaux exos-key.txt / ../selfty-engagements/pont/pont-key.txt / eng-console-key.txt
EXOS_URL = "https://script.google.com/macros/s/AKfycbyh4K7PfXc9r8v9GCqMzTlPdKzJ-AYlxchLhUOEJl4tXJuaxBJ424Nc5vwh0L9nh-cc/exec"
ENG_URL = "https://script.google.com/macros/s/AKfycbwROj9t1ce54CnwjVJ7piRzVhqug2eKzyjy8y3BhvpJcKDrXHEIF_YmZQ169cgw1zc/exec"


def _local_key(rel):
    f = HERE / rel
    return f.read_text().strip() if f.exists() else ""


def pont_post(url, payload, timeout=90):
    req = urllib.request.Request(url, data=json.dumps(payload).encode(), headers={"Content-Type": "text/plain"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:   # 302 -> GET suivi par urllib, comme curl -L
        return json.load(resp)


exos_key = os.environ.get("EXOS_KEY") or _local_key("exos-key.txt")
exos_ok, exos = False, []
if exos_key:
    try:
        j = pont_post(EXOS_URL, {"key": exos_key, "what": "exercices"})
        if j.get("ok"):
            exos_ok = True
            for x in j.get("exercices") or []:
                mail = str(x.get("email") or "").strip().lower()
                if mail in TEST_EMAILS and not SHOW_TEST:
                    continue
                exos.append({k: ("" if x.get(k) is None else str(x.get(k))) for k in
                             ("date", "prenom", "email", "module", "lecon", "exercice", "question", "fichiers", "statut", "commentaire", "id", "lien")})
        else:
            print("Exercices : refus du pont", j.get("error"))
    except Exception as ex:
        print("Exercices : pont injoignable", ex)
if SHOW_TEST and (HERE / "test-exos.json").exists():
    exos_ok = True
    exos += json.loads((HERE / "test-exos.json").read_text())
print(f"Exercices : {len(exos)} ({sum(1 for x in exos if x['statut'] != 'Corrigé')} à corriger)")

eng_key = os.environ.get("ENG_KEY") or _local_key("../selfty-engagements/pont/pont-key.txt")
eng_ckey = os.environ.get("ENG_CKEY") or _local_key("eng-console-key.txt")
eng = {"ok": False, "eleves": [], "rappels": [], "today": ""}
if eng_key and eng_ckey:
    try:
        j = pont_post(ENG_URL, {"key": eng_key, "ckey": eng_ckey, "what": "console"})
        if j.get("ok"):
            keep = lambda m: SHOW_TEST or str(m or "").lower() not in TEST_EMAILS
            eng = {"ok": True, "today": j.get("today", ""), "semaine": j.get("semaine"), "promo": j.get("promo"), "intakeForm": j.get("intake_form", ""),
           "semaineIso": j.get("semaineIso", ""), "semainePrecIso": j.get("semainePrecIso", ""),
                   "eleves": [e for e in j.get("eleves") or [] if keep(e.get("email"))],
                   "rappels": [r for r in j.get("rappels") or [] if keep(r.get("email"))]}
        else:
            print("Engagements : refus du pont", j.get("error"))
    except Exception as ex:
        print("Engagements : pont injoignable", ex)
if SHOW_TEST and (HERE / "test-engagements.json").exists():
    t = json.loads((HERE / "test-engagements.json").read_text())
    eng["ok"] = True
    eng["eleves"] += t.get("eleves", [])
    eng["rappels"] += t.get("rappels", [])
print(f"Engagements : {len(eng['eleves'])} élèves, {sum(1 for r in eng['rappels'] if not r.get('traite'))} relances à traiter")

# ---- Questionnaire d'arrivée des élèves (form Tally 81X5kk « Ton point de départ », même clé) ----
# e-mail = hidden field `email` du lien perso, sinon le champ « E-mail » ; seulement les soumissions complètes
INTAKE_FORM = "81X5kk"

def intake_parse(d, qlabels, out):
    for q in d.get("questions") or []:
        qlabels[q["id"]] = (str(q.get("title") or "").strip(), str(q.get("type") or ""))
    for s in d.get("submissions") or []:
        if s.get("isCompleted") is False:
            continue
        hid, mail = {}, ""
        for r in s.get("responses") or []:
            lab, qtype = qlabels.get(r.get("questionId"), ("", ""))
            ans = r.get("answer")
            if qtype == "HIDDEN_FIELDS" and isinstance(ans, dict):
                hid = ans
            elif lab == "E-mail" and not mail:
                mail = ty_txt(ans).strip().lower()
        mail = str(hid.get("email") or "").strip().lower() or mail
        if not mail or (mail in TEST_EMAILS and not SHOW_TEST):
            continue
        at = str(s.get("submittedAt") or "")
        if mail not in out or at < out[mail]:
            out[mail] = at   # première soumission complète

intake_mails, intake_ok = {}, False
if ty_key:
    try:
        qlabels, page = {}, 1
        while True:
            req = urllib.request.Request(
                f"https://api.tally.so/forms/{INTAKE_FORM}/submissions?filter=completed&page={page}",
                headers={"Authorization": "Bearer " + ty_key, "User-Agent": "curl/8.4.0"})
            d = json.load(urllib.request.urlopen(req, timeout=30))
            intake_parse(d, qlabels, intake_mails)
            if not d.get("hasMore"):
                break
            page += 1
        intake_ok = True
    except Exception as ex:
        print("Tally questionnaire d'arrivée KO (on garde la console sans) :", ex)
# repli : date d'intake connue du pont Engagements (webhook Tally -> onglet Élèves)
for e in eng.get("eleves") or []:
    m = str(e.get("email") or "").strip().lower()
    if m and e.get("intake_le") and m not in intake_mails:
        intake_mails[m] = str(e.get("intake_le"))
print(f"Questionnaire d'arrivée : {len(intake_mails)} rempli(s)")

data = {
    "maj": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
    "ecoleDebut": ECOLE_DEBUT,
    "eow": {"ok": eow_ok, "url": f"https://tally.so/r/{EOW_FORM}", "subs": eow_subs},
    "presences": presences,
    "exos": {"ok": exos_ok, "list": exos, "form": "https://tally.so/r/dWOQbz", "sheet": "https://docs.google.com/spreadsheets/d/1lb7WpgYw5YDQXaurQZ9nk_hMZ35sXpwDcS7us428uYg/edit"},
    "eng": eng,
    "webi": {
        "label": "Live du mercredi 9 septembre, 18h",
        "meet": "https://us06web.zoom.us/j/88555750551?pwd=LjHAbfU8giQRrlGs6a6LRTggG3Sd8K.1&jst=2",
        "meetPrec": "https://meet.google.com/oxf-vzjg-bhr",
        "groupe": "https://chat.whatsapp.com/JRCRXWKVg8qBUS0uUUTkCA?mode=gi_t",
        "lp": "https://selfty-academy.github.io/live-9-septembre/",
        "lpPrec": "https://selfty-academy.github.io/live-31-aout/",
    },
    "live3": {"label": "Live du lundi 28 septembre, 18h", "lp": "https://selfty-academy.github.io/live-28-septembre/",
              "visites": dict(live3_vis), "visitesSrc": dict(live3_vis_src)},
    "inscrits": sorted(inscrits, key=lambda x: x["ts"], reverse=True),
    "cands": sorted(cands, key=lambda c: (c["accord"] != "oui", c["ts"])),
    "la": la,
    "laDeja": la_deja,
    "ecole": ecole,
    "ecoleUniques": ecole_uniques,
    "icalls": {"ok": ic_ok, "calls": sorted(icalls, key=lambda x: x["utc"])},
    "leads": {"ok": leads_ok, "list": sorted(leads, key=lambda x: x["cree"], reverse=True)},
    "suivi": {
        "trackOk": track_rows is not None,
        "clientsOk": clients_rows is not None,
        "clientes": clientes,
        "contratsOk": contrats_rows is not None,
        "contrats": list(contrats.values()),
    },
    "schol": {"ok": schol_ok, "url": "https://tally.so/r/Np1Gy0",
              "stats": schol_stats, "subs": schol_subs},
    "coach": {"ok": coach_ok, "url": f"https://tally.so/r/{COACH_FORM}", "subs": coach_subs},
    "sondage": {"ok": sond_ok, "url": f"https://tally.so/r/{SOND_FORM}", "sujets": [c for _, c in SOND_SUJETS],
                "sources": SOND_SOURCES, "subs": sond_subs,
                "dest": sd_dest, "destSrc": dict(sd_src), "destExclus": len(sd_excl)},
    "obj8": obj8,
    "intake": {"ok": intake_ok, "url": f"https://tally.so/r/{INTAKE_FORM}", "mails": intake_mails},
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

# Pont Apps Script (suivi école éditable) : pont.json local ou env (CI)
pont = {"url": "", "key": ""}
pont_file = HERE / "pont.json"
if pont_file.exists():
    pont.update(json.loads(pont_file.read_text()))
pont["url"] = os.environ.get("PONT_URL", pont["url"])
pont["key"] = os.environ.get("PONT_KEY", pont["key"])

# Bot Telegram « SELFTY » (notif vente envoyée depuis la page) : telegram.json local ou env (CI)
tg = {"token": "", "chat_id": ""}
tg_file = HERE / "telegram.json"
if tg_file.exists():
    tg.update(json.loads(tg_file.read_text()))
tg["token"] = os.environ.get("TG_TOKEN", tg["token"])
tg["chat_id"] = os.environ.get("TG_CHAT", tg["chat_id"])

tpl = (HERE / "template.html").read_text()
# garde-fou général : un identifiant ou un téléphone qui ressort « 2509944.0 » = format d'export Google
# encore changé -> alerte Telegram (le bug du 16/09 avait vidé « Calls effectués » sans prévenir)
_flottants = re.findall(r'"(\d{5,}\.0)"', json.dumps(data, ensure_ascii=False))
if _flottants:
    print("⚠ ALERTE : valeurs en « .0 » dans les données :", _flottants[:5])
    if os.environ.get("TG_TOKEN") and os.environ.get("TG_CHAT"):
        try:
            import urllib.parse
            urllib.request.urlopen(urllib.request.Request(
                f"https://api.telegram.org/bot{os.environ['TG_TOKEN']}/sendMessage",
                data=urllib.parse.urlencode({"chat_id": os.environ["TG_CHAT"], "text":
                    f"⚠ Console Selfty : {len(_flottants)} valeurs lues avec un « .0 » (ex. {_flottants[0]}). "
                    "Le format d'export du Sheet a changé, IDs et téléphones à vérifier."}).encode()), timeout=10)
        except Exception as ex:
            print("alerte Telegram KO :", ex)
out = (tpl.replace("__DATA__", json.dumps(data, ensure_ascii=False)).replace("__LOGO__", logo)
       .replace("__PONT_URL__", pont["url"]).replace("__PONT_KEY__", pont["key"])
       .replace("__TG_TOKEN__", tg["token"]).replace("__TG_CHAT__", tg["chat_id"])
       .replace("__EXOS_URL__", EXOS_URL).replace("__EXOS_KEY__", exos_key)
       .replace("__ENG_URL__", ENG_URL).replace("__ENG_KEY__", eng_key).replace("__ENG_CKEY__", eng_ckey)
       .replace("__DRIVE_CONTRATS__", DRIVE_CONTRATS).replace("__DRIVE_FACTURES__", DRIVE_FACTURES))
(HERE / "console.html").write_text(out)
print(f"console.html : {len(inscrits)} inscrits, {len(cands)} candidatures live, {len(ecole)} lignes école ({ecole_uniques} personnes), {len(visites)} visites")

# ---- Objectif 8 : liste complète pour Alex, fichier LOCAL seulement (jamais en CI, jamais dans le repo) ----
if (HERE / "objectif8.json").exists() and not os.environ.get("CI"):
    md = [f"# Objectif 8 : qui rappeler, dans l'ordre ({datetime.datetime.now().strftime('%d/%m/%Y %H:%M')})", "",
          f"{len(obj8_list)} personnes à rappeler, du plus chaud au plus froid. Une seule proposition par message : un appel avec Anaïs ({ICLOSED_LINK}).",
          f"Ventes depuis le {OBJ8_DEBUT} : {ventes_depuis} / {OBJ8_CIBLE} (dont {ventes_mois} sur le mois). Coches partagées : onglet « Rappels objectif 8 » du Sheet École.", ""]
    for seg, lab in O8_SEG.items():
        grp = [x for x in obj8_list if x["seg"] == seg]
        if not grp:
            continue
        md += [f"## {O8_LVL[seg]}. {lab} ({len(grp)})", ""]
        for x in grp:
            etat = " · CONTACTÉE" if x["c"] else ""
            etat += " · CALL BOOKÉ" if x["b"] else ""
            md += [f"### {x['n']}{etat}", f"- Téléphone : {'+' + x['tel'] if x['tel'] else 'pas de numéro'} · {x['mail'] or 'pas d’e-mail'} · dernier signe : {x['lastLab'] or '?'}"]
            md += [f"- Pourquoi : {w}" for w in x["why"]]
            if x["note"]:
                md.append(f"- Note : {x['note']}")
            md += ["- Message" + (" (perso)" if x["perso"] else " (modèle)") + " :", "", "```", x["msg"], "```", ""]
    if obj8_exclues:
        md += ["## Exclues (pas à rappeler)", ""] + [f"- {x['n']} : {x['pourquoi']}" for x in obj8_exclues] + [""]
    (HERE / "objectif-8.md").write_text("\n".join(md))
    print(f"objectif-8.md : {len(obj8_list)} personnes")
