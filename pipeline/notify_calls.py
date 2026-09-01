#!/usr/bin/env python3
"""Notif Telegram à Alex quand un nouveau call est booké sur l'iClosed Selfty.

Compare les ids de icalls.json (écrit par build.py à chaque refresh) à l'état
commité state-icalls.json. L'état ne contient QUE des ids (aucune donnée perso :
le repo est public). Premier passage sans état : on initialise sans notifier.
Env : TG_TOKEN + TG_CHAT (mêmes valeurs que le bot Console I3).
"""
import datetime
import json
import os
import sys
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
CALLS = os.path.join(HERE, "icalls.json")
STATE = os.path.join(HERE, "state-icalls.json")

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("Europe/Paris")
except Exception:
    TZ = None


def quand(utc):
    try:
        d = datetime.datetime.fromisoformat(utc.replace("Z", "+00:00"))
        if TZ:
            d = d.astimezone(TZ)
        jours = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
        mois = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet",
                "août", "septembre", "octobre", "novembre", "décembre"]
        return f"{jours[d.weekday()]} {d.day} {mois[d.month - 1]} à {d.strftime('%Hh%M')}"
    except Exception:
        return utc or "?"


def main():
    token = os.environ.get("TG_TOKEN", "").strip()
    chat = os.environ.get("TG_CHAT", "").strip()
    if not token or not chat:
        print("notify : pas de TG_TOKEN/TG_CHAT, on saute")
        return
    if not os.path.exists(CALLS):
        print("notify : pas de icalls.json (fetch iClosed KO ?), on saute")
        return
    calls = json.load(open(CALLS))
    ids = sorted(str(c["id"]) for c in calls)
    first = not os.path.exists(STATE)
    old = set() if first else set(json.load(open(STATE)).get("ids", []))
    sent = 0
    if not first:
        for c in calls:
            if str(c["id"]) in old or c.get("cancel"):
                continue
            msg = (f"📞 Selfty : nouveau call booké\n"
                   f"{c.get('n', '?')} — {quand(c.get('utc', ''))}\n"
                   f"{c.get('event', '')}".strip())
            data = urllib.parse.urlencode({"chat_id": chat, "text": msg}).encode()
            urllib.request.urlopen(
                f"https://api.telegram.org/bot{token}/sendMessage", data, timeout=30)
            sent += 1
    tmp = STATE + ".tmp"
    with open(tmp, "w") as f:
        json.dump({"ids": ids}, f)
    os.replace(tmp, STATE)
    print(f"notify : {'init sans notif' if first else str(sent) + ' notif(s)'}, {len(ids)} calls en état")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # jamais bloquant pour le workflow
        print("notify KO :", e)
        sys.exit(0)
