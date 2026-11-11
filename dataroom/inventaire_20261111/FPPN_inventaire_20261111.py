# -*- coding: utf-8 -*-
"""
FPPN - Inventaire du dossier « 11-Documentations au 11112026 » et préparation de l'envoi des documents nouveaux.
Python 3.8 ou plus, bibliothèque standard uniquement. Aucun fichier d'origine n'est modifié ni déplacé.

Ce que fait le script :
  1. il parcourt tout le dossier (sous-dossiers compris) et calcule l'empreinte SHA-256 de chaque fichier ;
  2. il compare chaque empreinte à FPPN_empreintes_connues_20261111.csv (posé à côté du script) :
       - « Déjà reçu »                    : contenu identique à un fichier déjà transmis à Adenka ;
       - « Présent au 09/11, non transmis » : contenu identique à un fichier de la dataroom du 09/11 non encore envoyé ;
       - « Nouveau »                      : contenu jamais vu (y compris une nouvelle version d'un fichier connu) ;
       - « Doublon interne »              : copie exacte d'un autre fichier du même dossier (un seul exemplaire est envoyé) ;
  3. il copie les fichiers « Nouveau » dans <sortie>\\A_envoyer en gardant les sous-dossiers ;
  4. il range ces fichiers dans des archives .zip de 24 Mo au plus ; un fichier plus gros est zippé seul puis coupé
     en morceaux de 24 Mo (« ..._morceau01sur03.zip ») qu'Adenka recolle ;
  5. il écrit l'inventaire complet (CSV), la liste du contenu des archives et un résumé (texte).

Mode d'emploi (invite de commandes ou PowerShell) :
  python FPPN_inventaire_20261111.py
  python FPPN_inventaire_20261111.py "C:\\autre\\dossier" --sortie "C:\\autre\\sortie" --taille 24
  Option --avec-non-transmis : ajoute aussi à l'envoi les fichiers « Présent au 09/11, non transmis ».
"""
import argparse
import csv
import hashlib
import os
import shutil
import sys
import time
import zipfile
from pathlib import Path

DOSSIER = r"C:\Dev\Missions\05-FPPN\04-Documents reçus du client\11-Documentations au 11112026"
DATE = "20261111"
MO = 1024 * 1024


def long(p):
    """Chemin long Windows (plus de 260 caractères)."""
    s = str(Path(p).resolve())
    if os.name == "nt" and not s.startswith("\\\\?\\"):
        s = "\\\\?\\UNC\\" + s[2:] if s.startswith("\\\\") else "\\\\?\\" + s
    return s


def empreinte(p):
    h = hashlib.sha256()
    with open(long(p), "rb") as f:
        for bloc in iter(lambda: f.read(4 * MO), b""):
            h.update(bloc)
    return h.hexdigest()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dossier", nargs="?", default=DOSSIER)
    ap.add_argument("--sortie", default=None)
    ap.add_argument("--taille", type=int, default=24, help="taille maximale d'une archive, en Mo")
    ap.add_argument("--avec-non-transmis", action="store_true")
    a = ap.parse_args()

    src = Path(a.dossier)
    if not src.is_dir():
        sys.exit(f"Dossier introuvable : {src}")
    sortie = Path(a.sortie) if a.sortie else src.parent / f"FPPN_Envoi_{DATE}"
    ici = Path(__file__).resolve().parent
    ref_csv = ici / f"FPPN_empreintes_connues_{DATE}.csv"
    if not ref_csv.is_file():
        sys.exit(f"Fichier de référence introuvable : {ref_csv} (le placer à côté du script)")
    connus = {}
    with open(ref_csv, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f, delimiter=";"):
            connus[r["sha256"].lower()] = (r["statut"], r["chemin_connu"])
    # index des chemins connus (fin de chemin) pour signaler les nouvelles versions
    fins = {}
    for _, (st, ch) in connus.items():
        el = ch.replace("/", "\\").lower().split("\\")
        for k in (1, 2, 3):
            if len(el) >= k:
                fins.setdefault("\\".join(el[-k:]), ch)

    sortie.mkdir(parents=True, exist_ok=True)
    a_env = sortie / "A_envoyer"
    t0 = time.time()
    fichiers = sorted(p for p in src.rglob("*") if p.is_file())
    print(f"{len(fichiers)} fichiers trouvés dans {src}")
    vus, lignes, envoi = {}, [], []
    for i, p in enumerate(fichiers, 1):
        rel = str(p.relative_to(src))
        try:
            st_ = os.stat(long(p))
            taille, dmod = st_.st_size, time.strftime("%d/%m/%Y %H:%M", time.localtime(st_.st_mtime))
            h = empreinte(p)
        except OSError as e:
            lignes.append([rel, p.suffix.lower(), "", "", "", "Illisible", "", str(e)])
            continue
        remarque, chemin_connu = "", ""
        if h in vus:
            statut, remarque = "Doublon interne", f"copie de {vus[h]}"
        elif h in connus:
            statut = "Déjà reçu" if connus[h][0] == "Reçu" else "Présent au 09/11, non transmis"
            chemin_connu = connus[h][1]
        else:
            statut = "Nouveau"
            el = rel.lower().split("\\") if os.name == "nt" else rel.lower().split("/")
            for k in (3, 2, 1):
                if len(el) >= k and "\\".join(el[-k:]) in fins:
                    remarque = "nouvelle version d'un fichier connu" if k > 1 or len(p.name) > 12 else "même nom qu'un fichier connu"
                    chemin_connu = fins["\\".join(el[-k:])]
                    break
        vus.setdefault(h, rel)
        lignes.append([rel, p.suffix.lower(), taille, dmod, h, statut, chemin_connu, remarque])
        if statut == "Nouveau" or (a.avec_non_transmis and statut == "Présent au 09/11, non transmis"):
            envoi.append((p, rel, taille, h))
        if i % 200 == 0:
            print(f"  {i}/{len(fichiers)} fichiers lus ({time.time() - t0:.0f} s)")

    inv = sortie / f"FPPN_inventaire_{DATE}.csv"
    with open(inv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["chemin_relatif", "extension", "taille_octets", "date_modification", "sha256", "statut", "chemin_deja_connu", "remarque"])
        w.writerows(lignes)

    # copie et archives
    if a_env.exists():
        shutil.rmtree(long(a_env))
    for p, rel, _, _ in envoi:
        d = a_env / rel
        os.makedirs(long(d.parent), exist_ok=True)
        shutil.copy2(long(p), long(d))
    limite = a.taille * MO
    zips, cont, lot, cumul = [], [], [], 0

    def fermer(lot_):
        if not lot_:
            return
        nom = sortie / f"FPPN_Envoi_{DATE}_{len(zips) + 1:02d}.zip"
        with zipfile.ZipFile(long(nom), "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
            for p, rel, t, h in lot_:
                z.write(long(p), rel)
                cont.append([nom.name, rel, t, h])
        zips.append(nom)

    for e in sorted(envoi, key=lambda x: x[1]):
        if e[2] > limite:
            continue
        if cumul + e[2] > limite and lot:
            fermer(lot)
            lot, cumul = [], 0
        lot.append(e)
        cumul += e[2]
    fermer(lot)
    for p, rel, t, h in [e for e in envoi if e[2] > limite]:
        nom = sortie / f"FPPN_Envoi_{DATE}_{len(zips) + 1:02d}_gros.zip"
        with zipfile.ZipFile(long(nom), "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
            z.write(long(p), rel)
        tz = os.path.getsize(long(nom))
        if tz <= limite:
            zips.append(nom)
            cont.append([nom.name, rel, t, h])
            continue
        n = -(-tz // limite)
        with open(long(nom), "rb") as f:
            for k in range(n):
                m = nom.with_name(f"{nom.stem}_morceau{k + 1:02d}sur{n:02d}.zip")
                with open(long(m), "wb") as g:
                    g.write(f.read(limite))
                cont.append([m.name, rel, t, h])
                zips.append(m)
        os.remove(long(nom))
    with open(sortie / f"FPPN_contenu_envoi_{DATE}.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["archive", "chemin_relatif", "taille_octets", "sha256"])
        w.writerows(cont)

    from collections import Counter
    c = Counter(l[5] for l in lignes)
    tot = sum(l[2] for l in lignes if isinstance(l[2], int))
    res = [f"Inventaire FPPN du dossier « {src.name} » ({time.strftime('%d/%m/%Y %H:%M')})",
           f"{len(lignes)} fichiers, {tot / MO:.1f} Mo"]
    res += [f"   - {k} : {v}" for k, v in c.most_common()]
    nv = Counter(Path(l[0]).parts[0] if len(Path(l[0]).parts) > 1 else "(racine)" for l in lignes if l[5] == "Nouveau")
    res += ["", "Fichiers nouveaux par sous-dossier :"] + [f"   - {k} : {v}" for k, v in sorted(nv.items())]
    res += ["", f"Fichiers à envoyer : {len(envoi)} ({sum(e[2] for e in envoi) / MO:.1f} Mo)",
            f"Archives produites dans {sortie} : {len(zips)}"]
    res += [f"   - {z.name} ({os.path.getsize(long(z)) / MO:.1f} Mo)" for z in zips]
    res += ["", "À transmettre à Adenka : les archives .zip, ce résumé, l'inventaire et la liste du contenu (.csv)."]
    (sortie / f"FPPN_resume_{DATE}.txt").write_text("\n".join(res), encoding="utf-8-sig")
    print("\n".join(res))


if __name__ == "__main__":
    main()
