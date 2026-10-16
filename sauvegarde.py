"""Sauvegarde complète de Frais Agora : table `frais_depenses` + tickets.

Deux usages :
  1. Depuis l'application : bouton « Sauvegarde » réservé à l'administrateur.
  2. En ligne de commande / GitHub Actions :
         python sauvegarde.py ping     # requête légère : évite la mise en pause
         python sauvegarde.py backup   # crée sauvegardes/sauvegarde_....zip

En ligne de commande, il faut définir les variables d'environnement
SUPABASE_URL et SUPABASE_KEY (clé secrète / service_role).
Le programme se termine avec un code d'erreur en cas de problème, ce qui
déclenche un e-mail d'alerte de GitHub.
"""
from __future__ import annotations

import csv
import datetime
import io
import json
import os
import sys
import zipfile

TABLE = "frais_depenses"
BUCKET = "justificatifs"
PAGE = 1000  # Supabase renvoie au maximum 1000 lignes par requête
SANS_TICKET = ("", "Aucun justificatif", "None", "nan")


def lire_table(sb) -> list:
    """Lit TOUTES les lignes de la table, par pages de 1000."""
    lignes, debut = [], 0
    while True:
        res = (
            sb.table(TABLE)
            .select("*")
            .order("id")
            .range(debut, debut + PAGE - 1)
            .execute()
        )
        lot = res.data or []
        lignes.extend(lot)
        if len(lot) < PAGE:
            break
        debut += PAGE
    return lignes


def lister_bucket(sb):
    """Liste les fichiers du bucket. Renvoie None si la liste est impossible."""
    noms, offset = [], 0
    try:
        while True:
            lot = sb.storage.from_(BUCKET).list(
                "", {"limit": PAGE, "offset": offset}
            )
            lot = lot or []
            # les dossiers n'ont pas d'identifiant : on les ignore
            noms.extend(f["name"] for f in lot if f.get("id") is not None)
            if len(lot) < PAGE:
                break
            offset += PAGE
        return noms
    except Exception:
        return None


def _lisez_moi(resume: dict) -> str:
    manquants = resume["manquants"]
    texte = [
        "SAUVEGARDE FRAIS AGORA",
        f"Créée le : {resume['date']}",
        f"Lignes de dépenses : {resume['nb_lignes']}",
        f"Tickets sauvegardés : {resume['nb_tickets']}",
        "",
        "Contenu :",
        "  - frais_depenses.csv / .json : toutes les dépenses (tous les utilisateurs)",
        "  - justificatifs/             : les photos des tickets",
        "",
        "Restauration :",
        "  - Table : Supabase > Table Editor > frais_depenses > Insert >",
        "    Import data from CSV (utiliser frais_depenses.csv).",
        "  - Tickets : Supabase > Storage > justificatifs > Upload files.",
    ]
    if manquants:
        texte += ["", "ATTENTION - tickets introuvables dans le stockage :"]
        texte += [f"  - {nom}" for nom in manquants]
    return "\n".join(texte) + "\n"


def construire_sauvegarde(sb):
    """Construit le ZIP de sauvegarde. Renvoie (octets_du_zip, resume)."""
    lignes = lire_table(sb)

    references = set()
    for ligne in lignes:
        nom = str(ligne.get("justificatif") or "").strip()
        if nom not in SANS_TICKET:
            references.add(nom)

    dans_bucket = lister_bucket(sb)
    if dans_bucket is None:
        a_telecharger = sorted(references)
    else:
        a_telecharger = sorted(set(dans_bucket) | references)

    colonnes = []
    for ligne in lignes:
        for cle in ligne:
            if cle not in colonnes:
                colonnes.append(cle)

    manquants, nb_tickets = [], 0
    memoire = io.BytesIO()
    with zipfile.ZipFile(memoire, "w", zipfile.ZIP_DEFLATED) as zf:
        tampon = io.StringIO()
        ecrivain = csv.DictWriter(tampon, fieldnames=colonnes, extrasaction="ignore")
        ecrivain.writeheader()
        ecrivain.writerows(lignes)
        zf.writestr("frais_depenses.csv", tampon.getvalue())
        zf.writestr(
            "frais_depenses.json",
            json.dumps(lignes, ensure_ascii=False, indent=2, default=str),
        )

        for nom in a_telecharger:
            try:
                contenu = sb.storage.from_(BUCKET).download(nom)
                zf.writestr(f"justificatifs/{nom}", bytes(contenu))
                nb_tickets += 1
            except Exception:
                manquants.append(nom)

        resume = {
            "date": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
            "nb_lignes": len(lignes),
            "nb_tickets": nb_tickets,
            "manquants": manquants,
        }
        zf.writestr("LISEZMOI.txt", _lisez_moi(resume))

    return memoire.getvalue(), resume


def nom_fichier_sauvegarde() -> str:
    return (
        "sauvegarde_frais_agora_"
        + datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
        + ".zip"
    )


def main(argv: list) -> int:
    mode = argv[1] if len(argv) > 1 else "ping"
    from supabase import create_client

    sb = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])

    if mode == "ping":
        res = sb.table(TABLE).select("id").limit(1).execute()
        print(f"OK - base joignable ({len(res.data or [])} ligne lue).")
        return 0

    if mode == "backup":
        octets, resume = construire_sauvegarde(sb)
        os.makedirs("sauvegardes", exist_ok=True)
        chemin = os.path.join("sauvegardes", nom_fichier_sauvegarde())
        with open(chemin, "wb") as f:
            f.write(octets)
        print(
            f"Sauvegarde créée : {chemin} - {resume['nb_lignes']} lignes,"
            f" {resume['nb_tickets']} tickets."
        )
        if resume["nb_lignes"] == 0:
            print("ERREUR : la table est vide ou illisible (clé incorrecte ?).")
            return 1
        if resume["manquants"]:
            print("ERREUR : tickets introuvables :", resume["manquants"])
            return 1
        return 0

    print(f"Mode inconnu : {mode} (utiliser ping ou backup)")
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
