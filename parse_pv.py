#
# Auteur : sylvain.kieffer@univ-paris13.fr
#
# SPDX-License-Identifier: AGPL-3.0-or-later
# License-Filename: LICENSE
#

import csv
from dataclasses import dataclass


@dataclass
class PV:
    fixed_merge_fields: dict[str, str]
    etudiants: list[dict[str, str]]


def parse_pv(path: str) -> PV:
    with open(path) as f:
        reader = csv.reader(f, delimiter=";")
        rows = list(reader)

    formation = rows[1][0]
    libelle_session = rows[2][0]
    annee = rows[3][0]

    codes = rows[6]
    sessions = rows[7]
    champs = rows[8]

    formation_code = None
    latest_code = None
    latest_libelle = None
    latest_session = None
    ids = []
    matieres = {}
    for i in range(0, len(champs)):
        code = codes[i]
        if code:
            splitted_code = code.split(" - ")
            latest_code = splitted_code.pop(0)
            latest_libelle = " - ".join(splitted_code)
            if formation_code is None:
                formation_code = latest_code
        session = sessions[i]
        if session:
            latest_session = {
                "Session 1": "1",
                "Session 2": "2",
                "Evaluations Finales": "F",
            }.get(session, session)
        champ = champs[i]
        if champ:
            champ_trad = {}.get(champ, champ)
            id = "_".join(
                [x.upper() for x in [latest_code, champ_trad, latest_session] if x]
            )
            ids.append(id)
            if latest_code:
                matieres[f"{latest_code}_LIBELLE"] = latest_libelle

    etudiants = []
    for row in rows[9:]:
        etudiant = {}
        for i, id in enumerate(ids):
            etudiant[id] = row[i]
        etudiants.append(etudiant)

    fixed_merge_fields = {
        "FORMATION_CODE": formation_code,
        "FORMATION_LIBELLE": formation,
        "SESSION_LIBELLE": libelle_session,
        "ANNEE_LIBELLE": annee,
    } | matieres

    return PV(
        fixed_merge_fields=fixed_merge_fields,
        etudiants=etudiants,
    )


if __name__ == "__main__":
    pv = parse_pv("pv-de-jury.csv")
    for k in pv.fixed_merge_fields.keys():
        print(f"#{k}#: FIXE")
    for k in pv.etudiants[0].keys():
        print(f"#{k}#: ETUDIANT")
    # print(pv)
