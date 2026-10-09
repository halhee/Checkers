"""Persistance légère (SQLite) de l'historique des analyses Checkers.

Deux usages, au cœur du métier de référent fonctionnel BIM en phase
d'exploitation d'un SI :
- suivi dans le temps (une analyse = un point d'une courbe : taux de
  validité, empreinte carbone, nombre d'anomalies ouvertes) ;
- comparaison entre deux analyses (ex. avant/après correction d'une
  maquette par un corps d'état) pour voir quelles anomalies ont été
  corrigées, lesquelles sont nouvelles, lesquelles persistent -- exactement
  le suivi d'anomalies attendu d'une recette fonctionnelle.

Une "anomalie" ici est une non-conformité agrégée par (lot, étage, classe
IFC, PSet) -- pas par élément individuel. Ce choix de granularité reprend
une leçon du prototype BIM Tender Check (~/Desktop/mon-projet) : à l'échelle
réelle (des centaines d'éléments concernés), énumérer chaque élément rend
le suivi illisible ; une règle/PSet manquant sur N éléments d'un type/étage
donné est l'unité utile pour une fiche de recette. La clé stable (hash du
tuple) permet de faire correspondre la "même" anomalie entre deux analyses
même si les identifiants internes IFC des éléments ont changé.
"""

import sqlite3
import json
import os
import hashlib
from datetime import datetime, timezone

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'checkers_history.db')


def _connect():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db():
    conn = _connect()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS analyses (
            id TEXT PRIMARY KEY,
            timestamp TEXT NOT NULL,
            label TEXT NOT NULL,
            excel_filename TEXT,
            lots_json TEXT NOT NULL,
            total_elements INTEGER,
            valid_elements INTEGER,
            validation_rate REAL,
            carbon_total REAL,
            anomaly_count INTEGER
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS anomalies (
            analysis_id TEXT NOT NULL,
            anomaly_key TEXT NOT NULL,
            lot_name TEXT,
            floor TEXT,
            ifc_class TEXT,
            pset_name TEXT,
            count INTEGER
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_anomalies_analysis ON anomalies(analysis_id)")
    conn.commit()
    conn.close()


def _anomaly_key(lot_name, floor, ifc_class, pset_name) -> str:
    raw = f"{lot_name}|{floor}|{ifc_class}|{pset_name}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def record_analysis(analysis_id: str, label: str, excel_filename: str, results: dict, anomalies_raw: list):
    """`anomalies_raw` = liste de dicts {lot, floor, ifc_class, pset, count}
    (voir Checkers.py::process_files, dérivé de missing_combo_counts)."""
    total = results["total_elements"]
    valid = results["valid_elements"]
    validation_rate = (valid / total * 100) if total else None

    conn = _connect()
    conn.execute(
        "INSERT OR REPLACE INTO analyses "
        "(id, timestamp, label, excel_filename, lots_json, total_elements, valid_elements, "
        "validation_rate, carbon_total, anomaly_count) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (
            analysis_id,
            datetime.now(timezone.utc).isoformat(),
            label,
            excel_filename,
            json.dumps([{"name": l["name"], "corps_etat": l["corps_etat"]} for l in results["lots"]]),
            total,
            valid,
            validation_rate,
            results["carbon_footprint"]["total"],
            len(anomalies_raw),
        ),
    )
    conn.execute("DELETE FROM anomalies WHERE analysis_id = ?", (analysis_id,))
    conn.executemany(
        "INSERT INTO anomalies (analysis_id, anomaly_key, lot_name, floor, ifc_class, pset_name, count) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)",
        [
            (
                analysis_id,
                _anomaly_key(a["lot"], a["floor"], a["ifc_class"], a["pset"]),
                a["lot"], a["floor"], a["ifc_class"], a["pset"], a["count"],
            )
            for a in anomalies_raw
        ],
    )
    conn.commit()
    conn.close()


def list_analyses(limit: int = 100) -> list:
    conn = _connect()
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        "SELECT * FROM analyses ORDER BY timestamp DESC LIMIT ?", (limit,)
    ).fetchall()
    conn.close()
    result = []
    for row in rows:
        d = dict(row)
        d["lots"] = json.loads(d.pop("lots_json"))
        result.append(d)
    return result


def get_analysis(analysis_id: str):
    conn = _connect()
    conn.row_factory = sqlite3.Row
    row = conn.execute("SELECT * FROM analyses WHERE id = ?", (analysis_id,)).fetchone()
    conn.close()
    if row is None:
        return None
    d = dict(row)
    d["lots"] = json.loads(d.pop("lots_json"))
    return d


def get_anomalies(analysis_id: str) -> list:
    conn = _connect()
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        "SELECT * FROM anomalies WHERE analysis_id = ?", (analysis_id,)
    ).fetchall()
    conn.close()
    return [dict(row) for row in rows]


def compare_analyses(before_id: str, after_id: str) -> dict:
    """Diff des anomalies entre deux analyses, par clé stable (lot, étage,
    classe IFC, PSet). Retourne :
    - resolved : anomalies présentes dans `before`, absentes de `after`
      (corrigées) ;
    - new : anomalies absentes de `before`, présentes dans `after`
      (nouvelles -- ex. régression, ou nouveau lot analysé) ;
    - persistent : présentes dans les deux (avec le compte des deux côtés,
      pour voir une amélioration/dégradation partielle du nombre d'éléments
      concernés sans que l'anomalie ait totalement disparu)."""
    before = {a["anomaly_key"]: a for a in get_anomalies(before_id)}
    after = {a["anomaly_key"]: a for a in get_anomalies(after_id)}

    resolved = [a for key, a in before.items() if key not in after]
    new = [a for key, a in after.items() if key not in before]
    persistent = [
        {**a, "count_before": before[key]["count"]}
        for key, a in after.items() if key in before
    ]

    return {"resolved": resolved, "new": new, "persistent": persistent}
