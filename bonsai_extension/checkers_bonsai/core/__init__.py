"""Moteur de recette de Checkers pour Bonsai, sans dépendance à Blender.

- recette.py : orchestration (IDS et/ou règles Excel) et verdict ;
- ids.py : contrôle IDS (ifctester) avec garde-fous sur le fichier reçu ;
- rules.py / psets.py : règles Excel et contrôle des PSet ;
- spatial.py : niveau d'un élément (étage ou partie d'ouvrage IFC 4.3) ;
- trace.py : empreintes SHA-256 et en-tête IFC ;
- report.py : rapport Excel (Fiche de recette) ;
- bcf.py : export BCF 2.1 ;
- excel_safety.py : neutralisation des formules injectées.

Aucun module n'effectue d'accès réseau.
"""
