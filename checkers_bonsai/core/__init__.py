"""Moteur de recette de Checkers pour Bonsai, sans dépendance à Blender.

- recette.py : orchestration des contrôles et verdict ;
- structure.py : contrôles de structure du modèle, sans référentiel ;
- ids.py : contrôle IDS (ifctester) avec garde-fous sur le fichier reçu ;
- rules.py / psets.py : règles Excel et contrôle des PSet ;
- journal.py : journal JSON d'une recette et comparaison entre livraisons ;
- inventory.py : inventaire d'exploitation (équipements et locaux) ;
- spatial.py : niveau d'un élément (étage ou partie d'ouvrage IFC 4.3) ;
- trace.py : empreintes SHA-256 et en-tête IFC ;
- report.py : rapport Excel (fiche de recette) ;
- bcf.py : export BCF 2.1 ;
- excel_safety.py : neutralisation des formules injectées.

Aucun module n'effectue d'accès réseau.
"""
