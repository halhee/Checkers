"""Moteur de recette de Checkers pour Bonsai, sans dépendance à Blender.

- recette.py : orchestration des contrôles et verdict ;
- structure.py : contrôles de structure du modèle, sans référentiel ;
- ids.py : contrôle IDS (ifctester) avec garde-fous sur le fichier reçu ;
- rules.py / psets.py : règles Excel et contrôle des PSet ;
- federation.py : cohérence de référence avec les maquettes des autres lots ;
- clashes.py : conflits géométriques avec IfcClash, selon des règles ;
- derogations.py : fichier des dérogations (anomalies admises, avec leur motif) ;
- scene.py : coordonnées IFC vers la scène Blender, géométrie d'un élément ;
- survey.py : fiche de relevé terrain et comparaison avec la maquette ;
- journal.py : journal JSON d'une recette et comparaison entre livraisons ;
- tracking.py : tableau de suivi consolidé à partir des journaux ;
- inventory.py : inventaire d'exploitation (équipements et locaux) ;
- diagnostic.py : vérification de l'installation sur le poste ;
- spatial.py : niveau d'un élément (étage ou partie d'ouvrage IFC 4.3) ;
- trace.py : empreintes SHA-256 et en-tête IFC ;
- report.py : rapport Excel (fiche de recette) ;
- bcf.py : export BCF 2.1 ;
- excel_safety.py : neutralisation des formules injectées et des caractères de contrôle.

Aucun module n'effectue d'accès réseau.
"""
