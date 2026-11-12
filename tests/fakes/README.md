Doublures de test minimales (IfcOpenShell, ifctester, bpy, Bonsai) pour
exécuter les tests sans Blender ni IfcOpenShell. Elles n'implémentent que
ce que l'extension utilise, avec les mêmes signatures que les bibliothèques
réelles (vérifiées dans le code source d'IfcOpenShell 0.9.1) ; la validation
finale se fait dans Bonsai, sur de vraies maquettes.
