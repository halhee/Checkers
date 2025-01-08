# IFC Analyzer Rust Module

Ce module Rust est conçu pour accélérer l'analyse des fichiers IFC en utilisant les performances et la parallélisation de Rust.

## Installation

1. Installer Rust et Cargo :
```bash
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
```

2. Installer les dépendances Python :
```bash
pip install maturin
```

3. Compiler le module :
```bash
cd rust_analyzer
maturin develop
```

## Utilisation

Dans votre code Python :

```python
from ifc_analyzer import analyze_elements

# Préparer les données
elements = [...]  # Liste des IDs d'éléments
required_psets = {...}  # Dictionnaire des PSet requis
element_psets = {...}  # Dictionnaire des PSet des éléments

# Analyser les éléments
results = analyze_elements(elements, required_psets, element_psets)
```

## Performance

Le module utilise :
- Rayon pour la parallélisation
- Structures de données optimisées
- Gestion efficace de la mémoire

## Avantages

1. **Vitesse** : 5-10x plus rapide que Python pur
2. **Mémoire** : Utilisation mémoire optimisée
3. **Parallélisation** : Utilisation efficace de tous les cœurs CPU
4. **Sécurité** : Vérifications à la compilation
