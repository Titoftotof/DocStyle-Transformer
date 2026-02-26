# DocStyle Transformer

Transforme automatiquement vos documents `.docx` en appliquant un design system professionnel (style Apple Minimal). L'utilisateur obtient un document mis en forme sans toucher une seule option de style.

## Fonctionnalites

- **Analyse intelligente** : detection automatique des titres, paragraphes, tableaux, listes, images et encadres
- **Detection de callouts** : identifie les notes, avertissements et conseils par mots-cles ou structure
- **Detection d'etapes** : regroupe les sequences numerotees en blocs de procedures
- **Couverture automatique** : titre split sur deux lignes, barre d'accent, metadonnees
- **Table des matieres** : generee automatiquement avec numerotation des sections
- **Numerotation des sections** : format "Section 01", "Section 02", etc.
- **En-tete et pied de page** : titre du document + pagination automatique
- **Multi-themes** : design system configurable en YAML
- **Interface desktop** : application ttkbootstrap avec previsualisation

## Installation

```bash
pip install -r requirements.txt
```

## Utilisation CLI

```bash
# Transformation basique
python main.py document.docx

# Avec options
python main.py document.docx -o sortie.docx --cover-title "Mon Titre" --mention "Confidentiel" -v

# Sans couverture ni table des matieres
python main.py document.docx --no-cover --no-toc

# Avec un theme personnalise
python main.py document.docx --theme themes/apple-minimal.yaml
```

### Options CLI

| Option | Description |
|---|---|
| `-o, --output` | Chemin du fichier de sortie |
| `--theme` | Chemin vers un fichier theme YAML |
| `--no-cover` | Ne pas generer la page de couverture |
| `--no-toc` | Ne pas generer la table des matieres |
| `--no-numbering` | Ne pas numeroter les sections |
| `--no-header-footer` | Ne pas ajouter en-tete/pied de page |
| `--cover-title` | Titre personnalise pour la couverture |
| `--mention` | Mention en pied de page (defaut: "Confidentiel") |
| `-v, --verbose` | Activer les logs detailles |

## Interface graphique

```bash
python -m ui.app
```

## Architecture

```
docstyle-transformer/
├── main.py                  # Point d'entree CLI
├── config/
│   ├── design-system.yaml   # Theme par defaut
│   └── settings.ini         # Parametres applicatifs
├── core/
│   ├── models.py            # Modeles de donnees (DocumentTree, Section, etc.)
│   ├── parser.py            # Analyse du .docx source
│   ├── detector.py          # Detection callouts, etapes, structure
│   ├── mapper.py            # Mapping vers le design system
│   ├── generator.py         # Generation du .docx final
│   ├── cover.py             # Generateur de couverture
│   └── toc.py               # Generateur de table des matieres
├── themes/
│   └── apple-minimal.yaml   # Theme Apple Minimal
├── ui/
│   ├── app.py               # Application desktop
│   └── components.py        # Widgets reutilisables
├── tests/
│   ├── fixtures/            # Fichiers .docx de test
│   └── test_transform.py    # Suite de tests (40 tests)
└── requirements.txt
```

## Pipeline de transformation

```
.docx source
    |
    v
[Parser] --> DocumentTree (representation intermediaire)
    |
    v
[Detector] --> DocumentTree enrichi (callouts, etapes detectes)
    |
    v
[Generator] --> .docx stylise (avec couverture, TOC, headers)
```

## Tests

```bash
pytest tests/ -v
```

## Technologies

- Python 3.11+
- python-docx + lxml
- PyYAML
- ttkbootstrap (interface)
- pytest
