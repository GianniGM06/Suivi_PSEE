# Suivi PSEE — Amundi Immobilier

Scripts de génération et d'automatisation pour le suivi des Prestations de Services Essentielles et Externalisées (PSEE / PCI).

## Fichiers à utiliser

| Fichier | Description |
|---|---|
| `generer_matrice_pci_v3.py` | ✅ **Générateur principal** — 5 onglets : MODE D'EMPLOI, DATABASE, REMÉDIATION, CONTRÔLES LOD1, SYNTHÈSE |
| `MailsLOD1.bas` | ✅ **Module VBA Outlook** — campagnes L1-08 (événements majeurs) et L1-09 (tests PCA) |
| `generer_matrice_pci_v2.py` | ❌ Obsolète |
| `generer_matrice_pci.py` | ❌ Obsolète |
| `build_matrice.py` | ❌ Obsolète |
| `relances_lod1_outlook.py` | ❌ Obsolète — remplacé par MailsLOD1.bas |

## Utilisation

```bash
pip install openpyxl
python generer_matrice_pci_v3.py
```

Génère `Matrice_PCI_Remediation_v3.xlsx` dans le répertoire courant.

## Activation des macros VBA

### ⚠️ Copier-coller dans le VBE — supprimer la première ligne

1. Ouvrir le `.xlsx` dans Excel
2. `Alt+F11` → Insertion → Module
3. Copier-coller le contenu de `MailsLOD1.bas`
4. **Supprimer la toute première ligne** : `Attribute VB_Name = "MailsLOD1"` (provoque une erreur de syntaxe si laissée)
5. Adapter `EXPEDITEUR_NOM` et `EXPEDITEUR_EMAIL` en haut du module
6. Fermer VBE → Fichier → Enregistrer sous → **Classeur Excel avec macros (.xlsm)**

### Alternative : import direct du fichier .bas (pas besoin de supprimer la ligne)

1. `Alt+F11` → Fichier → **Importer un fichier...**
2. Sélectionner `MailsLOD1.bas` → OK
3. La ligne `Attribute` est gérée automatiquement par Excel

## Insérer les boutons dans CONTRÔLES LOD1

1. Insertion → Formes → Rectangle arrondi → dessiner dans la zone bandeau
2. Double-clic sur la forme → taper le label (`📧 LANCER L1-08`, etc.)
3. Clic droit → **Affecter une macro** → choisir :
   - `LancerCampagneL108`
   - `RelancesL108`
   - `LancerCampagneL109`
   - `RelancesL109`
4. Mise en forme : fond `#001C4B`, texte blanc, Arial 9

## Structure de la matrice v3

| Onglet | Contenu |
|---|---|
| MODE D'EMPLOI | Procédures L1-08, L1-09, activation VBA |
| DATABASE | 28 prestataires — référentiel central |
| REMÉDIATION | DO · QECI · Avis Risques · Stratégie Sortie · EASY · Contrat · Remédiation EBA |
| CONTRÔLES LOD1 | L1-08 semestriel (PCI HG) + L1-09 annuel T1 (toutes PCI) |
| SYNTHÈSE | Tableau de bord RAG consolidé |
