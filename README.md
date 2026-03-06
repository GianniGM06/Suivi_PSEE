# Suivi PSEE — Amundi Immobilier

Scripts de génération et d'automatisation pour le suivi des Prestations de Services Essentielles et Externalisées (PSEE / PCI).

## Fichiers

| Fichier | Description |
|---|---|
| `generer_matrice_pci_v3.py` | Générateur principal de la matrice Excel v3 (5 onglets : MODE D'EMPLOI, DATABASE, REMÉDIATION, CONTRÔLES LOD1, SYNTHÈSE) |
| `MailsLOD1.bas` | Module VBA Outlook — campagnes L1-08 (événements majeurs) et L1-09 (tests PCA) |
| `generer_matrice_pci_v2.py` | Version précédente (avec onglets EASY et PS) |
| `generer_matrice_pci.py` | Version initiale |
| `build_matrice.py` | Script utilitaire de construction |
| `relances_lod1_outlook.py` | Script Python alternatif pour les relances Outlook |

## Utilisation

```bash
pip install openpyxl
python generer_matrice_pci_v3.py
```

Génère `Matrice_PCI_Remediation_v3.xlsx` dans le répertoire courant.

## Activation des macros VBA

1. Ouvrir le `.xlsx` dans Excel
2. `Alt+F11` → Insertion → Module → coller `MailsLOD1.bas`
3. Adapter `EXPEDITEUR_NOM` et `EXPEDITEUR_EMAIL` en haut du module
4. Fichier → Enregistrer sous → **Classeur Excel avec macros (.xlsm)**

## Structure de la matrice v3

| Onglet | Contenu |
|---|---|
| MODE D'EMPLOI | Procédures L1-08, L1-09, activation VBA |
| DATABASE | 28 prestataires — référentiel central |
| REMÉDIATION | DO · QECI · Avis Risques · Stratégie Sortie · EASY · Contrat · EBA |
| CONTRÔLES LOD1 | L1-08 semestriel (PCI HG) + L1-09 annuel T1 (toutes PCI) |
| SYNTHÈSE | Tableau de bord RAG consolidé |
