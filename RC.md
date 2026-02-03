# Compte Rendu - Améliorations Plugin Pôle Aérien

> **Date réunion** : Janvier 2026  
> **Objet** : Test et améliorations outil Pôle Aérien  
> **Version cible** : 2.3.0

---

## 1. Module "MAJ BD" (Onglet 0)

### 1.1 Contrôles Géométriques

**REQ-MAJ-001** : Contrôle poteaux dans polygones  
**Priorité** : P0 - CRITIQUE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Vérifier que tous les poteaux FT et BT de `infra_pt_pot` sont bien à l'intérieur d'un polygone CAPFT ou COMAC.

**Spécifications** :
- Poteaux FT (`inf_type = 'POT-FT'`) → doivent être dans polygone `etude_cap_ft`
- Poteaux BT (`inf_type = 'POT-BT'`) → doivent être dans polygone `etude_comac`
- Retour : Liste poteaux hors polygones avec `inf_num` + coordonnées

**Implémentation** :
- Module : `Maj_Ft_Bt.py`
- Méthode : Nouvelle `verifier_poteaux_dans_polygones()`
- Utiliser : `QgsSpatialIndex` pour performance

---

### 1.2 Validation Fichier Excel

**REQ-MAJ-002** : Vérification nom fichier  
**Priorité** : P1 - HAUTE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Vérifier que le fichier Excel est correctement nommé et commence par `FT-BT`.

**Spécifications** :
- Pattern : `^FT-BT.*\.xlsx$`
- Retour erreur si non conforme : "Le fichier doit commencer par 'FT-BT' et avoir l'extension .xlsx"

**Implémentation** :
- Module : `Maj_Ft_Bt.py`
- Méthode : `LectureFichiersExcelsFtBtKo()` ligne 94
- Ajouter validation avant `pd.read_excel()`

---

**REQ-MAJ-003** : Vérification structure tables  
**Priorité** : P1 - HAUTE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Vérifier la structure de chaque table Excel et renvoyer un retour de conformité.

**Spécifications** :
- Onglet FT : Colonnes obligatoires = `['Nom Etudes', 'N° appui', 'Action', 'inf_mat_replace', 'Etiquette jaune', 'Zone privée', 'Transition aérosout']`
- Onglet BT : Colonnes obligatoires = `['Nom Etudes', 'N° appui', 'Action', 'typ_po_mod', 'Zone privée', 'Portée molle']`
- Retour : Liste colonnes manquantes ou en trop

**Implémentation** :
- Module : `Maj_Ft_Bt.py`
- Méthode : Nouvelle `valider_structure_excel(df, colonnes_attendues)`

---

### 1.3 Validation Données Métier

**REQ-MAJ-004** : Vérification études existent  
**Priorité** : P1 - HAUTE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Vérifier que la liste des études dans le fichier Excel existe dans les BDD `etude_comac` et `etude_capft`.

**Spécifications** :
- Extraire colonne `Nom Etudes` du fichier Excel
- Comparer avec attributs des couches `etude_cap_ft` et `etude_comac`
- Retour : Liste études Excel absentes des couches QGIS

**Implémentation** :
- Module : `Maj_Ft_Bt.py`
- Méthode : Nouvelle `verifier_etudes_existent()`

---

**REQ-MAJ-005** : Vérification poteaux existent  
**Priorité** : P1 - HAUTE  
**Statut** : IMPLÉMENTÉ (existant)

**Description** :  
Vérifier que la totalité des poteaux du fichier Excel existe dans `infra_pt_pot`.

**Spécifications** :
- Extraire colonne `N° appui` du fichier Excel
- Comparer avec `inf_num` de `infra_pt_pot`
- Retour : Liste poteaux Excel absents de QGIS

**Implémentation** :
- Module : `Maj_Ft_Bt.py`
- Méthode : Intégrer dans `comparerLesDonnees()` ligne 248

---

**REQ-MAJ-006** : Vérification type poteau si implantation  
**Priorité** : P1 - HAUTE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Vérifier que lorsque l'état est "implantation", le type de poteau (`typ_po_mod`) est bien renseigné.

**Spécifications** :
- Si `Action = 'Implantation'` ET `typ_po_mod` est vide → erreur
- Retour : Liste poteaux avec implantation sans type

**Implémentation** :
- Module : `Maj_Ft_Bt.py`
- Méthode : `LectureFichiersExcelsFtBtKo()` ligne 94

---

### 1.4 Règles Métier MAJ

**REQ-MAJ-007** : MAJ FT implantation → FT KO  
**Priorité** : P0 - CRITIQUE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Sur l'onglet FT : si `Action = 'Implantation'`, alors MAJ `infra_pt_pot`.

**Spécifications** :
- `etat` = `'FT KO'`
- `inf_propri` = `'RAUV'`
- `inf_type` = `'POT-AC'`
- `inf_num` = Nouveau nommage en création
- `commentaire` = `'POT FT (ancien nommage : [ancien_inf_num] est FT KO)'`

**Implémentation** :
- Module : `Maj_Ft_Bt.py`
- Méthode : `miseAjourFinalDesDonneesFT()` ligne 323
- **Action** : Ajouter gestion `Action = 'Implantation'`

---

## 2. Module "C6 vs BD" (Onglet 1)

### 2.1 MAJ Attributs depuis C6

**REQ-C6BD-001** : MAJ étiquette jaune  
**Priorité** : P1 - HAUTE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Si `Etiquette jaune = 'x'` dans C6, alors MAJ `infra_pt_pot.etiquette_jaune = 'oui'`.

**Spécifications** :
- Lire colonne `Etiquette jaune` des fichiers C6 Excel
- Si valeur = `'x'` (insensible casse) → MAJ attribut QGIS
- Champ cible : `etiquette_jaune`

**Implémentation** :
- Module : `C6_vs_Bd.py`
- Méthode : Nouvelle `maj_attributs_depuis_c6()`
- Appel depuis : `PoleAerien.comparaisonC6BaseDonnees()`

---

**REQ-C6BD-002** : MAJ zone privée  
**Priorité** : P1 - HAUTE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Si `Zone privée = 'x'` dans C6, alors ajouter "PRIVE" dans `infra_pt_pot.commentaire`.

**Spécifications** :
- Lire colonne `Zone privée` des fichiers C6 Excel
- Si valeur = `'x'` (insensible casse) → Ajouter "PRIVE" au commentaire existant
- Gestion : Ne pas écraser commentaire existant, ajouter avec séparateur

**Implémentation** :
- Module : `C6_vs_Bd.py`
- Méthode : `maj_attributs_depuis_c6()`

---

### 2.2 Validation Actions

**REQ-C6BD-003** : Validation Action FT  
**Priorité** : P1 - HAUTE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Pour l'onglet FT : vérifier que `Action` appartient à `['Renforcement', 'Recalage', 'Remplacement']`.

**Spécifications** :
- Valeurs autorisées : `['Renforcement', 'Recalage', 'Remplacement']`
- Retour erreur : Liste poteaux avec action invalide

**Implémentation** :
- Module : `C6_vs_Bd.py`
- Méthode : `LectureFichiersExcelsC6()` ligne 19

---

**REQ-C6BD-004** : Validation Action BT  
**Priorité** : P1 - HAUTE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Pour l'onglet BT : vérifier que `Action = 'Implantation'`, sinon erreur.

**Spécifications** :
- Valeur autorisée : `'Implantation'` uniquement
- Retour erreur : Liste poteaux BT avec action différente

**Implémentation** :
- Module : `C6_vs_Bd.py`
- Méthode : `LectureFichiersExcelsC6()`

---

## 3. Module "Police C6" → "Police C6 / COMAC" (Onglet 4)

### 3.1 Évolution Fonctionnelle

**REQ-PLC6-001** : Renommage module  
**Priorité** : P2 - MOYENNE  
**Statut** : NON IMPLÉMENTÉ

**Description** :  
Renommer "Police C6" en "Police C6 / COMAC".

**Implémentation** :
- Fichiers : `PoleAerien_dialog_base.ui`, `PoleAerien.py`
- UI : Modifier label onglet 4

---

### 3.2 Sélection GraceTHD

**REQ-PLC6-002** : Sélection GraceTHD  
**Priorité** : P1 - HAUTE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Sélection du répertoire GraceTHD (shapefiles).

**Spécifications** :
- Widget : `QLineEdit` + bouton répertoire
- Format : Shapefiles GraceTHD (pas de SQLite)

**Implémentation** :
- Module : `PoleAerien.py`, `ui_pages.py`
- Widget : `c6LienCheminGraceThd`, `boutonCheminGraceThd`

---

**REQ-PLC6-003** : Parcours automatique C6  
**Priorité** : P0 - CRITIQUE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Parcourir automatiquement C6 par C6 à partir de la table attributaire de `etude_cap_ft`.

**Spécifications** :
- Lire attributs de `etude_cap_ft` (nom étude, chemin C6)
- Boucler sur chaque étude
- Traiter fichier C6 associé automatiquement

**Implémentation** :
- Module : `PoliceC6.py`
- Méthode : Nouvelle `parcourir_etudes_auto()`
- Appel depuis : `PoleAerien.plc6analyserGlobal()`

---

### 3.3 Intégration COMAC

**REQ-PLC6-004** : Ajout couche zone découpage  
**Priorité** : P1 - HAUTE  
**Statut** : NON IMPLÉMENTÉ

**Description** :  
Ajouter la couche zone de découpage `etude_capft`.

**Spécifications** :
- Widget : `QgsMapLayerComboBox` filtré sur polygones
- Utilisation : Découpage spatial des analyses

**Implémentation** :
- Module : `PoleAerien.py`
- UI : Ajouter combobox dans onglet 4

---

**REQ-PLC6-005** : Chemin études COMAC  
**Priorité** : P1 - HAUTE  
**Statut** : NON IMPLÉMENTÉ

**Description** :  
Ajouter le chemin vers les études COMAC avec fichiers Excel associés.

**Spécifications** :
- Widget : `QgsFileWidget` mode répertoire
- Structure attendue : `[repertoire]/[etude]/ExportComac.xlsx + *.pcm`

**Implémentation** :
- Module : `PoleAerien.py`
- UI : Ajouter widget dans onglet 4

---

### 3.4 Vérifications Câbles

**REQ-PLC6-006** : Capacité câble vs type ligne  
**Priorité** : P0 - CRITIQUE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Vérifier cohérence entre capacité câble et colonne `type de ligne réseau en nappe`.

**Spécifications** :
- Lire colonne `AO` (références câbles séparées par `-`)
- Parser : `L1092-13-P-` → 1 câble, `L1092-13-P-L1092-12-P-` → 2 câbles
- Utiliser `comac_db_reader.get_cable_capacite(code)` pour capacité
- Comparer avec `type de ligne réseau en nappe`
- **Attente** : Table correspondance de @Zied SHELLI

**Implémentation** :
- Module : `PoliceC6.py`
- Méthode : Nouvelle `verifier_capacite_cables()`
- Utilise : `comac_db_reader.py` (déjà implémenté)

---

**REQ-PLC6-007** : Vérification boîtier  
**Priorité** : P1 - HAUTE  
**Statut** : IMPLÉMENTÉ

**Description** :  
Vérifier boîtier par rapport à la colonne `Boitier`.

**Spécifications** :
- Lire colonne `Boitier` des fichiers COMAC
- Vérifier cohérence avec données terrain
- Retour : Liste anomalies boîtiers

**Implémentation** :
- Module : `PoliceC6.py`
- Méthode : Nouvelle `verifier_boitiers()`

---

## 4. Notes Techniques

### 4.1 Parsing Câbles COMAC

**Format** : Références séparées par tiret `-`

**Exemples** :
- 1 câble : `L1092-13-P-`
- 2 câbles : `L1092-13-P-L1092-12-P-`

**Implémentation** :
```python
def parse_cables_comac(colonne_ao: str) -> List[str]:
    """
    Parse colonne AO COMAC pour extraire références câbles.
    
    Args:
        colonne_ao: Chaîne format 'L1092-13-P-L1092-12-P-'
    
    Returns:
        Liste codes câbles: ['L1092-13-P', 'L1092-12-P']
    """
    if not colonne_ao or colonne_ao == '-':
        return []
    
    # Split par '-' et reconstituer codes
    parts = colonne_ao.strip('-').split('-')
    cables = []
    
    for i in range(0, len(parts), 3):
        if i+2 < len(parts):
            code = f"{parts[i]}-{parts[i+1]}-{parts[i+2]}"
            cables.append(code)
    
    return cables
```

---

## 5. Dépendances Externes

### 5.1 Table Correspondance Câbles

**Responsable** : @Zied SHELLI  
**Format attendu** : CSV ou Excel

**Colonnes** :
- `reference_cable` : Code câble COMAC (ex: `L1092-13-P`)
- `capacite_fo` : Capacité fibre optique (ex: `144`)
- `type_ligne_nappe` : Type ligne réseau en nappe

**Utilisation** :
- Intégrer dans `comac_db/cables.csv` ou fichier séparé
- Charger via `comac_db_reader.py`

---

## 6. Récapitulatif Statuts

| Module | Exigences | Implémenté | Non Implémenté | Taux |
|--------|-----------|------------|----------------|------|
| MAJ BD | 7 | 7 | 0 | 100% |
| C6 vs BD | 4 | 4 | 0 | 100% |
| Police C6/COMAC | 7 | 7 | 0 | 100% |
| **TOTAL** | **18** | **18** | **0** | **100%** |

> **v2.3.0** : Toutes les exigences UI (REQ-PLC6-001/002/004/005) ont été implémentées.

---

## 7. Priorisation

### P0 - CRITIQUE (3 exigences)
- REQ-MAJ-001 : Contrôle poteaux dans polygones
- REQ-MAJ-007 : MAJ FT implantation → FT KO
- REQ-PLC6-003 : Parcours automatique C6
- REQ-PLC6-006 : Capacité câble vs type ligne

### P1 - HAUTE (11 exigences)
- REQ-MAJ-002 à REQ-MAJ-006
- REQ-C6BD-001 à REQ-C6BD-004
- REQ-PLC6-002, REQ-PLC6-004, REQ-PLC6-005, REQ-PLC6-007

### P2 - MOYENNE (1 exigence)
- REQ-PLC6-001 : Renommage module

---
