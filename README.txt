================================================================================
                              POLE AERIEN - Plugin QGIS
              Controle qualite et MAJ des poteaux aeriens FTTH (FT/BT)
================================================================================

DESCRIPTION
-----------
Plugin QGIS pour le controle qualite, la comparaison et la mise a jour des
donnees de poteaux aeriens dans le cadre des projets FTTH. Compare les donnees
terrain (Excel, CSV, PCM) avec la BDD PostgreSQL et les fichiers GraceTHD.
Produit un rapport Excel unifie multi-modules avec tableau de bord, glossaire
et dictionnaire de donnees.

Supporte deux bureaux d'etudes :
- NGE : donnees conformes en BDD PostgreSQL (fddcpiax, infra_pt_pot)
- Axione : donnees via GraceTHD (fichiers locaux SHP+CSV)

COMPATIBILITE
-------------
- QGIS : 3.28 a 4.99 (Qt5/PyQt5 et Qt6/PyQt6)
- Python : 3.9+
- OS : Windows, Linux, macOS

FONCTIONNALITES
---------------

0. MAJ BD
   - Import fichier Excel FT-BT KO (onglets FT et BT)
   - Extraction spatiale des poteaux par etude (CAP_FT et COMAC)
   - MAJ attributs : etat, action, inf_mat, inf_type, inf_propri, noe_usage
   - Execution SQL directe en background thread (PostgreSQL)

1. Verification CAP_FT
   - Comparaison poteaux FT QGIS vs fiches appuis Excel (FicheAppui_*.xlsx)
   - Detection doublons, hors perimetre, absents QGIS/fiches

2. GESPOT vs C6
   - Comparaison champ par champ CSV GESPOT vs Excel Annexe C6
   - 11 criteres compares par appui (type, strategie, milieu, electrique, etc.)
   - Regles de normalisation metier (BR-01 a BR-07)

3. Verification COMAC
   - Comparaison poteaux BT QGIS vs fichiers ExportComac.xlsx
   - Verification cables COMAC vs BDD (fddcpiax) ou GraceTHD
   - Verification boitiers FO vs BPE
   - Comparaison portees PCM vs BDD/GraceTHD
   - Schemas techniques des appuis (dessins polaires depuis fichiers PCM)

4. C6 vs BD
   - Comparaison annexes C6 vs poteaux FT en BDD QGIS
   - Detection ecarts IN/OUT par zone d'etude

5. Police C6
   - Analyse detaillee des annexes C6 vs BDD (cables, boitiers, appuis)
   - Filtrage cables gras uniquement (bold font)
   - Support attaches-aware (cable/BPE detection via rip_avg_nge.attaches)
   - Compteur par etude avec bilan anomalies

6. C6-C3A-C7 vs BD
   - Croisement annexes C6, C3A et C7 avec la base QGIS
   - Support QGIS ou Excel pour C3A

RAPPORT UNIFIE
--------------
Tous les modules alimentent un rapport Excel unique :
- TABLEAU DE BORD : KPIs par module, synthese visuelle
- Feuilles par module : formatage professionnel, filtres, freeze panes
- COMAC_PORTEES : ecarts portees PCM vs reference
- DESSIN_{etude} : schemas techniques polaires des appuis (Matplotlib PNG)
- GLOSSAIRE : 33 termes metier + codes couleur
- DICTIONNAIRE : registre ISO/IEC 11179, toutes colonnes documentees
- reference_values : listes de valeurs officielles CAPFT (Bases)

MODE PROJET
-----------
Selection d'un dossier projet -> SRO derive du nom de dossier -> couches
chargees directement depuis PostgreSQL avec filtre SRO. Detection automatique
du type BE (NGE/Axione) via GraceTHD ou table comac.sro_nge_axione.

COUCHES QGIS REQUISES
----------------------
- infra_pt_pot : poteaux (inf_num, inf_type, etat, noe_codext)
- etude_cap_ft : polygones zones d'etude CAP_FT
- etude_comac : polygones zones d'etude COMAC

Pour Axione, les poteaux sont charges depuis GraceTHD (t_ptech + t_noeud).

BASE DE DONNEES
---------------
- Serveur PostgreSQL (schema rip_avg_nge) : infra_pt_pot, cables, attaches, BPE
- Fonction SQL fddcpiax(sro) : decoupage cables par appui (NGE + Axione)
- Schema comac : supports, cables, hypotheses, armements, communes

STRUCTURE DU PLUGIN
-------------------
PoleAerien/
  |-- PoleAerien.py           # Point d'entree, instancie dialog_v2
  |-- dialog_v2.py            # Interface utilisateur (detection, modules, log)
  |-- batch_orchestrator.py   # Pont entre UI et workflows
  |-- batch_runner.py         # Moteur d'execution sequentielle
  |
  |-- workflows/              # Orchestrateurs par module
  |   |-- maj_workflow.py     # MAJ FT/BT
  |   |-- capft_workflow.py   # Verification CAP_FT
  |   |-- comac_workflow.py   # Verification COMAC
  |   |-- c6bd_workflow.py    # C6 vs BD
  |   |-- c6c3a_workflow.py   # C6-C3A-C7 vs BD
  |   |-- police_workflow.py  # Police C6
  |   |-- gespot_workflow.py  # GESPOT vs C6
  |
  |-- async_tasks.py          # Taches QgsTask (thread worker)
  |
  |-- Comac.py                # Logique metier COMAC
  |-- CapFt.py                # Logique metier CAP_FT
  |-- PoliceC6.py             # Logique metier Police C6
  |-- C6_vs_Bd.py             # Logique metier C6 vs BD
  |-- C6_vs_C3A_vs_Bd.py      # Logique metier C6/C3A/C7
  |-- Maj_Ft_Bt.py            # Logique metier MAJ FT/BT
  |-- maj_sql_background.py   # MAJ SQL directe PostgreSQL
  |
  |-- gespot_reader.py        # Parsing CSV GESPOT
  |-- gespot_c6_comparator.py # Comparaison GESPOT vs C6
  |-- gracethd_reader.py      # Parsing GraceTHD (SHP+CSV)
  |-- pcm_parser.py           # Parsing fichiers PCM (XML)
  |-- pcm_drawing.py          # Schemas polaires Matplotlib
  |-- pcm_bdd_comparator.py   # Comparaison PCM vs BDD
  |
  |-- cable_analyzer.py       # Analyse cables, portees, boitiers
  |-- security_rules.py       # Regles NFC 11201, portees max
  |-- comac_db_reader.py      # Requetes schema comac PostgreSQL
  |
  |-- unified_report.py       # Rapport Excel unifie (openpyxl)
  |-- report_export_task.py   # Tache export rapport (QgsTask)
  |-- data_dictionary.json    # Registre ISO/IEC 11179
  |
  |-- project_detector.py     # Detection automatique structure projet
  |-- preflight_checks.py     # Validations pre-batch
  |-- db_connection.py        # Connexion PostgreSQL (fddcpiax, BPE, attaches)
  |-- db_layer_loader.py      # Chargement couches depuis PostgreSQL
  |-- qgis_utils.py           # Utilitaires QGIS (spatial index, CRS, couches)
  |-- core_utils.py           # Utilitaires Python purs (normalisation, parsing)
  |-- dataclasses_results.py  # Structures de donnees partagees
  |-- compat.py               # Compatibilite Qt5/Qt6 (enums, flags, types)
  |
  |-- sql/                    # Fonctions SQL (fddcpi2, fddcpiax, fdrca)
  |-- styles/                 # Styles QML pour couches QGIS
  |-- tests/                  # Tests unitaires
  |-- docs/                   # Documentation

INSTALLATION
------------
1. Copier le dossier PoleAerien dans le repertoire plugins QGIS :
   - Windows (QGIS 3): %APPDATA%\QGIS\QGIS3\profiles\default\python\plugins\
   - Windows (QGIS 4): %APPDATA%\QGIS\QGIS4\profiles\default\python\plugins\
   - Linux: ~/.local/share/QGIS/QGIS3/profiles/default/python/plugins/
   - macOS: ~/Library/Application Support/QGIS/QGIS3/profiles/default/python/plugins/

2. Redemarrer QGIS

3. Activer le plugin : Extensions > Gerer les extensions > Pole Aerien

DEPENDANCES PYTHON
------------------
- pandas (manipulation DataFrames)
- openpyxl (lecture/ecriture Excel)
- matplotlib (schemas techniques PCM)

AUTEUR ORIGINAL
---------------
SOUMARE Abdoulayi
Email: abdoulayisoumare@gmail.com

EDITEUR / MAINTENEUR
--------------------
Youcef ADDA
Email: yadda@nge-es.fr
LinkedIn: https://www.linkedin.com/in/youcef-adda/

LICENCE
-------
GNU General Public License v2.0 ou ulterieure

VERSION
-------
3.1.0 (Mars 2026)

CHANGELOG
---------
v3.1.0 (2026-03) - Compatibilite QGIS 4.0 (Qt6/PyQt6)
  - [COMPAT] Couche compat.py : resolution enums Qt5/Qt6 au chargement
  - [COMPAT] Fix import sip (PyQt6.sip via try/except)
  - [COMPAT] Fix enums Qt non qualifies (QFrame, QEvent, QTextCursor, QDialogButtonBox)
  - [COMPAT] Compatible QGIS 3.28 a 4.99 (PyQt5 >= 5.15 et PyQt6)

v3.0.0 (2026-03) - Rapport unifie, schemas techniques, dictionnaire de donnees
  - [FEAT] Schemas techniques polaires des appuis depuis fichiers PCM (Matplotlib)
  - [FEAT] Dictionnaire de donnees ISO/IEC 11179 + glossaire 33 termes
  - [FEAT] Valeurs de reference CAPFT (Bases) : 17 listes officielles
  - [FEAT] Comparaison portees PCM vs BDD/GraceTHD (feuille COMAC_PORTEES)
  - [FEAT] Module GESPOT vs C6 : 11 criteres, regles BR-01 a BR-07
  - [FEAT] Comparaison PCM vs BDD : matching multi-strategie supports + cables
  - [FEAT] Verification boitiers COMAC vs BPE
  - [FEAT] Layout dessins grille par zone d'etude COMAC
  - [FEAT] Filtrage poteaux accrochage (BO) dans les dessins
  - [FIX] Faux doublons COMAC inter-communes (keep_commune)
  - [FIX] Cables souterrains GraceTHD comptes par erreur (posemode)
  - [FIX] Cables recales : detection via attaches (cable/BPE des deux cotes)
  - [PERF] Cache fddcpi2 inter-modules (COMAC + Police C6)
  - [PERF] Cache SRO + appuis WKB inter-modules
  - [PERF] Cache CRS validation (evite ~8 comparaisons redondantes)
  - [PERF] Connexion unique comac_db_reader (6 connexions PG -> 1)
  - [PERF] Rapport unifie en tache background (UnifiedReportExportTask)

v2.5.0 (2026-02) - GraceTHD, mode projet, interface V2
  - [FEAT] Pipeline GraceTHD complet (poteaux, cables, BPE depuis SHP+CSV)
  - [FEAT] Detection auto BE type (NGE/Axione) via GraceTHD ou BDD
  - [FEAT] Mode projet : SRO depuis nom dossier, couches PostgreSQL directes
  - [FEAT] Chargement couches asynchrone (background thread)
  - [FEAT] SRO editable dans l'interface
  - [FEAT] Detection inline multi-candidats (COMAC, CAP FT, C6, C3A, C7)
  - [FEAT] Filtrage cables gras dans C6 Excel (bold font)
  - [FEAT] Rapport unifie multi-modules avec TABLEAU DE BORD
  - [FEAT] Preflight checks global avant execution batch
  - [FEAT] Diagnostics actionables par module
  - [FEAT] Liens cliquables vers fichiers de sortie
  - [ARCH] Interface V2 (dialog_v2.py) : remplacement complet ancienne UI
  - [ARCH] Theme-aware : rendu natif Qt, compatible dark/light
  - [ARCH] Suppression ancienne interface (11 fichiers)
  - [FIX] Comptage cables physiques COMAC (GROUP BY GID)
  - [FIX] Routage fddcpi2 -> fddcpiax pour Axione
  - [FIX] Liaison vs cable confusion Police C6 (regex L\d)

v2.3.0 (2026-02) - Corrections performance, architecture workflows
  - [ARCH] Deconstruction God Object PoleAerien.py (2700 -> 128 lignes)
  - [ARCH] Couche workflows/ : 6 orchestrateurs dedies
  - [ARCH] Utils scinde : qgis_utils.py + core_utils.py
  - [ARCH] Normalisation appuis unifiee (normalize_appui_num)
  - [CRITICAL] Filtres spatiaux getFeatures() (-80% temps, -70% memoire)
  - [CRITICAL] Cleanup memoire explicite (evite fuite 500 MB)
  - [CRITICAL] Guards division par zero (security_rules.py)
  - [CRITICAL] Transactions atomiques avec rollback (Maj_Ft_Bt.py)
  - [CRITICAL] Fix injection SQL (PoliceC6.py)
  - [FIX] Migration comac.gpkg -> PostgreSQL (schema comac)

v2.2.2 (2026-01) - Refonte architecturale
  - [ARCH] Threading QGIS conforme : extraction donnees sur main thread
  - [PERF] Index spatial QgsSpatialIndex dans CapFt, Comac, C6_vs_Bd
  - [SEC] Protection injection SQL/Expression
  - [FIX] Acces couches securise via get_layer_safe()
  - [CLEAN] Suppression dependances tabulate, shapely

v2.2.1 (2026-01)
  - Refonte UI : police Corbel, labels humanises

v2.0.1 (2026-01)
  - Correction compatibilite QGIS 3.28 a 3.42 (protection NULL/QVariant)
  - Optimisation spatiale avec QgsSpatialIndex
  - Execution asynchrone MAJ FT/BT (QgsTask)

v2.0.0 (2025)
  - Refonte complete de l'interface
  - Ajout module Police C6
  - Ajout comparaison C6/C3A/C7

v1.0.0 (2022-03)
  - Version initiale (Verification COMAC/CAP_FT)
