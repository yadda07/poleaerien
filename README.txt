================================================================================
                              POLE AERIEN - Plugin QGIS
                    Controle et MAJ des Poteaux ENEDIS (FT/BT)
================================================================================

DESCRIPTION
-----------
Plugin QGIS pour le controle qualite et la mise a jour des donnees de poteaux
aeriens ENEDIS dans le cadre des projets FTTH. Il permet de comparer les
donnees terrain (fichiers Excel sous-traitants) avec la base de donnees QGIS
et de mettre a jour automatiquement les attributs des poteaux.

COMPATIBILITE
-------------
- QGIS : 3.28 a 3.42
- Python : 3.9+
- OS : Windows, Linux, macOS

FONCTIONNALITES
---------------

1. MAJ FT/BT (Onglet "MAJ BD")
   - Lecture du fichier Excel FT-BT KO (onglets FT et BT)
   - Extraction spatiale des poteaux par etude (CAP_FT et COMAC)
   - Comparaison et mise a jour des attributs :
     * etat (A RECALER, A REMPLACER, A RENFORCER, BT KO)
     * action (RECALAGE, REMPLACEMENT, RENFORCEMENT, IMPLANTATION)
     * inf_mat, inf_type, inf_propri, noe_usage, dce
   - Execution asynchrone (non-bloquante)

2. Verification CAP_FT (Onglet "Verif CAP_FT")
   - Detection des poteaux FT hors decoupage etude
   - Detection des doublons de noms d'etudes
   - Comparaison QGIS vs Fiches Appuis Excel
   - Export rapport d'analyse Excel

3. Verification COMAC (Onglet "Verif COMAC")
   - Detection des poteaux BT hors decoupage etude
   - Detection des doublons de noms d'etudes
   - Comparaison QGIS vs fichiers ExportComac.xlsx
   - Export rapport d'analyse Excel

4. C6 vs BD (Onglet "C6 vs BD")
   - Lecture des annexes C6 (fichiers Excel)
   - Comparaison avec les poteaux FT de la base QGIS
   - Detection des ecarts (presents dans C6, absents de QGIS et vice-versa)
   - Export rapport comparatif Excel

5. C6/C3A/C7 vs BD (Onglet "C6 C3A C7 vs BD")
   - Analyse croisee des annexes C6, C3A et C7
   - Support QGIS ou Excel pour les donnees C3A (table CMD)
   - Filtrage par decoupage et valeur de champ
   - Export rapport d'analyse Excel

6. Police C6 (Onglet "Police C6")
   - Analyse detaillee des annexes C6 avec donnees GraceTHD
   - Verification des correspondances appui/cable/PBO
   - Detection des incoherences entre C6 et tables GraceTHD
   - Generation de couches erreurs dans QGIS

TABLES QGIS REQUISES
--------------------
- infra_pt_pot : Table des poteaux (champs: inf_num, inf_type, etat, etc.)
- etude_cap_ft : Decoupage des etudes CAP_FT (polygones)
- etude_comac : Decoupage des etudes COMAC (polygones)
- Tables GraceTHD (pour Police C6) : t_cable, t_cableline, t_ptech, etc.

STRUCTURE DU PLUGIN
-------------------
PoleAerien/
  |-- PoleAerien.py        # Controleur UI principal
  |-- async_tasks.py       # Taches asynchrones (QgsTask) thread-safe
  |-- Maj_Ft_Bt.py         # MAJ asynchrone FT/BT
  |-- PoliceC6.py          # Analyse Police C6
  |-- Comac.py             # Verification COMAC (index spatial)
  |-- CapFt.py             # Verification CAP_FT (index spatial)
  |-- C6_vs_Bd.py          # Comparaison C6 vs BD (index spatial)
  |-- C6_vs_C3A_vs_Bd.py   # Comparaison C6/C3A/C7
  |-- SecondFile.py        # Fonctions UI utilitaires
  |-- utils.py             # Fonctions partagees (get_layer_safe, etc.)
  |-- interfaces/          # Fichiers UI Qt Designer
  |-- docs/                # Documentation HTML

INSTALLATION
------------
1. Copier le dossier PoleAerien dans le repertoire plugins QGIS :
   - Windows: %APPDATA%\QGIS\QGIS3\profiles\default\python\plugins\
   - Linux: ~/.local/share/QGIS/QGIS3/profiles/default/python/plugins/
   - macOS: ~/Library/Application Support/QGIS/QGIS3/profiles/default/python/plugins/

2. Redemarrer QGIS

3. Activer le plugin : Extensions > Gerer les extensions > Pole Aerien

DEPENDANCES PYTHON
------------------
- pandas (manipulation DataFrames)
- openpyxl (lecture/ecriture Excel)

Note: tabulate et shapely ne sont plus requis depuis v2.2.2.

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
2.2.2 (Janvier 2026)

CHANGELOG
---------
v2.2.2 (2026-01) - Refonte Architecturale
  - [ARCH] Threading QGIS conforme : extraction donnees sur main thread
  - [ARCH] async_tasks.py : CapFtTask/ComacTask recoivent donnees pures
  - [ARCH] Tous attributs PoliceC6 definis dans __init__
  - [PERF] Index spatial QgsSpatialIndex dans CapFt, Comac, C6_vs_Bd
  - [SEC] Protection injection SQL/Expression (Maj_Ft_Bt, C6_vs_C3A_vs_Bd)
  - [FIX] Bug modification pendant iteration (Comac.traitementResultatFinaux)
  - [FIX] Acces couches securise via get_layer_safe() avec isValid()
  - [CLEAN] Suppression dependances tabulate, shapely

v2.2.1 (2026-01)
  - Refonte UI : police Corbel, labels humanises, suppression boutons Actualiser

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
