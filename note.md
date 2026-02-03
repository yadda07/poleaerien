
Notes de réunion
Zone de test pour Youcef:
ZAPM = 63041/B1I/PMZ/00003


MAJ BDD :Vérifier la structure de chaque table et renvoyer un retour (comparer par rapport le tableau type): Pas OK
Vérifier si état implantation, il faut avoir le type de poteau
Sur l'onglet FT si Action = implantation donc MAJ infra_pt_pot colonne état = 'FT KO' et inf_propri = RAUV et inf_type = 'POT-AC' et donne un nouveau nommage en création pour ces poteaux et en champs commentaire = 'POT FT (ancien nommage du FT est FT KO)
MAJ champs étiquette Jaune = oui si excale/ etiquette jaune = x , étiquette orange si excel/Action= 'recalage' ; 
Manque zone privé si zone privé = 'x' donc infra_pt_pot -- commentaire rajoute 'PRIVE' 
Sur le excel FT BT KO dans ONGLET ft/ Vérifier si action est dans [Renforcement, Recalage, Remplacement]
Pour onglet BT, Vérifier si Action = Implantation si ligne non vide sinon erreur


C6 vs BD :Vérifier si tous les poteaux FT sontcouverte par le polygone capft. (Liste in/ out)
Vérifier si tous les nom étude dans le polygone capft existent dans le répertoire des C6 sélectionné
Vérifier le fonctionnement du mode SRO et découpage (à supprimer si besoin) --> rajouté par Youcef
C6vsBD le nom d'etudes à sélectionner à virer de l'interface


POLICE C6 :Donner la possibilité de Sélectionner le GRACE THD
Parcourir automatiquement C6 par C6 à partir de la table attributaire de la couche etude_cap_ft


Police C6 --> Devient POLICE C6/COMACRajouter la couche zone décopage etude capft et chemin études COMAC avec les fichiers excel
Capacité de câble par rapport à la colonne type de ligne réseau en nappe : @Zied de communiquer la liste de correspondance des capacités des câbles contre les références des câbles dans le fichier COMAC
Vérification boitier par rapport à la colonne Boitier

NB : les réf des capacités des câbles sur les études COMAC sont séparés par un tiret '-' ; Exemple : 
 1 câble -> colonne AO = L1092-13-P-
 2 câbles -> colonne AO = L1092-13-P-L1092-12-P- 
