#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
Dataclasses pour les résultats des modules métier.

QA-EXPERT: Tous les résultats sont immuables après création.
PERF-SPECIALIST: Pas d'état mutable, retours propres.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional, Any


@dataclass
class ValidationResult:
    """Résultat générique de validation."""
    valide: bool = True
    erreurs: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    
    def add_error(self, msg: str) -> None:
        self.erreurs.append(msg)
        self.valide = False
    
    def add_warning(self, msg: str) -> None:
        self.warnings.append(msg)


@dataclass
class ExcelValidationResult(ValidationResult):
    """Résultat validation fichier Excel."""
    nom_fichier: str = ""
    colonnes_manquantes: List[str] = field(default_factory=list)
    colonnes_en_trop: List[str] = field(default_factory=list)
    structure_ft_ok: bool = False
    structure_bt_ok: bool = False


@dataclass
class PoteauxPolygoneResult:
    """Résultat contrôle poteaux dans polygones (REQ-MAJ-001)."""
    ft_hors_polygone: List[Tuple[str, float, float]] = field(default_factory=list)
    bt_hors_polygone: List[Tuple[str, float, float]] = field(default_factory=list)
    
    @property
    def nb_ft_hors(self) -> int:
        return len(self.ft_hors_polygone)
    
    @property
    def nb_bt_hors(self) -> int:
        return len(self.bt_hors_polygone)
    
    @property
    def tous_dans_polygone(self) -> bool:
        return self.nb_ft_hors == 0 and self.nb_bt_hors == 0


@dataclass
class EtudesValidationResult:
    """Résultat vérification études existent (REQ-MAJ-004)."""
    etudes_absentes_cap_ft: List[str] = field(default_factory=list)
    etudes_absentes_comac: List[str] = field(default_factory=list)
    
    @property
    def toutes_existent(self) -> bool:
        return len(self.etudes_absentes_cap_ft) == 0 and len(self.etudes_absentes_comac) == 0


@dataclass
class ImplantationValidationResult:
    """Résultat vérification type poteau si implantation (REQ-MAJ-006)."""
    erreurs_implantation: List[Dict[str, Any]] = field(default_factory=list)
    
    @property
    def valide(self) -> bool:
        return len(self.erreurs_implantation) == 0


@dataclass
class ActionsValidationResult:
    """Résultat validation actions FT/BT (REQ-C6BD-003/004)."""
    erreurs_actions_ft: List[Dict[str, Any]] = field(default_factory=list)
    erreurs_actions_bt: List[Dict[str, Any]] = field(default_factory=list)
    
    @property
    def valide(self) -> bool:
        return len(self.erreurs_actions_ft) == 0 and len(self.erreurs_actions_bt) == 0


@dataclass
class MajAttributsC6Result:
    """Résultat MAJ attributs depuis C6 (REQ-C6BD-001/002)."""
    nb_etiquette_jaune: int = 0
    nb_zone_privee: int = 0
    erreurs: List[str] = field(default_factory=list)


@dataclass
class PoliceC6Result:
    """Résultat analyse Police C6 (refactoring PoliceC6.py).
    
    Remplace les attributs d'état mutable de PoliceC6.__init__
    PERF-SPECIALIST: Pas d'état global, résultats immuables.
    """
    nb_appui_corresp: int = 0
    nb_pbo_corresp: int = 0
    bpo_corresp: List[str] = field(default_factory=list)
    nb_appui_absent: int = 0
    nb_appui_absent_pot: int = 0
    pot_inf_num_present: List[str] = field(default_factory=list)
    inf_num_pot_absent: List[str] = field(default_factory=list)
    ebp_non_appui: List[str] = field(default_factory=list)
    absence: List[str] = field(default_factory=list)
    id_pot_present: List[int] = field(default_factory=list)
    id_pot_absent: List[int] = field(default_factory=list)
    bpe_pot_cap_ft: List[Any] = field(default_factory=list)
    ebp_appui_inconnu: List[str] = field(default_factory=list)
    liste_appui_ebp: List[List[Any]] = field(default_factory=list)
    presence_liste_appui_ebp: bool = False
    pbo_a_supprimer: set = field(default_factory=set)
    ebp_a_supprimer: set = field(default_factory=set)
    liste_cable_appui_trouve: List[List[Any]] = field(default_factory=list)
    liste_appui_capa_appui_absent: List[Any] = field(default_factory=list)


@dataclass
class CableCapaciteResult:
    """Résultat vérification capacité câbles (REQ-PLC6-006)."""
    anomalies: List[Dict[str, Any]] = field(default_factory=list)
    cables_traites: int = 0
    cables_valides: int = 0
    
    @property
    def taux_validite(self) -> float:
        if self.cables_traites == 0:
            return 100.0
        return (self.cables_valides / self.cables_traites) * 100


@dataclass
class BoitierValidationResult:
    """Résultat vérification boîtiers (REQ-PLC6-007)."""
    anomalies: List[Dict[str, Any]] = field(default_factory=list)
    boitiers_traites: int = 0
    boitiers_valides: int = 0


@dataclass
class EtudeC6Result:
    """Résultat analyse d'une étude C6."""
    etude: str = ""
    chemin_c6: str = ""
    statut: str = "PENDING"
    resultat: Optional[PoliceC6Result] = None
    erreur: Optional[str] = None


@dataclass 
class ParcourAutoC6Result:
    """Résultat parcours automatique C6 (REQ-PLC6-003)."""
    etudes_traitees: List[EtudeC6Result] = field(default_factory=list)
    
    @property
    def nb_ok(self) -> int:
        return sum(1 for e in self.etudes_traitees if e.statut == "OK")
    
    @property
    def nb_erreur(self) -> int:
        return sum(1 for e in self.etudes_traitees if e.statut == "ERREUR")
    
    @property
    def nb_total(self) -> int:
        return len(self.etudes_traitees)


# =============================================================================
# COMPARAISON BD PostgreSQL ↔ GraceTHD (G01-G10)
# =============================================================================

@dataclass
class PoteauxComparisonResult:
    """Résultat comparaison poteaux BD ↔ GraceTHD (G01/G02).
    
    G01: Poteaux BD manquants dans GraceTHD
    G02: Poteaux GraceTHD excédentaires (pas dans BD)
    """
    manquants_gracethd: set = field(default_factory=set)
    excedentaires_gracethd: set = field(default_factory=set)
    correspondances: set = field(default_factory=set)
    erreurs: List[str] = field(default_factory=list)
    
    @property
    def nb_total_bd(self) -> int:
        return len(self.correspondances) + len(self.manquants_gracethd)
    
    @property
    def nb_total_gracethd(self) -> int:
        return len(self.correspondances) + len(self.excedentaires_gracethd)
    
    @property
    def taux_correspondance(self) -> float:
        total = self.nb_total_bd
        return (len(self.correspondances) / total * 100) if total > 0 else 0.0


@dataclass
class BPEComparisonResult:
    """Résultat comparaison BPE BD ↔ GraceTHD (G03/G04/G09).
    
    G03: BPE BD manquants dans GraceTHD
    G04: BPE GraceTHD excédentaires
    G09: Types boîtiers différents
    """
    manquants_gracethd: set = field(default_factory=set)
    excedentaires_gracethd: set = field(default_factory=set)
    types_differents: List[Dict[str, Any]] = field(default_factory=list)
    correspondances: set = field(default_factory=set)
    erreurs: List[str] = field(default_factory=list)
    
    @property
    def nb_anomalies(self) -> int:
        return len(self.manquants_gracethd) + len(self.types_differents)
    
    @property
    def taux_correspondance(self) -> float:
        total = len(self.correspondances) + len(self.manquants_gracethd)
        return (len(self.correspondances) / total * 100) if total > 0 else 0.0


@dataclass
class CablesComparisonResult:
    """Résultat comparaison câbles BD ↔ GraceTHD (G05/G06/G07).
    
    G05: Câbles BD manquants dans GraceTHD
    G06: Câbles GraceTHD excédentaires
    G07: Capacités FO différentes
    """
    manquants_gracethd: set = field(default_factory=set)
    excedentaires_gracethd: set = field(default_factory=set)
    capacites_differentes: List[Dict[str, Any]] = field(default_factory=list)
    correspondances: set = field(default_factory=set)
    erreurs: List[str] = field(default_factory=list)
    
    @property
    def nb_anomalies(self) -> int:
        return len(self.manquants_gracethd) + len(self.capacites_differentes)
    
    @property
    def taux_correspondance(self) -> float:
        total = len(self.correspondances) + len(self.manquants_gracethd)
        return (len(self.correspondances) / total * 100) if total > 0 else 0.0


@dataclass
class CheminementsComparisonResult:
    """Résultat comparaison cheminements BD ↔ GraceTHD (G08)."""
    manquants_gracethd: set = field(default_factory=set)
    excedentaires_gracethd: set = field(default_factory=set)
    correspondances: set = field(default_factory=set)
    erreurs: List[str] = field(default_factory=list)


@dataclass
class GraceTHDComparisonResult:
    """Résultat agrégé de toutes les comparaisons BD ↔ GraceTHD.
    
    Regroupe les résultats de G01-G10 en un seul objet.
    """
    poteaux: Optional[PoteauxComparisonResult] = None
    bpe: Optional[BPEComparisonResult] = None
    cables: Optional[CablesComparisonResult] = None
    cheminements: Optional[CheminementsComparisonResult] = None
    gracethd_path: str = ""
    
    @property
    def nb_anomalies_total(self) -> int:
        count = 0
        if self.poteaux:
            count += len(self.poteaux.manquants_gracethd)
        if self.bpe:
            count += self.bpe.nb_anomalies
        if self.cables:
            count += self.cables.nb_anomalies
        if self.cheminements:
            count += len(self.cheminements.manquants_gracethd)
        return count
    
    @property
    def is_valid(self) -> bool:
        return self.nb_anomalies_total == 0
    
    def to_summary(self) -> Dict[str, Any]:
        """Génère un résumé pour affichage UI."""
        return {
            'poteaux': {
                'correspondances': len(self.poteaux.correspondances) if self.poteaux else 0,
                'manquants': len(self.poteaux.manquants_gracethd) if self.poteaux else 0,
                'excedentaires': len(self.poteaux.excedentaires_gracethd) if self.poteaux else 0,
            },
            'bpe': {
                'correspondances': len(self.bpe.correspondances) if self.bpe else 0,
                'manquants': len(self.bpe.manquants_gracethd) if self.bpe else 0,
                'types_diff': len(self.bpe.types_differents) if self.bpe else 0,
            },
            'cables': {
                'correspondances': len(self.cables.correspondances) if self.cables else 0,
                'manquants': len(self.cables.manquants_gracethd) if self.cables else 0,
                'capa_diff': len(self.cables.capacites_differentes) if self.cables else 0,
            },
            'total_anomalies': self.nb_anomalies_total,
        }
