"""
Analyseur de charge câbles par appui - Police C6 v2.0
Compte les câbles découpés qui intersectent chaque appui
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple
import re
from qgis.core import (
    QgsVectorLayer, QgsFeature, QgsGeometry, QgsPointXY,
    QgsMessageLog, Qgis, QgsSpatialIndex, QgsFeatureRequest
)

from .db_connection import CableSegment


@dataclass
class AppuiChargeResult:
    """Résultat de l'analyse de charge pour un appui"""
    num_appui: str
    geom: QgsGeometry
    
    # Données BDD (câbles découpés)
    nb_cables_bdd: int = 0
    capacite_totale_bdd: int = 0
    cables_details: List[CableSegment] = field(default_factory=list)
    
    # Données C6 (Annexe C6)
    nb_cables_c6: int = 0
    capacite_totale_c6: int = 0
    refs_cables_c6: List[Dict] = field(default_factory=list)
    
    # Comparaison
    match_nb: bool = False
    match_capa: bool = False
    anomalie: bool = False
    message_anomalie: str = ""


@dataclass
class CableRefC6:
    """Référence de câble extraite de l'Annexe C6"""
    reference: str      # Ex: "L192.11"
    capacite: int       # Ex: 26 (de "-26P")
    raw_text: str       # Texte original


class CableAnalyzer:
    """
    Analyse la charge de câbles par appui.
    Compare les câbles découpés (BDD) avec les déclarations Annexe C6.
    """
    
    def __init__(self, tolerance: float = 0.5):
        """
        Args:
            tolerance: Distance max (m) pour considérer qu'un câble touche un appui
        """
        self.tolerance = tolerance
        self.resultats: Dict[str, AppuiChargeResult] = {}
    
    def analyser_charge_appuis(
        self,
        appuis: List[Dict],
        cables_decoupes: List[CableSegment],
        only_aerien: bool = True
    ) -> Dict[str, AppuiChargeResult]:
        """
        Pour chaque appui, compte les câbles découpés qui le touchent.
        
        Args:
            appuis: Liste de dicts avec 'num_appui' et 'geom' (QgsGeometry ou QgsPointXY)
            cables_decoupes: Segments de câbles depuis fddcpi2
            only_aerien: Si True, ne compte que les câbles aériens (posemode=1)
        
        Returns:
            Dict[num_appui, AppuiChargeResult]
        """
        self.resultats = {}
        
        # Filtrer câbles aériens + façade si demandé
        if only_aerien:
            cables = [c for c in cables_decoupes if c.posemode in (1, 2)]
        else:
            cables = cables_decoupes
        
        QgsMessageLog.logMessage(
            f"Analyse charge: {len(appuis)} appuis, {len(cables)} câbles "
            f"({'aériens/façade' if only_aerien else 'tous'})",
            "PoleAerien", Qgis.Info
        )
        
        for appui in appuis:
            num_appui = str(appui.get('num_appui', ''))
            geom = appui.get('geom')
            
            if not num_appui or not geom:
                continue
            
            # Convertir en QgsGeometry si nécessaire
            if isinstance(geom, QgsPointXY):
                geom = QgsGeometry.fromPointXY(geom)
            
            # Trouver les câbles qui touchent cet appui
            cables_touchant = self._find_cables_touching_appui(geom, cables)
            
            # Calculer les stats
            result = AppuiChargeResult(
                num_appui=num_appui,
                geom=geom,
                nb_cables_bdd=len(cables_touchant),
                capacite_totale_bdd=sum(c.cab_capa for c in cables_touchant),
                cables_details=cables_touchant
            )
            
            self.resultats[num_appui] = result
        
        return self.resultats
    
    def _find_cables_touching_appui(
        self,
        appui_geom: QgsGeometry,
        cables: List[CableSegment]
    ) -> List[CableSegment]:
        """
        Trouve tous les câbles dont une extrémité touche l'appui.
        """
        touching = []
        appui_point = appui_geom.asPoint()
        
        for cable in cables:
            # Parser la géométrie WKT du câble
            cable_geom = QgsGeometry.fromWkt(cable.geom_wkt)
            if not cable_geom or cable_geom.isEmpty():
                continue
            
            # Vérifier les extrémités
            if cable_geom.type() == 1:  # LineString
                line = cable_geom.asPolyline()
                if len(line) >= 2:
                    start_point = line[0]
                    end_point = line[-1]
                    
                    # Distance aux extrémités
                    dist_start = appui_point.distance(start_point)
                    dist_end = appui_point.distance(end_point)
                    
                    if dist_start <= self.tolerance or dist_end <= self.tolerance:
                        touching.append(cable)
        
        return touching
    
    def enrichir_avec_c6(
        self,
        donnees_c6: Dict[str, str]
    ):
        """
        Enrichit les résultats avec les données de l'Annexe C6.
        
        Args:
            donnees_c6: Dict[num_appui, nom_cable] depuis l'Annexe C6
        """
        for num_appui, result in self.resultats.items():
            nom_cable_c6 = donnees_c6.get(num_appui, '')
            
            if nom_cable_c6:
                refs = self.parser_references_cables_c6(nom_cable_c6)
                result.nb_cables_c6 = len(refs)
                result.capacite_totale_c6 = sum(r['capacite'] for r in refs)
                result.refs_cables_c6 = refs
            
            # Comparer
            result.match_nb = (result.nb_cables_bdd == result.nb_cables_c6)
            result.match_capa = (result.capacite_totale_bdd == result.capacite_totale_c6)
            result.anomalie = not (result.match_nb and result.match_capa)
            
            if result.anomalie:
                result.message_anomalie = self._generer_message_anomalie(result)
    
    @staticmethod
    def parser_references_cables_c6(nom_cable: str) -> List[Dict]:
        """
        Parse la colonne 'Nom du câble' de l'Annexe C6.
        
        Formats reconnus:
        - "L192.11-26P" → ref='L192.11', capa=26
        - "L192.11-26P | L193.12-6P" → 2 références
        - "L192.11-26P L193.12-6P" → 2 références
        
        Args:
            nom_cable: Contenu de la colonne "Nom du câble"
        
        Returns:
            Liste de dicts {'reference': str, 'capacite': int, 'raw': str}
        """
        if not nom_cable or not isinstance(nom_cable, str):
            return []
        
        references = []
        
        # Pattern: Lxxx.xx-xxP ou similaire
        # Exemples: L192.11-26P, L193.12-6P, CB-12P
        pattern = r'([A-Z]{1,3}\d*\.?\d*)-(\d+)P'
        
        matches = re.findall(pattern, nom_cable, re.IGNORECASE)
        
        for ref, capa in matches:
            references.append({
                'reference': ref,
                'capacite': int(capa),
                'raw': f"{ref}-{capa}P"
            })
        
        # Si aucun match, essayer un pattern plus simple
        if not references:
            # Pattern: juste -xxP pour extraire les capacités
            simple_pattern = r'-(\d+)P'
            simple_matches = re.findall(simple_pattern, nom_cable, re.IGNORECASE)
            for idx, capa in enumerate(simple_matches):
                references.append({
                    'reference': f"REF{idx+1}",
                    'capacite': int(capa),
                    'raw': f"-{capa}P"
                })
        
        return references
    
    def _generer_message_anomalie(self, result: AppuiChargeResult) -> str:
        """Génère un message explicatif pour l'anomalie."""
        messages = []
        
        if result.nb_cables_bdd != result.nb_cables_c6:
            diff = result.nb_cables_bdd - result.nb_cables_c6
            if diff > 0:
                messages.append(
                    f"BDD: {result.nb_cables_bdd} câbles, C6: {result.nb_cables_c6} "
                    f"(+{diff} câbles en BDD non déclarés en C6)"
                )
            else:
                messages.append(
                    f"BDD: {result.nb_cables_bdd} câbles, C6: {result.nb_cables_c6} "
                    f"({abs(diff)} câbles C6 non trouvés en BDD)"
                )
        
        if result.capacite_totale_bdd != result.capacite_totale_c6:
            diff = result.capacite_totale_bdd - result.capacite_totale_c6
            messages.append(
                f"Capacité BDD: {result.capacite_totale_bdd} FO, C6: {result.capacite_totale_c6} FO "
                f"(différence: {diff:+d} FO)"
            )
        
        return " | ".join(messages) if messages else ""
    
    def get_anomalies(self) -> List[AppuiChargeResult]:
        """Retourne uniquement les appuis avec anomalie."""
        return [r for r in self.resultats.values() if r.anomalie]
    
    def get_stats_globales(self) -> Dict:
        """Retourne les statistiques globales."""
        total = len(self.resultats)
        anomalies = len(self.get_anomalies())
        
        return {
            'total_appuis': total,
            'appuis_ok': total - anomalies,
            'appuis_anomalie': anomalies,
            'total_cables_bdd': sum(r.nb_cables_bdd for r in self.resultats.values()),
            'total_cables_c6': sum(r.nb_cables_c6 for r in self.resultats.values()),
            'total_capa_bdd': sum(r.capacite_totale_bdd for r in self.resultats.values()),
            'total_capa_c6': sum(r.capacite_totale_c6 for r in self.resultats.values()),
        }


def extraire_appuis_from_layer(layer: QgsVectorLayer, field_num_appui: str = 'num_appui') -> List[Dict]:
    """
    Extrait les appuis depuis une couche QGIS.
    
    Args:
        layer: Couche de points (poteaux)
        field_num_appui: Nom du champ contenant le numéro d'appui
    
    Returns:
        Liste de dicts {'num_appui': str, 'geom': QgsGeometry}
    """
    appuis = []
    
    if not layer or not layer.isValid():
        return appuis
    
    # Chercher le bon nom de champ
    field_names = [f.name() for f in layer.fields()]
    actual_field = None
    
    for candidate in [field_num_appui, 'inf_num', 'INF_NUM', 'num_appui', 'NUM_APPUI', 'pt_ad_numsu', 'PT_AD_NUMSU']:
        if candidate in field_names:
            actual_field = candidate
            break
    
    if not actual_field:
        QgsMessageLog.logMessage(
            f"Champ numéro d'appui non trouvé dans {layer.name()}. "
            f"Champs disponibles: {field_names[:10]}...",
            "PoleAerien", Qgis.Warning
        )
        return appuis
    
    from .core_utils import normalize_appui_num
    
    for feature in layer.getFeatures():
        num_appui = feature[actual_field]
        if num_appui:
            num_norm = normalize_appui_num(num_appui)
            if num_norm:
                appuis.append({
                    'num_appui': num_norm,
                    'geom': feature.geometry(),
                    'feature_id': feature.id()
                })
    
    return appuis


def extraire_appuis_wkb(layer: QgsVectorLayer, field_num_appui: str = 'num_appui') -> List[Dict]:
    """Extrait appuis depuis couche QGIS et serialise geometries en WKB.
    
    Combine extraire_appuis_from_layer() + WKB serialization en un seul appel.
    Le resultat est thread-safe (pas de QgsGeometry, juste des bytes WKB).
    
    Args:
        layer: Couche de points (poteaux)
        field_num_appui: Nom du champ contenant le numero d'appui
    
    Returns:
        Liste de dicts {'num_appui': str, 'feature_id': int, 'geom_wkb': bytes|None}
    """
    raw_appuis = extraire_appuis_from_layer(layer, field_num_appui)
    appuis_wkb = []
    for appui in raw_appuis:
        geom = appui.get('geom')
        appuis_wkb.append({
            'num_appui': appui.get('num_appui', ''),
            'feature_id': appui.get('feature_id'),
            'geom_wkb': geom.asWkb().data() if geom and not geom.isNull() else None
        })
    return appuis_wkb


def compter_cables_par_appui(
    cables: List[CableSegment],
    appuis: List[Dict],
    tolerance: float = 0.5,
    group_by_gid: bool = False
) -> Dict[str, Dict]:
    """
    Compte les câbles qui touchent chaque appui.
    
    Args:
        cables: Liste de CableSegment depuis fddcpi2
        appuis: Liste de dicts avec 'num_appui' et 'geom'
        tolerance: Distance max en mètres pour l'intersection
        group_by_gid: Si True, compte les câbles physiques (distinct GID)
            au lieu des segments découpés (gid_dc2). Utiliser True pour COMAC
            (références = câbles physiques), False pour Police C6
            (références = câbles découpés par zone d'étude).
    
    Returns:
        Dict[num_appui, {'count': int, 'capacites': List[int], 'cables': List}]
    """
    result = {}
    
    if not cables or not appuis:
        return result
    
    # Créer index spatial des appuis
    appuis_by_id = {}
    for i, appui in enumerate(appuis):
        num = appui.get('num_appui', '')
        if num:
            appuis_by_id[num] = appui
            result[num] = {
                'count': 0,
                'capacites': [],
                'cables': [],
                '_gids_seen': set()
            }
    
    # Pour chaque câble, trouver les appuis aux extrémités
    # Garder câbles de distribution aériens + façade (cab_type='CDI' + posemode 1 ou 2)
    for cable in cables:
        if getattr(cable, 'cab_type', '') != 'CDI' or getattr(cable, 'posemode', 0) not in (1, 2):
            continue
        
        wkt = cable.geom_wkt if hasattr(cable, 'geom_wkt') else None
        if not wkt:
            continue
        
        try:
            cable_geom = QgsGeometry.fromWkt(wkt)
            if cable_geom.isNull() or cable_geom.isEmpty():
                continue
            
            # Récupérer les extrémités du câble
            if cable_geom.isMultipart():
                lines = cable_geom.asMultiPolyline()
                line = lines[0] if lines else []
            else:
                line = cable_geom.asPolyline()
            
            if len(line) < 2:
                continue
            
            start_point = QgsGeometry.fromPointXY(line[0])
            end_point = QgsGeometry.fromPointXY(line[-1])
            
            # Chercher l'appui le plus proche de chaque extrémité
            for appui_num, appui_data in appuis_by_id.items():
                appui_geom = appui_data.get('geom')
                if not appui_geom:
                    continue
                
                # Vérifier distance aux extrémités
                dist_start = appui_geom.distance(start_point)
                dist_end = appui_geom.distance(end_point)
                
                if dist_start <= tolerance or dist_end <= tolerance:
                    # Déduplication par GID physique si demandé
                    if group_by_gid:
                        gid = getattr(cable, 'gid', 0)
                        if gid in result[appui_num]['_gids_seen']:
                            continue
                        result[appui_num]['_gids_seen'].add(gid)
                    
                    result[appui_num]['count'] += 1
                    if cable.cab_capa:
                        result[appui_num]['capacites'].append(cable.cab_capa)
                    result[appui_num]['cables'].append({
                        'id': cable.gid_dc2,
                        'gid': getattr(cable, 'gid', 0),
                        'capacite': cable.cab_capa,
                        'posemode': cable.posemode
                    })
        except Exception as e:
            QgsMessageLog.logMessage(
                f"Erreur comptage câble: {e}",
                "PoleAerien", Qgis.Warning
            )
    
    # Nettoyer les sets internes
    for data in result.values():
        data.pop('_gids_seen', None)
    
    return result


def comparer_source_cables(
    dico_cables_source: Dict[str, List[str]],
    cables_par_appui: Dict[str, Dict],
    source_label: str = "SOURCE",
    get_capacites_fn=None
) -> List[Dict]:
    """Compare les cables declares dans une source (C6 ou COMAC) avec les cables BDD (fddcpi2).
    
    Fonction generique qui factorise la logique identique de
    Comac.comparer_comac_cables() et PoliceC6.comparer_c6_cables().
    
    Args:
        dico_cables_source: {appui_norm -> [refs_cables]} depuis parsing Excel
        cables_par_appui: {appui_norm -> {count, capacites, cables}} depuis compter_cables_par_appui()
        source_label: Label pour les messages et cles de sortie ('C6' ou 'COMAC')
        get_capacites_fn: Callable(ref) -> List[int] pour resoudre capacites possibles.
            Si None, utilise security_rules.get_capacites_possibles.
    
    Returns:
        Liste de dicts avec comparaison par appui:
        [{num_appui, nb_cables_source, cables_source, capas_source,
          nb_cables_bdd, capas_bdd, statut, message}]
    """
    if get_capacites_fn is None:
        from .security_rules import get_capacites_possibles
        get_capacites_fn = get_capacites_possibles

    comparaison = []
    label_lower = source_label.lower()

    for num_appui, refs_source in dico_cables_source.items():
        bdd_data = cables_par_appui.get(num_appui, {})
        nb_cables_bdd = bdd_data.get('count', 0)
        nb_cables_source = len(refs_source)

        capas_possibles = [get_capacites_fn(ref) for ref in refs_source]
        capas_bdd = sorted(bdd_data.get('capacites', []))
        capas_display = ['/'.join(str(c) for c in cp) for cp in capas_possibles]

        messages = []
        if not bdd_data:
            statut = "ABSENT_BDD"
            messages.append("Appui non trouvé dans les appuis QGIS")
        else:
            ecart_count = nb_cables_source != nb_cables_bdd
            ecart_capa = not capacites_compatibles(capas_possibles, capas_bdd)

            if ecart_count:
                diff = nb_cables_bdd - nb_cables_source
                messages.append(
                    f"Écart nombre: {diff:+d} câble(s) (BDD={nb_cables_bdd}, {source_label}={nb_cables_source})"
                )
            if ecart_capa:
                src_str = '+'.join(capas_display) or '?'
                bdd_str = '+'.join(str(c) for c in capas_bdd) or '?'
                messages.append(
                    f"Écart capacité: {source_label}=[{src_str}] vs BDD=[{bdd_str}]"
                )

            statut = "ECART" if (ecart_count or ecart_capa) else "OK"

        comparaison.append({
            'num_appui': num_appui,
            f'nb_cables_{label_lower}': nb_cables_source,
            f'cables_{label_lower}': '; '.join(refs_source),
            f'capas_{label_lower}': capas_display,
            'nb_cables_bdd': nb_cables_bdd,
            'capas_bdd': capas_bdd,
            'statut': statut,
            'message': ' | '.join(messages)
        })

    return comparaison


def capacites_compatibles(
    capas_possibles_source: List[List[int]],
    capas_bdd: List[int]
) -> bool:
    """
    Vérifie si les capacités BDD sont compatibles avec les câbles source (C6 ou COMAC).
    Chaque capacité BDD doit pouvoir être associée à un câble source
    dont les capacités possibles incluent cette valeur.
    
    Ex: source = [[24,36], [24,36]], BDD = [24, 36] → True
    Ex: source = [[24,36]], BDD = [72] → False
    
    Args:
        capas_possibles_source: Liste de listes de capacités possibles par câble source
        capas_bdd: Liste des capacités BDD (cab_capa)
    
    Returns:
        True si toutes les capacités BDD matchent un câble source
    """
    if len(capas_possibles_source) != len(capas_bdd):
        return False
    
    if not capas_bdd and not capas_possibles_source:
        return True
    
    # Matching glouton : pour chaque capacité BDD, trouver un câble source compatible
    used = [False] * len(capas_possibles_source)
    
    for bdd_capa in sorted(capas_bdd):
        matched = False
        for i, possibles in enumerate(capas_possibles_source):
            if not used[i] and bdd_capa in possibles:
                used[i] = True
                matched = True
                break
        if not matched:
            return False
    
    return True


def verifier_boitiers(
    boitier_source: Dict[str, str],
    appuis_data: List[Dict],
    bpe_geoms: List[Dict],
    tolerance: float = 1.0
) -> Dict[str, Dict]:
    """Verifie la presence de BPE pour chaque appui declarant un boitier.
    
    Fonction generique utilisee par ComacTask et PoliceC6Task.
    
    Args:
        boitier_source: {num_appui: valeur_boitier} (ex: "PB", "PEO", "oui")
        appuis_data: Liste de dicts avec 'num_appui' et 'geom' (QgsGeometry)
        bpe_geoms: Liste de dicts avec 'geom' (QgsGeometry), 'noe_type', 'gid'
        tolerance: Distance max en metres pour le matching spatial (defaut 1m)
    
    Returns:
        Dict[num_appui, {boitier_source, bpe_trouve, bpe_noe_type, statut}]
    """
    result = {}
    
    if not boitier_source:
        return result
    
    appuis_by_num = {}
    for appui in appuis_data:
        num = appui.get('num_appui', '')
        if num and appui.get('geom'):
            appuis_by_num[num] = appui['geom']
    
    for num_appui, type_boitier in boitier_source.items():
        appui_geom = appuis_by_num.get(num_appui)
        
        entry = {
            'boitier_source': type_boitier,
            'bpe_trouve': False,
            'bpe_noe_type': '',
            'statut': ''
        }
        
        if not appui_geom:
            entry['statut'] = 'ERREUR'
            entry['bpe_noe_type'] = 'appui non localisé'
            result[num_appui] = entry
            continue
        
        bpe_proche = None
        dist_min = 999999
        for bpe in bpe_geoms:
            dist = appui_geom.distance(bpe['geom'])
            if dist < tolerance and dist < dist_min:
                dist_min = dist
                bpe_proche = bpe
        
        if bpe_proche:
            entry['bpe_trouve'] = True
            entry['bpe_noe_type'] = bpe_proche['noe_type']
            entry['statut'] = 'OK'
        else:
            entry['statut'] = 'ERREUR'
        
        result[num_appui] = entry
    
    return result
