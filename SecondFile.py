#!/usr/bin/python
# -*-coding:Utf-8-*-

# Second pour contenir quelques fonctions qui seront utilisées dans le fichier principale Police.py
from qgis.utils import iface
from qgis.core import Qgis, QgsProject, QgsExpression, QgsFeatureRequest, QgsVectorLayer
from .qgis_utils import remove_group, layer_group_error, temps_ecoule


class SecondFile:
    """ Classe qui fait tout le travail pour accéder à la base de données,
    Réaliser des requêtes dans la base de données,
    Exporter le résultat des requêtes vers le fichier Excel """

    def __init__(self):
        """Le constructeur de ma classe
        Il prend pour attribut de classe les *** """

        self.compteGoup = 0

    def alerteInfo(self, monMessage, duree=5):
        """Fonction pour afficher un message d'information à destination de l'utilisateur"""
        iface.messageBar().pushMessage("Message : ", monMessage, level=Qgis.Info, duration=duree)

    def alerteCritique(self, monMessage, duree=10):
        """Fonction pour afficher des messages d'erreur à destination de l'utilisateur"""
        iface.messageBar().pushMessage("Erreur : ", monMessage, level=Qgis.Critical, duration=duree)

    def createNewLayer(self, colonne, condition, table, geom, name, error=""):
        """Crée une nouvelle couche mémoire avec les features filtrées."""
        lyrs = QgsProject.instance().mapLayersByName(table)
        if not lyrs:
            return
        layer = lyrs[0]
        if not layer.isValid():
            return
        # SEC-04: condition est un tuple passé directement - sûr
        requete = QgsExpression(f"{colonne} IN {condition}")
        feats = (feat for feat in layer.getFeatures(QgsFeatureRequest(requete)))
        epsg = layer.crs().postgisSrid()  # SRID

        # geom peut être : Point, Polygon, Polyline
        couche = QgsVectorLayer(f"{geom}?crs=epsg:{epsg}", f"error_{table}_{name}", "memory")
        couche_data = couche.dataProvider()
        attr = layer.dataProvider().fields().toList()
        couche_data.addAttributes(attr)
        couche.updateFields()
        couche_data.addFeatures(feats)

        # Update extent of the layer
        couche.updateExtents()

        QgsProject.instance().addMapLayer(couche, False)
        self.layerGroupError(couche, error)

    def tempsEcouler(self, seconde):
        return temps_ecoule(seconde)

    def layerGroupError(self, couche, nomEtude):
        """Fonction qui permet d'ajout une couche dans un groupe d'erreur"""
        layer_group_error(couche, nomEtude)

    def removeGroup(self, name):
        """Suppression de groupe de couche du nom de GraceTHD"""
        remove_group(name)

