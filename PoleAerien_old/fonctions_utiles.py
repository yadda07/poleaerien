# -*- coding: utf-8 -*-

from qgis.core import QgsSpatialIndex

class FonctionsUtiles:

    def __init__(self):
        pass

    # Créer un index spatial
    def create_qgsspatialindex(self, layer):
        index = QgsSpatialIndex()
        for feature in layer.getFeatures():
            index.insertFeature(feature)
        return index


fonctions_utiles = FonctionsUtiles()