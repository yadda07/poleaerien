#!/usr/bin/python
# -*-coding:Utf-8-*-

# Second pour contenir quelques fonctions qui seront utilisées dans le fichier principale Police.py
from qgis.utils import iface
from qgis.core import Qgis, QgsProject, QgsExpression, QgsFeatureRequest, QgsVectorLayer, QgsFeature, QgsLayerTreeLayer
import os
import re


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

    def alerteAvertissement(self, monMessage, duree=10):
        """Fonction pour afficher des messages d'avertissements à destination de l'utilisateur"""
        iface.messageBar().pushMessage(" Attention : ", monMessage, level=Qgis.Warning, duration=duree)

    def createNewLayer(self, colonne, condition, table, geom, name, error=""):
        """Fonctions qui permet de Créer une nouvelle couche """
        mesCouches = QgsProject.instance().mapLayers()
        layer = ""
        for coucheId, couche in mesCouches.items():
            if couche.name() == table:
                layer = couche
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
        seconds = seconde % (24 * 3600)
        hour = int(seconds // 3600)
        seconds %= 3600
        minutes = int(seconds // 60)
        seconds %= 60

        if hour > 0:
            #  "%heure:%minutes:%seconde" % (hour, minutes, seconds)
            return f"{hour}h: {int(minutes)}mn : {int(seconds)}sec"

        else:
            return f"{minutes}mn : {int(seconds)}sec"

    def barreP(self, numero):
        """ Pour faire appel à la barre de progression """
        return u"self.dlg.progressBar.setValue({})".format(numero)

    def insertLayerInGroupGraceTHD(self, couche):
        """Fonction qui permet d'ajout une couche dans un groupe d'erreur"""

        root = QgsProject.instance().layerTreeRoot()
        group = root.findGroup(u"GRACETHD")

        if not group:
            groupName = u"GRACETHD"
            root = QgsProject.instance().layerTreeRoot()
            group = root.addGroup(groupName)

        else:
            pass

        group.insertChildNode(1, QgsLayerTreeLayer(couche))

    def layerGroupError(self, couche, nomEtude):
        """Fonction qui permet d'ajout une couche dans un groupe d'erreur"""
        root = QgsProject.instance().layerTreeRoot()
        group = root.findGroup(u'ERROR_' + str(nomEtude))

        if not group:
            groupName = u'ERROR_' + str(nomEtude)
            root = QgsProject.instance().layerTreeRoot()
            group = root.addGroup(groupName)

        else:
            pass

        group.insertChildNode(1, QgsLayerTreeLayer(couche))

    def removeGroup(self, name):
        """Suppression de groupe de couche du nom de GraceTHD"""
        root = QgsProject.instance().layerTreeRoot()
        group = root.findGroup(name)
        if not group is None:
            for child in group.children():
                dump = child.dump()
                id = dump.split("=")[-1].strip()
                QgsProject.instance().removeMapLayer(id)
            root.removeChildNode(group)

    def appliquerstyle(self, nomcouche):
        """Fonction qui permet d'appliquer un style au fichier erreur qui sera généré """
        # On applique un style au fichier C3A qui aurait été généré
        cheminstyle = ''

        cheminAbsolu = "~/AppData/Roaming/QGIS/QGIS3/profiles/default/python/plugins/PoliceC6/styles/"

        if re.match(u'error_infra_pt_pot_', nomcouche):
            cheminstyle = os.path.expanduser(cheminAbsolu + "infra_pt_pot.qml")

        if re.match(u'error_infra_pt_pot_ebp', nomcouche):
            cheminstyle = os.path.expanduser(cheminAbsolu + "infra_pt_pot_ebp.qml")

        if re.match(u'error_appui_capa_appui', nomcouche):
            cheminstyle = os.path.expanduser(cheminAbsolu + "t_cable.qml")

        if re.match(u'error_bpe', nomcouche):
            cheminstyle = os.path.expanduser(cheminAbsolu + "bpe.qml")

        verifstyle = True
        self.messagestyle = ''
        # On vérifie si le fichier contenant le style existe pour l'appliquer au fichier C3A

        if verifstyle:
            mesCouches = QgsProject.instance().mapLayers()
            for coucheId, couche in mesCouches.items():
                if couche.name() == nomcouche:
                    try:
                        with open(cheminstyle):
                            couche.loadNamedStyle(cheminstyle)
                            self.messagestyle = u"Le style a bien été appliquée au fichier C3A"
                            # pass
                    except IOError:
                        verifstyle = False
                        self.messagestyle = u"Le style n'a pu été trouvé. A faire manuellement"
