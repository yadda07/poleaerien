# -*- coding: utf-8 -*-
"""
/***************************************************************************
 Pôle Aérien
                                 A QGIS plugin
 Contrôle qualité et mise à jour des poteaux aériens ENEDIS (FT/BT)
                            -------------------
        begin                : 2022-03-01
        copyright            : (C) 2022-2026 by NGE ES
        email                : contact@nge-es.fr
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""


class Initialisation:
    """Classe de métadonnées du plugin"""

    def __init__(self):
        pass

    def name(self):
        return "Pôle Aérien"

    def version(self):
        return "2.3.0"

    def description(self):
        return "Contrôle qualité et mise à jour des poteaux aériens ENEDIS (FT/BT) pour projets FTTH"

    def qgisMinimumVersion(self):
        return "3.28"

    def qgisMaximumVersion(self):
        return "3.99"

    def experimental(self):
        return False

    def author(self):
        return "NGE ES"

    def authorName(self):
        return self.author()

    def email(self):
        return "yadda@nge-es.fr"

    def icon(self):
        return "images/icon.svg"


# noinspection PyPep8Naming
def classFactory(iface):  # pylint: disable=invalid-name
    """Load PoleAerien class from file PoleAerien.

    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
    from .PoleAerien import PoleAerien
    return PoleAerien(iface)
