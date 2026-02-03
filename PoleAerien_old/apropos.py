# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'check_optyce_apropos.ui'
#
# Created by: PyQt4 UI code generator 4.11.4
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)


class Ui_AboutDialog(object):
    def setupUi(self, AboutDialog):
        AboutDialog.setObjectName(_fromUtf8("AboutDialog"))
        AboutDialog.resize(751, 628)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(AboutDialog.sizePolicy().hasHeightForWidth())
        AboutDialog.setSizePolicy(sizePolicy)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(_fromUtf8("icon.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        AboutDialog.setWindowIcon(icon)
        self.gridLayout = QtGui.QGridLayout(AboutDialog)
        self.gridLayout.setObjectName(_fromUtf8("gridLayout"))
        self.verticalLayout = QtGui.QVBoxLayout()
        self.verticalLayout.setObjectName(_fromUtf8("verticalLayout"))
        self.horizontalLayout_3 = QtGui.QHBoxLayout()
        self.horizontalLayout_3.setObjectName(_fromUtf8("horizontalLayout_3"))
        self.about = QtGui.QLabel(AboutDialog)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.about.sizePolicy().hasHeightForWidth())
        self.about.setSizePolicy(sizePolicy)
        self.about.setMinimumSize(QtCore.QSize(120, 50))
        self.about.setMaximumSize(QtCore.QSize(91, 25))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.about.setFont(font)
        self.about.setObjectName(_fromUtf8("about"))
        self.horizontalLayout_3.addWidget(self.about)
        self.logo = QtGui.QLabel(AboutDialog)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.logo.sizePolicy().hasHeightForWidth())
        self.logo.setSizePolicy(sizePolicy)
        self.logo.setMinimumSize(QtCore.QSize(70, 79))
        self.logo.setMaximumSize(QtCore.QSize(70, 79))
        self.logo.setObjectName(_fromUtf8("logo"))
        self.horizontalLayout_3.addWidget(self.logo)
        self.logo_2 = QtGui.QLabel(AboutDialog)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.logo_2.sizePolicy().hasHeightForWidth())
        self.logo_2.setSizePolicy(sizePolicy)
        self.logo_2.setMinimumSize(QtCore.QSize(70, 79))
        self.logo_2.setMaximumSize(QtCore.QSize(70, 79))
        self.logo_2.setObjectName(_fromUtf8("logo_2"))
        self.horizontalLayout_3.addWidget(self.logo_2)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.horizontalLayout = QtGui.QHBoxLayout()
        self.horizontalLayout.setObjectName(_fromUtf8("horizontalLayout"))
        self.version = QtGui.QLabel(AboutDialog)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.version.sizePolicy().hasHeightForWidth())
        self.version.setSizePolicy(sizePolicy)
        self.version.setMinimumSize(QtCore.QSize(90, 25))
        self.version.setMaximumSize(QtCore.QSize(90, 25))
        self.version.setAlignment(QtCore.Qt.AlignCenter)
        self.version.setObjectName(_fromUtf8("version"))
        self.horizontalLayout.addWidget(self.version)
        self.version_n = QtGui.QLabel(AboutDialog)
        self.version_n.setMaximumSize(QtCore.QSize(136, 89))
        self.version_n.setText(_fromUtf8(""))
        self.version_n.setObjectName(_fromUtf8("version_n"))
        self.horizontalLayout.addWidget(self.version_n)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtGui.QHBoxLayout()
        self.horizontalLayout_2.setObjectName(_fromUtf8("horizontalLayout_2"))
        self.autors = QtGui.QLabel(AboutDialog)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.autors.sizePolicy().hasHeightForWidth())
        self.autors.setSizePolicy(sizePolicy)
        self.autors.setMinimumSize(QtCore.QSize(90, 25))
        self.autors.setMaximumSize(QtCore.QSize(90, 25))
        self.autors.setObjectName(_fromUtf8("autors"))
        self.horizontalLayout_2.addWidget(self.autors)
        self.autors_name = QtGui.QLabel(AboutDialog)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.autors_name.sizePolicy().hasHeightForWidth())
        self.autors_name.setSizePolicy(sizePolicy)
        self.autors_name.setMinimumSize(QtCore.QSize(136, 89))
        self.autors_name.setMaximumSize(QtCore.QSize(136, 89))
        self.autors_name.setObjectName(_fromUtf8("autors_name"))
        self.horizontalLayout_2.addWidget(self.autors_name)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.textBrowser_3 = QtGui.QTextBrowser(AboutDialog)
        self.textBrowser_3.setEnabled(True)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textBrowser_3.sizePolicy().hasHeightForWidth())
        self.textBrowser_3.setSizePolicy(sizePolicy)
        self.textBrowser_3.setMinimumSize(QtCore.QSize(263, 121))
        self.textBrowser_3.setMaximumSize(QtCore.QSize(261, 121))
        self.textBrowser_3.setLineWidth(2)
        self.textBrowser_3.setOpenExternalLinks(True)
        self.textBrowser_3.setObjectName(_fromUtf8("textBrowser_3"))
        self.verticalLayout.addWidget(self.textBrowser_3)
        self.gridLayout.addLayout(self.verticalLayout, 0, 0, 1, 1)
        self.verticalLayout_2 = QtGui.QVBoxLayout()
        self.verticalLayout_2.setObjectName(_fromUtf8("verticalLayout_2"))
        self.label_help = QtGui.QLabel(AboutDialog)
        self.label_help.setMinimumSize(QtCore.QSize(0, 25))
        self.label_help.setMaximumSize(QtCore.QSize(16777215, 25))
        self.label_help.setObjectName(_fromUtf8("label_help"))
        self.verticalLayout_2.addWidget(self.label_help)
        self.textBrowser_2 = QtGui.QTextBrowser(AboutDialog)
        font = QtGui.QFont()
        font.setBold(False)
        font.setWeight(50)
        self.textBrowser_2.setFont(font)
        self.textBrowser_2.setObjectName(_fromUtf8("textBrowser_2"))
        self.verticalLayout_2.addWidget(self.textBrowser_2)
        self.buttonBox = QtGui.QDialogButtonBox(AboutDialog)
        self.buttonBox.setStandardButtons(QtGui.QDialogButtonBox.Close)
        self.buttonBox.setObjectName(_fromUtf8("buttonBox"))
        self.verticalLayout_2.addWidget(self.buttonBox)
        self.gridLayout.addLayout(self.verticalLayout_2, 0, 1, 1, 1)

        self.retranslateUi(AboutDialog)
        QtCore.QObject.connect(self.buttonBox, QtCore.SIGNAL(_fromUtf8("rejected()")), AboutDialog.reject)
        QtCore.QMetaObject.connectSlotsByName(AboutDialog)

    def retranslateUi(self, AboutDialog):
        AboutDialog.setWindowTitle(_translate("AboutDialog", "C3A Plugin: Info", None))
        self.about.setText(_translate("AboutDialog", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'Sans Serif\'; font-size:12pt; font-weight:600; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:20pt;\">About</span></p></body></html>", None))
        self.logo.setText(_translate("AboutDialog", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'Sans Serif\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><img src=\":/plugins/GenerateurC3A/images/logo_NGE.png\" /></p></body></html>", None))
        self.logo_2.setText(_translate("AboutDialog", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'Sans Serif\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><img src=\":/plugins/GenerateurC3A/images/icon_grand.png\" /></p></body></html>", None))
        self.version.setText(_translate("AboutDialog", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'Sans Serif\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt; font-weight:600;\">Version :</span></p></body></html>", None))
        self.autors.setText(_translate("AboutDialog", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'Sans Serif\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt; font-weight:600;\">Authors :</span></p></body></html>", None))
        self.autors_name.setText(_translate("AboutDialog", "NGE Infranet, \n"
"HACHANA Mohamed, \n"
"SOUMARE Abdoulayi, \n"
"TIEN Nguyen.", None))
        self.textBrowser_3.setHtml(_translate("AboutDialog", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.5pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:9pt;\">More info</span></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><a href=\"https://leportail.nge.fr/\"><span style=\" font-size:8pt; text-decoration: underline; color:#0000ff;\">https://leportail.nge.fr/</span></a></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><a href=\"https://www.nge.fr/filiales/nge-infranet\"><span style=\" font-size:8pt; text-decoration: underline; color:#0000ff;\">https://www.nge.fr/filiales/nge-infranet</span></a></p>\n"
"<p align=\"center\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:8pt; text-decoration: underline; color:#0000ff;\"><br /></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:9pt;\">You can add an issue or a bug</span></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><a href=\"asoumare@nge.net\"><span style=\" font-size:8pt; text-decoration: underline; color:#0000ff;\">asoumare@nge.net</span></a></p></body></html>", None))
        self.label_help.setText(_translate("AboutDialog", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'Sans Serif\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt; font-weight:600;\">How to use</span></p></body></html>", None))
        self.textBrowser_2.setHtml(_translate("AboutDialog", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.5pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600;\">C3A </span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">est un outil développé par NGE infranet pour répondre à la réalisation des commandes d\'accès Orange. Il s\'appuie sur de nombreuses données que sont GraceTHD et les données Annexes C7.</span></p>\n"
"<p align=\"justify\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'Sans Serif\'; font-size:10pt;\"><br /></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600;\">1. Importer</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> : la première partie de l\'outil importe des données utiles à la réalisation du C3A. L\'utilisateur doit simplement indiquer le répertoire dans lequel se trouve les données </span><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-style:italic; text-decoration: underline;\">GraceTHD</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> et l\'outil se charge de choisir les données qu\'il a besoin : </span><span style=\" font-family:\'Consolas\'; font-size:8pt; color:#000000;\">t_cheminement,</span><span style=\" font-family:\'Consolas\'; font-size:8pt; color:#cc7832;\"> </span><span style=\" font-family:\'Consolas\'; font-size:8pt; color:#000000;\">t_cableline, t_noeud, t_cab_cond, t_cond_chem, t_ebp, t_cable, t_ptech, t_conduite.</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> </span></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-style:italic; text-decoration: underline;\">En bonus</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600; font-style:italic;\"> :</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> un style est appliqué automatiquement aux données lors de leur importation dans QGIS. Ces styles peuvent être remplacés par vos propres styles en remplaçant les fichiers de styles situés ici : </span></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">C:\\Users\\</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-style:italic;\">votreNOM</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">\\.qgis2\\python\\plugins\\GenerateurC3A\\style\\</span></p>\n"
"<p align=\"justify\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'Sans Serif\'; font-size:10pt;\"><br /></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600;\">2. Générer C3A </span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">: après importation des données, la réalisation du fichier C3A en cliquant sur le bouton indiqué. Il ne vous restera plus qu\'attendre la fin de ce proccesus qui peut nécessiter 2 et 4mn en fonction de la quantité des données. La barre de progression est là pour indiquer la progression en cours. </span><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-style:italic; text-decoration: underline;\">Attention</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> : la barre de progression se bloque lorsqu\'on essaye de faire quoi que ce soit en parallèle sur le PC. Toutefois, cela n\'est rien de grave puisque le programme continuera ses tâches en arrière-plan et vous serait averti lorsqu\'il attendra ses objectifs.</span></p>\n"
"<p align=\"justify\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'Sans Serif\'; font-size:10pt;\"><br /></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600;\">3. Appuis à remplacer</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> : les données C3A générées ne tiennent pas compte des appuis à remplacer. Pour cela, il a besoin des données Orange des fichiers Annexes C7. A nouveau, vous devrez indiquer le répertoire où se trouve les fichiers, et l\'outil se charge de trouver les appuis a remplacer.</span></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">Il se base uniquement sur les fichiers Excel de format &quot;</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-style:italic; text-decoration: underline;\">xlsx</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">&quot;. Veuillez à ce que votre répertoire ne contient pas d\'autres fichiers Excel &quot;xlsx&quot;. Sinon des erreurs peuvent se produire (voire plus bas).</span></p>\n"
"<p align=\"justify\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'Sans Serif\'; font-size:10pt;\"><br /></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600;\">4. Exporter</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> : les exports des données se font directement dans un fichier modèle d\'Orange &quot;</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600; font-style:italic; color:#ffaa00;\">Annexe C3 a</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">&quot;. Les exportations répondent à 3 exigences fondamentales :</span></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">- Maximum de 5 communes par (commande) fichier</span></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">- Aucun câble divisé entre deux commandes distinctes</span></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">- Maximum 500 lignes par commandes.</span></p>\n"
"<p align=\"justify\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'Sans Serif\'; font-size:10pt;\"><br /></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-style:italic;\">Distribution</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> : Exporter les données concernant les cables de distribution.</span></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-style:italic;\">Transport</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> : Exporter les données concernant les cables de transport.</span></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">L\'un des deux doit au moins être coché lors de l\'exportation des données commandes C3A.</span></p>\n"
"<p align=\"justify\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'Sans Serif\'; font-size:10pt;\"><br /></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600; text-decoration: underline;\">Projection</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> : le système de projection utilisé est le RGF-93. EPSG : 2154.</span></p>\n"
"<p align=\"justify\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'Sans Serif\'; font-size:10pt;\"><br /></p>\n"
"<p align=\"justify\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'Sans Serif\'; font-size:10pt;\"><br /></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600; font-style:italic; text-decoration: underline; color:#ff0000;\">Erreurs connues</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt; color:#ff0000;\"> </span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">:</span></p>\n"
"<p align=\"justify\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'Sans Serif\'; font-size:10pt;\"><br /></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600;\">1.</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> Appuis à remplacer : Lorsqu\'on met à jour les informations concernant les appuis à remplacer, l\'outil se base uniquement sur les fichiers de format &quot;</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600;\">xlsx</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">&quot;. Si d\'autres fichiers de ce format existent dans le même répertoire, des erreurs peuvent se produire. Veuillez éviter cette situation, sinon l\'erreur suivante peut se produire : ValueError: invalid literal for </span><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600;\">int() </span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">with base 10.</span></p>\n"
"<p align=\"justify\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'Sans Serif\'; font-size:10pt;\"><br /></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600;\">2.</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\"> &quot;No module Name </span><span style=\" font-family:\'Sans Serif\'; font-size:10pt; font-weight:600;\">Openpyxl</span><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">&quot;: Module Openpyxl n\'est pas installé par défaut dans QGIS et il est nécessaire de l\'installer.</span></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">Pour répondre à cette problématique le dossier &quot;modules_a_copier vous a été préparé en avance.</span></p>\n"
"<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'Sans Serif\'; font-size:10pt;\">Copier tous les contenus de ce dossier et le copier dans le répertoire python de votre QGIS qui se trouve à cet emplacement : C:\\Program Files\\QGIS 2.18 \\bin\\</span></p></body></html>", None))

