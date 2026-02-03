import os
import webbrowser
from qgis.PyQt import QtWidgets

from .PoleAerien_apropos import Ui_AboutDialog


# create dialog for about/help window
class PoleAerienAboutDialog(QtWidgets.QDialog):
    def __init__(self):
        QtWidgets.QDialog.__init__(self)
        # Set up the user interface from Designer.
        self.uiAbout = Ui_AboutDialog()
        self.uiAbout.setupUi(self)
        
        # Connect link click to handler
        self.uiAbout.textBrowser_2.anchorClicked.connect(self.handleLinkClick)
        self.uiAbout.textBrowser_2.setOpenLinks(False)
    
    def handleLinkClick(self, url):
        """Open local docs or external links"""
        link = url.toString()
        if link.startswith("docs/"):
            # Local doc - build absolute path
            plugin_dir = os.path.dirname(__file__)
            doc_path = os.path.join(plugin_dir, link)
            if os.path.exists(doc_path):
                webbrowser.open(f"file:///{doc_path}")
        else:
            # External link
            webbrowser.open(link)
