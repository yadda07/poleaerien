# INSTALLATION DU PLUGIN PÔLE AÉRIEN

## ⚠️ IMPORTANT : NOM DU DOSSIER

Le plugin **DOIT** être installé dans un dossier nommé **`PoleAerien`** (sans espaces, sans accents).

### ❌ ERREUR FRÉQUENTE

Si vous voyez cette erreur :
```
ModuleNotFoundError: No module named 'new/PoleAerien'
```

C'est que le dossier du plugin s'appelle **`new`** au lieu de **`PoleAerien`**.

---

## 📦 INSTALLATION MANUELLE

### Étape 1 : Localiser le dossier des plugins QGIS

Le chemin dépend de votre système d'exploitation :

**Windows** :
```
C:\Users\[VotreNom]\AppData\Roaming\QGIS\QGIS3\profiles\default\python\plugins\
```

**Linux** :
```
~/.local/share/QGIS/QGIS3/profiles/default/python/plugins/
```

**macOS** :
```
~/Library/Application Support/QGIS/QGIS3/profiles/default/python/plugins/
```

### Étape 2 : Créer le dossier PoleAerien

1. Allez dans le dossier `plugins/`
2. Créez un dossier nommé **exactement** `PoleAerien` (respectez la casse)
3. Copiez **tous les fichiers** du plugin dans ce dossier

### Étape 3 : Vérifier la structure

Le dossier doit contenir :
```
PoleAerien/
├── __init__.py
├── PoleAerien.py
├── metadata.txt
├── workflows/
│   ├── __init__.py
│   ├── maj_workflow.py
│   ├── comac_workflow.py
│   ├── capft_workflow.py
│   ├── c6bd_workflow.py
│   ├── c6c3a_workflow.py
│   └── police_workflow.py
├── qgis_utils.py
├── core_utils.py
├── async_tasks.py
├── (autres fichiers...)
└── images/
```

### Étape 4 : Redémarrer QGIS

1. Fermez complètement QGIS
2. Relancez QGIS
3. Allez dans **Extensions → Gérer et installer les extensions**
4. Cherchez "Pôle Aérien" dans l'onglet **Installées**
5. Cochez la case pour activer le plugin

---

## 🔧 INSTALLATION VIA ZIP

### Étape 1 : Créer le ZIP

1. Créez un dossier nommé **`PoleAerien`**
2. Copiez tous les fichiers du plugin dedans
3. Compressez le dossier `PoleAerien` en **`PoleAerien.zip`**

**Important** : Le ZIP doit contenir le dossier `PoleAerien/`, pas directement les fichiers !

Structure correcte du ZIP :
```
PoleAerien.zip
└── PoleAerien/
    ├── __init__.py
    ├── PoleAerien.py
    ├── metadata.txt
    └── (autres fichiers...)
```

### Étape 2 : Installer via QGIS

1. Dans QGIS : **Extensions → Gérer et installer les extensions**
2. Onglet **Installer depuis un ZIP**
3. Sélectionnez `PoleAerien.zip`
4. Cliquez sur **Installer l'extension**

---

## ✅ VÉRIFICATION

Après installation, vérifiez que :

1. Le dossier s'appelle bien `PoleAerien` (pas `new`, `PoleAerien-master`, etc.)
2. Le fichier `metadata.txt` existe
3. Le fichier `__init__.py` contient la fonction `classFactory()`

---

## 🐛 DÉPANNAGE

### Erreur : "No module named 'new/PoleAerien'"

**Cause** : Le dossier du plugin ne s'appelle pas `PoleAerien`

**Solution** :
1. Allez dans le dossier `plugins/`
2. Renommez le dossier en **`PoleAerien`** (exactement)
3. Redémarrez QGIS

### Erreur : "Plugin broken"

**Cause** : Fichiers manquants ou structure incorrecte

**Solution** :
1. Vérifiez que tous les fichiers sont présents
2. Vérifiez que le dossier `workflows/` existe
3. Vérifiez que `__init__.py` existe à la racine

### Le plugin n'apparaît pas dans la liste

**Cause** : `metadata.txt` invalide ou manquant

**Solution** :
1. Vérifiez que `metadata.txt` existe
2. Vérifiez que la version QGIS est >= 3.28
3. Consultez les logs QGIS : **Extensions → Console Python**

---

## 📋 PRÉREQUIS

- **QGIS** : Version 3.28 minimum (testé jusqu'à 3.42)
- **Python** : 3.9+ (inclus avec QGIS)
- **Dépendances** : pandas, openpyxl (installées automatiquement par QGIS)

---

## 📞 SUPPORT

En cas de problème :
1. Consultez les logs QGIS : **Extensions → Console Python**
2. Vérifiez le nom du dossier (doit être `PoleAerien`)
3. Contactez : yadda@nge-es.fr
