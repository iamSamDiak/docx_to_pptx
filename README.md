# DOCX to PPTX Converter

Convertisseur de documents Word (.docx) en présentations PowerPoint (.pptx) avec interface graphique moderne.

## 🚀 Build sur Windows

### Prérequis
- Python 3.10+ installé depuis [python.org](https://www.python.org/downloads/)
- Chocolatey installé (pour Inno Setup)

### Instructions

1. **Cloner le projet**
```powershell
git clone https://github.com/iamSamDiak/docx_to_pptx.git
cd docx_to_pptx
```

2. **Lancer le build automatique**
```powershell
.\build.ps1
```

Ou manuellement :
```powershell
# Installer les dépendances
pip install -r requirements.txt
pip install pyinstaller

# Créer l'exécutable
pyinstaller --onefile --windowed --name="DOCX to PPTX" --icon=icons/app.ico --collect-all PyQt5 --collect-all docx --collect-all pptx --collect-all lxml main.py

# Installer Inno Setup
choco install innosetup -y

# Créer l'installeur
iscc setup.iss
```

3. **Récupérer l'installeur**
Le fichier `DOCX_to_PPTX_Setup.exe` sera dans le dossier `installer/`

## 📦 Fichiers essentiels

- `main.py`, `gui.py`, `convert.py` : Code source
- `requirements.txt` : Dépendances Python
- `icons/app.ico` : Icône de l'application
- `build.ps1` / `build.bat` : Scripts de build
- `setup.iss` : Configuration Inno Setup

## 📝 Licence

Projet open-source.
