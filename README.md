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
.\scripts\build.ps1
```

Ou manuellement :
```powershell
# Installer les dépendances
pip install -r requirements.txt
pip install pyinstaller

# Créer l'exécutable
pyinstaller --onefile --windowed --name="DOCX to PPTX" --icon=assets/app.ico --collect-all PyQt5 --collect-all docx --collect-all pptx --collect-all lxml src/main.py

# Installer Inno Setup (si non installé)
choco install innosetup -y

# Créer l'installeur
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup.iss
```

3. **Récupérer l'installeur**
Le fichier `DOCX_to_PPTX_Setup.exe` sera dans le dossier `installer/`

## 📦 Fichiers essentiels

- `src/` : Code source
  - `main.py` : Point d'entrée
  - `gui.py` : Interface PyQt5
  - `convert.py` : Logique de conversion
- `assets/` : Ressources (icône)
- `scripts/` : Scripts de build
  - `build.ps1` : Build PowerShell
  - `build.bat` : Build CMD
- `requirements.txt` : Dépendances Python
- `setup.iss` : Configuration Inno Setup

## 📝 Licence

Projet open-source.
