#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de test pour la conversion DOCX -> PPTX
"""

import sys
import os

# Ajoute le dossier parent au path pour importer depuis src/
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.convert import to_powerpoint
from src.handlers import NoHeadlineError

def test_conversion():
    """Test de conversion d'un document Word"""
    print("=== Test de conversion DOCX -> PPTX ===\n")
    
    # Chemin du fichier de test
    docx_file = os.path.join(os.path.dirname(__file__), "test - Copie.docx")
    pptx_file = os.path.join(os.path.dirname(__file__), "test.pptx")
    
    # Conversion
    pw = to_powerpoint()
    pw.open(docx_file)
    pw.to_pptx()
    
    # Sauvegarde avec chemin personnalisé
    pw.pptx.save(pptx_file)
    print(f"✅ PowerPoint exporté : {pptx_file}\n")
        
    return True

if __name__ == "__main__":
    success = test_conversion()
    sys.exit(0 if success else 1)
