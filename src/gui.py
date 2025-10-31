from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLineEdit, QFileDialog, 
                             QMessageBox, QLabel, QStyle, QMenuBar, QAction, 
                             QDialog, QDialogButtonBox, QFormLayout, QSpinBox)
from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtGui import QFont, QIcon
import sys
import io
from contextlib import redirect_stdout

try:
    from src.convert import to_powerpoint
except ModuleNotFoundError:
    from convert import to_powerpoint

if getattr(sys, 'frozen', False):
    import pyi_splash

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon("assets/app.png"))
        self.setWindowTitle("EcoDim")
        self.resize(600, 350)
        
        # Initialiser QSettings pour la persistance des paramètres
        self.settings = QSettings("EcoDim", "headline_size")
        
        # Charger les paramètres sauvegardés (ou utiliser la valeur par défaut)
        self.headline_size = float(self.settings.value("headline_size", 16.0))
        
        # Créer la barre de menu
        self.create_menu_bar()
        
        # Style moderne inspiré du design Google Calendar
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
            }
            QWidget {
                background-color: #ffffff;
                font-family: 'Segoe UI', 'Arial', sans-serif;
            }
        """)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal avec marges
        layout = QVBoxLayout()
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)
        central_widget.setLayout(layout)
        
        # Première rangée : bouton + champ de texte pour le chemin
        first_row = QHBoxLayout()
        first_row.setSpacing(10)
        
        self.browse_btn = QPushButton()
        # Utilise l'icône système de dossier
        folder_icon = self.style().standardIcon(QStyle.SP_DirIcon)
        self.browse_btn.setIcon(folder_icon)
        self.browse_btn.setFixedSize(50, 45)
        self.browse_btn.setCursor(Qt.PointingHandCursor)
        self.browse_btn.setStyleSheet("""
            QPushButton {
                background-color: #f1f3f4;
                border: 1px solid #dadce0;
                border-radius: 8px;
                font-size: 18px;
                padding: 8px;
            }
            QPushButton:hover {
                background-color: #e8eaed;
            }
            QPushButton:pressed {
                background-color: #dadce0;
            }
        """)
        self.browse_btn.clicked.connect(self.browse_file)
        
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText("Sélectionnez un document Word...")
        self.path_input.setFixedHeight(45)
        self.path_input.setStyleSheet("""
            QLineEdit {
                background-color: #ffffff;
                border: 1px solid #dadce0;
                border-radius: 8px;
                padding: 10px 15px;
                font-size: 14px;
                color: #202124;
            }
            QLineEdit:focus {
                border: 2px solid #1a73e8;
            }
        """)
        self.path_input.textChanged.connect(self.update_button_states)
        
        first_row.addWidget(self.browse_btn)
        first_row.addWidget(self.path_input)
        layout.addLayout(first_row)
        
        # Label d'information sur les paramètres de détection
        self.info_label = QLabel()
        self.info_label.setWordWrap(True)
        self.info_label.setAlignment(Qt.AlignCenter)
        self.update_info_label()  # Initialise le texte
        self.info_label.setStyleSheet("""
            QLabel {
                background-color: #fff;
                border-radius: 8px;
                padding: 12px;
                font-size: 12px;
                color: #000;
                border: 1px solid #000;
            }
        """)
        layout.addWidget(self.info_label)
        
        # Deuxième rangée : bouton de conversion
        self.convert_btn = QPushButton("Convertir en PowerPoint")
        self.convert_btn.setFixedHeight(50)
        self.convert_btn.setCursor(Qt.PointingHandCursor)
        self.convert_btn.setStyleSheet("""
            QPushButton {
                background-color: #34a853;
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 15px;
                font-weight: 500;
                padding: 12px;
            }
            QPushButton:hover {
                background-color: #2d9248;
            }
            QPushButton:pressed {
                background-color: #268042;
            }
            QPushButton:disabled {
                background-color: #e8eaed;
                color: #9aa0a6;
            }
        """)
        self.convert_btn.clicked.connect(self.convert_file)
        self.convert_btn.setEnabled(False)
        layout.addWidget(self.convert_btn)
        
        # Label pour les messages/logs
        self.log_label = QLabel("")
        self.log_label.setWordWrap(True)
        self.log_label.setAlignment(Qt.AlignCenter)
        self.log_label.setMinimumHeight(110)
        self.log_label.setStyleSheet("""
            QLabel {
                background-color: #f8f9fa;
                border-radius: 8px;
                padding: 15px;
                font-size: 13px;
                color: #5f6368;
            }
        """)
        layout.addWidget(self.log_label)
        
        # Troisième rangée : bouton d'export
        self.export_btn = QPushButton("Exporter PowerPoint")
        self.export_btn.setFixedHeight(50)
        self.export_btn.setCursor(Qt.PointingHandCursor)
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #3c4043;
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 15px;
                font-weight: 500;
                padding: 12px;
            }
            QPushButton:hover {
                background-color: #2d2f31;
            }
            QPushButton:pressed {
                background-color: #1f2021;
            }
            QPushButton:disabled {
                background-color: #e8eaed;
                color: #9aa0a6;
            }
        """)
        self.export_btn.clicked.connect(self.export_file)
        self.export_btn.setEnabled(False)
        layout.addWidget(self.export_btn)
        
        # Variable pour stocker l'objet Powerpoint
        self.pptx_obj = None
    
    def update_info_label(self):
        """Met à jour le label d'information avec les paramètres actuels"""
        self.info_label.setText(
            f"ℹ️ Détection : Titres (police ≥ {int(self.headline_size)} pt) • Paragraphes (police < {int(self.headline_size)} pt)"
        )
    
    def create_menu_bar(self):
        """Crée la barre de menu"""
        menubar = self.menuBar()
        
        # Menu Paramètres qui ouvre directement la boîte de dialogue
        settings_action = menubar.addAction("Paramètres")
        settings_action.triggered.connect(self.open_settings)
    
    def open_settings(self):
        """Ouvre la boîte de dialogue des paramètres"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Paramètres")
        dialog.setFixedSize(400, 150)
        
        # Layout du formulaire
        layout = QFormLayout()
        
        # Champ pour la taille de police des titres
        headline_spinbox = QSpinBox()
        headline_spinbox.setMinimum(10)
        headline_spinbox.setMaximum(50)
        headline_spinbox.setValue(int(self.headline_size))
        headline_spinbox.setSuffix(" pt")
        headline_spinbox.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                font-size: 13px;
                border: 1px solid #dadce0;
                border-radius: 4px;
            }
        """)
        
        layout.addRow("Taille minimum police : Titre", headline_spinbox)
        
        # Boutons OK / Annuler
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        dialog.setLayout(layout)
        
        # Si l'utilisateur clique sur OK, met à jour les paramètres
        if dialog.exec_() == QDialog.Accepted:
            self.headline_size = float(headline_spinbox.value())
            self.settings.setValue("headline_size", self.headline_size)  # Sauvegarde
            self.update_info_label()  # Met à jour le label d'information
    
    def update_button_states(self):
        """Met à jour l'état des boutons selon le contenu du champ de texte"""
        has_path = bool(self.path_input.text().strip())
        self.convert_btn.setEnabled(has_path)
        self.export_btn.setEnabled(has_path and self.pptx_obj is not None)
        
    def browse_file(self):
        """Ouvre une boîte de dialogue pour sélectionner un fichier DOCX"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Sélectionner un fichier Word",
            "",
            "Word Documents (*.docx);;All Files (*)"
        )
        if file_path:
            self.path_input.setText(file_path)
            self.log_label.setText("")  # Efface le label
            self.pptx_obj = None  # Réinitialise l'objet
            self.update_button_states()  # Met à jour les boutons
    
    def convert_file(self):
        """Convertit le fichier DOCX en PPTX"""
        file_path = self.path_input.text()
        
        if not file_path:
            self.log_label.setText("⚠️ Veuillez d'abord sélectionner un fichier")
            self.log_label.setStyleSheet("""
                QLabel {
                    background-color: #fef7e0;
                    border-radius: 8px;
                    padding: 15px;
                    font-size: 13px;
                    color: #ea8600;
                }
            """)
            return
        
        try:
            self.log_label.setText("⏳ Conversion en cours...")
            self.log_label.setStyleSheet("""
                QLabel {
                    background-color: #f1f3f4;
                    border-radius: 8px;
                    padding: 15px;
                    font-size: 13px;
                    color: #3c4043;
                }
            """)
            QApplication.processEvents()  # Met à jour l'affichage
            
            # Capture les outputs de print
            output = io.StringIO()
            
            # Crée l'objet Powerpoint et effectue la conversion
            self.pptx_obj = to_powerpoint(self.headline_size)
            
            with redirect_stdout(output):
                self.pptx_obj.open(file_path)
                
            # Vérifie si l'ouverture a échoué
            open_output = output.getvalue()
            if "Fichier introuvable" in open_output or "Erreur inconnue" in open_output:
                self.log_label.setText(f"{open_output.strip()}")
                self.log_label.setStyleSheet("""
                    QLabel {
                        background-color: #fce8e6;
                        border-radius: 8px;
                        padding: 15px;
                        font-size: 13px;
                        color: #d93025;
                    }
                """)
                return
            
            # Continue avec la conversion
            output = io.StringIO()
            with redirect_stdout(output):
                self.pptx_obj.to_pptx()
            
            conversion_output = output.getvalue()
            self.log_label.setText(f"✅ {conversion_output.strip()}")
            self.log_label.setStyleSheet("""
                QLabel {
                    background-color: #e6f4ea;
                    border-radius: 8px;
                    padding: 15px;
                    font-size: 13px;
                    color: #137333;
                }
            """)
            self.update_button_states()  # Active le bouton Export
        except Exception as e:
            self.log_label.setText(f"Erreur lors de la conversion:\n{str(e)}")
            self.log_label.setStyleSheet("""
                QLabel {
                    background-color: #fce8e6;
                    border-radius: 8px;
                    padding: 15px;
                    font-size: 13px;
                    color: #d93025;
                }
            """)
    
    def export_file(self):
        """Exporte le fichier PPTX via une boîte de dialogue"""
        if self.pptx_obj is None:
            self.log_label.setText("⚠️ Veuillez d'abord convertir un fichier")
            self.log_label.setStyleSheet("""
                QLabel {
                    background-color: #fef7e0;
                    border-radius: 8px;
                    padding: 15px;
                    font-size: 13px;
                    color: #ea8600;
                }
            """)
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Enregistrer le fichier PowerPoint",
            "presentation.pptx",
            "PowerPoint Files (*.pptx);;All Files (*)"
        )
        
        if file_path:
            try:
                self.log_label.setText("💾 Exportation en cours...")
                self.log_label.setStyleSheet("""
                    QLabel {
                        background-color: #f1f3f4;
                        border-radius: 8px;
                        padding: 15px;
                        font-size: 13px;
                        color: #3c4043;
                    }
                """)
                QApplication.processEvents()
                
                # Capture les outputs de print
                output = io.StringIO()
                with redirect_stdout(output):
                    self.pptx_obj.pptx.save(file_path)
                
                export_output = output.getvalue()
                if export_output:
                    self.log_label.setText(f"✅ {export_output.strip()}\nFichier : {file_path}")
                else:
                    self.log_label.setText(f"✅ Fichier exporté avec succès :\n{file_path}")
                self.log_label.setStyleSheet("""
                    QLabel {
                        background-color: #e6f4ea;
                        border-radius: 8px;
                        padding: 15px;
                        font-size: 13px;
                        color: #137333;
                    }
                """)
            except Exception as e:
                self.log_label.setText(f"Erreur lors de l'exportation :\n{str(e)}")
                self.log_label.setStyleSheet("""
                    QLabel {
                        background-color: #fce8e6;
                        border-radius: 8px;
                        padding: 15px;
                        font-size: 13px;
                        color: #d93025;
                    }
                """)