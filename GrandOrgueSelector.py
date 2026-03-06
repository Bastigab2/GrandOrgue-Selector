import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
try:
    from PIL import Image, ImageTk, ImageGrab
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
import subprocess
import os, json, shutil, zipfile
from datetime import datetime
import time
import ctypes
if os.name == "nt":
    from ctypes import wintypes

CONFIG_FILE = "organs.json"
IMAGES_DIR = "images"
SETTINGS_FILE = "settings.json"
VERSION = "1.2.0"
FONT_FAMILY = "Segoe UI" if os.name == "nt" else "DejaVu Sans"

# Traduzioni complete per 5 lingue
LANG = {
    "en": {
        "title": "GrandOrgue Selector",
        "file": "File", "load": "Import organ", "import_folder": "Import folder",
        "export": "Export configuration", "import": "Import configuration", "exit": "Exit",
        "view": "View", "stats": "Statistics",
        "settings": "Settings", "images": "Images folder", "help": "Help",
        "guide": "Guide", "about": "About", "search": "Search...",
        "category": "Category", "all": "All", "baroque": "Baroque",
        "romantic": "Romantic", "modern": "Modern", "contemporary": "Contemporary",
        "other": "Other", "priority_up": "Priority +", "priority_down": "Priority -",
        "edit": "Edit", "delete": "Delete",
        "import_btn": "Import Folder", "load_btn": "Import Organ",
        "launch": "LAUNCH ORGAN", "no_img": "No image\n\nClick to add screenshot",
        "select": "Select an organ", "add_title": "Add Organ",
        "name": "Organ name:", "desc": "Description (optional):",
        "save": "Save", "cancel": "Cancel", "count": "organs",
        "language": "Language", "auto_title": "Automatic Launch",
        "auto_desc": "Auto-launch favorite organ on startup",
        "delay": "Delay (seconds):", "min_3": "(minimum 3 seconds)",
        "select_img": "Select image", "capture": "Capture screenshot",
        "remove_img": "Remove image", "position_window": "Position GrandOrgue window",
        "images_note": "Tip: You can crop images in the images folder:",
        "export_success": "Configuration exported successfully!",
        "import_success": "Configuration imported successfully!",
        "select_lang": "Select Language",
        "set_favorite": "Set as favorite",
        "remove_favorite": "Remove from favorites",
        "favorite_organ": "Favorite organ:",
        "no_favorite": "No favorite set",
        "press_any_key": "Press ANY KEY to cancel",
        "auto_launch_in": "Auto-launch in",
        "cancelled": "Cancelled",
        "skin": "Theme",
        "skin_classic": "Classic Dark",
        "skin_light_wood": "Light Wood",
        "skin_walnut": "Walnut",
        "select_skin": "Select Theme",
        "launched": "Launched",
        "times": "times",
        "last": "Last",
        "never": "Never",
        "confirm_delete": "Delete",
        "warning": "Warning",
        "error": "Error",
        "success": "Success",
        "ok": "OK",
        "select_organ_file": "Select organ file",
        "select_image": "Select image",
        "image_selected": "Image selected!",
        "imported": "Imported",
        "organs_imported": "organs",
        "cannot_save": "Cannot save",
        "cannot_load": "Cannot load",
        "file_not_found": "File not found",
        "cannot_launch": "Cannot launch",
        "cannot_capture": "Cannot capture",
        "screenshot_captured": "Screenshot captured!",
        "crop_tip": "You can crop it with an image editor.",
        "images_in": "Images are in",
        "cannot_open_folder": "Cannot open folder",
        "cannot_export": "Cannot export",
        "cannot_import": "Cannot import",
        "lang_changed": "Language changed!",
        "restart_app": "Please restart the application.",
        "settings_saved": "Settings saved!",
        "min_delay_warning": "Minimum delay is 3 seconds!\n\nThis gives you time to cancel.",
        "how_it_works": "HOW IT WORKS",
        "auto_info_1": "On next startup, countdown appears",
        "auto_info_2": "Press ANY KEY during countdown to cancel",
        "auto_info_3": "Minimum 3 seconds gives you time to react",
        "auto_info_4": "Launches favorite organ (set with right-click)",
        "edit_organ": "Edit Organ",
        "name_label": "Name:",
        "category_label": "Category:",
        "description_label": "Description:",
        "capture_title": "Capture Screenshot",
        "capture_procedure": "PROCEDURE:",
        "capture_step1": "1. Click 'Launch Organ' to open GrandOrgue",
        "capture_step2": "2. Wait for organ to load completely",
        "capture_step3": "3. Position window as you prefer",
        "capture_step4": "4. Return here and click 'Capture Screenshot'",
        "capture_step5": "5. 5 seconds countdown - Position on GrandOrgue!",
        "capture_step6": "6. Screenshot taken automatically",
        "launch_organ_btn": "Launch Organ",
        "capture_btn": "Capture Screenshot",
        "error_loading_image": "Error loading image",
        "confirm_remove_image": "Remove image?",
        "total_organs": "Total organs",
        "total_launches": "Total launches",
        "by_category": "By category",
        "most_used": "Most used"
    },
    "it": {
        "title": "GrandOrgue Selector",
        "file": "File", "load": "Importa organo", "import_folder": "Importa cartella",
        "export": "Esporta configurazione", "import": "Importa configurazione", "exit": "Esci",
        "view": "Visualizza", "stats": "Statistiche",
        "settings": "Impostazioni", "images": "Cartella immagini", "help": "Aiuto",
        "guide": "Guida", "about": "Info", "search": "Cerca...",
        "category": "Categoria", "all": "Tutti", "baroque": "Barocco",
        "romantic": "Romantico", "modern": "Moderno", "contemporary": "Contemporaneo",
        "other": "Altro", "priority_up": "Priorita +", "priority_down": "Priorita -",
        "edit": "Modifica", "delete": "Elimina",
        "import_btn": "Importa Cartella", "load_btn": "Importa Organo",
        "launch": "AVVIA ORGANO", "no_img": "Nessuna immagine\n\nClicca per aggiungere",
        "select": "Seleziona un organo", "add_title": "Aggiungi Organo",
        "name": "Nome organo:", "desc": "Descrizione (opzionale):",
        "save": "Salva", "cancel": "Annulla", "count": "organi",
        "language": "Lingua", "auto_title": "Avvio Automatico",
        "auto_desc": "Avvia automaticamente l'organo preferito all'apertura",
        "delay": "Ritardo (secondi):", "min_3": "(minimo 3 secondi)",
        "select_img": "Seleziona immagine", "capture": "Cattura screenshot",
        "remove_img": "Rimuovi immagine", "position_window": "Posiziona finestra GrandOrgue",
        "images_note": "Suggerimento: Puoi ritagliare le immagini nella cartella:",
        "export_success": "Configurazione esportata con successo!",
        "import_success": "Configurazione importata con successo!",
        "select_lang": "Seleziona Lingua",
        "set_favorite": "Imposta come preferito",
        "remove_favorite": "Rimuovi dai preferiti",
        "favorite_organ": "Organo preferito:",
        "no_favorite": "Nessun preferito impostato",
        "press_any_key": "Premi un TASTO QUALSIASI per annullare",
        "auto_launch_in": "Avvio automatico tra",
        "cancelled": "Annullato",
        "skin": "Tema",
        "skin_classic": "Classico Scuro",
        "skin_light_wood": "Legno Chiaro",
        "skin_walnut": "Legno Noce",
        "select_skin": "Seleziona Tema",
        "launched": "Avviato",
        "times": "volte",
        "last": "Ultimo",
        "never": "Mai",
        "confirm_delete": "Eliminare",
        "warning": "Attenzione",
        "error": "Errore",
        "success": "Successo",
        "ok": "OK",
        "select_organ_file": "Seleziona file organo",
        "select_image": "Seleziona immagine",
        "image_selected": "Immagine selezionata!",
        "imported": "Importati",
        "organs_imported": "organi",
        "cannot_save": "Impossibile salvare",
        "cannot_load": "Impossibile caricare",
        "file_not_found": "File non trovato",
        "cannot_launch": "Impossibile avviare",
        "cannot_capture": "Impossibile catturare",
        "screenshot_captured": "Screenshot catturato!",
        "crop_tip": "Puoi ritagliarlo con un editor di immagini.",
        "images_in": "Le immagini sono in",
        "cannot_open_folder": "Impossibile aprire cartella",
        "cannot_export": "Impossibile esportare",
        "cannot_import": "Impossibile importare",
        "lang_changed": "Lingua cambiata!",
        "restart_app": "Riavvia l'applicazione.",
        "settings_saved": "Impostazioni salvate!",
        "min_delay_warning": "Il ritardo minimo e 3 secondi!\n\nQuesto ti da tempo per annullare.",
        "how_it_works": "COME FUNZIONA",
        "auto_info_1": "Al prossimo avvio, appare il countdown",
        "auto_info_2": "Premi un TASTO QUALSIASI per annullare",
        "auto_info_3": "Minimo 3 secondi ti danno tempo per reagire",
        "auto_info_4": "Avvia l'organo preferito (imposta con tasto destro)",
        "edit_organ": "Modifica Organo",
        "name_label": "Nome:",
        "category_label": "Categoria:",
        "description_label": "Descrizione:",
        "capture_title": "Cattura Screenshot",
        "capture_procedure": "PROCEDURA:",
        "capture_step1": "1. Clicca 'Avvia Organo' per aprire GrandOrgue",
        "capture_step2": "2. Aspetta che l'organo si carichi completamente",
        "capture_step3": "3. Posiziona la finestra come preferisci",
        "capture_step4": "4. Torna qui e clicca 'Cattura Screenshot'",
        "capture_step5": "5. Countdown 5 secondi - Posizionati su GrandOrgue!",
        "capture_step6": "6. Screenshot catturato automaticamente",
        "launch_organ_btn": "Avvia Organo",
        "capture_btn": "Cattura Screenshot",
        "error_loading_image": "Errore caricamento immagine",
        "confirm_remove_image": "Rimuovere l'immagine?",
        "total_organs": "Organi totali",
        "total_launches": "Avvii totali",
        "by_category": "Per categoria",
        "most_used": "Piu usati"
    },
    "fr": {
        "title": "GrandOrgue Selector",
        "file": "Fichier", "load": "Importer orgue", "import_folder": "Importer dossier",
        "export": "Exporter configuration", "import": "Importer configuration", "exit": "Quitter",
        "view": "Affichage", "stats": "Statistiques",
        "settings": "Parametres", "images": "Dossier images", "help": "Aide",
        "guide": "Guide", "about": "A propos", "search": "Rechercher...",
        "category": "Categorie", "all": "Tous", "baroque": "Baroque",
        "romantic": "Romantique", "modern": "Moderne", "contemporary": "Contemporain",
        "other": "Autre", "priority_up": "Priorite +", "priority_down": "Priorite -",
        "edit": "Modifier", "delete": "Supprimer",
        "import_btn": "Importer Dossier", "load_btn": "Importer Orgue",
        "launch": "LANCER ORGUE", "no_img": "Pas d'image\n\nCliquez pour ajouter",
        "select": "Selectionnez un orgue", "add_title": "Ajouter Orgue",
        "name": "Nom:", "desc": "Description (optionnelle):",
        "save": "Enregistrer", "cancel": "Annuler", "count": "orgues",
        "language": "Langue", "auto_title": "Lancement Auto",
        "auto_desc": "Lancer automatiquement l'orgue favori au demarrage",
        "delay": "Delai (secondes):", "min_3": "(minimum 3 secondes)",
        "select_img": "Selectionner image", "capture": "Capturer",
        "remove_img": "Supprimer", "position_window": "Positionner fenetre",
        "images_note": "Astuce: Vous pouvez recadrer les images dans:",
        "export_success": "Configuration exportee!",
        "import_success": "Configuration importee!",
        "select_lang": "Choisir Langue",
        "set_favorite": "Definir comme favori",
        "remove_favorite": "Retirer des favoris",
        "favorite_organ": "Orgue favori:",
        "no_favorite": "Aucun favori defini",
        "press_any_key": "Appuyez sur une TOUCHE pour annuler",
        "auto_launch_in": "Lancement auto dans",
        "cancelled": "Annule",
        "skin": "Theme",
        "skin_classic": "Classique Sombre",
        "skin_light_wood": "Bois Clair",
        "skin_walnut": "Noyer",
        "select_skin": "Selectionner Theme",
        "launched": "Lance",
        "times": "fois",
        "last": "Dernier",
        "never": "Jamais",
        "confirm_delete": "Supprimer",
        "warning": "Attention",
        "error": "Erreur",
        "success": "Succes",
        "ok": "OK",
        "select_organ_file": "Selectionner fichier orgue",
        "select_image": "Selectionner image",
        "image_selected": "Image selectionnee!",
        "imported": "Importes",
        "organs_imported": "orgues",
        "cannot_save": "Impossible de sauvegarder",
        "cannot_load": "Impossible de charger",
        "file_not_found": "Fichier non trouve",
        "cannot_launch": "Impossible de lancer",
        "cannot_capture": "Impossible de capturer",
        "screenshot_captured": "Capture d'ecran effectuee!",
        "crop_tip": "Vous pouvez le recadrer avec un editeur d'images.",
        "images_in": "Les images sont dans",
        "cannot_open_folder": "Impossible d'ouvrir le dossier",
        "cannot_export": "Impossible d'exporter",
        "cannot_import": "Impossible d'importer",
        "lang_changed": "Langue changee!",
        "restart_app": "Veuillez redemarrer l'application.",
        "settings_saved": "Parametres sauvegardes!",
        "min_delay_warning": "Le delai minimum est de 3 secondes!\n\nCela vous donne le temps d'annuler.",
        "how_it_works": "COMMENT CA MARCHE",
        "auto_info_1": "Au prochain demarrage, le compte a rebours apparait",
        "auto_info_2": "Appuyez sur une TOUCHE pour annuler",
        "auto_info_3": "Minimum 3 secondes vous donne le temps de reagir",
        "auto_info_4": "Lance l'orgue favori (defini avec clic droit)",
        "edit_organ": "Modifier Orgue",
        "name_label": "Nom:",
        "category_label": "Categorie:",
        "description_label": "Description:",
        "capture_title": "Capturer Screenshot",
        "capture_procedure": "PROCEDURE:",
        "capture_step1": "1. Cliquez 'Lancer Orgue' pour ouvrir GrandOrgue",
        "capture_step2": "2. Attendez que l'orgue se charge completement",
        "capture_step3": "3. Positionnez la fenetre comme vous preferez",
        "capture_step4": "4. Revenez ici et cliquez 'Capturer'",
        "capture_step5": "5. Compte a rebours 5 secondes - Positionnez-vous!",
        "capture_step6": "6. Capture effectuee automatiquement",
        "launch_organ_btn": "Lancer Orgue",
        "capture_btn": "Capturer",
        "error_loading_image": "Erreur chargement image",
        "confirm_remove_image": "Supprimer l'image?",
        "total_organs": "Total orgues",
        "total_launches": "Total lancements",
        "by_category": "Par categorie",
        "most_used": "Plus utilises"
    },
    "de": {
        "title": "GrandOrgue Selector",
        "file": "Datei", "load": "Orgel importieren", "import_folder": "Ordner importieren",
        "export": "Konfiguration exportieren", "import": "Konfiguration importieren", "exit": "Beenden",
        "view": "Ansicht", "stats": "Statistiken",
        "settings": "Einstellungen", "images": "Bilderordner", "help": "Hilfe",
        "guide": "Anleitung", "about": "Uber", "search": "Suchen...",
        "category": "Kategorie", "all": "Alle", "baroque": "Barock",
        "romantic": "Romantisch", "modern": "Modern", "contemporary": "Zeitgenossisch",
        "other": "Andere", "priority_up": "Prioritat +", "priority_down": "Prioritat -",
        "edit": "Bearbeiten", "delete": "Loschen",
        "import_btn": "Ordner Importieren", "load_btn": "Orgel Importieren",
        "launch": "ORGEL STARTEN", "no_img": "Kein Bild\n\nKlicken zum Hinzufugen",
        "select": "Wahlen Sie eine Orgel", "add_title": "Orgel Hinzufugen",
        "name": "Name:", "desc": "Beschreibung (optional):",
        "save": "Speichern", "cancel": "Abbrechen", "count": "Orgeln",
        "language": "Sprache", "auto_title": "Automatischer Start",
        "auto_desc": "Lieblingsorgel automatisch beim Start starten",
        "delay": "Verzogerung (Sek.):", "min_3": "(mindestens 3 Sek.)",
        "select_img": "Bild auswahlen", "capture": "Aufnahme",
        "remove_img": "Entfernen", "position_window": "Fenster positionieren",
        "images_note": "Tipp: Sie konnen Bilder zuschneiden in:",
        "export_success": "Konfiguration exportiert!",
        "import_success": "Konfiguration importiert!",
        "select_lang": "Sprache Wahlen",
        "set_favorite": "Als Favorit setzen",
        "remove_favorite": "Aus Favoriten entfernen",
        "favorite_organ": "Lieblingsorgel:",
        "no_favorite": "Kein Favorit gesetzt",
        "press_any_key": "Drucken Sie eine TASTE zum Abbrechen",
        "auto_launch_in": "Automatischer Start in",
        "cancelled": "Abgebrochen",
        "skin": "Thema",
        "skin_classic": "Klassisch Dunkel",
        "skin_light_wood": "Helles Holz",
        "skin_walnut": "Nussbaum",
        "select_skin": "Thema Wahlen",
        "launched": "Gestartet",
        "times": "mal",
        "last": "Letzter",
        "never": "Nie",
        "confirm_delete": "Loschen",
        "warning": "Warnung",
        "error": "Fehler",
        "success": "Erfolg",
        "ok": "OK",
        "select_organ_file": "Orgeldatei auswahlen",
        "select_image": "Bild auswahlen",
        "image_selected": "Bild ausgewahlt!",
        "imported": "Importiert",
        "organs_imported": "Orgeln",
        "cannot_save": "Kann nicht speichern",
        "cannot_load": "Kann nicht laden",
        "file_not_found": "Datei nicht gefunden",
        "cannot_launch": "Kann nicht starten",
        "cannot_capture": "Kann nicht aufnehmen",
        "screenshot_captured": "Screenshot aufgenommen!",
        "crop_tip": "Sie konnen es mit einem Bildbearbeitungsprogramm zuschneiden.",
        "images_in": "Bilder sind in",
        "cannot_open_folder": "Kann Ordner nicht offnen",
        "cannot_export": "Kann nicht exportieren",
        "cannot_import": "Kann nicht importieren",
        "lang_changed": "Sprache geandert!",
        "restart_app": "Bitte starten Sie die Anwendung neu.",
        "settings_saved": "Einstellungen gespeichert!",
        "min_delay_warning": "Minimale Verzogerung ist 3 Sekunden!\n\nDas gibt Ihnen Zeit zum Abbrechen.",
        "how_it_works": "WIE ES FUNKTIONIERT",
        "auto_info_1": "Beim nachsten Start erscheint der Countdown",
        "auto_info_2": "Drucken Sie eine TASTE zum Abbrechen",
        "auto_info_3": "Minimum 3 Sekunden gibt Ihnen Zeit zu reagieren",
        "auto_info_4": "Startet Lieblingsorgel (per Rechtsklick festlegen)",
        "edit_organ": "Orgel Bearbeiten",
        "name_label": "Name:",
        "category_label": "Kategorie:",
        "description_label": "Beschreibung:",
        "capture_title": "Screenshot Aufnehmen",
        "capture_procedure": "VORGEHENSWEISE:",
        "capture_step1": "1. Klicken Sie 'Orgel Starten' um GrandOrgue zu offnen",
        "capture_step2": "2. Warten Sie bis die Orgel vollstandig geladen ist",
        "capture_step3": "3. Positionieren Sie das Fenster wie gewunscht",
        "capture_step4": "4. Kommen Sie hierher zuruck und klicken Sie 'Aufnahme'",
        "capture_step5": "5. 5 Sekunden Countdown - Positionieren Sie sich!",
        "capture_step6": "6. Screenshot wird automatisch aufgenommen",
        "launch_organ_btn": "Orgel Starten",
        "capture_btn": "Aufnehmen",
        "error_loading_image": "Fehler beim Laden des Bildes",
        "confirm_remove_image": "Bild entfernen?",
        "total_organs": "Gesamte Orgeln",
        "total_launches": "Gesamte Starts",
        "by_category": "Nach Kategorie",
        "most_used": "Meistgenutzt"
    },
    "es": {
        "title": "GrandOrgue Selector",
        "file": "Archivo", "load": "Cargar organo", "import_folder": "Importar carpeta",
        "export": "Exportar configuracion", "import": "Importar configuracion", "exit": "Salir",
        "view": "Ver", "stats": "Estadisticas",
        "settings": "Configuracion", "images": "Carpeta imagenes", "help": "Ayuda",
        "guide": "Guia", "about": "Acerca de", "search": "Buscar...",
        "category": "Categoria", "all": "Todos", "baroque": "Barroco",
        "romantic": "Romantico", "modern": "Moderno", "contemporary": "Contemporaneo",
        "other": "Otro", "priority_up": "Prioridad +", "priority_down": "Prioridad -",
        "edit": "Editar", "delete": "Eliminar",
        "import_btn": "Importar Carpeta", "load_btn": "Cargar Organo",
        "launch": "INICIAR ORGANO", "no_img": "Sin imagen\n\nClic para agregar",
        "select": "Selecciona un organo", "add_title": "Agregar Organo",
        "name": "Nombre:", "desc": "Descripcion (opcional):",
        "save": "Guardar", "cancel": "Cancelar", "count": "organos",
        "language": "Idioma", "auto_title": "Inicio Automatico",
        "auto_desc": "Iniciar automaticamente el organo favorito al abrir",
        "delay": "Retraso (segundos):", "min_3": "(minimo 3 segundos)",
        "select_img": "Seleccionar imagen", "capture": "Capturar",
        "remove_img": "Eliminar", "position_window": "Posicionar ventana",
        "images_note": "Consejo: Puedes recortar imagenes en:",
        "export_success": "Configuracion exportada!",
        "import_success": "Configuracion importada!",
        "select_lang": "Seleccionar Idioma",
        "set_favorite": "Establecer como favorito",
        "remove_favorite": "Quitar de favoritos",
        "favorite_organ": "Organo favorito:",
        "no_favorite": "Sin favorito establecido",
        "press_any_key": "Presiona CUALQUIER TECLA para cancelar",
        "auto_launch_in": "Inicio automatico en",
        "cancelled": "Cancelado",
        "skin": "Tema",
        "skin_classic": "Clasico Oscuro",
        "skin_light_wood": "Madera Clara",
        "skin_walnut": "Nogal",
        "select_skin": "Seleccionar Tema",
        "launched": "Iniciado",
        "times": "veces",
        "last": "Ultimo",
        "never": "Nunca",
        "confirm_delete": "Eliminar",
        "warning": "Advertencia",
        "error": "Error",
        "success": "Exito",
        "ok": "OK",
        "select_organ_file": "Seleccionar archivo de organo",
        "select_image": "Seleccionar imagen",
        "image_selected": "Imagen seleccionada!",
        "imported": "Importados",
        "organs_imported": "organos",
        "cannot_save": "No se puede guardar",
        "cannot_load": "No se puede cargar",
        "file_not_found": "Archivo no encontrado",
        "cannot_launch": "No se puede iniciar",
        "cannot_capture": "No se puede capturar",
        "screenshot_captured": "Captura realizada!",
        "crop_tip": "Puedes recortarla con un editor de imagenes.",
        "images_in": "Las imagenes estan en",
        "cannot_open_folder": "No se puede abrir la carpeta",
        "cannot_export": "No se puede exportar",
        "cannot_import": "No se puede importar",
        "lang_changed": "Idioma cambiado!",
        "restart_app": "Por favor reinicia la aplicacion.",
        "settings_saved": "Configuracion guardada!",
        "min_delay_warning": "El retraso minimo es 3 segundos!\n\nEsto te da tiempo para cancelar.",
        "how_it_works": "COMO FUNCIONA",
        "auto_info_1": "Al proximo inicio, aparece la cuenta regresiva",
        "auto_info_2": "Presiona CUALQUIER TECLA para cancelar",
        "auto_info_3": "Minimo 3 segundos te dan tiempo para reaccionar",
        "auto_info_4": "Inicia el organo favorito (establecer con clic derecho)",
        "edit_organ": "Editar Organo",
        "name_label": "Nombre:",
        "category_label": "Categoria:",
        "description_label": "Descripcion:",
        "capture_title": "Capturar Pantalla",
        "capture_procedure": "PROCEDIMIENTO:",
        "capture_step1": "1. Clic en 'Iniciar Organo' para abrir GrandOrgue",
        "capture_step2": "2. Espera a que el organo cargue completamente",
        "capture_step3": "3. Posiciona la ventana como prefieras",
        "capture_step4": "4. Regresa aqui y clic en 'Capturar'",
        "capture_step5": "5. Cuenta regresiva 5 segundos - Posicionate!",
        "capture_step6": "6. Captura tomada automaticamente",
        "launch_organ_btn": "Iniciar Organo",
        "capture_btn": "Capturar",
        "error_loading_image": "Error al cargar imagen",
        "confirm_remove_image": "Eliminar imagen?",
        "total_organs": "Total organos",
        "total_launches": "Total inicios",
        "by_category": "Por categoria",
        "most_used": "Mas usados"
    }
}

# Bandiere per le lingue
LANG_FLAGS = {
    "en": "GB",
    "it": "IT",
    "fr": "FR",
    "de": "DE",
    "es": "ES"
}

# Definizione delle skin/temi
SKINS = {
    "classic": {
        "name_key": "skin_classic",
        "bg": "#2b2b2b",
        "sidebar_bg": "#1e1e1e",
        "fg": "#ffffff",
        "accent": "#4a90e2",
        "hover": "#357abd",
        "button_bg": "#4a90e2",
        "button_fg": "#ffffff",
        "list_bg": "#2d2d2d",
        "list_fg": "#ffffff",
        "list_select_bg": "#4a90e2",
        "header_bg": "#4a90e2",
        "info_bg": "#1e1e1e",
        "border_color": "#555555",
        "btn_priority_up": "#4a90e2",
        "btn_priority_down": "#4a90e2",
        "btn_edit": "#5cb85c",
        "btn_delete": "#d9534f",
        "btn_load": "#6f42c1",
        "btn_import": "#f0ad4e",
        "nav_btn_bg": "#3a3a3a",
        "nav_btn_fg": "#ffffff"
    },
    "light_wood": {
        "name_key": "skin_light_wood",
        "bg": "#d4a574",
        "sidebar_bg": "#c4956a",
        "fg": "#3d2914",
        "accent": "#8b5a2b",
        "hover": "#6b4423",
        "button_bg": "#f5deb3",
        "button_fg": "#3d2914",
        "list_bg": "#e8d4b8",
        "list_fg": "#3d2914",
        "list_select_bg": "#8b5a2b",
        "header_bg": "#8b5a2b",
        "info_bg": "#c4956a",
        "border_color": "#8b5a2b",
        "btn_priority_up": "#8B4513",
        "btn_priority_down": "#8B4513",
        "btn_edit": "#B87333",
        "btn_delete": "#A52A2A",
        "btn_load": "#A0522D",
        "btn_import": "#A0522D",
        "nav_btn_bg": "#f5deb3",
        "nav_btn_fg": "#3d2914"
    },
    "walnut": {
        "name_key": "skin_walnut",
        "bg": "#4a3728",
        "sidebar_bg": "#3d2d22",
        "fg": "#f5f5dc",
        "accent": "#8b4513",
        "hover": "#a0522d",
        "button_bg": "#1a1a1a",
        "button_fg": "#f5f5dc",
        "list_bg": "#5d4037",
        "list_fg": "#f5f5dc",
        "list_select_bg": "#8b4513",
        "header_bg": "#5d4037",
        "info_bg": "#3d2d22",
        "border_color": "#8b4513",
        "btn_priority_up": "#8B4513",
        "btn_priority_down": "#8B4513",
        "btn_edit": "#B87333",
        "btn_delete": "#A52A2A",
        "btn_load": "#A0522D",
        "btn_import": "#A0522D",
        "nav_btn_bg": "#f5f5dc",
        "nav_btn_fg": "#1a1a1a"
    }
}


class OrgansApp:
    def __init__(self, root):
        self.root = root

        # Impostazioni
        self.settings = {
            "auto_launch": False,
            "auto_launch_delay": 5,
            "first_run": True,
            "language": "en",
            "skin": "classic",
            "favorite_organ": None
        }
        self.load_settings()
        self.t = LANG[self.settings["language"]]
        self.skin = SKINS[self.settings.get("skin", "classic")]

        self.root.title(f"{self.t['title']} v{VERSION}")

        # Applica tema
        self.apply_skin()

        # Fullscreen
        if os.name == "nt":
            self.root.state("zoomed")
        else:
            self.root.attributes("-fullscreen", True)

        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass

        if not os.path.exists(IMAGES_DIR):
            os.makedirs(IMAGES_DIR)

        self.organs = []
        self.filtered_organs = []
        self.update_categories()
        self.auto_launch_timer = None
        self.countdown_active = False
        self.countdown_window = None

        self.create_ui()
        self.load_config()
        self.setup_keybindings()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def apply_skin(self):
        """Applica il tema corrente"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.root.configure(bg=self.skin["bg"])

        # Configura stili ttk
        self.style.configure("TEntry", fieldbackground=self.skin["list_bg"],
                            foreground=self.skin["list_fg"])
        self.style.configure("TCombobox", fieldbackground=self.skin["list_bg"],
                            foreground=self.skin["list_fg"])

    def center_window(self, window, width=None, height=None):
        """Centra una finestra rispetto alla finestra principale"""
        window.update_idletasks()
        if width is None:
            width = window.winfo_width()
        if height is None:
            height = window.winfo_height()

        # Posizione della finestra principale
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()

        # Calcola posizione centrata
        x = root_x + (root_w - width) // 2
        y = root_y + (root_h - height) // 2

        window.geometry(f"{width}x{height}+{x}+{y}")

    def get_current_monitor_bbox(self):
        """Restituisce le coordinate (left, top, right, bottom) del monitor corrente"""
        if os.name != "nt":
            # Su sistemi non-Windows, cattura tutto lo schermo
            return None

        try:
            # Struttura per le info del monitor
            class MONITORINFO(ctypes.Structure):
                _fields_ = [
                    ("cbSize", ctypes.c_ulong),
                    ("rcMonitor", wintypes.RECT),
                    ("rcWork", wintypes.RECT),
                    ("dwFlags", ctypes.c_ulong)
                ]

            # Ottieni posizione della finestra principale
            x = self.root.winfo_x() + self.root.winfo_width() // 2
            y = self.root.winfo_y() + self.root.winfo_height() // 2

            # Trova il monitor che contiene questo punto
            MONITOR_DEFAULTTONEAREST = 2
            hMonitor = ctypes.windll.user32.MonitorFromPoint(
                wintypes.POINT(x, y),
                MONITOR_DEFAULTTONEAREST
            )

            # Ottieni info del monitor
            mi = MONITORINFO()
            mi.cbSize = ctypes.sizeof(MONITORINFO)
            ctypes.windll.user32.GetMonitorInfoW(hMonitor, ctypes.byref(mi))

            return (mi.rcMonitor.left, mi.rcMonitor.top,
                    mi.rcMonitor.right, mi.rcMonitor.bottom)
        except:
            return None

    def update_categories(self):
        """Aggiorna categorie in base alla lingua"""
        self.categories = [
            self.t["all"], self.t["baroque"], self.t["romantic"],
            self.t["modern"], self.t["contemporary"], self.t["other"]
        ]

    def create_ui(self):
        """Crea interfaccia utente"""
        # TOP BAR
        top_bar = tk.Frame(self.root, bg=self.skin["header_bg"], height=50)
        top_bar.pack(side="top", fill="x")

        title_label = tk.Label(top_bar, text=f"{self.t['title']}",
                               font=(FONT_FAMILY, 16, "bold"),
                               bg=self.skin["header_bg"], fg="white")
        title_label.pack(side="left", padx=20, pady=10)

        self.stats_label = tk.Label(top_bar, text="", font=(FONT_FAMILY, 9),
                                     bg=self.skin["header_bg"], fg="white")
        self.stats_label.pack(side="right", padx=20)

        self.countdown_label = tk.Label(top_bar, text="", font=(FONT_FAMILY, 10, "bold"),
                                        bg=self.skin["header_bg"], fg="yellow")
        self.countdown_label.pack(side="right", padx=10)

        # SIDEBAR
        sidebar = tk.Frame(self.root, bg=self.skin["sidebar_bg"], width=300)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        # Ricerca/Filtro
        search_frame = tk.Frame(sidebar, bg=self.skin["sidebar_bg"])
        search_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(search_frame, text=f"{self.t['search'].replace('...', '')}:",
                bg=self.skin["sidebar_bg"], fg=self.skin["fg"],
                font=(FONT_FAMILY, 9)).pack(anchor="w")

        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_organs)
        search_entry = tk.Entry(search_frame, textvariable=self.search_var,
                                font=(FONT_FAMILY, 10), bg=self.skin["list_bg"],
                                fg=self.skin["list_fg"], insertbackground=self.skin["fg"])
        search_entry.pack(fill="x", pady=5)

        # Filtro categoria
        filter_frame = tk.Frame(sidebar, bg=self.skin["sidebar_bg"])
        filter_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(filter_frame, text=f"{self.t['category']}:", bg=self.skin["sidebar_bg"],
                fg=self.skin["fg"], font=(FONT_FAMILY, 9)).pack(anchor="w")

        self.category_var = tk.StringVar(value=self.t["all"])
        self.category_combo = ttk.Combobox(filter_frame, textvariable=self.category_var,
                                          values=self.categories, state="readonly",
                                          font=(FONT_FAMILY, 9))
        self.category_combo.pack(fill="x", pady=5)
        self.category_combo.bind("<<ComboboxSelected>>", self.filter_organs)

        # Lista organi con frecce navigazione
        list_container = tk.Frame(sidebar, bg=self.skin["sidebar_bg"])
        list_container.pack(fill="both", expand=True, padx=10, pady=5)

        # Freccia SU per navigazione
        self.nav_up_btn = tk.Button(list_container, text="^",
                                    command=self.navigate_up,
                                    bg=self.skin["nav_btn_bg"], fg=self.skin["nav_btn_fg"],
                                    font=(FONT_FAMILY, 10, "bold"),
                                    relief="flat", cursor="hand2", height=1)
        self.nav_up_btn.pack(fill="x", pady=(0, 2))

        list_frame = tk.Frame(list_container, bg=self.skin["sidebar_bg"])
        list_frame.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(list_frame, font=(FONT_FAMILY, 10),
                                  bg=self.skin["list_bg"], fg=self.skin["list_fg"],
                                  selectbackground=self.skin["list_select_bg"],
                                  selectforeground="white",
                                  borderwidth=0, highlightthickness=0,
                                  yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        # Freccia GIU per navigazione
        self.nav_down_btn = tk.Button(list_container, text="v",
                                      command=self.navigate_down,
                                      bg=self.skin["nav_btn_bg"], fg=self.skin["nav_btn_fg"],
                                      font=(FONT_FAMILY, 10, "bold"),
                                      relief="flat", cursor="hand2", height=1)
        self.nav_down_btn.pack(fill="x", pady=(2, 0))

        self.listbox.bind("<Return>", self.launch_orgue)
        self.listbox.bind("<Double-Button-1>", self.launch_orgue)
        self.listbox.bind("<<ListboxSelect>>", self.show_preview)
        self.listbox.bind("<Button-3>", self.show_context_menu)

        # Pulsanti sidebar
        btn_frame = tk.Frame(sidebar, bg=self.skin["sidebar_bg"])
        btn_frame.pack(fill="x", padx=10, pady=10)

        buttons = [
            (self.t["priority_up"], self.move_up, self.skin["btn_priority_up"]),
            (self.t["priority_down"], self.move_down, self.skin["btn_priority_down"]),
            (self.t["edit"], self.edit_organ, self.skin["btn_edit"]),
            (self.t["delete"], self.delete_organ, self.skin["btn_delete"]),
            (self.t["load_btn"], self.load_orgue, self.skin["btn_load"]),
            (self.t["import_btn"], self.import_folder, self.skin["btn_import"]),
        ]

        for text, command, color in buttons:
            btn = tk.Button(btn_frame, text=text, command=command,
                          bg=color, fg=self.skin["button_fg"],
                          font=(FONT_FAMILY, 9, "bold"),
                          relief="flat", cursor="hand2",
                          padx=10, pady=8)
            btn.pack(fill="x", pady=2)
            btn.bind("<Enter>", lambda e, b=btn, c=color: b.config(bg=self.adjust_color(c, -20)))
            btn.bind("<Leave>", lambda e, b=btn, c=color: b.config(bg=c))

        # AREA CENTRALE
        center_frame = tk.Frame(self.root, bg=self.skin["bg"])
        center_frame.pack(side="top", fill="both", expand=True)

        # Preview immagine
        preview_container = tk.Frame(center_frame, bg=self.skin["info_bg"],
                                    relief="solid", borderwidth=1)
        preview_container.pack(fill="both", expand=True, padx=10, pady=10)

        self.preview_label = tk.Label(preview_container, bg=self.skin["info_bg"],
                                      fg=self.skin["fg"],
                                      text=self.t["select"],
                                      font=(FONT_FAMILY, 12))
        self.preview_label.pack(fill="both", expand=True, padx=20, pady=20)

        # Info organo
        info_frame = tk.Frame(center_frame, bg=self.skin["info_bg"],
                             relief="solid", borderwidth=1)
        info_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.name_label = tk.Label(info_frame, text="",
                                   font=(FONT_FAMILY, 14, "bold"),
                                   bg=self.skin["info_bg"], fg=self.skin["fg"], anchor="w")
        self.name_label.pack(fill="x", padx=20, pady=(15, 5))

        self.category_label = tk.Label(info_frame, text="",
                                       font=(FONT_FAMILY, 9),
                                       bg=self.skin["info_bg"], fg="#aaaaaa", anchor="w")
        self.category_label.pack(fill="x", padx=20, pady=(0, 5))

        self.desc_label = tk.Label(info_frame, text="", wraplength=800,
                                   justify="left",
                                   font=(FONT_FAMILY, 10),
                                   bg=self.skin["info_bg"], fg=self.skin["fg"], anchor="w")
        self.desc_label.pack(fill="x", padx=20, pady=(0, 5))

        self.stats_organ_label = tk.Label(info_frame, text="",
                                          font=(FONT_FAMILY, 8),
                                          bg=self.skin["info_bg"], fg="#888888", anchor="w")
        self.stats_organ_label.pack(fill="x", padx=20, pady=(0, 15))

        # Pulsante AVVIA
        self.launch_btn = tk.Button(info_frame, text=self.t["launch"],
                                    command=self.launch_orgue,
                                    bg=self.skin["accent"], fg="white",
                                    font=(FONT_FAMILY, 14, "bold"),
                                    relief="flat", cursor="hand2",
                                    padx=20, pady=15)
        self.launch_btn.pack(fill="x", padx=20, pady=(0, 15))
        self.launch_btn.bind("<Enter>", lambda e: self.launch_btn.config(bg=self.skin["hover"]))
        self.launch_btn.bind("<Leave>", lambda e: self.launch_btn.config(bg=self.skin["accent"]))

        # MENU
        self.create_menu()

    def navigate_up(self):
        """Naviga all'organo precedente nella lista"""
        selection = self.listbox.curselection()
        if selection:
            current = selection[0]
            if current > 0:
                self.listbox.selection_clear(0, "end")
                self.listbox.selection_set(current - 1)
                self.listbox.see(current - 1)
                self.listbox.activate(current - 1)
                self.show_preview()
        elif self.filtered_organs:
            self.listbox.selection_set(0)
            self.show_preview()

    def navigate_down(self):
        """Naviga all'organo successivo nella lista"""
        selection = self.listbox.curselection()
        if selection:
            current = selection[0]
            if current < len(self.filtered_organs) - 1:
                self.listbox.selection_clear(0, "end")
                self.listbox.selection_set(current + 1)
                self.listbox.see(current + 1)
                self.listbox.activate(current + 1)
                self.show_preview()
        elif self.filtered_organs:
            self.listbox.selection_set(0)
            self.show_preview()

    def show_context_menu(self, event):
        """Mostra menu contestuale per impostare preferito"""
        selection = self.listbox.nearest(event.y)
        if selection >= 0 and selection < len(self.filtered_organs):
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set(selection)
            self.show_preview()

            organ = self.filtered_organs[selection]
            menu = tk.Menu(self.root, tearoff=0)

            is_favorite = (self.settings.get("favorite_organ") == organ.get("file"))

            if is_favorite:
                menu.add_command(label=f"{self.t['remove_favorite']}",
                               command=lambda: self.remove_favorite())
            else:
                menu.add_command(label=f"{self.t['set_favorite']}",
                               command=lambda: self.set_favorite(organ))

            menu.add_separator()
            menu.add_command(label=self.t["edit"], command=self.edit_organ)
            menu.add_command(label=self.t["delete"], command=self.delete_organ)

            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

    def set_favorite(self, organ):
        """Imposta organo come preferito"""
        self.settings["favorite_organ"] = organ.get("file")
        self.save_settings()
        self.refresh_list_display()

    def remove_favorite(self):
        """Rimuove preferito"""
        self.settings["favorite_organ"] = None
        self.save_settings()
        self.refresh_list_display()

    def refresh_list_display(self):
        """Aggiorna display lista con stelle per preferiti"""
        current_selection = self.listbox.curselection()
        self.listbox.delete(0, "end")

        for organ in self.filtered_organs:
            is_favorite = (self.settings.get("favorite_organ") == organ.get("file"))
            prefix = "* " if is_favorite else "  "
            self.listbox.insert("end", f"{prefix}{organ['name']}")

        if current_selection and current_selection[0] < len(self.filtered_organs):
            self.listbox.selection_set(current_selection[0])

    def create_menu(self):
        """Crea menu"""
        menubar = tk.Menu(self.root)

        # File
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label=self.t["load"], command=self.load_orgue,
                            accelerator="Ctrl+O")
        filemenu.add_command(label=self.t["import_folder"],
                            command=self.import_folder, accelerator="Ctrl+I")
        filemenu.add_separator()
        filemenu.add_command(label=self.t["export"], command=self.export_config)
        filemenu.add_command(label=self.t["import"], command=self.import_config)
        filemenu.add_separator()
        filemenu.add_command(label=self.t["exit"], command=self.on_close,
                            accelerator="Esc")
        menubar.add_cascade(label=self.t["file"], menu=filemenu)

        # Visualizza
        viewmenu = tk.Menu(menubar, tearoff=0)
        viewmenu.add_command(label=self.t["stats"], command=self.show_statistics)
        viewmenu.add_command(label=self.t["images"],
                            command=self.open_images_folder)
        viewmenu.add_separator()
        viewmenu.add_command(label=self.t["settings"], command=self.show_settings)
        viewmenu.add_command(label=self.t["language"],
                            command=self.show_language_selector)
        viewmenu.add_command(label=self.t["skin"],
                            command=self.show_skin_selector)
        menubar.add_cascade(label=self.t["view"], menu=viewmenu)

        # Aiuto
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label=self.t["guide"], command=self.show_help)
        helpmenu.add_command(label=self.t["about"], command=self.show_about)
        menubar.add_cascade(label=self.t["help"], menu=helpmenu)

        self.root.config(menu=menubar)

    def adjust_color(self, hex_color, amount):
        """Schiarisce/scurisce colore"""
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        rgb = tuple(max(0, min(255, c + amount)) for c in rgb)
        return f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'

    def setup_keybindings(self):
        """Tasti rapidi"""
        self.root.bind("<Control-o>", lambda e: self.load_orgue())
        self.root.bind("<Control-i>", lambda e: self.import_folder())
        self.root.bind("<Delete>", lambda e: self.delete_organ())
        self.root.bind("<F1>", lambda e: self.show_help())
        self.root.bind("<F5>", lambda e: self.refresh_all())

    def save_config(self):
        """Salva configurazione organi"""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.organs, f, indent=2, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror(parent=self.root, title=self.t["error"], message=f"{self.t['cannot_save']}: {e}")

    def load_settings(self):
        """Carica impostazioni"""
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    self.settings.update(loaded)
            except:
                pass

    def save_settings(self):
        """Salva impostazioni"""
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=2)
        except Exception as e:
            messagebox.showerror(parent=self.root, title=self.t["error"], message=f"{self.t['cannot_save']}: {e}")

    def load_orgue(self):
        """Carica nuovo organo"""
        orgue_file = filedialog.askopenfilename(
            parent=self.root,
            title=self.t["select_organ_file"],
            filetypes=[("GrandOrgue files", "*.orgue *.organ"), ("All", "*.*")]
        )
        if not orgue_file:
            return

        dialog = tk.Toplevel(self.root)
        dialog.title(self.t["add_title"])
        dialog.configure(bg=self.skin["bg"])
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 450, 450)

        tk.Label(dialog, text=self.t["name"], bg=self.skin["bg"], fg=self.skin["fg"],
                font=(FONT_FAMILY, 9)).pack(pady=(15, 5))
        name_entry = tk.Entry(dialog, width=50, font=(FONT_FAMILY, 10),
                             bg=self.skin["list_bg"], fg=self.skin["list_fg"])
        name_entry.insert(0, os.path.splitext(os.path.basename(orgue_file))[0])
        name_entry.pack(pady=5)

        tk.Label(dialog, text=self.t["category"], bg=self.skin["bg"], fg=self.skin["fg"],
                font=(FONT_FAMILY, 9)).pack(pady=(10, 5))
        cat_var = tk.StringVar(value=self.t["other"])
        ttk.Combobox(dialog, textvariable=cat_var, values=self.categories[1:],
                    state="readonly", font=(FONT_FAMILY, 9)).pack(pady=5)

        tk.Label(dialog, text=self.t["desc"], bg=self.skin["bg"], fg=self.skin["fg"],
                font=(FONT_FAMILY, 9)).pack(pady=(10, 5))
        desc_text = tk.Text(dialog, width=50, height=5, font=(FONT_FAMILY, 9),
                           bg=self.skin["list_bg"], fg=self.skin["list_fg"])
        desc_text.pack(pady=5, padx=20)

        image_path = [None]

        def select_image():
            img = filedialog.askopenfilename(
                parent=self.root,
                title=self.t["select_image"],
                filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp")]
            )
            if img:
                image_path[0] = img
                messagebox.showinfo(parent=self.root, title=self.t["ok"], message=self.t["image_selected"])

        tk.Button(dialog, text=self.t["select_img"],
                 command=select_image,
                 bg=self.skin["button_bg"], fg=self.skin["button_fg"],
                 font=(FONT_FAMILY, 9),
                 relief="flat", padx=15, pady=8).pack(pady=10)

        btn_frame = tk.Frame(dialog, bg=self.skin["bg"])
        btn_frame.pack(pady=15)

        def save_organ():
            img_dest = None
            if image_path[0]:
                try:
                    ext = os.path.splitext(image_path[0])[1]
                    img_name = f"{name_entry.get().replace(' ', '_')}_{int(time.time())}{ext}"
                    img_dest = os.path.join(IMAGES_DIR, img_name)
                    shutil.copy(image_path[0], img_dest)
                except:
                    pass

            organ = {
                "name": name_entry.get(),
                "file": orgue_file,
                "image": img_dest,
                "desc": desc_text.get("1.0", "end-1c"),
                "category": cat_var.get(),
                "added": datetime.now().isoformat(),
                "launches": 0,
                "last_used": None
            }
            self.organs.append(organ)
            self.save_config()
            self.refresh_all()

            for i, o in enumerate(self.filtered_organs):
                if o == organ:
                    self.listbox.selection_clear(0, "end")
                    self.listbox.selection_set(i)
                    self.listbox.see(i)
                    self.show_preview()
                    break

            dialog.destroy()

        tk.Button(btn_frame, text=self.t["save"], command=save_organ,
                 bg=self.skin["accent"], fg="white",
                 font=(FONT_FAMILY, 10, "bold"),
                 relief="flat", padx=20, pady=10).pack(side="left", padx=5)

        tk.Button(btn_frame, text=self.t["cancel"], command=dialog.destroy,
                 bg="#6c757d", fg="white",
                 font=(FONT_FAMILY, 10, "bold"),
                 relief="flat", padx=20, pady=10).pack(side="left", padx=5)

    def import_folder(self):
        """Importa organi da cartella"""
        folder = filedialog.askdirectory(parent=self.root, title=self.t["import_folder"])
        if not folder:
            return

        count = 0
        for root, dirs, files in os.walk(folder):
            for file in files:
                if file.lower().endswith(('.organ', '.orgue')):
                    filepath = os.path.join(root, file)
                    if any(o['file'] == filepath for o in self.organs):
                        continue

                    organ = {
                        "name": os.path.splitext(file)[0],
                        "file": filepath,
                        "image": None,
                        "desc": "",
                        "category": self.t["other"],
                        "added": datetime.now().isoformat(),
                        "launches": 0,
                        "last_used": None
                    }
                    self.organs.append(organ)
                    count += 1

        self.save_config()
        self.refresh_all()
        messagebox.showinfo(parent=self.root, title=self.t["success"],
                           message=f"{self.t['imported']} {count} {self.t['organs_imported']}!")

    def filter_organs(self, *args):
        """Filtra organi per ricerca e categoria"""
        search = self.search_var.get().lower()
        category = self.category_var.get()

        self.filtered_organs = [
            o for o in self.organs
            if (search in o['name'].lower() or search in o.get('desc', '').lower())
            and (category == self.t["all"] or o.get('category', self.t['other']) == category)
        ]

        self.refresh_list_display()
        self.update_stats()

    def show_preview(self, event=None):
        """Mostra anteprima con possibilita di aggiungere screenshot"""
        selection = self.listbox.curselection()
        if not selection:
            return

        organ = self.filtered_organs[selection[0]]

        # Gestione immagine
        if organ.get("image") and os.path.exists(organ["image"]):
            try:
                if HAS_PIL:
                    img = Image.open(organ["image"])
                    w = self.preview_label.winfo_width() or 800
                    h = self.preview_label.winfo_height() or 600
                    img.thumbnail((w-40, h-40), Image.LANCZOS)
                    self.imgtk = ImageTk.PhotoImage(img)
                else:
                    self.imgtk = tk.PhotoImage(file=organ["image"])
                self.preview_label.config(image=self.imgtk, text="", cursor="hand2")
                self.preview_label.bind("<Button-1>",
                                       lambda e: self.add_or_change_image(organ))
            except:
                self.preview_label.config(image="",
                                         text=self.t["error_loading_image"],
                                         cursor="hand2")
                self.preview_label.bind("<Button-1>",
                                       lambda e: self.add_or_change_image(organ))
        else:
            self.preview_label.config(
                image="",
                text=self.t["no_img"],
                font=(FONT_FAMILY, 14),
                cursor="hand2",
                fg="#888888"
            )
            self.preview_label.bind("<Button-1>",
                                   lambda e: self.add_or_change_image(organ))

        # Info organo
        is_favorite = (self.settings.get("favorite_organ") == organ.get("file"))
        name_text = f"* {organ['name']}" if is_favorite else organ["name"]
        self.name_label.config(text=name_text)
        self.category_label.config(text=f"{organ.get('category', self.t['other'])}")
        self.desc_label.config(text=organ.get("desc", ""))

        launches = organ.get("launches", 0)
        last = organ.get("last_used", self.t["never"])
        if last and last != self.t["never"]:
            try:
                last = datetime.fromisoformat(last).strftime("%d/%m/%Y %H:%M")
            except:
                pass

        self.stats_organ_label.config(
            text=f"{self.t['launched']} {launches} {self.t['times']} | {self.t['last']}: {last}"
        )

    def add_or_change_image(self, organ):
        """Menu per aggiungere/cambiare/rimuovere immagine"""
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label=self.t["select_img"],
                        command=lambda: self.select_image_for_organ(organ))
        menu.add_command(label=self.t["capture"],
                        command=lambda: self.capture_screenshot_for_organ(organ))
        if organ.get("image"):
            menu.add_separator()
            menu.add_command(label=self.t["remove_img"],
                            command=lambda: self.remove_image_from_organ(organ))

        try:
            menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        finally:
            menu.grab_release()

    def select_image_for_organ(self, organ):
        """Seleziona file immagine esistente"""
        img = filedialog.askopenfilename(
            parent=self.root,
            title=self.t["select_image"],
            filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp")]
        )
        if img:
            try:
                ext = os.path.splitext(img)[1]
                img_name = f"{organ['name'].replace(' ', '_')}_{int(time.time())}{ext}"
                img_dest = os.path.join(IMAGES_DIR, img_name)
                shutil.copy(img, img_dest)

                if organ.get("image") and os.path.exists(organ["image"]):
                    try:
                        os.remove(organ["image"])
                    except:
                        pass

                idx = self.organs.index(organ)
                self.organs[idx]["image"] = img_dest
                self.save_config()
                self.show_preview()

            except Exception as e:
                messagebox.showerror(parent=self.root, title=self.t["error"], message=f"{self.t['cannot_save']}:\n{e}")

    def launch_orgue(self, event=None):
        """Avvia organo"""
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showwarning(parent=self.root, title=self.t["warning"], message=self.t["select"])
            return

        organ = self.filtered_organs[selection[0]]

        if not os.path.exists(organ["file"]):
            messagebox.showerror(parent=self.root, title=self.t["error"], message=f"{self.t['file_not_found']}:\n{organ['file']}")
            return

        # Chiudi GrandOrgue precedente
        try:
            if os.name == "nt":
                subprocess.Popen(["taskkill", "/IM", "GrandOrgue.exe", "/F", "/T"],
                               stdout=subprocess.DEVNULL,
                               stderr=subprocess.DEVNULL,
                               creationflags=subprocess.CREATE_NO_WINDOW)
            else:
                subprocess.Popen(["pkill", "-9", "GrandOrgue"],
                               stdout=subprocess.DEVNULL,
                               stderr=subprocess.DEVNULL)
        except:
            pass

        self.root.after(100, lambda: self._do_launch(organ))

    def _do_launch(self, organ):
        """Esegue lancio organo"""
        try:
            if os.name == "nt":
                subprocess.Popen([organ["file"]], shell=True,
                               creationflags=subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS)
            else:
                subprocess.Popen(["xdg-open", organ["file"]],
                               stdout=subprocess.DEVNULL,
                               stderr=subprocess.DEVNULL)

            idx = self.organs.index(organ)
            self.organs[idx]["launches"] = organ.get("launches", 0) + 1
            self.organs[idx]["last_used"] = datetime.now().isoformat()
            self.save_config()
            self.show_preview()

        except Exception as e:
            messagebox.showerror(parent=self.root, title=self.t["error"], message=f"{self.t['cannot_launch']}:\n{e}")

    def move_up(self):
        """Sposta su (aumenta priorita)"""
        selection = self.listbox.curselection()
        if not selection or selection[0] == 0:
            return

        idx = self.organs.index(self.filtered_organs[selection[0]])
        if idx > 0:
            self.organs[idx-1], self.organs[idx] = self.organs[idx], self.organs[idx-1]
            self.save_config()
            new_pos = selection[0] - 1
            self.filter_organs()
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set(new_pos)
            self.listbox.see(new_pos)
            self.listbox.activate(new_pos)

    def move_down(self):
        """Sposta giu (diminuisce priorita)"""
        selection = self.listbox.curselection()
        if not selection:
            return

        idx = self.organs.index(self.filtered_organs[selection[0]])
        if idx < len(self.organs) - 1:
            self.organs[idx+1], self.organs[idx] = self.organs[idx], self.organs[idx+1]
            self.save_config()
            new_pos = selection[0] + 1
            self.filter_organs()
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set(new_pos)
            self.listbox.see(new_pos)
            self.listbox.activate(new_pos)

    def delete_organ(self):
        """Elimina organo"""
        selection = self.listbox.curselection()
        if not selection:
            return

        organ = self.filtered_organs[selection[0]]
        if messagebox.askyesno(parent=self.root, title=self.t["confirm_delete"], message=f"{self.t['confirm_delete']} '{organ['name']}'?"):
            self.organs.remove(organ)
            self.save_config()
            self.refresh_all()

    def edit_organ(self):
        """Modifica organo"""
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showwarning(parent=self.root, title=self.t["warning"], message=self.t["select"])
            return

        organ = self.filtered_organs[selection[0]]

        dialog = tk.Toplevel(self.root)
        dialog.title(self.t["edit_organ"])
        dialog.configure(bg=self.skin["bg"])
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 400, 350)

        tk.Label(dialog, text=self.t["name_label"], bg=self.skin["bg"],
                fg=self.skin["fg"]).pack(pady=5)
        name_entry = tk.Entry(dialog, width=40, bg=self.skin["list_bg"],
                             fg=self.skin["list_fg"])
        name_entry.insert(0, organ["name"])
        name_entry.pack(pady=5)

        tk.Label(dialog, text=self.t["category_label"], bg=self.skin["bg"],
                fg=self.skin["fg"]).pack(pady=5)
        cat_var = tk.StringVar(value=organ.get("category", self.t["other"]))
        ttk.Combobox(dialog, textvariable=cat_var,
                    values=self.categories[1:], state="readonly").pack(pady=5)

        tk.Label(dialog, text=self.t["description_label"], bg=self.skin["bg"],
                fg=self.skin["fg"]).pack(pady=5)
        desc_text = tk.Text(dialog, width=40, height=6, bg=self.skin["list_bg"],
                           fg=self.skin["list_fg"])
        desc_text.insert("1.0", organ.get("desc", ""))
        desc_text.pack(pady=5)

        def save_changes():
            idx = self.organs.index(organ)
            self.organs[idx]["name"] = name_entry.get()
            self.organs[idx]["category"] = cat_var.get()
            self.organs[idx]["desc"] = desc_text.get("1.0", "end-1c")
            self.save_config()
            self.refresh_all()
            dialog.destroy()

        tk.Button(dialog, text=self.t["save"], command=save_changes,
                 bg=self.skin["accent"], fg="white").pack(pady=10)

    def remove_image_from_organ(self, organ):
        """Rimuovi immagine"""
        if messagebox.askyesno(parent=self.root, title=self.t["confirm_delete"], message=self.t["confirm_remove_image"]):
            if organ.get("image") and os.path.exists(organ["image"]):
                try:
                    os.remove(organ["image"])
                except:
                    pass

            idx = self.organs.index(organ)
            self.organs[idx]["image"] = None
            self.save_config()
            self.show_preview()

    def update_stats(self):
        """Aggiorna statistiche header"""
        total = len(self.organs)
        filtered = len(self.filtered_organs)
        text = f"{filtered}/{total} {self.t['count']}"
        self.stats_label.config(text=text)

    def refresh_all(self):
        """Aggiorna tutto"""
        current_selection = self.listbox.curselection()
        self.filter_organs()

        if self.filtered_organs:
            if current_selection and current_selection[0] < len(self.filtered_organs):
                self.listbox.selection_set(current_selection[0])
            else:
                self.listbox.selection_set(0)
            self.show_preview()

    def capture_screenshot_for_organ(self, organ):
        """Cattura screenshot con countdown 5 sec"""
        info = tk.Toplevel(self.root)
        info.title(self.t["capture_title"])
        info.configure(bg=self.skin["bg"])
        info.attributes("-topmost", True)
        info.transient(self.root)
        info.grab_set()
        self.center_window(info, 500, 350)

        images_path = os.path.abspath(IMAGES_DIR)

        tk.Label(info,
                text=f"{self.t['capture_title']}\n{organ['name']}\n\n"
                     f"{self.t['capture_procedure']}\n"
                     f"{self.t['capture_step1']}\n"
                     f"{self.t['capture_step2']}\n"
                     f"{self.t['capture_step3']}\n"
                     f"{self.t['capture_step4']}\n"
                     f"{self.t['capture_step5']}\n"
                     f"{self.t['capture_step6']}\n\n"
                     f"{self.t['images_note']}\n{images_path}",
                bg=self.skin["bg"], fg=self.skin["fg"],
                font=(FONT_FAMILY, 9),
                justify="left").pack(pady=20, padx=20)

        btn_frame = tk.Frame(info, bg=self.skin["bg"])
        btn_frame.pack(pady=10)

        def launch_first():
            info.withdraw()
            self.launch_orgue()
            self.root.after(1000, lambda: info.deiconify())

        def do_capture():
            info.withdraw()

            countdown_win = tk.Toplevel(self.root)
            countdown_win.title("Countdown")
            countdown_win.configure(bg="#ff6b6b")
            countdown_win.attributes("-topmost", True)
            countdown_win.overrideredirect(True)
            self.center_window(countdown_win, 400, 150)

            msg_label = tk.Label(countdown_win,
                                text=self.t["position_window"].upper(),
                                bg="#ff6b6b", fg="white",
                                font=(FONT_FAMILY, 12, "bold"))
            msg_label.pack(pady=10)

            countdown_label = tk.Label(countdown_win,
                                      text="",
                                      bg="#ff6b6b", fg="white",
                                      font=(FONT_FAMILY, 48, "bold"))
            countdown_label.pack(expand=True)

            for i in range(5, 0, -1):
                countdown_label.config(text=str(i))
                countdown_win.update()
                time.sleep(1)

            countdown_win.destroy()

            try:
                # Cattura solo il monitor dove si trova l'applicazione
                img_name_base = f"{organ['name'].replace(' ', '_')}_screenshot_{int(time.time())}"
                if HAS_PIL:
                    monitor_bbox = self.get_current_monitor_bbox()
                    screenshot = ImageGrab.grab(bbox=monitor_bbox)
                    ext = ".jpg"
                    img_name = img_name_base + ext
                    img_dest = os.path.join(IMAGES_DIR, img_name)
                    screenshot = screenshot.convert("RGB")
                    screenshot.save(img_dest, "JPEG", quality=85)
                else:
                    ext = ".png"
                    img_name = img_name_base + ext
                    img_dest = os.path.join(IMAGES_DIR, img_name)
                    captured = False
                    for cmd in [
                        ["gnome-screenshot", "-f", img_dest],
                        ["scrot", img_dest],
                        ["import", "-window", "root", img_dest],
                    ]:
                        if shutil.which(cmd[0]):
                            subprocess.run(cmd, check=True)
                            captured = True
                            break
                    if not captured:
                        raise FileNotFoundError(
                            "No screenshot tool found.\n"
                            "Install one: sudo apt install gnome-screenshot\n"
                            "or: sudo apt install scrot\n"
                            "or: pip install Pillow"
                        )

                if organ.get("image") and os.path.exists(organ["image"]):
                    try:
                        os.remove(organ["image"])
                    except:
                        pass

                idx = self.organs.index(organ)
                self.organs[idx]["image"] = img_dest
                self.save_config()

                info.destroy()
                self.show_preview()

                messagebox.showinfo(parent=self.root, title=self.t["ok"],
                    message=f"{self.t['screenshot_captured']}\n\n"
                    f"{self.t['crop_tip']}\n"
                    f"{self.t['images_in']}: {images_path}")

            except Exception as e:
                messagebox.showerror(parent=self.root, title=self.t["error"], message=f"{self.t['cannot_capture']}:\n{e}")
                info.deiconify()

        tk.Button(btn_frame, text=self.t["launch_organ_btn"],
                 command=launch_first,
                 bg=self.skin["btn_edit"], fg="white",
                 font=(FONT_FAMILY, 9, "bold"),
                 relief="flat", padx=15, pady=8).pack(side="left", padx=5)

        tk.Button(btn_frame, text=self.t["capture_btn"],
                 command=do_capture,
                 bg=self.skin["accent"], fg="white",
                 font=(FONT_FAMILY, 9, "bold"),
                 relief="flat", padx=15, pady=8).pack(side="left", padx=5)

        tk.Button(btn_frame, text=self.t["cancel"],
                 command=info.destroy,
                 bg="#6c757d", fg="white",
                 font=(FONT_FAMILY, 9, "bold"),
                 relief="flat", padx=15, pady=8).pack(side="left", padx=5)

    def open_images_folder(self):
        """Apri cartella immagini"""
        images_path = os.path.abspath(IMAGES_DIR)
        try:
            if os.name == "nt":
                os.startfile(images_path)
            elif os.name == "darwin":
                subprocess.Popen(["open", images_path])
            else:
                subprocess.Popen(["xdg-open", images_path])
        except Exception as e:
            messagebox.showerror(parent=self.root, title=self.t["error"], message=f"{self.t['cannot_open_folder']}:\n{e}")

    def export_config(self):
        """Export completo con ZIP"""
        export_file = filedialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".zip",
            filetypes=[("ZIP Archive", "*.zip"), ("All files", "*.*")],
            initialfile=f"grandorgue_config_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        )

        if not export_file:
            return

        try:
            with zipfile.ZipFile(export_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                zipf.write(CONFIG_FILE, arcname=CONFIG_FILE)

                if os.path.exists(SETTINGS_FILE):
                    zipf.write(SETTINGS_FILE, arcname=SETTINGS_FILE)

                for organ in self.organs:
                    if organ.get("image") and os.path.exists(organ["image"]):
                        zipf.write(organ["image"],
                                 arcname=os.path.join(IMAGES_DIR,
                                                     os.path.basename(organ["image"])))

            messagebox.showinfo(parent=self.root, title=self.t["success"],
                message=f"{self.t['export_success']}\n\n"
                f"File: {os.path.basename(export_file)}\n"
                f"{self.t['count'].capitalize()}: {len(self.organs)}")

        except Exception as e:
            messagebox.showerror(parent=self.root, title=self.t["error"], message=f"{self.t['cannot_export']}:\n{e}")

    def import_config(self):
        """Import completo da ZIP"""
        import_file = filedialog.askopenfilename(
            parent=self.root,
            title=self.t["import"],
            filetypes=[("ZIP Archive", "*.zip"), ("JSON", "*.json"), ("All", "*.*")]
        )

        if not import_file:
            return

        try:
            if import_file.endswith('.zip'):
                with zipfile.ZipFile(import_file, 'r') as zipf:
                    if CONFIG_FILE in zipf.namelist():
                        zipf.extract(CONFIG_FILE, path=".")

                    if SETTINGS_FILE in zipf.namelist():
                        zipf.extract(SETTINGS_FILE, path=".")

                    for file in zipf.namelist():
                        if file.startswith(IMAGES_DIR + "/"):
                            zipf.extract(file, path=".")

                self.load_settings()
                self.organs = []
                self.load_config()

                messagebox.showinfo(parent=self.root, title=self.t["success"],
                    message=f"{self.t['import_success']}\n\n"
                    f"{self.t['count'].capitalize()}: {len(self.organs)}")

            else:
                with open(import_file, "r", encoding="utf-8") as f:
                    imported = json.load(f)
                self.organs.extend(imported)
                self.save_config()
                self.refresh_all()
                messagebox.showinfo(parent=self.root, title=self.t["success"],
                    message=f"{self.t['imported']} {len(imported)} {self.t['organs_imported']}!")

        except Exception as e:
            messagebox.showerror(parent=self.root, title=self.t["error"], message=f"{self.t['cannot_import']}:\n{e}")

    def show_language_selector(self):
        """Selettore lingua"""
        dialog = tk.Toplevel(self.root)
        dialog.title(self.t["select_lang"])
        dialog.configure(bg=self.skin["bg"])
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 350, 450)

        tk.Label(dialog, text=f"{self.t['language']}",
                bg=self.skin["bg"], fg=self.skin["fg"],
                font=(FONT_FAMILY, 14, "bold")).pack(pady=20)

        lang_frame = tk.Frame(dialog, bg=self.skin["bg"])
        lang_frame.pack(pady=10)

        languages = [
            ("en", "English"),
            ("it", "Italiano"),
            ("fr", "Francais"),
            ("de", "Deutsch"),
            ("es", "Espanol")
        ]

        for lang_code, lang_name in languages:
            flag = LANG_FLAGS[lang_code]
            is_current = (lang_code == self.settings["language"])

            btn = tk.Button(lang_frame,
                          text=f"[{flag}]  {lang_name}" + (" *" if is_current else ""),
                          command=lambda lc=lang_code: self.change_language(lc, dialog),
                          bg=self.skin["accent"] if is_current else self.skin["button_bg"],
                          fg="white" if is_current else self.skin["button_fg"],
                          font=(FONT_FAMILY, 11, "bold" if is_current else "normal"),
                          relief="flat",
                          cursor="hand2",
                          width=20,
                          padx=15, pady=12)
            btn.pack(pady=5)

            if not is_current:
                btn.bind("<Enter>", lambda e, b=btn: b.config(bg=self.skin["hover"]))
                btn.bind("<Leave>", lambda e, b=btn: b.config(bg=self.skin["button_bg"]))

    def change_language(self, lang_code, dialog):
        """Cambia lingua"""
        self.settings["language"] = lang_code
        self.save_settings()

        messagebox.showinfo(parent=self.root, title=self.t["success"],
            message=f"{self.t['lang_changed']}\n{self.t['restart_app']}")

        dialog.destroy()

    def show_skin_selector(self):
        """Selettore tema/skin"""
        dialog = tk.Toplevel(self.root)
        dialog.title(self.t["select_skin"])
        dialog.configure(bg=self.skin["bg"])
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 400, 400)

        tk.Label(dialog, text=f"{self.t['skin']}",
                bg=self.skin["bg"], fg=self.skin["fg"],
                font=(FONT_FAMILY, 14, "bold")).pack(pady=20)

        skin_frame = tk.Frame(dialog, bg=self.skin["bg"])
        skin_frame.pack(pady=10)

        for skin_id, skin_data in SKINS.items():
            is_current = (skin_id == self.settings.get("skin", "classic"))
            skin_name = self.t[skin_data["name_key"]]

            # Preview colori
            preview_frame = tk.Frame(skin_frame, bg=self.skin["bg"])
            preview_frame.pack(fill="x", pady=5, padx=20)

            # Campione colori
            color_sample = tk.Frame(preview_frame, bg=skin_data["bg"],
                                   width=30, height=30, relief="solid", bd=1)
            color_sample.pack(side="left", padx=5)
            color_sample.pack_propagate(False)

            accent_sample = tk.Frame(color_sample, bg=skin_data["accent"],
                                    width=15, height=15)
            accent_sample.place(relx=0.5, rely=0.5, anchor="center")

            btn = tk.Button(preview_frame,
                          text=f"{skin_name}" + (" *" if is_current else ""),
                          command=lambda sid=skin_id: self.change_skin(sid, dialog),
                          bg=self.skin["accent"] if is_current else self.skin["button_bg"],
                          fg="white" if is_current else self.skin["button_fg"],
                          font=(FONT_FAMILY, 11, "bold" if is_current else "normal"),
                          relief="flat",
                          cursor="hand2",
                          width=25,
                          padx=15, pady=10)
            btn.pack(side="left", padx=10)

            if not is_current:
                btn.bind("<Enter>", lambda e, b=btn: b.config(bg=self.skin["hover"]))
                btn.bind("<Leave>", lambda e, b=btn: b.config(bg=self.skin["button_bg"]))

    def change_skin(self, skin_id, dialog):
        """Cambia tema"""
        self.settings["skin"] = skin_id
        self.save_settings()

        messagebox.showinfo(parent=self.root, title=self.t["success"],
            message=f"{self.t['lang_changed']}\n{self.t['restart_app']}")

        dialog.destroy()

    def show_settings(self):
        """Finestra impostazioni"""
        dialog = tk.Toplevel(self.root)
        dialog.title(self.t["settings"])
        dialog.configure(bg=self.skin["bg"])
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 550, 450)

        auto_frame = tk.LabelFrame(dialog, text=self.t["auto_title"],
                                   bg=self.skin["bg"], fg=self.skin["fg"],
                                   font=(FONT_FAMILY, 10, "bold"),
                                   padx=20, pady=15)
        auto_frame.pack(fill="x", padx=20, pady=20)

        auto_var = tk.BooleanVar(value=self.settings["auto_launch"])

        tk.Checkbutton(auto_frame, text=self.t["auto_desc"],
                      variable=auto_var, bg=self.skin["bg"], fg=self.skin["fg"],
                      selectcolor=self.skin["bg"],
                      activebackground=self.skin["bg"],
                      activeforeground=self.skin["fg"],
                      font=(FONT_FAMILY, 9)).pack(anchor="w", pady=5)

        # Mostra organo preferito corrente
        fav_frame = tk.Frame(auto_frame, bg=self.skin["bg"])
        fav_frame.pack(fill="x", pady=5)

        fav_name = self.t["no_favorite"]
        if self.settings.get("favorite_organ"):
            for o in self.organs:
                if o.get("file") == self.settings["favorite_organ"]:
                    fav_name = f"* {o['name']}"
                    break

        tk.Label(fav_frame, text=f"{self.t['favorite_organ']} {fav_name}",
                bg=self.skin["bg"], fg="#f0ad4e",
                font=(FONT_FAMILY, 9, "italic")).pack(anchor="w")

        delay_frame = tk.Frame(auto_frame, bg=self.skin["bg"])
        delay_frame.pack(fill="x", pady=10)

        tk.Label(delay_frame, text=self.t["delay"],
                bg=self.skin["bg"], fg=self.skin["fg"],
                font=(FONT_FAMILY, 9)).pack(side="left")

        delay_var = tk.IntVar(value=self.settings["auto_launch_delay"])
        tk.Spinbox(delay_frame, from_=3, to=30, textvariable=delay_var,
                  width=5, font=(FONT_FAMILY, 9)).pack(side="left", padx=10)

        tk.Label(delay_frame, text=self.t["min_3"],
                bg=self.skin["bg"], fg="#888888",
                font=(FONT_FAMILY, 8)).pack(side="left")

        info_frame = tk.Frame(auto_frame, bg=self.skin["info_bg"],
                             relief="solid", borderwidth=1)
        info_frame.pack(fill="x", pady=10, padx=5)

        tk.Label(info_frame,
                text=f"{self.t['how_it_works']}:\n\n"
                     f"- {self.t['auto_info_1']}\n"
                     f"- {self.t['auto_info_2']}\n"
                     f"- {self.t['auto_info_3']}\n"
                     f"- {self.t['auto_info_4']}",
                bg=self.skin["info_bg"], fg="#f0ad4e",
                font=(FONT_FAMILY, 8), justify="left",
                padx=10, pady=10).pack()

        btn_frame = tk.Frame(dialog, bg=self.skin["bg"])
        btn_frame.pack(fill="x", padx=20, pady=10)

        def save_settings():
            delay_value = delay_var.get()
            if delay_value < 3:
                messagebox.showwarning(parent=self.root, title=self.t["warning"], message=self.t["min_delay_warning"])
                return

            self.settings["auto_launch"] = auto_var.get()
            self.settings["auto_launch_delay"] = delay_value
            self.settings["first_run"] = False
            self.save_settings()
            messagebox.showinfo(parent=self.root, title=self.t["ok"], message=self.t["settings_saved"])
            dialog.destroy()

        tk.Button(btn_frame, text=self.t["save"], command=save_settings,
                 bg=self.skin["accent"], fg="white",
                 font=(FONT_FAMILY, 9, "bold"),
                 relief="flat", padx=20, pady=8).pack(side="left", padx=5)

        tk.Button(btn_frame, text=self.t["cancel"], command=dialog.destroy,
                 bg="#6c757d", fg="white",
                 font=(FONT_FAMILY, 9, "bold"),
                 relief="flat", padx=20, pady=8).pack(side="left", padx=5)

    def show_statistics(self):
        """Mostra statistiche"""
        total = len(self.organs)
        total_launches = sum(o.get("launches", 0) for o in self.organs)

        by_category = {}
        for o in self.organs:
            cat = o.get("category", self.t["other"])
            by_category[cat] = by_category.get(cat, 0) + 1

        most_used = sorted(self.organs,
                          key=lambda x: x.get("launches", 0),
                          reverse=True)[:5]

        msg = f"{self.t['stats'].upper()}\n\n"
        msg += f"{self.t['total_organs']}: {total}\n"
        msg += f"{self.t['total_launches']}: {total_launches}\n\n"
        msg += f"{self.t['by_category']}:\n"
        for cat, count in by_category.items():
            msg += f"  - {cat}: {count}\n"
        msg += f"\n{self.t['most_used']}:\n"
        for i, o in enumerate(most_used, 1):
            msg += f"  {i}. {o['name']} ({o.get('launches', 0)}x)\n"

        messagebox.showinfo(parent=self.root, title=self.t["stats"], message=msg)

    def show_help(self):
        """Guida"""
        help_text = f"""
{self.t['guide'].upper()}

KEYBOARD:
- Ctrl+O: {self.t['load']}
- Ctrl+I: {self.t['import_folder']}
- Enter/Double-click: {self.t['launch']}
- Delete: {self.t['delete']}
- F1: {self.t['guide']}
- F5: Refresh

FEATURES:
- {self.t['search']}: Filter
- {self.t['category']}: Organize
- {self.t['stats']}: Usage
- Export/Import: Backup (ZIP)
- Multi-language: 5 languages
- Auto-launch: Favorite organ
- Themes: 3 skins available

VERSION: {VERSION}
        """
        messagebox.showinfo(parent=self.root, title=self.t["guide"], message=help_text)

    def show_about(self):
        """Info app"""
        messagebox.showinfo(parent=self.root, title=self.t["about"],
            message=f"GrandOrgue Selector\n"
            f"Version {VERSION}\n\n"
            f"Visual manager for GrandOrgue virtual organs\n\n"
            f"Author: Gabriele Bastianelli\n"
            f"Urbino, Italy\n"
            f"bastigab@gmail.com\n\n"
            f"Developed with Python + Tkinter\n"
            f"2024\n\n"
            f"Language: [{LANG_FLAGS[self.settings['language']]}] "
            f"{self.settings['language'].upper()}\n"
            f"Theme: {self.t[SKINS[self.settings.get('skin', 'classic')]['name_key']]}")

    def load_config(self):
        """Carica configurazione"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.organs = json.load(f)
                self.refresh_all()

                if self.filtered_organs:
                    self.listbox.selection_clear(0, "end")
                    self.listbox.selection_set(0)
                    self.listbox.activate(0)
                    self.listbox.focus_set()
                    self.root.after(200, self.show_preview)

                    if self.settings["auto_launch"] and not self.settings["first_run"]:
                        self.start_auto_launch_countdown()
                    else:
                        if self.settings["first_run"]:
                            self.settings["first_run"] = False
                            self.save_settings()
            except Exception as e:
                messagebox.showerror(parent=self.root, title=self.t["error"], message=f"{self.t['cannot_load']}: {e}")
                self.organs = []

    def start_auto_launch_countdown(self):
        """Countdown auto-launch con finestra grande"""
        # Trova organo preferito
        favorite_organ = None
        if self.settings.get("favorite_organ"):
            for o in self.organs:
                if o.get("file") == self.settings["favorite_organ"]:
                    favorite_organ = o
                    break

        if not favorite_organ:
            # Se non c'e preferito, usa il primo
            if self.filtered_organs:
                favorite_organ = self.filtered_organs[0]
            else:
                return

        self.countdown_active = True
        delay = self.settings["auto_launch_delay"]
        self.countdown_remaining = delay

        # Crea finestra countdown grande
        self.countdown_window = tk.Toplevel(self.root)
        self.countdown_window.title("Auto-Launch")
        self.countdown_window.configure(bg="#ff6b6b")
        self.countdown_window.attributes("-topmost", True)
        self.countdown_window.overrideredirect(True)
        self.center_window(self.countdown_window, 500, 300)

        # Contenuto
        tk.Label(self.countdown_window,
                text=f"{self.t['auto_launch_in']}...",
                bg="#ff6b6b", fg="white",
                font=(FONT_FAMILY, 14, "bold")).pack(pady=20)

        self.countdown_number_label = tk.Label(self.countdown_window,
                                              text=str(delay),
                                              bg="#ff6b6b", fg="white",
                                              font=(FONT_FAMILY, 72, "bold"))
        self.countdown_number_label.pack(pady=10)

        tk.Label(self.countdown_window,
                text=f"* {favorite_organ['name']}",
                bg="#ff6b6b", fg="yellow",
                font=(FONT_FAMILY, 12, "bold")).pack(pady=10)

        tk.Label(self.countdown_window,
                text=self.t["press_any_key"],
                bg="#ff6b6b", fg="white",
                font=(FONT_FAMILY, 10)).pack(pady=10)

        # Bind QUALSIASI tasto per cancellare
        self.root.bind("<Key>", self.cancel_auto_launch)
        self.root.bind("<Button-1>", self.cancel_auto_launch)
        self.root.bind("<Escape>", self.cancel_auto_launch)
        self.countdown_window.bind("<Key>", self.cancel_auto_launch)
        self.countdown_window.bind("<Button-1>", self.cancel_auto_launch)

        def update_countdown():
            if not self.countdown_active:
                return

            if self.countdown_remaining > 0:
                self.countdown_number_label.config(text=str(self.countdown_remaining))
                self.countdown_remaining -= 1
                self.auto_launch_timer = self.root.after(1000, update_countdown)
            else:
                self.countdown_active = False
                if self.countdown_window:
                    self.countdown_window.destroy()
                    self.countdown_window = None
                self.unbind_countdown_keys()

                # Seleziona e avvia l'organo preferito
                for i, o in enumerate(self.filtered_organs):
                    if o.get("file") == favorite_organ.get("file"):
                        self.listbox.selection_clear(0, "end")
                        self.listbox.selection_set(i)
                        self.listbox.see(i)
                        break

                self._do_launch(favorite_organ)

        update_countdown()

    def cancel_auto_launch(self, event=None):
        """Cancella auto-launch"""
        if not self.countdown_active:
            return

        self.countdown_active = False

        if self.auto_launch_timer:
            self.root.after_cancel(self.auto_launch_timer)
            self.auto_launch_timer = None

        if self.countdown_window:
            self.countdown_window.destroy()
            self.countdown_window = None

        self.countdown_label.config(text=f"{self.t['cancelled']}")
        self.root.after(2000, lambda: self.countdown_label.config(text=""))

        self.unbind_countdown_keys()

    def unbind_countdown_keys(self):
        """Rimuove binding tasti countdown"""
        self.root.unbind("<Key>")
        self.root.unbind("<Button-1>")
        self.root.bind("<Escape>", lambda e: self.on_close())

    def on_close(self):
        """Chiusura"""
        if self.auto_launch_timer:
            self.root.after_cancel(self.auto_launch_timer)
        self.countdown_active = False
        if self.countdown_window:
            self.countdown_window.destroy()
        self.save_config()
        self.root.destroy()


# MAIN
if __name__ == "__main__":
    root = tk.Tk()
    app = OrgansApp(root)
    root.mainloop()
