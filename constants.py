"""
Here we will load all the constants to the program. Therefore it is more efficient to find the constants and to change
them.
@ Fabio serra Pereira
contact: serrafabio10@outlook.com
"""
import os
from pathlib import Path
# Time of varification
OFFICE = 24
LABOR = 12
LIMIT_UNDO_ACTIONS = 200
# Layout
WIDTH = 50
HEIGHT = 50
MAIN_WINDOW_WIDTH = 1010
MAIN_WINDOW_HEIGHT = 650
NAME_OF_PROGRAM = 'EPrüfung'
DESKTOP = 'Desktop'
# Table View
MINIMU_HEIGHT_TABLE_VIEW_MAIN_WINDOW = 400
MINIMU_WIDTH_TABLE_VIEW_MAIN_WINDOW = 600
#Dialogs
DIALOG_PROBLEM_WIDHT = 400
DIALOG_OPERATION_CONCLUEDED_WIDHT= 200
DIALOG_PROBLEM_HEIGHT = 150
DIALOG_HEIGHT_ID = 132
DIALOG_WIDTH_ID = 334
LINE_EDITOR_HEIGHT = 30
DIALOG_HEIGHT_CALENDER = 200
DIALOG_WIDTH_CALENDER = 300
DIALOG_ADD_NEW_LOCATION_WIDTH = 525
DIALOG_ADD_NEW_LOCATION_HEIGTH = 200
DIALOG_ADD_NEW_LOCATION_COMBOBOX = 230
DIALOG_ADD_NEW_LOCATION_LINEEDIT = 30
DIALOG_DELETE_LOCATION_WIDTH = 255
DIALOG_DELETE_LOCATION_HEIGHT = 160
DIALOG_REFACTOR_LOCATION_HEIGHT = 185
DIALOG_BUTTONS_WIDTH = 100
DIALOG_SHOW_RESULTS_WIDTH = 500
DIALOG_SHOW_RESULTS_HEIGHT = 400
DIALOG_WIDTH_EMAIL = 300
DIALOG_HEIGHT_EMAIL = 100
DIALOG_SIZE_WIDTH_DATE, DIALOG_SIZE_HEIGHT_DATE = 250, 30
DIALOG_WIDTH_NOTVERFIED_BUTTON = 350
DIALOG_SIZE_CHECKBOX = 75,40
DIALOG_WIDTH_URLACTUALIZE = 600
#Minimums of the bottom
MINIMUM_WIDTH_MAIN_BUTTONS = 400
MINIMUM_WIDTH_TYPE_SELECTION = 250
#Font
FONT = 'Arial'
SIZE_FONT = 11
# ICONS
ICON_TU_DARMSTADT = os.getcwd() + os.sep +'Icon' + os.sep + 'TU_Darmstadt.ico'
ICON_PTW_LOGO = os.getcwd() + os.sep +'Icon'+ os.sep +'ptw_logo.ico'
ICON_UNDO = os.getcwd() + os.sep +'Icon'+ os.sep +'undo-arrow.ico'
# Data Base
NAME_OFFICIAL_DATA_BASE= os.getcwd() + os.sep + 'official_db.db'
NAME_CURRENT_DATA_BASE = os.getcwd() + os.sep +'memory.db'
NAME_RECOVERY_DATA_BASE = os.getcwd() + os.sep + 'undo' + os.sep +'recovery'
NAME_DB_END = '.db'
VARIABLES_WANTED_IN_TABLE_VIEW = [0,1,2,3,4,5]
home = str(Path.home())
DIRECTORY_IMPORT_FILES = home + os.sep +DESKTOP + os.sep +'VDE-Sicherheitsprüfung'+ os.sep +'Datenbank'+ os.sep +'EPRUEFUNG Backups'
# Help Information
VERIFICATION_PDF = os.getcwd() + os.sep + 'help_info'+ os.sep +'Anleitung zum Prüfung.pdf'
GROSSENMETRAWATT_PDF = os.getcwd() + os.sep + 'help_info'+ os.sep +'Anleitung zur Gossen Metrwatt Datenbank.pdf'
GENERAL_PDF = os.getcwd() + os.sep + 'help_info'+ os.sep +'Anleitung zur EPruefung Software.pdf'
GROSSENMETRAWATT_IMPORT_EXPORT_PDF = os.getcwd() + os.sep + 'help_info'+ os.sep +'Anleitung zum Gossen Metrawatt Importieren_Exportieren.pdf'
GROSSENMETRAWATT_CONFIGURATE_EMAIL_PDF = os.getcwd() + os.sep + 'help_info'+ os.sep +'Anleitung zur automatischen E-Mail.pdf'
# Default e-mail to rooms' resposibles
MESSAGE_EMAIL_ROOM_RESPONSABLES =  home + os.sep + DESKTOP + os.sep +'VDE-Sicherheitsprüfung'+ os.sep + 'Responsable_email-Setup'+ os.sep +'message.txt'
EMAIL_SETUP_ROOM_RESPONSABLES = home + os.sep + DESKTOP + os.sep +'VDE-Sicherheitsprüfung'+ os.sep +'Responsable_email-Setup'+ os.sep +'email-setup.txt'
EMAIL_CONTACTS_ROOM_RESPONSABLES = home + os.sep + DESKTOP + os.sep +'VDE-Sicherheitsprüfung'+ os.sep +'Responsable_email-Setup'+ os.sep  + 'contact.txt'
BACKUPFILE_MESSAGE_EMAIL_ROOM_RESPONSABLES = os.getcwd() + os.sep + 'help_txts' + os.sep + 'message.txt'
BACKUPFILE_EMAIL_CONTACTS_ROOM_RESPONSABLES = os.getcwd() + os.sep + 'help_txts' + os.sep + 'contact.txt'
# Default e-mail Backup
MESSAGE_EMAIL_ENGLISH = home + os.sep + DESKTOP + os.sep +'VDE-Sicherheitsprüfung'+ os.sep +'email-Setup'+ os.sep +'message_english.txt'
BACKUPFILE_MESSAGE_EMAIL_ENGLISH = os.getcwd() + os.sep + 'help_txts' + os.sep + 'message_english.txt'
MESSAGE_EMAIL_GERMAN = home + os.sep + DESKTOP + os.sep +'VDE-Sicherheitsprüfung'+ os.sep +'email-Setup'+ os.sep +'message_german.txt'
BACKUPFILE_MESSAGE_EMAIL_GERMAN = os.getcwd() + os.sep + 'help_txts' + os.sep + 'message_german.txt'
CONTACTS_EMAIL = home + os.sep + DESKTOP + os.sep +'VDE-Sicherheitsprüfung'+ os.sep +'email-Setup'+ os.sep +'contact.txt'
EMAIL_SETUP = home + os.sep + DESKTOP + os.sep +'VDE-Sicherheitsprüfung'+ os.sep +'email-Setup'+ os.sep +'email-setup.txt'
# Grossen Metrawatt CONFIGURATION
GROSSENMETRAWATT_DIRECTORY = home + os.sep + DESKTOP + os.sep +'VDE-Sicherheitsprüfung'+ os.sep +'Datenbank'+ os.sep +'IZYTRON.IQ Export'
AUXILIAR_TXT_FILE_READ = 'IZYTRON_exportFileAuxiliarRead.txt' # File that is created and eliminated
AUXILIAR_TXT_FILE_WRITE = 'IZYTRON_exportFileAuxiliarWrite.txt' # File that is created and then transformed
# Grossen Metrawatt XML CONFIG
READ_BUILDING_NAME = 'Buildings'
READ_ITEMS_NAME = 'Items'
READ_FLOOR_NAME = 'Rooms'
READ_ID_NAME_LOCATIONS = 'Id'
READ_ID_NAME = 'Identification'
READ_MEASURMENT_NAME = 'Measurements'
READ_NAME_NAME= 'Name'
READ_NOTE_NAME = 'Remark'
READ_LOCATION_NAME = 'LocationId'
READ_COMPANY_NAME = 'Manufacturer'
READ_SERIALNUMBER = 'SerialNumber'
READ_ROOMTYPE_NUMBER = 'TestingInterval'
READ_LAST_TESTING_DATE = 'LastTestingDate'
READ_NEXT_TESTING_DATE = 'NextTestingDate'
# Laguange Setting
DEFAULT_LANGUAGE_ENGLISH = 1
DEFAULT_LANGUAGE_GERMAN = 2
DEFAULT_LAGUAGE_SETTING = os.getcwd() + os.sep + 'Language_Package'+ os.sep +'default_language.txt'
# Finding responsibles in website
DEFAULT_FILE_WITH_WEBSITE = os.getcwd() + os.sep + 'help_info'+ os.sep +'linkToWebSite.txt'
# Reading the Link to find the contact of certain responsible of a room
HTML_MAIN = 'main'
HTML_SECTION = 'section'
HTML_TR = 'tr'
HTML_TD = 'td'
HTML_DIV = 'div'
HTML_TBODY = 'tbody'
HTML_TABLE = 'table'
FINAL_EMAIL = 'darmstadt.de'
NOTFOUND = -404
EXCEPTION_NAMES=['Institutsleitung', 'Sekretariat', 'Administrativer Support', 'IT Support', 'Mechanischer Support']
EXCEPTION_NAMES_2 = ['Emeritus', 'Oberingenieure','Lehrbeauftragte', 'Finanzen | Controlling', 'Wissenschaftliche Mitarbeiter' ]

TABLE_VIEW_STYLE= """ 
        QTableWidget {
            background-color: rgb(255,255,255);
            alternate-background-color: rgb(135,206,250);
            selection-background-color: rgb(255,255,255);
            selection-color: rgb(0,0,205);
        }
        QHeaderView::section {
            background-color: rgb(205,205,205);
            color: rgb(25,25,25);
            font-size: 16px;
            margin: 4px;
        }
        QHeaderView::section::selected {
            background-color: rgb(205,205,205);
            color: rgb(25,25,25);
            font-size: 18px;
            margin: 4px;
        }
        QTableWidget::item::selected{
            background-color: rgb(135,206,250);
            color: rgb(25,25,25);
        }        
        QTableWidgetItem {
            background-color: rgb(255,255,255);
            color: rgb(25,25,25);
        }
    """
# Menu bar
MENU_BAR_STYLE= """ 
        QMenuBar {
            background-color: rgb(148,28,38);
            color: rgb(255,255,255);
        }
        QMenuBar::item {
            background-color: rgb(148,28,38);
            color: rgb(255,255,255);
            padding: -10px 15px;
        }
        QMenuBar::item::selected {
            background-color: rgb(100,28,38);
        }
        QMenu {
            background-color: rgb(148,28,38);
            color: rgb(255,255,255);         
        }
        QMenu::item::selected {
            background-color: rgb(100,28,38);
        }
    """