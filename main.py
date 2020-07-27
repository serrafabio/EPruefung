"""
This software was developed to help the verification of the devices of the institute PTW of Technische Universität Darmstadt.
The Software has the goal of organizing the devices of the institute. A secondary Data base of Gossen Metrawatt has the same goal,
but due to it difficult, this software came to help the user of this organization and increase his/her production.

The Software is used only by the PTW institute and doesn't have commercial purpose.
This software conts with 4(four) scripts:
    main.py (graphic interface)
    DataBase.py (Data Base of manipulation)
    XMLDataBase (integration with Grossen Metrawatt Data Base)
    constants.py (Arquive with all constants: for any manipulation, it is indicated to manipulate this script)
For further questions about the software, please contact the developer:

@ Fabio Serra Pereira
contact: serrafabio10@outlook.com (write in english or german)
Problem: Inside the e-mail hiwiEPruefung is a back-UP of the software
Version: 1.1
"""

import sys
import os
import re
import webbrowser as wd
from os.path import basename
from pathlib import Path, PureWindowsPath
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.Qt import *
from PyQt5 import QtCore
from constants import *
from DataBase import DB_Manipulation
from XMLDataBase import GMW
from datetime import datetime
from datetime import date
from string import Template
import SearchResposibles as sr
import time
import psutil
import win32com.client as win32


# Create a custom "QProxyStyle" to enlarge the QMenu icons
class MyProxyStyle(QProxyStyle):
    pass
    def pixelMetric(self, QStyle_PixelMetric, option=None, widget=None):

        if QStyle_PixelMetric == QStyle.PM_SmallIconSize:
            return 65
        else:
            return QProxyStyle.pixelMetric(self, QStyle_PixelMetric, option, widget)

class MainWindow(QMainWindow):
    def __init__(self):
        """
        Here we just call the main layout to load
        """
        super().__init__()
        # font configuration
        self.font = QFont(FONT)
        self.main_window()

    def __init__variables(self):
        """
        Here we start all the variables, which musst be inizialized.
        :return: None
        """
        self.name_building_selected, self.name_floor_selected, self.name_room_selected, self.type_room = '', '', '', OFFICE
        self.canInsertID, self.saved, self.undo_database, self.TableWidgetEditing, self.undoLimit, self.ContinuesAction = False, True, 0, False, 0, True


    def main_window(self):
        """
        Here we load the widgets used in the main window
        :return: (self).show()
        """
        self.__init__variables()# Inicialize Variables
        # Language Setting
        with open(DEFAULT_LAGUAGE_SETTING, 'r', encoding='utf-8') as FileLaguangeSetting:  # open the file
            for line in FileLaguangeSetting: self.currentLanguageSetting = int(line)  # get the data
        FileLaguangeSetting.close()
        if self.currentLanguageSetting == DEFAULT_LANGUAGE_ENGLISH: from Language_Package import english as laguage  # English
        else: from Language_Package import german as laguage  # German
        self.laguage = laguage
        self.setGeometry(WIDTH,HEIGHT,MAIN_WINDOW_WIDTH,MAIN_WINDOW_HEIGHT), self.setWindowTitle(NAME_OF_PROGRAM) # Set the first size of the main window
        self.setWindowIcon(QIcon(ICON_PTW_LOGO)) # Set the icon of ptw as logo of software
        self.db = DB_Manipulation(); self.building_available = sorted(self.db.load_building())# Inicialize the Data Bank

        self.font.setPointSize(11) # Font Size to Menu bar
        ##### Menu bar #######
        menu_bar = QMenuBar(); menu_bar.setFont(self.font), menu_bar.setAutoFillBackground(True), menu_bar.setStyleSheet(MENU_BAR_STYLE) # set style Menu Bar
        icon_ptw_menu_bar = menu_bar.addMenu(QIcon(ICON_PTW_LOGO), '&'); icon_ptw_menu_bar.setEnabled(False); # set the icon in the menu bar
        file_menu_bar = menu_bar.addMenu(self.laguage.menuBar_file) # insert File Option
        self.building_menu_bar = menu_bar.addMenu(self.laguage.menuBar_building) # Insert the building option
        edit_menu_bar = menu_bar.addMenu(self.laguage.edit_menuBar) # Insert Edit option
        gmw_menu_bar = menu_bar.addMenu(self.laguage.menuBar_grossenMetrawatt) # Insert Gossen Metrawatt option
        help_menu_bar = menu_bar.addMenu(self.laguage.menuBar_help) # Insert the Help option
        language_menu_bar = menu_bar.addMenu(self.laguage.menuBar_language)  # Insert the Language option
        #undo_arrow_menu = menu_bar.add # undo option
        # File Option
        save_action_menuBar = QAction(self.laguage.saveAll, self); save_action_menuBar.setShortcut('Ctrl+S'), file_menu_bar.addAction(save_action_menuBar), file_menu_bar.addSeparator()  # Save
        import_DB_menuBar = QAction(self.laguage.importBU, self); import_DB_menuBar.setShortcut('Ctrl+I'), file_menu_bar.addAction(import_DB_menuBar) # Import Data Base, if wanted
        backup_action_menuBar = QAction(self.laguage.exportBU, self); backup_action_menuBar.setShortcut('Ctrl+B'),file_menu_bar.addAction(backup_action_menuBar), file_menu_bar.addSeparator()  # Create Back-UP file
        email_automatic = QAction(self.laguage.send_email_to_responsible, self); file_menu_bar.addAction(email_automatic), file_menu_bar.addSeparator()
        exit_action_menuBar = QAction(self.laguage.exit_label, self); exit_action_menuBar.setShortcut('Ctrl+Q'), exit_action_menuBar.triggered.connect(self.close), file_menu_bar.addAction(exit_action_menuBar) # Exit system
        # Edit
        # Undo Arrow
        undo_arrow_menuBar = QAction(QIcon(ICON_UNDO), self.laguage.undo, self); undo_arrow_menuBar.setShortcut('Ctrl+Z'), edit_menu_bar.addAction(undo_arrow_menuBar); edit_menu_bar.addSeparator() # we undo actions
        seeRoomsnotVerified = QAction(self.laguage.verify_room, self); seeRoomsnotVerified.setShortcut('Alt+S'), edit_menu_bar.addAction(seeRoomsnotVerified) # Here we see the rooms with devices that need to verify
        actualizeLinkRoomsNotVerified = QAction(self.laguage.actualizeLink, self); edit_menu_bar.addAction(actualizeLinkRoomsNotVerified), edit_menu_bar.addSeparator() # Here we can actualize the link, where we find the Institute's employees
        insertItem = QAction(self.laguage.insertItem, self); insertItem.setShortcut('Ctrl+Y'), edit_menu_bar.addAction(insertItem) # Insert Item into Table View
        deleteItem = QAction(self.laguage.deleteItem, self); deleteItem.setShortcut('Ctrl+W'), edit_menu_bar.addAction(deleteItem), edit_menu_bar.addSeparator()  # Delete Item from Table View
        addLocation_menuBar = QAction(self.laguage.addLocation, self); addLocation_menuBar.setShortcut('Ctrl+L'), edit_menu_bar.addAction(addLocation_menuBar) # Habilite to add or eliminate Buildings, floor and Rooms
        # Gossen Metrawatt
        import_GMW = QAction(self.laguage.importGW, self);import_GMW.setShortcut('Ctrl+Shift+G'), gmw_menu_bar.addAction(import_GMW), gmw_menu_bar.addSeparator()  # export data to  Metrawatt
        go_to_save_realDB = QAction(self.laguage.exportGW, self); go_to_save_realDB.setShortcut('Ctrl+G'), gmw_menu_bar.addAction(go_to_save_realDB) # export data to Gossen Metrawatt
        # Help
        open_help_measument_menuBar = QAction(self.laguage.help_measure, self); help_menu_bar.addAction(open_help_measument_menuBar), open_help_measument_menuBar.triggered.connect(lambda ch, namePDF= VERIFICATION_PDF: self.helpPDF(VERIFICATION_PDF)) # Tutorial how to Measure
        open_help_metrawatt_selfediting_menuBar = QAction(self.laguage.GW_tutorial_help, self); help_menu_bar.addAction(open_help_metrawatt_selfediting_menuBar), open_help_metrawatt_selfediting_menuBar.triggered.connect(lambda ch, namePDF= GROSSENMETRAWATT_PDF: self.helpPDF(GROSSENMETRAWATT_PDF)) # Tutorial how to use Gossen Metrawatt
        open_help_add_data_menuBar = QAction(self.laguage.help_tutorial_2, self); help_menu_bar.addAction(open_help_add_data_menuBar) # Tutorial how to use this software
        open_help_export_data_menuBar = QAction(self.laguage.help_tutorial_3, self); help_menu_bar.addAction(open_help_export_data_menuBar) # Tutorial how to export data
        open_help_config_email = QAction(self.laguage.help_tutorial_4, self); help_menu_bar.addAction(open_help_config_email) # Tutorial to configure the automatic e-mail
        # Languages:
        englis_menuBar = QAction(self.laguage.english, self); language_menu_bar.addAction(englis_menuBar); englis_menuBar.triggered.connect(lambda ch, index = DEFAULT_LANGUAGE_ENGLISH: self.LaguangeSetting(index))
        german_menuBar = QAction(self.laguage.german, self); language_menu_bar.addAction(german_menuBar); german_menuBar.triggered.connect(lambda ch, index=DEFAULT_LANGUAGE_GERMAN: self.LaguangeSetting(index))
        self.setMenuBar(menu_bar)

        self.font.setPointSize(SIZE_FONT) # Font Size to Main Window
        ###### Central Body ######
        main_grid_layout = QGridLayout() # This is the main grid, where we will have 2 parts
                                            # the table view, where we will edit the ids and others parameters
                                            # then we have the part of select floor, room and type
        self.table_view_main_window = QTableWidget() # Create the Table view
        self.table_view_main_window.setGeometry(0,0,MINIMU_WIDTH_TABLE_VIEW_MAIN_WINDOW,MINIMU_HEIGHT_TABLE_VIEW_MAIN_WINDOW), self.table_view_main_window.setMinimumWidth(MAIN_WINDOW_WIDTH), self.table_view_main_window.setMinimumHeight(MINIMU_HEIGHT_TABLE_VIEW_MAIN_WINDOW) # geometry of the table view
        self.table_view_main_window.setColumnCount(6), self.table_view_main_window.setShowGrid(True), self.table_view_main_window.setFont(self.font), self.table_view_main_window.setStyleSheet(TABLE_VIEW_STYLE) # Define as fixed the number of columns
        self.table_view_main_window.setHorizontalHeaderLabels([self.laguage.id_body, self.laguage.devices_Body, self.laguage.manufacturer_body, self.laguage.serialNumber_body, self.laguage.note_body, self.laguage.current_Date_Body]),self.table_view_main_window.horizontalHeader().setFont(self.font), self.table_view_main_window.itemDelegateForColumn(5) # Add the Columns
        for i in range(1,5): self.table_view_main_window.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch) # adjust the size of the Table view to auto-adjust
        self.table_view_main_window.itemPressed.connect(self.TableWidget_clicked); self.table_view_main_window.setSelectionBehavior(QTableView.SelectRows)
        main_grid_layout.addWidget(self.table_view_main_window) # add the model to the layout
        # buttom of the main window
        secondary_grid_layout = QGridLayout() # This are the button in the bottom, where we have the selection of the room, type and the buttons
        secondary_1_box_layout = QVBoxLayout(); secondary_grid_layout.addLayout(secondary_1_box_layout,0,0), secondary_1_box_layout.setContentsMargins(10,5,0,5) # Receives the dectetion of the room
        secondary_2_box_layout = QVBoxLayout(); secondary_grid_layout.addLayout(secondary_2_box_layout,0,2), secondary_2_box_layout.setContentsMargins(25,5,0,55)# Receives the type of the room
        secondary_3_box_layout = QVBoxLayout(); secondary_grid_layout.addLayout(secondary_3_box_layout,0,3), secondary_3_box_layout.setContentsMargins(25,5,5,5) # Receives the bottons
        #Layout 1
        self.label_floor = QLabel(); self.label_floor.setText(self.laguage.floor_body), self.label_floor.setFont(self.font), self.label_floor.setFixedWidth(MINIMUM_WIDTH_MAIN_BUTTONS), secondary_1_box_layout.addWidget(self.label_floor) #Label Floor
        self.check_box_floor = QComboBox(); self.check_box_floor.setFont(self.font), self.check_box_floor.setFixedWidth(MINIMUM_WIDTH_MAIN_BUTTONS), self.check_box_floor.activated.connect(lambda ch, name= self.check_box_floor.currentText() : self.Floor_selected(self.check_box_floor.currentText())), secondary_1_box_layout.addWidget(self.check_box_floor) #Select the floor
        room_label = QLabel();room_label.setText(self.laguage.Room_body), room_label.setFont(self.font), room_label.setFixedWidth(MINIMUM_WIDTH_MAIN_BUTTONS), secondary_1_box_layout.addWidget(room_label)  # Label Room
        self.check_box_room = QComboBox();self.check_box_room.setFont(self.font), self.check_box_room.setFixedWidth(MINIMUM_WIDTH_MAIN_BUTTONS), self.check_box_room.activated.connect(lambda ch, name= self.check_box_room.currentText() : self.Load_devices(self.check_box_room.currentText())), secondary_1_box_layout.addWidget(self.check_box_room)  # Select the floor
        #Layout 2
        label_type = QLabel(); label_type.setFont(self.font), label_type.setText(self.laguage.typeRoom_body),label_type.setFixedWidth(MINIMUM_WIDTH_TYPE_SELECTION), secondary_2_box_layout.addWidget(label_type) #Label of the type of room
        radiobutton_office_type = QRadioButton(checkable=True); radiobutton_office_type.setChecked(True), radiobutton_office_type.setFont(self.font), radiobutton_office_type.setText(self.laguage.office_body), secondary_2_box_layout.addWidget(radiobutton_office_type), radiobutton_office_type.clicked.connect(lambda ch, name = radiobutton_office_type.text(): self.Radio_Button_Type(radiobutton_office_type.text()))  # Radio button of selection type of room
        radiobutton_labor_type = QRadioButton(checkable=True); radiobutton_labor_type.setChecked(False), radiobutton_labor_type.setFont(self.font),  radiobutton_labor_type.setText(self.laguage.labor_body), secondary_2_box_layout.addWidget(radiobutton_labor_type), radiobutton_labor_type.clicked.connect(lambda ch, name = radiobutton_labor_type.text(): self.Radio_Button_Type(radiobutton_labor_type.text())) # Radio button of selection type of room
        #Layout 3
        # Buttons
        insert_button = QPushButton(self.laguage.insertDevice); insert_button.setFont(self.font), insert_button.setMinimumWidth(MINIMUM_WIDTH_MAIN_BUTTONS), insert_button.clicked.connect(self.insertNewDeviceintoRoom), secondary_3_box_layout.addWidget(insert_button) # Here we create the button to insert new devices
        delete_button = QPushButton(self.laguage.deleteDevice); delete_button.setFont(self.font), delete_button.clicked.connect(self.deleteRowAndItem), secondary_3_box_layout.addWidget(delete_button) # Here we delete the selected row (device)
        save_button = QPushButton(self.laguage.saveRoom); save_button.setFont(self.font), save_button.clicked.connect(self.saveRoom),secondary_3_box_layout.addWidget(save_button) # Here we save the new and atualized ids
        # Building (load buildings available )
        self.Building_load_menu(); insertItem.triggered.connect(self.insertNewDeviceintoRoom); deleteItem.triggered.connect(self.deleteRowAndItem), go_to_save_realDB.triggered.connect(lambda ch, action = 2: self.ImportExportFile(2))
        save_action_menuBar.triggered.connect(self.SaveAll), addLocation_menuBar.triggered.connect(self.CreateNewLocation), seeRoomsnotVerified.triggered.connect(self.SeeRoomsNotVerified), actualizeLinkRoomsNotVerified.triggered.connect(self.ActualizeLinkOfEmployeesList)
        import_DB_menuBar.triggered.connect(lambda ch, action = 0: self.ImportExportFile(0)), backup_action_menuBar.triggered.connect(lambda ch, action = 1: self.ImportExportFile(1)), import_GMW.triggered.connect(lambda ch, action = 3: self.ImportExportFile(3)), email_automatic.triggered.connect(self.SendEmailForResponsible) # Addtionals Actions
        open_help_add_data_menuBar.triggered.connect(lambda ch, name = GENERAL_PDF: self.helpPDF(GENERAL_PDF)); open_help_export_data_menuBar.triggered.connect(lambda ch, name = GROSSENMETRAWATT_IMPORT_EXPORT_PDF: self.helpPDF(GROSSENMETRAWATT_IMPORT_EXPORT_PDF)); open_help_config_email.triggered.connect(lambda ch, name = GROSSENMETRAWATT_CONFIGURATE_EMAIL_PDF: self.helpPDF(GROSSENMETRAWATT_CONFIGURATE_EMAIL_PDF));
        undo_arrow_menuBar.triggered.connect(self.restoreUndoAction)

        main_grid_layout.addLayout(secondary_grid_layout,1,0)
        widget = QWidget(self); widget.setLayout(main_grid_layout); self.setCentralWidget(widget)

        # Dialog of Problem
        self.problemDialog = QDialog(); self.problemDialog.setFont(self.font), self.problemDialog.move(DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER), self.problemDialog.setFixedSize(DIALOG_PROBLEM_WIDHT,DIALOG_PROBLEM_HEIGHT) # create Dialog
        vbox = QVBoxLayout(); self.problemDialog.setWindowTitle(self.laguage.problemDetected), self.problemDialog.setWindowIcon(QIcon(ICON_PTW_LOGO))
        qlabel = QLabel(self.laguage.problemDetected_2); qlabel.setFont(self.font), qlabel.setAlignment(Qt.AlignCenter), qlabel.setFixedWidth(DIALOG_PROBLEM_WIDHT-10), vbox.addWidget(qlabel)
        self.labelProblem = QLabel(); self.labelProblem.setFont(self.font), self.labelProblem.setFixedWidth(DIALOG_PROBLEM_WIDHT-10), self.labelProblem.setAlignment(Qt.AlignCenter), vbox.addWidget(self.labelProblem)
        self.closeBtt = QPushButton(self.laguage.closeBtt, self.problemDialog); self.closeBtt.setFont(self.font), self.closeBtt.clicked.connect(self.problemDialog.close), self.closeBtt.setFixedWidth(DIALOG_BUTTONS_WIDTH), self.closeBtt.setPalette(QPalette(Qt.red)), vbox.addWidget(self.closeBtt), vbox.setAlignment(self.closeBtt, Qt.AlignCenter)
        vbox.setAlignment(Qt.AlignCenter),self.problemDialog.setLayout(vbox)

        self.update() ,self.show()

    def restoreUndoAction(self):
        """
        Here we activated always the action of undo the database
        :return: Data Base actualized
        """
        if self.undo_database > self.undoLimit:
            self.undo_database -= 1; self.db.RecoverUndoAction(self.undo_database) # here we restore
            self.building_available = sorted(self.db.load_building())
            self.Building_load_menu()
            if (self.name_room_selected != ''): self.Load_devices(self.name_room_selected)

    def keyReleaseEvent(self, event):
        if(event.key() == Qt.Key_Tab and self.TableWidgetEditing):
            self.SaveInformationOfTableWidget()

    def LaguangeSetting(self, laguangeReceveid):
        """"
        The goal is to change the language of the user
        :return None
        """
        if(laguangeReceveid != self.currentLanguageSetting):
            fileToWrite = open(DEFAULT_LAGUAGE_SETTING, 'w', encoding='utf-8')
            if(laguangeReceveid == DEFAULT_LANGUAGE_ENGLISH): fileToWrite.write('%s' %DEFAULT_LANGUAGE_ENGLISH) # change to english
            else: fileToWrite.write('%s' %DEFAULT_LANGUAGE_GERMAN) # change to german
            fileToWrite.close()
            self.main_window()

    def TableWidget_clicked(self, item):
        """
        The goal is to limited the column of selection
        :return: None
        """
        if (self.table_view_main_window.currentColumn() == 0): self.table_view_main_window.setEditTriggers(QTableWidget.NoEditTriggers)
        else:
            self.table_view_main_window.editItem(item)
            self.SaveInformationOfTableWidget()
        self.table_view_main_window.update()

    def ActualizeLinkOfEmployeesList(self):
        """
        here we actualize the txt, where we save the link
        :return: None
        """
        self.newDialog_URLActualization = QDialog(self); self.newDialog_URLActualization.setFont(self.font); self.newDialog_URLActualization.move(DIALOG_WIDTH_CALENDER, DIALOG_HEIGHT_CALENDER), self.newDialog_URLActualization.setFixedSize(DIALOG_WIDTH_URLACTUALIZE, DIALOG_HEIGHT_ID)  # Create dialog
        vbox, hbox, hbox_1 = QVBoxLayout(), QHBoxLayout(), QHBoxLayout(); self.newDialog_URLActualization.setWindowTitle(self.laguage.actualizeLink_title), self.newDialog_URLActualization.setWindowIcon(QIcon(ICON_PTW_LOGO))  # layout
        qlabel = QLabel(); qlabel.setFont(self.font), qlabel.setText(self.laguage.pleaseInsertURLhere), qlabel.setFixedWidth(DIALOG_WIDTH_CALENDER + 10), hbox_1.addWidget(qlabel), hbox_1.setAlignment(Qt.AlignLeft)  # Set first label
        self.qlabel_saved = QLabel(); self.qlabel_saved.setFont(self.font), self.qlabel_saved.setText(self.laguage.saved), self.qlabel_saved.setFixedWidth( DIALOG_WIDTH_CALENDER + 10), self.qlabel_saved.hide(), self.qlabel_saved.setStyleSheet("QLabel { color : blue; }") ,hbox_1.addWidget(self.qlabel_saved), vbox.addLayout(hbox_1)  # Set first label
        with open(DEFAULT_FILE_WITH_WEBSITE, 'r', encoding='utf-8') as urlfile:
            for line in urlfile: url = str(line)
        qbox = QLineEdit(); qbox.setFont(self.font), qbox.setPlaceholderText(url), qbox.setFixedHeight(LINE_EDITOR_HEIGHT), qbox.setFixedWidth(DIALOG_WIDTH_URLACTUALIZE - 20), vbox.addWidget(qbox)  # Place to insert the url
        savebtt = QPushButton(self.laguage.save); savebtt.setFont(self.font), savebtt.setFixedWidth(DIALOG_WIDTH_CALENDER // 2), savebtt.setPalette(QPalette(Qt.blue)), savebtt.clicked.connect(lambda ch, text = qbox.text(): self.ActualizeLinkOfEmployeesListSaved(qbox.text())), hbox.addWidget(savebtt) # Save button
        closeBtt = QPushButton(self.laguage.cancelBtt);closeBtt.setFont(self.font), closeBtt.setFixedWidth(DIALOG_WIDTH_CALENDER // 2), closeBtt.setPalette(QPalette(Qt.red)), closeBtt.clicked.connect(self.newDialog_URLActualization.close), hbox.addWidget(closeBtt), vbox.addLayout(hbox)  # Close Button
        vbox.setContentsMargins(10, 10, 10, 10), self.newDialog_URLActualization.setLayout(vbox), self.newDialog_URLActualization.open()

    def ActualizeLinkOfEmployeesListSaved(self, text):
        """
        Here we save the link in the corresponding file
        :return: TXT saved and actualized
        """
        if not text == '':
            with open(DEFAULT_FILE_WITH_WEBSITE, 'w', encoding='utf-8') as urlfile: urlfile.write(text)
            self.qlabel_saved.show(), QCoreApplication.processEvents()
            time.sleep(3)
            self.newDialog_URLActualization.close()

    def Building_load_menu(self):
        """
        We load the Menu Bar in the section of the Building selection
        :return: (self).building_menu_bar (QMenu): Contends the building available
        """
        self.TableWidgetEditing = False
        self.building_menu_bar.clear()
        for buildings in self.building_available: building_action_menuBar = QAction(buildings, self, checkable= True);building_action_menuBar.setChecked(False) ;self.building_menu_bar.addAction(building_action_menuBar) #Load the Menu Bar section of the building
        self.building_menu_bar.update(); self.canInsertID = False; self.building_menu_bar.triggered.connect(self.actionClicked); self.label_floor.setText(self.laguage.floor_body)

    def Building_selected(self, name):
        """
        We receive the name of the building selected and then
        :param name: is the name of the building selected
        :return: (self).check_box_floor (QCheckBox): Contends the Rooms available
        """
        self.check_box_room.clear(); self.check_box_floor.clear(); self.table_view_main_window.setRowCount(0); self.devices = None; self.name_room_selected = ''; self.name_floor_selected = ''
        floor_available = sorted(self.db.load_floor(name))  # Load the floors available
        self.check_box_floor.clear(), self.check_box_floor.addItem(''), self.check_box_floor.addItems(floor_available); # Insert into the Combo Box
        self.name_building_selected = name; self.canInsertID = False; self.label_floor.setText(self.laguage.floor_2_body +str(self.name_building_selected)) # Atualize name of the building

    @QtCore.pyqtSlot(QAction)
    def actionClicked(self, action):
        for acc in self.building_menu_bar.actions(): acc.setChecked(False)
        action.setChecked(True)
        self.Building_selected(action.text())

    def Floor_selected(self, name):
        """
        Here we call the function that we load the devices into the Table View
        :param name: name of the floor selected
        :return: (self).check_box_room (QCheckBox)
        """
        self.table_view_main_window.setRowCount(0)
        if name != '':
            rooms_available = sorted(self.db.load_rooms(self.name_building_selected, name)) # Load the rooms available
            self.check_box_room.clear(); self.check_box_room.addItem(''), self.check_box_room.addItems(rooms_available) # Insert into the Combo Box
        else: self.name_room_selected = name; self.check_box_room.clear(); self.devices = None
        self.name_floor_selected = name; self.name_room_selected = ''
        self.canInsertID = False  # Atualize name of the floor

    def Load_devices(self, name):
        """
        Here we atualize the table Widget with the IDs available for a specific room
        :param name: Name of the room selected
        :return: (self).table_view_main_window (QTableWidget): table that contends the IDs available in the corresponding room
        """
        self.saved, self.TableWidgetEditing = False, True
        if name != '':
            self.name_room_selected = name; self.canInsertID = True
            self.devices = self.db.load_devices(self.name_building_selected, self.name_floor_selected, name); self.device_neg = False # Load all devices available in a specific room
            for device in self.devices:
                if int(device[0]) <= 0 : self.device_neg = device; self.devices.remove(device)
            self.table_view_main_window.setRowCount(len(self.devices)) # We load the number of row required
            for row in range(len(self.devices)): # Atualizing the table widget
                for column in VARIABLES_WANTED_IN_TABLE_VIEW:
                    if(column !=5): self.table_view_main_window.setItem(row, column, QTableWidgetItem(str(self.devices[row][column]))) # Load the information into the table widget
                    else:
                        deviceDate = re.findall(r'\d+',str(self.devices[row][column]) )
                        qDate = QDate(int(deviceDate[0]), int(deviceDate[1]), int(deviceDate[2]));
                        if(self.currentLanguageSetting == DEFAULT_LANGUAGE_ENGLISH): date_edit = QPushButton('{0}/{1}/{2}'.format(qDate.month(), qDate.day(), qDate.year()))
                        else: date_edit = QPushButton('{0}/{1}/{2}'.format(qDate.day(), qDate.month(), qDate.year()))
                        date_edit.clicked.connect(self.Calender_for_Table_Widget)
                        self.table_view_main_window.setCellWidget(row, column, date_edit) # Load button to the date selection
        else: self.table_view_main_window.setRowCount(0); self.devices = None;  self.name_room_selected = name

    def Calender_for_Table_Widget(self):
        """
        This function load dialog possible to select the date of edition
        :param  currentDay (string): current date of the measurement of specific ID
        :return: dialog_calender(QDialog): load a dialog conteds the Calender
        """
        self.undoLimit = self.db.SaveUndoAction(self.undo_database); self.undo_database += 1  # save undo action
        row_of_buttton= self.table_view_main_window.currentRow() # We get the row selected by the button
        currentYear = datetime.now().year; currentMonth = datetime.now().month; currentDay = datetime.now().day # load today
        dialog_calender = QDialog(self); dialog_calender.setGeometry(DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER, DIALOG_HEIGHT_CALENDER) #Dialog of the calender
        vbox = QVBoxLayout(); dialog_calender.setWindowIcon(QIcon(ICON_PTW_LOGO)), dialog_calender.setWindowTitle(self.laguage.calenderName) #Layout of the dialog
        calendar = QCalendarWidget()
        if(self.currentLanguageSetting == DEFAULT_LANGUAGE_ENGLISH): calendar.setLocale(QLocale(QLocale.English))# Calender & Default Language: english
        else: calendar.setLocale(QLocale(QLocale.German))# Calender & Language: german
        calendar.setMinimumDate(QDate(currentYear-5 , currentMonth, currentDay)), calendar.setMaximumDate(QDate(currentYear+5, currentMonth, currentDay)), calendar.setSelectedDate(QDate(currentYear, currentMonth, currentDay)) # Configure the calender
        self.date_to_change = QDate(currentYear, currentMonth, currentDay)
        calendar.clicked.connect(self.get_date_calender)
        vbox.addWidget(calendar) # Add the Calender
        save_only = QPushButton(); save_only.setText(self.laguage.calenderSaveOnly); save_only.clicked.connect(lambda ch, row = row_of_buttton :self.save_only_calender(row_of_buttton)); save_only.clicked.connect(dialog_calender.close); vbox.addWidget(save_only) # Button to save just for the measured selected
        save_all= QPushButton(); save_all.setText(self.laguage.calenderSaveAll); save_all.clicked.connect(self.save_all_calender); save_all.clicked.connect(dialog_calender.close) ;vbox.addWidget(save_all) # Button to save for all measurments
        close_dialog = QPushButton(); close_dialog.setText(self.laguage.cancelBtt); close_dialog.clicked.connect(dialog_calender.close), vbox.addWidget(close_dialog) # Button to close dialog
        dialog_calender.setLayout(vbox) # set layout
        dialog_calender.show() # show the dialog

    def get_date_calender(self, date):
        """
        Get the date selected in the calender
        :param date: contends the date selected
        :return: (self).date(QDate) : contends the new date selected
        """
        self.date_to_change = date

    def SaveInformationOfTableWidget(self):
        """
        We save the information of the table widget into the data Base
        :return: DataBase atualized
        """
        self.undoLimit = self.db.SaveUndoAction(self.undo_database); self.undo_database += 1  # save undo action
        for row in range(self.table_view_main_window.rowCount()):
            id_name = int(self.devices[row][0])  # get the id
            try:name_object = self.table_view_main_window.item(row,1);name_object = name_object.text()  # get the name of device
            except:name_object = ''  # Bug fixer
            try:manufacturer = self.table_view_main_window.item(row,2); manufacturer = manufacturer.text()  # get the name of company
            except:manufacturer = ''  # bug fixer
            try:inv_num = self.table_view_main_window.item(row,3); inv_num = inv_num.text()  # get the Serial Number of the device
            except:inv_num = ''  # bug fixer
            try:note_aux = self.table_view_main_window.item(row,4); note_aux = note_aux.text()  # get the Note of the device
            except:note_aux = ''  # bug fixer
            self.db.insertInformationIntoCurrent(id_name, name_object, manufacturer, inv_num, note_aux)  # We update the information about the device

    def save_only_calender(self, row):
        """
        This function only save the date for the selected index
        :param: row (int): contends the row of the selcted item
        :return: save the date to item selected
        """
        self.SaveInformationOfTableWidget()
        current_date_for_db = str(self.date_to_change.year()) + '-' + str(self.date_to_change.month()) + '-' + str(self.date_to_change.day()); next_date_for_db = str(self.date_to_change.year()+(self.type_room//12)) + '-' + str(self.date_to_change.month()) + '-' + str(self.date_to_change.day()) # we update the data of modification
        self.db.insert_date_into_data_base(self.devices[row][0], current_date_for_db, self.type_room, next_date_for_db) # insert into the data bank and atualize
        self.Load_devices(self.name_room_selected) # reload the Table Widget

    def save_all_calender(self):
        """
        This function save the date for all selected items
        :return: save the date to all items
        """
        self.SaveInformationOfTableWidget()
        for row in range(len(self.devices)):
            current_date_for_db = str(self.date_to_change.year()) + '-' + str(self.date_to_change.month()) + '-' + str(self.date_to_change.day()); next_date_for_db = str(self.date_to_change.year()+(self.type_room//12)) + '-' + str(self.date_to_change.month()) + '-' + str(self.date_to_change.day())  # we update the data of modification
            self.db.insert_date_into_data_base(self.devices[row][0], current_date_for_db, self.type_room, next_date_for_db)  # insert into the data bank and atualize
        self.Load_devices(self.name_room_selected)  # reload the Table Widget

    def insertNewDeviceintoRoom(self):
        """
        here we want to insert new devices to table widget
        :return: new device into Table Widget
        """
        if(self.canInsertID):
            self.saveRoom()
            newDialog = QDialog(self); newDialog.setFont(self.font); newDialog.move(DIALOG_WIDTH_CALENDER,DIALOG_WIDTH_CALENDER), newDialog.setFixedSize(DIALOG_WIDTH_ID, DIALOG_HEIGHT_ID) # Create dialog
            vbox = QVBoxLayout(); newDialog.setWindowTitle(self.laguage.newId), newDialog.setWindowIcon(QIcon(ICON_PTW_LOGO)) # layout
            qlabel = QLabel(); qlabel.setFont(self.font), qlabel.setText(self.laguage.newId_Inset), qlabel.setFixedWidth(DIALOG_WIDTH_CALENDER+10), vbox.addWidget(qlabel) # Set a label
            qbox = QLineEdit();qbox.setFont(self.font), qbox.setPlaceholderText(self.laguage.newId_Insert_Back), qbox.setFixedHeight(LINE_EDITOR_HEIGHT), qbox.setFixedWidth(DIALOG_WIDTH_CALENDER+10), vbox.addWidget(qbox) # Place to insert the id
            hbox = QHBoxLayout(); insertBtt = QPushButton(self.laguage.insert_new_ID); insertBtt.setFont(self.font), insertBtt.setFixedWidth(DIALOG_WIDTH_CALENDER//2), insertBtt.clicked.connect(lambda ch, text = qbox.text(): self.inserNewDeviceButton(qbox.text())),insertBtt.setPalette(QPalette(Qt.blue)),insertBtt.clicked.connect(newDialog.close), hbox.addWidget(insertBtt) # Button to insert
            closeBtt = QPushButton(self.laguage.cancelBtt); closeBtt.setFont(self.font), closeBtt.setFixedWidth(DIALOG_WIDTH_CALENDER//2), closeBtt.setPalette(QPalette(Qt.red)) ,closeBtt.clicked.connect(newDialog.close), hbox.addWidget(closeBtt), vbox.addLayout(hbox) #Close Button
            vbox.setContentsMargins(10,10,10,10), newDialog.setLayout(vbox),newDialog.open()
        else: self.labelProblem.setText(self.laguage.noRoomDetected), self.problemDialog.update(), self.problemDialog.open()

    def inserNewDeviceButton(self, text):
        """
        Here we need to see if the ID already exist or we need to create
        :return: Update into data Bank and atualize Table Widget
        """
        try:
            id = int(text)
            if(id > 0):
                device = self.db.seeIfIDexists(id)  # Check if exits
                if (device): self.db.insert_location_into_data_base(id, self.name_building_selected, self.name_floor_selected,self.name_room_selected, self.type_room)  # We atualize the location
                else:
                    date_now = QDate(datetime.now().year, datetime.now().month, datetime.now().day)  # we create the data of modification
                    date_next = QDate(datetime.now().year + (self.type_room // 12), datetime.now().month, datetime.now().day)  # we create the data of modification
                    date_now = str(date_now.year()) + '-' + str(date_now.month()) + '-' + str(date_now.day());  # we update the data of modification
                    date_next = str(date_next.year()) + '-' + str(date_next.month()) + '-' + str(date_next.day())  # we update the data of modification
                    self.db.CreatenewID(id, self.name_building_selected, self.name_floor_selected, self.name_room_selected,self.db.load_room_id(self.name_building_selected, self.name_floor_selected,self.name_room_selected)[0], self.type_room, date_now,date_next)  # Create new device
                    if self.device_neg != False: self.db.deleteIDs(self.device_neg[0])
                self.Load_devices(self.name_room_selected)
            else: self.labelProblem.setText(self.laguage.negative_ID_message), self.problemDialog.update(), self.problemDialog.open()
        except: self.labelProblem.setText(self.laguage.id_not_number_error), self.problemDialog.update(), self.problemDialog.open()


    def deleteRowAndItem(self):
        """
        We will delete an Item in the room
        :return: None
        """
        try:
            self.undoLimit = self.db.SaveUndoAction(self.undo_database); self.undo_database += 1  # save undo action
            row = self.table_view_main_window.currentRow()
            self.db.deleteIDs(self.devices[row][0])
            self.Load_devices(self.name_room_selected)
        except: self.labelProblem.setText(self.laguage.noIdDetected), self.problemDialog.update(), self.problemDialog.open()

    def saveRoom(self):
        """
        Here we save the information manipulated in the Table Widget
        :return: None
        """
        self.undoLimit = self.db.SaveUndoAction(self.undo_database); self.undo_database += 1  # save undo action
        try:
            self.SaveInformationOfTableWidget() # First we save the information
            for row in range(self.table_view_main_window.rowCount()):
                id_name = int(self.devices[row][0]) # get the id
                date_aux = re.findall(r'\d+', self.table_view_main_window.cellWidget(row, 5).text())
                date_now = QDate(int(date_aux[2]), int(date_aux[0]), int(date_aux[1]))  # we create the data of modification
                date_next = QDate(int(date_aux[2])+(self.type_room//12) , int(date_aux[0]), int(date_aux[1]))  # we create the data of modification
                date_now = str(date_now.year()) + '-' + str(date_now.month()) + '-' + str(date_now.day());  # we update the data of modification
                date_next = str(date_next.year()) + '-' + str(date_next.month()) + '-' + str(date_next.day())  # we update the data of modification
                self.db.insert_date_into_data_base(id_name, date_now, self.type_room, date_next) # we atualize the information about the date
            self.Load_devices(self.name_room_selected)
        except: self.labelProblem.setText(self.laguage.id_sheet_error), self.problemDialog.update(), self.problemDialog.open()

    def SaveAll(self):
        """
        Here we save the current Data Base to the Official Data base
        :return: Data Base Official saved
        """
        if(self.table_view_main_window.rowCount() != 0): self.saveRoom() # Save if necessary
        self.db.CopyBackupToOfficial(NAME_CURRENT_DATA_BASE) # We atualize the official
        self.db = DB_Manipulation() # Reopen the class
        self.saved = True

    def CreateNewLocation(self):
        """
        This option offers the functionality of search the locations available inside the Gossen Metrawatt
        :return: None
        """
        self.WarningDialog(self.laguage.message_tutorial_addLocation)

    def SaveNewLocation(self):
        """
        Here we verify if the receved locations already exists, if they exists we atualize the name, otherwise we create a new ID as default
        :return: None
        """
        fileDialog = QFileDialog();fileDialog.setWindowIcon(QIcon(ICON_PTW_LOGO)); fileDialog.setWindowTitle(self.laguage.title_xml_import)
        importFile = fileDialog.getOpenFileName(fileDialog, 'Import .xml', GROSSENMETRAWATT_DIRECTORY, '.xml (*.xml)'); importFile = PureWindowsPath(importFile[0])
        dialog = QDialog(self); dialog.setFont(self.font), dialog.setWindowIcon(QIcon(ICON_PTW_LOGO)), dialog.setWindowTitle(self.laguage.processing), dialog.move(DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER), dialog.setMinimumSize(DIALOG_WIDTH_EMAIL, DIALOG_HEIGHT_EMAIL); vbox = QVBoxLayout()  # set dialog
        bolStatus = True; qlabe = QLabel(self.laguage.processing); qlabe.setFont(self.font), qlabe.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), vbox.addWidget(qlabe)
        processBar = QProgressBar(dialog); processBar.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), vbox.addWidget(processBar)
        closeBtt = QPushButton(self.laguage.cancelBtt); closeBtt.setFont(self.font), closeBtt.setPalette(QPalette(Qt.red)), closeBtt.setFixedWidth( DIALOG_BUTTONS_WIDTH), closeBtt.clicked.connect(dialog.close), closeBtt.clicked.connect(lambda bolean: not bolStatus), vbox.addWidget(closeBtt), vbox.setAlignment(closeBtt, Qt.AlignCenter)
        dialog.setLayout(vbox), dialog.adjustSize(), dialog.open(), QCoreApplication.processEvents()
        try:
            self.undoLimit = self.db.SaveUndoAction(self.undo_database); self.undo_database+=1 # Save undo database
            processBar.setMaximum(5), QCoreApplication.processEvents()
            grossenMetraWatt = GMW(importFile, False)  # get the path
            if (bolStatus): locations = grossenMetraWatt.readAllLocations();  processBar.setValue(2) # first we load all the available locations
            if(bolStatus): self.db.inserNewLocation(locations); processBar.setValue(4)
            if(bolStatus): self.SaveAll();
            processBar.setValue(5), closeBtt.setPalette(QPalette(Qt.blue)), QCoreApplication.processEvents()
            closeBtt.setText(self.laguage.done), closeBtt.clicked.connect(self.main_window), qlabe.setText(self.laguage.operation_concluded), QCoreApplication.processEvents()
        except: dialog.close(), self.labelProblem.setText(self.laguage.file_not_supported), self.problemDialog.update(), self.problemDialog.show()

    def Radio_Button_Type(self, name):
        """
        Here we atualize the type of the room
        :param name: can be Office or Labor
        :return: (self).type_room (int): the number of months for the next verification
        """
        if(name == 'Office'): self.type_room = OFFICE # atualize for office
        else: self.type_room = LABOR # atualize for labor

    def BrainOFSearch(self):
        """
        This function find the locations to the functionality required
        :param firstDate: The first Date of the intervall of interest
        :param lastDate:  the last date of the intervall of interest
        :param typeOfSelection: If we want the Not Verified or the Verified
        :return: The location required to be evaluated
        """
        # Due to problem with the Gossen Metrawatt, this operation is required to correct missing data
        def CreateDateNextVerication(deviceDate, numberOFMonths):
            return date(int(deviceDate[0]) + (numberOFMonths // 12), int(deviceDate[1]), int(deviceDate[2]))

        def TransformDateToLanguage(qdate, languageSetting):
            if(languageSetting == DEFAULT_LANGUAGE_ENGLISH): return str(qdate.month) + "/" + str(qdate.day) + "/"+ str(qdate.year)
            else: return str(qdate.day) + "/" + str(qdate.month) + "/"+ str(qdate.year)

        def ReTransformDateToLanguage(qdate, languageSetting):
            qdate = re.findall(r'\d+', qdate)
            if(languageSetting == DEFAULT_LANGUAGE_ENGLISH): return date(int(qdate[2]), int(qdate[0]), int(qdate[1]))
            else: return date(int(qdate[2]), int(qdate[1]), int(qdate[0]))

        # Brain
        sys.setrecursionlimit(1500)
        devices, location = self.db.loadAll(), list()  # we receive all devices
        for device in devices:
            if (int(device[0]) > 0):
                if (self.roomNotVerified):  # Not Verified (default)
                    deviceDate = re.findall(r'\d+', str(device[11]))
                    try: deviceDate = date(int(deviceDate[0]), int(deviceDate[1]), int(deviceDate[2]))
                    except: deviceDate = CreateDateNextVerication(re.findall(r'\d+', str(device[5])), int(device[10]))
                else:
                    deviceDate = re.findall(r'\d+', str(device[5]))
                    deviceDate = date(int(deviceDate[0]), int(deviceDate[1]), int(deviceDate[2]))
                if (self.firstDate <= deviceDate <= self.lastDate): location.append([device[6], device[7], device[8], TransformDateToLanguage(deviceDate, self.currentLanguageSetting)])  # we save the building, floor and room
        try:
            if (location != []):
                location = sorted(location)
                if(len(location)>1):
                    location = sr.CleanRepetitivePositions(location[len(location)-1], location, len(location)-2)
                    for j in range(len(location)-1, 0, -1):
                        building, floor, room, dateFinal = location[j][0], location[j][1], location[j][2], ReTransformDateToLanguage(location[j][3], self.currentLanguageSetting)
                        for i in range(j, -1, -1):
                            building_c, floor_c, room_c, date_c = location[i][0], location[i][1], location[i][2], ReTransformDateToLanguage(location[i][3], self.currentLanguageSetting)
                            if(building == building_c) and (floor == floor_c) and (room_c == room):
                                if date_c < dateFinal: dateFinal = date_c; location[j][3] = TransformDateToLanguage(dateFinal, self.currentLanguageSetting)
                                else: date_c = dateFinal;  location[i][3] = TransformDateToLanguage(date_c, self.currentLanguageSetting)
                    location = sr.CleanRepetitivePositions(location[len(location)-1], location, len(location)-2)
            return location
        except: return []

    def TableSeeRoom(self, location, contactsFound):
        """
        Create Table and Adjust the Data
        :return: Table Atualized
        """
        self.table_view.clear()
        self.table_view.setColumnCount(len(self.columnHeader)), self.table_view.setHorizontalHeaderLabels(self.columnHeader), self.table_view.horizontalHeader().setFont(self.font) # Add the Columns
        if contactsFound == None: aux_list, aux_list_2 = [1,4], [0,1,2,3,4]
        else: aux_list,  aux_list_2 = [0,3] ,[0,1,2,3]
        for i in aux_list_2: self.table_view.horizontalHeader().setSectionResizeMode(i,QHeaderView.ResizeToContents)
        for i in aux_list: self.table_view.horizontalHeader().setSectionResizeMode(i,QHeaderView.Stretch)  # adjust the size of the Table view to auto-adjust
        self.table_view.setEditTriggers(QAbstractItemView.NoEditTriggers)
        if contactsFound == None: self.table_view.setRowCount(len(location))
        else: self.table_view.setRowCount(len(contactsFound))
        self.table_view.reset(), QCoreApplication.processEvents()
        if location != None:
            if (contactsFound == None):
                for row in range(len(location)):
                    for column in range(len(location[0])):
                        if(column == 0): item = QCheckBox(checkable=True); item.setStyleSheet("QCheckBox::indicator { width:"+str(DIALOG_SIZE_CHECKBOX[0])+"; height:"+ str(DIALOG_SIZE_CHECKBOX[1])+";}"),item.setChecked(False); self.table_view.setCellWidget(row, column,  item) ,self.table_view.isEnabledTo(item), item.clicked.connect(lambda ch, index = self.table_view.currentRow(): self.CheckedCheckBoxSeeRooms(self.table_view.currentRow()) )
                        self.table_view.setItem(row, column + 1, QTableWidgetItem(str(location[row][column])))
            else:
                for row in range(len(contactsFound)):
                    for column in range(len(contactsFound[0])):
                        self.table_view.setItem(row, column, QTableWidgetItem(str(contactsFound[row][column])))

    def ComboBoxDelegateSeeRooms(self, index):
        """
        Here if the combo Box is selected, the Table View is actualized
        :param index: correspondig type of evalution required
        :return: Table View actualized
        """
        self.positionsClicked = list()
        if(index == 0): # Not Verified
            self.qlabel1_to.hide(), self.buttonFirstDate.hide()
            self.roomNotVerified, self.columnHeader, self.lastDate, self.firstDate = True, self.laguage.tableShowResultsHeader_NotVerified_1,self.buttonLastDate.date(), QDate(1900, 1, 1)
            self.buttonFirstDate.setEnabled(False), QCoreApplication.processEvents()
            self.qlabel1_untill.setText(self.laguage.untill)
        else:
            self.qlabel1_to.show(), self.buttonFirstDate.show()
            self.roomNotVerified, self.columnHeader, self.lastDate, self.firstDate = False, self.laguage.tableShowResultsHeader_NotVerified_1, self.buttonLastDate.date(), self.buttonFirstDate.date()
            self.buttonFirstDate.setEnabled(True), QCoreApplication.processEvents()
            self.qlabel1_untill.setText(self.laguage.fromlabel)
        self.TableSeeRoom(self.BrainOFSearch(), self.contactsFound)

    def SeeRoomsNotVerified(self):
        """
        we create a dialog with the rooms to verify
        :return: Dialog
        """
        # Default settings
        self.lastDate, self.firstDate, self.roomNotVerified, self.columnHeader, self.positionsClicked = QDate(datetime.now().year, datetime.now().month, datetime.now().day), QDate(1900, 1, 1), True, self.laguage.tableShowResultsHeader_NotVerified_1, []  # we create the data of modification
        month = datetime.now().month-1
        if(month<1): firstDate = QDate(datetime.now().year - 1, 12, datetime.now().day)  # we create the data of modification
        elif(month>12): firstDate = QDate(datetime.now().year + 1, 1, datetime.now().day)  # we create the data of modification
        else: firstDate = QDate(datetime.now().year-1, month, datetime.now().day)  # we create the data of modification
        self.showDialog = QDialog(self); self.showDialog.setFont(self.font); self.showDialog.move(DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER)
        self.showDialog.setWindowTitle(self.laguage.titleSearch), self.showDialog.setWindowIcon(QIcon(ICON_PTW_LOGO))
        locations, contacts, self.contactsFound = self.BrainOFSearch(), [], None
        # Create the Table Widget
        vbox = QVBoxLayout() # additionals
        self.table_view = QTableWidget()  # Create the Table view
        self.table_view.setGeometry(0, 0, MINIMU_WIDTH_TABLE_VIEW_MAIN_WINDOW, MINIMU_HEIGHT_TABLE_VIEW_MAIN_WINDOW), self.table_view.setMinimumSize(MAIN_WINDOW_WIDTH, MINIMU_HEIGHT_TABLE_VIEW_MAIN_WINDOW)  # geometry of the table view
        self.table_view.setShowGrid(True), self.table_view.setFont(self.font), self.table_view.setStyleSheet(TABLE_VIEW_STYLE)  # Define as fixed the number of columns
        self.TableSeeRoom(locations, None), vbox.addWidget(self.table_view), self.table_view.setSelectionBehavior(QTableView.SelectRows)
        # Create the UI, which the user can interect
        vbox2, hbox, vbox3, hbox1, hbox2 = QVBoxLayout(), QHBoxLayout(), QVBoxLayout(), QHBoxLayout(), QHBoxLayout()
        self.comboBoxVerifyNotVerify = QComboBox(); self.comboBoxVerifyNotVerify.setFont(self.font), self.comboBoxVerifyNotVerify.addItems(self.laguage.comboBoxData) , self.comboBoxVerifyNotVerify.activated.connect(lambda ch, index = self.comboBoxVerifyNotVerify.currentIndex(): self.ComboBoxDelegateSeeRooms(self.comboBoxVerifyNotVerify.currentIndex())), self.comboBoxVerifyNotVerify.setFixedWidth(DIALOG_SIZE_WIDTH_DATE) ,vbox2.addWidget(self.comboBoxVerifyNotVerify)
        # For Verified Rooms, Widget necessary
        self.qlabel1_untill = QLabel(self.laguage.untill); self.qlabel1_untill.setFont(self.font), hbox1.addWidget(self.qlabel1_untill)
        self.buttonLastDate = QDateEdit(calendarPopup=True); self.buttonLastDate.setFont(self.font) ,self.buttonLastDate.setDate(self.lastDate), self.buttonLastDate.setFixedSize(DIALOG_SIZE_WIDTH_DATE, DIALOG_SIZE_HEIGHT_DATE) , hbox1.addWidget(self.buttonLastDate), hbox1.setAlignment(Qt.AlignLeft), vbox2.addLayout(hbox1)
        self.qlabel1_to = QLabel(self.laguage.tolabel); self.qlabel1_to.setFont(self.font), self.qlabel1_to.hide(), hbox2.addWidget(self.qlabel1_to)
        self.buttonFirstDate = QDateEdit(calendarPopup=True); self.buttonFirstDate.setFont(self.font), self.buttonFirstDate.setDate(firstDate), self.buttonFirstDate.hide(), self.buttonFirstDate.setFixedSize(DIALOG_SIZE_WIDTH_DATE,DIALOG_SIZE_HEIGHT_DATE), hbox2.addWidget(self.buttonFirstDate), hbox2.setAlignment(Qt.AlignLeft), vbox2.addLayout(hbox2)
        if(self.currentLanguageSetting == DEFAULT_LANGUAGE_ENGLISH): self.buttonLastDate.setDisplayFormat("M/d/yyyy"), self.buttonFirstDate.setDisplayFormat("M/d/yyyy")
        else: self.buttonLastDate.setDisplayFormat("d/M/yyyy"), self.buttonFirstDate.setDisplayFormat("d/M/yyyy")
        self.buttonFirstDate.setEnabled(False), self.buttonFirstDate.dateChanged.connect(lambda ch, i = 1: self.ComboBoxDelegateSeeRooms(1)), self.buttonLastDate.dateChanged.connect(lambda ch, i = self.comboBoxVerifyNotVerify.currentIndex(): self.ComboBoxDelegateSeeRooms(self.comboBoxVerifyNotVerify.currentIndex())) # we all start withthe not verified Rooms
        # For not Verfied Rooms, Widgets necessary
        self.findResponsibles = QPushButton(self.laguage.searchResponsible); self.findResponsibles.setFont(self.font) ,self.findResponsibles.setFixedWidth(DIALOG_WIDTH_NOTVERFIED_BUTTON), vbox3.addWidget(self.findResponsibles); self.findResponsibles.clicked.connect(self.FindResponsibles) #Button to send E-Mail to responsible
        self.closebtt = QPushButton(self.laguage.closeBtt); self.closebtt.setFont(self.font), self.closebtt.setPalette(QPalette(Qt.red)) ,self.closebtt.clicked.connect(self.showDialog.close), self.closebtt.setFixedWidth(DIALOG_WIDTH_NOTVERFIED_BUTTON), vbox3.addWidget(self.closebtt), vbox3.setAlignment(Qt.AlignCenter)
        hbox.addLayout(vbox2), hbox.addLayout(vbox3), vbox.addLayout(hbox), vbox.setContentsMargins(10,10,10,10), self.showDialog.setLayout(vbox), self.showDialog.adjustSize(), self.showDialog.show()


    def CheckedCheckBoxSeeRooms(self, index):
        """
        Load the positions Clicked and wanted to find the responsibles
        :param index: row clicked
        :return: self.positionsClicked actualized with positions required to find inside the website
        """
        try:
            self.positionsClicked.index(index)
            del self.positionsClicked[self.positionsClicked.index(index)]
        except: self.positionsClicked.append(index)
        self.positionsClicked = sorted(self.positionsClicked)

    def FindResponsibles(self):
        """
        Here we implement the class SearchResponsibles, and find the name and e-mail from all the people selected
        :return: will return the list with the name of responsibles and e-mail of him or them
        """
        #Configurate the date option
        self.buttonLastDate.hide(), self.qlabel1_untill.hide(), self.comboBoxVerifyNotVerify.setEnabled(False), self.buttonFirstDate.disconnect(), self.qlabel1_to.setText(self.laguage.choose_appointment), self.qlabel1_to.show(), self.buttonFirstDate.show(), QCoreApplication.processEvents()
        self.buttonFirstDate.setDate(QDate(datetime.now().year, datetime.now().month, datetime.now().day).addDays(7) )# Increase one week
        dialog = QDialog(self);
        dialog.setFont(self.font), dialog.setWindowIcon(QIcon(ICON_PTW_LOGO)), dialog.setWindowTitle( self.laguage.processing), dialog.move(DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER), dialog.setMinimumSize(DIALOG_WIDTH_EMAIL, DIALOG_HEIGHT_EMAIL)  # set dialog
        vbox = QVBoxLayout(); qlabe = QLabel(self.laguage.processing); qlabe.setFont(self.font), qlabe.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), vbox.addWidget(qlabe)
        processBar = QProgressBar(dialog); processBar.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), processBar.setMaximum(10), processBar.setValue(0) ,vbox.addWidget(processBar)
        dialog.setLayout(vbox), dialog.adjustSize(), dialog.open(), QCoreApplication.processEvents()
        if(self.positionsClicked != []):
            location, floor = list(), list()
            for i in range(self.table_view.rowCount()):
                if i in self.positionsClicked: location.append([str(self.table_view.item(i, 1).text()), str(self.table_view.item(i, 3).text())]), floor.append(str(self.table_view.item(i, 2).text()))
            processBar.setValue(2)
            contacts, emails = sr.FindResponsibles(), []; processBar.setValue(5)
            if contacts == Exception:
                processBar.close(), qlabe.setText(self.laguage.error_message_finding_responsibles), QCoreApplication.processEvents()
            else:
                contacts, contactsFound, floorID, rooms_label = sr.CleanContactsToNameandEmail(contacts), [], 0, [] ; processBar.setValue(7)
                for l in location:
                    elementsFound = sr.FindResponsibleForEachRoom(l, contacts)
                    if elementsFound == []: contactsFound.append([l[0], floor[floorID] ,l[1], self.laguage.notFound]); rooms_label.append(l[1])
                    else:
                        for k in elementsFound: contactsFound.append([l[0], floor[floorID] ,l[1], k[0]]), emails.append(k[0]); rooms_label.append(l[1])
                    floorID += 1
                processBar.setValue(9)
                self.columnHeader = self.laguage.tableShowResultsHeader_NotVerified_2; self.TableSeeRoom(1, contactsFound)
                self.findResponsibles.setText(self.laguage.send_email_to_responsibles), processBar.setValue(10), dialog.close(), self.findResponsibles.disconnect()
                if self.roomNotVerified: self.findResponsibles.clicked.connect(lambda ch, email= emails: self.SendEMailResponsablesForRoom(email)), self.buttonFirstDate.setEnabled(True)
                else: self.findResponsibles.setDisabled(True), QCoreApplication.processEvents(), self.buttonFirstDate.setEnabled(False)
                if len(rooms_label) >= 2:
                    rooms_label = sr.CleanRepetitivePositions(rooms_label[len(rooms_label)-1], rooms_label, len(rooms_label)-2)
                self.rooms_label = ''
                for room in rooms_label: self.rooms_label += str(room) + ' '
        else: processBar.close(), qlabe.setText(self.laguage.error_message_finding_responsibles_2), QCoreApplication.processEvents()

    def SendEMailResponsablesForRoom(self, emails):
        """
        The idea is to send an automatic e-mail to all responsibles, which were previous selected
        :return: None
        """
        def get_contacts(filename): # we open the file w/ the contacts
            names = []; emails = []
            with open(filename, mode='r', encoding='utf-8') as contacts_file:
                for a_contact in contacts_file:
                    if (a_contact == "\n"):
                        break
                    else:
                        name, email, i = '', '', 0
                        while i < len(a_contact):
                            if(a_contact[i] != ':'): name += a_contact[i]
                            else: i = 10000
                            i += 1
                        email = a_contact.split()[len(a_contact.split())-1]
                        names.append(name)
                        emails.append(email)
            return emails

        def read_template(filename): # Read the message
            with open(filename, 'r', encoding='utf-8') as template_file:
                template_file_content = template_file.read()
            return Template(template_file_content)

        # Dialog of processing
        date_title = self.buttonFirstDate.date().toString(Qt.LocaleDate)
        for e in get_contacts(EMAIL_CONTACTS_ROOM_RESPONSABLES): emails.append(e)
        dialog = QDialog(self); dialog.setFont(self.font), dialog.setWindowIcon(QIcon(ICON_PTW_LOGO)), dialog.setWindowTitle( self.laguage.processing), dialog.move(DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER), dialog.setMinimumSize( DIALOG_WIDTH_EMAIL, DIALOG_HEIGHT_EMAIL)  # set dialog
        vbox = QVBoxLayout(); qlabe = QLabel(self.laguage.processing); qlabe.setFont(self.font), qlabe.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), vbox.addWidget(qlabe)
        processBar = QProgressBar(dialog); processBar.setMaximum(7), processBar.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), vbox.addWidget(processBar)
        closeBtt = QPushButton(self.laguage.done); closeBtt.setFont(self.font), closeBtt.setPalette(QPalette(Qt.red)), closeBtt.setFixedWidth(DIALOG_BUTTONS_WIDTH), closeBtt.clicked.connect(dialog.close), vbox.addWidget(closeBtt), vbox.setAlignment(closeBtt, Qt.AlignCenter)
        dialog.setLayout(vbox), dialog.adjustSize(), dialog.open(), QCoreApplication.processEvents()
        # send e-mail
        try:
            self.InializateOutlook()
            outlook = win32.Dispatch('outlook.application')
            mail, email_contact = outlook.CreateItem(0), ""
            message_template = read_template(MESSAGE_EMAIL_ROOM_RESPONSABLES)
            processBar.setValue(4)
            # For each contact, send the email:
            for e in emails: email_contact += str(e) + ";"
            # add in the actual person name to the message template
            try:
                try:
                    try: message = message_template.substitute(TAG=date_title, RAEUME=self.rooms_label)
                    except: message = message_template.substitute(RAEUME=self.rooms_label)
                except: message = message_template.substitute(TAG=date_title)
            except: message = message_template.substitute(None)
            # setup the parameters of the messagelogin
            mail.To = email_contact
            mail.Subject = self.laguage.titleEmailSendToResponsibles
            # add in the message body
            mail.Body = message
            mail.Display(True)
            processBar.setValue(6)
            qlabe.setText(self.laguage.done), processBar.setValue(7), dialog.update()
        except:  processBar.setValue(7), qlabe.setPalette(QPalette(Qt.red)), qlabe.setText(self.laguage.error_message_send_email), qlabe.update(), dialog.update(), QCoreApplication.processEvents()

    def ImportExportFile(self, typeOfAction):
        """
        Here we open the dialog to import or export a file
        :param typeOfAction: We have 3 different Actions: Export BackUp, Import BackUp or Actualize Gossen Metrawatt
        :return: None
        """
        endDialog = QDialog(self); endDialog.move(DIALOG_WIDTH_CALENDER,DIALOG_WIDTH_CALENDER), endDialog.setFixedSize(DIALOG_OPERATION_CONCLUEDED_WIDHT, DIALOG_PROBLEM_HEIGHT-50)  # create Dialog
        vbox = QVBoxLayout(); endDialog.setWindowTitle(self.laguage.done), endDialog.setWindowIcon(QIcon(ICON_PTW_LOGO))
        qlabel = QLabel(self.laguage.operation_concluded); qlabel.setFont(self.font),  qlabel.setFixedWidth( DIALOG_OPERATION_CONCLUEDED_WIDHT- 10), vbox.addWidget(qlabel)
        closeBtt = QPushButton(self.laguage.closeBtt, endDialog); closeBtt.setFont(self.font), closeBtt.clicked.connect(endDialog.close), closeBtt.setFixedWidth(DIALOG_OPERATION_CONCLUEDED_WIDHT- 40), closeBtt.setPalette(QPalette(Qt.darkRed)), vbox.addWidget(closeBtt)
        vbox.setAlignment(Qt.AlignCenter), endDialog.setLayout(vbox)
        fileDialog = QFileDialog(); fileDialog.setWindowIcon(QIcon(ICON_PTW_LOGO))
        try:
            if(typeOfAction == 0):
                self.undoLimit = self.db.SaveUndoAction(self.undo_database); self.undo_database+=1 # save undo action
                fileDialog.setWindowTitle(self.laguage.importBU)
                ImportDataBaseName = fileDialog.getOpenFileName(fileDialog,'Import DB', DIRECTORY_IMPORT_FILES,'DB (*.db)')
                if ImportDataBaseName[0] != "":
                    ImportDataBaseName = PureWindowsPath(ImportDataBaseName[0])
                    exist = self.db.CheckIFBackupISCompatible(str(ImportDataBaseName))
                    if not exist: self.labelProblem.setText(self.laguage.dataBank_not_compatible), self.problemDialog.update(), self.problemDialog.open()
                    else:  endDialog.open(), self.db.CopyBackupToOfficial(str(ImportDataBaseName)); self.saved = True; self.main_window()
            elif(typeOfAction == 1):
                fileDialog.setWindowTitle(self.laguage.exportBU)
                fileToExportName = DIRECTORY_IMPORT_FILES + '\\BackUP_EPruefung ' + str(datetime.now().year) +'-' + str(datetime.now().month) +'-'+str(datetime.now().day)
                nameOfFile = fileDialog.getSaveFileName(fileDialog,'Export DB', fileToExportName,'DB (*.db)'); nameOfFile = PureWindowsPath(nameOfFile[0])
                self.db.CopyOfficialToBackUP(str(nameOfFile)); endDialog.open()
            else:
                if (typeOfAction == 2): self.WarningDialog(self.laguage.ExportWarningMessage)
                else: self.WarningDialog(self.laguage.ImportWarningMessage)
                if self.ContinuesAction:
                    self.undoLimit = self.db.SaveUndoAction(self.undo_database); self.undo_database+=1 # save undo action
                    fileDialog.setWindowTitle(self.laguage.title_xml_import)
                    importFile = fileDialog.getOpenFileName(fileDialog, 'Import .xml', GROSSENMETRAWATT_DIRECTORY, '.xml (*.xml)')
                    if importFile[0] != "":
                        importFile = PureWindowsPath(importFile[0])
                        if(typeOfAction == 2): self.GrossenMetrawattDialog(importFile)
                        else: self.ImportDataFROMGrossenMetrawatt(importFile)
        except: self.labelProblem.setText(self.laguage.operation_not_concluded), self.problemDialog.update(), self.problemDialog.open()

    def GrossenMetrawattDialog(self, filePath):
        """
        Here we create a dialog to comunicate with the dialog
        :param filePath: is the file .xml opened
        :return: New file atualized
        """
        dialog = QDialog(self);
        dialog.setFont(self.font), dialog.setWindowIcon(QIcon(ICON_PTW_LOGO)), dialog.setWindowTitle(self.laguage.processing), dialog.move(DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER), dialog.setMinimumSize(DIALOG_WIDTH_EMAIL, DIALOG_HEIGHT_EMAIL)  # set dialog
        vbox = QVBoxLayout(); bolStatus = True
        qlabe = QLabel(self.laguage.processing); qlabe.setFont(self.font), qlabe.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), vbox.addWidget(qlabe)
        processBar = QProgressBar(dialog); processBar.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), vbox.addWidget(processBar)
        closeBtt = QPushButton(self.laguage.cancelBtt); closeBtt.setFont(self.font), closeBtt.setPalette(QPalette(Qt.red)), closeBtt.setFixedWidth(DIALOG_BUTTONS_WIDTH), closeBtt.clicked.connect(dialog.close), closeBtt.clicked.connect(lambda bolean: not bolStatus), vbox.addWidget(closeBtt), vbox.setAlignment(closeBtt, Qt.AlignCenter)
        dialog.setLayout(vbox), dialog.adjustSize(), dialog.open(), QCoreApplication.processEvents()
        try:
            devices = self.db.loadAll()  # get all devices
            processBar.setMaximum(len(devices)+2), QCoreApplication.processEvents()
            grossenMetraWatt = GMW(filePath, True) # get the path
            for device in devices:
                id, name, company, inv_num, note, building, floor, room, locationID, room_type = str(device[0]), str(device[1]), str(device[2]), str(device[3]), str(device[4]),  str(device[6]), str(device[7]), str(device[8]), str(device[9]), str(device[10])
                grossenMetraWatt.insertIntoXMLFile(locationID, name, company, inv_num, note, room_type, id)
                processBar.setValue(processBar.value()+1)
                if(not bolStatus): break
            grossenMetraWatt.TransformAuxiliarFileIntoDef()
            processBar.setValue(len(devices)+2), closeBtt.setPalette(QPalette(Qt.blue)), QCoreApplication.processEvents()
            closeBtt.setText(self.laguage.done), qlabe.setText(self.laguage.operation_concluded),QCoreApplication.processEvents()
        except: dialog.close(), self.labelProblem.setText(self.laguage.file_not_supported) ,self.problemDialog.update() ,self.problemDialog.show()

    def ImportDataFROMGrossenMetrawatt(self, filePath):
        """
        Here we import all devices available in Gossen Metrawatt if we lost the information
        :param filePath: is the file .xml opened
        :return: return the data base actualiyed, however it isn't saved
        """
        dialog = QDialog(self); dialog.setFont(self.font), dialog.setWindowIcon(QIcon(ICON_PTW_LOGO)), dialog.setWindowTitle(self.laguage.processing), dialog.move(DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER), dialog.setMinimumSize(DIALOG_WIDTH_EMAIL, DIALOG_HEIGHT_EMAIL)  # set dialog
        vbox = QVBoxLayout(); bolStatus = True; qlabe = QLabel(self.laguage.processing); qlabe.setFont(self.font), qlabe.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), vbox.addWidget(qlabe) # Dialog Labels
        processBar = QProgressBar(dialog);  processBar.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), vbox.addWidget(processBar) # Progress bar
        closeBtt = QPushButton(self.laguage.cancelBtt); closeBtt.setFont(self.font), closeBtt.setPalette(QPalette(Qt.red)), closeBtt.setFixedWidth(DIALOG_BUTTONS_WIDTH), closeBtt.clicked.connect(dialog.close), closeBtt.clicked.connect(lambda bolean: not bolStatus), vbox.addWidget(closeBtt), vbox.setAlignment(closeBtt, Qt.AlignCenter) # Buttons
        dialog.setLayout(vbox), dialog.adjustSize(), dialog.open(), QCoreApplication.processEvents()
        try:
            processBar.setMaximum(10) # set the value of
            grossenMetraWatt = GMW(filePath, False); processBar.setValue(1) # incialize the class
            if bolStatus: locations = grossenMetraWatt.readAllLocations();  processBar.setValue(2) # first we load all the available locations
            if bolStatus: self.db.inserNewLocation(locations); processBar.setValue(4) # we save all the locations available
            if bolStatus:
                for i in locations: # loop to colect all devices in each location
                    devices = grossenMetraWatt.getDataFromXMLFile(i[3])
                    if devices != False and devices != None: # verify the locations contends devices
                        for device in devices:  # insert into the data base
                            if not self.db.seeIfIDexists(device[0]): self.db.CreatenewID(device[0], i[0], i[1], i[2], i[3], device[7], device[3], device[5])
                            else: self.db.insert_date_into_data_base(device[0], device[3], device[7], device[5]); self.db.insert_location_into_data_base(device[0],  i[0], i[1], i[2], i[3])
                            self.db.insertInformationIntoCurrent(device[0], device[1], device[4], device[6], device[2])
                processBar.setValue(9)
            if bolStatus: self.SaveAll(), self.main_window()
            processBar.setValue(10), closeBtt.setText(self.laguage.done), qlabe.setText(self.laguage.operation_concluded),  QCoreApplication.processEvents()
        except: dialog.close(), self.labelProblem.setText(self.laguage.file_not_supported) ,self.problemDialog.update() ,self.problemDialog.show()

    def WarningDialog(self, message):
        """
        This is a function able to generate a warning dialog
        :param message: this is what message is wanted
        :return: None
        """
        self.ContinuesAction = True
        newDialog = QDialog(); newDialog.setFont(self.font); newDialog.move(DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER), newDialog.setFixedSize(DIALOG_ADD_NEW_LOCATION_WIDTH, DIALOG_ADD_NEW_LOCATION_HEIGTH)  # Create
        vbox = QVBoxLayout(); newDialog.setWindowTitle(self.laguage.warning), newDialog.setWindowIcon(QIcon(ICON_PTW_LOGO))  # layout
        label = QLabel(); label.setFont(self.font); label.setAlignment(Qt.AlignJustify)
        if message == self.laguage.message_tutorial_addLocation: label.setText(self.laguage.message_tutorial_addLocation)
        elif message == self.laguage.ImportWarningMessage: label.setText(self.laguage.ImportWarningMessage)
        else: label.setText(self.laguage.ExportWarningMessage)
        vbox.addWidget(label)
        hbox = QHBoxLayout();
        if message == self.laguage.message_tutorial_addLocation: insertBtt = QPushButton(self.laguage.addLocation); insertBtt.clicked.connect(self.SaveNewLocation)
        else: insertBtt = QPushButton('OK'); insertBtt.clicked.connect(lambda ch, b = True: self.ChangeStateOfContinuesAction(True))
        insertBtt.setFont(self.font), insertBtt.setFixedWidth((DIALOG_ADD_NEW_LOCATION_WIDTH - 20) // 2), insertBtt.setPalette(QPalette(Qt.blue)), insertBtt.clicked.connect(newDialog.close), hbox.addWidget(insertBtt)  # Button to save
        closeBtt = QPushButton(self.laguage.cancelBtt); closeBtt.setFont(self.font), closeBtt.setFixedWidth((DIALOG_ADD_NEW_LOCATION_WIDTH - 20) // 2), closeBtt.setPalette(QPalette(Qt.red)), closeBtt.clicked.connect(newDialog.close), hbox.addWidget(closeBtt), hbox.setContentsMargins(0, 5, 0, 0), vbox.addLayout(hbox)  # Close Button
        closeBtt.clicked.connect(lambda b = False: self.ChangeStateOfContinuesAction(False)); newDialog.closeEvent = self.ChangeStateOfContinuesAction(False)
        vbox.setContentsMargins(10, 10, 10, 10), newDialog.setLayout(vbox); newDialog.exec_()

    def ChangeStateOfContinuesAction(self, bolean):
        """
        Here we want to say if we continues with the process or not
        :return: None
        """
        self.ContinuesAction = bolean


    def helpPDF(self, namePDF):
        """
        We open the PDF required
        :param: Open the PDF
        :return: None
        """
        wd.open_new(str(namePDF))

    def InializateOutlook(self):
        """
        Here we inializite the Outlook Application
        :return: None
        """
        for p in psutil.process_iter(attrs=['pid', 'name']):
            if "OUTLOOK.EXE" in p.info['name']:
                print("Yes", p.info['name'], "is running")
                break
        else:
            print("No, Outlook is not running")
            os.startfile("outlook")
            print("Outlook is starting now...")

    def SendEmailForResponsible(self):
        """
        Here we create a e-mail to send for the responsible
        :return: None
        """
        def get_contacts(filename): # we open the file w/ the contacts
            names = []; emails = []
            with open(filename, mode='r', encoding='utf-8') as contacts_file:
                for a_contact in contacts_file:
                    if (a_contact == "\n"):
                        break
                    else:
                        name, email, i = '', '', 0
                        while i < len(a_contact):
                            if(a_contact[i] != ':'): name += a_contact[i]
                            else: i = 10000
                            i += 1
                        email = a_contact.split()[len(a_contact.split())-1]
                        names.append(name)
                        emails.append(email)
            return names, emails

        def read_template(filename): # Read the message
            with open(filename, 'r', encoding='utf-8') as template_file:
                template_file_content = template_file.read()
            return Template(template_file_content)

        # Dialog of processing
        dialog = QDialog(self); dialog.setFont(self.font), dialog.setWindowIcon(QIcon(ICON_PTW_LOGO)), dialog.setWindowTitle(self.laguage.processing), dialog.move(DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER), dialog.setMinimumSize( DIALOG_WIDTH_EMAIL, DIALOG_HEIGHT_EMAIL)  # set dialog
        vbox = QVBoxLayout();
        qlabe = QLabel(self.laguage.processing);qlabe.setFont(self.font), qlabe.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), vbox.addWidget(qlabe)
        processBar = QProgressBar(dialog); processBar.setMaximum(7), processBar.setFixedWidth(DIALOG_WIDTH_EMAIL - 20), vbox.addWidget(processBar)
        closeBtt = QPushButton(self.laguage.done); closeBtt.setFont(self.font), closeBtt.setPalette(QPalette(Qt.red)), closeBtt.setFixedWidth(DIALOG_BUTTONS_WIDTH), closeBtt.clicked.connect(dialog.close), vbox.addWidget(closeBtt), vbox.setAlignment(closeBtt, Qt.AlignCenter)
        dialog.setLayout(vbox), dialog.adjustSize(), dialog.open(), QCoreApplication.processEvents()
        # Send message via outlook
        try:
            self.InializateOutlook()
            outlook = win32.Dispatch('outlook.application')
            names, emails = get_contacts(CONTACTS_EMAIL)
            processBar.setValue(3)
            if(self.currentLanguageSetting == DEFAULT_LANGUAGE_ENGLISH): message_template = read_template(MESSAGE_EMAIL_ENGLISH)
            else: message_template = read_template(MESSAGE_EMAIL_GERMAN)
            processBar.setValue(4)
            # Load the file of back UP to load
            fileToExportName = DIRECTORY_IMPORT_FILES + '\\BackUP_EPruefung ' + str(datetime.now().year) + '-' + str(datetime.now().month) + '-' + str(datetime.now().day) + '.db'
            self.db.CopyOfficialToBackUP(fileToExportName); # Create Back-UP
            processBar.setValue(5)
            # For each contact, send the email:
            for name, email in zip(names, emails):
                mail = outlook.CreateItem(0)
                # add in the actual person name to the message template
                try: message = message_template.substitute(PERSON_NAME=name.title())
                except: message = message_template.substitute(None)
                # setup the parameters of the message
                mail.To = email
                mail.Subject = self.laguage.backupDayMessage
                # add in the message body
                mail.Body = message
                # Attach the Back-UP
                mail.Attachments.Add(fileToExportName)
                mail.Display(True)
            processBar.setValue(6)
            with open(fileToExportName, 'w') as f: f.close()
            os.remove(fileToExportName)
            qlabe.setText(self.laguage.done),processBar.setValue(7), dialog.update()
        except: processBar.setValue(7), qlabe.setPalette(QPalette(Qt.red)), qlabe.setText(self.laguage.error_message_send_email),qlabe.update() ,dialog.update(), QCoreApplication.processEvents()

    def closeEvent(self, event):
        """
        We verified if the system was closed, then we can not forget to save
        :param: (self).saved (boolean): verified if the system was saved
        :return:
        """
        if not (self.saved):
            newDialog = QDialog();newDialog.setFont(self.font); newDialog.move(DIALOG_WIDTH_CALENDER, DIALOG_WIDTH_CALENDER), newDialog.setFixedSize(DIALOG_WIDTH_ID,DIALOG_HEIGHT_ID)  # Create dialog
            vbox = QVBoxLayout();newDialog.setWindowTitle(self.laguage.Forgot_to_Save), newDialog.setWindowIcon(QIcon(ICON_PTW_LOGO))  # layout
            qlabel = QLabel(); qlabel.setFont(self.font), qlabel.setText(self.laguage.Forgot_to_Save), qlabel.setAlignment(Qt.AlignCenter),qlabel.setFixedWidth(DIALOG_WIDTH_CALENDER + 10), vbox.addWidget(qlabel)  # Set a label
            hbox = QHBoxLayout();insertBtt = QPushButton(self.laguage.save);insertBtt.setFont(self.font), insertBtt.setFixedWidth(DIALOG_WIDTH_CALENDER // 2), insertBtt.clicked.connect(self.SaveAll), insertBtt.setPalette(QPalette(Qt.blue)), insertBtt.clicked.connect(newDialog.close), hbox.addWidget(insertBtt)  # Button to save
            closeBtt = QPushButton(self.laguage.cancelBtt);closeBtt.setFont(self.font), closeBtt.setFixedWidth(DIALOG_WIDTH_CALENDER // 2), closeBtt.setPalette(QPalette(Qt.red)), closeBtt.clicked.connect(newDialog.close), hbox.addWidget(closeBtt), vbox.addLayout(hbox)  # Close Button
            vbox.setContentsMargins(10, 10, 10, 10), newDialog.setLayout(vbox), newDialog.exec_()
        # here we clean all files inside the folder undo
        f = []
        for (dirpath, dirnames, filenames) in os.walk(os.getcwd() + os.sep + 'undo' + os.sep):
            f.extend(filenames)
            break
        for file in f:
            os.remove(os.getcwd() + os.sep + 'undo' + os.sep + file)


if __name__ == '__main__':
    def copyfile_2_dir(source, dest):
        # we make sure that we have the txts required
        string = ''
        with open(source, 'r') as src, open(dest, 'w') as dst:
            for lines in src: string += lines + '\n'
            dst.write(string)

    # Creating the folder if required
    folderhome = str(Path.home()) + os.sep + DESKTOP + os.sep + 'VDE-Sicherheitsprüfung'
    try: os.mkdir(folderhome)
    except: None
    try:
        os.mkdir(folderhome + os.sep + 'email-Setup')
        copyfile_2_dir(BACKUPFILE_EMAIL_CONTACTS_ROOM_RESPONSABLES, CONTACTS_EMAIL)
        copyfile_2_dir(BACKUPFILE_MESSAGE_EMAIL_ENGLISH, MESSAGE_EMAIL_ENGLISH)
        copyfile_2_dir(BACKUPFILE_MESSAGE_EMAIL_GERMAN, MESSAGE_EMAIL_GERMAN)
    except: None
    try:
        os.mkdir(folderhome + os.sep + 'Responsable_email-Setup')
        copyfile_2_dir(BACKUPFILE_EMAIL_CONTACTS_ROOM_RESPONSABLES, EMAIL_CONTACTS_ROOM_RESPONSABLES)
        copyfile_2_dir(BACKUPFILE_MESSAGE_EMAIL_ROOM_RESPONSABLES, MESSAGE_EMAIL_ROOM_RESPONSABLES)
    except: None
    try:
        os.mkdir(folderhome + os.sep + 'Datenbank')
        try: os.mkdir(folderhome + os.sep + 'Datenbank'+ os.sep +'EPRUEFUNG Backups')
        except: None
        try:os.mkdir(folderhome + os.sep + 'Datenbank'+ os.sep +'IZYTRON.IQ Export')
        except: None
    except: None
    # Starting the Software normally
    app = QApplication(sys.argv)
    myStyle = MyProxyStyle('Fusion')
    app.setStyle(myStyle)
    window= MainWindow()
    sys.exit(app.exec_() )