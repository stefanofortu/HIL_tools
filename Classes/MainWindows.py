from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon, QAction
from PySide6.QtWidgets import QMainWindow, QWidget, QFrame, QTabWidget, QToolBar, QStatusBar

from Classes.Configuration_Data import HIL_Function_Configuration_Data, TC_Highlight_Configuration_Data
from Classes.Configuration_File import Configuration_File
from Classes.HIL_Function_Widget import HIL_Function_Widget
from Classes.TC_Highlight_Widget import TC_Highlight_Widget
from Classes.TC_Substitution_Handler import TC_Substitution_Configuration_Data
from Classes.TC_Substitution_Widget import TC_Substitution_Widget
from icons.resources import resource_path


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ferrari TC tools")
        self.left = 100
        self.top = 100
        self.width = 320
        self.height = 240
        self.setGeometry(self.left, self.top, self.width, self.height)
        # self.setStyleSheet("background-color: rgb(220,220,220)")
        self.configuration_file = Configuration_File()

        toolbar_action_new = QAction(QIcon(resource_path("new_configuration.png")), "New", self)
        toolbar_action_new.setStatusTip("Create new configuration")
        toolbar_action_new.triggered.connect(self.configuration_file.new)

        toolbar_action_open = QAction(QIcon(resource_path("open.jpg")), "Open", self)
        toolbar_action_open.setStatusTip("Open existing configuration")
        toolbar_action_open.triggered.connect(self.open_configuration_file)

        toolbar_action_save = QAction(QIcon(resource_path("save.ico")), "Save", self)
        toolbar_action_save.setStatusTip("Save current configuration")
        toolbar_action_save.triggered.connect(self.save_configuration_file)

        toolbar_action_save_as = QAction(QIcon(resource_path("save_as.jpeg")), "Save as", self)
        toolbar_action_save_as.setStatusTip("Save new configuration")
        toolbar_action_save_as.triggered.connect(self.save_configuration_file_as)

        # ####################
        toolbar = QToolBar("My main toolbar")
        toolbar.setIconSize(QSize(24, 24))
        toolbar.addAction(toolbar_action_new)
        toolbar.addAction(toolbar_action_open)
        toolbar.addAction(toolbar_action_save)
        toolbar.addAction(toolbar_action_save_as)
        self.addToolBar(toolbar)
        #
        menu = self.menuBar()
        #
        file_menu = menu.addMenu("File")
        file_menu.addAction(toolbar_action_new)
        file_menu.addAction(toolbar_action_open)
        file_menu.addAction(toolbar_action_save)
        file_menu.addAction(toolbar_action_save_as)

        self.setStatusBar(QStatusBar(self))

        ######################
        # try:
        #     with open('pathFile.json', 'r') as json_file:
        #         json_file_no_comment = ''.join(line for line in json_file if not line.startswith('#'))
        #         data_dict = json.loads(json_file_no_comment)
        #
        #
        # except FileNotFoundError:
        #     print('File pathFile.json does not exist')
        #     sys.exit()

        # print(json_data)

        # definisci il widget principale
        main_widget = QWidget()

        # self.hil_function_widget = HIL_Function_Widget(HIL_Function_Configuration_Data())
        # self.tc_highlight_widget = TC_Highlight_Widget(TC_Highlight_Configuration_Data())
        self.tc_substitution_widget = TC_Substitution_Widget(TC_Substitution_Configuration_Data())

        self.setCentralWidget(main_widget)

        main_widget = QTabWidget()
        main_widget.setDocumentMode(True)
        main_widget.setTabPosition(QTabWidget.North)
        main_widget.setMovable(False)

        # main_widget.insertTab(2, self.hil_function_widget, "HIL function")
        # main_widget.insertTab(1, self.tc_highlight_widget, "Test Case Highlight")
        main_widget.insertTab(0, self.tc_substitution_widget, "Test Case Substitutions")

        self.setCentralWidget(main_widget)

    def open_configuration_file(self):
        try:
            hil_function_file_data, tc_highlight_data, tc_substitution_data = self.configuration_file.open()
            self.tc_substitution_widget.update_handler(tc_substitution_data)
        except ValueError:
            self.statusBar().showMessage("No file selected", 2500)

    def save_configuration_file(self):
        hil_function_file_data = HIL_Function_Configuration_Data()
        tc_highlight_data = TC_Highlight_Configuration_Data()
        tc_substitution_data = self.tc_substitution_widget.tc_substitution_handler.cfg_data
        self.configuration_file.save(hil_function_file_data=hil_function_file_data.return_json_dict(),
                                     tc_highlight_data=tc_highlight_data.return_json_dict(),
                                     tc_substitution_data=tc_substitution_data.return_json_dict(),
                                     select_new_file=False)

    def save_configuration_file_as(self):
        hil_function_file_data = HIL_Function_Configuration_Data()
        tc_highlight_data = TC_Highlight_Configuration_Data()
        tc_substitution_data = self.tc_substitution_widget.tc_substitution_handler.cfg_data
        self.configuration_file.save(hil_function_file_data=hil_function_file_data.return_json_dict(),
                                     tc_highlight_data=tc_highlight_data.return_json_dict(),
                                     tc_substitution_data=tc_substitution_data.return_json_dict(),
                                     select_new_file=True)
