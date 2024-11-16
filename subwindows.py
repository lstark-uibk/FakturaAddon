import PyQt5.QtWidgets as QtWidgets
from PyQt5.QtWidgets import QMainWindow
import sys


class Subwindow(QMainWindow):
    def __init__(self, Windowname,  Menubardata,*args,**kwargs):
        # menubar has to be [[Text,Shortcut,function]]
        super(Subwindow, self).__init__(*args, **kwargs)
        print("Initializing Subwindow")
        self.windowname = Windowname
        self.menubardata = Menubardata
        self.setWindowTitle(self.windowname)

        self.init_Ui_overview()

    def init_Ui_overview(self):
        self.centralwidget = QtWidgets.QWidget(self)
        # main layout setup
        self.overallverticallayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.verticalLayout0 = QtWidgets.QVBoxLayout()  # layout on the left with the masslist, and other stuff
        self.verticalLayout1 = QtWidgets.QVBoxLayout()  # laout on the right with the graph
        menubar = QtWidgets.QMenuBar()
        self.actionFile = menubar.addMenu("Datei")
        # the po.importanythingact triggers init_UI_file_loaded() and init_plots()
        for menuline in self.menubardata:
            action = QtWidgets.QAction(menuline[0], self)
            action.triggered.connect(menuline[2])
            if menuline[1]:
                action.setShortcut(menuline[1])
            self.actionFile.addAction(action)

        self.actionFile.addSeparator()
        quit = QtWidgets.QAction("Schließen", self)
        quit.setShortcut("Alt+F4")
        quit.triggered.connect(lambda: sys.exit(0))
        self.actionFile.addAction(quit)

        self.overallverticallayout.addWidget(menubar)
        self.overallverticallayout.addLayout(self.horizontalLayout)
        self.setCentralWidget(self.centralwidget)


# class Selectfromnextcloudwindow():
    #https: // pythonspot.com / pyqt5 - directory - view /




#
# class MainWindow(QtWidgets.QMainWindow):
#     def __init__(self, *args, **kwargs):
#         super(MainWindow, self).__init__(*args, **kwargs)
#         print("Initializing Window")
#         self.setWindowTitle("Faktura Infinity Addon")
#
#         self.init_Ui_file_not_loaded()
#
#     def init_Ui_file_not_loaded(self):
#         self.centralwidget = QtWidgets.QWidget(self)
#         # main layout setup
#         self.overallverticallayout = QtWidgets.QVBoxLayout(self.centralwidget)
#         self.horizontalLayout = QtWidgets.QHBoxLayout()
#         self.verticalLayout0 = QtWidgets.QVBoxLayout()  # layout on the left with the masslist, and other stuff
#         self.verticalLayout1 = QtWidgets.QVBoxLayout()  # laout on the right with the graph
#         self.table_0_0 = TableView()
#         self.table_0_1 = TableView()
#         self.creditor_ID_layout = QtWidgets.QHBoxLayout()
#         self.creditor_ID_label = QtWidgets.QLabel()
#         self.table_1_0 = TableView()
#
#         self.horizontalLayout.addLayout(self.verticalLayout0)
#         self.horizontalLayout.addLayout(self.verticalLayout1)
#         self.verticalLayout0.addWidget(QLabel("Rechnungsdaten KonsumentInnen"))
#         self.verticalLayout0.addWidget(self.table_0_0)
#         self.verticalLayout0.addWidget(QLabel("Mandatsdaten KonsumentInnen"))
#         self.verticalLayout0.addWidget(self.table_0_1)
#         self.verticalLayout0.addLayout(self.creditor_ID_layout)
#         self.creditor_ID_layout.addWidget(QtWidgets.QLabel("Creditor ID:"))
#         self.creditor_ID_layout.addWidget(self.creditor_ID_label)
#         self.verticalLayout0.setStretch(1, 7)
#         self.verticalLayout0.setStretch(3, 7)
#
#
#         # plot widget for the verticalLayout1
#         self.verticalLayout1.addWidget(QLabel("Rechnungsdaten ProduzentInnen"))
#         self.verticalLayout1.addWidget(self.table_1_0)
#
#         menubar = QtWidgets.QMenuBar()
#         self.actionFile = menubar.addMenu("Datei")
#         # the po.importanythingact triggers init_UI_file_loaded() and init_plots()
#         importanythingact = QtWidgets.QAction("Importieren", self)
#         importanythingact.setShortcut("Ctrl+I")
#         importanythingact.triggered.connect(self.importanything)
#         makeinvact = QtWidgets.QAction("Rechnungen erstellen", self)
#         makeinvact.triggered.connect(self.makeinvoice)
#         makeinfexpact = QtWidgets.QAction("Für Infinity vorbereiten", self)
#         makeinfexpact.triggered.connect(self.makeinfexport)
#         mailingact = QtWidgets.QAction("Emails Senden", self)
#         mailingact.triggered.connect(self.mailingselect)
#         self.actionFile.addAction(importanythingact)
#         self.actionFile.addAction(makeinvact)
#         self.actionFile.addAction(makeinfexpact)
#         self.actionFile.addAction(mailingact)
#
#
#         self.actionFile.addSeparator()
#         quit = QtWidgets.QAction("Schließen", self)
#         quit.setShortcut("Alt+F4")
#         quit.triggered.connect(lambda: sys.exit(0))
#         self.actionFile.addAction(quit)
#
#         self.overallverticallayout.addWidget(menubar)
#         self.overallverticallayout.addLayout(self.horizontalLayout)
#         self.setCentralWidget(self.centralwidget)
#
#
#
#     def importanything(self):
#         print("Import Action")
#         dlg = ImportDialog(self)
#         dlg.exec()
#
#     def makeinvoice(self):
#         print("Invoice Action")
#         # auswählen wohin und welches format
#
#     def makeinfexport(self):
#         print("Inf exp action")
#         # zeige für welche kundInnen kein sepa mandat vorhanden
#
#     def mailingselect(self):
#         print("Send Mails")
#         # auswählen an wen (woher bekommen wir die mail daten?)
#
#         # dialog = QtWidgets.QFileDialog()
#         # filepath, filter = dialog.getimportanythingactName(None, "Window name", "", "HDF5_files (*.hdf5)")
#         # self.filename = filepath
#         # # if self.file_loaded:
#         # #     print("remove old plot stuff")
#         # #     pyqtgraph_objects.remove_all_plot_items(parent)
#         # self.init_basket_objects()
#         # self.init_UI_file_loaded()
#         # self.init_plots()
#         # self.file_loaded = True
#
#     def init_basket_objects(self):
#         # those are the "basket" objects, where the data is in sp = all data that has to do with the spectrum, ml = all data to the masslist
#
#         self.plot_settings = {"vert_lines_color_suggestions": (97, 99, 102, 70),
#                               "vert_lines_color_masslist": (38, 135, 20),
#                               "vert_lines_color_masslist_without_composition": (13, 110, 184),
#                               "vert_lines_color_isotopes": (252, 3, 244, 70),
#                               # RGB tubel and last number gives the transparency (from 0 to 255)
#                               "vert_lines_width_suggestions": 1,
#                               "vert_lines_width_masslist": 2,
#                               "vert_lines_width_isotopes": 1.5,
#                               "average_spectrum_color": (252, 49, 3),
#                               "max_spectrum_color": (122, 72, 6, 80),
#                               "min_spectrum_color": (11, 125, 191, 80),
#                               "sub_spectrum_color": (103, 42, 201, 80),
#                               "color_cycle": ['r', 'g', 'b', 'c', 'm', 'y'],
#                               "current_color": 0,
#                               "current_color_fixed": 0,
#                               "background_color": "w",
#                               "show_plots": [True, False, False, False],
#                               # plots corresponding to [avg spectrum, min spec, max spect, subspectr]
#                               "avg": False,
#                               "raw": True
#                               }
#
#
#     def init_UI_file_loaded(self):
#         pass
#
#     def init_plots(self):
#         pass
