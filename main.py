import numpy as np
import pandas as pd
import os
from PyQt5.QtCore import *
from PyQt5 import QtWidgets, QtGui, QtCore
import sys
from PyQt5.QtWidgets import QLabel

from PyQt5.QtWidgets import QHBoxLayout


class TableView(QtWidgets.QTableWidget):
    def __init__(self, data={"1":[0]}, *args):
        QtWidgets.QTableWidget.__init__(self, *args)
        self.data = data
        self.setData()
        self.resizeColumnsToContents()
        self.resizeRowsToContents()

    def setData(self):
        horHeaders = []
        for n, key in enumerate(sorted(self.data.keys())):
            horHeaders.append(key)
            for m, item in enumerate(self.data[key]):
                newitem = QtWidgets.QTableWidgetItem(item)
                self.setItem(m, n, newitem)
        self.setHorizontalHeaderLabels(horHeaders)

class ImportDialog(QtWidgets.QDialog):
    def __init__(self,mainwind):
        super().__init__(mainwind)

        self.setWindowTitle("Import")

        layout = QtWidgets.QVBoxLayout()
        importvariables = ["Rechnungsdaten aus EEG Faktura","Daten über Mandate","Vorlage Rechnungen","Vorlage Infinity Export"]
        buttons = [0]*len(importvariables)
        okbutton = QtWidgets.QPushButton("OK")
        okbutton.pressed.connect(self.accept)
        layouts = [0]*len(importvariables)
        for i,importvariable in enumerate(importvariables):
            layouts[i] = QtWidgets.QHBoxLayout()
            layouts[i].addWidget(QtWidgets.QLabel(importvariable))
            buttons[i] = QtWidgets.QPushButton("Laden")
            layouts[i].addWidget(buttons[i])
            layout.addLayout(layouts[i])
        layout.addWidget(okbutton)
        self.setLayout(layout)

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        print("Initializing Window")
        self.setWindowTitle("Faktura Infinity Addon")

        self.init_Ui_file_not_loaded()

    def init_Ui_file_not_loaded(self):
        self.centralwidget = QtWidgets.QWidget(self)
        # main layout setup
        self.overallverticallayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.verticalLayout0 = QtWidgets.QVBoxLayout()  # layout on the left with the masslist, and other stuff
        self.verticalLayout1 = QtWidgets.QVBoxLayout()  # laout on the right with the graph
        self.table_0_0 = TableView()
        self.table_0_1 = TableView()
        self.creditor_ID_layout = QtWidgets.QHBoxLayout()
        self.creditor_ID_label = QtWidgets.QLabel()
        self.table_1_0 = TableView()

        self.horizontalLayout.addLayout(self.verticalLayout0)
        self.horizontalLayout.addLayout(self.verticalLayout1)
        self.verticalLayout0.addWidget(QLabel("Rechnungsdaten KonsumentInnen"))
        self.verticalLayout0.addWidget(self.table_0_0)
        self.verticalLayout0.addWidget(QLabel("Mandatsdaten KonsumentInnen"))
        self.verticalLayout0.addWidget(self.table_0_1)
        self.verticalLayout0.addLayout(self.creditor_ID_layout)
        self.creditor_ID_layout.addWidget(QtWidgets.QLabel("Creditor ID:"))
        self.creditor_ID_layout.addWidget(self.creditor_ID_label)
        self.verticalLayout0.setStretch(1, 7)
        self.verticalLayout0.setStretch(3, 7)


        # plot widget for the verticalLayout1
        self.verticalLayout1.addWidget(QLabel("Rechnungsdaten ProduzentInnen"))
        self.verticalLayout1.addWidget(self.table_1_0)

        menubar = QtWidgets.QMenuBar()
        self.actionFile = menubar.addMenu("Datei")
        # the po.importanythingact triggers init_UI_file_loaded() and init_plots()
        importanythingact = QtWidgets.QAction("Importieren", self)
        importanythingact.setShortcut("Ctrl+I")
        importanythingact.triggered.connect(self.importanything)
        makeinvact = QtWidgets.QAction("Rechnungen erstellen", self)
        makeinvact.triggered.connect(self.makeinvoice)
        makeinfexpact = QtWidgets.QAction("Für Infinity vorbereiten", self)
        makeinfexpact.triggered.connect(self.makeinfexport)
        mailingact = QtWidgets.QAction("Emails Senden", self)
        mailingact.triggered.connect(self.mailingselect)
        self.actionFile.addAction(importanythingact)
        self.actionFile.addAction(makeinvact)
        self.actionFile.addAction(makeinfexpact)
        self.actionFile.addAction(mailingact)


        self.actionFile.addSeparator()
        quit = QtWidgets.QAction("Schließen", self)
        quit.setShortcut("Alt+F4")
        quit.triggered.connect(lambda: sys.exit(0))
        self.actionFile.addAction(quit)

        self.overallverticallayout.addWidget(menubar)
        self.overallverticallayout.addLayout(self.horizontalLayout)
        self.setCentralWidget(self.centralwidget)



    def importanything(self):
        print("Import Action")
        dlg = ImportDialog(self)
        dlg.exec()

    def makeinvoice(self):
        print("Invoice Action")
        # auswählen wohin und welches format

    def makeinfexport(self):
        print("Inf exp action")
        # zeige für welche kundInnen kein sepa mandat vorhanden

    def mailingselect(self):
        print("Send Mails")
        # auswählen an wen (woher bekommen wir die mail daten?)

        # dialog = QtWidgets.QFileDialog()
        # filepath, filter = dialog.getimportanythingactName(None, "Window name", "", "HDF5_files (*.hdf5)")
        # self.filename = filepath
        # # if self.file_loaded:
        # #     print("remove old plot stuff")
        # #     pyqtgraph_objects.remove_all_plot_items(parent)
        # self.init_basket_objects()
        # self.init_UI_file_loaded()
        # self.init_plots()
        # self.file_loaded = True

    def init_basket_objects(self):
        # those are the "basket" objects, where the data is in sp = all data that has to do with the spectrum, ml = all data to the masslist

        self.plot_settings = {"vert_lines_color_suggestions": (97, 99, 102, 70),
                              "vert_lines_color_masslist": (38, 135, 20),
                              "vert_lines_color_masslist_without_composition": (13, 110, 184),
                              "vert_lines_color_isotopes": (252, 3, 244, 70),
                              # RGB tubel and last number gives the transparency (from 0 to 255)
                              "vert_lines_width_suggestions": 1,
                              "vert_lines_width_masslist": 2,
                              "vert_lines_width_isotopes": 1.5,
                              "average_spectrum_color": (252, 49, 3),
                              "max_spectrum_color": (122, 72, 6, 80),
                              "min_spectrum_color": (11, 125, 191, 80),
                              "sub_spectrum_color": (103, 42, 201, 80),
                              "color_cycle": ['r', 'g', 'b', 'c', 'm', 'y'],
                              "current_color": 0,
                              "current_color_fixed": 0,
                              "background_color": "w",
                              "show_plots": [True, False, False, False],
                              # plots corresponding to [avg spectrum, min spec, max spect, subspectr]
                              "avg": False,
                              "raw": True
                              }


    def init_UI_file_loaded(self):
        pass

    def init_plots(self):
        pass

def main():
    app = QtWidgets.QApplication(sys.argv)
    main = MainWindow()
    main.show()
    sys._excepthook = sys.excepthook

    def exception_hook(exctype, value, traceback):
        print("silent error")
        print(exctype, value, traceback)
        sys._excepthook(exctype, value, traceback)
        sys.exit(1)

    sys.excepthook = exception_hook
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()