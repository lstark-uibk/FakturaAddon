from time import sleep
import numpy as np
import pandas as pd
import os
import subwindows as sw
from PyQt5.QtCore import *
from PyQt5 import QtWidgets, QtGui, QtCore
import sys
from PyQt5.QtWidgets import QLabel, QFileDialog, QMessageBox,QTableWidget,QTableWidgetItem
from importing import mandates,invoices,emails
from PyQt5.QtWidgets import QHBoxLayout


class TableView(QtWidgets.QTableWidget):
    def __init__(self, data={"1":[0]}, *args):
        QtWidgets.QTableWidget.__init__(self, *args)
        self.data = data
        self.setData()
        self.resizeColumnsToContents()
        self.resizeRowsToContents()
    def set_new_data(self,data):
        """
        sets the Table to new data
        :param data: pd.Dataframe
        :return:
        """
        self.data = data.to_dict(orient="list")
        self.setData(data.shape[0],data.shape[1])
    def setData(self,rowcount = 0, colcount = 0):
        print(colcount,rowcount)
        self.setColumnCount(colcount)
        self.setRowCount(rowcount)
        horHeaders = []
        for n, key in enumerate(sorted(self.data.keys())):
            horHeaders.append(key)
            for m, item in enumerate(self.data[key]):
                newitem = QtWidgets.QTableWidgetItem(str(item))
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
        def sel_filepath_and_import():
            dialog = QFileDialog()
            foo_dir = dialog.getExistingDirectory(self, 'Select an awesome directory')
        buttons[0].pressed.connect()

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        print("Initializing Window")
        self.setWindowTitle("Faktura Infinity Addon")
        self.second_window = None
        self.home_directory = "/home/leander/gei"

        self.init_Ui_overview()
        self.init_data()

    def init_Ui_overview(self):
        self.centralwidget = QtWidgets.QWidget(self)
        # main layout setup
        self.overallverticallayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.setCentralWidget(self.centralwidget)
        button1 = QtWidgets.QPushButton("SEPA Export erstellen")
        button2= QtWidgets.QPushButton("Rechnung erstellen (noch nicht implementiert)")
        button3 = QtWidgets.QPushButton("Mitglied anmelden (noch nicht implementiert)")
        button1.pressed.connect(self.init_SEPA_export_UI)
        self.init_SEPA_export_UI() ###########################
        self.overallverticallayout.addWidget(button1)
        self.overallverticallayout.addWidget(button2)
        self.overallverticallayout.addWidget(button3)

    def init_data(self):
        self.mandates = mandates
        self.invoices = invoices
        self.emails = emails
    def init_SEPA_export_UI(self):
        if self.second_window is None:
            def load_filepath(title,filter= "Excel (*.xlsx)"):
                dlg = QMessageBox(self)
                dlg.setWindowTitle("Importiern")
                dlg.setText("Von wo willst du importierten?")
                local = dlg.addButton('Lokal', QMessageBox.YesRole)
                nextcloud = dlg.addButton('Nextcloud', QMessageBox.NoRole)
                button = dlg.exec()

                if dlg.clickedButton() == local:
                    print("Lokal")
                    filepath,filter = QFileDialog.getOpenFileName(self, title, "/home/leander/gei", filter)
                elif dlg.clickedButton() == nextcloud:
                    print("Nextcloud")
                return filepath

            def import_mandates():
                print("Import mandates")
                filepath = load_filepath("Importiere SEPA Mandate")
                mandatedata = self.mandates.load_data(filepath)
                self.second_window.reload_table_view("0_1",mandatedata)

            def import_invoice_data():
                print("import invoice data")
                filepath = load_filepath("Importiere Rechnungen von EEG Faktura")
                invoicedata = self.invoices.load_data(filepath)
                self.second_window.reload_table_view("0_0",invoicedata["debit"])
                self.second_window.reload_table_view("1_0",invoicedata["transfer"])


            def select_templates():
                print("Select templates")
                filepath1 = load_filepath("Wähle Exportvorlage für SEPA Lastschrift aus",filter =  "csv (*.csv)")
                filepath2 = load_filepath("Wähle Exportvorlage für SEPA Lastschrift aus",filter =  "csv (*.csv)")
                template = self.mandates.load_template(filepath1,filepath2)
                self.second_window.creditor_ID_label.setText("Ja")
            def export_csv():
                print("Export cvs")

            menubardata = [["Importiere Mandate","Strg+I",import_mandates],["Importiere Rechnungsdaten von EEG Faktura","",import_invoice_data],["Vorlage für SEPA Export auswählen","",select_templates],["Exportiere .csv Datei für Raiffeisen Infinty","",export_csv]]
            self.second_window = sw.Subwindow("SEPA Export", menubardata)
            self.second_window.horizontalLayout = QtWidgets.QHBoxLayout()
            self.second_window.verticalLayout0 = QtWidgets.QVBoxLayout()  # layout on the left with the masslist, and other stuff
            self.second_window.verticalLayout1 = QtWidgets.QVBoxLayout()  # laout on the right with the graph
            self.second_window.table_0_0 = TableView()
            self.second_window.table_0_1 = TableView()
            self.second_window.creditor_ID_layout = QtWidgets.QHBoxLayout()
            self.second_window.creditor_ID_label = QtWidgets.QLabel("Nein")
            self.second_window.table_1_0 = TableView()

            def reload_table_view(tablenr,data):
                """

                :param tablenr: in columns on the grid "0_0","0_1","1_0"
                :param data: as a Pandas Dataframe
                :return:
                """
                tabledict = {"0_0":self.second_window.table_0_0,
                             "0_1": self.second_window.table_0_1,
                             "1_0": self.second_window.table_1_0,
                             }
                tabledict[tablenr].set_new_data(data)
            self.second_window.reload_table_view = reload_table_view

            self.second_window.horizontalLayout.addLayout(self.second_window.verticalLayout0)
            self.second_window.horizontalLayout.addLayout(self.second_window.verticalLayout1)
            self.second_window.verticalLayout0.addWidget(QLabel("Rechnungsdaten KonsumentInnen"))
            self.second_window.verticalLayout0.addWidget(self.second_window.table_0_0)
            self.second_window.verticalLayout0.addWidget(QLabel("Mandatsdaten KonsumentInnen"))
            self.second_window.verticalLayout0.addWidget(self.second_window.table_0_1)
            self.second_window.verticalLayout0.addLayout(self.second_window.creditor_ID_layout)
            self.second_window.creditor_ID_layout.addWidget(QtWidgets.QLabel("Exportvolage geladen? "))
            self.second_window.creditor_ID_layout.addWidget(self.second_window.creditor_ID_label)
            self.second_window.verticalLayout0.setStretch(1, 7)
            self.second_window.verticalLayout0.setStretch(3, 7)


            # plot widget for the verticalLayout1
            self.second_window.verticalLayout1.addWidget(QLabel("Rechnungsdaten ProduzentInnen"))
            self.second_window.verticalLayout1.addWidget(self.second_window.table_1_0)
            self.second_window.overallverticallayout.addLayout(self.second_window.horizontalLayout)


            self.second_window.show()

        else:
            self.second_window.close()  # Close window.
            self.second_window = None  # Discard reference.



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