from time import sleep
from nc_py_api import Nextcloud
import pandas as pd
import os
from subwindows import LoginPrompt, Subwindow
from PyQt5.QtCore import *
from PyQt5 import QtWidgets, QtGui, QtCore
import sys
from PyQt5.QtWidgets import QLabel, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QListWidget, QWidget, QListWidgetItem, QCheckBox, QListWidgetItem, QPushButton, QVBoxLayout
from importing import mandates,invoices,emails
from exporting import produce_sepa_export_dfs
from PyQt5.QtWidgets import QHBoxLayout
import datetime as dt



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
        self.setColumnCount(colcount)
        self.setRowCount(rowcount)
        horHeaders = []
        for n, key in enumerate(self.data.keys()):
            horHeaders.append(key)
            for m, item in enumerate(self.data[key]):
                newitem = QtWidgets.QTableWidgetItem(str(item))
                self.setItem(m, n, newitem)
        self.setHorizontalHeaderLabels(horHeaders)


# class ImportDialog(QtWidgets.QDialog):
#     def __init__(self,mainwind):
#         super().__init__(mainwind)
#
#         self.setWindowTitle("Import")
#
#         layout = QtWidgets.QVBoxLayout()
#         importvariables = ["Rechnungsdaten aus EEG Faktura","Daten über Mandate","Vorlage Rechnungen","Vorlage Infinity Export"]
#         buttons = [0]*len(importvariables)
#         okbutton = QtWidgets.QPushButton("OK")
#         okbutton.pressed.connect(self.accept)
#         layouts = [0]*len(importvariables)
#         for i,importvariable in enumerate(importvariables):
#             layouts[i] = QtWidgets.QHBoxLayout()
#             layouts[i].addWidget(QtWidgets.QLabel(importvariable))
#             buttons[i] = QtWidgets.QPushButton("Laden")
#             layouts[i].addWidget(buttons[i])
#             layout.addLayout(layouts[i])
#         layout.addWidget(okbutton)
#         self.setLayout(layout)
#         def sel_filepath_and_import():
#             dialog = QFileDialog()
#             foo_dir = dialog.getExistingDirectory(self, 'Select an awesome directory')
#         buttons[0].pressed.connect()

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        print("Initializing Window")
        self.setWindowTitle("Faktura Infinity Addon")
        self.resize(1000, 600)
        self.move(20, 20)
        self.second_window = None
        self.exportwindow = None
        self.home_directory = "/home/leander/gei"
        self.loaded_filepaths = pd.DataFrame({"Daten":["Mandate","Rechnungsdaten","Rechnungsdaten Vorlagen"],
                                  "Speicherort":["","",""],

                                              })
        self.loginprompt = None
        self.nc_auth_user = ''
        self.nc_auth_pass = ''
        self.creditor_ID = "AT94ZZZ00000079821"
        self.nc_mandatefilepath = "Gemeinwohlenergie/Rechnungswesen, IT/Abrechnung Faktura/SEPA Lastschriftmandate/lastschriftmandate.xlsx"
        self.mandatesdata_loaded = False
        self.invoicesdata_loaded = False
        self.init_Ui()
        self.init_data()

    def init_Ui(self):
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget = QtWidgets.QWidget(self)
        self.overallverticallayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.menubardata = self.init_menubardata_mandates()
        if self.menubardata:
            menubar = QtWidgets.QMenuBar()
            self.actionFile = menubar.addMenu("Infinity export")
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

        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.verticalLayout0 = QtWidgets.QVBoxLayout() 
        self.verticalLayout1 = QtWidgets.QVBoxLayout()
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.verticalLayout0 = QtWidgets.QVBoxLayout()  # layout on the left with the masslist, and other stuff
        self.verticalLayout1 = QtWidgets.QVBoxLayout()  # laout on the right with the graph
        self.table_0_0 = TableView()
        self.table_0_1 = TableView()
        self.table_1_0 = TableView()
        self.table_1_1 = TableView()
        self.table_1_1.set_new_data(self.loaded_filepaths)
        self.horizontalLayout.addLayout(self.verticalLayout1)
        self.horizontalLayout.addLayout(self.verticalLayout0)
        self.verticalLayout0.addWidget(QLabel("Rechnungsdaten KonsumentInnen"))
        self.verticalLayout0.addWidget(self.table_0_0)
        self.verticalLayout0.addWidget(QLabel("Mandatsdaten KonsumentInnen"))
        self.verticalLayout0.addWidget(self.table_0_1)
        self.verticalLayout0.setStretch(1, 7)
        self.verticalLayout0.setStretch(3, 7)

        self.verticalLayout1.addWidget(QLabel("Rechnungsdaten ProduzentInnen"))
        self.verticalLayout1.addWidget(self.table_1_0)
        self.verticalLayout1.addWidget(QLabel("Dateien geladen:"))
        self.verticalLayout1.addWidget(self.table_1_1)
        self.verticalLayout1.setStretch(1, 7)
        self.verticalLayout1.setStretch(3, 4)

        self.overallverticallayout.addLayout(self.horizontalLayout)



        self.overallverticallayout.addLayout(self.horizontalLayout)
        self.setCentralWidget(self.centralwidget)

    def init_data(self):
        self.mandates = mandates
        self.invoices = invoices
        self.emails = emails
    
    def init_menubardata_mandates(self):
        def load_filepath(title,filter= "Excel (*.xlsx)",fileex=True):
            dlg = QMessageBox(self)

                # dlg.setWindowTitle("Importiern")
                # dlg.setText("Von wo willst du importierten?")
                # local = dlg.addButton('Lokal', QMessageBox.YesRole)
                # nextcloud = dlg.addButton('Nextcloud', QMessageBox.NoRole)
                # button = dlg.exec()
                #
                # if dlg.clickedButton() == local:
            print("Lokal")
            if fileex:
                filepath,filter = QFileDialog.getOpenFileName(self, title, self.home_directory, filter)
            else:
                filepath, filter = QFileDialog.getSaveFileName(self, title,  self.home_directory, filter)
                # elif dlg.clickedButton() == nextcloud:
                #     print("Nextcloud")
            if filepath:
                return filepath
            else: return None

        def updatetable_1_1():
            self.table_1_1.set_new_data(self.loaded_filepaths.iloc[0:3])

        def import_mandates():
            print("Import mandates")
            dlg = QMessageBox(self)
            questiontext = f"Ich kann die Mandate von folgendem Pfad in nextcloud herunterladen:"
            questiontext += f"\n\n{self.nc_mandatefilepath}"
            questiontext += "\n\nSoll ich es von diesem Pfad herunterladen, oder willst du lokal eine Datei von deinem Computer auswählen?"
            dlg.setText(questiontext)
            dlg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            prompt = dlg.exec()

            def load_mandate(filepath,nc_loading=False,nc_instance=""):
                if filepath is not None:
                    mandatedata = self.mandates.load_data(filepath,nc = nc_loading,nc_instance=nc_instance)
                    if mandatedata is not None:
                        self.reload_table_view("0_1",mandatedata)
                        self.loaded_filepaths.loc[self.loaded_filepaths["Daten"][self.loaded_filepaths["Daten"] == "Mandate"].index, "Speicherort"] = filepath
                        updatetable_1_1()
                        self.mandatesdata_loaded = True
                else:
                    return None


            if prompt == QMessageBox.Yes:
                nc_loading = True
                self.loginprompt = LoginPrompt(load_mandate,self.nc_mandatefilepath)
                self.loginprompt.show()
            else:
                nc_loading = False
                filepath = load_filepath("Lade Daten von SEPA Mandate")
                load_mandate(filepath)


        def import_invoice_data():
            print("import invoice data")
            filepath = load_filepath("Importiere Rechnungen von EEG Faktura")
            if filepath is not None:
                self.loaded_filepaths.loc[self.loaded_filepaths["Daten"][self.loaded_filepaths["Daten"] == "Rechnungsdaten"].index, "Speicherort"] = filepath


                invoicedata = self.invoices.load_data(filepath)
                if invoicedata is not None:
                    debit = invoicedata["list"][(invoicedata["list"]["Dokumenttyp"] == "Rechnung")]
                    transfer = invoicedata["list"][(invoicedata["list"]["Dokumenttyp"] == "Gutschrift")|(invoicedata["list"]["Dokumenttyp"] == "Information")]

                    self.reload_table_view("0_0",debit)
                    self.reload_table_view("1_0",transfer)

                    updatetable_1_1()
                    self.invoicesdata_loaded = True
                else:
                    print()

        def select_templates():
            print("Select templates")
            filepath1 = load_filepath("Wähle Exportvorlage für SEPA Lastschrift aus",filter =  "csv (*.csv)")
            if filepath1 is not None:
                filepath2 = load_filepath("Wähle Exportvorlage für SEPA Lastschrift aus",filter =  "csv (*.csv)")
                if filepath2 is not None:
                    self.loaded_filepaths.loc[self.loaded_filepaths["Daten"][self.loaded_filepaths["Daten"] == "Mandate Vorlagen"].index, "Speicherort1"] = filepath1
                    self.loaded_filepaths.loc[self.loaded_filepaths["Daten"][self.loaded_filepaths["Daten"] == "Mandate Vorlagen"].index, "Speicherort2"] = filepath2

                    template = self.mandates.load_template(filepath1,filepath2)
                    updatetable_1_1()

        def export_csv():
            print("Export cvs")
            if self.exportwindow is None:
                #data check
                if not self.invoicesdata_loaded:
                    errorbox = QMessageBox()
                    errorbox.setText("Es wurden keine Rechungsdaten ausgewählt, wähle zuerst diese aus und versuch es nochmal.")
                    errorbox.exec_()
                    return None
                if not self.mandatesdata_loaded:
                    errorbox = QMessageBox()
                    errorbox.setText("Es wurden keine Mandatssdaten ausgewählt, wähle zuerst diese aus und versuch es nochmal.")
                    errorbox.exec_()

                    return None

                self.exportwindow = Subwindow("Exportiere .csv für SEPA")
                self.exportwindow.resize(500, 500)
                self.exportwindow.move(30, 30)
                self.exportwindow.verticalLayout = QVBoxLayout()
                header_layout = QHBoxLayout()

                # Add header labels to the header layout
                header_label1 = QLabel("Name")
                header_label3 = QLabel("Betrag")
                header_layout.addWidget(header_label1)
                header_layout.addWidget(header_label3)

                # Adjust the header layout
                header_layout.addStretch(1)
                header_layout.setSpacing(20)


                self.exportwindow.list_widget = QListWidget()
                self.exportwindow.list_data = []

                names = []
                amounts = []
                for idx, person in self.invoices.data["list"].iterrows():
                    name = person["Empfänger Vorame"]
                    if not pd.isna(person["Empfänger Nachname"]):
                        name += f" {person['Empfänger Nachname']}"
                    names.append(name)
                    amounts.append(person["Rechnungsbetrag Brutto"])
                # mandatesexist = []
                # for name in names:
                #     if (mandates.data["Zahlungspflichtiger Name"] == name).any():
                #         mandatesexist.append("x")
                #     else: mandatesexist.append("")

                for name,amount in zip(names,amounts):
                    item = QListWidgetItem(self.exportwindow.list_widget )
                    item.setSizeHint(QSize(500, 30))

                    row_widget = QWidget()
                    row_layout = QHBoxLayout()

                    checkbox = QCheckBox()
                    checkbox.setChecked(True)  # Default: unchecked
                    row_layout.addWidget(checkbox)
                    self.exportwindow.list_data.append(checkbox)

                    col1 = QLabel(str(name))
                    col2 = QLabel(str(amount))

                    row_layout.addWidget(col1)
                    row_layout.addWidget(col2)


                    row_layout.setContentsMargins(0, 0,0,0)
                    # row_layout.setSpacing(15)
                    row_widget.setLayout(row_layout)

                    self.exportwindow.list_widget.setItemWidget(item, row_widget)

                def get_selected_names():
                    nr_list_widgets = len(self.exportwindow.list_data)
                    selected_names = [False] * nr_list_widgets
                    for index,checkbox in enumerate(self.exportwindow.list_data):
                        if checkbox.isChecked():
                            selected_names[index] = True

                    print(self.invoices.data["list"].loc[selected_names])
                    invoices_selected_names = self.invoices.data["list"].loc[selected_names]

                    exportingdebit,exportingtransfer, missingmandates = produce_sepa_export_dfs(invoices_selected_names,self.mandates,self.creditor_ID)
                    print(missingmandates)
                    if missingmandates:
                        dlg = QMessageBox(self)
                        questiontext = f"Für folgende Personen gibt es Daten zur Lastschrift, aber keine Daten zu einem Mandat:\n\n"
                        for name in missingmandates:
                            questiontext += f"{name} \n"
                        questiontext += "\nWillst du trotzdem fortfahren? \n(Es ist eigentlich kein Problem, wenn ein Mandat fehlt, da du in Infinity noch ein Mandat hinzufügen kannst. Jedoch ist es 'Good practice' dies im Mandatenfile zu machen."

                        dlg.setText(questiontext)
                        dlg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                        prompt = dlg.exec()
                        if prompt == QMessageBox.No:
                            print("Abort")
                            return

                    print(f"df = {exportingdebit,exportingtransfer}")
                    filepath1 = load_filepath("Wähle Speicherort für Export für SEPA Lastschrift aus", filter="csv (*.csv)", fileex=False)
                    if filepath1 is not None:
                        if ".csv" not in filepath1:
                            filepath1 = f"{filepath1}.csv"
                        print(f"Export to: {filepath1}")
                        try:
                            exportingdebit.to_csv(filepath1, index=False,sep=";")
                        except:
                            errorbox = QMessageBox("Saving didnot work")
                            print("savning didnot work")
                    else: return

                    filepath2 = load_filepath("Wähle Speicherort für Export für Überweisungen aus",
                                              filter="csv (*.csv)", fileex=False)

                    if filepath2 is not None:
                        if ".csv" not in filepath2:
                            filepath2 = f"{filepath2}.csv"
                        print(f"Export to: {filepath2}")
                        try:
                            exportingtransfer.to_csv(filepath2, index=False,sep=";")
                        except:
                            errorbox = QMessageBox("Saving didnot work")
                            print("savning didnot work")
                    else:
                        return
                    self.exportwindow.close()
                    return selected_names


                self.exportwindow.ok_button = QPushButton("OK")
                self.exportwindow.ok_button.pressed.connect(get_selected_names)
                self.exportwindow.overallverticallayout.addLayout(self.exportwindow.verticalLayout)

                self.exportwindow.verticalLayout.addLayout(header_layout)
                self.exportwindow.verticalLayout.addWidget(self.exportwindow.list_widget)
                self.exportwindow.verticalLayout.addWidget(self.exportwindow.ok_button)


                self.exportwindow.show()
            else:
                self.exportwindow.close()  # Close window.
                self.exportwindow = None  # Discard reference.

        def reload_table_view(tablenr,data):
            """

            :param tablenr: in columns on the grid "0_0","0_1","1_0"
            :param data: as a Pandas Dataframe
            :return:
            """
            tabledict = {"0_0":self.table_0_0,
                         "0_1": self.table_0_1,
                         "1_0": self.table_1_0,
                         }
            tabledict[tablenr].set_new_data(data)

        self.reload_table_view = reload_table_view

        menubardata = [["Importiere Rechnungdaten von EEG Faktura", "", import_invoice_data],
                            ["Lade Daten von SEPA Mandate", "", import_mandates],
                            ["Exportiere .csv Datei für Raiffeisen Infinty", "", export_csv]]
        return menubardata


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