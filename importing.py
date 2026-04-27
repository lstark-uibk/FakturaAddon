import time
from tempfile import template
import numpy as np
from functools import partial
from docxtpl import DocxTemplate
import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QWidget, QVBoxLayout, QPushButton, QLabel,  QMessageBox, QDialog
from PyQt5.QtCore import pyqtSignal, QThread, Qt
from io import BytesIO
from jinja2 import Environment, FileSystemLoader, PackageLoader, select_autoescape
import os
from numpy.testing.print_coercion_tables import print_new_cast_table
from openpyxl import load_workbook


# mandate =pd.read_excel("/home/leander/gei/faktura/abrechnung_24_q3/CC100438_abrechnung_Abr_YQ-2024-3_export(1).xlsx",sheet_name="Liste")
# df = pd.read_excel("/home/leander/gei/faktura/abrechnung_24_q3/CC100438_abrechnung_Abr_YQ-2024-3_export(1).xlsx",sheet_name="Liste")
class Data():
    def __init__(self,F_for_data_loading,F_for_template_loading, F_for_metadata_loading = lambda x: x):
        self.f_load_data = F_for_data_loading
        self.f_load_template = F_for_template_loading
        self.f_load_metadata = F_for_metadata_loading

        self.data = None
        self.metadata = None
        self.template = None
    def load_data(self,**kwargs):
        self.data = self.f_load_data(**kwargs)
        return self.data
    def load_metadata(self,**kwargs):
        self.metadata = self.f_load_metadata(**kwargs)
        return self.metadata
    def load_template(self,**kwargs):
        self.template = self.f_load_template(**kwargs)
        return self.template

# def load_mandates(filepath='',nc =False, nc_instance = ''):
#     print(f"Load {filepath}")
#     if not nc:
#         try:
#             mandates = pd.read_excel(filepath)
#         except:
#             errorbox = QMessageBox()
#             errorbox.setText("Ausgewählte Date ist nicht lesbar (ist sie im richtigen Format?)")
#             errorbox.exec_()
#             return
#     else:
#         print("Nextcloud loading")
#         mandates = nc_instance.files.download(filepath)
#         mandates = pd.read_excel(BytesIO(mandates), engine='openpyxl')
#         print(mandates)
#     try:
#         mandates = mandates[~mandates['Mandatsausstellungsdatum (Datum auf dem Vertrag)'].isna()]
#         mandates['Mandatsausstellungsdatum'] = pd.to_datetime(mandates['Mandatsausstellungsdatum (Datum auf dem Vertrag)'],dayfirst=True)
#         mandates = mandates.drop('Mandatsausstellungsdatum (Datum auf dem Vertrag)', axis=1)
#         mandates = mandates.rename(columns={'Vorname (gleich wie in eegfaktura)': 'Vorname', 'Nachname (gleich wie in eegfaktura)': 'Nachname','Mitgliedsnummer aus eegfaktura ist auch die Mandatsreferenz':'Mitgliedsnummer'})
#     except:
#         errorbox = QMessageBox()
#         errorbox.setText("Ausgewählte Date ist nicht lesbar (ist sie im richtigen Format?)")
#         errorbox.exec_()
#         return
#     return mandates
# def load_mandate_template(filepath = '', filepath2= ''):
#     print(f"Load {filepath},{filepath2}")
#     templates = {}
#     templates["debit"] = pd.read_csv(filepath,delimiter=";")
#     templates["transfer"]  = pd.read_csv(filepath2,delimiter=";")
#     return templates


def load_invoices(filepath='',nc = False):
    print(filepath)
    if not nc:
        try:
            data = pd.read_excel(filepath,sheet_name="Liste")
            datadetailed = pd.read_excel(filepath,sheet_name="Details")
        except:
            data = None
            datadetailed = None
            errorbox = QMessageBox()
            errorbox.setText("Ausgewählte Date ist nicht lesbar (ist sie im richtigen Format?)")
            errorbox.exec_()
    else:
        print("Nextcloud loading")
    if data is not None:
        invoicedata = {}
        invoicedata["list"] = data
        invoicedata["detailed"] = datadetailed

        return invoicedata
    else:
        return None

def load_invoice_template(filepath, nc = False):
    print(f"Load invoice template from {filepath}")
    print(filepath)
    if not nc:
        try:
            data = DocxTemplate(filepath)
        except:
            data = None
            errorbox = QMessageBox()
            errorbox.setText("Ausgewählte Date ist nicht lesbar (ist sie im richtigen Format?)")
            errorbox.exec_()
    else:
        print("Nextcloud loading")
    return data

def load_mail_adresses(filepath = '', nc = ''):
    print(f"Load Emaildata from {filepath}")
    if not nc:
        try:
            data = pd.read_excel(filepath,sheet_name="Mitglieder")["E-Mail"]
        except Exception as e:
            print(e)
            data = None
            errorbox = QMessageBox()
            errorbox.setText("Ausgewählte Date ist nicht lesbar (ist sie im richtigen Format?)")
            errorbox.exec_()
    else:
        print("Nextcloud loading (not implemented)")
    return data

def load_mail_template(filepath = ''):
    print("Email Template")
    print(filepath)
    env = Environment(loader=FileSystemLoader(os.path.dirname(filepath)), autoescape=select_autoescape())
    template = env.get_template(os.path.basename(filepath))
    return template



def load_energy_qovdata(filepath = "", nc = False):
    print(f"Load {filepath}")
    if not nc:
        try:
            # to acess the data you have to user multiindex.like energydata.data.loc[:,pd.IndexSlice["AT005120000000000000000030160301P",:,:,:]] or data.xs("Mario Buchinger",level="Name",axis=1)
            # First entry is zählpunktnummer, dann Name, dann prod/cons dann Art der Daten
            qovdata = pd.read_excel(filepath, sheet_name="QoV Log", skiprows=[7, 8, 9], header=[1, 2, 3, 6],
                                    index_col=[0])
            newnameindex = {x: x.replace(" ", "") for x in qovdata.columns.get_level_values(level="Name")}
            qovdata.index = pd.to_datetime(qovdata.index, format="%d.%m.%Y %H:%M:%S")
            qovdata = qovdata.sort_index()

        except Exception as e:
            print(e)
            errorbox = QMessageBox()
            errorbox.setText("Ausgewählte Date ist nicht lesbar (ist sie im richtigen Format?)")
            errorbox.exec_()
            return
    else:
        print("Nextcloud loading")

    return qovdata

def load_energy_data(filepath = "", nc = False, qov = False):
    print(f"Load {filepath}, qov: {qov}")
    if not nc:
        class LoadingDialog(QDialog):
            def __init__(self, message="Laden"):
                super().__init__()
                self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)  # keep on top
                self.setWindowTitle("Loading")
                self.setModal(True)  # modal dialog
                self.resize(250, 80)

                layout = QVBoxLayout()
                self.label = QLabel(message)
                self.label.setAlignment(Qt.AlignCenter)
                layout.addWidget(self.label)
                self.setLayout(layout)


        class Worker(QThread):
            finished = pyqtSignal(list)  # signal to return data
            progress = pyqtSignal(str)  # optional signal for messages

            def __init__(self, filepath, qov):
                super().__init__()
                self.filepath = filepath
                self.qov = qov

            def run(self):
                self.progress.emit("Loading Excel file...")
                try:
                    if self.qov:
                        data = pd.read_excel(self.filepath, sheet_name="QoV Log", skiprows=[7, 8, 9], header=[1, 2, 3, 6],
                                                index_col=[0])
                    else:
                        data = pd.read_excel(self.filepath,sheet_name="Energiedaten",skiprows=[7,8,9], header=[1,2,3,6],index_col=[0])
                    data.index = pd.to_datetime(data.index, format="%d.%m.%Y %H:%M:%S")
                    data = data.sort_index()

                    self.finished.emit([True,data])
                except Exception as e:
                    self.finished.emit([False,f"There was an error loading: {e}"])

        dlg = LoadingDialog("Laden, bitte warten...")
        dlg.show()
        df_container = {}
        def on_finished(emit):
            if emit[0]:
                df_container["df"] = emit[1]
                print(df_container)

            else:
                print(emit[1])
                errorbox = QMessageBox()
                errorbox.setText("Ausgewählte Datei ist nicht lesbar (ist sie im richtigen Format?)")
                errorbox.exec_()
                return
        worker = Worker(filepath,qov)
        worker.finished.connect(lambda df: dlg.close())
        worker.finished.connect(on_finished)

        worker.start()


        dlg.exec()


    else:
        print("Nextcloud loading")
    if df_container:
        return df_container["df"]
    else:
        return None

def load_faktura_member_export_template(filepath = "",nc =False, nc_instance = ''):
    print(f"Load {filepath}")
    if not nc:
        try:
            template = load_workbook(filename = filepath)
            print(template)
        except:
            errorbox = QMessageBox()
            errorbox.setText("Ausgewählte Date ist nicht lesbar (ist sie im richtigen Format?)")
            errorbox.exec_()
            return
    else:
        print("Nextcloud loading")
        template = nc_instance.files.download(filepath)
        template = load_workbook(BytesIO(template))
        print(template)
    return template

def load_new_member_data(data = ''):
    return data

def load_filepath(parent, title, filter="Excel (*.xlsx)", fileex=True, pathisdir = False,homedir = "", defaultfilename = ""):
    print("Lokal")
    if not pathisdir:
        if fileex:
            filepath, filter = QFileDialog.getOpenFileName(parent, title, homedir, filter)
        else:
            if defaultfilename:
                homedir = os.path.join(homedir,defaultfilename)
            filepath, filter = QFileDialog.getSaveFileName(parent, title, homedir, filter)
    else:
        filepath = QFileDialog.getExistingDirectory(parent, title, homedir)
    if filepath:
        return filepath
    else:
        return None

def load_masterdata(filepath, nc = False):
    print(f"Load Masterdata from {filepath}")
    if not nc:
        try:
            data = pd.read_excel(filepath,sheet_name="Mitglieder")
        except Exception as e:
            print(e)
            data = None
            errorbox = QMessageBox()
            errorbox.setText("Ausgewählte Date ist nicht lesbar (ist sie im richtigen Format?)")
            errorbox.exec_()
    else:
        print("Nextcloud loading (not implemented)")
    return data

def load_masterdata_meta(filepath):
    print(f"Load Metadata from {filepath}")
    df = pd.read_excel(filepath, header=None, dtype=str)

    result = {}
    for col in df.columns:
        col_values = df[col].tolist()  # remove empty cells
        col_values.append(np.nan)
        print(col_values)
        for i in range(len(col_values) - 2):
            this = col_values[i]
            next = col_values[i + 1]
            nextnext = col_values[i + 2]
            if not pd.isna(col_values[i + 1]) and pd.isna(col_values[i + 2]):
                result[this] = next
                print(this, next)
            else:
                print("no viable")

            i += 1
    return result


def check_whether_data_exists(mandates = None,masterdata = None,invoices = None,energydata = None, emails = None,
                              mandatesrequired = False, invoicedatarequired = False, masterdatarequired = False,
                              masterdataexporttemprequired = False, emailstemprequired = False,invoicestemprequired = False,energymetadatarequired = False,
                              energydataoptional = False, newmember = None, newmemberdatarequired = False, newmembertemprequired = False):
    data_missing = []
    check = True
    if mandatesrequired:
        if mandates.data is None:
            check = False
            data_missing.append("Mandate")
    if invoicedatarequired:
        if invoices.data is None:
            check = False
            data_missing.append("Rechnungsdaten")
    if emailstemprequired:
        if emails.template is None:
            check = False
            data_missing.append("Email Template")
    if masterdatarequired:
        if masterdata.data is None:
            check = False
            data_missing.append("EEG Faktura Stammdaten")
    if masterdataexporttemprequired:
        if masterdata.template is None:
            check = False
            data_missing.append("Vorlage EEG Faktura Stammdaten Expor")
    if invoicestemprequired:
        if invoices.template is None:
            check = False
            data_missing.append("Rechnungen Vorlage")
    if energydataoptional:
        if energydata.data is None:
            data_missing.append("Energiedaten")
    if energymetadatarequired:
        if energydata.metadata is None:
            check = False
            data_missing.append("Energiedaten QoV")
    if newmemberdatarequired:
        if newmember.data is None:
            check = False
            data_missing.append("Daten zum Neuen Mitglied")
    if newmembertemprequired:
        if newmember.template is None:
            check = False
            data_missing.append("Template zum export der Neuen Mitglied")
    return check, data_missing



masterdata = Data(load_masterdata,"",F_for_metadata_loading= load_masterdata_meta)
invoices = Data(load_invoices,load_invoice_template)
emails = Data(load_mail_adresses,load_mail_template)
energydata = Data(load_energy_data,"", F_for_metadata_loading=partial(load_energy_data,qov = True))
newmember = Data(load_new_member_data,load_faktura_member_export_template)

