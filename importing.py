from tempfile import template
import datetime as dt
import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox

# mandate =pd.read_excel("/home/leander/gei/faktura/abrechnung_24_q3/CC100438_abrechnung_Abr_YQ-2024-3_export(1).xlsx",sheet_name="Liste")
# df = pd.read_excel("/home/leander/gei/faktura/abrechnung_24_q3/CC100438_abrechnung_Abr_YQ-2024-3_export(1).xlsx",sheet_name="Liste")
class Data():
    def __init__(self,F_for_data_loading,F_for_template_loading):
        self.f_load_data = F_for_data_loading
        self.f_load_template = F_for_template_loading

        self.data = []
        self.template_for_export = []
    def load_data(self,filepath):
        self.data = self.f_load_data(filepath)
        return self.data
    def load_template(self,filepath1,filepath2 = ''):
        self.template_for_export = self.f_load_template(filepath1,filepath2 = filepath2)
        return self.template_for_export

def load_mandates(filepath,nc = False):
    print(f"Load {filepath}")
    if not nc:
        try:
            mandates = pd.read_excel(filepath)
        except:
            errorbox = QMessageBox()
            errorbox.setText("Ausgewählte Date ist nicht lesbar (ist sie im richtigen Format?)")
            errorbox.exec_()
            return
    else:
        print("Nextcloud loading")
    try:
        mandates = mandates[~mandates['Mandatsausstellungsdatum (Datum auf dem Vertrag)'].isna()]
        mandates['Mandatsausstellungsdatum'] = pd.to_datetime(mandates['Mandatsausstellungsdatum (Datum auf dem Vertrag)'],dayfirst=True)
        mandates = mandates.drop('Mandatsausstellungsdatum (Datum auf dem Vertrag)', axis=1)
        mandates = mandates.rename(columns={'Vorname (gleich wie in eegfaktura)': 'Vorname', 'Nachname (gleich wie in eegfaktura)': 'Nachname','Mitgliedsnummer aus eegfaktura ist auch die Mandatsreferenz':'Mitgliedsnummer'})
    except:
        errorbox = QMessageBox()
        errorbox.setText("Ausgewählte Date ist nicht lesbar (ist sie im richtigen Format?)")
        errorbox.exec_()
        return
    return mandates
def load_mandate_template(filepath_lastschrift, filepath2= ''):
    print(f"Load {filepath_lastschrift},{filepath2}")
    templates = {}
    templates["debit"] = pd.read_csv(filepath_lastschrift,delimiter=";")
    templates["transfer"]  = pd.read_csv(filepath2,delimiter=";")
    return templates


def load_invoices(filepath,nc = False):
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

def load_invoice_template(filepath):
    print(4)
    print(filepath)
def load_mail_adresses(filepath):
    print("Email")
    print(filepath)
def load_mail_template(filepath):
    print("Email Template")
    print(filepath)

mandates = Data(load_mandates,load_mandate_template)
invoices = Data(load_invoices,load_invoice_template)
emails = Data(load_mail_adresses,load_mail_template)

template_debit = "/home/leander/gei/faktura/pythonProject/data/Musterdatei Import Lastschriften.csv"
template_transfer = "/home/leander/gei/faktura/pythonProject/data/Musterdatei Import Überweisungen.csv"
datamandate = "/home/leander/gei/export_infinity/lastschriftmandate.xlsx"
datainvoices = "/home/leander/gei/faktura/pythonProject/data/CC100438_abrechnung_final.xlsx"

# mandates.load_template(template_debit,template_transfer)
# mandates.load_data(datamandate)
# invoices.load_data(datainvoices)
