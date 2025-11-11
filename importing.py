from tempfile import template
import datetime as dt
from functools import partial
from docxtpl import DocxTemplate
import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox
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
        try:
            # to acess the data you have to user multiindex.like energydata.data.loc[:,pd.IndexSlice["AT005120000000000000000030160301P",:,:,:]] or data.xs("Mario Buchinger",level="Name",axis=1)
            # First entry is zählpunktnummer, dann Name, dann prod/cons dann Art der Daten
            if qov:
                data = pd.read_excel(filepath, sheet_name="QoV Log", skiprows=[7, 8, 9], header=[1, 2, 3, 6],
                                        index_col=[0])
            else:
                data = pd.read_excel(filepath,sheet_name="Energiedaten",skiprows=[7,8,9], header=[1,2,3,6],index_col=[0])
            newnameindex = {x:x.replace(" ","") for x in data.columns.get_level_values(level = "Name")}
            data.index = pd.to_datetime(data.index, format = "%d.%m.%Y %H:%M:%S")
            data = data.sort_index()
        except Exception as e:
            print(e)
            errorbox = QMessageBox()
            errorbox.setText("Ausgewählte Date ist nicht lesbar (ist sie im richtigen Format?)")
            errorbox.exec_()
            return
    else:
        print("Nextcloud loading")

    return data

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



# mandates = Data(load_mandates,load_mandate_template)
masterdata = Data(load_masterdata,"")
invoices = Data(load_invoices,load_invoice_template)
emails = Data(load_mail_adresses,load_mail_template)
energydata = Data(load_energy_data,"", F_for_metadata_loading=partial(load_energy_data,qov = True))
newmember = Data(load_new_member_data,load_faktura_member_export_template)

template_debit = "/home/leander/gei/Abrechnung/Mandate/Musterdatei Import Lastschriften.csv"
template_transfer = "/home/leander/gei/Abrechnung/Mandate/Musterdatei Import Überweisungen.csv"
datamandate = "/home/leander/gei/Abrechnung/Mandate/lastschriftmandate.xlsx"
datainvoices = "/home/leander/gei/Abrechnung/2025 q1/CC100438_abrechnung_Abr_YQ-2025-1_export_ausgebessert.xlsx"
template_invoice_fp = "/home/leander/gei/faktura/pythonProject/template_invoice.docx"
fakturaexport = "/home/leander/gei/faktura/stammdaten_import/241206-vorlage-import-stammdaten_ls.xlsx"
masterdatafp = "/home/leander/gei/Abrechnung/2025 q1/CC100438-EEG-Masterdata-20250420.xlsx"
energydatafp = "/home/leander/gei/Abrechnung/2025 q2/CC100438-Energy-Report-20250401_20250630.xlsx"
emailtemplatefp = "/home/leander/gei/faktura/pythonProject/email_template.html"

# mandates.load_template(filepath = template_debit,filepath2 = template_transfer)
# mandates.load_data(filepath = datamandate)
# invoices.load_data(filepath = datainvoices)
invoices.load_template(filepath = template_invoice_fp)
# masterdata.load_data(filepath = masterdatafp)
# energydata.load_data(filepath = energydatafp)
# y = energydata.load_metadata(filepath = energydatafp)

# emails.load_template(filepath = emailtemplatefp)
# # newmember.load_template(filepath = fakturaexport)

# checking energydata
# qov_values = energydata.metadata.copy()
# qov_cols = energydata.metadata.columns.get_level_values(0) == "QoV"
# qov_values.columns = range(qov_values.shape[1])
# qov_values = qov_values.loc[:,qov_cols]
# qov_L3_values = (energydata.metadata.loc[:,qov_cols] == "L3").values
# times_qov_L3 = qov_L3_values.any(axis = 1)
# change = np.diff(times_qov_L3.astype(int))
# starts = np.where(change == 1)[0] + 1
# ends = np.where(change == -1)[0]
# if times_qov_L3[0]:
#     starts = np.insert(starts, 0, 0)
# if times_qov_L3[-1]:
#     ends = np.append(ends, len(times_qov_L3) - 1)
#
# qovL3_startendgroups = [(start,end) for start,end in zip(starts,ends)]
# report_list = []
# for start, end in qovL3_startendgroups:
#     columns_this_L3, = np.where(qov_L3_values[start])
#     names = np.unique(energydata.metadata.columns[qov_values.columns[columns_this_L3]-1].get_level_values('Name'))
#     shownames =', '.join(names)
#     if names.shape[0] > 3:
#         shownames = "All"
#     Metering_points = np.unique(energydata.metadata.columns[qov_values.columns[columns_this_L3]-1].get_level_values('MeteringpointID'))
#     days = (energydata.metadata.index[start].date(),energydata.metadata.index[end].date())
#     timerange = f"{energydata.metadata.index[start]} - {energydata.metadata.index[end]}"
#     print(days,timerange,names,Metering_points)
#     line = [energydata.metadata.index[start].date(),energydata.metadata.index[end].date(),shownames,timerange,', '.join(names),', '.join(Metering_points)]
#     report_list.append(line)
# report_df = x = pd.DataFrame(report_list,columns = ["Start Datum", "End Datum", "Namen Übersicht", "Zeitraum Details","Namen Details", "ZP Details"])
# names = np.unique( np.array([x[2] for x in energydata.data.columns]))
# for name in names:
# check for L3 in the columns and give start and end time