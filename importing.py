from tempfile import template
import datetime as dt

from Cython.Shadow import returns
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
    def __init__(self,F_for_data_loading,F_for_template_loading):
        self.f_load_data = F_for_data_loading
        self.f_load_template = F_for_template_loading

        self.data = None
        self.template = None
    def load_data(self,**kwargs):
        self.data = self.f_load_data(**kwargs)
        return self.data
    def load_template(self,**kwargs):
        self.template = self.f_load_template(**kwargs)
        return self.template

def load_mandates(filepath='',nc =False, nc_instance = ''):
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
        mandates = nc_instance.files.download(filepath)
        mandates = pd.read_excel(BytesIO(mandates), engine='openpyxl')
        print(mandates)
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


def load_energy_data(filepath = "", nc = False):
    print(f"Load {filepath}")
    if not nc:
        try:
            # to acess the data you have to user multiindex.like energydata.data.loc[:,pd.IndexSlice["AT005120000000000000000030160301P",:,:,:]] or data.xs("Mario Buchinger",level="Name",axis=1)
            # First entry is zählpunktnummer, dann Name, dann prod/cons dann Art der Daten
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


def load_filepath(parent, title, filter="Excel (*.xlsx)", fileex=True, pathisdir = False,homedir = ""):
    print("Lokal")
    if not pathisdir:
        if fileex:
            filepath, filter = QFileDialog.getOpenFileName(parent, title, homedir, filter)
        else:
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
               mandatesrequired = False, invoicedatarequired = False, masterdatarequired = False,masterdataexporttemprequired = False, emailstemprequired = False,invoicestemprequired = False, energydataoptional = False):
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
    return check, data_missing



mandates = Data(load_mandates,load_mandate_template)
masterdata = Data(load_masterdata,load_faktura_member_export_template)
invoices = Data(load_invoices,load_invoice_template)
emails = Data(load_mail_adresses,load_mail_template)
energydata = Data(load_energy_data,"")

template_debit = "/home/leander/gei/faktura/pythonProject/data/Musterdatei Import Lastschriften.csv"
template_transfer = "/home/leander/gei/faktura/pythonProject/data/Musterdatei Import Überweisungen.csv"
datamandate = "/home/leander/gei/export_infinity/lastschriftmandate.xlsx"
datainvoices = "/home/leander/gei/Abrechnung/2025 q1/CC100438_abrechnung_Abr_YQ-2025-1_export.xlsx"
template_invoice_fp = "/home/leander/gei/faktura/pythonProject/template_invoice.docx"
fakturaexport = "/home/leander/gei/faktura/stammdaten_import/241206-vorlage-import-stammdaten_ls.xlsx"
masterdatafp = "/home/leander/gei/Abrechnung/2025 q1/CC100438-EEG-Masterdata-20250420.xlsx"
energydatafp = "/home/leander/gei/Abrechnung/2025 q1/CC100438-Energy-Report-20250101_20250331(2).xlsx"
emailtemplatefp = "/home/leander/gei/faktura/pythonProject/email_template.html"

# mandates.load_template(template_debit,template_transfer)
mandates.load_data(filepath = datamandate)
invoices.load_data(filepath = datainvoices)
# invoices.load_template(filepath = template_invoice_fp)
masterdata.load_data(filepath = masterdatafp)
energydata.load_data(filepath = energydatafp)
emails.load_template(filepath = emailtemplatefp)
# # newmember.load_template(filepath = fakturaexport)

#
#
# #personaldata =  masterdata.data.iloc[0]
# # invoicedata = invoices.data['detailed'][invoices.data['detailed']['Empfänger Name'] == fullname]
# #invoicetemplate = invoices.template
#
# import matplotlib.pyplot as plt
# from matplotlib.dates import DateFormatter, MonthLocator
# from exporting import check_doubles
#
# # fullname = f"{personaldata['Name 1']} {personaldata['Name 2']}"
# invoicedata = invoices .data["detailed"]
# energydata = energydata.data
# debit,transfer, doublesprocess = check_doubles(invoicedata)
# # debit, transfer = debit[cols_for_tbl], transfer[cols_for_tbl]
#
# name = "Leander Stark"
# print(f"Make invoice for {name}")
# invoicethis = invoicedata[invoicedata["Empfänger Name"] == name]
# debits_this =  debit[debit["Empfänger Name"] == name]
# transfers_this = transfer[transfer["Empfänger Name"] == name]
# total_transfer_debit = "Rechnung"
# #if diff > 0 debit bigger than transfer -> Rechnung otherwise Gutschrift
# diff = debits_this["Pos. Bruttobetrag"].sum(axis = 0) -transfers_this["Pos. Bruttobetrag"].sum(axis = 0)
# if diff > 0:
#     total_transfer_debit = "Rechnung"
#     transfers_this.loc[:,['Pos. Nettobetrag','Pos. Bruttobetrag']]= -transfers_this.loc[:,['Pos. Nettobetrag','Pos. Bruttobetrag']]
#
# if diff < 0:
#     total_transfer_debit = "Gutschrift"
#     debits_this.loc[:,['Pos. Nettobetrag','Pos. Bruttobetrag']]= -debits_this.loc[:,['Pos. Nettobetrag','Pos. Bruttobetrag']]
#
#
#
#
# all_position_for_tbl = pd.concat([debits_this,transfers_this], axis = 0)
# all_position_for_tbl['Pos. Invoicetext'] = (all_position_for_tbl['Pos. Tarif'] + "\n"+ all_position_for_tbl['Pos. Text'])
# all_position_for_tbl['Pos. Preis / Einheit'] = all_position_for_tbl['Pos. Preis / Einheit'].astype(float)
# all_position_for_tbl.loc[all_position_for_tbl['Pos. Preis / Einheit Euro/Cent'] == "Ct",'Pos. Preis / Einheit'] = all_position_for_tbl.loc[all_position_for_tbl['Pos. Preis / Einheit Euro/Cent'] == "Ct",'Pos. Preis / Einheit']/ 100
# sumnetto = str(round(all_position_for_tbl['Pos. Nettobetrag'].sum(),2))
# sumbrutto = str(round(all_position_for_tbl['Pos. Bruttobetrag'].sum(),2))
#
#
#
# cols_for_tbl = ['Pos. Invoicetext','Pos. Menge','Pos. Preis / Einheit','Pos. Nettobetrag','Pos. UST %','Pos. Bruttobetrag']
# moneycols = ['Pos. Menge','Pos. Preis / Einheit','Pos. Nettobetrag','Pos. Bruttobetrag']
# all_position_for_tbl[moneycols] = all_position_for_tbl[moneycols].map("{0:.2f}".format)
# all_position_for_tbl = all_position_for_tbl[cols_for_tbl]
#
# invoicequart = invoicethis["Abrechnung"].iloc[0]
# year,quart = invoicequart.split("-")[-2],invoicequart.split("-")[-1]
#
# parsing_dict = {}
# parsing_dict["EmpfängerName"] = invoicethis["Empfänger Name"].values[0]
# parsing_dict["EmpfängerAdresse1"] =invoicethis["Empfänger Adresse 1"].values[0]
# parsing_dict["EmpfängerAdresse2"] =invoicethis["Empfänger Adresse 2"].values[0]
# parsing_dict["invoiceitemstbl_contents"]= all_position_for_tbl.values.tolist()
# parsing_dict["Rechnung_Gutschrift"] = total_transfer_debit
# invoicenumberstr = ""
# if total_transfer_debit == "Rechnung":
#     invoicenumberstr = debits_this["Nummer"].iloc[0]
#     parsing_dict["Überweisungstext"] = "Die gegenständliche Rechnungsforderung wird vereinbarungsgemäß von Ihr Konto eingezogen."
# if total_transfer_debit == "Gutschrift":
#     invoicenumberstr = transfers_this["Nummer"].iloc[0]
#     parsing_dict["Überweisungstext"] = "Die gegenständliche Rechnungsforderung wird vereinbarungsgemäß auf Ihr Konto überwiesen."
#
# parsing_dict["Rechnungsnummer"] = invoicenumberstr
# parsing_dict["DatumHeute"] = dt.datetime.today().strftime("%d.%m.%Y")
# parsing_dict["Jahr"] = year
# parsing_dict["Quartal"] = quart
# parsing_dict["TotalSumNetto"] = sumnetto
# parsing_dict["TotalSumBrutto"] = sumbrutto
#
#
#
# print(parsing_dict)
# image_stream = BytesIO()
# fig, axs = plt.subplots(2)
# # prepare data for plot
# # check whether the energy direction is always the same
# energydirections = energydata.xs(name, level="Name", axis=1).columns.get_level_values(level="Energy direction")
# meteringpointids = energydata.xs(name, level="Name", axis=1).columns.get_level_values(level="MeteringpointID")
# for energydirection,meteringpointid in zip(energydirections.unique(),meteringpointids.unique()):
#     print(energydirection,meteringpointid)
#     total_energy_consumption = energydata.loc[:, pd.IndexSlice[:, name,energydirection, "Gesamtverbrauch lt. Messung (bei Teilnahme gem. Erzeugung) [KWH]"]].sort_index()
#     energy_consumption_through_eg = energydata.loc[:, pd.IndexSlice[:, name, energydirection,"Eigendeckung gemeinschaftliche Erzeugung [KWH]"]].sort_index()
#     # dailysums_total_energy = total_energy_consumption.groupby(total_energy_consumption.index.strftime('%d.%m.%Y')).sum()
#     total_energy = total_energy_consumption.groupby(total_energy_consumption.index.strftime('%Y-%U')).sum()
#     total_energy_through_eeg =  energy_consumption_through_eg.groupby(energy_consumption_through_eg.index.strftime('%Y-%U')).sum()
#     xaxis = pd.date_range(start=total_energy_consumption.index[0], end=total_energy_consumption.index[-1], freq='W-MON')
#     #throuw away the first and last week since they are only partly and will change the sum
#     total_energy = total_energy.iloc[1:-1]
#     total_energy_through_eeg = total_energy_through_eeg.iloc[1:-1]
#     xaxis = xaxis[1:-1]
#
#     dailyhourmean_total_energy  = total_energy_consumption.groupby(total_energy_consumption.index.hour).mean()
#     dailyhourmean_through_eeg = energy_consumption_through_eg.groupby(energy_consumption_through_eg.index.hour).mean()
#     barwidth = dt.timedelta(days=6)
#     axs[0].bar(xaxis,total_energy.values.T[0],barwidth, label = "Gesamte Energie")
#     axs[0].bar(xaxis,total_energy_through_eeg.values.T[0],barwidth, label = "Energie über die Energiegemeinschaft")
#     axs[0].xaxis.set_major_locator(MonthLocator())
#     axs[0].xaxis.set_major_formatter(DateFormatter('%b %Y'))
#     # axs[0].xaxis.set_label_position("right")
#     axs[0].set_ylabel("kWh in der Woche")
#     axs[0].legend(loc = 1)
#     # Creating the secondary x-axis
#     # ax2 = axs[0].twiny()
#     # Setting the secondary x-axis ticks as month labels
#     # ax2.set_xlim(axs[0].get_xlim())  # Make sure both x-axes have the same range
#     # ax2.set_xticks(dailysums_total_energy.index.values)  # Set tick positions on the secondary axis
#     # ax2.set_xticklabels([date.strftime('%b') for date in dailysums_total_energy.index.values])  # Format as month names (e.g., Jan, Feb, etc.)
#     # ax2.set_xlabel('Month')
#
#     axs[1].bar(dailyhourmean_total_energy.index,dailyhourmean_total_energy.values.T[0])
#     axs[1].bar(dailyhourmean_through_eeg.index,dailyhourmean_through_eeg.values.T[0])
#     axs[1].set_xticks([0, 6, 12, 18,24])
#     axs[1].set_xticklabels(["0 Uhr","6 Uhr","12 Uhr","18 Uhr", "24 Uhr"])
#     axs[1].set_ylabel("Durchschnittliche Leistung \nzur Tageszeit [kW]")
#
#
#
# plt.show()