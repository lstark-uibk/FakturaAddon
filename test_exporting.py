from tempfile import template
import datetime as dt
import numpy as np
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
            errorbox.setText("Ausgewählte Datei ist nicht lesbar (ist sie im richtigen Format?)")
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
        errorbox.setText("Ausgewählte Datei ist nicht lesbar (ist sie im richtigen Format?)")
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
            errorbox.setText("Ausgewählte Datei ist nicht lesbar (ist sie im richtigen Format?)")
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
            errorbox.setText("Ausgewählte Datei ist nicht lesbar (ist sie im richtigen Format?)")
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
            errorbox.setText("Ausgewählte Datei ist nicht lesbar (ist sie im richtigen Format?)")
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
            errorbox.setText("Ausgewählte Datei ist nicht lesbar (ist sie im richtigen Format?)")
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
            errorbox.setText("Ausgewählte Datei ist nicht lesbar (ist sie im richtigen Format?)")
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
            errorbox.setText("Ausgewählte Datei ist nicht lesbar (ist sie im richtigen Format?)")
            errorbox.exec_()
    else:
        print("Nextcloud loading (not implemented)")
    return data

def check_whether_data_exists(mandates = None,masterdata = None,invoices = None,energydata = None,
               mandatesrequired = False, invoicedatarequired = False, masterdatarequired = False,masterdataexporttemprequired = False, invoicestemprequired = False, energydataoptional = False):
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
# mandates.load_data(datamandate)
invoices.load_data(filepath = datainvoices)
# invoices.load_template(filepath = template_invoice_fp)
masterdata.load_data(filepath = masterdatafp)
energydata.load_data(filepath = energydatafp)
# emails.load_template(filepath = emailtemplatefp)
# # newmember.load_template(filepath = fakturaexport)

#
from cProfile import label
from pathlib import Path
import os
import pandas as pd
import datetime as dt
# from docxtpl import DocxTemplate
# import openpyxl
# from docx.enum.table import WD_TABLE_ALIGNMENT
# from docx import Document
# from docx.shared import Cm
# import numpy as np
# from datetime import date
# import datetime as dt
# from PyInquirer import prompt
# import pprint
# import tkinter as tk
import subprocess
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter,MonthLocator
from io import BytesIO


def check_doubles(invoices):
    debit = invoices[(invoices["Dokumenttyp"] == "Rechnung")]
    transfer = invoices[(invoices["Dokumenttyp"] == "Gutschrift")|(invoices["Dokumenttyp"] == "Information")]
    debitdoubles = debit["Empfänger Name"][debit["Empfänger Name"].isin(transfer["Empfänger Name"])].index
    transferdoubles = transfer["Empfänger Name"][transfer["Empfänger Name"].isin(debit["Empfänger Name"])].index
    doublesprocess = {"Name":[],"Debit":[],"Transfer":[],"Type":[],"Final":[]}
    for i,j in zip(debitdoubles,transferdoubles):
        doublesprocess["Name"].append(debit.loc[i,"Empfänger Name"])
        doublesprocess["Debit"].append(debit.loc[i,"Pos. Bruttobetrag"])
        doublesprocess["Transfer"].append(transfer.loc[j,"Pos. Bruttobetrag"])
        if debit.loc[i,'Pos. Bruttobetrag'] < transfer.loc[j,'Pos. Bruttobetrag']:
            finalsum = transfer.loc[j,'Pos. Bruttobetrag'] - debit.loc[i,'Pos. Bruttobetrag']
            # transfer.loc[j, 'Pos. Bruttobetrag'] = finalsum
            # debit = debit.drop(i)
            doublesprocess["Type"].append("Überweisung")
            doublesprocess["Final"].append(finalsum)

        else:
            finalsum = debit.loc[i,'Pos. Bruttobetrag'] - transfer.loc[j,'Pos. Bruttobetrag']
            # debit.loc[i, 'Pos. Bruttobetrag'] = finalsum
            # transfer = transfer.drop(j)
            doublesprocess["Type"].append("Lastschrift")
            doublesprocess["Final"].append(finalsum)

    return debit,transfer, doublesprocess


energydata,invoicedata,invoicetemplate = energydata.data,invoices.data["detailed"], invoices.template
debit, transfer, doublesprocess = check_doubles(invoicedata)

name = "Johannes Riedel"

print(f"Make invoice for {name}")
invoicethis = invoicedata[invoicedata["Empfänger Name"] == name]
print(invoicethis)
debits_this = debit[debit["Empfänger Name"] == name]
transfers_this = transfer[transfer["Empfänger Name"] == name]
total_transfer_debit = "Rechnung"
# if diff > 0 debit bigger than transfer -> Rechnung otherwise Gutschrift
diff = debits_this["Pos. Bruttobetrag"].sum(axis=0) - transfers_this["Pos. Bruttobetrag"].sum(axis=0)
if diff > 0:
    total_transfer_debit = "Rechnung"
    transfers_this.loc[:, ['Pos. Nettobetrag', 'Pos. Bruttobetrag']] = -transfers_this.loc[:,
                                                                        ['Pos. Nettobetrag', 'Pos. Bruttobetrag']]

if diff < 0:
    total_transfer_debit = "Gutschrift"
    debits_this.loc[:, ['Pos. Nettobetrag', 'Pos. Bruttobetrag']] = -debits_this.loc[:,
                                                                     ['Pos. Nettobetrag', 'Pos. Bruttobetrag']]

all_position_for_tbl = pd.concat([debits_this, transfers_this], axis=0)
postext = all_position_for_tbl.copy()
postext.loc[postext['Pos. Text'].str.contains('Mitgliedsgebuer'), 'Pos. Text'] = ''
all_position_for_tbl['Pos. Invoicetext'] = (all_position_for_tbl['Pos. Tarif'] + "\n" + postext['Pos. Text'])
all_position_for_tbl['Pos. Preis / Einheit'] = all_position_for_tbl['Pos. Preis / Einheit'].astype(float)

all_position_for_tbl['Pos. Menge'] = all_position_for_tbl['Pos. Menge'].astype(str) + " " + all_position_for_tbl[
    'Pos. Mengeneinheit'].fillna("")
all_position_for_tbl.loc[all_position_for_tbl['Pos. Preis / Einheit Euro/Cent'] == "Ct", 'Pos. Preis / Einheit'] = \
all_position_for_tbl.loc[all_position_for_tbl['Pos. Preis / Einheit Euro/Cent'] == "Ct", 'Pos. Preis / Einheit'] / 100
sumnetto = str(round(all_position_for_tbl['Pos. Nettobetrag'].sum(), 2))
sumbrutto = str(round(all_position_for_tbl['Pos. Bruttobetrag'].sum(), 2))

cols_for_tbl = ['Pos. Invoicetext', 'Pos. Menge', 'Pos. Preis / Einheit', 'Pos. Nettobetrag', 'Pos. UST %',
                'Pos. Bruttobetrag']
moneycols = ['Pos. Preis / Einheit', 'Pos. Nettobetrag', 'Pos. Bruttobetrag']
all_position_for_tbl[moneycols] = all_position_for_tbl[moneycols].map("{0:.2f}".format)
all_position_for_tbl = all_position_for_tbl[cols_for_tbl]

invoicequart = invoicethis["Abrechnung"].iloc[0]
year, quart = invoicequart.split("-")[-2], invoicequart.split("-")[-1]

parsing_dict = {}
parsing_dict["EmpfängerName"] = invoicethis["Empfänger Name"].values[0]
parsing_dict["EmpfängerAdresse1"] = invoicethis["Empfänger Adresse 1"].values[0]
parsing_dict["EmpfängerAdresse2"] = invoicethis["Empfänger Adresse 2"].values[0]
parsing_dict["invoiceitemstbl_contents"] = all_position_for_tbl.values.tolist()
parsing_dict["Rechnung_Gutschrift"] = total_transfer_debit
invoicenumberstr = ""
if total_transfer_debit == "Rechnung":
    invoicenumberstr = debits_this["Nummer"].iloc[0]
    parsing_dict[
        "Überweisungstext"] = "Die gegenständliche Rechnungsforderung wird vereinbarungsgemäß von Ihrem Konto eingezogen."
if total_transfer_debit == "Gutschrift":
    invoicenumberstr = transfers_this["Nummer"].iloc[0]
    parsing_dict[
        "Überweisungstext"] = "Die gegenständliche Rechnungsforderung wird vereinbarungsgemäß auf Ihr Konto überwiesen."

parsing_dict["Rechnungsnummer"] = invoicenumberstr
parsing_dict["DatumHeute"] = dt.datetime.today().strftime("%d.%m.%Y")
parsing_dict["Jahr"] = year
parsing_dict["Quartal"] = quart
parsing_dict["TotalSumNetto"] = sumnetto
parsing_dict["TotalSumBrutto"] = sumbrutto

# print(parsing_dict)
image_stream = BytesIO()
sizemutiplier = 1.9
fig, axs = plt.subplots(3,figsize=(4.2*sizemutiplier, 2.7*sizemutiplier))
fig.suptitle("jöaögigeuaögu\n\söafui4wn\nasoxdgeibriuög\niaoghoe")
# prepare data for plot
# check whether the energy direction is always the same
energydata = energydata.replace('NaN', 0)
try:
    energydata.loc[:, pd.IndexSlice[:, name, :, :]]
except:
    print("This Name is not in the Energydata try adding a whitespace")
    name += " "
energydirections = energydata.xs(name, level="Name", axis=1).columns.get_level_values(level="Energy direction")
meteringpointids = energydata.xs(name, level="Name", axis=1).columns.get_level_values(level="MeteringpointID")
hatches = ['','/', '\\', '|', '-', '+', 'x', 'o', 'O', '.', '*']
edgecolors = ['none','black','green','red']
hatches_this_person = [x for x,y in zip(hatches,range(0,meteringpointids.unique().shape[0]))] # take nr meteringpoints hatches
hatchnr = -1
for energydirection in energydirections.unique():
    for meteringpointid in meteringpointids.unique():
        print(energydirection, meteringpointid)
        hatchnr += 1
        hatch = hatches[hatchnr]
        edgecolor = edgecolors[hatchnr]

        if energydirection == "GENERATION":
            total_energy = energydata.loc[:, pd.IndexSlice[meteringpointid, name, energydirection,
                                             "Gesamte gemeinschaftliche Erzeugung [KWH]"]].sort_index()

            energy_through_evu = energydata.loc[:, pd.IndexSlice[meteringpointid, name, energydirection,
                                                   "Gesamt/Überschusserzeugung, Gemeinschaftsüberschuss [KWH]"]].sort_index()
            energy_through_eg = pd.DataFrame((np.nan_to_num(total_energy.values,0 )- np.nan_to_num(energy_through_evu.values,0)), index=total_energy.index)

            plottext1 = f"Energielieferung an außerhalb der Energiegemeinschaft von ZP {meteringpointid[-6:]}"
            plottext2 = f"Energielieferung über unsere Energiegemeinschaft von ZP {meteringpointid[-6:]}"
            totalsumeg = energy_through_eg.values.sum()
            shareeg = totalsumeg / total_energy.sum() * 100
            parsing_dict[
                "TextfürVerbrauch"] = f"Insgesamt wurden {totalsumeg:.1f}kWh an die Energiegemeischaft verkauft. \nDies ist {shareeg:.1f}% der gesamten erzeugten Energie."
        else:
            total_energy = energydata.loc[:, pd.IndexSlice[meteringpointid, name, energydirection,
                                             "Gesamtverbrauch lt. Messung (bei Teilnahme gem. Erzeugung) [KWH]"]].sort_index()
            energy_through_eg = energydata.loc[:, pd.IndexSlice[meteringpointid, name, energydirection,
                                                  "Eigendeckung gemeinschaftliche Erzeugung [KWH]"]].sort_index()
            energy_through_evu = pd.DataFrame((total_energy.values - energy_through_eg.values), index=total_energy.index)

            plottext1 = f"Energiebezug von Stromlieferant für ZP {meteringpointid[-6:]}"
            plottext2 = f"Energiebezug über unsere Energiegemeinschaft für ZP {meteringpointid[-6:]}"
            totalsumeg = energy_through_eg.sum()
            shareeg = totalsumeg / total_energy.sum() * 100
            parsing_dict[
                "TextfürVerbrauch"] = f"Insgesamt wurden {totalsumeg:.1f}kWh über die Energiegemeischaft bezogen. \nDies ist {shareeg:.1f}% des Gesamtenergieverbrauchs."
        # dailysums_total_energy = total_energy_consumption.groupby(total_energy_consumption.index.strftime('%d.%m.%Y')).sum()
        # plt.plot(total_energy)
        # plt.plot(energy_through_eg)
        # plt.plot(energy_through_evu)
        energy_through_evu_weeklysum = energy_through_evu.groupby(energy_through_evu.index.strftime('%Y-%W')).sum()
        xaxis_energy_through_evu = []
        xaxis_energy_through_eg = []
        for week, weektotalenergy in energy_through_evu.groupby(energy_through_evu.index.strftime('%Y-%W')):
            xaxis_energy_through_evu.append(dt.datetime.strptime(f"{week}-1", "%Y-%U-%w"))  # the one saying we take the monday of the week
        energy_through_eg_weeklysum = energy_through_eg.groupby(energy_through_eg.index.strftime('%Y-%W')).sum()
        for week, weekenergy_through_eg in energy_through_eg.groupby(energy_through_eg.index.strftime('%Y-%W')):
            xaxis_energy_through_eg.append(dt.datetime.strptime(f"{week}-1", "%Y-%U-%w"))

        # throuw away the first and last week since they are only partly and will change the sum
        energy_through_evu_weeklysum = energy_through_evu_weeklysum.iloc[1:-1]
        energy_through_eg_weeklysum = energy_through_eg_weeklysum.iloc[1:-1]
        xaxis_energy_through_evu = xaxis_energy_through_evu[1:-1]
        #shift one day earlier

        xaxis_energy_through_eg = xaxis_energy_through_eg[1:-1]


        dailyhourmean_energy_through_evu = energy_through_evu.groupby(energy_through_evu.index.hour).mean()
        dailyhourmean_through_eeg = energy_through_eg.groupby(
            energy_through_eg.index.hour).mean()
        barwidth = dt.timedelta(days=3)
        xaxis_energy_through_evu = [x-0.5*barwidth for x in xaxis_energy_through_evu]
        xaxis_energy_through_eg = [x+0.5*barwidth for x in xaxis_energy_through_eg]


        energy_through_evu_weekday =energy_through_evu.groupby(energy_through_evu.index.strftime('%Y-%w')).sum()
        energy_through_evu_weekday = np.roll(energy_through_evu_weekday.values, shift=-1, axis=0).T # because it starts with sunday roll by one day
        energy_through_eg_weekday =energy_through_eg.groupby(energy_through_eg.index.strftime('%Y-%w')).sum()
        energy_through_eg_weekday = np.roll(energy_through_eg_weekday.values, shift=-1, axis=0).T # because it starts with sunday roll by one day
        xaxis_weekdays = ['Mo', 'Do', 'Mi', 'Do', 'Fr', 'Sa', 'So']
        # axs[0].bar(xaxis_energy_through_evu, energy_through_evu_weeklysum.values.T[0], width=barwidth,
        #            label=plottext1,
        #            color="#69a4dc", hatch=hatch, edgecolor=edgecolor)
        # print(f"energy_through_evu_weeklysum:{energy_through_evu_weeklysum}")
        energy_through_eg_weeklysum.values.T
        print(energy_through_evu_weeklysum)
        axs[1].bar(xaxis_energy_through_evu, energy_through_evu_weeklysum, width=barwidth,
                   label=plottext1,
                   color="#69a4dc", hatch=hatch, edgecolor=edgecolor)
        # print(f"energy_through_evu_weeklysum:{energy_through_evu_weeklysum}")
        # axs[1].bar(xaxis_energy_through_eg, energy_through_eg_weeklysum, width=barwidth,
        #            label=plottext2,
        #            color="#f8ae42", hatch=hatch, edgecolor=edgecolor)
        # print(f"energy_through_evu_weeklysum:{energy_through_eeg_weeklysum}")
        axs[1].xaxis.set_major_locator(MonthLocator())
        axs[1].xaxis.set_major_formatter(DateFormatter('%b %Y'))
        # axs[1].xaxis.set_label_position("right")
        axs[1].set_ylabel("kWh")

        # Creating the secondary x-axis
        # ax2 = axs[1].twiny()
        # Setting the secondary x-axis ticks as month labels
        # ax2.set_xlim(axs[1].get_xlim())  # Make sure both x-axes have the same range
        # ax2.set_xticks(dailysums_total_energy.index.values)  # Set tick positions on the secondary axis
        # ax2.set_xticklabels([date.strftime('%b') for date in dailysums_total_energy.index.values])  # Format as month names (e.g., Jan, Feb, etc.)
        # ax2.set_xlabel('Month')
        barwidth = 0.4
        axs[2].bar(dailyhourmean_energy_through_evu.index - 0.5 * barwidth,
                   dailyhourmean_energy_through_evu, width=barwidth, color="#69a4dc", hatch=hatch,
                   edgecolor=edgecolor)
        axs[2].bar(dailyhourmean_through_eeg.index + 0.5 * barwidth, dailyhourmean_through_eeg,
                   width=barwidth, color="#f8ae42", hatch=hatch, edgecolor=edgecolor)
        axs[2].set_xticks([0, 6, 12, 18, 24])
        axs[2].set_xticklabels(["0 Uhr", "6 Uhr", "12 Uhr", "18 Uhr", "24 Uhr"])
        axs[2].set_ylabel("kW")

        barwidth = 0.4
        # axs[3].bar(np.linspace(0, 6, 7) - 0.5 * barwidth, energy_through_evu_weekday.flatten(), width=barwidth,
        #            color="#69a4dc", tick_label=xaxis_weekdays, hatch=hatch, edgecolor=edgecolor)
        # axs[3].bar(np.linspace(0, 6, 7) + 0.5 * barwidth, energy_through_eg_weekday.flatten(), width=barwidth,
        #            color="#f8ae42", tick_label=xaxis_weekdays, hatch=hatch, edgecolor=edgecolor)
        # axs[3].set_ylabel("kWh")
        # axs[1].legend(fontsize='small')
#
plt.show()