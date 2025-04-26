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
import numpy as np
import subprocess
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter,MonthLocator
import matplotlib
matplotlib.use('Agg')
from io import BytesIO
from PyQt5.QtWidgets import QMessageBox


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



def produce_sepa_export_dfs(invoices_selected_persons,mandates,creditor_ID):
    debit,transfer, doublesprocess = check_doubles(invoices_selected_persons)

    def create_one_line_debit(invoicelistline,creditor_ID,mandates,datatype = "debit"):
        print(invoicelistline["Empfänger Name"])
        columns_debit_export = ['Fälligkeitsdatum', 'Zahlungspflichtiger Name',
       'Zahlungspflichtiger Adresse', 'Zahlungspflichtiger Ort',
       'Zahlungspflichtiger IBAN', 'Zahlungspflichtiger BIC', 'Betrag in EUR',
       'Zahlungsreferenz/Verwendungszweck', 'Auftraggeberinformation',
       'Geschäftsvorfallcode', 'Auftraggeber IBAN',
       'Abweichender Auftraggeber', 'Mandatsausstellungsdatum', 'Creditor ID',
       'Mandatsreferenz', 'Art der Verwendung', 'Firmenlastschrift']

        columns_transfer_export = ['Durchführungsdatum', 'Empfänger Name', 'Empfänger Adresse',
       'Empfänger Ort', 'Empfänger IBAN', 'Empfänger BIC', 'Betrag in EUR',
       'Zahlungsreferenz/Verwendungszweck', 'Auftraggeberinformation',
       'Geschäftsvorfallcode', 'Dringlichkeit', 'Auftraggeber IBAN',
       'Abweichender Auftraggeber']
        if datatype == "debit":
            exportline = pd.Series([""] * len(columns_debit_export), index=columns_debit_export)
            exportline["Fälligkeitsdatum"] = dt.datetime.today().strftime("%d.%m.%Y")
            matchingdictinvoice = {
                "Zahlungspflichtiger Name": "Empfänger Name",
                "Zahlungspflichtiger Adresse": "Empfänger Adresse 1",
                "Zahlungspflichtiger Ort": "Empfänger Adresse 2",
                "Zahlungspflichtiger IBAN": "Empfänger Konto IBAN",
                "Betrag in EUR": "Pos. Bruttobetrag",
                "Auftraggeber IBAN": "Ersteller IBAN"
            }
            exportline["Mandatsreferenz"] = f"{invoicelistline['Empfänger Mitgliedsnummer']:03}"
            exportline["Creditor ID"] = creditor_ID


            mandateline = mandates.data[(mandates.data["Vorname"] == invoicelistline["Empfänger Vorame"])]
            matchingmandate = True
            if not pd.isna(invoicelistline["Empfänger Nachname"]):
                mandateline = mandateline[(mandateline["Nachname"] == invoicelistline["Empfänger Nachname"])]
            if mandateline.size > 0:
                exportline["Mandatsausstellungsdatum"] = mandateline["Mandatsausstellungsdatum"].iloc[0].strftime("%d.%m.%Y")
                exportline["Firmenlastschrift"] = mandateline["Firmenlastschrift"].iloc[0]
            else:
                exportline["Creditor ID"] = 0
                matchingmandate = False


        elif datatype == "transfer":
            exportline = pd.Series([""] * len(columns_transfer_export), index=columns_transfer_export)
            exportline["Durchführungsdatum"] = dt.datetime.today().strftime("%d.%m.%Y")
            matchingdictinvoice = {
                "Empfänger Name": "Empfänger Name",
                "Empfänger Adresse": "Empfänger Adresse 1",
                "Empfänger Ort": "Empfänger Adresse 2",
                "Empfänger IBAN": "Empfänger Konto IBAN",
                "Betrag in EUR": "Pos. Bruttobetrag",
                "Auftraggeber IBAN": "Ersteller IBAN"
            }

        for exportcol in matchingdictinvoice.keys():
            exportline[exportcol] = invoicelistline[matchingdictinvoice[exportcol]]


        def get_quartal_out_of_str(string):
            y = (string.split("-"))
            return y[1], y[2]

        year, quartal = get_quartal_out_of_str(invoicelistline["Abrechnung"])
        exportline["Zahlungsreferenz/Verwendungszweck"] = f"Gemeinwohlenergie Rechung {year} Quartal {quartal}"
        if datatype == "debit":
            return exportline, matchingmandate
        else:
            return exportline


    serieslist = []
    missingmandates = []
    for index, line in debit.iterrows():
        exportline,matchingmandate = create_one_line_debit(line,creditor_ID,mandates,datatype="debit")
        if exportline is not None:
            serieslist.append(exportline)
            if not matchingmandate:
                missingmandates.append(f"{line['Empfänger Vorame']} {line['Empfänger Nachname']}")
    debitexport = pd.concat(serieslist, axis=1).T

    serieslist = []
    for index, line in transfer.iterrows():
        exportline = create_one_line_debit(line,creditor_ID,mandates,datatype="transfer")
        if exportline is not None:
            serieslist.append(exportline)
    transferexport = pd.concat(serieslist, axis=1).T

    reply = QMessageBox.question(None,
        'Frage',
        'Willst du die einzelne Positionen zusammenfassen??',
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.No
    )

    # Check the user's response
    if reply == QMessageBox.Yes:
        print("Merge single positions")

        def group_and_aggregate(df, group_cols, sum_cols):
            # Create a dictionary for aggregation:
            # - 'sum' for the columns you want to sum
            # - 'first' for the other columns
            agg_dict = {col: 'sum' for col in sum_cols}  # Sum columns
            other_cols = [col for col in df.columns if col not in group_cols + sum_cols]
            agg_dict.update({col: 'first' for col in other_cols})  # First entry for the rest

            grouped_df = df.groupby(group_cols, as_index=False).agg(agg_dict)

            return grouped_df
        debitexport = group_and_aggregate(debitexport,["Zahlungspflichtiger Name"],['Betrag in EUR'])
        transferexport = group_and_aggregate(transferexport,["Empfänger Name"],['Betrag in EUR'])

        doubles_names = debitexport[debitexport['Zahlungspflichtiger Name'].isin(transferexport['Empfänger Name'])]['Zahlungspflichtiger Name']
        for double_name in doubles_names:
            print(double_name)
            diff = debitexport.loc[debitexport['Zahlungspflichtiger Name'] == double_name, 'Betrag in EUR'].values[0] -transferexport.loc[transferexport['Empfänger Name'] == double_name, 'Betrag in EUR'].values[0]
            # if diff > 0  --> more debit than transfer
            if diff > 0:
                print("We have more debit than transfer")
                debitexport.loc[debitexport['Zahlungspflichtiger Name'] == double_name, 'Betrag in EUR'] = diff
                transferexport = transferexport.loc[~(transferexport['Empfänger Name'] == double_name),:].reset_index(drop=True)

            else:
                print("We have more transfer than debit")
                transferexport.loc[transferexport['Empfänger Name'] == double_name, 'Betrag in EUR'] = -diff
                debitexport = debitexport.loc[~(debitexport['Zahlungspflichtiger Name'] == double_name),:].reset_index(drop=True)






    else:
        print("donot merge single positions.")
    debitexport["Betrag in EUR"] = debitexport["Betrag in EUR"].apply(lambda x: f"{x:.2f}".replace('.', ','))
    transferexport["Betrag in EUR"] = transferexport["Betrag in EUR"].apply(lambda x: f"{x:.2f}".replace('.', ','))


    return debitexport,transferexport, missingmandates, doublesprocess

#    fullname = f"{personaldata['Name 1']} {personaldata['Name 2']}"
#personaldata =  masterdata.data.iloc[0]
# invoicedata = invoices.data['detailed'][invoices.data['detailed']['Empfänger Name'] == fullname]
#invoicetemplate = invoices.template_for_export

def produce_invoices_and_save(energydata,invoicedata,invoicetemplate,savedirfp,callback,finished):
    # fullname = f"{personaldata['Name 1']} {personaldata['Name 2']}"
    debit,transfer, doublesprocess = check_doubles(invoicedata)
    # debit, transfer = debit[cols_for_tbl], transfer[cols_for_tbl]
    nr_all_persons = invoicedata["Empfänger Name"].unique().shape[0]
    for index,name in enumerate(invoicedata["Empfänger Name"].unique()):
        print(f"({index+1}/{nr_all_persons}) Make invoice for {name}")
        callback(f"({index+1}/{nr_all_persons}) Ich mache die Rechnung für {name}...")
        invoicethis = invoicedata[invoicedata["Empfänger Name"] == name]
        print(invoicethis)
        debits_this =  debit[debit["Empfänger Name"] == name]
        transfers_this = transfer[transfer["Empfänger Name"] == name]
        total_transfer_debit = "Rechnung"
        #if diff > 0 debit bigger than transfer -> Rechnung otherwise Gutschrift
        diff = debits_this["Pos. Bruttobetrag"].sum(axis = 0) -transfers_this["Pos. Bruttobetrag"].sum(axis = 0)
        if diff > 0:
            total_transfer_debit = "Rechnung"
            transfers_this.loc[:,['Pos. Nettobetrag','Pos. Bruttobetrag']]= -transfers_this.loc[:,['Pos. Nettobetrag','Pos. Bruttobetrag']]

        if diff < 0:
            total_transfer_debit = "Gutschrift"
            debits_this.loc[:,['Pos. Nettobetrag','Pos. Bruttobetrag']]= -debits_this.loc[:,['Pos. Nettobetrag','Pos. Bruttobetrag']]




        all_position_for_tbl = pd.concat([debits_this,transfers_this], axis = 0)
        postext =  all_position_for_tbl.copy()
        postext.loc[postext['Pos. Text'].str.contains('Mitgliedsgebuer'),'Pos. Text'] = ''
        all_position_for_tbl['Pos. Invoicetext'] = (all_position_for_tbl['Pos. Tarif'] + "\n"+ postext['Pos. Text'])
        all_position_for_tbl['Pos. Preis / Einheit'] = all_position_for_tbl['Pos. Preis / Einheit'].astype(float)

        all_position_for_tbl['Pos. Menge'] = all_position_for_tbl['Pos. Menge'].astype(str) + " " + all_position_for_tbl[ 'Pos. Mengeneinheit'].fillna("")
        all_position_for_tbl.loc[all_position_for_tbl['Pos. Preis / Einheit Euro/Cent'] == "Ct",'Pos. Preis / Einheit'] = all_position_for_tbl.loc[all_position_for_tbl['Pos. Preis / Einheit Euro/Cent'] == "Ct",'Pos. Preis / Einheit']/ 100
        sumnetto = str(round(all_position_for_tbl['Pos. Nettobetrag'].sum(),2))
        sumbrutto = str(round(all_position_for_tbl['Pos. Bruttobetrag'].sum(),2))


        cols_for_tbl = ['Pos. Invoicetext','Pos. Menge','Pos. Preis / Einheit','Pos. Nettobetrag','Pos. UST %','Pos. Bruttobetrag']
        moneycols = ['Pos. Preis / Einheit','Pos. Nettobetrag','Pos. Bruttobetrag']
        all_position_for_tbl[moneycols] = all_position_for_tbl[moneycols].map("{0:.2f}".format)
        all_position_for_tbl = all_position_for_tbl[cols_for_tbl]

        invoicequart = invoicethis["Abrechnung"].iloc[0]
        year,quart = invoicequart.split("-")[-2],invoicequart.split("-")[-1]

        parsing_dict = {}
        parsing_dict["EmpfängerName"] = invoicethis["Empfänger Name"].values[0]
        parsing_dict["EmpfängerAdresse1"] =invoicethis["Empfänger Adresse 1"].values[0]
        parsing_dict["EmpfängerAdresse2"] =invoicethis["Empfänger Adresse 2"].values[0]
        parsing_dict["invoiceitemstbl_contents"]= all_position_for_tbl.values.tolist()
        parsing_dict["Rechnung_Gutschrift"] = total_transfer_debit
        invoicenumberstr = ""
        if total_transfer_debit == "Rechnung":
            invoicenumberstr = debits_this["Nummer"].iloc[0]
            parsing_dict["Überweisungstext"] = "Die gegenständliche Rechnungsforderung wird vereinbarungsgemäß von Ihrem Konto eingezogen."
        if total_transfer_debit == "Gutschrift":
            invoicenumberstr = transfers_this["Nummer"].iloc[0]
            parsing_dict["Überweisungstext"] = "Die gegenständliche Gutschrift wird vereinbarungsgemäß auf Ihr Konto überwiesen."

        parsing_dict["Rechnungsnummer"] = invoicenumberstr
        parsing_dict["DatumHeute"] = dt.datetime.today().strftime("%d.%m.%Y")
        parsing_dict["Jahr"] = year
        parsing_dict["Quartal"] = quart
        parsing_dict["TotalSumNetto"] = sumnetto
        parsing_dict["TotalSumBrutto"] = sumbrutto



        # print(parsing_dict)
        image_stream = BytesIO()
        fig, axs = plt.subplots(2)
        # prepare data for plot
        # check whether the energy direction is always the same
        try:
            energydata.loc[:, pd.IndexSlice[:, name, :, :]]
        except:
            print("This Name is not in the Energydata try adding a whitespace")
            name += " "
        energydirections = energydata.xs(name, level="Name", axis=1).columns.get_level_values(level="Energy direction")
        meteringpointids = energydata.xs(name, level="Name", axis=1).columns.get_level_values(level="MeteringpointID")
        for energydirection, meteringpointid in zip(energydirections.unique(), meteringpointids.unique()):
            # print(energydirection, meteringpointid)

            if energydirection == "GENERATION":
                total_energy = energydata.loc[:, pd.IndexSlice[:, name, energydirection,
                                                             "Gesamte gemeinschaftliche Erzeugung [KWH]"]].sort_index()

                energy_through_evu = energydata.loc[:, pd.IndexSlice[:, name, energydirection,
                                                                  "Gesamt/Überschusserzeugung, Gemeinschaftsüberschuss [KWH]"]].sort_index()
                energy_through_eg = pd.DataFrame((np.nan_to_num(total_energy.values,0 )- np.nan_to_num(energy_through_evu.values,0)),index=total_energy.index)

                plottext1  = "Energielieferung an außerhalb der Energiegemeinschaft"
                plottext2  = "Energielieferung über unsere Energiegemeinschaft"
                totalsumeg = energy_through_eg.sum().values[0]
                shareeg = totalsumeg/total_energy.sum().values[0]*100
                parsing_dict["TextfürVerbrauch"] = f"Insgesamt wurden {totalsumeg:.1f}kWh an die Energiegemeischaft verkauft. \nDies ist {shareeg:.1f}% der gesamten erzeugten Energie."
            else:
                total_energy= energydata.loc[:, pd.IndexSlice[:, name, energydirection,
                                                             "Gesamtverbrauch lt. Messung (bei Teilnahme gem. Erzeugung) [KWH]"]].sort_index()
                energy_through_eg = energydata.loc[:, pd.IndexSlice[:, name, energydirection,
                                                                  "Eigendeckung gemeinschaftliche Erzeugung [KWH]"]].sort_index()
                energy_through_evu = pd.DataFrame((np.nan_to_num(total_energy.values,0 )- np.nan_to_num(energy_through_eg.values,0)), index=total_energy.index)

                plottext1  = "Energiebezug von Stromlieferant"
                plottext2  = "Energiebezug über unsere Energiegemeinschaft"
                totalsumeg = energy_through_eg.sum().values[0]
                shareeg = totalsumeg/total_energy.sum().values[0]*100
                parsing_dict["TextfürVerbrauch"] = f"Insgesamt wurden {totalsumeg:.1f}kWh über die Energiegemeischaft bezogen. \nDies ist {shareeg:.1f}% des Gesamtenergieverbrauchs."
            # dailysums_total_energy = total_energy_consumption.groupby(total_energy_consumption.index.strftime('%d.%m.%Y')).sum()

            energy_through_evu_weeklysum = energy_through_evu.groupby(energy_through_evu.index.strftime('%Y-%W')).sum()
            xaxis_energy_through_evu = []
            xaxis_energy_through_eeg = []
            for week, weektotalenergy in energy_through_evu.groupby(energy_through_evu.index.strftime('%Y-%W')):
                xaxis_energy_through_evu.append(dt.datetime.strptime(f"{week}-1","%Y-%U-%w"))#the one saying we take the monday of the week
            energy_through_eeg_weeklysum = energy_through_eg.groupby( energy_through_eg.index.strftime('%Y-%W')).sum()
            for week, weekenergy_through_eg in energy_through_eg.groupby(energy_through_eg.index.strftime('%Y-%W')):
                xaxis_energy_through_eeg.append(dt.datetime.strptime(f"{week}-1","%Y-%U-%w"))

            # throuw away the first and last week since they are only partly and will change the sum
            energy_through_evu_weeklysum = energy_through_evu_weeklysum.iloc[1:-1]
            energy_through_eeg_weeklysum = energy_through_eeg_weeklysum.iloc[1:-1]
            xaxis_energy_through_evu = xaxis_energy_through_evu[1:-1]
            xaxis_energy_through_eeg = xaxis_energy_through_eeg[1:-1]

            dailyhourmean_energy_through_evu = energy_through_evu.groupby(energy_through_evu.index.hour).mean()
            dailyhourmean_through_eeg = energy_through_eg.groupby(
                energy_through_eg.index.hour).mean()
            barwidth = dt.timedelta(days=3)
            xaxis_energy_through_evu = [x - 0.5* barwidth for x in xaxis_energy_through_evu]
            xaxis_energy_through_eeg = [x + 0.5 * barwidth for x in xaxis_energy_through_eeg]
            axs[0].bar(xaxis_energy_through_evu, energy_through_evu_weeklysum.values.T[0], width=barwidth,
                       label=plottext1,
                       color="#69a4dc")
            axs[0].bar(xaxis_energy_through_eeg, energy_through_eeg_weeklysum.values.T[0], width=barwidth,
                       label=plottext2,
                       color="#f8ae42")
            axs[0].xaxis.set_major_locator(MonthLocator())
            axs[0].xaxis.set_major_formatter(DateFormatter('%b %Y'))
            # axs[0].xaxis.set_label_position("right")
            axs[0].set_ylabel("kWh in der Woche")
            axs[0].legend(loc=1)
            # Creating the secondary x-axis
            # ax2 = axs[0].twiny()
            # Setting the secondary x-axis ticks as month labels
            # ax2.set_xlim(axs[0].get_xlim())  # Make sure both x-axes have the same range
            # ax2.set_xticks(dailysums_total_energy.index.values)  # Set tick positions on the secondary axis
            # ax2.set_xticklabels([date.strftime('%b') for date in dailysums_total_energy.index.values])  # Format as month names (e.g., Jan, Feb, etc.)
            # ax2.set_xlabel('Month')

            barwidth = 0.4
            axs[1].bar(dailyhourmean_energy_through_evu.index - 0.5 * barwidth,
                       dailyhourmean_energy_through_evu.values.T[0], width=barwidth, color="#69a4dc")
            axs[1].bar(dailyhourmean_through_eeg.index + 0.5 * barwidth, dailyhourmean_through_eeg.values.T[0],
                       width=barwidth, color="#f8ae42")
            axs[1].set_xticks([0, 6, 12, 18, 24])
            axs[1].set_xticklabels(["0 Uhr", "6 Uhr", "12 Uhr", "18 Uhr", "24 Uhr"])
            axs[1].set_ylabel("Durchschnittliche Leistung \nzur Tageszeit [kW]")


        invoicetemplate.render(parsing_dict)
        # fig.set_size(3.49, 1.97)
        fig.tight_layout(rect=(0.05, 0.1, 1, 0.8))
        plt.savefig(image_stream, format="png")
        # plt.show()

        plt.close()
        image_stream.seek(0)

        invoicetemplate.replace_pic("Image2", image_stream)



        #save
        namefile = f"Rechnung_{year}_q{quart}_{invoicethis["Empfänger Vorame"].iloc[0]}_{invoicethis["Empfänger Nachname"].iloc[0]}"
        savepathdocx = os.path.join(savedirfp, f"{namefile}.docx")
        print(f"Save invoice of {name} to {savepathdocx}")
        invoicetemplate.save(savepathdocx)

        def generate_pdf(doc_path, path):

            subprocess.call(['soffice',
                             # '--headless',
                             '--convert-to',
                             'pdf',
                             '--outdir',
                             path,
                             doc_path])
            return doc_path

        # print(f"Save invoice of {name} to {os.path.join(savedirfp, f'{namefile}.pdf')}")
        print(f"Save invoice of {name} to {savedirfp}")

        generate_pdf(savepathdocx, savedirfp)
    finished()
    return invoicetemplate

# base_dir = current_directory = os.getcwd()
# parent_dir = os.path.dirname(base_dir)
# supparentdir = os.path.dirname(parent_dir)
# template_path = os.path.join(parent_dir,"Vorlage.docx")
# excel_template_path = os.path.join(parent_dir,"Jahresübersicht_Vorlage.xlsx")
# allhourdata_path = os.path.join(parent_dir ,"Stundendaten.xlsx")
# allclientdata_path = os.path.join(parent_dir ,"PatientInneninformationen.xlsx")
#
# outputdir_path = os.path.join(supparentdir,f"{dt.datetime.now().year}")
#
# if not os.path.isdir(outputdir_path):
#     os.mkdir(outputdir_path)
# else:
#     print(f"We already have a directory {outputdir_path}")
#
#
# outputfile_path = os.path.join(outputdir_path, f"RE {1} {dt.date.today().strftime('%d_%m_%Y')}.docx")
#
# # invoicedata has to be a dict with keys which are the same as the placeholders in the template
# # input the client data  in word
# doc = DocxTemplate(template_path)
# doc.render(invoicedata)
# doc.save(outputfile_path)
#
# ## input the hour table in word
# doc = Document(outputfile_path)
# doc.tables #a list of all tables in document
# # table nr. 0 is the data table and table nr. 1 is the sum table
#
# # change a table in the template
# wordtable = pd.concat([namehourdata["Datum"].apply(lambda x: x.strftime("%d.%m.%Y")), namehourdata["Minuten"].apply(lambda x: str(x) + " min"), amountpersession.apply(lambda x: "%0.2f" % x + " €") ], axis=1)
# print("Die Stunden sind: " )
# print(wordtable)
#
#
# # insert the table in the Word document
# for index, row in wordtable.iterrows():
#     hourdatatable = doc.tables[0]   #so hourdatatable is the first table in the document
#     data_row = hourdatatable.add_row().cells
#     for i,(name,entry) in enumerate(row.items()):
#             data_row[i].text = entry
# #format it
# for row in doc.tables[0].rows:
#     row.height = Cm(0.8)
#     row.alignment = WD_TABLE_ALIGNMENT.CENTER
#
#
#
# #insert total amount into tables[1]
# totalamount = sum(np.array(amountpersession))
# totalamountstring = (str(totalamount)+"0").replace(".",",")
# doc.tables[1].cell(0, 2).text = str(totalamount) + "0" + " €"
#
# doc.save(outputfile_path)
#
#
# os.startfile(outputfile_path)