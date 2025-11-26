import pandas as pd
from lxml import etree
import matplotlib.pyplot as plt
import matplotlib
import datetime as dt
import numpy as np
import glob
# from pandas.conftest import names

matplotlib.use('Qt5Agg')

# marlene = pd.read_xml("/home/leander/gei/faktura/test_data/AT0050000000000000000000000407441__AT005000202409121819299710489012233.xml")
# Define namespaces based on the XML declaration
namespaces = {
    'ns0': 'http://www.ebutilities.at/schemata/customerprocesses/consumptionrecord/01p40',
    'ns1': 'http://www.ebutilities.at/schemata/customerprocesses/common/types/01p20'
}

# Path to your XML file
file_paths = ['/home/leander/gei/faktura/abrechnung_24_q3/AirG_CC100438_DATEN_CRMSG_462009221.xmlv',
              '/home/leander/gei/faktura/test_data/AT0051000602000000000000000020341_marcell_CC100438_DATEN_CRMSG_449659721.xml',
              '/home/leander/gei/faktura/test_data/AT0051000607200000000000000146104_gerhard_CC100438_DATEN_CRMSG_449659921.xml',
              '/home/leander/gei/faktura/test_data/marlene_AT005000202409161804252160490931709.xml'

              ]

file_paths =glob.glob('/home/leander/gei/faktura/abrechnung_24_q3/*')
file_paths.append("/home/leander/gei/faktura/test_data/marlene_AT005000202409161804252160490931709.xml")
names = ["AirG","Marcell","Rosmarie","Leander","Gerhard","Diele","Marlene"]
metercodes = {"1-1:2.9.0 G.01" : "QH Erzeugung laut Messung",
              "1-1:2.9.0 G.01T": "QH Erzeugung laut Messung entsprechend dem Teilnahmefaktor bei der EG und je ZP" ,
              "1-1:2.9.0 P.01" : "QH Lastgang ¼ h Wirkenergiewerte, Bezug vom Endkunden, Gemeinschaftsüberschuss ",
              "1-1:2.9.0 P.01T": "QH Restnetzüberschuss bei Energiegemeinschaft ",
              "1-1:2.9.0 G.02": "QH Anteil an der Erzeugung ",
              "1-1:1.9.0 G.01" : "QH Verbrauch laut Messung ",
              "1-1:2.9.0 G.03" : "QH Eigendeckung" }

def read_energy_data_xml(file_path):
    # Parse the XML file
    tree = etree.parse(file_path)
    root = tree.getroot()

    # Find all <ns0:EnergyData> elements
    dataframes = {}
    for metercode in metercodes:
        datatype =  metercodes[metercode]
        starttime = root.findall(f'.//ns0:EnergyData[@MeterCode="{metercode}"]/ns0:EP/ns0:DTF', namespaces)
        energy_value = root.findall(f'.//ns0:EnergyData[@MeterCode="{metercode}"]/ns0:EP/ns0:BQ', namespaces)

        data = []
        for time,ev in zip(starttime,energy_value):
            data.append({'starttime':time.text,'energy_value': ev.text})

        # Convert list of dictionaries to pandas DataFrame
        df = pd.DataFrame(data)
        if df.size > 0:
            df.index = pd.to_datetime(df['starttime'])
            df = df.energy_value.astype(float)
            # dataframes[datatype] = df
            dataframes[metercode] = df
    return dataframes


persons = {}
for path, name in zip(file_paths,names):
    persons[name] = read_energy_data_xml(path)

plotdata = "1-1:1.9.0 G.01"
def plotdaily(person,plotdata,ax):
    daily  = [x[1] for x in  persons[person][plotdata].resample('D')]
    hour = daily[0].index.hour + daily[0].index.minute/60
    dailykWh = np.array([x.values for x in daily]).mean(axis=0)*4
    ax.plot(hour,dailykWh,label=f"{person} - {plotdata}")
    ax.legend()

def plotall(person,plotdata,ax):
    data = persons[person][plotdata]
    ax.plot(data.index,data,label=f"{person} - {plotdata}")
    ax.legend()



fig,ax = plt.subplots(1)
ax.grid()
ax.set_ylabel("kW")
plotdaily( "AirG","1-1:1.9.0 G.01",ax)
# plotdaily("AirG","1-1:2.9.0 G.02",ax)
# plotdaily("AirG","1-1:2.9.0 G.03",ax)

plotdaily("Gerhard","1-1:2.9.0 G.01T",ax)
plotdaily("Marlene","1-1:2.9.0 G.01T",ax)
plotdaily("Marcell","1-1:1.9.0 G.01",ax)

fig2,ax2 = plt.subplots(1)
ax2.grid()

plotall("Gerhard","1-1:2.9.0 G.01T",ax2)
plotall("Gerhard","1-1:2.9.0 G.01",ax2)

fig3,ax3 = plt.subplots(1)
ax3.grid()

plotall("Gerhard","1-1:2.9.0 G.01T",ax3)
plotall("Marlene","1-1:2.9.0 G.01T",ax3)
plotall("AirG","1-1:2.9.0 G.03",ax3)
plotall("Leander","1-1:2.9.0 G.03",ax3)
plotall("Leander","1-1:1.9.0 G.01",ax3)