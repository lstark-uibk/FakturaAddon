import pandas as pd
from lxml import etree
import matplotlib.pyplot as plt
import matplotlib
import datetime as dt
import numpy as np
# from pandas.conftest import names

matplotlib.use('Qt5Agg')

# marlene = pd.read_xml("/home/leander/gei/faktura/test_data/AT0050000000000000000000000407441__AT005000202409121819299710489012233.xml")
# Define namespaces based on the XML declaration
namespaces = {
    'ns0': 'http://www.ebutilities.at/schemata/customerprocesses/consumptionrecord/01p40',
    'ns1': 'http://www.ebutilities.at/schemata/customerprocesses/common/types/01p20'
}

# Path to your XML file
file_paths = ['/home/leander/gei/faktura/test_data/130046_Leander_CC100438_DATEN_CRMSG_449659521.xml',
              '/home/leander/gei/faktura/test_data/AT0051000602000000000000000020341_marcell_CC100438_DATEN_CRMSG_449659721.xml',
              '/home/leander/gei/faktura/test_data/AT0051000607200000000000000146104_gerhard_CC100438_DATEN_CRMSG_449659921.xml',
              '/home/leander/gei/faktura/test_data/marlene_AT005000202409161804252160490931709.xml'

              ]
names = ["leander","marcell","gerhard","marlene"]
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
            dataframes[datatype] = df
            # dataframes[metercode] = df
    return dataframes


persons = {}
for path, name in zip(file_paths,names):
    persons[name] = read_energy_data_xml(path)
# mache weiter
# for name in names:
#     fig,ax = plt.subplots(1)
#     data = persons[name]
#     for line in data:
#         ax.plot(data[line].index,data[line],label = f"{name}-{line}")
#     ax.legend()
# toplot= ['QH Restnetzüberschuss bei Energiegemeinschaft ']
toplot= ['QH Verbrauch laut Messung ','QH Restnetzüberschuss bei Energiegemeinschaft ']
dirs = ["Verbrauch","Einspeisung"]
filtereddata = {"data":[],
                "datatype":np.array([]),
                "dir":np.array([]),
                "name":np.array([])}
for name in names:
    data = persons[name]
    for type in data:

        if type == 'QH Verbrauch laut Messung 'or type == "QH Eigendeckung" :
            filtereddata["data"].append(data[type])
            filtereddata["datatype"] = np.append(filtereddata["datatype"],type)
            filtereddata["name"] = np.append(filtereddata["name"],name)
            filtereddata["dir"] = np.append(filtereddata["dir"],"Verbrauch")
        if type == 'QH Restnetzüberschuss bei Energiegemeinschaft 'or type == "QH Erzeugung laut Messung entsprechend dem Teilnahmefaktor bei der EG und je ZP":
            filtereddata["data"].append(data[type])
            filtereddata["datatype"] = np.append(filtereddata["datatype"],type)
            filtereddata["name"] = np.append(filtereddata["name"],name)
            filtereddata["dir"] = np.append(filtereddata["dir"],"Einspeisung")
filtereddata["data"].append(persons["gerhard"]["QH Erzeugung laut Messung"])
filtereddata["datatype"] = np.append(filtereddata["datatype"],"QH Erzeugung laut Messung")
filtereddata["name"] = np.append(filtereddata["name"], "gerhard")
filtereddata["dir"] = np.append(filtereddata["dir"], "Einspeisung")


# fig, ax = plt.subplots(1)
# for x in range(0,8):
#     line = "--"
#     if (x % 2) == 0:
#         line = "-"
#     plt.plot(filtereddata["data"][x].index,filtereddata["data"][x],label=f"{filtereddata['name'][x]} - {filtereddata['datatype'][x]}",linestyle=line)
# plt.legend()

sums = {"Verbrauch": filtereddata["data"][1]+filtereddata["data"][3],
"Einspeisung":filtereddata["data"][4]+filtereddata["data"][6]
        }


fig,axs = plt.subplots(3, sharex=True)

for ind in range(0,4):
    color = f"C{ind}"
    # mask = (filtereddata["name"]== name) & ((filtereddata["datatype"] =="QH Erzeugung laut Messung entsprechend dem Teilnahmefaktor bei der EG und je ZP" )|  (filtereddata["datatype"] =="QH Eigendeckung" ))
    for y in range(0,2):
        z = ind*2 + y
        data = [y for i,y in enumerate(filtereddata["data"]) if i==z][0]
        dir = filtereddata["dir"][z]
        type = filtereddata["datatype"][z]
        name = filtereddata["name"][z]
        linestyle = "-"
        if z %2 ==0 :
            linestyle = "--"
        axs[0].plot(data.index, data, label=f"{type}: {name}",linestyle =linestyle,color = color)

        totalperday = data.resample('D').sum()
        shift = dt.timedelta(minutes=120)
        if z in [1,3,4,6]:
            axs[2].bar(totalperday.index+dt.timedelta(hours=6)+(shift*ind),totalperday,width = shift,color = color)
        # for x, y in zip(totalperday.index+dt.timedelta(hours=5)+(shift*ind),totalperday):
        #     plt.text(x, y+10, f"{round(y,1)}", ha='left', va='bottom', rotation=90,fontsize=8)

color = f"C2"
z = 8
data = [y for i,y in enumerate(filtereddata["data"]) if i==z][0]
dir = filtereddata["dir"][z]
type = filtereddata["datatype"][z]
name = filtereddata["name"][z]
linestyle = ":"
axs[0].plot(data.index, data, label=f"{type}: {name}", linestyle=linestyle, color=color)
    # fig.suptitle(path)
plt.yscale('linear')
axs[0].legend(bbox_to_anchor=(0.6, 0.3))

for dir in dirs:
    ind+=1
    color = f"C{ind}"
    if dir == "Verbrauch":
        text = f"Total {dir} matched"
    else:
        text = f"Total {dir} EEG und EVU"
    axs[1].plot(sums[dir].index,sums[dir],label = text,color= color)
    totalperday = sums[dir].resample('D').sum()
    axs[2].bar(totalperday.index+dt.timedelta(hours=6)+(shift*ind),totalperday,width = shift,color = color)


for p in axs[2].patches:
    axs[2].annotate(f'{round(p.get_height(),4)}',
                (p.get_x() + p.get_width() / 2., p.get_height()),  # Position of the annotation
                ha='center', va='center',
                xytext=(0, 10), textcoords='offset points',rotation=90,fontsize=8)  # Offset the text slightly above the bar


axs[1].legend()
for x in [0,1,2]:
    axs[x].grid()
    box = axs[x].get_position()
    axs[x].set_position([box.x0, box.y0, box.width * 0.9, box.height])
plt.show()

