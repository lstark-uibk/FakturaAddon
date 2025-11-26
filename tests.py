import pandas as pd
import numpy as np
fp ="/home/leander/gei/Abrechnung/2025 q3/CC100438-EEG-Masterdata-20251112.xlsx"

df = pd.read_excel(fp, header=None, dtype=str)

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
            print(this,next)
        else:
            print("no viable")

        i += 1
