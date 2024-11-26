# %%
import pandas as pd
import calendar
import numpy as np
from datetime import date
import locale
import traceback
locale.setlocale(locale.LC_TIME, 'es_ES.utf8') 

# %%
def calculate_age_ranges(year,month,bins,labels):
    try:
        dt = date(year,month,1)
        month_name = calendar.month_name[dt.month]
        file_path = f"{month_name} {year}.ods"
        sheet_name = "mes" 
        df = pd.read_excel(file_path, engine="odf", sheet_name=sheet_name,dtype=str)
        days = calendar.monthrange(year, month)[1]
        ages = []
        for i in range(days):
            li=7+43*i
            lf=li+32
            df_day = df.iloc[li:lf,7].dropna()
            ages.append(df_day)
        df_ages = pd.concat(ages).reset_index(drop=True).astype(int)
        df_ages.name="edad"
        group=pd.cut(df_ages, bins=bins, labels=labels, right=False)
        info = group.value_counts().reset_index()
        return info
    except Exception as e:
        print(month_name)
        print(traceback.format_exc())


# %%
if __name__=="__main__":
    year = 2024
    month = 8
    bins = [0, 12, 19, 31, 61, np.inf]
    labels = ['0-11', '12-18', '19-30', '30-60', '60-']
    calculate_age_ranges(year,month,bins,labels).to_excel(f"{date.today()}.xlsx")


