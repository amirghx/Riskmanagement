import os
import pandas as pd

def Diff(li1, li2):
    li_dif = [i for i in li1 + li2 if i not in li1 or i not in li2]
    return li_dif

df_1 = pd.read_excel(r"C:\Users\Amgh\PycharmProjects\Riskmanagement\data\states.xlsx",usecols=[1])
ref = df_1.values.tolist()
print(ref)
df_2 = pd.read_excel(r"C:\Users\Amgh\PycharmProjects\Riskmanagement\Finaldata\shakhes_final.xlsx", usecols=[0])
cprs = df_2.values.tolist()

print(Diff(ref,cprs))

