import math

import pandas as pd
import numpy as np
import statistics
import os

# define array of price for RAYAn base ETFs
final_agas = []
final_karis = []
final_sarv = []
final_atimes = []
final_Bazr = []
final_Almas = []
# define array of amount for shakhes
final_sakhes = []

# define arraye of price for TADBir base ETFs
final_Asas = []
final_Atlas = []
final_firooze = []
final_kardan = []
final_Ofogh = []
# Find path of current working path
path = os.getcwd()

# Find path of reference date excel

Created_path = path + '\data\states.xlsx'

# Read ref and create new folder for transformed data of each ETF
date_ref = pd.read_excel(Created_path)
date_ref_list = date_ref.values.tolist()
Path_raw_data = path + '\Finaldata'

try:
    os.mkdir(Path_raw_data)
except OSError:
    print("Creation of the directory %s failed" % path)
else:
    print("Successfully created the directory %s " % path)


# Read RAYAn base ETF and make df out of it
def read_and_make_rayan(Name):
    path = os.getcwd()
    Created_path = path + '\data' + Name

    ETF_df = pd.read_excel(Created_path, usecols=[1, 3])
    return ETF_df


# extract data out of shakhes
def ext_shakhes():
    path = os.getcwd()
    Created_path = path + '\data\shakhes.xlsx'
    shakhes_df = pd.read_excel(Created_path, usecols=[1, 5])
    return shakhes_df


# make list out of RAYAn base ETF column
def make_list_out_of_df(df):
    date = df['تاریخ'].to_list()
    price = df["قیمت ابطال خالص ارزش روز"].to_list()
    return date, price


def make_list_out_of_sdf(df):
    date = df['date'].to_list()
    price = df["<CLOSE>"].to_list()
    return date, price


# make Agas file and lablabalab
date, price = make_list_out_of_df(read_and_make_rayan("\Agas details.xlsx"))
Merged = (list(zip(date, price)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_agas.append(date_merged)

agas = pd.DataFrame(final_agas, columns=['date', 'price'])
agas.to_excel(Path_raw_data + "\Agas_final.xlsx", index=False)

# make karis file and lablabalab

date, price = make_list_out_of_df(read_and_make_rayan("\karis details.xlsx"))
Merged = (list(zip(date, price)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_karis.append(date_merged)

karis = pd.DataFrame(final_karis, columns=['date', 'price'])
karis.to_excel(Path_raw_data + "\karis_final.xlsx", index=False)

# make sarv file and lablabalab

date, price = make_list_out_of_df(read_and_make_rayan("\sarv details.xlsx"))
Merged = (list(zip(date, price)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_sarv.append(date_merged)

karis = pd.DataFrame(final_sarv, columns=['date', 'price'])
karis.to_excel(Path_raw_data + "\sarv_final.xlsx", index=False)

# make Atimes file and lablabalab

date, price = make_list_out_of_df(read_and_make_rayan("\Atimes details.xlsx"))
Merged = (list(zip(date, price)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_atimes.append(date_merged)

karis = pd.DataFrame(final_atimes, columns=['date', 'price'])
karis.to_excel(Path_raw_data + "\Atimes_final.xlsx", index=False)

# make Bazr file and lablabalab

date, price = make_list_out_of_df(read_and_make_rayan("\Bazr details.xlsx"))
Merged = (list(zip(date, price)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_Bazr.append(date_merged)

karis = pd.DataFrame(final_Bazr, columns=['date', 'price'])
karis.to_excel(Path_raw_data + "\Bazr_final.xlsx", index=False)

# make Alamas file and lablabalab

date, price = make_list_out_of_df(read_and_make_rayan("\Almas details.xlsx"))
Merged = (list(zip(date, price)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_Almas.append(date_merged)

karis = pd.DataFrame(final_Almas, columns=['date', 'price'])
karis.to_excel(Path_raw_data + "\Almas_final.xlsx", index=False)


# ************************************************************************************************************************************
# make list out of Tadbir  base ETF column

def read_and_make_tadbir(Name):
    path = os.getcwd()
    Created_path = path + '\data' + Name

    ETF_df = pd.read_excel(Created_path, header=5, usecols=[10, 8])
    return ETF_df


def make_list_out_of_Tdf(df):
    date = df['تاریخ'].to_list()
    price = df["(قیمت ابطال (خالص ارزش روز"].to_list()
    return date, price


# make Asas file and lablabalab

date, price = make_list_out_of_Tdf(read_and_make_tadbir("\Asas Detail.xlsx"))
Merged = (list(zip(date, price)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_Asas.append(date_merged)

asas = pd.DataFrame(final_Asas, columns=['date', 'price'])
asas.to_excel(Path_raw_data + "\Asas_final.xlsx", index=False)

# make Atlas file and lablabalab

date, price = make_list_out_of_Tdf(read_and_make_tadbir("\Atlas Detail.xlsx"))
Merged = (list(zip(date, price)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_Atlas.append(date_merged)

Atlas = pd.DataFrame(final_Atlas, columns=['date', 'price'])
Atlas.to_excel(Path_raw_data + "\Atlas_final.xlsx", index=False)

# make Atlas file and lablabalab

date, price = make_list_out_of_Tdf(read_and_make_tadbir("\Firooze Detail.xlsx"))
Merged = (list(zip(date, price)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_firooze.append(date_merged)

Atlas = pd.DataFrame(final_firooze, columns=['date', 'price'])
Atlas.to_excel(Path_raw_data + "\Firooze_final.xlsx", index=False)

# make Kardan file and lablabalab

date, price = make_list_out_of_Tdf(read_and_make_tadbir("\kardan Detail.xlsx"))
Merged = (list(zip(date, price)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_kardan.append(date_merged)

Atlas = pd.DataFrame(final_kardan, columns=['date', 'price'])
Atlas.to_excel(Path_raw_data + "\kardan_final.xlsx", index=False)

# make Ofogh file and lablabalab

date, price = make_list_out_of_Tdf(read_and_make_tadbir("\ofoghmelat Detail.xlsx"))
Merged = (list(zip(date, price)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_Ofogh.append(date_merged)

Atlas = pd.DataFrame(final_Ofogh, columns=['date', 'price'])
Atlas.to_excel(Path_raw_data + "\Ofogh_final.xlsx", index=False)


# final file reader we generate with all the stuff we done in upper part

def get_all():
    path = os.listdir(Path_raw_data)
    return path


for path in get_all():
    df = pd.read_excel(Path_raw_data + "/" + path)

date, amount = make_list_out_of_sdf(ext_shakhes())
Merged = (list(zip(date, amount)))

for Maindate in date_ref_list:
    for date_merged in Merged:
        if Maindate[1] == date_merged[0]:
            final_sakhes.append(date_merged)

karis = pd.DataFrame(final_sakhes, columns=['date', 'price'])
karis.to_excel(Path_raw_data + "\shakhes_final.xlsx", index=False)


def get_all():
    path = os.listdir(Path_raw_data)
    return path


def make_list_out_of_finals(df):
    date = df['date'].to_list()
    price = df["price"].to_list()
    return date, price


def shakhes():
    df = pd.read_excel(Path_raw_data + "\shakhes_final.xlsx", usecols=[0, 1])
    dt_array = [0]
    date, prices = make_list_out_of_finals(df)
    last_month = int((date[-1].split('/'))[1])
    last_12month_diff = []
    for i in range(len(prices)):
        if i + 1 in range(len(prices)):
            dt_array.append(((prices[i + 1] / prices[i]) - 1) * 100)
    hello = []
    for i in range(len(date)):
        if (int((date[i].split('/'))[1]) in range(last_month - 6, last_month + 1) and int(
                (date[i].split('/'))[0]) == int((date[-1].split('/'))[0])) or \
                ((int((date[i].split('/'))[0]) == int((date[-1].split('/'))[0]) - 1) and int(
                    (date[i].split('/'))[1]) in range(last_month + 1, 13)):
            hello.append(date[i])
            last_12month_diff.append(dt_array[i])
    return last_12month_diff, hello


differ_shakhes, date_shakhes = shakhes()


def Diff(li1, li2):
    li_dif = [i for i in li1 + li2 if i not in li1 or i not in li2]
    return li_dif


for path in get_all():
    last_month_re = []
    last_3month_re = []
    last_6month_re = []
    last_12month_re = []
    Two_year_re = []
    Three_year_re = []
    # diffs
    last_month_diff = []
    last_3month_diff = []
    last_6month_diff = []
    last_12month_diff = []
    print(path)

    df = pd.read_excel((Path_raw_data + "/" + path))
    date, prices = make_list_out_of_finals(df)
    dt_array = [0]
    for i in range(len(prices)):
        if i + 1 in range(len(prices)):
            dt_array.append(((prices[i + 1] / prices[i]) - 1) * 100)
    last_month = int((date[-1].split('/'))[1])
    for i in range(len(date)):
        if last_month == int((date[i].split('/'))[1]) and int((date[i].split('/'))[0]) == int((date[-1].split('/'))[0]):
            last_month_diff.append(dt_array[i])
            last_month_re.append(prices[i])
            one_month_re = ((last_month_re[-1] / last_month_re[0]) - 1) * 100

    sharpe_one_m = (one_month_re - 20 / 12) / ((statistics.stdev(last_month_diff)) * (math.sqrt(len(last_month_diff))))
    for i in range(len(date)):
        if int((date[i].split('/'))[1]) in range(last_month - 2, last_month + 1) and int(
                (date[i].split('/'))[0]) == int((date[-1].split('/'))[0]):
            last_3month_diff.append(dt_array[i])
            last_3month_re.append(prices[i])
            three_month_re = ((last_3month_re[-1] / last_3month_re[0]) - 1) * 100
    sharpe_three_m = (three_month_re - 20 / 4) / (
            (statistics.stdev(last_3month_diff)) * (math.sqrt(len(last_3month_diff))))
    for i in range(len(date)):
        if int((date[i].split('/'))[1]) in range(last_month - 6, last_month + 1) and (int(
                (date[i].split('/'))[0]) == int((date[-1].split('/'))[0]) or (int(
            (date[i].split('/'))[0]) - 1) == int((date[-1].split('/'))[0])):
            last_6month_diff.append(dt_array[i])
            last_6month_re.append(prices[i])
            six_month_re = ((last_6month_re[-1] / last_6month_re[0]) - 1) * 100
    sharpe_six_m = (six_month_re - 20 / 2) / ((statistics.stdev(last_6month_diff)) * (math.sqrt(len(last_6month_diff))))
    days_of_year = []
    for i in range(len(date)):
        if (int((date[i].split('/'))[1]) in range(last_month - 6, last_month + 1) and int(
                (date[i].split('/'))[0]) == int((date[-1].split('/'))[0])) or \
                ((int((date[i].split('/'))[0]) == int((date[-1].split('/'))[0]) - 1) and int(
                    (date[i].split('/'))[1]) in range(last_month + 1, 13)):
            days_of_year.append(date[i])
            last_12month_diff.append(dt_array[i])
            last_12month_re.append(prices[i])
            one_year_re = ((last_12month_re[-1] / last_12month_re[0]) - 1) * 100
    sharpe_one_y = (one_year_re - 20) / ((statistics.stdev(last_12month_diff)) * (math.sqrt(len(last_12month_diff))))
    days_of_year = list(dict.fromkeys(days_of_year))

    if len(differ_shakhes) == len(last_12month_diff):
     beta = (np.cov(last_12month_diff, differ_shakhes)[0][1])/statistics.variance(differ_shakhes)
     print(np.cov(last_12month_diff, differ_shakhes)[0][1])
     print(statistics.variance(differ_shakhes))
    tr = ((one_year_re)/100-2/10)/beta

    df.insert(2, 'one month stdev', statistics.stdev(last_month_diff), True)
    df.insert(3, 'three month stdev', statistics.stdev(last_3month_diff), True)
    df.insert(4, 'six month stdev', statistics.stdev(last_6month_diff), True)
    df.insert(5, 'one year stdev', statistics.stdev(last_12month_diff), True)
    df.insert(6, 'one month return', one_month_re, True)
    df.insert(7, 'three month return', three_month_re, True)
    df.insert(8, 'six month return', six_month_re, True)
    df.insert(9, 'one year return', one_year_re, True)
    df.insert(10, 'sharpe one month', sharpe_one_m, True)
    df.insert(11, 'sharpe three month', sharpe_three_m, True)
    df.insert(12, 'sharpe six month', sharpe_six_m, True)
    df.insert(13, 'sharpe one year', sharpe_one_y, True)
    df.insert(14, 'beta', beta, True)
    df.insert(15, 'treynor ratio', tr, True)
    df.insert(2, "differ", dt_array, True)
    df.to_excel((Path_raw_data + "/" + path), index=False)
    print("*************************************************************************************")
