import streamlit as st
import pandas as pd
import sys
import functools
import pip
pip.main(["install", "openpyxl"])
sys.setrecursionlimit(1000)


def lv(s1, s2):
    n = len(s1)
    m = len(s2)

    @functools.lru_cache(maxsize=None)
    def distance(i, j):
        if i == 0 and j == 0:
            return 0
        if i == 0:
            return j
        if j == 0:
            return i

        a = 1
        if s1[i - 1] == s2[j - 1]:
            a = 0

        return min(distance(i, j - 1) + 1, distance(i - 1, j) + 1, distance(i - 1, j - 1) + a)

    return distance(n, m)


def main():
    st.markdown('## Сравнение таблиц')
    f1, f2 = False, False

    our_register = st.file_uploader("Выбрать внутренний реестр", type=['xlsx', "xls"], key=1)
    if our_register is not None:
        f1 = True

    mavr_register = st.file_uploader("Выбрать реестр от МАВРа", type=['xlsx', "xls"], key=2)
    if mavr_register is not None:
        f2 = True

    if st.button('Сравнить', disabled=not f1 or not f2):
        start(our_register, mavr_register)


def start(our_register, mavr_register):
    our_register_df = pd.read_excel(our_register, header=1)
    # st.dataframe(our_register_df)
    # print(our_register_df)

    mavr_register_df = pd.read_excel(mavr_register, header=0)
    # st.dataframe(mavr_register_df)

    last_date = ""
    # our_register_df['Дата подачи'] = pd.to_datetime(our_register_df['Дата подачи'], errors='coerce')
    for i in range(len(our_register_df['Дата подачи'])):
        if our_register_df['Дата подачи'][i].count(".") == 2 and len(our_register_df['Дата подачи'][i]) == 10:
            # print(pd.to_datetime(".".join(our_register_df['Дата подачи'][i].split(".")[::-1])))
            our_register_df['Дата подачи'][i] = pd.to_datetime(".".join(our_register_df['Дата подачи'][i].split(".")[::-1]))
        else:
            our_register_df['Дата подачи'][i] = None

    for i in [1, 4, 7]:
        for j in range(len(mavr_register_df.iloc[:, i])):
            # print(str(mavr_register_df.iloc[:, i][j]))
            date = mavr_register_df.iloc[:, i][j]
            # print(str(date))
            if last_date == "" or str(date) == "nan" or date.isoformat() != "NaT":
                last_date = str(date)
            else:
                mavr_register_df.iloc[:, i][j] = pd.to_datetime(last_date)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    # print(mavr_register_df)

    dataframe1 = mavr_register_df.iloc[:, 3:6].rename({"3": "", "4дата": "дата", "50№вагона": "№вагона"})
    dataframe2 = mavr_register_df.iloc[:, 6:9].rename({"6": "", "7дата": "дата", "8№вагона": "№вагона"})
    mavr_register_df = mavr_register_df.drop(mavr_register_df.columns[[j for j in range(3, 9)]], axis=1)
    mavr_register_df.columns = ["", "Дата", "№ вагона"]
    dataframe1.columns = ["", "Дата", "№ вагона"]
    dataframe2.columns = ["", "Дата", "№ вагона"]
    mavr_register_df = pd.concat([mavr_register_df, dataframe1, dataframe2], ignore_index=True).set_index("").sort_index()
    mavr_register_df = mavr_register_df[pd.isna(mavr_register_df['№ вагона']) == False]
    # mavr_register_df = mavr_register_df[mavr_register_df['№ вагона'] != "NaT"]
    mavr_register_df[["№ вагона"]] = mavr_register_df[["№ вагона"]].astype(int)
    # print(mavr_register_df)

    result_data = {"№ вагона": [], "Дата": [], "Проблема": []}
    year, month, day = map(int, str(mavr_register_df.iloc[0, 0]).split(" ")[0].split("-"))
    # print(day, month, year)
    start_month = (2 - len(str(month - 1))) * "0" + str(month - 1)
    start_year = str(year)
    if start_month == "00":
        start_month = "12"
        start_year = str(int(year) - 1)
    start_date = f"{start_year}-{start_month}-15"
    start_date = pd.to_datetime(f"{start_date}")
    # print(start_date)
    #
    # print(pd.to_datetime("02.05.2022"))
    # print(pd.to_datetime(".".join("02.05.2022".split(".")[::-1])))

    # our_register_df['Дата подачи'] = pd.to_datetime(our_register_df['Дата подачи'], errors='coerce')

    our_register_df = our_register_df.loc[:, ['№ вагона', "Дата подачи"]]
    # print(our_register_df)
    filter1 = our_register_df['Дата подачи'] > start_date
    our_register_df = our_register_df[filter1]
    # print(our_register_df)
    our_register_df = our_register_df[pd.isna(our_register_df['№ вагона']) == False]
    our_register_df = our_register_df[pd.isna(our_register_df['Дата подачи']) == False]
    used = []
    for i in range(len(mavr_register_df["№ вагона"])):
        # print(1)
        mavr_register_df["№ вагона"].iloc[i] = str(mavr_register_df["№ вагона"].iloc[i]).strip()
        # if mavr_register_df["№ вагона"].iloc[i] == "":
        #     continue
        closest = [100, "", ""]
        for j in range(len(our_register_df["№ вагона"])):
            # print(str(our_register_df["№ вагона"].iloc[j]).strip(), str(mavr_register_df["№ вагона"].iloc[i]).strip())
            if lv(str(our_register_df["№ вагона"].iloc[j]).strip(), str(mavr_register_df["№ вагона"].iloc[i]).strip()) < closest[0]:
                # print(str(our_register_df["№ вагона"].iloc[j]).strip(), str(mavr_register_df["№ вагона"].iloc[i]).strip())
                closest[0] = lv(str(our_register_df["№ вагона"].iloc[j]).strip(), str(mavr_register_df["№ вагона"].iloc[i]).strip())
                closest[1] = str(our_register_df["№ вагона"].iloc[j]).strip()
                closest[2] = str(our_register_df["Дата подачи"].iloc[j])
        if closest[0] > 0:
            result_data["№ вагона"].append(mavr_register_df["№ вагона"].iloc[i])
            result_data["Дата"].append(mavr_register_df["Дата"].iloc[i])
            result_data["Проблема"].append(f'Номера нет в таблице Ближаший номер: {closest[1]} Дата подачи: {closest[2].replace("-", ".")[:-9]}')
        if not pd.isna(mavr_register_df["№ вагона"].iloc[i]) and mavr_register_df["№ вагона"].value_counts()[mavr_register_df["№ вагона"].iloc[i]] > 1 and str(mavr_register_df["№ вагона"].iloc[i]) not in used:
            used.append(str(mavr_register_df["№ вагона"].iloc[i]))
            result_data["№ вагона"].append(mavr_register_df["№ вагона"].iloc[i])
            result_data["Дата"].append(mavr_register_df["Дата"].iloc[i])
            result_data["Проблема"].append(f'Номер встретился больше 1 раза')

    res = pd.DataFrame(result_data)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    # print(res)

    @st.cache
    def convert_df(df):
        return df.to_excel("Результат.xlsx")

    convert_df(res)
    with open("Результат.xlsx", "rb") as file:
        st.download_button(
            label="Скачать результат",
            data=file,
            file_name="Результат.xlsx",
            mime="table/xlsx"
        )

    st.warning("Успех!\nФайл успешно создан.")

    """
    Вагона нет/номер вагона не совподает
    Задвоение вагона
    """


if __name__ == '__main__':
    main()

