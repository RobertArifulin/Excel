from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox
from PIL import ImageTk, Image
import pandas as pd
from docx import Document
import Levenshtein as lv

W1_MIN_WIDTH = 600
W1_MIN_HEIGHT = 400

W1_FONT = 'Times 13'
W2_FONT = 'Times 16'
TOUR_FONT = 'J 16'

FRAME_PADX = 5
FRAME_PADY = 15
GRID_PADX = 5
GRID_PADY = 10

BG_COLOR = "white"
WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
RED = (255, 0, 0)
GREEN = (0, 255, 0)
BLUE = (0, 0, 255)

LB_HEIGHT = 5


class StartWindow(Tk):

    def __init__(self):
        super().__init__()

        self.selected_tables = ["", ""]

        self.text_var_1 = StringVar()
        self.text_var_2 = StringVar()
        self.text_var_1.set("Не выбрано")
        self.text_var_2.set("Не выбрано")

        self.create_ui()

    def create_ui(self):
        """Настраивает окно, создает и настраивет все виджеты."""

        # Блок 1
        # задаются параметры окна
        w = self.winfo_screenwidth() // 2
        h = self.winfo_screenheight() // 2
        w = w - W1_MIN_WIDTH // 2
        h = h - W1_MIN_HEIGHT // 2
        self.minsize(W1_MIN_WIDTH, W1_MIN_HEIGHT)
        self.geometry(f'{W1_MIN_WIDTH}x{W1_MIN_HEIGHT}+{w}+{h}')
        self.title("Сравнение Таблиц")

        # ______________________________
        # Блок 1
        # создание контейнера frame1
        frame1 = Frame(self, background=BG_COLOR)
        frame1.pack(expand=True, fill=BOTH)
        for i in range(3):
            frame1.grid_rowconfigure(i, weight=1)
        frame1.grid_columnconfigure(2, weight=1)

        # создание надписи "Выбрать Ботов"
        table_1_label = Label(frame1, text="Выбрать внутренний реестр",
                              font=W1_FONT, background=BG_COLOR)
        table_1_label.grid(row=1, column=0, padx=GRID_PADX, pady=GRID_PADY, sticky=E)

        # создание надписи "Удалить Бота"
        table_2_label = Label(frame1, text="Выбрать реест от МАВРа",
                              font=W1_FONT, background=BG_COLOR)
        table_2_label.grid(row=2, column=0, padx=GRID_PADX, pady=GRID_PADY, sticky=E)

        # импорт картинки проводника
        image = Image.open('images/Windows_Explorer_Icon.png')
        image = image.resize((32, 32))
        win_explorer = ImageTk.PhotoImage(image)
        # создание кнопки вызова проводника
        file_explorer_bt_1 = Button(frame1, command=self.file_explorer_bt_1_press,
                                    image=win_explorer)
        file_explorer_bt_1.grid(row=1, column=1, padx=GRID_PADX, pady=GRID_PADY, sticky=W)

        file_explorer_bt_2 = Button(frame1, command=self.file_explorer_bt_2_press,
                                    image=win_explorer)
        file_explorer_bt_2.grid(row=2, column=1, padx=GRID_PADX, pady=GRID_PADY, sticky=W)

        # self.text_var_1 = StringVar()
        # self.text_var_2 = StringVar()
        # self.text_var_1.set("Не выбрано")
        # self.text_var_2.set("Не выбрано")

        selected_table_1_label = Label(frame1, text=0, textvariable=self.text_var_1,
                                       font=W1_FONT, background=BG_COLOR)
        selected_table_1_label.grid(row=1, column=2, padx=GRID_PADX, pady=GRID_PADY, sticky=W + E)

        selected_table_2_label = Label(frame1, text=0, textvariable=self.text_var_2,
                                       font=W1_FONT, background=BG_COLOR)
        selected_table_2_label.grid(row=2, column=2, padx=GRID_PADX, pady=GRID_PADY, sticky=W + E)

        # создание надписи "Выбранные Боты"
        selected_tables_label = Label(frame1, text="Выбранные файлы", width=15,
                                      font=W1_FONT, background=BG_COLOR)
        selected_tables_label.grid(row=0, column=2, padx=GRID_PADX, pady=GRID_PADY, sticky=E + W)

        # ______________________________
        # Блок 2
        # создание контейнера frame3
        frame2 = Frame(self, background=BG_COLOR)
        frame2.pack(fill=BOTH, expand=True)

        # создание кнопки начала турнира
        self.__start_prog_bt = Button(frame2, width=15, height=2, command=self.start_prog_bt_press,
                                      text="Начать", font="Times 16", state="disable")
        self.__start_prog_bt.pack(expand=True, padx=FRAME_PADX, pady=FRAME_PADY)
        # ______________________________
        # mainloop
        self.mainloop()

    def start_prog_bt_press(self):
        # try:
        df1 = pd.read_excel(rf"{self.selected_tables[0]}", header=1)
        # print(df1.loc[:, ['№ вагона', "Дата подачи"]])
        dataframe = pd.read_excel(rf"{self.selected_tables[1]}", header=0)
        # print(dataframe.iloc[:, 1])
        # print(dataframe.iloc[:, 4])
        # print(dataframe.iloc[:, 7])
        last_date = ""
        df1['Дата подачи'] = pd.to_datetime(df1['Дата подачи'], errors='coerce')
        print(dataframe)
        for i in [1, 4, 7]:
            for j in range(len(dataframe.iloc[:, i])):
                # print(str(dataframe.iloc[:, i][j]))
                date = dataframe.iloc[:, i][j]
                # print(str(date))
                if last_date == "" or str(date) == "nan" or date.isoformat() != "NaT":
                    last_date = str(date)
                else:
                    dataframe.iloc[:, i][j] = pd.to_datetime(last_date)
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        # print(dataframe)

        dataframe1 = dataframe.iloc[:, 3:6].rename({"3": "", "4дата": "дата", "50№вагона": "№вагона"})
        dataframe2 = dataframe.iloc[:, 6:9].rename({"6": "", "7дата": "дата", "8№вагона": "№вагона"})
        dataframe = dataframe.drop(dataframe.columns[[j for j in range(3, 9)]], axis=1)
        dataframe.columns = ["", "Дата", "№ вагона"]
        dataframe1.columns = ["", "Дата", "№ вагона"]
        dataframe2.columns = ["", "Дата", "№ вагона"]
        dataframe = pd.concat([dataframe, dataframe1, dataframe2], ignore_index=True).set_index("").sort_index()
        dataframe = dataframe[pd.isna(dataframe['№ вагона']) == False]
        # dataframe = dataframe[dataframe['№ вагона'] != "NaT"]
        dataframe[["№ вагона"]] = dataframe[["№ вагона"]].astype(int)
        # print(dataframe)

        result_data = {"№ вагона": [], "Дата": [], "Проблема": []}
        year, month, day = map(int, str(dataframe.iloc[0, 0]).split(" ")[0].split("-"))
        # print(day, month, year)
        start_month = (2 - len(str(month - 1))) * "0" + str(month - 1)
        start_year = str(year)
        if start_month == "00":
            start_month = "12"
            start_year = str(int(year) - 1)
        start_date = f"{start_year}-{start_month}-15"
        start_date = pd.to_datetime(f"{start_date}")

        df1['Дата подачи'] = pd.to_datetime(df1['Дата подачи'], errors='coerce')
        filter1 = df1['Дата подачи'] > start_date
        df1 = df1[filter1]
        df1 = df1.loc[:, ['№ вагона', "Дата подачи"]]
        df1 = df1[pd.isna(df1['№ вагона']) == False]
        # print(df1)
        used = []
        for i in range(len(dataframe["№ вагона"])):
            # print(1)
            dataframe["№ вагона"].iloc[i] = str(dataframe["№ вагона"].iloc[i]).strip()
            # if dataframe["№ вагона"].iloc[i] == "":
            #     continue
            closest = [100, "", ""]
            for j in range(len(df1["№ вагона"])):
                # print(str(df1["№ вагона"].iloc[j]).strip(), str(dataframe["№ вагона"].iloc[i]).strip())
                if lv.distance(str(df1["№ вагона"].iloc[j]).strip(), str(dataframe["№ вагона"].iloc[i]).strip()) < closest[0]:
                    # print(str(df1["№ вагона"].iloc[j]).strip(), str(dataframe["№ вагона"].iloc[i]).strip())
                    closest[0] = lv.distance(str(df1["№ вагона"].iloc[j]).strip(), str(dataframe["№ вагона"].iloc[i]).strip())
                    closest[1] = str(df1["№ вагона"].iloc[j]).strip()
                    closest[2] = str(df1["Дата подачи"].iloc[j])
            if closest[0] > 0:
                result_data["№ вагона"].append(dataframe["№ вагона"].iloc[i])
                result_data["Дата"].append(dataframe["Дата"].iloc[i])
                result_data["Проблема"].append(f'Номера нет в таблице Ближаший номер: {closest[1]} Дата подачи: {closest[2].replace("-", ".")[:-9]}')
            if not pd.isna(dataframe["№ вагона"].iloc[i]) and dataframe["№ вагона"].value_counts()[dataframe["№ вагона"].iloc[i]] > 1 and str(dataframe["№ вагона"].iloc[i]) not in used:
                used.append(str(dataframe["№ вагона"].iloc[i]))
                result_data["№ вагона"].append(dataframe["№ вагона"].iloc[i])
                result_data["Дата"].append(dataframe["Дата"].iloc[i])
                result_data["Проблема"].append(f'Номер встретился больше 1 раза')

        res = pd.DataFrame(result_data)
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        # print(res)
        res.to_excel('Результат.xlsx')
        messagebox.showinfo("Успех!", "Файл успешно создан.")
        self.quit()
        # except:
        #     messagebox.showinfo("Ошибка!", "Неизвестная ошибка.\nПроверьте таблицы.\nЗакройте окно файла Результат.xlsx.")
        #     self.quit()

        """
        Вагона нет/номер вагона не совподает
        Задвоение вагона
        """

    def file_explorer_bt_2_press(self):
        """ Отвечает за работу file_explorer_bt. Выводит имена выбранных ботов. Удаляет повторения.
            Проверяет, достаточно ли данных для начала турнира."""
        path = fd.askopenfilename(title='Выберете таблицу для сравнения',
                                  filetypes=[('*', '.xlsx'), ('*', '.xls')])
        if path != self.selected_tables[1]:
            self.selected_tables[1] = path

            self.text_var_2.set("")
            self.text_var_2.set(path.split("/")[-1])

        if self.text_var_2.get() == "":
            self.text_var_2.set("Не выбрано")
        if all(self.selected_tables):
            self.__start_prog_bt.configure(state="normal")
        else:
            self.__start_prog_bt.configure(state="disable")

    def file_explorer_bt_1_press(self):
        """ Отвечает за работу file_explorer_bt. Выводит имена выбранных ботов. Удаляет повторения.
            Проверяет, достаточно ли данных для начала турнира."""
        path = fd.askopenfilename(title='Выберете внутреннюю таблицу',
                                  filetypes=[('*', '.xlsx'), ('*', '.xls')])
        if path != self.selected_tables[0]:
            self.selected_tables[0] = path

            self.text_var_1.set("")
            self.text_var_1.set(path.split("/")[-1])
        if self.text_var_1.get() == "":
            self.text_var_1.set("Не выбрано")
        if all(self.selected_tables):
            self.__start_prog_bt.configure(state="normal")
        else:
            self.__start_prog_bt.configure(state="disable")

    def read_docx_table(self, document, table_num=1, nheader=1):
        ans = []
        for table in document.tables:
            data = [[cell.text for cell in row.cells[:10]] for row in table.rows]
            df = pd.DataFrame(data)
            if nheader == 1:
                df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
            elif nheader == 2:
                outside_cold, inside_col = df.iloc[0], df.iloc[1]
                hier_index = pd.MultiIndex.from_tuples(list(zip(outside_cold, inside_col)))
                df = pd.DataFrame(data, columns=hier_index).drop(df.index[[0, 1]]).reset_index(drop=True)
            elif nheader > 2:
                print("More than two headers not currently supported")
                df = pd.DataFrame()
            ans.append(df)
        return ans


def main():
    s = StartWindow()


if __name__ == '__main__':
    main()
