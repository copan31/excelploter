import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import Reference, LineChart, AreaChart, BarChart
import os

class A:
    def __init__(self, csvfile):
        self._df = pd.read_csv(csvfile)
        self._df["DateTime"] = pd.to_datetime(self._df["公表_年月日"])
        self._df = self._df.set_index("DateTime")
        self._file = "test.xlsx"
        print(self._df)

    def _plot(self, sheet, chart, title, xlabel, ylabel, ymax = 0):
        # plot on the sheet
        chart.title = title
        chart.x_axis.title = xlabel
        chart.y_axis.title = ylabel
        chart.y_axis.scaling.min = 0
        if ymax != 0:
            chart.y_axis.scaling.max = ymax # 指定があるならy軸の最大値を設定
        chart.legend.position = 'r' # 凡例は右に設定
        chart.add_data(Reference(sheet, min_col = 2, max_col = sheet.max_column, min_row = 1, max_row = sheet.max_row), titles_from_data = True)
        chart.set_categories(Reference(sheet, min_col = 1, max_col = 1, min_row = 2, max_row = sheet.max_row)) # ラベル
        sheet.add_chart(chart, "A1")

    def reload(self):
        self._df = pd.read_csv(csvfile)

    def print(self):
        print(self._df)

    def plot_bar(self):
        pass

    def plot_line(self, sheetname, column, value):
        # 必要な値と項目を残した時系列データを作成
        df = self._df[[value, column]].pivot_table(index = "DateTime", values = value, columns = column, fill_value = 0)

        with pd.ExcelWriter(self._file) as writer:
            # excelファイルがあるなら追加書きする
            if os.path.isfile(self._file):
                writer.book = load_workbook(self._file)
            # write the df on excel at once
            df.to_excel(writer, sheet_name = sheetname)

            # plot line chart
            chart = LineChart()            
            self._plot(writer.book[sheetname], chart, title = "線グラフ", xlabel = "日付", ylabel = "回数")

    def plot_Area(self, sheetname, column, value):
        # 必要な値と項目を残した時系列データを作成
        df = self._df[[value, column]].pivot_table(index = "DateTime", values = value, columns = column, fill_value = 0)

        with pd.ExcelWriter(self._file) as writer:
            # excelファイルがあるなら追加書きする
            if os.path.isfile(self._file):
                writer.book = load_workbook(self._file)

            # write the df on excel at once
            df.to_excel(writer, sheet_name = sheetname)

            # plot Area chart
            chart = AreaChart()
            chart.grouping = "stacked"
            self._plot(writer.book[sheetname], chart, title = "面グラフ", xlabel = "日付", ylabel = "回数")

    def plot_Stack(self, sheetname, column, value, unit):
        df = self._df[[value, column]].pivot_table(index = "DateTime", values = value, columns = column, fill_value = 0)
        df = df.resample(unit).sum()

        # self._df.to_csv("a.csv", index = True, header = True, encoding='utf_8_sig')

        with pd.ExcelWriter(self._file) as writer:
            # excelファイルがあるなら追加書きする
            if os.path.isfile(self._file):
                writer.book = load_workbook(self._file)

            # write the df on excel at once
            df.to_excel(writer, sheet_name = sheetname)

            # plot Area chart
            chart = BarChart()
            chart.type = "col"
            chart.grouping = "stacked"
            chart.overlap = 100
            self._plot(writer.book[sheetname], chart, title = "積み上げ棒グラフ", xlabel = "月", ylabel = "回数")

    def plot_test(self):
        sep = 4
        l = len(self._df.columns)
        n = int(l / sep)
        # dfs = [self._df.iloc[ :, i : i + sep] for i in range(0, n, sep)]
        dfs = [self._df.iloc[:, i: i + sep].head() for i in range(0, l , sep)]
        print(dfs)

if __name__ == "__main__":
    csvfile = "130001_tokyo_covid19_patients(1).csv"
    a = A(csvfile)
    a.plot_line(sheetname = "Line", column = "患者_年代", value = "退院済フラグ")
    a.plot_Area(sheetname = "Area", column = "患者_年代", value = "退院済フラグ")
    a.plot_Stack(sheetname = "Stack", column = "患者_年代", value = "退院済フラグ", unit = 'M')
