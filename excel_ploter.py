import pandas as pd
from openpyxl.chart import Reference, LineChart, AreaChart

class A:
    def __init__(self, csvfile):
        self._df = pd.read_csv(csvfile)
        self._df["DateTime"] = pd.to_datetime(self._df["公表_年月日"])
        self._df = self._df.set_index("DateTime")
        self._file = "test.xlsx"
        self._sheetname = "test"
        print(self._df)

    def _plot(self, chart, title, xlabel, ylabel, col_start = 0, col_end = 0, col_step = 0):
        if col_start == 0:
            col_start = 2 # データの開始列
        if col_end == 0:
            col_end = len(self._df.columns) + 1 # データの終了列

        with pd.ExcelWriter(self._file) as writer:
            # write the df on excel at once
            self._df.to_excel(writer, sheet_name=self._sheetname)

            # plot on the sheet
            sheet = writer.book.active
            chart.title = title
            chart.x_axis.title = xlabel
            chart.y_axis.title = ylabel
            chart.add_data(Reference(sheet, min_col = col_start, max_col = col_end, min_row = 1, max_row = sheet.max_row), titles_from_data = True)
            chart.set_categories(Reference(sheet, min_col = 1, max_col = 1, min_row = 2, max_row = sheet.max_row)) # ラベル
            sheet.add_chart(chart, "A1")

    def print(self):
        print(self._df)

    def plot_bar(self):
        pass

    def plot_line(self, column, value):
        self._df = self._df[[value, column]].pivot_table(index = "DateTime", values = value, columns = column, fill_value = 0)

        chart = LineChart()
        self._plot(chart, title = "線グラフ", xlabel = "日付", ylabel = "回数")

    def plot_Area(self, column, value):
        self._df = self._df[[value, column]].pivot_table(index = "DateTime", values = value, columns = column, fill_value = 0)
        self.print()

        chart = AreaChart()
        chart.grouping = "stacked"
        self._plot(chart, title = "面グラフ", xlabel = "日付", ylabel = "回数")

    def plot_Stack(self, column, value, unit):
        self._df = self._df[[value, column]].pivot_table(index = "DateTime", values = value, columns = column, fill_value = 0)
        self._df = self._df.resample(unit).sum()

        # self._df.to_csv("a.csv", index = True, header = True, encoding='utf_8_sig')

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
    a.plot_Area(column = "患者_年代", value = "退院済フラグ")
    # a.plot_line(column = "患者_年代", value = "退院済フラグ")