from win32com.client import Dispatch


def just_open(filename):
    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = False
    xlBook = xlApp.Workbooks.Open(filename)
    xlBook.Save()
    xlBook.Close()

if __name__ == '__main__':
    path_excel = "C:\\Users\sai\AppData\Local\Temp\TSM_Export_Excel.py\Alliance - 比格沃斯.xlsx"
    just_open(path_excel)
