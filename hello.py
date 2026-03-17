import xlwings as xw

def my_function():
    msg = "増田三莉音"
    wb = xw.Book.caller()
    wb.sheets("Sheet1").range('A1').value = msg 