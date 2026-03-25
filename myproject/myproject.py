import xlwings as xw
from xlwings import func


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"


@func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("myproject.xlsm").set_mock_caller()
    main()
