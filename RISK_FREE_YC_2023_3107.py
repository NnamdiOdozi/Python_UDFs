import xlwings as xw

@xw.sub
def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"
        
@xw.func
def linear(x,z):
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        #sheet["A1"].value = "Bye Interpolation !"
        y =  1
    else:
        #sheet["A1"].value = "Hello Interpolation!"
        y =  0
    return y *(x+z)

@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("myproject.xlsm").set_mock_caller()
    main()