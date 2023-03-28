import pytest
import openpyxl
from tkinter import ttk, Frame, Label
import tkinter as tk
from dc_appt_log import populate_main_window

# change filepath to match location of dc.xlsx file
path = (r"C:\Users\Jennifer\Documents\jen_school\byui\programming\cse111_programing_with_functions\cse111\dc.xlsx")

def test_insert_row():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["DC", "Delivery#"])
    sheet.append([6003, "D0001"])
    sheet.append([7035, "D0002"])
    workbook.save(path)

    root = tk.Tk()
    frame = Frame(root)
    frame.master.title("DC Scheduling Log")
    frame.pack()

    populate_main_window(frame)

    assert sheet["A1"].value == "DC"
    assert sheet["A2"].value == 6003
    assert sheet["A3"].value == 7035
    assert sheet["B1"].value == "Delivery#"
    assert sheet["B2"].value == "D0001"
    assert sheet["B3"].value == "D0002"


# Call the main function that is part of pytest so that the
# computer will execute the test functions in this file.
pytest.main(["-v", "--tb=line", "-rN", __file__])
