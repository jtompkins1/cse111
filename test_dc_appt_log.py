import pytest
import openpyxl
import tkinter as tk
from tkinter import Frame
from dc_appt_log import populate_main_window


def test_insert_row():
    # change filepath to where you wish to save the test_dc.xlsx file that will be created.
    filepath = (r"C:\Users\Jennifer\Documents\jen_school\byui\programming\cse111_programing_with_functions\cse111\test_dc.xlsx")

    root = tk.Tk()
    frame = Frame(root)
    frame.master.title("DC Scheduling Log")
    frame.pack()

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    heading = ["DC", "Delivery#"]
    sheet.append(heading)

    dc = 6003
    delnum = "98765432"
    row_values = [dc, delnum]

    sheet.append(row_values)

    workbook.save(filepath)
    populate_main_window(frame, filepath)

    assert sheet["A1"].value == "DC"
    assert sheet["B1"].value == "Delivery#"
    assert sheet["A2"].value == 6003
    assert sheet["B2"].value == "98765432"

# Call the main function that is part of pytest so that the
# computer will execute the test functions in this file.
pytest.main(["-v", "--tb=line", "-rN", __file__])
