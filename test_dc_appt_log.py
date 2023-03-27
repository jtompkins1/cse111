from dc_appt_log import populate_main_window
import pytest
import openpyxl
import tkinter as tk
from tkinter import ttk, Frame, Label


def test_populate_main_window():
    pass


# Call the main function that is part of pytest so that the
# computer will execute the test functions in this file.
pytest.main(["-v", "--tb=line", "-rN", __file__])
