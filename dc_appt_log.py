import tkinter as tk
from tkinter import ttk, Frame, Label, Button
import openpyxl

def main():

    root = tk.Tk()
    frame = Frame(root)
    frame.master.title("DC Scheduling Log")
    frame.pack()

    populate_main_window(frame)


    root.mainloop()

def populate_main_window(frame):

    # change the file path to match where the dc.xlsx is saved
    path = (r"C:\Users\Jennifer\Documents\jen_school\byui\programming\cse111_programing_with_functions\cse111\dc.xlsx")

    widgets_frame = ttk.LabelFrame(frame, text="Insert Row")
    widgets_frame.grid(row=0, column=0, padx=20, pady=10)

    dc_entry = ttk.Entry(widgets_frame)
    dc_entry.insert(0,"DC")
    # when you click on this entry it clears the text
    dc_entry.bind("<FocusIn>", lambda e: dc_entry.delete("0", "end"))
    dc_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

    # drop down selection
    dc_list = ["Select DC#", 6003, 6018, 7035, 7034, 7026, 6094, 6017, 6011, 6020, 6030, 6040]
    dc_entry = ttk.Combobox(widgets_frame, values=dc_list)
    dc_entry.current(0)
    dc_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

    del_entry = ttk.Entry(widgets_frame)
    del_entry.insert(0,"Delivery#")
    # when you click on this entry it clears the text
    del_entry.bind("<FocusIn>", lambda e: del_entry.delete("0", "end"))
    del_entry.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

    def load_data():
        """load data from excel file to treeview"""
        
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active

        list_values = list(sheet.values)
        
        # add column names to treeview
        for col_name in list_values[0]:
            treeview.heading(col_name, text=col_name)
        #skip row one
        for value_tuple in list_values[1:]:
            treeview.insert("", tk.END, values=value_tuple)

    def insert_row():
        dc = dc_entry.get()
        delnum = del_entry.get()

        # insert row into excel sheet
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        row_values = [dc, delnum]
        sheet.append(row_values)
        workbook.save(path)
        
        # insert row into treeview
        treeview.insert("", tk.END, values=row_values)

        # clear values
        # dc_entry.delete(0, "end")
        # dc_entry.insert(0, "DC")
        dc_entry.set(dc_list[0])
        del_entry.delete(0, "end")
        del_entry.insert(0, "Delivery#")

    # # check button
    # a = tk.BooleanVar()
    # checkbutton = ttk.Checkbutton(widgets_frame, text="Emailed", variable=a)
    # checkbutton.grid(row=3, column=0, padx=5, pady=(0, 5), sticky="news")

    # insert button
    button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
    button.grid(row=2,  column=0, padx=5, pady=(0, 5), sticky="news")

    # create a separator
    separator = ttk.Separator(widgets_frame)
    separator.grid(row=3, column=0, padx=(20, 10), pady=10, sticky="ew")

    # widget to show count of rows
    lbl_row_count = Label(widgets_frame, text="Total Number of Rows: " )
    lbl_row_count.grid(row=4, column=0, padx=(20, 10), pady=10, sticky="ew")
    lbl_show_row_count = Label(widgets_frame, width=3)
    lbl_show_row_count.grid(row=5, column=0, padx=(20, 10), pady=10, sticky="ew")

    # create treeview frame
    treeFrame = ttk.Frame(frame)
    treeFrame.grid(row=0, column=1, pady=10)

    # create scrollbar
    treeScroll = ttk.Scrollbar(treeFrame)
    treeScroll.pack(side="right", fill="y")

    # create columns
    cols = ("DC", "Delivery#")

    # create treeview
    treeview = ttk.Treeview(treeFrame, show="headings", yscrollcommand=treeScroll.set, columns=cols, height=20)

    # set width of treeview
    # treeview["columns"] = ("DC", "Delivery#")
    treeview.column("DC", width=100)
    treeview.column("Delivery#", width=100)

    # Configure column headings
    treeview.heading("#0", text="Item")
    treeview.heading("DC", text="DC")
    treeview.heading("Delivery#", text="Delivery#")


    # position treeview
    treeview.pack()
    # config scrollbar
    treeScroll.config(command=treeview.yview)

    return ttk.Treeview

    load_data()


# TODO :
    # running total of rows not including heading

    # def calculate_row_count():
    #     lbl_show_row_count.config()
        

    # insert current datetime into column at time of delnum entry


if __name__ == "__main__":
    main()