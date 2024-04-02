from tkinter import *
import customtkinter
from tkinter import filedialog
from excelController import ExcelModificationsController
import os

class ViewApp(customtkinter.CTk):
    def __init__(self):
        super().__init__()        

        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("blue")

        self.title("FM Excel Controller App")
        self.geometry("300x300")

        self.choose_excel_file_button = customtkinter.CTkButton(self, text="Choose excel file", command=self.on_choose_excel_file_button_click)
        self.choose_excel_file_button.place(relx=0.5, rely=0.2, anchor=CENTER)

        self.excel_file_label = customtkinter.CTkLabel(self, text="Excel file name: ")
        self.excel_file_label.place(relx=0.2, rely=0.25) 
        self.excel_file_name_label = customtkinter.CTkLabel(self, text="none")
        self.excel_file_name_label.place(relx=0.5, rely=0.25)

        self.choose_csv_file_button = customtkinter.CTkButton(self, text="Choose csv file", command=self.on_choose_csv_file_button_click)
        self.choose_csv_file_button.place(relx=0.5, rely=0.5, anchor=CENTER)
        self.choose_csv_file_button.configure(state="disabled")

        self.csv_file_label = customtkinter.CTkLabel(self, text="CSV file name: ")
        self.csv_file_label.place(relx=0.2, rely = 0.55) 
        self.csv_file_name_label = customtkinter.CTkLabel(self, text="none")
        self.csv_file_name_label.place(relx=0.5, rely=0.55)
        
        self.update_excel_button = customtkinter.CTkButton(self, text="Update", width=100, command=self.on_update_excel_button_click)
        self.update_excel_button.place(relx=0.3, rely=0.8, anchor=CENTER)
        self.update_excel_button.configure(state="disabled")

        self.delete_player_excel_button = customtkinter.CTkButton(self, text="Delete", width=100, command=self.on_delete_player_excel_button_click)
        self.delete_player_excel_button.place(relx=0.7, rely=0.8, anchor=CENTER)
        self.delete_player_excel_button.configure(state="disabled")

    def on_choose_excel_file_button_click(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.excel_file_name = os.path.basename(file_path)
        self.excel_file_name_label.configure(text=self.excel_file_name)        
        self.choose_csv_file_button.configure(state="normal")

    def on_choose_csv_file_button_click(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.csv_file_name = os.path.basename(file_path)
        self.csv_file_name_label.configure(text=self.csv_file_name)        
        self.update_excel_button.configure(state="normal")        
        self.insert_player_excel_button.configure(state="normal")
        self.delete_player_excel_button.configure(state="normal")

    def on_update_excel_button_click(self):  
        self.excel_modifications_controller = ExcelModificationsController(self.excel_file_name)
        self.excel_modifications_controller.update_player_by_file(self.csv_file_name)

    def on_delete_player_excel_button_click(self):
        self.excel_modifications_controller = ExcelModificationsController(self.excel_file_name)
        self.excel_modifications_controller.delete_player_by_file(self.csv_file_name)

viewApp = ViewApp()
viewApp.mainloop()