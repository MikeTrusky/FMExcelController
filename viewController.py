from tkinter import *
import customtkinter
from tkinter import filedialog
from excelController import ExcelModificationsController

class ViewApp(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.excel_modifications_controller = ExcelModificationsController()

        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("blue")

        self.title("FM Excel Controller App")
        self.geometry("300x400")

        self.choose_file_button = customtkinter.CTkButton(self, text="Choose csv file", command=self.on_choose_file_button_click)
        self.choose_file_button.place(relx=0.5, rely=0.2, anchor=CENTER)

        self.file_label = customtkinter.CTkLabel(self, text="")
        self.file_label.place(relx=0.5, rely=0.4, anchor=CENTER)

        self.update_excel_button = customtkinter.CTkButton(self, text="Update player with file!", command=self.on_update_excel_button_click)
        self.update_excel_button.place(relx=0.5, rely=0.6, anchor=CENTER)

    def on_choose_file_button_click(self):
        file_path = filedialog.askopenfilename()
        self.file_label.configure(text=file_path)        

    def on_update_excel_button_click(self):
        print("Update player by file")
        #self.excel_modifications_controller.update_player_by_file()

viewApp = ViewApp()
viewApp.mainloop()