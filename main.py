import os
import subprocess
import win32print
import win32api
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog


class Window:
    def __init__(self, master):
        # Settings
        self.master = master
        self.master.title("Many PDF Print")
        self.master.geometry("600x800")
        self.master.configure(background="#a3a3a3")

        self.master.bind("<Escape>", self.confirm_exit)
        self.printers = self.get_available_printers()
        self.choosen_printer = False

        # Widgets
        self.choose_printer_button = tk.Button(master=self.master,
                                               text="Wybierz drukarkę",
                                               command=self.choose_printer)
        self.choose_folder_button = tk.Button(master=self.master,
                                              text="Wybierz folder")
        self.number_of_files = tk.Label(master=self.master,
                                        text="Nie wybrano plików")
        self.submit_button = tk.Button(master=self.master,
                                       text="Wyślij do wydruku")

        # Widgets Placing
        self.choose_printer_button.place(relx=0.25, rely=0.1, anchor="center",
                                         relwidth=0.4, relheight=0.15)
        self.choose_folder_button.place(relx=0.75, rely=0.1, anchor="center",
                                        relwidth=0.4, relheight=0.15)
        self.number_of_files.place(relx=0.5, rely=0.5, anchor="center",
                                   relwidth=0.9, relheight=0.6)
        self.submit_button.place(relx=0.5, rely=0.9, anchor="center",
                                 relwidth=0.9, relheight=0.15)

    def confirm_exit(self, event=None):
        if messagebox.askyesno("Wyjście", "Czy na pewno chcesz wyjść z programu?"):
            self.master.destroy()

    def get_available_printers(self):
        try:
            result = subprocess.run(["wmic", "printer", "get", "name"], capture_output=True, text=True)
            printers = result.stdout.strip().split("\n")
            return [printer.strip() for printer in printers[1:] if printer.strip()]
        except Exception:
            return []

    def set_printer(self, printer):
        self.choosen_printer = printer
        self.choose_printer_button.configure(text=self.choosen_printer)
        self.printer_window.destroy()

    def choose_printer(self):
        if self.printers:
            self.printer_window = tk.Toplevel(self.master)
            self.printer_window.title("Wybierz drukarkę")

            for printer in self.printers:
                printer_widget = tk.Button(master=self.printer_window,
                                           text=printer,
                                           command=lambda printer=printer: self.set_printer(printer))
                printer_widget.pack()
        else:
            print("Brak drukarek")


def main():
    root = tk.Tk()
    window = Window(root)
    root.mainloop()


if __name__ == "__main__":
    main()
