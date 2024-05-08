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
        self.choosen_folder = False

        # Widgets
        self.choose_printer_button = tk.Button(master=self.master,
                                               text="Wybierz drukarkę",
                                               command=self.choose_printer)
        self.choose_folder_button = tk.Button(master=self.master,
                                              text="Wybierz folder",
                                              command=self.choose_folder)
        self.number_of_files = tk.Frame(master=self.master)
        self.files_list = tk.Listbox(self.number_of_files)
        self.submit_button = tk.Button(master=self.master,
                                       text="Wyślij do wydruku")

        # Widgets Placing
        self.choose_printer_button.place(relx=0.5, rely=0.05, anchor="n",
                                         relwidth=0.9, relheight=0.1)
        self.choose_folder_button.place(relx=0.5, rely=0.2, anchor="n",
                                        relwidth=0.9, relheight=0.1)
        self.number_of_files.place(relx=0.5, rely=0.35, anchor="n",
                                   relwidth=0.9, relheight=0.45)
        self.files_list.pack(fill=tk.BOTH, expand=True)
        self.submit_button.place(relx=0.5, rely=0.85, anchor="n",
                                 relwidth=0.9, relheight=0.1)

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

    def choose_folder(self):
        self.choosen_folder = filedialog.askdirectory()
        if self.choosen_folder:
            self.choose_folder_button.configure(text=self.choosen_folder)

            self.pdf_files = [
                filename for filename in os.listdir(self.choosen_folder) if filename.lower().endswith('.pdf')
            ]

            self.files_list.delete(0, tk.END)
            for index, file in enumerate(self.pdf_files):
                self.files_list.insert(tk.END, f"{index}: {file}")


def main():
    root = tk.Tk()
    window = Window(root)
    root.mainloop()


if __name__ == "__main__":
    main()
