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
        self.master.title("Print Many Files")
        self.master.geometry("600x800")
        self.master.configure(background="#a3a3a3")

        self.master.bind("<Escape>", self.confirm_exit)

        self.file_types = [("Dokumenty", "*.doc *.docx *.pdf")]
        self.printers = self.get_available_printers()
        self.choosen_printer = False
        self.choosen_files = []
        self.not_printed = []

        self.testing = False

        # Widgets
        self.choose_printer_button = tk.Button(master=self.master,
                                               text="Wybierz drukarkę",
                                               command=self.choose_printer)
        self.choose_files_button = tk.Button(master=self.master,
                                             text="Wybierz pliki",
                                             command=self.choose_files)
        self.files_list = tk.Listbox(master=self.master)
        self.submit_button = tk.Button(master=self.master,
                                       text="Wyślij do wydruku",
                                       command=self.send_to_print,
                                       state="disabled")

        # Widgets Placing
        self.choose_printer_button.place(relx=0.5, rely=0.05, anchor="n",
                                         relwidth=0.9, relheight=0.1)
        self.choose_files_button.place(relx=0.5, rely=0.2, anchor="n",
                                       relwidth=0.9, relheight=0.1)
        self.files_list.place(relx=0.5, rely=0.35, anchor="n",
                              relwidth=0.9, relheight=0.45)
        self.submit_button.place(relx=0.5, rely=0.85, anchor="n",
                                 relwidth=0.9, relheight=0.1)

    def confirm_exit(self, event=None):
        if messagebox.askyesno("Wyjście", "Czy na pewno chcesz wyjść z programu?"):
            self.master.destroy()

    def check_printer_and_files(self):
        if self.choosen_files and self.choosen_printer:
            self.submit_button.configure(state="normal")
        else:
            self.submit_button.configure(state="disabled")

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
        self.check_printer_and_files()

    def choose_printer(self):
        if not hasattr(self, "printer_window") or not self.printer_window.winfo_exists():
            if self.printers:
                self.printer_window = tk.Toplevel(self.master)
                self.printer_window.title("Wybierz drukarkę")

                for printer in self.printers:
                    printer_widget = tk.Button(master=self.printer_window,
                                               text=printer,
                                               command=lambda printer=printer: self.set_printer(printer),
                                               padx=20,
                                               pady=20)
                    printer_widget.pack(fill=tk.BOTH, expand=True)
            else:
                messagebox.showerror("Brak danych", "Brak drukarek")

    def choose_files(self):
        self.choosen_files = filedialog.askopenfilenames(filetypes=self.file_types)
        if self.choosen_files:
            if len(self.choosen_files) != 0:
                if len(self.choosen_files) == 1:
                    self.choose_files_button.configure(text=f"Wybrano {len(self.choosen_files)} plik")
                elif len(self.choosen_files) % 10 in [2, 3, 4] and len(self.choosen_files) % 100 not in [12, 13, 14]:
                    self.choose_files_button.configure(text=f"Wybrano {len(self.choosen_files)} pliki")
                else:
                    self.choose_files_button.configure(text=f"Wybrano {len(self.choosen_files)} plików")

                self.files_list.delete(0, tk.END)
                self.submit_button.configure(text=f"Wyślij do wydruku ({len(self.choosen_files)})")
                for index, file in enumerate(self.choosen_files):
                    self.files_list.insert(tk.END, f"{index}: {file}")
            else:
                messagebox.showinfo("Brak danych", "Nie wybrano plików")

        self.check_printer_and_files()

    def send_to_print(self):
        if self.choosen_files and self.choosen_printer:
            if not self.testing:
                for file in self.choosen_files:
                    try:
                        file_path = file
                        filename = os.path.basename(file)
                        hPrinter = win32print.OpenPrinter(self.choosen_printer)

                        if file.endswith('.pdf'):
                            hJob = win32print.StartDocPrinter(hPrinter, 1, (filename, None, "RAW"))
                            win32print.StartPagePrinter(hPrinter)
                            win32print.WritePrinter(hPrinter, open(file_path, 'rb').read())
                            win32print.EndPagePrinter(hPrinter)
                            win32print.EndDocPrinter(hPrinter)
                            win32print.ClosePrinter(hPrinter)
                        elif file.endswith('.doc') or file.endswith('.docx'):
                            win32print.SetDefaultPrinter(self.choosen_printer)
                            win32api.ShellExecute(0, "print", file_path, None, ".", 0)

                    except Exception as e:
                        self.not_printed.append(file)
                        messagebox.showerror("Błąd", e)

            messagebox.showinfo("Wynik", "Wysłano pliki do drukarki")

            self.files_list.delete(0, tk.END)
            self.submit_button.configure(text=f"Wyślij do wydruku")

            if self.not_printed:
                messagebox.showerror("Wynik", f"Nie wydrukowano {len(self.not_printed)}")

                for index, file in enumerate(self.not_printed):
                    self.files_list.insert(tk.END, f"{index}: {file}")

                self.choose_files_button.configure(text=f"Nie wydrukowano {len(self.not_printed)}")
                self.choosen_files = self.not_printed
            else:
                self.choose_files_button.configure(text="Wybierz pliki")
                self.choosen_files = False

            self.check_printer_and_files()
        else:
            if not self.choosen_printer and not self.choosen_files:
                messagebox.showerror("Brak danych", "Nie wybrano drukarki i plików")
            elif not self.choosen_printer and self.choosen_files:
                messagebox.showerror("Brak danych", "Nie wybrano drukarki")
            elif self.choosen_printer and not self.choosen_files:
                messagebox.showerror("Brak danych", "Nie wybrano plików")


def main():
    root = tk.Tk()
    window = Window(root)
    root.mainloop()


if __name__ == "__main__":
    main()
