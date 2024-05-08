import os
import subprocess
import win32print
import win32api
import tkinter


def get_available_printers():
    try:
        result = subprocess.run(["wmic", "printer", "get", "name"], capture_output=True, text=True)
        printers = result.stdout.strip().split("\n")
        return [printer.strip() for printer in printers[1:]]
    except Exception as e:
        print("Wystąpił błąd podczas pobierania dostępnych drukarek:", e)
        return []


def print_pdf_files(folder_path, selected_printer):
    if not os.path.isdir(folder_path):
        print("Podana ścieżka nie jest katalogiem")
        return

    pdf_files = [filename for filename in os.listdir(folder_path) if filename.lower().endswith('.pdf')]

    print(f"W folderze znaleziono {len(pdf_files)} plików PDF")

    not_printed = []

    i = 1
    for filename in pdf_files:
        file_path = os.path.join(folder_path, filename)
        print(f"{i}. Drukowanie pliku {filename}")
        try:
            hPrinter = win32print.OpenPrinter(selected_printer)
            hJob = win32print.StartDocPrinter(hPrinter, 1, (filename, None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, open(file_path, 'rb').read())
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            win32print.ClosePrinter(hPrinter)
            print(f"{i}. {filename} -  Wysłano do drukowania")
        except Exception as e:
            print(f"{i}. {filename} - Wystąpił błąd podczas drukowania")
            print(e)
            not_printed.append(f"{i}. {filename}")
        i += 1

    return not_printed


current_directory = os.path.dirname(os.path.realpath(__file__))
printers = get_available_printers()


if __name__ == '__main__':
    current_directory = os.path.dirname(os.path.realpath(__file__))
    printers = get_available_printers()

    if printers:
        print("Dostępne drukarki:")
        for i, printer in enumerate(printers):
            print(f"{i + 1}. {printer}")

        selected_printer_index = input("Wybierz numer drukarki: ")
        try:
            selected_printer_index = int(selected_printer_index)
            if 1 <= selected_printer_index <= len(printers):
                selected_printer = printers[selected_printer_index - 1]
                not_printed = print_pdf_files(current_directory, selected_printer)
                if not_printed:
                    print("Nie wydrukowano:\n")
                    for file in not_printed:
                        print(f"{file}\n")
            else:
                print("Podano nieprawidłowy numer drukarki.")
        except ValueError:
            print("Podano nieprawidłowy numer drukarki.")
    else:
        print("Brak dostępnych drukarek.")
