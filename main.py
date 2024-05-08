import os
import subprocess
import win32print
import win32api


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

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        i = 1

        if filename.lower().endswith('.pdf'):
            print(f"{i}. Drukowanie pliku {filename}")
            try:
                subprocess.run(["lp", "-d", selected_printer, file_path])
                print(f"{i}. Plik został wydrukowany pomyślnie")
                i += 1
            except Exception as e:
                print(f"{i}. Wystąpił błąd podczas drukowania pliku {filename}")
                print(e)
                i += 1


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
            print_pdf_files(current_directory, selected_printer)
        else:
            print("Podano nieprawidłowy numer drukarki.")
    except ValueError:
        print("Podano nieprawidłowy numer drukarki.")
else:
    print("Brak dostępnych drukarek.")
