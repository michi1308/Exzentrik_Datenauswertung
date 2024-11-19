
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Funktion, um die Excel-Datei zu analysieren
def analyze_file(path, index, side, output_widget):
    try:
        # Excel-Datei einlesen
        df_isokin_links = pd.read_excel(path, sheet_name='Isokin_Exz_Kon_60_60_Links')
        df_isokin_rechts = pd.read_excel(path, sheet_name='Isokin_Exz_Kon_60_60_Rechts')

        # Wähle die richtige Spalte basierend auf der Auswahl
        if side == "Links":
            data = df_isokin_links.iloc[:, 1]  # Spalte B für Winkel links
        else:
            data = df_isokin_rechts.iloc[:, 1]  # Spalte B für Winkel rechts

        # Peaks suchen
        left_peak, right_peak = find_neighboring_peaks_with_plateaus(data, index)

        # Ergebnis in das Text-Widget schreiben
        output_to_widget(output_widget, f"Linker benachbarter Peak (Index, Wert): {left_peak[0]}, {left_peak[1]}")
        output_to_widget(output_widget, f"Rechter benachbarter Peak (Index, Wert): {right_peak[0]}, {right_peak[1]}")

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler bei der Analyse: {e}")

# Funktion zur Suche nach benachbarten Peaks mit Plateaus
def find_neighboring_peaks_with_plateaus(data, peak_index):
    def find_peak_in_direction(data, start_index, step):
        for idx in range(start_index, len(data) if step > 0 else -1, step):
            if idx == 0 or idx == len(data) - 1:
                if (idx == 0 and data[idx] <= data[idx + 1]) or \
                   (idx == len(data) - 1 and data[idx] <= data[idx - 1]) or \
                   (idx == 0 and data[idx] >= data[idx + 1]) or \
                   (idx == len(data) - 1 and data[idx] >= data[idx - 1]):
                    return idx, data[idx]
            if (data[idx] > data[idx - 1] and data[idx] >= data[idx + 1]) or \
               (data[idx] < data[idx - 1] and data[idx] <= data[idx + 1]):
                plateau_start = idx
                plateau_end = idx
                while plateau_end + step >= 0 and plateau_end + step < len(data) and \
                      data[plateau_end + step] == data[idx]:
                    plateau_end += step
                return plateau_start, data[idx]
        return None

    left_peak = find_peak_in_direction(data, peak_index - 1, -1)
    right_peak = find_peak_in_direction(data, peak_index + 1, 1)
    return left_peak, right_peak

# Funktion zur Ausgabe in das Text-Widget
def output_to_widget(text_widget, message):
    text_widget.insert(tk.END, message + "\n")
    text_widget.see(tk.END)
    text_widget.update_idletasks()

# Funktion zur Dateiauswahl
def select_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)


# GUI erstellen
def main():
    root = tk.Tk()
    root.title("Isokinetik: ROM nachberechnen.")

    # Fenstergröße festlegen
    root.geometry("680x500")  # Breiter und höheres Fenster

    # Pfadeingabe
    tk.Label(root, text="Pfad zur Excel-Datei:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    path_entry = tk.Entry(root, width=70)  # Breiteres Eingabefeld
    path_entry.grid(row=0, column=1, padx=10, pady=5)
    browse_button = tk.Button(root, text="Durchsuchen", command=lambda: select_file(path_entry))
    browse_button.grid(row=0, column=2, padx=10, pady=5)

    # Index-Eingabe
    tk.Label(root, text="Index für den Peak:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    index_entry = tk.Entry(root, width=10)
    index_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")

    # Dropdown-Menü zur Seitenauswahl
    tk.Label(root, text="Seite wählen:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    side_var = tk.StringVar(value="Links")
    side_dropdown = tk.OptionMenu(root, side_var, "Links", "Rechts")
    side_dropdown.grid(row=2, column=1, padx=7, pady=5, sticky="w")

    # Text-Widget zur Ausgabe
    output_text = tk.Text(root, wrap="word", height=20, width=80)  # Größeres Text-Widget
    output_text.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

    # Start-Button zur Analyse
    start_button = tk.Button(root, text="Starten", command=lambda: analyze_file(
        path_entry.get(), int(index_entry.get()), side_var.get(), output_text))
    start_button.grid(row=3, column=0, columnspan=3, pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()