

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from scipy.signal import find_peaks

def output_to_widget(text_widget, message):
    text_widget.insert(tk.END, message + "\n")
    text_widget.see(tk.END)
    text_widget.update_idletasks()

def analyze_file(file_path, text_widget):
    try:
        # Excel-Datei einlesen
        df_isokin_links = pd.read_excel(file_path, sheet_name='Isokin_Exz_Kon_60_60_Links')
        df_isokin_rechts = pd.read_excel(file_path, sheet_name='Isokin_Exz_Kon_60_60_Rechts')

        # Drehmoment und Winkel extrahieren
        drehmoment_links = df_isokin_links.iloc[:, 2]  # Spalte C
        drehmoment_rechts = df_isokin_rechts.iloc[:, 2]  # Spalte C
        winkel_links = df_isokin_links.iloc[:, 1]  # Spalte B
        winkel_rechts = df_isokin_rechts.iloc[:, 1]  # Spalte B

        # Peaks finden mit Prominenz von 50
        peaks_links, _ = find_peaks(drehmoment_links, prominence=50)
        peaks_rechts, _ = find_peaks(drehmoment_rechts, prominence=50)

        # Ergebnisse sammeln
        resultate_links = pd.DataFrame({
            'Index': peaks_links,
            'Winkel': winkel_links.iloc[peaks_links].values,
            'Drehmoment': drehmoment_links.iloc[peaks_links].values
        })

        resultate_rechts = pd.DataFrame({
            'Index': peaks_rechts,
            'Winkel': winkel_rechts.iloc[peaks_rechts].values,
            'Drehmoment': drehmoment_rechts.iloc[peaks_rechts].values
        })

        # Ergebnisse in das Text-Widget ausgeben
        output_to_widget(text_widget, "Peaks für das linke Bein:")
        output_to_widget(text_widget, resultate_links.to_string(index=False))
        output_to_widget(text_widget, "\nPeaks für das rechte Bein:")
        output_to_widget(text_widget, resultate_rechts.to_string(index=False))

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Verarbeiten der Datei: {e}")

def select_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

def main():
    # GUI erstellen
    root = tk.Tk()
    root.title("Isokinetik-Nachbearbeitung: Index, Winkel und Drehmomente berechnen")

    # Datei-Pfad-Eingabe
    frame = tk.Frame(root)
    frame.pack(padx=10, pady=5)

    tk.Label(frame, text="Pfad zur Excel-Datei:").pack(side=tk.LEFT)
    path_entry = tk.Entry(frame, width=50)
    path_entry.pack(side=tk.LEFT, padx=(5, 0))

    # Datei durchsuchen Button
    browse_button = tk.Button(frame, text="Durchsuchen", command=lambda: select_file(path_entry))
    browse_button.pack(side=tk.LEFT, padx=(5, 0))

    # Start-Button
    start_button = tk.Button(root, text="Starten", command=lambda: analyze_file(path_entry.get(), output_text))
    start_button.pack(pady=5)

    # Text-Widget für die Ausgabe
    output_text = tk.Text(root, wrap="word", height=25, width=70)
    output_text.pack(padx=10, pady=10)

    # GUI ausführen
    root.mainloop()

if __name__ == "__main__":
    main()

