# Programm zur Verwaltung von Ratenzahlungen
# zum Windows-EXE packen:
# pyinstaller.exe --hidden-import babel.numbers --onefile dateiname.py

# from tkinter import INSERT, font
from tkinter.constants import SUNKEN
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import tkinter as tk
from tkinter import Menu, messagebox, filedialog, Tk
from tkcalendar import DateEntry
import xlsxwriter
import os


# Erstelle Klasse die von tkinter.Frame erbt
# Frame ist ein Rahmen-Widget und ist der Ausgangspunkt für weitere
# Steuerelemente in der Grafischen Oberfläche
class myApp(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        # Die Methode initUI wird vom Konstruktor initialisiert für die weiteren Steuerelemente
        self.pack()
        self.initUI()

    def initUI(self):
        # Legt Titel des Fensters fest
        self.master.title("Ratenzahlung")
        self.text = tk.StringVar()
        self.text.set("Keine Datei geladen!")
        # Menüleiste erstellen
        self.menubar = Menu(self.master)
        self.master.config(menu=self.menubar)
        self.fileMenu = Menu(self, tearoff=False)
        self.menubar.add_cascade(label="Datei", menu=self.fileMenu)
        self.fileMenu.add_command(label="Öffnen", command=self.openFile)
        # fileMenu.add_command(label="Speichern als ...", command=self.saveFile)
        self.fileMenu.add_separator()
        self.fileMenu.add_command(label="Beenden", command=self.quit)
        self.helpMenu = tk.Menu(self, tearoff=False)
        self.menubar.add_cascade(label="Hilfe", menu=self.helpMenu)
        self.helpMenu.add_command(label="Hilfe Index", command=self.helpText)
        self.helpMenu.add_command(label="Über", command=self.about)

        # Datumseingabe mit Kalender Widget anzeigen
        self.cal = DateEntry(self, locale='en_US', date_pattern='dd-MM-yyyy')
        self.cal.grid(row=0, column=1)
        self.uebernehmen_button = tk.Button(self, text="Datum/Bearbeiter übernehmen", command=self.get_my_date, width=25)
        self.uebernehmen_button.grid(row=5, column=1, pady=5)

        self.bearbeiter_label = tk.Label(self, text="Bearbeiter Kürzel:")
        self.bearbeiter_label.grid(row=4, column=1, pady=5)
        self.bearbeiter_eingabe = tk.Entry(self)
        self.bearbeiter_eingabe.grid(row=4, column=2, pady=5)

        self.vorabliste_button = tk.Button(self, text="Vorabliste ausgeben", command=lambda: self.vorabListe(False), width=25)
        self.vorabliste_button.grid(row=6, column=1, pady=5)
        self.alletilgungen_button = tk.Button(self, text="Vollständige Liste ausgeben", command=lambda: self.vorabListe(True), width=25)
        self.alletilgungen_button.grid(row=7, column=1, pady=5)
        self.verarbeiten_button = tk.Button(self, text="Ratenzahlungen verarbeiten", command=self.verarbeiteRaten, width=25)
        self.verarbeiten_button.grid(row=8, column=1, pady=5)
        # status bar
        self.statusbar = tk.Label(self, textvariable=self.text, bd=1, relief=tk.SUNKEN, anchor=tk.CENTER, width=25)
        self.statusbar.grid(row=10, column=1, pady=15)

    def get_my_date(self):
        # hole Datum und wandle in TMJ um
        self.datum1 = self.cal.get_date()
        self.datum = self.datum1.strftime("%d-%m-%Y")
        self.bearbeiter = self.bearbeiter_eingabe.get()
        # Ratenzahlungen.datum = self.datum
        print(self.datum1, self.bearbeiter)
        return(self.datum1)

    def openFile(self):
        # Funktion zum Anzeigen eines Datei-Dialogs und Laden der Ratenzahlung-Kartei im Excel-Format.
        self.filename = filedialog.askopenfilename(title='Datei Öffnen', initialdir=os.getcwd(), filetypes=[("Excel files", ".xlsx")])
        # messagebox.showinfo(title='Gewählte Datei', message=filename)
        # versuche Datei zu öffnen
        try:
            self.ratenzahlungen = pd.read_excel(self.filename, engine='openpyxl', sheet_name=None)
            self.text.set("Ratenzahlungen geladen!")
        except:
            messagebox.showerror(title='Datei nicht gefunden!', message=self.filename)

    def vorabListe(self, alle):
        '''
        Funktion um eine Liste der anstehenden Ratenzahlungen zum Stichtag als Excel-Datei auszugeben.
        Dateiname wird festgelegt, ob alle oder nur aktuelle Raten ausgewählt wurden.
        '''
        self.datum = self.datum1.strftime("%d-%m-%Y")
        self.alle = alle
        print(self.datum, "Alle: ", self.alle)

        self.liste = pd.DataFrame(columns=['Name', 'Vorname', 'Empfaenger', 'IBAN', 'Ref1', 'Ref2', 'Rest', 'Rate', 'RestNeu'])
        # mdatum wandelt str in Format datetime YYYY-MM-DD HH:MM:SS
        self.mdatum = datetime.strptime(self.datum, "%d-%m-%Y")
        # vdatum wandelt in Format Str DD-MM-YYYY
        vdatum = self.mdatum.strftime('%d-%m-%Y')
        print('Verarbeitungsdatum: ' + vdatum)
        keys = self.ratenzahlungen.keys()
        # erzeuge Liste aller Tilgungen aus einzelnen Blättern der Excel Datei
        # ausser denen die noch nicht begonnen haben
        for i in keys:
            print('verarbeite Blatt: ' + i)
            self.blatt = self.ratenzahlungen.get(i)
            # überspringe Blatt wenn Datum vor Ratenbeginn liegt
            if self.blatt.iloc[5, 1] > self.mdatum and self.alle is False:
                print(i, '; Beginn: ', self.blatt.iloc[5, 1], 'vor Verarbeitungsdatum: ', self.mdatum)
            else:
                zeile = self.erzeugeListe()
                # pandas append deprecated with Pandas 1.4.0 use concat instead
                # self.liste = self.liste.append(zeile)
                self.liste = pd.concat([self.liste, zeile])

        # Sortiere Liste nach Alphabet
        sortliste = self.liste.sort_values(by='Name')
        # Passe Name der Datei an
        if self.alle is True:
            nameVorabliste = ('Liste alle Ratenzahlungen ' + str(vdatum) + '.xlsx')
        else:
            nameVorabliste = ('Vorabliste Ratenzahlungen ' + str(vdatum) + '.xlsx')
        # sortliste.to_excel(nameVorabliste, sheet_name='Liste Tilgungen', index=None)
        # Neue Excel-Arbeitsmappe mit Pandas Excelwriter anlegen
        Excelwriter = pd.ExcelWriter(nameVorabliste, engine="xlsxwriter")
        # Schreibt die sortierte Liste in Arbeitsblatt in Excel-Arbeitsmappe
        sortliste.to_excel(Excelwriter, sheet_name='Liste Ratenzahlungen', index=None)
        # Zugriff auf Excelwriter Blätter ermöglichen
        workbook = Excelwriter.book
        worksheet = Excelwriter.sheets['Liste Ratenzahlungen']
        # Zellenformate für Excel-Arbeitsmappe festlegen
        numFormat = workbook.add_format({'num_format': '#,##0.00'})
        dateFormat = workbook.add_format({'num_format': 'dd/mm/yyyy'})
        # Format und Zellenbreite in Excel-Arbeitsblatt festlegen
        worksheet.set_column('A:A', 20, None)
        worksheet.set_column('B:B', 18, None)
        worksheet.set_column('C:C', 28, None)
        worksheet.set_column('D:D', 26, None)
        worksheet.set_column('E:E', 20, None)
        worksheet.set_column('F:F', 20, None)
        worksheet.set_column('G:G', 13, numFormat)
        worksheet.set_column('H:H', 13, numFormat)
        worksheet.set_column('I:I', 13, numFormat)
        worksheet.set_column('J:J', 12, None)
        worksheet.set_column('K:K', 13, dateFormat)
        # Schreibe Excel-Datei
        Excelwriter.save()

    def erzeugeListe(self):
        # verwende iloc für positionsbasierte Indizierung
        '''
        Funktion um Daten aus einem Excel-Blatt zu holen. Wird an die Funktion als Pandas-Dataframe
        übergeben und Name, Vorname, Rate, Empfänger, IBAN, Ref1, Ref2 werden in eine Zeile als
        Pandas-DataFrame geschrieben und von der Funktion zurückgegeben.
        iloc dient der positionsbasierten Indizierung innerhalb des DataFrames.
        Das Schema ist erste Zahl gibt die Zeile, die zweite die Spalte an.
        In Python wird immer ab 0 angefangen zu zählen, d.h. die erste Zeile ist 0!
        '''
        # Hole Text/Zahlen aus der Tabelle anhand ihrer Zellen Position
        self.name = self.blatt.iloc[1, 1]
        self.vorname = self.blatt.iloc[2, 1]
        self.gesamt = float(self.blatt.iloc[3, 1])
        self.rate = float(self.blatt.iloc[4, 1])
        self.beginn = self.blatt.iloc[5, 1]
        # Beginn in passendes Format für Liste umwandeln
        self.dbeginn = self.beginn.strftime('%d.%m.%Y')
        self.empfaenger = self.blatt.iloc[1, 3]
        self.iban = self.blatt.iloc[3, 3]
        self.ref1 = self.blatt.iloc[5, 3]
        self.ref2 = self.blatt.iloc[6, 3]
        self.aussetzen = self.blatt.iloc[6, 1]
        self.rest = float(self.blatt.iloc[7, 1])
        # Berechne Rest-Betrag
        if (self.rate >= self.rest):
            self.rate = self.rest
            self.restNeu = self.rest - self.rate
        else:
            self.restNeu = self.rest - self.rate
        # erzeuge neue Tabelle (Dataframe) mit Informationen für die Liste
        self.ratenzahlungen1 = pd.DataFrame([[self.name, self.vorname, self.empfaenger,
                                              self.iban, self.ref1, self.ref2, self.rest,
                                              self.rate, self.restNeu, self.aussetzen,
                                              self.dbeginn]], columns=['Name', 'Vorname',
                                                                       'Empfaenger', 'IBAN', 'Ref1', 'Ref2', 'Rest',
                                                                       'Rate', 'RestNeu', 'Aussetzen', 'Beginn'])
        return(self.ratenzahlungen1)

    def verarbeiteRaten(self):
        '''
        Funktion zum Verarbeiten der Ratenzahlungen. Alle gültigen Ratenzahlungen werden
        zum Verarbeitungsdatum an die Karteiblätter angehängt und die Rest-Rate
        angepasst.
        '''
        self.datum = self.datum1.strftime("%d-%m-%Y")
        print(self.datum)
        # self.ratenz.tilgungen_verarbeiten()

        # mdatum wandelt str in Format datetime YYYY-MM-DD HH:MM:SS
        mdatum = datetime.strptime(self.datum, "%d-%m-%Y")
        # vdatum wandelt in Format Str DD-MM-YYYY
        vdatum = mdatum.strftime('%d-%m-%Y')
        # ddatum wandelt in Format Str DD/MM/YYYY
        ddatum = mdatum.strftime('%d.%m.%Y')
        print('Verarbeitungsdatum: ' + vdatum)

        # Blattnamen sind die keys aus dem dictionary
        keys = self.ratenzahlungen.keys()
        ratenzahlungenNeu = self.ratenzahlungen.copy()
        for i in keys:
            print('verarbeite Blatt: ' + i)
            blatt = ratenzahlungenNeu.get(i)
            # überspringe Blatt wenn Datum vor Ratenbeginn liegt oder Rate aussetzen auf 'j' ist
            if blatt.iloc[5, 1] > mdatum:
                print(i, '; Beginn: ', blatt.iloc[5, 1], 'vor Verarbeitungsdatum: ', mdatum)
            elif blatt.iloc[6, 1] == 'j':
                print(i, ' Rate ist ausgesetzt')
            else:
                # hole Blatt i, ändere RestNeu und hänge aktuelle Tilgungszeile an
                blattNeu = ratenzahlungenNeu.get(i)
                blattNeu = self.schreibeTilgungen(blattNeu, ddatum)
                ratenzahlungenNeu[i] = blattNeu

        # Erzeuge neue Excel Datei mit allen aktualisierten Blättern
        dateiName = ('Ratenzahlungen_neu_' + str(vdatum) + '.xlsx')
        Excelwriter = pd.ExcelWriter(dateiName, engine="xlsxwriter")
        for i in keys:
            print('schreibe Blatt: ' + i)
            # Die Methode get holt sich das Blatt aus dem Dictionary
            blatt = ratenzahlungenNeu.get(i)
            beginn = blatt.iloc[5, 1]
            blatt.to_excel(Excelwriter, sheet_name=i, index=False)
            # Zugriff auf Excelwriter Blätter ermöglichen
            workbook = Excelwriter.book
            worksheet = Excelwriter.sheets[i]
            # Zellenformate festlegen
            numFormat = workbook.add_format({'num_format': '#,##0.00'})
            dateFormat = workbook.add_format({'num_format': 'dd/mm/yyyy'})
            # Format und Zellenbreite festlegen
            worksheet.set_column('A:A', 20, dateFormat)
            # schreibe Verarbeitungsdatum in letzte Zeile mit Datumsformat
            # worksheet.write_datetime(letzteZeile, 0, mdatum, dateFormat)
            worksheet.write('B7:B7', beginn, dateFormat)
            worksheet.set_column('B5:B6', 15, numFormat)
            worksheet.set_column('C:C', 15, numFormat)
            worksheet.set_column('D:D', 34, None)
        # Speichern der Excel Datei
        Excelwriter.save()

    def schreibeTilgungen(self, blatt, ddatum):
        '''
        Funktion um gebuchte und überwiesene Raten in die Blätter zu schreiben.
        iloc wird für die positionsbasierte Indizierung innerhalb des DataFrames
        verwendet. Das Schema ist erste Zahl gibt die Zeile, die zweite die Spalte an.
        In Python wird immer ab 0 angefangen zu zählen, d.h. die erste Zeile ist 0!
        '''
        # Hole Text/Zahlen aus der Tabelle anhand ihrer Zellen Position
        rate = float(blatt.iloc[4, 1])
        rest = float(blatt.iloc[7, 1])
        # Prüfen ob Rate größer als Rest. Ist Rate größer als Rest, dann neuer Rest gleich Rate
        if (rate >= rest):
            rate = rest
            restNeu = rest - rate
        else:
            restNeu = rest - rate
        # ändere restNeu auf Fliesskommazahl
        blatt.iloc[7, 1] = float(restNeu)
        # hänge Zeile mit Datum, Rate, RestNeu, Bearbeiter unten an die Tabelle an
        blatt.loc[len(blatt.index)] = [ddatum, rate, restNeu, self.bearbeiter]
        return(blatt)

    def saveFile(self):
        # Funtkion wird noch nicht benutzt
        self.savename = filedialog.asksaveasfilename(initialdir="/", title="Save as", filetypes=[("Excel files", "*.xlsx")])

    def helpText(self):
        hilfetext = '''Die Ratenzahlungen müssen in einer Excel Datei vorliegen.\n\nVorabliste: Druckt eine Liste der Ratenzahlungen zum Ausführungsdatum.\n\nListe aller Ratenzahlungen: Druckt alle erfassten Ratenzahlungen.\n\nRatenzahlungen verarbeiten: Schreibt Zahlungen zum Ausführungsdatum in Excel-Datei.'''
        self.helpWin = tk.Toplevel()
        self.helpWin.wm_title("Hilfe")
        self.helpWin.wm_geometry("650x250")
        text = tk.Text(self.helpWin, height=10, bg="light yellow")
        text.configure(font=("Arial", 11, "normal"))
        text.insert(tk.INSERT, hilfetext)
        text.pack(side=tk.TOP, expand=True, fill=tk.BOTH)
        b = tk.Button(self.helpWin, text="Ok", command=self.helpWin.destroy)
        b.pack()

    def about(self):
        messagebox.showinfo("Ratenzahlung", "Programm zur Verwaltung der Ratenzahlungen")


# Starte Programm
if __name__ == "__main__":
    root = Tk()
    root.geometry("400x300")
    app = myApp(root)
    app.mainloop()
