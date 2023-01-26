import os
import tkinter
from tkinter import filedialog
import openpyxl
import sys

def nullpunkt(file):
    znullpunkt = "SP1;\n" \
        + "PU;\n" \
        + "PA 0.00, 0.00;\n" \
        + "PD;\n" \
        + "PR 1.00, 1.00;\n" 
    file.write(znullpunkt)

def halterung(file,x):
    zhalterung = "PU;\n" \
        + "PA " + str(x) + ", 0.00;\n" \
        + "PD;\n" \
        + "PR 0.00, 30.00;\n" 
    file.write(zhalterung)

def trennungen(file,x):
    ztrennungen = "PU;\n" \
        + "PA " + str(x) + ", 0.00;\n" \
        + "PD;\n" \
        + "PR 0.00, 70.00;\n" 
    file.write(ztrennungen)

def laenge_schuerze(file,x):
    zlaenge_schuerze = "PU;\n" \
        + "PA " + str(x) + ", 0.00;\n" \
        + "PD;\n" \
        + "PR 0.00, 150.00;\n" 
    file.write(zlaenge_schuerze)

def mehrzahl(row):
    

    os.chdir(script_dir)
    
    folder_path = os.path.join(script_dir, "Liner 5 Schürzen")

    file_name = row[0] + ".plt"

    file_path = os.path.join(folder_path, file_name)

    with open(file_path, "w") as file:

        nullpunkt(file)

        if row[1]:
            anzahl_halterungen=int(row[1])
            for i in range (anzahl_halterungen):
                    halterung(file,float(pos_halterungen[i]))
        
        if row[3]:
            anzahl_trennungen=int(row[3])
            for i in range (anzahl_trennungen):
                trennungen(file,float(pos_trennungen[i]))

        if row[5]:
            laenge_schuerze(file,float(row[5]))

    


def PLT(file):
# Öffnen Sie die ausgewählte Excel-Datei mit openpyxl
    wb = openpyxl.load_workbook(file)

        # Machen Sie etwas mit der Excel-Datei (z.B. Werte in einer Zelle ändern)
    ws = wb[wb.sheetnames[0]]

    row_count = 0
    for rows in ws.rows:
        #print(rows)
        if row_count == 0:  # Skip the first row
            row_count += 1
            continue
        
        global pos_halterungen
        global pos_trennungen

        pos_halterungen = str(rows[2].value).split(',')
        pos_trennungen = str(rows[4].value).split(',')

        row = []
        for cell in rows:
            row.append(cell.value)
       

        if row[0]:
            mehrzahl(row)
        else:
            break

def upload_file(filepath):
    # Print the path of the selected file
    global script_dir
    if os.path.splitext(sys.argv[0])[1] == '.exe':
        script_dir = os.path.dirname(os.path.realpath(sys.executable))
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))

    if os.path.exists(filepath) and os.path.isfile(filepath):
        PLT(filepath)

def upload_file_manuell():
    # Open the file dialog and get the selected file
    global script_dir
    if os.path.splitext(sys.argv[0])[1] == '.exe':
        script_dir = os.path.dirname(os.path.realpath(sys.executable))
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))

    filepath = filedialog.askopenfile(filetypes=[('Excel-Dateien', '*.xlsx')], initialdir = script_dir)
    # Print the path of the selected file
    if os.path.exists(filepath.name) and os.path.isfile(filepath.name):
        PLT(filepath.name)

if __name__ == "__main__":
    upload_file_manuell()
    sys.exit()
