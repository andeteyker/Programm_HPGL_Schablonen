import os
import tkinter
from tkinter import filedialog
import openpyxl

def planspnr(file,x):
    zplanspnr = "PU;\n" \
        + "PA " + str(x) + ", 62.00;\n" \
        + "PD;\n" \
        + "PR 100.00, 0.00;\n" \
        + "PR 0.00, 39.00;\n" \
        + "PR -100.00, 0.00;\n" \
        + "PR 0.00, -39.00;\n" 
    file.write(zplanspnr)
    
def runge(file,x):
    zrunge = "PU;\n" \
        + "PA " + str(x) + ", 103.00;\n" \
        + "PD;\n" \
        + "PR 150.00, 0.00;\n" \
        + "PR 0.00, 50.00;\n" \
        + "PR -150.00, 0.00;\n" \
        + "PR 0.00, -50.00;\n" 
    file.write(zrunge)     
    
def trennung_leiste(file,x):
    ztrennung = "PU;\n" \
            + "PA " + str(x) + ", 86.00;\n" \
            + "PD;\n" \
            + "PR 0.00, 44.00;\n"
    file.write(ztrennung) 

def trennung_schuerze(file,x):
    ztrennung = "PU;\n" \
            + "PA " + str(x) + ", 0.00;\n" \
            + "PD;\n" \
            + "PR 0.00, 86.00;\n" 
    file.write(ztrennung) 

def ausschnitt(file,x,y,l,h):
    zausschnitt = "PU;\n" \
        + "PA " + str(x) + ", "+ str(y+10) +";\n" \
        + "PD;\n" \
        + "PR 0.00, -10.00;\n" \
        + "PR 10.00, 0.00;\n" \
        + "PU;\n"\
        + "PA " + str(x+l) + ", "+ str(y+10) +";\n"\
        + "PD;\n" \
        + "PR 0.00, -10.00;\n" \
        + "PR -10.00, 0.00;\n" \
        + "PU;\n"\
        + "PA " + str(x+l) + ", "+ str(y+h-10) +";\n"\
        + "PD;\n" \
        + "PR 0.00, 10.00;\n" \
        + "PR -10.00, 0.00;\n" \
        + "PU;\n"\
        + "PA " + str(x) + ", "+ str(y+h-10) +";\n"\
        + "PD;\n" \
        + "PR 0.00, 10.00;\n" \
        + "PR 10.00, 0.00;\n" 
    file.write(zausschnitt)  

def bohrung(file,x):
    zbohrung = "PU;\n" \
        + "PA " + str(x+9) + ", 116.60;\n" \
        + "PD;\n" \
        + "PR 0.00, 7.00;\n" \
        + "PR 7.00, 0.00;\n" \
        + "PR 0.00, -7.00;\n" \
        + "PR -7.00, 0.00;\n" \
        + "PU;\n" \
        + "PA " + str(x+159) + ", 116.60;\n" \
        + "PD;\n" \
        + "PR 0.00, 7.00;\n" \
        + "PR 7.00, 0.00;\n" \
        + "PR 0.00, -7.00;\n" \
        + "PR -7.00, 0.00;\n" 
    file.write(zbohrung)

def winkel_L(file):
    zwinkel_L = "PU;\n" \
        + "PA 0.00, 0.00;\n" \
        + "PD;\n" \
        + "PR 53.00, 100.00;\n" 
    file.write(zwinkel_L)

def winkel_R(file,x):
    zwinkel_R = "PU;\n" \
        + "PA " + str(x) + ", 0.00;\n" \
        + "PD;\n" \
        + "PR -53.00, 100.00;\n"
    file.write(zwinkel_R)

def nullpunkt(file):
    nullpunkt = "SP1;\n" \
        + "PU;\n" \
        + "PA 0.00, 0.00;\n" \
        + "PD;\n" \
        + "PR 1.00, 1.00;\n" 
    file.write(nullpunkt)

def einzahl(row):

    script_dir = os.path.dirname(os.path.abspath(__file__))
    folder_path = os.path.join(script_dir, "Liner 4 Schürzen")
    file_name = row[0] + ".plt"
    file_path = os.path.join(folder_path, file_name)

    with open(file_path, "w") as file:
    
        anzahl_trennungen_schuerze_NPL = int(row[13])

        anzahl_trennungen_schuerze_NPR = int(row[15])
        
        anzahl_trennungen_leiste_NPL = int(row[17])

        anzahl_trennungen_leiste_NPR = int(row[19])
        
        anzahl_rungen_NPL = int(row[7])
    
        anzahl_rungen_NPR = int(row[9])
        
        anzahl_planspnr_NPL = int(row[3])

        anzahl_planspnr_NPR = int(row[5])

        if row[11]:
            abstand_leiste_schuerze_NPL = float(row[11])

        if row[12]:
            abstand_leiste_schuerze_NPR = float(row[12])

        nullpunkt(file)

        if anzahl_trennungen_schuerze_NPR>=1 or anzahl_trennungen_schuerze_NPL>=1:
                abstand=3700.00
        else:
            abstand=float(row[2])
                
        if anzahl_trennungen_leiste_NPL>0:
            trennung_leiste(file,float(abstand_leiste_schuerze_NPL))
            bohrung(file,float(abstand_leiste_schuerze_NPL))
            for i in range(anzahl_trennungen_leiste_NPL):
                pos_trennungen_leiste_npl[i]=float(pos_trennungen_leiste_npl[i])+abstand_leiste_schuerze_NPL
                trennung_leiste(file,float(pos_trennungen_leiste_npl[i]))  

        if anzahl_trennungen_leiste_NPR>0:
            
            trennung_leiste(file,float(abstand-abstand_leiste_schuerze_NPR))
            bohrung(file,float(abstand-abstand_leiste_schuerze_NPR-171.50))
            for i in range(anzahl_trennungen_leiste_NPR):
                pos_trennungen_leiste_npr[i] = abstand-float(pos_trennungen_leiste_npr[i])-abstand_leiste_schuerze_NPR
                trennung_leiste(file,float(pos_trennungen_leiste_npr[i]))

        if anzahl_trennungen_schuerze_NPL>0:
            for i in range(anzahl_trennungen_schuerze_NPL):
                trennung_schuerze(file,float(pos_trennungen_schuerze_npl[i]))

        if anzahl_trennungen_schuerze_NPR>0:
            for i in range(anzahl_trennungen_schuerze_NPR):
                pos_trennungen_schuerze_npr[i] = 3700.00-float(pos_trennungen_schuerze_npr[i])
                trennung_schuerze(file,float(pos_trennungen_schuerze_npr[i]))

        if anzahl_rungen_NPL>0:
            for i in range(anzahl_rungen_NPL):
                pos_rungen_npl[i]=float(pos_rungen_npl[i])+abstand_leiste_schuerze_NPL
                runge(file,float(pos_rungen_npl[i]))

        if anzahl_rungen_NPR>0:
            for i in range(anzahl_rungen_NPR):
                pos_rungen_npr[i] = abstand-float(pos_rungen_npr[i])-abstand_leiste_schuerze_NPR-150.00
                runge(file,float(pos_rungen_npr[i]))
                
        if anzahl_planspnr_NPL>0: 
            for i in range(anzahl_planspnr_NPL):
                planspnr(file,float(pos_planspanner_npl[i]))

        if anzahl_planspnr_NPR>0:
            for i in range(anzahl_planspnr_NPR):
                pos_planspanner_npr[i] = abstand-float(pos_planspanner_npr[i])-100.00
                planspnr(file,float(pos_planspanner_npr[i]))
            
        if int(row[21])==1:
            winkel_L(file)

        if int(row[22])==1:
            winkel_R(file,abstand)


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
        
        global pos_planspanner_npl, pos_planspanner_npr, pos_rungen_npl, pos_rungen_npr, pos_trennungen_schuerze_npl, pos_trennungen_schuerze_npr, pos_trennungen_leiste_npl, pos_trennungen_leiste_npr

        pos_planspanner_npl, pos_planspanner_npr, pos_rungen_npl, pos_rungen_npr, pos_trennungen_schuerze_npl, pos_trennungen_schuerze_npr, pos_trennungen_leiste_npl, pos_trennungen_leiste_npr = [
        str(rows[i].value).split(',') if str(rows[i].value) and not str(rows[i].value).isspace() else []
        for i in [4, 6, 8, 10, 14, 16, 18, 20]
    ]
        row = []
        for cell in rows:
            row.append(cell.value)
        print(row)
        if row[0]:
            einzahl(row)
            #mehrzahl(row)
        else:
            break
            
def upload_file(filepath):
    # Print the path of the selected file
    if os.path.exists(filepath) and os.path.isfile(filepath):
        PLT(filepath)

def upload_file_manuell():
    # Open the file dialog and get the selected file
    filepath = filedialog.askopenfile(filetypes=[('Excel-Dateien', '*.xlsx')])
    # Print the path of the selected file
    if os.path.exists(filepath.name) and os.path.isfile(filepath.name):
        PLT(filepath.name)

if __name__ == "__main__":
    upload_file_manuell()
    #exit()
#quit