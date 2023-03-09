import os
import sys
import tkinter
from tkinter import filedialog
import openpyxl

#Überall wo die Stiftfarbe SP5 ausgewählt wurde, NICHT ändern. DIe Zeichnungen zeigen die Komplettansicht der Schürze

#HPGL Commands: SP N    Stiftfabe auswählen
#               PU      Pointer heben um neue Position anzufahren
#               PA      Position anfahren
#               PD      Pointer absetzen
#               PR      Pointer relativ zur Position in X und Y Richtung bewegen

#Zeichnung Planspanner

def planspnr(file,x):
    zplanspnr = "SP2;\n"\
        + "PU;\n" \
        + "PA " + str(x) + ", 135.00;\n" \
        + "PD;\n" \
        + "PR 0.00, 32.00;\n" \
        + "PU;\n" \
        + "PR 100.00, 0.00;\n" \
        + "PD;\n" \
        + "PR 0.00, -32.00;\n" 
    file.write(zplanspnr)
   
    
#Zeichnung Rungen

def runge(file,x):
    zrunge = "SP1;\n"\
        + "PU;\n" \
        + "PA " + str(x) + ", 135.00;\n" \
        + "PD;\n" \
        + "PR 0.00, 32.00;\n" \
        + "PU;\n" \
        + "PR 150.00, 0.00;\n" \
        + "PD;\n" \
        + "PR 0.00, -32.00;\n" 
    file.write(zrunge)     
   
    
   
#Zeichnung wo die Aluleisten getrennt werden
 
def trennung_leiste(file,x):
    ztrennung = "SP4;\n"\
        + "PU;\n" \
        + "PA " + str(x) + ", 111.00;\n" \
        + "PD;\n" \
        + "PR 0.00, 80.00;\n"
    file.write(ztrennung) 
 
    
#Zeichnung wo die Schürzen getrennt werden

def trennung_schuerze(file,x):
    ztrennung = "SP3;\n"\
        + "PU;\n" \
        + "PA " + str(x) + ", 0.00;\n" \
        + "PD;\n" \
        + "PR 0.00, 101.00;\n" 
    file.write(ztrennung) 
   

#Zeichnung für einen Custom Ausschnitt (Wird im Menü nicht genutzt)

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

#Zeichnung für die Position der Bohrung an der Alu Leiste

def bohrung(file,x):
    zbohrung = "SP4;\n"\
        + "PU;\n" \
        + "PA " + str(x+9) + ", 131.0;\n" \
        + "PD;\n" \
        + "PR 0.00, 7.00;\n" \
        + "PR 7.00, 0.00;\n" \
        + "PR 0.00, -7.00;\n" \
        + "PR -7.00, 0.00;\n" \
        + "PU;\n" \
        + "PA " + str(x+158) + ", 131.0;\n" \
        + "PD;\n" \
        + "PR 0.00, 7.00;\n" \
        + "PR 7.00, 0.00;\n" \
        + "PR 0.00, -7.00;\n" \
        + "PR -7.00, 0.00;\n" 
    file.write(zbohrung)
  
 
#Zeichnung für den Winkel an der linken Seite

def winkel_L(file,x):
    zwinkel_L = "SP3;\n"\
        + "PU;\n" \
        + "PA " + str(x) + ", 0.00;\n" \
        + "PD;\n" \
        + "PR 60.00, 100.90;\n" 
    file.write(zwinkel_L)
 
  
#Zeichnung für den Winkel an der rechten Seite

def winkel_R(file,x1,):
    zwinkel_R = "SP3;\n"\
        + "PU;\n" \
        + "PA " + str(x1) + ", 0.00;\n" \
        + "PD;\n" \
        + "PR -60.00, 100.90;\n"
    file.write(zwinkel_R)

    
#Zeichnung für den Nullpunkt (WICHTIG)

def nullpunkt(file):
    nullpunkt = "SP3;\n" \
        + "PU;\n" \
        + "PA 0.00, 0.00;\n" \
        + "PD;\n" \
        + "PR 1.00, 1.00;\n" 
    file.write(nullpunkt)

 



#Funktion einzahl: die Funktion nimmt die durch die Parameter weitergegebene exceltabelle (row) und erstellt ein .plt Dokument und schreibt die einzelnen Positionsangaben in das Dokument.  
def einzahl(row):

    #script_dir = os.path.dirname(os.path.realpath(sys.executable))

    os.chdir(script_dir)
    
    folder_path = os.path.join(script_dir, "Liner 4 Schürzen")

    file_name = row[0] + ".plt"
    
    file_path = os.path.join(folder_path, file_name)

    #print(file_path)
    

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

        if anzahl_trennungen_schuerze_NPR>=1 or anzahl_trennungen_schuerze_NPL>=1 or row[1]==None:
                abstand=3920.00
        else:
            abstand=float(row[1])
                
        if int(row[21])==1:
            if row[1]==None:
                x=abstand-float(row[2])
                winkel_L(file,x)
            else:
                winkel_L(file,0)
                
        if int(row[22])==1:
                    winkel_R(file,abstand)      
        
        if anzahl_trennungen_leiste_NPL>0:
            trennung_leiste(file,float(abstand_leiste_schuerze_NPL))
            bohrung(file,float(abstand_leiste_schuerze_NPL))
            for i in range(anzahl_trennungen_leiste_NPL):
                pos_trennungen_leiste_npl[i]=float(pos_trennungen_leiste_npl[i])+abstand_leiste_schuerze_NPL 
               
                trennung_leiste(file,float(pos_trennungen_leiste_npl[i]))

        if anzahl_trennungen_leiste_NPR>0:
            trennung_leiste(file,float(abstand-abstand_leiste_schuerze_NPR))
            bohrung(file,float(abstand-abstand_leiste_schuerze_NPR-175.0))
            for i in range(anzahl_trennungen_leiste_NPR):
                pos_trennungen_leiste_npr[i] = abstand-float(pos_trennungen_leiste_npr[i])-abstand_leiste_schuerze_NPR
             
                trennung_leiste(file,float(pos_trennungen_leiste_npr[i]))

        if anzahl_trennungen_schuerze_NPL>0:
            for i in range(anzahl_trennungen_schuerze_NPL):
                trennung_schuerze(file,float(pos_trennungen_schuerze_npl[i]))

        if anzahl_trennungen_schuerze_NPR>0:
            for i in range(anzahl_trennungen_schuerze_NPR):
                pos_trennungen_schuerze_npr[i] = abstand-float(pos_trennungen_schuerze_npr[i])
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
            

        
        

#Funktion PLT: öffnet das Excel Dokument und schreibt die einzelnen Zellen in die Variablen.     


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
       
        if row[0]:
            einzahl(row)
            #mehrzahl(row)
        else:
            break
            
#öffnet ein dialogfenster wo du die Excleldatei auswählen kannst

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
    #exe_path = os.path.dirname(os.path.realpath(sys.executable))
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
