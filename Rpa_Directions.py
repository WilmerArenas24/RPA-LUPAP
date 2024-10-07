import subprocess
import pyautogui as pt
import pandas as pd
import time
import pyperclip
import re
import os
import math

# Especifica la ruta del ejecutable de Chrome
chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"  # Ajusta la ruta si es necesario

# Abre Google usando Chrome
subprocess.Popen([chrome_path, 'https://www.google.com'])

coord = []  # Lista para almacenar coordenadas

# Cambia al directorio donde está el archivo Excel
os.chdir(r"C:\Users\wilme\Documents\Dir")  # Directorio del archivo
data = pd.read_excel("Direcciones.xlsx")
data["Buscar"] = data["Departamento"] + " " + data["Provincia"] + " " + data["Distrito"] + " " + data["Dirección"]
buscar = data["Buscar"].values

def isnan(value):
    try:
        return math.isnan(float(value))
    except:
        return False

p1 = (746, 448)  # Click en buscador
p3 = (750, 63)   # Click en URL
pobs = (255, 5)  # Posición alternativa si hay un problema

pt.click(p1)
time.sleep(2)

for i in range(len(buscar)):
    if not isnan(buscar[i]):
        pt.typewrite(buscar[i])
        pt.hotkey("enter")
        time.sleep(5)

        pt.hotkey("ctrl", "f")
        time.sleep(5)
        pt.typewrite("maps")
        time.sleep(5)
        pt.hotkey("ctrl", "enter")
        time.sleep(10)

        pt.click(p3)
        pt.hotkey("ctrl", "c")

        a = re.findall(r"@-?\d+\.\d+,-?\d+\.\d+", pyperclip.paste())  # Regex para coordenadas
        try:
            if a:  # Asegúrate de que hay coincidencias
                a = re.sub("@", "", a[0])
                coord.append(a)
            else:
                print(f"No se encontraron coordenadas para: {buscar[i]}")
                pt.click(pobs)
                time.sleep(9)
                pt.click(p3)
                time.sleep(3)
                pt.hotkey("ctrl", "c")

                a = re.findall(r"@-?\d+\.\d+,-?\d+\.\d+", pyperclip.paste())
                if a:
                    a = re.sub("@", "", a[0])
                    coord.append(a)
                else:
                    print(f"No se encontraron coordenadas en el segundo intento para: {buscar[i]}")
        except Exception as e:
            print(f"Error en la obtención de coordenadas: {str(e)}")
            coord.append("")  # Agrega un valor vacío en caso de error
    else:
        coord.append("")  # Corrige el error tipográfico en 'appen'

# Convierte la lista de coordenadas a DataFrame y guarda en Excel
coord_df = pd.DataFrame(coord, columns=["Coordenadas"])  # Asigna un nombre a la columna
coord_df.to_excel("Direcciones.xlsx", index=False)
