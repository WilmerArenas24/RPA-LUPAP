from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
from openpyxl import load_workbook

# Ruta del archivo de Excel
ruta_excel = r"C:\Users\wilme\PycharmProjects\RPA_Direcciones\Direcciones.xlsx"

# Leer el archivo Excel con pandas para obtener los datos
data = pd.read_excel(ruta_excel)

# Asegurarse de que la columna "Direccion_normalizada" sea del tipo str
data["Direccion_normalizada"] = data.get("Direccion_normalizada", pd.Series(dtype='str'))

# Buscar las direcciones y ciudades del Excel
buscar = data["Direccion"].values
ciudades = data["Ciudad"].values  # Obtener las ciudades

# Especificar la ruta al ChromeDriver usando la clase Service
service = Service(executable_path=r"C:\Users\wilme\PycharmProjects\RPA_Direcciones\chromedriver-win64\chromedriver.exe")

# Inicializar el navegador (en este caso Chrome)
driver = webdriver.Chrome(service=service)

# Rutas de las ciudades
city_urls = {
    "FUNZA": "https://lupap.com/address/funza/CL+10A+12A+4",
    "BOGOTA": "https://lupap.com/address/bogota/KR+13H+29+9+Sur",
    "MOSQUERA": "https://lupap.com/address/cun_mosquera/CL+1+3+04"
    # Agrega más ciudades aquí según sea necesario
}

# Buscar las direcciones del Excel
for i in range(len(buscar)):
    ciudad = ciudades[i].upper()  # Convertir a mayúsculas para coincidir con las claves
    direccion = buscar[i]

    # Verificar si la ciudad está en el diccionario
    if ciudad in city_urls:
        # Abrir la página web de la ciudad correspondiente
        driver.get(city_urls[ciudad])
        time.sleep(5)  # Esperar a que la página cargue

        # Encontrar el campo de búsqueda usando su XPath
        search_box = driver.find_element("xpath", "/html/body/div[4]/form/div/input")

        # Limpiar el campo de búsqueda antes de escribir
        search_box.clear()

        # Escribir en el campo de búsqueda
        search_box.send_keys(direccion)

        # Simular el botón Enter para realizar la búsqueda
        search_box.send_keys(Keys.RETURN)

        # Esperar unos segundos para que los resultados aparezcan
        time.sleep(5)

        # Tomar el valor del elemento ubicado en /html/body/div[5]/div[1]/div/span
        try:
            element_value = driver.find_element("xpath", "/html/body/div[5]/div[1]/div/span").text
            # Guardar el valor en la columna "Direccion_normalizada"
            data.at[i, "Direccion_normalizada"] = str(element_value)  # Convertir a str
            # Imprimir el valor en la consola
            print("El valor encontrado es:", element_value)
        except Exception as e:
            print("No se pudo encontrar el valor:", e)
            # Guardar un mensaje de error en la columna "Direccion_normalizada"
            data.at[i, "Direccion_normalizada"] = "No encontrado"
    else:
        print(f"La ciudad '{ciudad}' no está en la lista de URL de ciudades.")

# Cerrar el navegador
driver.quit()

# Guardar los resultados en la columna "Direccion_normalizada" sin perder el formato de tabla
book = load_workbook(ruta_excel)
sheet = book.active

for i in range(len(buscar)):
    # Escribir los resultados en la celda correspondiente
    sheet[f"C{i + 2}"] = data.at[i, "Direccion_normalizada"]  # Cambia "C" por la letra de la columna deseada

# Guardar el archivo
book.save(ruta_excel)
book.close()
print("Los resultados se han guardado en el archivo Excel.")

