from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
from datetime import datetime
import sys

# Chemin vers le fichier exécutable ChromeDriver
chrome_driver_path = "chemin/vers/chromedriver"

# Instancie le navigateur Chrome
service = Service(chrome_driver_path)
options = Options()
options.add_argument("--headless")  # Exécuter Chrome en mode headless
options.add_argument("--disable-javascript") #désactive l'exécution de JavaScript dans le navigateur
options.add_argument("--log-level=3") #définit le niveau de journalisation sur 3
driver = webdriver.Chrome(service=service, options=options) #nitialise une nouvelle instance du navigateur Chrome

#Permet à l'utilisateur d'entrer une année dans la plage demandé
current_year = datetime.now().year
print('Pour quelle année voulez-vous le classement ?')
while True:
    year = input()
    try:
        year = int(year)
        if year < 1918 or year > current_year:
            raise ValueError()
        break
    except ValueError:
        print(f"Veuillez entrer une année entre 1918 et {current_year}.")
    except KeyboardInterrupt:
        sys.exit()


# Faites une requête GET vers l'URL contenant le tableau HTML
url = 'https://www.hockey-reference.com/leagues/NHL_' + str(year) + '.html'
driver.get(url)

# Attends que le contenu dynamique soit chargé (ajustez le délai si nécessaire)
driver.implicitly_wait(10)

# Trouve le tableau par son ID
table = driver.find_element("id", "stats")

# Trouve l'en-tête du tableau
header = table.find_element("tag name", "thead")
header_rows = header.find_elements("tag name", "tr")
header_row = header_rows[1]

# # Crée un nouveau classeur Excel
wb = Workbook()
# # Sélectionne la première feuille de calcul
ws = wb.active

# Parcours les cellules <th> de l'en-tête
header_th_cells = header_row.find_elements("tag name", "th")
header_th_texts = [cell.text for cell in header_th_cells]

ws.append(header_th_texts)

# # Parcours les lignes du tableau
body = table.find_element("tag name", "tbody")
rows = body.find_elements("tag name", "tr")

# # Parcours les lignes du tableau
for row in rows:
    # print(row.get_attribute("class"))
    if row.get_attribute("class") != "thead" and row.get_attribute("class") != "over_header thead" :
        # Parcours les cellules <th> et <td> de chaque ligne
       
        th_cells = row.find_elements("tag name", "th")
        th_texts = [cell.text for cell in th_cells]

        td_cells = row.find_elements("tag name", "td")
        td_texts = [cell.text for cell in td_cells]

        row_values = th_texts + td_texts

        ws.append(row_values)

# # Sauvegarde le classeur Excel

wb.save("tableau2.xlsx")

# Ferme le navigateur
driver.quit()
