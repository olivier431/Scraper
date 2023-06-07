import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
# Faites une requête GET vers l'URL contenant le tableau HTML
url = "https://www.quanthockey.com/nhl/seasons/nhl-players-stats.html"
response = requests.get(url)

# Analyse le contenu HTML avec BeautifulSoup
soup = BeautifulSoup(response.content, "html.parser")

table = soup.find("table", {"id": "statistics"})

# Crée un nouveau classeur Excel
wb = Workbook()
# Sélectionne la première feuille de calcul
ws = wb.active

rows = table.find_all("tr")

# Récupère l'en-tête du tableau
header_row = rows[1]
header_cells = header_row.find_all(["th", "td"])
header_texts = [cell.get_text(strip=True) for cell in header_cells]
ws.append(header_texts)

# Parcours les autres lignes du tableau
for row in rows[2:]:
    # Parcours les cellules de chaque ligne
    cell_texts = [cell.get_text(strip=True) for cell in row.find_all(["th", "td"])]
    ws.append(cell_texts)

# Sauvegarde le classeur Excel
wb.save("tableau.xlsx")

