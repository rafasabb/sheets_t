import time
import re
import math
import gspread
from oauth2client.service_account import ServiceAccountCredentials

SCOPE = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
NAME = "Engenharia de Software – Desafio [Rafael Sabbagh]"

def get_total_registers(s):
  return len(s.col_values(1)[3:])
def get_total_classes(s):
  return re.search('[0-9]+', s.cell(2,1).value).group(0)

def update_spreadsheet(sheet):
  total_classes = get_total_classes(sheet)
  total_registers = get_total_registers(sheet)
  for i in range(total_registers):
    time.sleep(10)
    absences = sheet.cell(i+4,3).value

    if (int(absences) / int(total_classes) > 0.25):
      sheet.update_cell(i+4,7, "Reprovado por Falta")
      sheet.update_cell(i+4,8, "0")
      continue

    p1 = int(sheet.cell(i+4,4).value)
    p2 = int(sheet.cell(i+4,5).value)
    p3 = int(sheet.cell(i+4,6).value)
    m = (p1 + p2 + p3)/3
        
    if (m < 50):
      sheet.update_cell(i+4,7, "Reprovado por Nota")
      sheet.update_cell(i+4,8, "0")
      continue

    if (m >= 70):
      sheet.update_cell(i+4,7, "Aprovado")
      sheet.update_cell(i+4,8, "0")
      continue

    sheet.update_cell(i+4,7, "Exame Final")
    sheet.update_cell(i+4,8, math.ceil(((50*2) - m)))
    
def main():
  creds = ServiceAccountCredentials.from_json_keyfile_name('gupy.json', SCOPE)
  client = gspread.authorize(creds)
  sheet = client.open(NAME).sheet1
  update_spreadsheet(sheet)

if __name__ == '__main__':
  main()