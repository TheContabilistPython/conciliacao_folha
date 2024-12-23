import pyautogui
import time
import tkinter as tk
from tkinter import simpledialog
import os
import pandas as pd
import openpyxl

# Função para obter o código da empresa, o mês e o ano do usuário
def get_user_input():
    root = tk.Tk()
    root.withdraw()  # Esconder a janela principal

    # Solicitar o código da empresa
    company_code = simpledialog.askstring(title="Código da Empresa", prompt="Digite o código da empresa:")

    # Solicitar o mês e o ano
    month_year = simpledialog.askstring(title="Mês e Ano", prompt="Digite o mês e o ano (MMYYYY):")

    return company_code, month_year

# Obter o código da empresa e o mês e o ano do usuário
company_code, month_year = get_user_input()

# Criar a variável day_month_year
day_month_year = '01' + month_year

# Caminho para o executável na rede
network_path = r'\\ap05\modulos\UNICO.EXE'

# Caminho para salvar o CSV
output_csv_path = r'C:\FOLHA\data\output\planilhas\relatorio_ferias.csv'

# Pressionar Win + R
pyautogui.hotkey('win', 'r')
time.sleep(1)  # Esperar um momento para a janela de execução abrir

# Digitar o caminho do executável na rede
pyautogui.typewrite(network_path)
pyautogui.press('enter')

# Esperar 12 segundos para o executável carregar
time.sleep(12)

# Digitar "contabil"
pyautogui.typewrite('contabil')

# Pressionar "tab"
pyautogui.press('tab')

# Digitar "1234"
pyautogui.typewrite('1234')

# Pressionar "enter"
pyautogui.press('enter')

# Esperar um momento para a próxima ação
time.sleep(5)

# Pressionar "Ctrl + 0"
pyautogui.hotkey('ctrl', '0')

# Esperar 10 segundos para a próxima ação
time.sleep(10)

# Pressionar "Alt + R", "O", "F"
pyautogui.hotkey('alt', 'r')
time.sleep(0.5)
pyautogui.press('o')
time.sleep(0.5)
pyautogui.press('f')

# Esperar um momento para a próxima ação
time.sleep(1)

# Digitar o código da empresa
pyautogui.typewrite(company_code)
pyautogui.press('enter')

# Pressionar "tab"
pyautogui.press('tab')

# Digitar o mês e o ano no formato MMYYYY
pyautogui.typewrite(month_year)
pyautogui.press('enter')

# Pressionar "Alt + F8"
pyautogui.hotkey('alt', 'f8')

# Esperar 3 segundos
time.sleep(3)

# Clicar em x=632, y=125
pyautogui.click(x=632, y=125)

# Esperar 0,5 segundos
time.sleep(0.5)

# Clicar em x=735, y=241
pyautogui.click(x=735, y=241)

time.sleep(3)

# Clicar em x=634, y=125
pyautogui.click(x=634, y=125)
time.sleep(1)

time.sleep(5)

# Clicar em x=636, y=219
pyautogui.click(x=636, y=219)
time.sleep(5)

# Pressionar "enter" para salvar o CSV
pyautogui.press('enter')
time.sleep(5)

# Digitar o caminho para salvar o CSV
pyautogui.typewrite(output_csv_path)
time.sleep(5)

pyautogui.press('enter')

# Aguardar um momento para garantir que o arquivo seja salvo
time.sleep(5)

# Verificar se o arquivo foi salvo corretamente
if os.path.exists(output_csv_path):
    print(f"Arquivo CSV salvo em: {output_csv_path}")
else:
    print("Erro ao salvar o arquivo CSV.")

# Ler o CSV e procurar por "(=) Saldo final da conta" na coluna A
df = pd.read_csv(output_csv_path, sep=';', encoding='latin1')

# Substituir os nomes das colunas conforme solicitado
df.rename(columns={
    'Unnamed: 5': 'FERIAS',
    'Unnamed: 11': 'FGTS_f',
    'Unnamed: 16': 'INSS_f',
    'Unnamed: 23': 'PIS_f',
    'Unnamed: 29': 'TOTAL_f'
}, inplace=True)

# Procurar por "(=) Saldo final da conta" na coluna A
saldo_final_row = df[df.iloc[:, 0].str.contains(r"\(=\) Saldo final da conta", na=False)]

# Variáveis para armazenar os dados
saldo_final_valores = None

# Se encontrado, armazenar os valores correspondentes à linha
if not saldo_final_row.empty:
    saldo_final_valores = saldo_final_row[['FERIAS', 'FGTS_f', 'INSS_f', 'PIS_f', 'TOTAL_f']].iloc[0].dropna().to_dict()
    print("Saldo final da conta encontrado. Valores correspondentes:")
    print(saldo_final_valores)
else:
    print("Saldo final da conta não encontrado.")

# Converter o valor de INSS_f para número
if 'INSS_f' in saldo_final_valores:
    inss_f_valor = float(saldo_final_valores['INSS_f'].replace('.', '').replace(',', '.'))
    print(f"Valor de INSS_f: {inss_f_valor}")

# Salvar os valores encontrados em um novo CSV
output_csv_path_final = r'C:\FOLHA\data\output\planilhas\saldo_final_valores.csv'
pd.DataFrame([saldo_final_valores]).to_csv(output_csv_path_final, index=False, sep=';')

print(f"Arquivo CSV com valores encontrados salvo em: {output_csv_path_final}")

# Imprimir o caminho do ambiente virtual
print("Caminho do ambiente virtual: c:/FOLHA/.venv/Scripts/python.exe")

time.sleep(3)

pyautogui.hotkey('ctrl', '2')

time.sleep(3)

pyautogui.press('alt')
time.sleep(2)
pyautogui.press('r')
time.sleep(2)
pyautogui.press('c')
time.sleep(2)

pyautogui.typewrite(day_month_year)
time.sleep(2)

# Pressionar "tab" duas vezes
pyautogui.press('tab')
time.sleep(2)
pyautogui.press('tab')
time.sleep(2)

# Digitar "181"
pyautogui.typewrite('181')
pyautogui.press('enter')
time.sleep(2)

# Pressionar "pg dn"
pyautogui.press('pagedown')
time.sleep(2)

pyautogui.click(x=63, y=121)

time.sleep(2)

pyautogui.press('tab')

time.sleep(2)

pyautogui.click(x=95, y=156)

time.sleep(1)

pyautogui.click(x=95, y=254)

time.sleep(2)

pyautogui.press('enter')

time.sleep(4)

pyautogui.write(r'C:\FOLHA\data\output\planilhas\cc_inss_prov_f')

time.sleep(1)

pyautogui.press('enter')

# Aguardar um momento para garantir que o arquivo seja salvo
time.sleep(5)

# Caminho do segundo relatório
second_report_path = r'C:\FOLHA\data\output\planilhas\cc_inss_prov_f.csv'

# Verificar se o segundo arquivo foi salvo corretamente
if os.path.exists(second_report_path):
    print(f"Segundo arquivo CSV salvo em: {second_report_path}")
else:
    print("Erro ao salvar o segundo arquivo CSV.")

# Ler o segundo relatório e trazer toda a coluna D
df_second = pd.read_csv(second_report_path, sep=';', encoding='latin1')

# Trazer toda a coluna D
coluna_d = df_second.iloc[:, 3]

# Separar a primeira e a última observação
primeiro_valor = coluna_d.iloc[0]
ultimo_valor = coluna_d.iloc[-1]

# Verificar se os valores são strings antes de converter
if isinstance(primeiro_valor, str):
    primeiro_valor = float(primeiro_valor.replace('.', '').replace(',', '.'))
if isinstance(ultimo_valor, str):
    ultimo_valor = float(ultimo_valor.replace('.', '').replace(',', '.'))

# Criar uma nova variável com o valor da última menos a primeira
cc_prov_inss_f = round(ultimo_valor - primeiro_valor, 2)

print(f"Primeiro valor da coluna D: {primeiro_valor}")
print(f"Último valor da coluna D: {ultimo_valor}")
print(f"Diferença entre o último e o primeiro valor (cc_prov_inss_f): {cc_prov_inss_f}")

# Verificar se inss_f_valor e cc_prov_inss_f são iguais
if inss_f_valor == cc_prov_inss_f:
    resultado = "OK"
else:
    resultado = "verificar"

# Caminho da planilha de conciliação
conciliacao_path = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{company_code}_{month_year}.xlsx'

# Abrir a planilha de conciliação
wb = openpyxl.load_workbook(conciliacao_path)
ws = wb.active

# Procurar por 181 na coluna A
for row in ws.iter_rows(min_row=1, max_col=1, max_row=ws.max_row):
    for cell in row:
        if cell.value == 181:
            # Escrever o resultado na coluna I da mesma linha
            ws.cell(row=cell.row, column=9, value=resultado)
            break

# Salvar a planilha
wb.save(conciliacao_path)

print(f"Resultado '{resultado}' escrito na planilha de conciliação.")