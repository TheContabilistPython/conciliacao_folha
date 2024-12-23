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
time.sleep(1)
# Pressionar "tab"
pyautogui.press('tab')

# Digitar o mês e o ano no formato MMYYYY
pyautogui.typewrite(month_year)
pyautogui.press('enter')

pyautogui.click(x=100, y=165)
time.sleep(1)

pyautogui.click(x=100, y=256)
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

# Converter os valores de saldo_final_valores para float
for key in saldo_final_valores:
    saldo_final_valores[key] = float(saldo_final_valores[key].replace('.', '').replace(',', '.'))

# Converter o valor de INSS_f para número
if 'INSS_f' in saldo_final_valores:
    inss_f_valor = saldo_final_valores['INSS_f']
    print(f"Valor de INSS_f: {inss_f_valor}")

# Salvar os valores encontrados em um novo CSV
output_csv_path_final = r'C:\FOLHA\data\output\planilhas\saldo_final_valores.csv'
pd.DataFrame([saldo_final_valores]).to_csv(output_csv_path_final, index=False, sep=';')

print(f"Arquivo CSV com valores encontrados salvo em: {output_csv_path_final}")

# Imprimir o caminho do ambiente virtual
print("Caminho do ambiente virtual: c:/FOLHA/.venv/Scripts/python.exe")

# Caminho da planilha de conciliação
conciliacao_path = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{company_code}_{month_year}.xlsx'

# Verificar se o arquivo de conciliação existe e é válido
if not os.path.exists(conciliacao_path):
    print(f"Erro: O arquivo {conciliacao_path} não existe.")
else:
    try:
        # Abrir a planilha de conciliação
        wb = openpyxl.load_workbook(conciliacao_path)

        # Atualizar a Sheet1
        ws1 = wb.active

        # Mapeamento dos valores da coluna H com os valores da coluna A
        valores_mapeados = {}
        for row in ws1.iter_rows(min_row=1, max_row=ws1.max_row, min_col=1, max_col=8):
            cell_a = row[0]  # Coluna A
            cell_h = row[7]  # Coluna H
            if cell_a.value is not None:
                valores_mapeados[cell_a.value] = cell_h.value

        # Procurar os valores onde A = 177, 179 e 181 e retornar os valores da coluna H
        vlr_prov_ferias = valores_mapeados.get(177, "Valor não encontrado")
        vlr_prov_fgts_f = valores_mapeados.get(179, "Valor não encontrado")
        vlr_prov_inss_f = valores_mapeados.get(181, "Valor não encontrado")

        print(f"vlr_prov_ferias: {vlr_prov_ferias}")
        print(f"vlr_prov_fgts_f: {vlr_prov_fgts_f}")
        print(f"vlr_prov_inss_f: {vlr_prov_inss_f}")

        # Comparar os valores
        resultado_ferias = "OK" if saldo_final_valores['FERIAS'] == vlr_prov_ferias else "verificar"
        resultado_fgts_f = "OK" if saldo_final_valores['FGTS_f'] == vlr_prov_fgts_f else "verificar"
        resultado_inss_f = "OK" if saldo_final_valores['INSS_f'] == vlr_prov_inss_f else "verificar"

        print(f"Comparação FERIAS: {resultado_ferias}")
        print(f"Comparação FGTS_f: {resultado_fgts_f}")
        print(f"Comparação INSS_f: {resultado_inss_f}")

        # Escrever os resultados na coluna I ao lado dos valores
        for row in ws1.iter_rows(min_row=1, max_row=ws1.max_row, min_col=1, max_col=1):
            for cell in row:
                if cell.value == 177:
                    ws1.cell(row=cell.row, column=9, value=resultado_ferias)
                elif cell.value == 179:
                    ws1.cell(row=cell.row, column=9, value=resultado_fgts_f)
                elif cell.value == 181:
                    ws1.cell(row=cell.row, column=9, value=resultado_inss_f)

        # Salvar a planilha novamente
        wb.save(conciliacao_path)

        print(f"Resultados escritos na planilha de conciliação.")
    except Exception as e:
        print(f"Erro ao abrir ou processar o arquivo de conciliação: {e}")