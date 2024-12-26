from bs4 import BeautifulSoup
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

# Caminho para salvar os CSVs
output_csv_path = r'C:\FOLHA\data\output\planilhas\relatorio_ferias.csv'
output_html_path_13 = r'C:\FOLHA\data\output\planilhas\relatorio_13'

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

# Mover o clique
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
if saldo_final_valores:
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

# Repetir o processo para relatorio_13.csv
# Pressionar "Alt + R", "O", "1"
pyautogui.hotkey('alt', 'r')
time.sleep(0.5)
pyautogui.press('o')
time.sleep(0.5)
pyautogui.press('1')

time.sleep(2)

# Pressionar "tab" e digitar o mês
pyautogui.press('tab')
pyautogui.typewrite(month_year)
pyautogui.press('tab')

# Mover o clique
pyautogui.click(x=100, y=165)
time.sleep(1)

pyautogui.click(x=100, y=228)
time.sleep(5)

# Pressionar "enter" para salvar o html 
pyautogui.press('enter')    
time.sleep(2)
pyautogui.write(output_html_path_13)
pyautogui.press('enter')

# Aguardar um momento para garantir que o arquivo seja salvo
time.sleep(5)

# Caminho do seu arquivo HTML
caminho_arquivo = output_html_path_13

# Função para formatar os números no formato desejado
def formatar_valor(valor):
    # Substituindo a vírgula por ponto (para garantir que a conversão para float funcione)
    valor = valor.replace('.', '').replace(',', '.')
    try:
        # Convertendo para float e depois formatando para a string no formato desejado
        return '{:,.2f}'.format(float(valor)).replace(',', 'X').replace('.', ',').replace('X', '.')
    except ValueError:
        return valor  # Retorna o valor original se não for um número válido

# Abrindo o arquivo HTML
with open(r"C:\FOLHA\data\output\planilhas\relatorio_13.htm", 'r', encoding='utf-8') as file:
    html_content = file.read()

# Usando BeautifulSoup para fazer o parsing do HTML
soup = BeautifulSoup(html_content, 'html.parser')

# Procurando pela tag <td> com o conteúdo específico
td_elemento = soup.find('td', string="(=) Saldo final da conta")

# Verificando se o conteúdo foi encontrado
if td_elemento:
    # Encontrando todos os <td> na mesma linha após o <td> encontrado
    td_siguentes = td_elemento.find_all_next('td')
    
    # Definindo os rótulos personalizados para os 6 primeiros valores
    rotulos = [
        "salario_13", 
        "fgts_13", 
        "inss_13", 
        "pis_13", 
        "Total_13"
    ]
    
    # Exibindo os 6 primeiros valores com rótulos personalizados
    valores_13 = {}
    for i, td in enumerate(td_siguentes[1:6]):
        valores_13[rotulos[i]] = td.get_text()
        print(f'{rotulos[i]}: {td.get_text()}')
else:
    print("Elemento específico não encontrado.")

# Transformar o valor de salario_13 no estilo 88340.34
if 'salario_13' in valores_13:
    salario_13_formatado = valores_13['salario_13'].replace('.', '').replace(',', '.')
    valores_13['salario_13'] = salario_13_formatado
    print(f"salario_13 formatado: {salario_13_formatado}")

# Transformar o valor de inss_13 no estilo 88340.34
if 'inss_13' in valores_13:
    inss_13_formatado = valores_13['inss_13'].replace('.', '').replace(',', '.')
    valores_13['inss_13'] = inss_13_formatado
    print(f"inss_13 formatado: {inss_13_formatado}")

# Transformar o valor de fgts_13 no estilo 88340.34
if 'fgts_13' in valores_13:
    fgts_13_formatado = valores_13['fgts_13'].replace('.', '').replace(',', '.')
    valores_13['fgts_13'] = fgts_13_formatado
    print(f"fgts_13 formatado: {fgts_13_formatado}")

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

        # Procurar os valores e retornar os valores da coluna H
        vlr_prov_ferias = valores_mapeados.get(177, "Valor não encontrado")
        vlr_prov_fgts_f = valores_mapeados.get(179, "Valor não encontrado")
        vlr_prov_inss_f = valores_mapeados.get(181, "Valor não encontrado")
        vlr_prov_13 = valores_mapeados.get(176, "Valor não encontrado")  
        vlr_prov_inss_13 = valores_mapeados.get(180, "Valor não encontrado")
        vlr_prov_fgts_13 = valores_mapeados.get(178, "Valor não encontrado")

        print(f"vlr_prov_ferias: {vlr_prov_ferias}")
        print(f"vlr_prov_fgts_f: {vlr_prov_fgts_f}")
        print(f"vlr_prov_inss_f: {vlr_prov_inss_f}")
        print(f"vlr_prov_13: {vlr_prov_13}")
        print(f"vlr_prov_inss_13: {vlr_prov_inss_13}")
        print(f"vlr_prov_fgts_13: {vlr_prov_fgts_13}")

        # Imprimir os valores que estão sendo usados para a comparação
        print(f"Valor de salario_13 usado para comparação: {salario_13_formatado}")
        print(f"Valor de inss_13 usado para comparação: {inss_13_formatado}")
        print(f"Valor de fgts_13 usado para comparação: {fgts_13_formatado}")

        # Comparar os valores
        try:
            salario_13_float = float(salario_13_formatado)
            vlr_prov_13_float = float(vlr_prov_13)
            resultado_13 = "OK" if salario_13_float == vlr_prov_13_float else "verificar"
        except ValueError as e:
            print(f"Erro ao converter valores para float: {e}")
            resultado_13 = "verificar"

        try:
            inss_13_float = float(inss_13_formatado)
            vlr_prov_inss_13_float = float(vlr_prov_inss_13)
            resultado_inss_13 = "OK" if inss_13_float == vlr_prov_inss_13_float else "verificar"
        except ValueError as e:
            print(f"Erro ao converter valores para float: {e}")
            resultado_inss_13 = "verificar"

        try:
            fgts_13_float = float(fgts_13_formatado)
            vlr_prov_fgts_13_float = float(vlr_prov_fgts_13)
            resultado_fgts_13 = "OK" if fgts_13_float == vlr_prov_fgts_13_float else "verificar"
        except ValueError as e:
            print(f"Erro ao converter valores para float: {e}")
            resultado_fgts_13 = "verificar"

        resultado_ferias = "OK" if saldo_final_valores['FERIAS'] == vlr_prov_ferias else "verificar"
        resultado_fgts_f = "OK" if saldo_final_valores['FGTS_f'] == vlr_prov_fgts_f else "verificar"
        resultado_inss_f = "OK" if saldo_final_valores['INSS_f'] == vlr_prov_inss_f else "verificar"

        print(f"Comparação FERIAS: {resultado_ferias}")
        print(f"Comparação FGTS_f: {resultado_fgts_f}")
        print(f"Comparação INSS_f: {resultado_inss_f}")
        print(f"Comparação salario_13: {resultado_13}")
        print(f"Comparação inss_13: {resultado_inss_13}")
        print(f"Comparação fgts_13: {resultado_fgts_13}")

        # Escrever os resultados na coluna I ao lado dos valores
        for row in ws1.iter_rows(min_row=1, max_row=ws1.max_row, min_col=1, max_col=1):
            for cell in row:
                if cell.value == 177:
                    ws1.cell(row=cell.row, column=9, value=resultado_ferias)
                elif cell.value == 179:
                    ws1.cell(row=cell.row, column=9, value=resultado_fgts_f)
                elif cell.value == 181:
                    ws1.cell(row=cell.row, column=9, value=resultado_inss_f)
                elif cell.value == 176:
                    ws1.cell(row=cell.row, column=9, value=resultado_13)
                elif cell.value == 180:
                    ws1.cell(row=cell.row, column=9, value=resultado_inss_13)
                elif cell.value == 178:
                    ws1.cell(row=cell.row, column=9, value=resultado_fgts_13)

        # Salvar a planilha novamente
        wb.save(conciliacao_path)

        print(f"Resultados escritos na planilha de conciliação.")
    except Exception as e:
        print(f"Erro ao abrir ou processar o arquivo de conciliação: {e}")

# Pressionar "Esc" três vezes
pyautogui.press('esc')
pyautogui.press('esc')
pyautogui.press('esc')