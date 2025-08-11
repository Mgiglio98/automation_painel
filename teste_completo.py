import pandas as pd
import os
import shutil
import time
import glob
import pyperclip
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
from collections import defaultdict

PASTA_FILA = r"C:\Users\Matheus\Desktop\Fila_Pedidos"
PASTA_PRONTOS = r"C:\Users\Matheus\Desktop\PedidosProntos"
CAMINHO_MATERIAIS_BASICOS = r"C:\Users\Matheus\Desktop\AutomatizaReq\MateriaisBasicos.xlsx"

def listar_arquivos_pendentes():
    return glob.glob(os.path.join(PASTA_FILA, "*.xlsx"))

def ler_dados_do_excel(arquivo):
    df = pd.read_excel(arquivo, sheet_name="Pedido", header=None)
    valor_c7 = str(df.iloc[6, 2])
    valor_c7_curtado = valor_c7[:4]
    coluna_b = df.iloc[12:, 1].dropna()
    ultima_linha = coluna_b.last_valid_index()
    intervalo = df.loc[12:ultima_linha, [1, 2, 4, 5]]
    intervalo.columns = ['B', 'C', 'E', 'F']
    return valor_c7_curtado, intervalo

def mover_para_concluidos(arquivo):
    if not os.path.exists(PASTA_PRONTOS):
        os.makedirs(PASTA_PRONTOS)
    shutil.move(arquivo, os.path.join(PASTA_PRONTOS, os.path.basename(arquivo)))

def identificar_basicos(intervalo, codigos_basicos):
    return [codigo for codigo in intervalo['B'].astype(str) if codigo in codigos_basicos]

def montar_lista_basicos(intervalo, df_basicos, valor_c7):
    lista = []

    for _, row in intervalo.iterrows():
        codigo = str(row['B']).strip()
        qtd_pedida = float(row['E'])
        linha_basico = df_basicos[df_basicos.iloc[:, 1].astype(str) == codigo]
        if not linha_basico.empty:
            linha = linha_basico.iloc[0]
            try:
                minimo = float(str(linha.iloc[4]).replace(',', '.'))
                maximo = float(str(linha.iloc[5]).replace(',', '.'))
                if minimo <= qtd_pedida <= maximo:
                    lista.append({
                        'codigo': codigo,
                        'descricao': row['C'],
                        'quantidade': qtd_pedida,
                        'empreendimento': valor_c7,
                        'fornecedor': linha.iloc[8],
                        'valor_unitario': str(linha.iloc[7]).replace('.', ',') if pd.notna(linha.iloc[7]) else "0,00"
                    })
            except Exception as e:
                print(f"⚠️ Erro ao validar faixa para código {codigo}: {e}")
    return lista

def executar_automacao(valor_c7, intervalo, driver):
    pyautogui.getWindowsWithTitle("Google Chrome")[0].activate()
    pyautogui.click(500, 230); time.sleep(0.5); pyautogui.write(valor_c7); pyautogui.press('enter'); time.sleep(2)
    pyautogui.click(1050, 230); time.sleep(2)
    texto_busca = ", ".join([i.split()[0] for i in intervalo['C'].dropna().astype(str).head(3)])
    pyautogui.click(315, 320); time.sleep(0.5); pyautogui.write(texto_busca); time.sleep(2)
    pyautogui.click(1135, 230); time.sleep(2)
    pyautogui.click(327, 452, clicks=1); time.sleep(0.5); pyautogui.hotkey('ctrl', 'c'); time.sleep(1)
    linhas = pyperclip.paste().strip().splitlines()
    numero_requisicao = linhas[2].split("\t")[0] if len(linhas) >= 2 else "NÃO ENCONTRADO"
    print(f"📋 Requisição criada: {numero_requisicao}")
    pyautogui.click(500, 395); time.sleep(2)

    for _, row in intervalo.iterrows():
        pyautogui.click(310, 432); time.sleep(1)
        pyautogui.click(422, 538); time.sleep(0.5); pyautogui.write("100.01.02"); pyautogui.press('enter'); time.sleep(1.5)
        pyautogui.click(363, 587); time.sleep(0.5); pyautogui.write(str(row['B'])); pyautogui.press('enter'); time.sleep(1.5)
        pyautogui.click(789, 587); time.sleep(0.5); pyautogui.write(str(row['E'])); time.sleep(1.5)
        pyautogui.click(908, 587); time.sleep(0.5); pyautogui.write("5"); time.sleep(1.5)
        if pd.notna(row['F']) and str(row['F']).strip().lower() != "nan":
            pyautogui.click(434, 637); pyautogui.hotkey('ctrl', 'shift', 'home'); time.sleep(0.5); pyautogui.press('backspace'); time.sleep(0.5)
            pyautogui.write(str(row['F'])); time.sleep(1.5)
        pyautogui.click(405, 434); time.sleep(2)

    print("🔒 Aprovando requisição...")
    pyautogui.click(1103, 234); time.sleep(2)
    pyautogui.click(1333, 287); time.sleep(2)
    pyautogui.click(1136, 234); time.sleep(4)
    print("✅ Requisição aprovada.")

from collections import defaultdict

def gerar_ordem_fornecimento(lista_basicos, valor_c7):
    if not lista_basicos:
        return

    print("🛠️ Iniciando geração da OF...")

    pyautogui.click(92, 750); time.sleep(10)

    # Agrupa os insumos por fornecedor
    insumos_por_fornecedor = defaultdict(list)
    for item in lista_basicos:
        fornecedor = item['fornecedor']
        insumos_por_fornecedor[fornecedor].append(item)

    for fornecedor, itens in insumos_por_fornecedor.items():
        print(f"📦 Gerando OF para fornecedor: {fornecedor}")

        # Cabeçalho da OF
        pyautogui.click(319, 318); time.sleep(3)
        pyautogui.click(1056, 226); time.sleep(3)
        pyautogui.click(477, 239); time.sleep(0.5); pyautogui.write("MATERIAL BASICO"); time.sleep(2)
        pyautogui.click(321, 361); time.sleep(0.5); pyautogui.write(valor_c7); pyautogui.press("enter"); time.sleep(2)
        pyautogui.click(629, 361); time.sleep(0.5); pyautogui.write(fornecedor); pyautogui.press("enter"); time.sleep(2)
        pyautogui.click(514, 454); time.sleep(0.5); pyautogui.write("3"); pyautogui.press("enter"); time.sleep(2)

        data_hoje = datetime.today().strftime("%d%m%Y")
        pyautogui.click(1274, 410); time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace'); time.sleep(0.3)
        pyautogui.write(data_hoje, interval=0.2); time.sleep(1.5)

        pyautogui.click(1415, 410); time.sleep(0.5); pyautogui.write("MATHEUS")
        pyautogui.click(1118, 565); time.sleep(0.5); pyautogui.write("21")
        pyautogui.click(1163, 565); time.sleep(0.5); pyautogui.write("25217235")
        pyautogui.click(1213, 225); time.sleep(5)

        # Adiciona cada insumo
        for item in itens:
            pyautogui.click(940, 322); time.sleep(2)
            pyautogui.click(307, 358); time.sleep(2)
            pyautogui.click(677, 358); time.sleep(1); pyautogui.press('home'); time.sleep(1); pyautogui.press('down', presses=7); time.sleep(1); pyautogui.press('enter'); time.sleep(2)
            pyautogui.click(718, 355); time.sleep(0.5); pyautogui.write(item['codigo']); time.sleep(0.5)
            pyautogui.click(1398, 342); time.sleep(2)
            pyautogui.click(496, 522); time.sleep(2)
            pyautogui.click(1398, 380); time.sleep(5)
            pyautogui.click(399, 402); time.sleep(1)
            pyautogui.click(506, 361); time.sleep(1)
            pyautogui.click(824, 566); time.sleep(0.5); pyautogui.write("2"); time.sleep(2); pyautogui.click(910, 615); time.sleep(2) 
            valor_unitario = item.get('valor_unitario')
            if pd.notna(valor_unitario):
                valor_formatado = str(valor_unitario).strip()
            else:
                print(f"⚠️ Valor unitário ausente para item {item['codigo']}")
                valor_formatado = "0,00"
            pyautogui.click(788, 404); time.sleep(1.5)
            pyautogui.write(valor_formatado); time.sleep(0.5)
            pyautogui.press('enter'); time.sleep(3)
            pyautogui.click(600, 351); time.sleep(3)

        # Finalização da OF
        pyautogui.click(1182, 223); time.sleep(3)
        pyautogui.click(672, 280); time.sleep(3)
        pyautogui.click(1213, 225); time.sleep(3) 
        pyautogui.click(1230, 323); time.sleep(5)
        pyautogui.click(879, 617); time.sleep(3)

    print("✅ Todas as Ordens de Fornecimento foram geradas.")

def main():
    arquivos = listar_arquivos_pendentes()
    if not arquivos:
        print("⚠️ Nenhum arquivo na Fila_Pedidos."); return

    chrome_options = Options()
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    driver.get("https://cloud4.siecon.com.br/"); time.sleep(3)
    driver.maximize_window(); time.sleep(2)
    driver.find_element(By.ID, "Editbox1").send_keys("oc.matheus.almeida")
    driver.find_element(By.ID, "Editbox2").send_keys("#Osborne")
    driver.find_element(By.ID, "buttonLogOn").click(); time.sleep(10)

    pyautogui.getWindowsWithTitle("Google Chrome")[0].activate()
    pyautogui.moveTo(92, 589); pyautogui.click(); time.sleep(13)
    pyautogui.click(874, 562); time.sleep(0.5); pyautogui.write("MATHEUS.ALMEIDA"); time.sleep(5)
    pyautogui.click(874, 610); time.sleep(0.5); pyautogui.write("2025"); time.sleep(5)
    pyautogui.click(1089, 560); time.sleep(7)
    pyautogui.click(92, 350); time.sleep(4)
    pyautogui.click(92, 675); time.sleep(4)

    df_basicos = pd.read_excel(CAMINHO_MATERIAIS_BASICOS, header=0, dtype={7: str})
    for arquivo in arquivos:
        print(f"\n🚀 Processando: {arquivo}")
        valor_c7, intervalo = ler_dados_do_excel(arquivo)
        executar_automacao(valor_c7, intervalo, driver)
        lista_basicos = montar_lista_basicos(intervalo, df_basicos, valor_c7)
        if lista_basicos:
            print("📋 Materiais básicos encontrados. Gerando OF...")
            gerar_ordem_fornecimento(lista_basicos, valor_c7)
        else:
            print("⚠️ Nenhum material básico encontrado para este pedido.")
        mover_para_concluidos(arquivo)

    print("\n🎉 Todos os pedidos foram processados!")
    input("\n🛑 Pressione Enter para finalizar...")

if __name__ == "__main__":
    main()