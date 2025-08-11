import pandas as pd
import os
import shutil
import time
import glob
import pyautogui
import pyperclip
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

PASTA_FILA = r"C:\Users\Matheus\Desktop\Fila_Pedidos"
PASTA_PRONTOS = r"C:\Users\Matheus\Desktop\PedidosProntos"

def listar_arquivos_pendentes():
    return glob.glob(os.path.join(PASTA_FILA, "*.xlsx"))

def ler_dados_do_excel(arquivo):
    df = pd.read_excel(arquivo, sheet_name="Pedido", header=None)
    valor_c7 = str(df.iloc[6, 2])[:4]
    coluna_b = df.iloc[12:, 1].dropna()
    ultima_linha = coluna_b.last_valid_index()
    intervalo = df.loc[12:ultima_linha, [1, 2, 4, 5]]
    intervalo.columns = ['B', 'C', 'E', 'F']
    return valor_c7, intervalo

def mover_para_concluidos(arquivo):
    if not os.path.exists(PASTA_PRONTOS):
        os.makedirs(PASTA_PRONTOS)
    shutil.move(arquivo, os.path.join(PASTA_PRONTOS, os.path.basename(arquivo)))

def executar_requisicao(valor_c7, intervalo):
    pyautogui.getWindowsWithTitle("Google Chrome")[0].activate()
    pyautogui.click(500, 230); time.sleep(0.5)
    pyautogui.write(valor_c7); pyautogui.press('enter'); time.sleep(2)
    pyautogui.click(1050, 230); time.sleep(2)

    texto_busca = ", ".join([i.split()[0] for i in intervalo['C'].dropna().astype(str).head(3)])
    pyautogui.click(315, 320); time.sleep(0.5)
    pyautogui.write(texto_busca); time.sleep(2)
    pyautogui.click(1135, 230); time.sleep(2)

    pyautogui.click(327, 452); time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'c'); time.sleep(1)
    linhas = pyperclip.paste().strip().splitlines()
    numero_requisicao = linhas[2].split("\t")[0] if len(linhas) >= 2 else "NÃO ENCONTRADO"
    print(f"📋 Requisição criada: {numero_requisicao}")

    pyautogui.click(500, 395); time.sleep(2)

    for _, row in intervalo.iterrows():
        # 🔒 Ignora insumos sem código
        if pd.isna(row['B']) or str(row['B']).strip() == "":
            print(f"⚠️ Insumo sem código ignorado: {row['C']}")
            continue
        pyautogui.click(310, 432); time.sleep(1)
        pyautogui.click(422, 538); time.sleep(0.5); pyautogui.write("100.01.02"); pyautogui.press('enter'); time.sleep(1.5)
        pyautogui.click(363, 587); time.sleep(0.5); pyautogui.write(str(row['B'])); pyautogui.press('enter'); time.sleep(1.5)
        pyautogui.click(789, 587); time.sleep(0.5); pyautogui.write(str(row['E'])); time.sleep(1.5)
        pyautogui.click(908, 587); time.sleep(0.5); pyautogui.write("5"); time.sleep(1.5)
        complemento = str(row['F']).strip().upper()
        if complemento and complemento.lower() != "nan":
            print(f'Digitando complemento: "{complemento}"')
            pyautogui.click(434, 637); time.sleep(0.5)
            for _ in range(14):
                pyautogui.press('backspace')
                time.sleep(0.05)
            pyautogui.click(434, 637); time.sleep(0.5)
            pyautogui.write(complemento); time.sleep(1)
        pyautogui.click(405, 434); time.sleep(2)
    
    print("✅ Requisição finalizada.")

def main():
    arquivos = listar_arquivos_pendentes()
    if not arquivos:
        print("⚠️ Nenhum arquivo na Fila_Pedidos.")
        return

    # Inicia navegador e login
    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    driver.get("https://cloud4.siecon.com.br/")
    time.sleep(3)
    driver.maximize_window()
    time.sleep(2)
    driver.find_element(By.ID, "Editbox1").send_keys("oc.matheus.almeida")
    driver.find_element(By.ID, "Editbox2").send_keys("#Osborne")
    driver.find_element(By.ID, "buttonLogOn").click()
    time.sleep(10)

    # Abre módulo de requisição
    pyautogui.getWindowsWithTitle("Google Chrome")[0].activate()
    pyautogui.moveTo(92, 589); pyautogui.click(); time.sleep(13)
    pyautogui.click(874, 562); time.sleep(0.5); pyautogui.write("MATHEUS.ALMEIDA"); time.sleep(5)
    pyautogui.click(874, 610); time.sleep(0.5); pyautogui.write("2025"); time.sleep(5)
    pyautogui.click(1089, 560); time.sleep(7)
    pyautogui.click(92, 350); time.sleep(4)
    pyautogui.click(92, 675); time.sleep(4)

    # Processa todos os pedidos
    for arquivo in arquivos:
        print(f"\n🚀 Processando pedido: {os.path.basename(arquivo)}")
        valor_c7, intervalo = ler_dados_do_excel(arquivo)
        executar_requisicao(valor_c7, intervalo)
        mover_para_concluidos(arquivo)

    print("\n🎉 Todos os pedidos foram lançados com sucesso.")
    input("\n🛑 Pressione Enter para finalizar...")

if __name__ == "__main__":
    main()