# Automação apenas para login e geração da OF de materiais básicos
import pandas as pd
import os
import glob
import time
import re
import pyperclip
from datetime import datetime
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from collections import defaultdict

# Caminhos
PASTA_FILA = r"C:\Users\Matheus\Desktop\Fila_Pedidos"
CAMINHO_MATERIAIS_BASICOS = r"C:\Users\Matheus\Desktop\AutomatizaReq\MateriaisBasicos.xlsx"
CAMINHO_EMPREENDIMENTOS = r"C:\Users\Matheus\Desktop\AutomatizaReq\Empreendimentos.xlsx"

def mover_para_concluidos(arquivo):
    pasta_destino = r"C:\Users\Matheus\Desktop\PedidosProntos"
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
    destino = os.path.join(pasta_destino, os.path.basename(arquivo))
    os.rename(arquivo, destino)

def listar_arquivos_pendentes():
    return glob.glob(os.path.join(PASTA_FILA, "*.xlsx"))

def formatar_brl(valor):
    s = str(valor).strip()
    if s == "" or s.lower() == "nan":
        return "0,00"
    # Normaliza: se tiver . e , assume . milhar e , decimal; senão troca , por .
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        num = float(s)
    except:
        num = 0.0
    # Formata “pt-BR”
    return f"{num:.2f}".replace(".", ",")

def ler_dados_do_excel(arquivo):
    df = pd.read_excel(arquivo, sheet_name="Pedido", header=None)
    # C7 pode vir como "2504 - ..." ou só "2504"; pega exatamente 4 dígitos
    raw_c7 = str(df.iloc[6, 2]).strip()
    m = re.search(r'(\d{4})', raw_c7)
    if not m:
        raise ValueError(f"Não encontrei um código de 4 dígitos em C7: {raw_c7}")
    valor_c7_curtado = m.group(1)
    coluna_b = df.iloc[12:, 1].dropna()
    ultima_linha = coluna_b.last_valid_index()
    intervalo = df.loc[12:ultima_linha, [1, 2, 4, 5]]
    intervalo.columns = ['B', 'C', 'E', 'F']
    return valor_c7_curtado, intervalo

def montar_lista_basicos(intervalo, df_basicos, valor_c7):
    lista = []
    for _, row in intervalo.iterrows():
        codigo = str(row['B']).strip()
        qtd_pedida = float(str(row['E']).replace(',', '.'))
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
                        'fornecedor': linha.iloc[8],       # Coluna I
                        'valor_unitario': linha.iloc[7]     # Coluna H
                    })
            except Exception as e:
                print(f"⚠️ Erro ao validar faixa para código {codigo}: {e}")
    return lista

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
                valor_formatado = formatar_brl(valor_unitario)
            else:
                print(f"⚠️ Valor unitário ausente para item {item['codigo']}")
                valor_formatado = "0,00"
            pyautogui.click(788, 404); time.sleep(1.0)
            pyautogui.hotkey('ctrl', 'a'); time.sleep(0.1)
            pyautogui.press('backspace'); time.sleep(0.1)
            pyperclip.copy(valor_formatado); time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v'); time.sleep(0.2)
            pyautogui.press('enter'); time.sleep(3)
            pyautogui.click(620, 353); time.sleep(3)

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
        print("⚠️ Nenhum arquivo na Fila_Pedidos.")
        return

    # Login no Siecon
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
    pyautogui.click(92, 350); time.sleep(3)
    pyautogui.click(92, 750); time.sleep(10)

    df_basicos = pd.read_excel(CAMINHO_MATERIAIS_BASICOS, header=None)
    df_empreend = pd.read_excel(CAMINHO_EMPREENDIMENTOS, header=None, usecols=[0, 4])
    df_empreend.columns = ["Obra", "Estado"]
    # Extrai exatamente 4 dígitos do início do identificador da obra
    df_empreend["Cod4"] = (
        df_empreend["Obra"].astype(str).str.extract(r'(\d{4})', expand=False)
    )

    for arquivo in arquivos:
        print(f"\n🚀 Processando: {arquivo}")
        valor_c7, intervalo = ler_dados_do_excel(arquivo)  # ex.: "2504"

        # Busca o estado da obra pelo código de 4 dígitos
        estado_obra = df_empreend.loc[
            df_empreend["Cod4"].astype(str).str.strip() == str(valor_c7).strip(),
            "Estado"
        ]

        if estado_obra.empty:
            print(f"⚠️ Obra {valor_c7} não encontrada na planilha de empreendimentos. Pulando arquivo.")
            mover_para_concluidos(arquivo)
            continue

        if str(estado_obra.values[0]).strip().upper() != "RJ":
            print(f"🏳️ Obra {valor_c7} é do estado {estado_obra.values[0]}. OF não será gerada (apenas obras do RJ são processadas).")
            mover_para_concluidos(arquivo)
            continue

        # Geração da OF para obras do RJ
        lista_basicos = montar_lista_basicos(intervalo, df_basicos, valor_c7)
        if lista_basicos:
            print("📋 Materiais básicos encontrados. Gerando OF...")
            gerar_ordem_fornecimento(lista_basicos, valor_c7)
        else:
            print("⚠️ Nenhum material básico encontrado para este pedido.")

        mover_para_concluidos(arquivo)

    print("\n🎉 Finalizado!")
    input("\n🛑 Pressione Enter para sair...")

if __name__ == "__main__":
    main()