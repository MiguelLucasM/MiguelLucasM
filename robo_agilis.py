# -*- coding: utf-8 -*-
"""
Robô RPA — Abertura de chamados no Agilis (Pagamento de Frete FOB / MRV)
=========================================================================
Lê a planilha de pedidos desaprovados e, para cada linha, abre um chamado
no Agilis, anexa o comunicado (PPTX), captura o número do chamado gerado
e grava esse número na coluna G da própria linha, salvando a planilha
imediatamente (permite retomar de onde parou em caso de queda).

Requisitos: Python 3.x, Selenium 4, openpyxl (e pyautogui apenas se for
usar o fallback de anexo via janela nativa do Windows).

Autor: Time de Logística (GPD) — MRV
"""

import logging
import os
import re
import sys
import time

from openpyxl import load_workbook
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    ElementNotInteractableException,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# =========================================================================
# CONFIG — caminhos, URL e parâmetros gerais
# =========================================================================

CAMINHO_EDGEDRIVER = r"C:\Webdriver\msedgedriver.exe"

CAMINHO_PLANILHA = (
    r"C:\Users\miguel.albuquerque\OneDrive - MRV"
    r"\Logística MRV & CO - Documentos\4 - GPD"
    r"\5 - Orçamento & Finanças\Gestão de Pagamentos"
    r"\Agilis - Pedidos Desaprovados.xlsx"
)

CAMINHO_ANEXO = (
    r"C:\Users\miguel.albuquerque\OneDrive - MRV"
    r"\Área de Trabalho\Comunicado Planilha Erros FOB.pptx"
)

URL_AGILIS = "https://interfaceagilis.mrv.com.br/formularios/central-servicos-suprimentos"

ARQUIVO_LOG = "robo_agilis.log"

# Timeout dos WebDriverWait (generoso por causa de rede instável)
TIMEOUT = 40
# Nº de tentativas dos helpers com retry
MAX_TENTATIVAS = 3
# Pausa entre tentativas de retry (segundos)
PAUSA_ENTRE_TENTATIVAS = 2.0
# Pausa do fluxo entre Tab e Enter (respiro do autocomplete Angular Material)
PAUSA_TAB_ENTER = 0.5
# Espera após clicar no ícone de anexo (fallback com janela nativa)
PAUSA_JANELA_ANEXO = 4.0
# Linha inicial dos dados na planilha (linha 1 = cabeçalho)
LINHA_INICIAL = 2

# Colunas da planilha (letras)
COL_PEDIDO_FRETE = "B"
COL_ORDEM_COLETA = "C"
COL_PEDIDO_MATERIAL = "D"
COL_ERRO = "E"
COL_CHAMADO = "G"  # saída — nº do chamado gravado pelo robô

# =========================================================================
# CONFIG — XPaths do fluxo
# -------------------------------------------------------------------------
# ATENÇÃO: os IDs "mat-input-N" / "mat-option-N" são gerados automaticamente
# pelo Angular Material e PODEM MUDAR de número entre versões da tela.
# Se algum passo quebrar, ajuste aqui (tudo centralizado) e, quando possível,
# troque pela alternativa por texto visível indicada nos comentários.
# =========================================================================

# Passo 1 — campo de seleção do serviço e opção desejada
XPATH_CAMPO_SERVICO = '//*[@id="mat-input-0"]'
XPATH_OPCAO_SERVICO = '//*[@id="mat-option-2"]'
# Alternativa por texto visível (troque 'TEXTO DA OPÇÃO' pelo texto real):
# XPATH_OPCAO_SERVICO = "//mat-option[.//span[contains(normalize-space(.), 'TEXTO DA OPÇÃO')]]"

# Passo 2 — caixa de texto do título
XPATH_CAMPO_TITULO = '//*[@id="mat-input-1"]'
TEXTO_TITULO = "CTE Divergente"

# Passo 3 — botão de avançar do primeiro passo do formulário
XPATH_BOTAO_AVANCAR_1 = (
    "/html/body/app-root/app-solicitacao-agilis/app-cabecalho/div/div[2]/div[2]"
    "/div/div[2]/div[3]/div/app-solicitacao-passo/div/div/div/div/div[2]/div"
    "/div[2]/app-solicitacao-fluxos/div/div/div[2]/div/button"
)

# Passo 4 — campo de descrição do erro
XPATH_CAMPO_ERRO = '//*[@id="mat-input-4"]'

# Passo 6 — opção do primeiro dropdown
XPATH_OPCAO_DROPDOWN_1 = '//*[@id="mat-option-8"]'
# Alternativa por texto visível:
# XPATH_OPCAO_DROPDOWN_1 = "//mat-option[.//span[contains(normalize-space(.), 'TEXTO DA OPÇÃO')]]"

# Passo 8 — opção do segundo dropdown
XPATH_OPCAO_DROPDOWN_2 = '//*[@id="mat-option-72"]'
# Alternativa por texto visível:
# XPATH_OPCAO_DROPDOWN_2 = "//mat-option[.//span[contains(normalize-space(.), 'TEXTO DA OPÇÃO')]]"

# Passo 10 — botão de avançar do segundo passo do formulário
XPATH_BOTAO_AVANCAR_2 = (
    "/html/body/app-root/app-solicitacao-agilis/app-cabecalho/div/div[2]/div[2]"
    "/div/div[2]/app-solicitacao-fluxos-renderizacao/div/div[4]"
    "/app-solicitacao-passo/div/div/div/div/div[2]/div/button"
)

# Passo 12 — campo do pedido de frete
XPATH_CAMPO_PEDIDO_FRETE = '//*[@id="mat-input-11"]'

# Passo 20 — ícone de anexo (usado apenas no fallback via janela nativa)
XPATH_ICONE_ANEXO = (
    '//*[@id="completo"]/div[1]/div/app-solicitacao-passo/div/div/div/div'
    "/div[2]/div/div[2]/app-solicitacao-anexo/mat-form-field/div/div[1]"
    "/div[1]/div/mat-toolbar/mat-icon"
)

# Passo 20/21 (preferencial) — input de arquivo oculto do componente de anexo.
# O Selenium consegue mandar o caminho direto para <input type="file"> mesmo
# que ele esteja invisível, sem abrir a janela nativa do Windows.
XPATH_INPUT_ARQUIVO = "//app-solicitacao-anexo//input[@type='file']"
XPATH_INPUT_ARQUIVO_GENERICO = "//input[@type='file']"

# Passo 22 — botão de finalizar
XPATH_BOTAO_FINALIZAR = (
    '//*[@id="completo"]/div[2]/div/app-solicitacao-passo/div/div/div/div'
    "/div[2]/div/div[2]/app-solicitacao-finalizar/div/div[3]/button"
)

# Passo 23 — elemento com o texto "Nº DO CHAMADO: NNNNNN"
XPATH_NUMERO_CHAMADO = (
    "/html/body/app-root/app-solicitacao-agilis/app-cabecalho/div/div[2]"
    "/div[2]/div/app-solicitacao-lista-concluido/div/div[2]/div[1]"
    "/app-solicitacao-passo/div/div/div/div/div[2]/div/div[1]"
)

# =========================================================================
# LOGGING — console + arquivo, com timestamp
# =========================================================================

logger = logging.getLogger("robo_agilis")
logger.setLevel(logging.INFO)
_formato = logging.Formatter(
    "%(asctime)s | %(levelname)s | %(message)s", datefmt="%d/%m/%Y %H:%M:%S"
)
_handler_arquivo = logging.FileHandler(ARQUIVO_LOG, encoding="utf-8")
_handler_arquivo.setFormatter(_formato)
_handler_console = logging.StreamHandler(sys.stdout)
_handler_console.setFormatter(_formato)
logger.addHandler(_handler_arquivo)
logger.addHandler(_handler_console)

# =========================================================================
# HELPERS DE INTERAÇÃO — sempre com WebDriverWait + retry
# =========================================================================


def esperar_elemento(driver, xpath, condicao=EC.presence_of_element_located, timeout=TIMEOUT):
    """Espera explícita por um elemento, segundo a condição informada."""
    return WebDriverWait(driver, timeout).until(condicao((By.XPATH, xpath)))


def esperar_e_clicar(driver, xpath, descricao=""):
    """
    Espera o elemento ficar clicável e clica, com até MAX_TENTATIVAS tentativas.
    Trata StaleElement, Timeout e ElementClickIntercepted; se o clique normal
    falhar na última tentativa, faz fallback para clique via JavaScript.
    """
    ultimo_erro = None
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            elemento = esperar_elemento(driver, xpath, EC.element_to_be_clickable)
            try:
                elemento.click()
            except (ElementClickInterceptedException, ElementNotInteractableException):
                # Fallback: clique via JavaScript (ignora sobreposições/overlays)
                logger.warning(
                    "Clique normal interceptado em '%s' — usando clique via JavaScript.",
                    descricao or xpath,
                )
                driver.execute_script("arguments[0].click();", elemento)
            return
        except (StaleElementReferenceException, TimeoutException, WebDriverException) as erro:
            ultimo_erro = erro
            logger.warning(
                "Tentativa %d/%d de clicar em '%s' falhou: %s",
                tentativa, MAX_TENTATIVAS, descricao or xpath, type(erro).__name__,
            )
            time.sleep(PAUSA_ENTRE_TENTATIVAS)
    raise TimeoutException(
        f"Não foi possível clicar em '{descricao or xpath}' após "
        f"{MAX_TENTATIVAS} tentativas. Último erro: {ultimo_erro}"
    )


def esperar_e_digitar(driver, xpath, texto, descricao="", limpar=True):
    """
    Espera o campo ficar visível, clica nele e digita o texto,
    com até MAX_TENTATIVAS tentativas.
    """
    ultimo_erro = None
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            elemento = esperar_elemento(driver, xpath, EC.visibility_of_element_located)
            elemento.click()
            if limpar:
                elemento.clear()
            elemento.send_keys(texto)
            return
        except (
            StaleElementReferenceException,
            TimeoutException,
            ElementClickInterceptedException,
            ElementNotInteractableException,
            WebDriverException,
        ) as erro:
            ultimo_erro = erro
            logger.warning(
                "Tentativa %d/%d de digitar em '%s' falhou: %s",
                tentativa, MAX_TENTATIVAS, descricao or xpath, type(erro).__name__,
            )
            time.sleep(PAUSA_ENTRE_TENTATIVAS)
    raise TimeoutException(
        f"Não foi possível digitar em '{descricao or xpath}' após "
        f"{MAX_TENTATIVAS} tentativas. Último erro: {ultimo_erro}"
    )


def pressionar_tab(driver):
    """Envia Tab para o elemento atualmente em foco."""
    driver.switch_to.active_element.send_keys(Keys.TAB)


def pressionar_enter(driver):
    """Envia Enter para o elemento atualmente em foco."""
    driver.switch_to.active_element.send_keys(Keys.ENTER)


def digitar_no_foco(driver, texto):
    """Digita o texto no elemento atualmente em foco."""
    driver.switch_to.active_element.send_keys(texto)


def tab_pausa_enter(driver):
    """Sequência recorrente do fluxo: Tab → espera 0,5s → Enter."""
    pressionar_tab(driver)
    time.sleep(PAUSA_TAB_ENTER)
    pressionar_enter(driver)


# =========================================================================
# FUNÇÕES PRINCIPAIS
# =========================================================================


def validar_pre_requisitos():
    """
    Valida a existência do msedgedriver, da planilha e do anexo antes de
    iniciar. Se algo faltar, registra erro claro e encerra o programa.
    """
    faltando = []
    if not os.path.isfile(CAMINHO_EDGEDRIVER):
        faltando.append(f"Edge WebDriver não encontrado: {CAMINHO_EDGEDRIVER}")
    if not os.path.isfile(CAMINHO_PLANILHA):
        faltando.append(f"Planilha não encontrada: {CAMINHO_PLANILHA}")
    if not os.path.isfile(CAMINHO_ANEXO):
        faltando.append(f"Anexo não encontrado: {CAMINHO_ANEXO}")

    if faltando:
        for mensagem in faltando:
            logger.error(mensagem)
        logger.error("Pré-requisitos ausentes — encerrando o robô.")
        sys.exit(1)

    logger.info("Pré-requisitos validados: driver, planilha e anexo encontrados.")


def inicializar_driver():
    """Inicializa o Microsoft Edge com o WebDriver local."""
    logger.info("Inicializando o Microsoft Edge WebDriver...")
    opcoes = EdgeOptions()
    opcoes.add_argument("--start-maximized")
    # Evita logs verbosos do próprio Edge no console
    opcoes.add_experimental_option("excludeSwitches", ["enable-logging"])
    servico = Service(executable_path=CAMINHO_EDGEDRIVER)
    driver = webdriver.Edge(service=servico, options=opcoes)
    driver.set_page_load_timeout(TIMEOUT * 2)  # rede instável: tolerância extra
    logger.info("Driver inicializado com sucesso.")
    return driver


def ler_planilha():
    """
    Abre a planilha com openpyxl (preservando o arquivo original) e monta a
    lista de linhas a processar. Linhas com a coluna G já preenchida são
    puladas (idempotência — permite retomar de onde parou).
    Retorna (workbook, worksheet, lista_de_dados).
    """
    logger.info("Abrindo a planilha: %s", CAMINHO_PLANILHA)
    wb = load_workbook(CAMINHO_PLANILHA)
    ws = wb.active

    linhas = []
    for numero_linha in range(LINHA_INICIAL, ws.max_row + 1):
        pedido_frete = ws[f"{COL_PEDIDO_FRETE}{numero_linha}"].value
        ordem_coleta = ws[f"{COL_ORDEM_COLETA}{numero_linha}"].value
        pedido_material = ws[f"{COL_PEDIDO_MATERIAL}{numero_linha}"].value
        erro = ws[f"{COL_ERRO}{numero_linha}"].value
        chamado_existente = ws[f"{COL_CHAMADO}{numero_linha}"].value

        # Linha totalmente vazia (sem pedido de frete): ignora
        if pedido_frete is None or str(pedido_frete).strip() == "":
            continue

        # Coluna G já preenchida: chamado já foi aberto — pula (retomada)
        if chamado_existente is not None and str(chamado_existente).strip() != "":
            logger.info(
                "Linha %d | Pedido %s | JÁ PROCESSADA (chamado %s) — pulando.",
                numero_linha, str(pedido_frete).strip(), str(chamado_existente).strip(),
            )
            continue

        linhas.append(
            {
                "linha": numero_linha,
                "pedido_frete": str(pedido_frete).strip(),
                "ordem_coleta": str(ordem_coleta).strip() if ordem_coleta is not None else "",
                "pedido_material": str(pedido_material).strip() if pedido_material is not None else "",
                "erro": str(erro).strip() if erro is not None else "",
            }
        )

    logger.info("Planilha lida: %d linha(s) pendente(s) de processamento.", len(linhas))
    return wb, ws, linhas


def anexar_arquivo(driver):
    """
    Passos 20 e 21 — anexa o PPTX do comunicado.

    ABORDAGEM PREFERENCIAL (ativa): localiza o <input type="file"> do
    componente de anexo (mesmo oculto) e envia o caminho direto com
    send_keys(), SEM clicar no ícone. Evita a janela nativa do Windows,
    que o Selenium não controla — muito mais estável.

    FALLBACK (comentado abaixo): se o input de arquivo não existir no DOM,
    clica no ícone, espera ~4s e usa pyautogui para digitar o caminho na
    janela nativa. ATENÇÃO: o fallback depende de a janela nativa estar em
    FOCO — não mexa no mouse/teclado durante a execução e não bloqueie a
    tela. Para usar, comente a abordagem preferencial e descomente o bloco.
    """
    # ---------------- ABORDAGEM PREFERENCIAL (ativa) ----------------
    ultimo_erro = None
    for xpath_input in (XPATH_INPUT_ARQUIVO, XPATH_INPUT_ARQUIVO_GENERICO):
        try:
            input_arquivo = WebDriverWait(driver, TIMEOUT).until(
                EC.presence_of_element_located((By.XPATH, xpath_input))
            )
            # O input costuma vir oculto (display:none); o send_keys funciona
            # mesmo assim, mas garantimos a visibilidade para navegadores
            # mais chatos, sem alterar o comportamento do componente.
            driver.execute_script(
                "arguments[0].style.display = 'block';"
                "arguments[0].style.visibility = 'visible';",
                input_arquivo,
            )
            input_arquivo.send_keys(CAMINHO_ANEXO)
            logger.info("Anexo enviado via <input type='file'>: %s", CAMINHO_ANEXO)
            # Respiro para o upload do arquivo concluir antes de finalizar
            time.sleep(PAUSA_JANELA_ANEXO)
            return
        except TimeoutException as erro:
            ultimo_erro = erro
            logger.warning(
                "Input de arquivo não encontrado pelo XPath '%s' — tentando o próximo.",
                xpath_input,
            )

    raise TimeoutException(
        "Nenhum <input type='file'> encontrado no DOM para anexar o arquivo. "
        "Avalie usar o fallback com pyautogui (bloco comentado em anexar_arquivo). "
        f"Último erro: {ultimo_erro}"
    )

    # ---------------- FALLBACK via janela nativa (comentado) ----------------
    # Descomente este bloco (e comente a abordagem preferencial acima) caso o
    # componente de anexo não exponha um <input type="file"> no DOM.
    # Requer: pip install pyautogui  (já listado no requirements.txt)
    #
    # import pyautogui  # importe no topo do arquivo se for usar o fallback
    #
    # # Passo 20 — clica no ícone de anexo (abre a janela nativa do Windows)
    # esperar_e_clicar(driver, XPATH_ICONE_ANEXO, "ícone de anexo")
    #
    # # Passo 21 — aguarda a janela nativa abrir e digita o caminho completo.
    # # IMPORTANTE: a janela nativa precisa estar em foco. Não use o computador
    # # durante esta etapa.
    # time.sleep(PAUSA_JANELA_ANEXO)
    # pyautogui.write(CAMINHO_ANEXO, interval=0.02)
    # time.sleep(1.0)
    # pyautogui.press("enter")
    # time.sleep(PAUSA_JANELA_ANEXO)  # respiro para o upload concluir
    # logger.info("Anexo enviado via janela nativa (pyautogui): %s", CAMINHO_ANEXO)


def capturar_numero_chamado(driver):
    """
    Passo 23 — lê o texto do elemento de conclusão e extrai apenas o número
    que vem depois de 'Nº DO CHAMADO:'.
    """
    elemento = esperar_elemento(
        driver, XPATH_NUMERO_CHAMADO, EC.visibility_of_element_located
    )
    texto = elemento.text or ""
    # Aceita variações de caixa/espaçamento: "Nº DO CHAMADO: 123456"
    correspondencia = re.search(r"N[ºO°]?\s*DO\s*CHAMADO\s*:\s*(\d+)", texto, re.IGNORECASE)
    if not correspondencia:
        raise ValueError(
            f"Não foi possível extrair o número do chamado do texto: '{texto}'"
        )
    numero = correspondencia.group(1)
    logger.info("Número do chamado capturado: %s", numero)
    return numero


def gravar_e_salvar(wb, ws, numero_linha, numero_chamado):
    """
    Passo 24 — grava o número do chamado na coluna G da linha e salva a
    planilha IMEDIATAMENTE (não perde progresso se cair rede/energia).
    """
    ws[f"{COL_CHAMADO}{numero_linha}"] = numero_chamado
    wb.save(CAMINHO_PLANILHA)
    logger.info(
        "Linha %d | Chamado %s gravado na coluna %s e planilha salva.",
        numero_linha, numero_chamado, COL_CHAMADO,
    )


def processar_linha(driver, dados_linha):
    """
    Executa o fluxo completo de abertura de chamado (passos 1 a 23) para uma
    linha da planilha. Retorna o número do chamado capturado.
    """
    linha = dados_linha["linha"]
    pedido_frete = dados_linha["pedido_frete"]
    ordem_coleta = dados_linha["ordem_coleta"]
    pedido_material = dados_linha["pedido_material"]
    erro = dados_linha["erro"]

    logger.info("Linha %d | Pedido %s | Iniciando abertura do chamado...", linha, pedido_frete)

    # Reabre a URL no início de CADA linha para garantir formulário limpo
    driver.get(URL_AGILIS)

    # Passo 1 — seleciona o serviço no campo inicial
    esperar_e_clicar(driver, XPATH_CAMPO_SERVICO, "campo de serviço (mat-input-0)")
    esperar_e_clicar(driver, XPATH_OPCAO_SERVICO, "opção do serviço (mat-option-2)")

    # Passo 2 — título do chamado
    esperar_e_digitar(driver, XPATH_CAMPO_TITULO, TEXTO_TITULO, "campo de título (mat-input-1)")

    # Passo 3 — avança para o formulário do fluxo
    esperar_e_clicar(driver, XPATH_BOTAO_AVANCAR_1, "botão avançar (passo 3)")

    # Passo 4 — descrição do erro: "{Erro} Pedido desaprovado - Ordem de coleta: {Ordem}"
    texto_erro = f"{erro} Pedido desaprovado - Ordem de coleta: {ordem_coleta}"
    esperar_e_digitar(driver, XPATH_CAMPO_ERRO, texto_erro, "campo de erro (mat-input-4)")

    # Passo 5 — Tab → 0,5s → Enter (confirma o autocomplete)
    tab_pausa_enter(driver)

    # Passo 6 — Tab e seleciona a opção do primeiro dropdown
    pressionar_tab(driver)
    esperar_e_clicar(driver, XPATH_OPCAO_DROPDOWN_1, "opção do dropdown 1 (mat-option-8)")

    # Passo 7 — Tab → 0,5s → Enter
    tab_pausa_enter(driver)

    # Passo 8 — Tab e seleciona a opção do segundo dropdown
    pressionar_tab(driver)
    esperar_e_clicar(driver, XPATH_OPCAO_DROPDOWN_2, "opção do dropdown 2 (mat-option-72)")

    # Passo 9 — Tab → 0,5s → Enter
    tab_pausa_enter(driver)

    # Passo 10 — botão de avançar do segundo passo
    esperar_e_clicar(driver, XPATH_BOTAO_AVANCAR_2, "botão avançar (passo 10)")

    # Passo 11 — Tab → 0,5s → Enter
    tab_pausa_enter(driver)

    # Passo 12 — pedido de frete (coluna B)
    esperar_e_digitar(
        driver, XPATH_CAMPO_PEDIDO_FRETE, pedido_frete, "campo pedido de frete (mat-input-11)"
    )

    # Passo 13 — Tab → 0,5s → Enter
    tab_pausa_enter(driver)

    # Passo 14 — Tab e digita o pedido de material (coluna D)
    pressionar_tab(driver)
    digitar_no_foco(driver, pedido_material)

    # Passo 15 — Tab → 0,5s → Enter
    tab_pausa_enter(driver)

    # Passo 16 — digita a ordem de coleta (coluna C) no campo em foco
    digitar_no_foco(driver, ordem_coleta)

    # Passo 17 — Tab → 0,5s → Enter
    tab_pausa_enter(driver)

    # Passo 18 — Tab e digita o recado, concatenando os valores da linha
    pressionar_tab(driver)
    recado = (
        f"Chamado aberto pois o pedido {pedido_frete} está com o erro {erro} "
        "e precisa ser solicitado para que sigamos com o pagamento do CTE. "
        "Para pegar mais detalhes sobre o pedido de frete, gentileza validar "
        "a planilha de frete FOB."
    )
    digitar_no_foco(driver, recado)

    # Passo 19 — Tab → 0,5s → Enter
    tab_pausa_enter(driver)

    # Passos 20 e 21 — anexa o comunicado (ver docstring de anexar_arquivo)
    anexar_arquivo(driver)

    # Passo 22 — finaliza o chamado
    esperar_e_clicar(driver, XPATH_BOTAO_FINALIZAR, "botão finalizar (passo 22)")

    # Passo 23 — captura o número do chamado gerado
    numero_chamado = capturar_numero_chamado(driver)

    logger.info(
        "Linha %d | Pedido %s | Chamado %s aberto com sucesso.",
        linha, pedido_frete, numero_chamado,
    )
    return numero_chamado


def salvar_screenshot_erro(driver, numero_linha):
    """Salva um screenshot da tela no momento do erro (erro_linha_{n}.png)."""
    nome_arquivo = f"erro_linha_{numero_linha}.png"
    try:
        driver.save_screenshot(nome_arquivo)
        logger.info("Screenshot do erro salvo em: %s", nome_arquivo)
    except WebDriverException:
        logger.warning("Não foi possível salvar o screenshot da linha %d.", numero_linha)


def main():
    """Orquestra o robô: valida, lê a planilha e processa linha a linha."""
    logger.info("=" * 70)
    logger.info("ROBÔ AGILIS — Abertura de chamados de frete FOB | Iniciando")
    logger.info("=" * 70)

    validar_pre_requisitos()

    wb, ws, linhas = ler_planilha()
    if not linhas:
        logger.info("Nenhuma linha pendente — nada a fazer. Encerrando.")
        return

    sucessos = []
    falhas = []
    driver = None

    try:
        driver = inicializar_driver()

        for dados_linha in linhas:
            numero_linha = dados_linha["linha"]
            pedido_frete = dados_linha["pedido_frete"]
            try:
                numero_chamado = processar_linha(driver, dados_linha)
                # Passo 24 — grava a coluna G e salva a planilha imediatamente
                gravar_e_salvar(wb, ws, numero_linha, numero_chamado)
                sucessos.append((numero_linha, pedido_frete, numero_chamado))
            except Exception as erro:  # noqa: BLE001 — não derruba o robô por uma linha
                logger.error(
                    "Linha %d | Pedido %s | FALHOU: %s — %s",
                    numero_linha, pedido_frete, type(erro).__name__, erro,
                )
                if driver is not None:
                    salvar_screenshot_erro(driver, numero_linha)
                falhas.append((numero_linha, pedido_frete))
                # Continua para a próxima linha
                continue
    finally:
        if driver is not None:
            try:
                driver.quit()
                logger.info("Driver encerrado.")
            except WebDriverException:
                logger.warning("Falha ao encerrar o driver (ignorada).")

    # ------------------------- RESUMO FINAL -------------------------
    logger.info("=" * 70)
    logger.info("RESUMO DA EXECUÇÃO")
    logger.info("Linhas processadas com sucesso: %d", len(sucessos))
    for numero_linha, pedido_frete, numero_chamado in sucessos:
        logger.info(
            "  OK    | Linha %d | Pedido %s | Chamado %s",
            numero_linha, pedido_frete, numero_chamado,
        )
    logger.info("Linhas com falha: %d", len(falhas))
    for numero_linha, pedido_frete in falhas:
        logger.info("  FALHA | Linha %d | Pedido %s", numero_linha, pedido_frete)
    if falhas:
        logger.info(
            "Rode o robô novamente para reprocessar apenas as linhas com falha "
            "(as já concluídas serão puladas pela coluna %s).", COL_CHAMADO,
        )
    logger.info("=" * 70)


if __name__ == "__main__":
    main()
