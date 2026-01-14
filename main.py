from oauth2client.service_account import ServiceAccountCredentials
import gspread

import webbrowser
import pyautogui
import time
from urllib.parse import quote

CAMINHO_CREDENCIAL = "formulariosolicitacaopagamento-f683a63c3e41.json"
PLANILHA_ID = "1lkM9yOjhu_D2nQjRFl-Wt6lNgWPvzl2wbQiaO633-KM"

ABA_BMS = "BMs 2026"
ABA_CONTATOS = "Contatos"

DIVISOR = "\n------------------------------\n"

ESPERA_CARREGAR_WHATSAPP = 12
ESPERA_APOS_ENVIAR = 3

def norm(s: str) -> str:
    if s is None:
        return ""
    return "".join(str(s).strip().lower().split())

def get_value(row: dict, *possible_keys: str) -> str:
    for k in possible_keys:
        if k in row and row[k] is not None:
            return str(row[k]).strip()
    return ""

def find_col_index(headers: list[str], *possible_names: str) -> int:
    h_norm = [norm(h) for h in headers]
    for name in possible_names:
        n = norm(name)
        if n in h_norm:
            return h_norm.index(n) + 1
    return -1

def conectar_google_sheets():
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CAMINHO_CREDENCIAL, scopes)
    return gspread.authorize(creds)

def carregar_contatos(ws_contatos):
    dados = ws_contatos.get_all_records()
    contatos_por_diretoria = {}

    for r in dados:
        nome = get_value(r, "Seu nome", "Nome", "nome")
        tel  = get_value(r, "Número do WhatsA", "Número do WhatsApp", "Numero do WhatsApp", "WhatsApp", "telefone")
        diretoria = get_value(r, "Sua diretoria", "Diretoria", "setor")

        if not nome or not tel or not diretoria:
            continue

        tel_digits = "".join(ch for ch in tel if ch.isdigit())
        if not tel_digits.startswith("55"):
            tel_digits = "55" + tel_digits
        tel_fmt = f"+{tel_digits}"

        contatos_por_diretoria.setdefault(diretoria.strip(), []).append((nome.strip(), tel_fmt))

    return contatos_por_diretoria

def montar_bloco_item(row):
    return (
        f"Objeto do Contrato: {get_value(row, 'Objeto do contrato', 'Objeto do Contrato')}\n"
        f"Local da obra ou serviço: {get_value(row, 'Local da obra ou serviço', 'Local da obra ou serviço ') }\n"
        f"Tipo de BM: {get_value(row, 'Tipo de BM', 'Tipo BM', 'Tipo da BM')}\n"
        f"Empresa: {get_value(row, 'Nome completo da empresa', 'Empresa', 'Nome da empresa')}\n"
        f"Número do Contrato: {get_value(row, 'N° do contrato', 'Nº do contrato', 'N° contrato', 'Nº Contrato', 'N° Contrato')}\n"
        f"BM nº: {get_value(row, 'N° do BM', 'Nº do BM', 'BM nº', 'BM nº ', 'BM', 'BM n°')}\n"
        f"Valor: {get_value(row, 'Valor', 'Valor (R$)', 'Valor R$')}\n"
        f"Fonte de Recursos do Pagamento: {get_value(row, 'Fonte de Recursos', 'Fonte de Recursos do Pagamento')}\n"
        f"Número do SEI: {get_value(row, 'N° do SEI', 'Nº do SEI', 'N° SEI', 'Nº SEI')}"
    )

def montar_mensagem(nome_contato, diretoria, row):
    bloco = montar_bloco_item(row)
    return (
        f"Olá, {nome_contato} da {diretoria}\n"
        f"Você já pode solicitar a disponibilidade financeira para pagamento do(s) contrato(s) abaixo:\n\n"
        f"{bloco}"
    )


def enviar_msg_whatsapp(texto, telefone):
    phone = telefone.replace("+", "")
    url = f"https://web.whatsapp.com/send?phone={phone}&text={quote(texto)}"

    webbrowser.open(url)
    time.sleep(ESPERA_CARREGAR_WHATSAPP)

    time.sleep(ESPERA_APOS_ENVIAR)

    pyautogui.hotkey("ctrl", "w")

def processar_liberados(ws_bms, contatos_por_diretoria):
    print("🔍 Lendo BMs 2026...")

    dados = ws_bms.get_all_records()
    headers = ws_bms.row_values(1)

    col_status = find_col_index(headers, "STATUS", "Status")
    if col_status == -1:
        print("❌ Coluna 'Status/STATUS' não encontrada.")
        return

    setor_key_name = None
    for h in headers:
        if norm(h) == norm("Setor da demanda"):
            setor_key_name = h
            break
    if not setor_key_name:
        print("❌ Coluna 'Setor da demanda' não encontrada.")
        return

    alterados = 0
    enviados = 0
    avisos_sem_contato = 0

    for i, row in enumerate(dados, start=2):
        status = get_value(row, "STATUS", "Status").strip().upper()
        if status != "LIBERADO":
            continue

        diretoria = (row.get(setor_key_name) or "").strip()
        if not diretoria:
            print(f"[AVISO] Linha {i}: sem 'Setor da demanda'. Pulando.")
            continue

        contatos = contatos_por_diretoria.get(diretoria, [])
        if not contatos:
            print(f"[AVISO] Linha {i}: sem contato cadastrado para '{diretoria}'.")
            avisos_sem_contato += 1
        else:
            print(f"\n== Linha {i} | Diretoria: {diretoria} | Destinatários: {len(contatos)} ==")
            for nome, tel in contatos:
                msg = montar_mensagem(nome, diretoria, row)
                print(f"📤 Enviando para {nome} ({tel})")
                enviar_msg_whatsapp(msg, tel)
                enviados += 1

        a1 = gspread.utils.rowcol_to_a1(i, col_status)
        ws_bms.update_acell(a1, "AGUARDANDO SEI")
        alterados += 1

    print("\n✅ Finalizado!")
    print(f"🔁 Linhas alteradas para 'AGUARDANDO SEI': {alterados}")
    print(f"📨 Mensagens enviadas: {enviados}")
    if avisos_sem_contato:
        print(f"⚠️ Linhas sem contato cadastrado: {avisos_sem_contato}")

    input("\n➡️ Pressione ENTER para encerrar")


def main():
    client = conectar_google_sheets()
    sh = client.open_by_key(PLANILHA_ID)

    ws_bms = sh.worksheet(ABA_BMS)
    ws_contatos = sh.worksheet(ABA_CONTATOS)

    contatos_por_diretoria = carregar_contatos(ws_contatos)
    processar_liberados(ws_bms, contatos_por_diretoria)


if __name__ == "__main__":
    main()
