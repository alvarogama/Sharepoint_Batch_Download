import os, requests, time, csv, random, re, json, unicodedata
import xml.etree.ElementTree as ET
from html import unescape
from bs4 import BeautifulSoup
from urllib.parse import quote, unquote
import tkinter as tk
from tkinter import simpledialog, messagebox

# --- 1. CONFIGURAÇÕES E VARIÁVEIS GLOBAIS ---
BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))
SECRETS_FILE = os.path.join(BASE_FOLDER, "secrets.json")

def atualizar_cookies_gui():
    """Janela Pop-up para colar cookies novos"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    raw_text = simpledialog.askstring(
        "Atualizar Cookies",
        "Cole aqui o texto COMPLETO dos cookies (RTFA, FedAuth, etc):",
        parent=root
    )

    if not raw_text:
        messagebox.showerror("Erro", "Cancelado pelo utilizador.")
        exit()

    novos_cookies = {}
    try:
        parsed = json.loads(raw_text)
        if isinstance(parsed, list):
            for obj in parsed:
                if not isinstance(obj, dict): continue
                name = obj.get('name') or obj.get('Name')
                val = obj.get('value') or obj.get('Value')
                if name and val:
                    if name == 'rtFa': novos_cookies['rtFa'] = val
                    if name == 'FedAuth': novos_cookies['FedAuth'] = val
        elif isinstance(parsed, dict):
            if 'cookies' in parsed and isinstance(parsed['cookies'], dict):
                if 'rtFa' in parsed['cookies']: novos_cookies['rtFa'] = parsed['cookies']['rtFa']
                if 'FedAuth' in parsed['cookies']: novos_cookies['FedAuth'] = parsed['cookies']['FedAuth']
            if 'rtFa' in parsed and isinstance(parsed['rtFa'], str):
                novos_cookies['rtFa'] = parsed['rtFa']
            if 'FedAuth' in parsed and isinstance(parsed['FedAuth'], str):
                novos_cookies['FedAuth'] = parsed['FedAuth']
    except Exception:
        pass

    if not (novos_cookies.get('rtFa') and novos_cookies.get('FedAuth')):
        rtFa_match = re.search(r"rtFa[\"'=:\s]+([^;\"',}\s]+)", raw_text)
        FedAuth_match = re.search(r"FedAuth[\"'=:\s]+([^;\"',}\s]+)", raw_text)

        if not rtFa_match: rtFa_match = re.search(r'"rtFa"\s*:\s*"([^"]+)"', raw_text)
        if not FedAuth_match: FedAuth_match = re.search(r'"FedAuth"\s*:\s*"([^"]+)"', raw_text)
        if not rtFa_match: rtFa_match = re.search(r"rtFa=([^;\s]+)", raw_text)
        if not FedAuth_match: FedAuth_match = re.search(r"FedAuth=([^;\s]+)", raw_text)

        if rtFa_match: novos_cookies['rtFa'] = rtFa_match.group(1).strip().strip('"')
        if FedAuth_match: novos_cookies['FedAuth'] = FedAuth_match.group(1).strip().strip('"')

    if not (novos_cookies.get('rtFa') and novos_cookies.get('FedAuth')):
        messagebox.showerror("Erro", "Não encontrei rtFa/FedAuth. Tente novamente.")
        root.destroy()
        return atualizar_cookies_gui()

    # Carregar dados existentes para não perder os feeds
    data = {"telegram_token": "", "chat_id": 0, "sharepoint_feeds": []}
    if os.path.exists(SECRETS_FILE):
        try:
            with open(SECRETS_FILE, "r", encoding="utf-8") as f:
                existing_data = json.load(f)
                data.update(existing_data) # Mantém feeds e outros dados
        except:
            pass

    data["cookies"] = novos_cookies

    if not data.get("telegram_token"):
        data["telegram_token"] = simpledialog.askstring("Telegram", "Token Bot (Opcional):", parent=root) or ""
        try:
            cid = simpledialog.askstring("Telegram", "Chat ID (Opcional):", parent=root)
            data["chat_id"] = int(cid) if cid else 0
        except:
            data["chat_id"] = 0

    with open(SECRETS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)

    root.destroy()
    return data

def load_secrets():
    if not os.path.exists(SECRETS_FILE):
        return atualizar_cookies_gui()
    try:
        with open(SECRETS_FILE, "r", encoding="utf-8") as f:
            secrets = json.load(f)
            if "cookies" not in secrets or not secrets["cookies"].get("rtFa"):
                return atualizar_cookies_gui()
            return secrets
    except:
        return atualizar_cookies_gui()

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept": "*/*",
    "Connection": "keep-alive",
    "Referer": "https://ismaipt.sharepoint.com/",
}

print("A carregar credenciais...")
secrets = load_secrets()

session = requests.Session()
if secrets and isinstance(secrets.get('cookies'), dict):
    session.cookies.update(secrets['cookies'])
session.headers.update({
    "User-Agent": headers["User-Agent"],
    "Referer": headers["Referer"]
})

MY_TOKEN = secrets.get('telegram_token') if secrets else ""
CHAT_ID = secrets.get('chat_id') if secrets else 0

# Allowed extensions
allowed_extensions = None 

def clean_url(url):
    return quote(unescape(url).strip(), safe=":/")

def extract_attachment_links(description):
    if description is None:
        return []
    soup = BeautifulSoup(description, "html.parser")
    links = [clean_url(a["href"]) for a in soup.find_all("a", href=True)]
    return links

def get_filename_from_url(url):
    raw = os.path.basename(url.split("?")[0])
    decoded = unquote(raw)
    decoded = decoded.replace('+', ' ')
    return decoded

def is_valid_file(url):
    if not allowed_extensions:
        return True
    lower = url.lower()
    return any(lower.endswith(ext) for ext in allowed_extensions)

def download_file(url, folder, retries=6):
    filename = get_filename_from_url(url)
    file_path = os.path.join(folder, filename)

    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        return filename, "Existia"

    attempt = 0
    while attempt < retries:
        try:
            wait_time = random.uniform(1, 2) * (2 ** attempt)
            #print(f"-> GET {url} (attempt {attempt+1}/{retries})")
            response = session.get(url, stream=True, timeout=30, allow_redirects=True)
            response.raise_for_status()

            content_type = response.headers.get("Content-Type", "")
            content_length = response.headers.get("Content-Length")

            if "text/html" in content_type or (content_length and int(content_length) < 100):
                return filename, f"Falhou: HTML recebido (status {response.status_code})"

            with open(file_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
                #print(f"   -> saved file: {file_path}")
                return filename, "Sucesso"
            else:
                return filename, "Falhou: ficheiro vazio"

        except requests.exceptions.RequestException as e:
            print(f"   -> request exception: {e}")
            attempt += 1
            if attempt < retries:
                time.sleep(wait_time + random.uniform(0, 1))
            else:
                return filename, f"Falhou: {e}"

def send_message(text):
    if not MY_TOKEN or not CHAT_ID:
        print(f"Telegram not configured. Message: {text}")
        return False
    try:
        url = f"https://api.telegram.org/bot{MY_TOKEN}/sendMessage"
        payload = {"chat_id": CHAT_ID, "text": str(text)}
        resp = requests.post(url, data=payload, timeout=10)
        return resp.status_code == 200
    except Exception as e:
        print(f"Telegram send failed: {e}")
        return False

def sanitize_name(name: str) -> str:
    if not name:
        return "Unknown"
    name = unicodedata.normalize('NFKD', name)
    name = name.encode('ascii', 'ignore').decode('ascii')
    # Permitir hífens e espaços, remover outros caracteres problemáticos
    name = re.sub(r'[<>:\\"/\\|?*]', '_', name)
    name = re.sub(r'\s+', ' ', name).strip()
    if not name:
        return 'Unknown'
    return name

def process_xml(url_feed, feed_name_fallback):
    """
    url_feed: URL do feed RSS
    feed_name_fallback: Nome definido manualmente no JSON, caso a extração falhe
    """
    print(f"\n--- A processar Feed: {feed_name_fallback} ---")
    
    # Download do XML
    try:
        resp = session.get(url_feed, timeout=30)
        resp.raise_for_status()
        xml_content = resp.content
    except Exception as e:
        print(f"Erro a baixar XML de {url_feed}: {e}")
        return

    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        print(f"Erro parse XML: {e}")
        return

    # --- NOVA LÓGICA: Extrair Título do Canal para Nome da Pasta ---
    channel = root.find("channel")
    folder_context = feed_name_fallback # Fallback inicial

    if channel is not None:
        title_node = channel.find("title")
        if title_node is not None and title_node.text:
            raw_title = title_node.text
            print(f"Título do Feed encontrado: '{raw_title}'")
            
            # Tenta separar por " : "
            # Ex: "FEED RSS para UMAIA : 25-26 : TIC : 1º : ITIC : Turma A: Entrega de Trabalhos"
            parts = raw_title.split(" : ")
            
            # Se tiver pelo menos 6 partes, assumimos que:
            # Índice 4 = Disciplina (ex: ITIC)
            # Índice 5 = Turma (ex: Turma A)
            if len(parts) >= 6:
                subject = parts[4].strip()
                class_name = parts[5].strip()
                # Limpar caracteres extra se houver " : " colado
                class_name = class_name.split(":")[0].strip() 
                
                folder_context = f"{subject} - {class_name}"
            else:
                print("Aviso: Formato do título não corresponde ao padrão esperado para extração. A usar nome do JSON.")

    # Sanitizar o nome da pasta da disciplina
    folder_context = sanitize_name(folder_context)
    #print(f"Pasta de Destino do Feed: [{folder_context}]")

    # CSV Log por feed
    csv_file = os.path.join(BASE_FOLDER, f"log_{folder_context}.csv")
    download_rows = []

    # Iterar itens
    for item in root.findall('.//item'):
        title = item.find('title').text if item.find('title') is not None else ''
        author = item.find('author').text if item.find('author') is not None else 'Unknown'
        author = sanitize_name(author)
        comments = item.find('comments').text if item.find('comments') is not None else ''
        pubDate = item.find('pubDate').text if item.find('pubDate') is not None else ''
        description = item.find('description').text if item.find('description') is not None else ''

        attachment_links = extract_attachment_links(description)

        # Caminho: Base -> Disciplina-Turma -> Autor
        target_folder = os.path.join(BASE_FOLDER, folder_context, author)
        os.makedirs(target_folder, exist_ok=True)

        if not attachment_links:
            download_rows.append([title, author, '', 'Nenhum anexo', pubDate, comments, ''])
            continue

        for link in attachment_links:
            if not link: continue
            #print(f"Ficheiro de {author}: {link}")
            filename, status = download_file(link, target_folder)
            download_rows.append([title, author, filename, status, pubDate, comments, link])
            time.sleep(random.uniform(0.5, 1.5))

    # Guardar CSV
    try:
        with open(csv_file, mode='a', newline='', encoding='utf-8') as f: # Mode 'a' para append ou 'w' se quiseres resetar sempre
            writer = csv.writer(f)
            # Se o ficheiro estiver vazio, escreve header
            if f.tell() == 0:
                writer.writerow(['Title', 'Author', 'Filename', 'Status', 'PubDate', 'Comments', 'URL'])
            writer.writerows(download_rows)
    except Exception as e:
        print(f"Erro a escrever CSV: {e}")
    
    return folder_context

if __name__ == '__main__':
    # Buscar lista de feeds ao ficheiro secrets
    feeds_list = secrets.get("sharepoint_feeds", [])

    if not feeds_list:
        print("Nenhum feed encontrado em 'secrets.json'.")
    else:
        for feed in feeds_list:
            url = feed.get("rss_url")
            nome_fallback = feed.get("nome", "SemNome")
            
            if url:
                processed_folder = process_xml(url, nome_fallback)
                msg = f"Processamento concluído para: {processed_folder}"
                send_message(msg)
            else:
                print(f"Feed '{nome_fallback}' sem URL válido.")
    
    print("--- FIM ---")