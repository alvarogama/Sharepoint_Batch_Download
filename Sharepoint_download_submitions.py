import os, requests, time, csv, random, asyncio, telegram, re, json, unicodedata
import xml.etree.ElementTree as ET
from html import unescape
from bs4 import BeautifulSoup
from urllib.parse import quote, unquote
import tkinter as tk
from tkinter import simpledialog, messagebox

# --- 1. CONFIGURAÇÕES E VARIÁVEIS GLOBAIS ---
# Define a pasta base como a diretoria onde o script está localizado
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

    # First try: parse pasted text as JSON (handles cookie-export arrays from browser devtools)
    novos_cookies = {}
    try:
        parsed = json.loads(raw_text)
        # If it's a list (cookie export), look for objects with name == rtFa/FedAuth
        if isinstance(parsed, list):
            for obj in parsed:
                if not isinstance(obj, dict):
                    continue
                name = obj.get('name') or obj.get('Name')
                val = obj.get('value') or obj.get('Value')
                if name and val:
                    if name == 'rtFa':
                        novos_cookies['rtFa'] = val
                    if name == 'FedAuth':
                        novos_cookies['FedAuth'] = val
        elif isinstance(parsed, dict):
            # Nested structures like {"cookies": {"rtFa": "..."}}
            if 'cookies' in parsed and isinstance(parsed['cookies'], dict):
                if 'rtFa' in parsed['cookies']:
                    novos_cookies['rtFa'] = parsed['cookies']['rtFa']
                if 'FedAuth' in parsed['cookies']:
                    novos_cookies['FedAuth'] = parsed['cookies']['FedAuth']
            # direct keys
            if 'rtFa' in parsed and isinstance(parsed['rtFa'], str):
                novos_cookies['rtFa'] = parsed['rtFa']
            if 'FedAuth' in parsed and isinstance(parsed['FedAuth'], str):
                novos_cookies['FedAuth'] = parsed['FedAuth']
    except Exception:
        # not JSON — fall back to regex parsing below
        pass

    # If JSON parsing didn't find both, try multiple regex patterns
    if not (novos_cookies.get('rtFa') and novos_cookies.get('FedAuth')):
        rtFa_match = re.search(r"rtFa[\"'=:\s]+([^;\"',}\s]+)", raw_text)
        FedAuth_match = re.search(r"FedAuth[\"'=:\s]+([^;\"',}\s]+)", raw_text)

        # JSON-like: "rtFa":"..."
        if not rtFa_match:
            rtFa_match = re.search(r'"rtFa"\s*:\s*"([^"]+)"', raw_text)
        if not FedAuth_match:
            FedAuth_match = re.search(r'"FedAuth"\s*:\s*"([^"]+)"', raw_text)

        # cookie header style: rtFa=...; FedAuth=...;
        if not rtFa_match:
            rtFa_match = re.search(r"rtFa=([^;\s]+)", raw_text)
        if not FedAuth_match:
            FedAuth_match = re.search(r"FedAuth=([^;\s]+)", raw_text)

        if rtFa_match:
            novos_cookies['rtFa'] = rtFa_match.group(1).strip().strip('"')
        if FedAuth_match:
            novos_cookies['FedAuth'] = FedAuth_match.group(1).strip().strip('"')

    # If still missing, show an error and re-prompt
    if not (novos_cookies.get('rtFa') and novos_cookies.get('FedAuth')):
        messagebox.showerror("Erro", "Não encontrei rtFa/FedAuth no texto. Por favor cole o Cookie header ou o JSON com os campos rtFa e FedAuth.")
        root.destroy()
        return atualizar_cookies_gui()

    data = {"telegram_token": "", "chat_id": 0}
    if os.path.exists(SECRETS_FILE):
        try:
            with open(SECRETS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except:
            pass

    data["cookies"] = novos_cookies

    # Se não tiver token, pede
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

print("🔑 A carregar credenciais...")
secrets = load_secrets()

# Criação da Sessão Global
session = requests.Session()
if secrets and isinstance(secrets.get('cookies'), dict):
    session.cookies.update(secrets['cookies'])
session.headers.update({
    "User-Agent": headers["User-Agent"],
    "Referer": headers["Referer"]
})

MY_TOKEN = secrets.get('telegram_token') if secrets else ""
CHAT_ID = secrets.get('chat_id') if secrets else 0

print("Session cookies:", session.cookies.get_dict())

# Default XML folder inside workspace (there's a ___Xml directory in the repo)
XML_DIR = os.path.join(BASE_FOLDER, "___Xml")
if not os.path.isdir(XML_DIR):
    XML_DIR = BASE_FOLDER

# List all XML files in the XML directory
xml_files = [f for f in os.listdir(XML_DIR) if f.lower().endswith('.xml')]

# Allowed extensions. Set to None to accept all file extensions (no filtering).
allowed_extensions = None  # or set to a set([...]) to restrict
# Extensões permitidas
#allowed_extensions = {".3g2",".3gp",".7z",".aac",".ai",".aif",".aiff",".af",".arw",".avi",".bmp",".bz2",".cr2",".csv",".dng",".doc",".docx",".eps",".flac",".flv",".gif",".gz",".heic",".heif",".htm",".html",".ico",".indd",".jpeg",".jpg",".json",".m4a",".m4v",".mkv",".mov",".mp3",".mp4",".mpeg",".mpg",".nef",".nrw",".odg",".odp",".ods",".odt",".ogg",".orf",".pdf",".pef",".png",".pps",".ppsx",".ppt",".pptx",".psd",".rar",".raw",".rtf",".rw2",".rwl",".sr2",".svg",".tar",".tif",".tiff",".txt",".url",".wav",".webm",".webp",".wma",".wmv",".xls",".xlsx",".xml",".xmp",".xz",".zip",".drp",".sketch",".xcf",".md",".tex",".log",".ini",".yaml",".yml",".toml",".xlsb",".tsv",".sav",".dta",".parquet",".avif",".xcf",".skp",".blend",".stl",".obj",".fbx",".braw",".mts",".vob",".mxf",".opus",".mid",".midi",".caf",".tar.gz",".tar.xz",".zst",".cab",".dmg",".iso",".py",".java",".cpp",".c",".cs",".sh",".bat",".ps1",".r",".jl",".go",".swift",".kt",".dart",".sql",".db",".mdb",".accdb",".sqlite3",".cfg",".plist",".reg"}



def clean_url(url):
    #Corrige espaços e caracteres HTML nas URLs.
    return quote(unescape(url).strip(), safe=":/")


def extract_attachment_links(description):
    #Extrai links corretamente da tag <description>.
    if description is None:
        return []  # Retorna lista vazia se description não existir

    soup = BeautifulSoup(description, "html.parser")
    links = [clean_url(a["href"]) for a in soup.find_all("a", href=True)]
    return links


def get_filename_from_url(url):
    #Extrai o nome do ficheiro da URL.
    raw = os.path.basename(url.split("?")[0])
    # Decode percent-encoding (e.g., "%C3%A9" -> "é ") and plus->space
    decoded = unquote(raw)
    decoded = decoded.replace('+', ' ')
    return decoded


def is_valid_file(url):
    #Verifica se a URL tem uma extensão válida.
    # If allowed_extensions is None, accept all files
    if not allowed_extensions:
        return True
    lower = url.lower()
    return any(lower.endswith(ext) for ext in allowed_extensions)


def download_file(url, folder, retries=6):
    """Download com tentativas e backoff. Retorna (filename, status)."""
    filename = get_filename_from_url(url)
    file_path = os.path.join(folder, filename)

    # If already exists, skip
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        return filename, "Existia"

    attempt = 0
    while attempt < retries:
        try:
            wait_time = random.uniform(1, 2) * (2 ** attempt)
            print(f"-> GET {url} (attempt {attempt+1}/{retries})")
            response = session.get(url, stream=True, timeout=30, allow_redirects=True)
            response.raise_for_status()

            content_type = response.headers.get("Content-Type", "")
            content_length = response.headers.get("Content-Length")

            print(f"   status={response.status_code} content-type={content_type} content-length={content_length}")

            # basic heuristic to avoid HTML errors
            if "text/html" in content_type or (content_length and int(content_length) < 100):
                # save debug HTML to file for inspection
                try:
                    debug_path = os.path.join(folder, f"{filename}.debug.html")
                    with open(debug_path, "wb") as dbg:
                        dbg.write(response.content[:10000])
                    print(f"   -> saved debug HTML: {debug_path}")
                except Exception as e:
                    print(f"   -> failed saving debug html: {e}")
                return filename, f"Falhou: HTML recebido (status {response.status_code})"

            with open(file_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            # small validation
            if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
                print(f"   -> saved file: {file_path} ({os.path.getsize(file_path)} bytes)")
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


def validar_ficheiro(file_path):
    """ Verifica se o ficheiro tem um tamanho razoável """
    if not os.path.exists(file_path):
        return False

    tamanho = os.path.getsize(file_path)
    if tamanho < 1:  # Se for muito pequeno, pode estar corrompido
        print(f"⚠️ Ficheiro suspeito (tamanho pequeno): {file_path}")
        return False
    return True


def send_message(text):
    """Send a Telegram message if token and chat id are configured. Safe no-op otherwise."""
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
    """Sanitize folder/file names to safe filesystem names."""
    if not name:
        return "Unknown"
    # Normalize accents
    name = unicodedata.normalize('NFKD', name)
    name = name.encode('ascii', 'ignore').decode('ascii')
    # Replace problematic chars
    name = re.sub(r'[<>:\\"/\\|?*]', '_', name)
    name = re.sub(r'\s+', ' ', name).strip()
    if not name:
        return 'Unknown'
    return name


def process_xml(xml_path_or_url):
    """Ler o XML (do ficheiro local ou URL), faz o download e guarda o registo num CSV.

    xml_path_or_url: either a local file path or an http(s) URL
    """
    # Load XML content
    xml_content = None
    source_name = xml_path_or_url
    try:
        if str(xml_path_or_url).lower().startswith('http'):
            resp = session.get(xml_path_or_url, timeout=30)
            resp.raise_for_status()
            xml_content = resp.content
        else:
            if not os.path.exists(xml_path_or_url):
                print(f"XML não encontrado: {xml_path_or_url}")
                return
            with open(xml_path_or_url, 'rb') as f:
                xml_content = f.read()
    except Exception as e:
        print(f"Erro a ler XML {xml_path_or_url}: {e}")
        return

    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        print(f"Erro a parsear XML: {e}")
        return

    # Prepare CSV log
    base_name = os.path.splitext(os.path.basename(source_name))[0]
    csv_file = os.path.join(BASE_FOLDER, f"download_log_{base_name}.csv")
    download_rows = []

    for item in root.findall('.//item'):
        title = item.find('title').text if item.find('title') is not None else ''
        author = item.find('author').text if item.find('author') is not None else 'Unknown'
        author = sanitize_name(author)
        comments = item.find('comments').text if item.find('comments') is not None else ''
        pubDate = item.find('pubDate').text if item.find('pubDate') is not None else ''
        description = item.find('description').text if item.find('description') is not None else ''

        attachment_links = extract_attachment_links(description)

        target_folder = os.path.join(BASE_FOLDER, author)
        os.makedirs(target_folder, exist_ok=True)

        if not attachment_links:
            # nothing to download, write a row indicating no attachments
            download_rows.append([title, author, '', 'Nenhum anexo encontrado', pubDate, comments])
            continue

        for link in attachment_links:
            if not link:
                continue
            print(f"Encontrado ficheiro: {link} (autor: {author})")
            filename, status = download_file(link, target_folder)
            download_rows.append([title, author, filename, status, pubDate, comments, link])
            time.sleep(random.uniform(0.5, 2.0))

    # Save CSV
    try:
        with open(csv_file, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Title', 'Author', 'Filename', 'Status', 'PubDate', 'Comments', 'URL'])
            writer.writerows(download_rows)
        print(f"📄 CSV criado: {csv_file}")
    except Exception as e:
        print(f"Erro a escrever CSV: {e}")


if __name__ == '__main__':
    # Process each XML file found in XML_DIR
    for xml_file in xml_files:
        xml_path = os.path.join(XML_DIR, xml_file)
        print(f"Processing: {xml_path}")
        process_xml(xml_path)
        msg = f"Processing finished: {xml_path}"
        send_message(msg)
