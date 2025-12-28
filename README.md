# SharePoint Batch Download

Esta ferramenta em Python automatiza o processo de download de anexos (submissões) a partir de feeds RSS do SharePoint (ficheiros XML). O script analisa o XML, identifica links de anexos nas descrições e organiza os ficheiros descarregados em pastas baseadas no nome do autor.

## Funcionalidades

- **Processamento em Lote**: Processa automaticamente todos os ficheiros XML encontrados na diretoria `___Xml` (ou na pasta base).
- **Saída Organizada**: Cria uma pasta específica para cada autor e guarda lá os respetivos anexos.
- **Gestão de Autenticação**: 
  - Utiliza cookies `rtFa` e `FedAuth` para autenticação no SharePoint.
  - Inclui uma janela pop-up (GUI) para colar e atualizar facilmente os cookies quando estes expiram.
  - Suporta a colagem de cabeçalhos de cookies brutos ou exportações JSON.
- **Registos (Logging)**: Gera um ficheiro de log CSV para cada XML processado, detalhando o estado de cada download (Sucesso, Falhou, Existe, URL).
- **Notificações**: Integração opcional com o Telegram para enviar atualizações de estado.
- **Robustez**: Inclui lógica de repetição (retry) com *exponential backoff* para pedidos de rede e higienização de nomes de ficheiros.

## Pré-requisitos

- Python 3.x
- Pacotes Python necessários:
  ```bash
  pip install requests beautifulsoup4 python-telegram-bot
  ```
  *(Note: `tkinter` is required for the GUI, which is typically included with standard Python installations)*

## Setup

1. **Clone the repository**.
2. **Configure Paths**:
   Open `Sharepoint_download_submitions.py` and check the `BASE_FOLDER` variable. You may want to change this to your preferred working directory:
   ```python
   BASE_FOLDER = r"C:\Path\To\Your\Output\Folder"
   ```
3. **Prepare XML Files**:
   Export your SharePoint list/library as an RSS Feed (XML). Place these XML files in a folder named `___Xml` inside your `BASE_FOLDER` (or directly in the `BASE_FOLDER` if the subfolder doesn't exist).

## Usage

Run the script via the terminal:

```bash
python Sharepoint_download_submitions.py
```

### Authentication
If the script cannot find valid cookies (or if they have expired), a popup window will appear.
1. Log in to your SharePoint site in a web browser.
2. Open Developer Tools (F12) -> **Application** (or Storage) -> **Cookies**.
3. Locate the cookies named `rtFa` and `FedAuth`.
4. Copy their values (or the entire cookie header string).
5. Paste them into the script's popup window.

### Telegram Notifications (Optional)
When prompted by the GUI (or by editing `secrets.json` manually), you can provide a Telegram Bot Token and Chat ID to receive notifications when processing finishes.

## Output

- **Files**: Saved in `BASE_FOLDER/<Author_Name>/<Filename>`.
- **Logs**: A CSV file (e.g., `download_log_filename.csv`) is generated in the `BASE_FOLDER` containing the download history.
- **Secrets**: A `secrets.json` file is created locally to store your cookies and tokens securely.


# SharePoint Batch Download

This Python tool automates the process of downloading attachments (submissions) from SharePoint RSS feeds (XML files). It parses the XML, identifies attachment links within the descriptions, and organizes the downloaded files into folders based on the author's name.

## Features

- **Batch Processing**: Automatically processes all XML files found in the `___Xml` directory (or the base folder).
- **Organized Output**: Creates a specific folder for each author and saves their attachments there.
- **Authentication Handling**: 
  - Uses `rtFa` and `FedAuth` cookies for SharePoint authentication.
  - Includes a GUI popup to easily paste and update cookies when they expire.
  - Supports pasting raw cookie headers or JSON exports.
- **Logging**: Generates a CSV log file for each XML processed, detailing the status of every download (Success, Failed, Exists, URL).
- **Notifications**: Optional integration with Telegram to send status updates.
- **Robustness**: Includes retry logic with exponential backoff for network requests and filename sanitization.

## Prerequisites

- Python 3.x
- Required Python packages:
  ```bash
  pip install requests beautifulsoup4 python-telegram-bot
  ```
  *(Note: `tkinter` is required for the GUI, which is typically included with standard Python installations)*

## Setup

1. **Clone the repository**.
2. **Configure Paths**:
   Open `Sharepoint_download_submitions.py` and check the `BASE_FOLDER` variable. You may want to change this to your preferred working directory:
   ```python
   BASE_FOLDER = r"C:\Path\To\Your\Output\Folder"
   ```
3. **Prepare XML Files**:
   Export your SharePoint list/library as an RSS Feed (XML). Place these XML files in a folder named `___Xml` inside your `BASE_FOLDER` (or directly in the `BASE_FOLDER` if the subfolder doesn't exist).

## Usage

Run the script via the terminal:

```bash
python Sharepoint_download_submitions.py
```

### Authentication
If the script cannot find valid cookies (or if they have expired), a popup window will appear.
1. Log in to your SharePoint site in a web browser.
2. Open Developer Tools (F12) -> **Application** (or Storage) -> **Cookies**.
3. Locate the cookies named `rtFa` and `FedAuth`.
4. Copy their values (or the entire cookie header string).
5. Paste them into the script's popup window.

### Telegram Notifications (Optional)
When prompted by the GUI (or by editing `secrets.json` manually), you can provide a Telegram Bot Token and Chat ID to receive notifications when processing finishes.

## Output

- **Files**: Saved in `BASE_FOLDER/<Author_Name>/<Filename>`.
- **Logs**: A CSV file (e.g., `download_log_filename.csv`) is generated in the `BASE_FOLDER` containing the download history.
- **Secrets**: A `secrets.json` file is created locally to store your cookies and tokens securely.
