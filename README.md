# SharePoint Batch Download

Esta ferramenta em Python automatiza o processo de download de anexos (submissões) a partir de feeds RSS do SharePoint (ficheiros XML). Ela analisa o XML, identifica os links dos anexos nas descrições e organiza os ficheiros descarregados em pastas com base no nome do autor.

## Funcionalidades

- **Processamento em Lote**: Processa automaticamente todos os ficheiros XML encontrados na pasta `___Xml` (ou na pasta base).
- **Saída Organizada**: Cria uma pasta específica para cada autor e guarda aí os respetivos anexos.
- **Gestão de Autenticação**: 
  - Utiliza os cookies `rtFa` e `FedAuth` para autenticação no SharePoint.
  - Inclui uma janela GUI para colar e atualizar facilmente os cookies quando expiram.
  - Suporta colagem de cabeçalhos de cookies brutos ou exportações em JSON.
- **Registo (Logging)**: Gera um ficheiro CSV de log para cada XML processado, com o estado de cada download (Sucesso, Falhou, Existe, URL).
- **Notificações**: Integração opcional com o Telegram para envio de atualizações de estado.
- **Robustez**: Inclui lógica de repetição (retry) com backoff exponencial para pedidos de rede e sanitização de nomes de ficheiros.

## Pré-requisitos

- Python 3.6+
- Pacotes Python necessários:
- pip install requests beautifulsoup4 python-telegram-bot

*(Nota: o `tkinter` é necessário para a GUI e normalmente já vem incluído nas instalações padrão do Python)*

### macOS Specific Note (Tkinter)
  The script uses a GUI popup for cookies (tkinter).

  Standard Python: If you installed Python from python.org, tkinter is included.
  Homebrew: If you installed Python via Homebrew (brew install python), you might need to install the GUI toolkit separately:
  bash
  brew install python-tk



## Configuração

1. **Clonar o repositório**.
2. **Configurar Caminhos**:
   Por defeito, o script utiliza a pasta onde está localizado como `BASE_FOLDER`.
   Os downloads e ficheiros de configuração ficarão na mesma pasta do script.

3. **Preparar os Ficheiros XML**:
Exporte a sua lista/biblioteca do SharePoint como um feed RSS (XML). Coloque estes ficheiros XML numa pasta chamada `___Xml` dentro da sua `BASE_FOLDER` (ou diretamente na `BASE_FOLDER` se a subpasta não existir).

## Utilização

Execute o script através do terminal:
python Sharepoint_download_submitions.py


### Autenticação
Se o script não encontrar cookies válidos (ou se já tiverem expirado), surgirá uma janela popup.
1. Inicie sessão no seu site SharePoint num navegador web.
2. Abra as Ferramentas de Programador (F12) -> **Application** (ou Armazenamento) -> **Cookies**.
3. Procure os cookies com os nomes `rtFa` e `FedAuth`.
4. Copie os valores (ou todo o cabeçalho de cookies).
5. Cole-os na janela popup do script.

### Notificações no Telegram (Opcional)
Quando solicitado pela GUI (ou editando manualmente o `secrets.json`), pode fornecer um Token de Bot e um ID de Chat do Telegram para receber notificações quando o processamento terminar.

## Saída

- **Ficheiros**: Guardados em `BASE_FOLDER/<Nome_Autor>/<Nome_Ficheiro>`.
- **Registos**: Um ficheiro CSV (por exemplo, `download_log_nomeficheiro.csv`) é gerado em `BASE_FOLDER`, contendo o histórico de downloads.
- **Segredos**: Um ficheiro `secrets.json` é criado localmente para armazenar com segurança os seus cookies e tokens.


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
  
  ### macOS Specific Note (Tkinter)
    The script uses a GUI popup for cookies (tkinter).

    Standard Python: If you installed Python from python.org, tkinter is included.
    Homebrew: If you installed Python via Homebrew (brew install python), you might need to install the GUI toolkit separately:
    bash
    brew install python-tk

## Setup

1. **Clone the repository**.
2. **Configure Paths**:
   By default, the script uses the directory where it is located as the `BASE_FOLDER`.
   Downloads and configuration files will be stored in the same folder as the script.

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
