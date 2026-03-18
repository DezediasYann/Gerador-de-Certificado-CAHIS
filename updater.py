import os
import sys
import subprocess
import requests
import tempfile

GITHUB_REPO = "DezediasYann/Gerador-de-Certificado-CAHIS"

def get_current_version():
    try:
        base = sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.dirname(__file__)
        with open(os.path.join(base, 'version.txt'), 'r') as f:
            return f.read().strip()
    except Exception:
        return "v1.0.0"

def check_for_update(allow_prerelease=False):
    try:
        if allow_prerelease:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases"
            res = requests.get(url, timeout=5)
            res.raise_for_status()
            releases = res.json()
            data = next((r for r in releases if r.get("prerelease")), releases[0])
        else:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            res = requests.get(url, timeout=5)
            res.raise_for_status()
            data = res.json()
            
        latest = data.get("tag_name", "")
        if latest and latest != get_current_version():
            changelog = data.get("body", "Sem notas de lançamento adicionais.")
            if data.get("assets"):
                return latest, data["assets"][0]["browser_download_url"], changelog
    except Exception:
        pass
    return None, None, None

def download_and_update(download_url):
    # 1. Pega o caminho da pasta temporária do Windows (permissão de escrita livre)
    temp_dir = tempfile.gettempdir()
    installer_path = os.path.join(temp_dir, "Instalador_CAHIS_Update.exe")
    
    # 2. Baixa o instalador do GitHub direto para a pasta temporária
    res = requests.get(download_url, stream=True)
    with open(installer_path, 'wb') as f:
        for chunk in res.iter_content(8192):
            f.write(chunk)
            
    # 3. Roda o instalador silenciosamente
    # O Inno Setup vai pedir a permissão de administrador na tela automaticamente
    subprocess.Popen(
        [installer_path, '/VERYSILENT', '/SUPPRESSMSGBOXES', '/FORCECLOSEAPPLICATIONS'],
        creationflags=0x00000008
    )
    
    # 4. Encerra o programa atual para que o instalador possa substituir os arquivos
    os._exit(0)