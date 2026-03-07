import os
import sys
import subprocess
import requests

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
    exe_path = sys.executable
    exe_dir = os.path.dirname(exe_path)
    update_path = exe_path + ".new"
    
    res = requests.get(download_url, stream=True)
    with open(update_path, 'wb') as f:
        for chunk in res.iter_content(8192):
            f.write(chunk)
            
    bat_path = os.path.join(exe_dir, "update.bat")
    
    with open(bat_path, "w") as bat:
        bat.write(
            '@echo off\n'
            ':loop\n'
            'timeout /t 1 /nobreak > NUL\n'
            f'move /y "{update_path}" "{exe_path}" > NUL 2>&1\n'
            'if errorlevel 1 goto loop\n'
            f'cd /d "{exe_dir}"\n'
            f'start "" "{exe_path}"\n'
            f'del "%~f0"\n'
        )
        
    # Limpa as variáveis do PyInstaller para o novo processo não buscar a pasta velha
    env_limpo = os.environ.copy()
    env_limpo.pop('_MEIPASS2', None)
    env_limpo.pop('_MEIPASS1', None)
    env_limpo.pop('_MEIPASS', None)
        
    DETACHED_PROCESS = 0x00000008
    subprocess.Popen(
        [bat_path], 
        shell=True, 
        creationflags=DETACHED_PROCESS, 
        close_fds=True,
        cwd=exe_dir,
        env=env_limpo  # Aplica o ambiente limpo
    )
    
    os._exit(0)