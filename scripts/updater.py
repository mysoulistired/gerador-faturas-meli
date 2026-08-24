"""
Checa e aplica atualizações do GeradorDeFaturas via GitHub Releases.

Só usa a biblioteca padrão (sem dependências externas) para que o check de
atualização funcione mesmo se o usuário não tiver clicado em "Instalar
dependências" ainda.
"""
from __future__ import annotations

import json
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path

REPO = "mysoulistired/gerador-faturas-meli"
ASSET_NAME = "GeradorDeFaturas.exe"
API_URL = f"https://api.github.com/repos/{REPO}/releases/latest"


def _version_tuple(v: str) -> tuple[int, ...]:
    v = v.strip().lstrip("vV")
    parts = []
    for p in v.split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


def check_for_update(current_version: str, timeout: float = 5.0) -> dict | None:
    """Retorna {'version', 'download_url'} se há uma versão mais nova
    publicada no GitHub, ou None. Nunca levanta erro (sem internet, repo
    sem release ainda, GitHub fora do ar etc. tudo vira None)."""
    try:
        req = urllib.request.Request(
            API_URL,
            headers={"User-Agent": "GeradorDeFaturas", "Accept": "application/vnd.github+json"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.load(resp)
        latest_tag = data.get("tag_name", "")
        if _version_tuple(latest_tag) <= _version_tuple(current_version):
            return None
        for asset in data.get("assets", []):
            if asset.get("name") == ASSET_NAME:
                return {"version": latest_tag, "download_url": asset["browser_download_url"]}
        return None
    except (urllib.error.URLError, TimeoutError, OSError, ValueError, KeyError):
        return None


def apply_update(download_url: str, log=print) -> bool:
    """Baixa o .exe novo e agenda a troca para depois que este processo
    terminar (não dá pra sobrescrever o próprio .exe rodando). Quem chamar
    deve encerrar o programa logo em seguida se isto retornar True."""
    if not getattr(sys, "frozen", False):
        log("[!] Rodando a partir do código-fonte (não é o .exe): atualize com 'git pull'.")
        return False

    exe_path = Path(sys.executable)
    new_path = exe_path.with_name(exe_path.stem + ".new.exe")
    bat_path = exe_path.with_name("_atualizar.bat")

    log(f"Baixando atualização de {download_url} ...")
    req = urllib.request.Request(download_url, headers={"User-Agent": "GeradorDeFaturas"})
    with urllib.request.urlopen(req, timeout=120) as resp, open(new_path, "wb") as f:
        f.write(resp.read())

    # O .bat espera este processo liberar o arquivo do .exe (por isso o
    # laço de retry), troca pelo novo, reabre e apaga a si mesmo.
    bat_path.write_text(
        "@echo off\r\n"
        ":retry\r\n"
        "timeout /t 1 /nobreak > nul\r\n"
        f'move /y "{new_path}" "{exe_path}" > nul 2> nul\r\n'
        "if errorlevel 1 goto retry\r\n"
        f'start "" "{exe_path}"\r\n'
        'del "%~f0"\r\n',
        encoding="utf-8",
    )
    subprocess.Popen(
        ["cmd", "/c", str(bat_path)],
        creationflags=subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS,
        close_fds=True,
    )
    log("Atualização baixada, reiniciando...")
    return True
