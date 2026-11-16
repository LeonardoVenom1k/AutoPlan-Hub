"""Atualização segura do executável publicado no GitHub."""
from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlparse

from app_info import APP_VERSION


MANIFEST_URL = (
    'https://raw.githubusercontent.com/LeonardoVenom1k/AutoPlan-Hub/'
    'main/executavel/manifest.json'
)
ALLOWED_DOWNLOAD_HOSTS = {'raw.githubusercontent.com', 'github.com', 'objects.githubusercontent.com'}
FILENAME_PATTERN = re.compile(r'^AutoPlan-Hub-v\d+\.\d+\.\d+\.exe$')


@dataclass(frozen=True)
class ReleaseInfo:
    version: str
    filename: str
    download_url: str
    sha256: str


def version_tuple(version: str) -> tuple[int, int, int]:
    match = re.fullmatch(r'(\d+)\.(\d+)\.(\d+)', version.strip())
    if not match:
        raise ValueError(f'Versão inválida: {version}')
    return tuple(int(part) for part in match.groups())


def fetch_release(timeout: int = 12) -> ReleaseInfo:
    request = urllib.request.Request(
        MANIFEST_URL,
        headers={'User-Agent': f'AutoPlan-Hub/{APP_VERSION}'},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        payload = json.loads(response.read().decode('utf-8'))
    release = ReleaseInfo(
        version=str(payload['version']),
        filename=str(payload['filename']),
        download_url=str(payload['download_url']),
        sha256=str(payload['sha256']).lower(),
    )
    validate_release(release)
    return release


def validate_release(release: ReleaseInfo) -> None:
    version_tuple(release.version)
    if not FILENAME_PATTERN.fullmatch(release.filename):
        raise ValueError('Nome do executável publicado é inválido.')
    if release.filename != f'AutoPlan-Hub-v{release.version}.exe':
        raise ValueError('O nome do executável não corresponde à versão publicada.')
    parsed = urlparse(release.download_url)
    if parsed.scheme != 'https' or parsed.hostname not in ALLOWED_DOWNLOAD_HOSTS:
        raise ValueError('Endereço de download não autorizado.')
    if not re.fullmatch(r'[0-9a-f]{64}', release.sha256):
        raise ValueError('Assinatura SHA-256 inválida no manifesto.')


def is_newer(release: ReleaseInfo, current_version: str = APP_VERSION) -> bool:
    return version_tuple(release.version) > version_tuple(current_version)


def running_executable() -> Path | None:
    if not getattr(sys, 'frozen', False):
        return None
    return Path(sys.executable).resolve()


def download_release(release: ReleaseInfo, destination_dir: Path) -> Path:
    destination_dir.mkdir(parents=True, exist_ok=True)
    temporary = destination_dir / f'{release.filename}.download'
    digest = hashlib.sha256()
    request = urllib.request.Request(
        release.download_url,
        headers={'User-Agent': f'AutoPlan-Hub/{APP_VERSION}'},
    )
    try:
        with urllib.request.urlopen(request, timeout=60) as response, temporary.open('wb') as output:
            while True:
                block = response.read(1024 * 1024)
                if not block:
                    break
                output.write(block)
                digest.update(block)
        if digest.hexdigest().lower() != release.sha256:
            raise ValueError('O arquivo baixado não passou na verificação de integridade.')
        return temporary
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def schedule_replacement(release: ReleaseInfo, downloaded: Path) -> Path:
    current = running_executable()
    if current is None:
        raise RuntimeError('A atualização automática funciona somente no executável do AutoPlan Hub.')
    target = current.parent / release.filename
    script = Path(tempfile.gettempdir()) / f'autoplan-update-{os.getpid()}.cmd'
    script.write_text(
        '@echo off\n'
        'setlocal\n'
        f':wait\n'
        f'tasklist /FI "PID eq {os.getpid()}" /NH | findstr /R /C:" {os.getpid()} " >nul\n'
        'if not errorlevel 1 (\n'
        '  timeout /t 1 /nobreak >nul\n'
        '  goto wait\n'
        ')\n'
        f'move /Y "{downloaded}" "{target}" >nul\n'
        f'if not exist "{target}" exit /b 1\n'
        f'start "" "{target}"\n'
        f'del /F /Q "{current}"\n'
        'del /F /Q "%~f0"\n',
        encoding='utf-8',
    )
    creation_flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
    subprocess.Popen(
        ['cmd.exe', '/c', str(script)],
        creationflags=creation_flags,
        close_fds=True,
    )
    return target


def prepare_update(release: ReleaseInfo) -> Path:
    current = running_executable()
    if current is None:
        raise RuntimeError('Abra o AutoPlan Hub pelo executável para instalar uma atualização.')
    downloaded = download_release(release, current.parent)
    return schedule_replacement(release, downloaded)

