"""Preferências persistentes e abertura segura do ambiente de trabalho."""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import webbrowser
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable


WEB_URLS = {
    'whatsapp': 'https://web.whatsapp.com/',
    'outlook': 'https://outlook.cloud.microsoft/mail/',
    'teams': 'https://teams.cloud.microsoft/',
    'maxtrack': 'https://go.maxtrack.com.br/#List?name=suporte/visitaTecnica',
    'zendesk': 'https://maxtrackhelp.zendesk.com/agent/home/tickets',
    'chatgpt': 'https://chatgpt.com/',
}

DESKTOP_URIS = {
    'whatsapp': 'whatsapp:',
    'outlook': 'ms-outlook:',
    'teams': 'msteams:',
}


@dataclass
class WorkEnvironmentConfig:
    whatsapp_enabled: bool = True
    whatsapp_mode: str = 'web'
    outlook_enabled: bool = True
    outlook_mode: str = 'web'
    teams_enabled: bool = True
    teams_mode: str = 'web'
    maxtrack_enabled: bool = True
    zendesk_enabled: bool = True
    chatgpt_enabled: bool = True
    downloads_enabled: bool = True
    notepad_enabled: bool = True
    chrome_restores_tabs: bool = False


def config_path() -> Path:
    base = Path(os.environ.get('APPDATA') or Path.home())
    return base / 'AutoPlan Hub' / 'ambiente_trabalho.json'


def load_config(path: Path | None = None) -> WorkEnvironmentConfig:
    target = path or config_path()
    if not target.is_file():
        return WorkEnvironmentConfig()
    try:
        raw = json.loads(target.read_text(encoding='utf-8'))
        allowed = WorkEnvironmentConfig.__dataclass_fields__
        values = {key: value for key, value in raw.items() if key in allowed}
        config = WorkEnvironmentConfig(**values)
        for name in ('whatsapp_mode', 'outlook_mode', 'teams_mode'):
            if getattr(config, name) not in {'web', 'desktop'}:
                setattr(config, name, 'web')
        return config
    except (OSError, ValueError, TypeError):
        return WorkEnvironmentConfig()


def save_config(config: WorkEnvironmentConfig, path: Path | None = None) -> Path:
    target = path or config_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_suffix('.tmp')
    temporary.write_text(
        json.dumps(asdict(config), ensure_ascii=False, indent=2),
        encoding='utf-8',
    )
    temporary.replace(target)
    return target


def find_chrome() -> Path | None:
    found = shutil.which('chrome') or shutil.which('chrome.exe')
    if found:
        return Path(found)
    roots = [
        os.environ.get('PROGRAMFILES'),
        os.environ.get('PROGRAMFILES(X86)'),
        os.environ.get('LOCALAPPDATA'),
    ]
    candidates = []
    for root in roots:
        if root:
            candidates.append(Path(root) / 'Google' / 'Chrome' / 'Application' / 'chrome.exe')
    return next((item for item in candidates if item.is_file()), None)


def selected_web_urls(config: WorkEnvironmentConfig) -> list[str]:
    urls = []
    for app in ('whatsapp', 'outlook', 'teams'):
        if getattr(config, f'{app}_enabled') and getattr(config, f'{app}_mode') == 'web':
            urls.append(WEB_URLS[app])
    for app in ('maxtrack', 'zendesk', 'chatgpt'):
        if getattr(config, f'{app}_enabled'):
            urls.append(WEB_URLS[app])
    return urls


def _open_uri(uri: str) -> None:
    if os.name != 'nt':
        raise OSError('Aplicativos Desktop estão disponíveis somente no Windows.')
    os.startfile(uri)  # type: ignore[attr-defined]


def launch_environment(
    config: WorkEnvironmentConfig,
    uri_opener: Callable[[str], None] = _open_uri,
) -> tuple[list[str], list[str]]:
    """Abre os itens selecionados e devolve (sucessos, falhas)."""
    successes, failures = [], []

    for app, label in (('whatsapp', 'WhatsApp'), ('outlook', 'Outlook'), ('teams', 'Teams')):
        if getattr(config, f'{app}_enabled') and getattr(config, f'{app}_mode') == 'desktop':
            try:
                uri_opener(DESKTOP_URIS[app])
                successes.append(f'{label} Desktop')
            except Exception as exc:
                failures.append(f'{label} Desktop: {exc}')

    urls = selected_web_urls(config)
    if urls:
        try:
            chrome = find_chrome()
            if chrome:
                command = [str(chrome)] if config.chrome_restores_tabs else [str(chrome), *urls]
                subprocess.Popen(command)
                if config.chrome_restores_tabs:
                    successes.append('Chrome com as abas da sessão anterior')
                else:
                    successes.append(f'{len(urls)} ferramenta(s) no navegador')
            else:
                for url in urls:
                    webbrowser.open_new_tab(url)
                successes.append(f'{len(urls)} ferramenta(s) no navegador padrão')
        except Exception as exc:
            failures.append(f'Navegador: {exc}')

    if config.downloads_enabled:
        try:
            downloads = Path.home() / 'Downloads'
            if not downloads.is_dir():
                raise FileNotFoundError(f'Pasta não encontrada: {downloads}')
            _open_uri(str(downloads))
            successes.append('Pasta Downloads')
        except Exception as exc:
            failures.append(f'Pasta Downloads: {exc}')

    if config.notepad_enabled:
        try:
            subprocess.Popen(['notepad.exe'])
            successes.append('Bloco de Notas')
        except Exception as exc:
            failures.append(f'Bloco de Notas: {exc}')

    return successes, failures

