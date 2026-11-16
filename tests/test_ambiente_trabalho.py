from automacoes.ambiente_trabalho.service import (
    WorkEnvironmentConfig,
    load_config,
    save_config,
    selected_web_urls,
)


def test_preferencias_sao_salvas_e_carregadas(tmp_path):
    path = tmp_path / 'preferencias.json'
    original = WorkEnvironmentConfig(
        whatsapp_mode='desktop',
        outlook_enabled=False,
        teams_mode='web',
        notepad_enabled=False,
    )
    save_config(original, path)
    assert load_config(path) == original


def test_urls_respeitam_modo_e_selecao():
    config = WorkEnvironmentConfig(
        whatsapp_mode='desktop',
        outlook_enabled=False,
        teams_mode='web',
        maxtrack_enabled=True,
        zendesk_enabled=False,
        chatgpt_enabled=True,
        downloads_enabled=False,
        notepad_enabled=False,
    )
    urls = selected_web_urls(config)
    assert 'https://web.whatsapp.com/' not in urls
    assert 'https://teams.cloud.microsoft/' in urls
    assert 'https://go.maxtrack.com.br/#List?name=suporte/visitaTecnica' in urls
    assert 'https://chatgpt.com/' in urls


def test_configuracao_invalida_volta_para_web(tmp_path):
    path = tmp_path / 'preferencias.json'
    path.write_text('{"whatsapp_mode": "qualquer"}', encoding='utf-8')
    assert load_config(path).whatsapp_mode == 'web'
