import pytest

from atualizacao.service import ReleaseInfo, is_newer, validate_release, version_tuple


def release(version='0.7.0'):
    return ReleaseInfo(
        version=version,
        filename=f'AutoPlan-Hub-v{version}.exe',
        download_url=(
            'https://raw.githubusercontent.com/LeonardoVenom1k/'
            f'AutoPlan-Hub/main/executavel/AutoPlan-Hub-v{version}.exe'
        ),
        sha256='a' * 64,
    )


def test_compara_versoes_sem_comparacao_textual():
    assert version_tuple('0.10.0') > version_tuple('0.9.9')
    assert is_newer(release('0.7.0'), '0.6.3')
    assert not is_newer(release('0.6.3'), '0.6.3')


def test_manifesto_valido():
    validate_release(release())


def test_rejeita_download_fora_do_github():
    invalid = ReleaseInfo('0.7.0', 'AutoPlan-Hub-v0.7.0.exe', 'https://example.com/app.exe', 'a' * 64)
    with pytest.raises(ValueError):
        validate_release(invalid)
