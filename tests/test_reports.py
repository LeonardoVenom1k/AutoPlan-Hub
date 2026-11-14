import unittest
from datetime import datetime
from pathlib import Path

from app_info import REVIEW_NOTICE
from automacoes.planilha_mae.service import JobConfig, build_report_text as mother_report
from automacoes.planilha_rank.service import RankJobConfig, build_report_text as rank_report


class ProcessingReportTests(unittest.TestCase):
    def test_relatorio_mae_explica_entradas_resultado_e_revisao(self):
        config = JobConfig(
            mode='update', daily_files=('01-10-2026.xlsx', '02-10-2026.xlsx'),
            final_file='final.xlsx', year=2026, month=10,
            mother_file='mae.xlsx',
        )
        text = mother_report(
            config, Path('mae.xlsx'), Path('backup.xlsx'), 12,
            ['Painel criado.', 'CONFLITO ID 10.'], datetime(2026, 11, 14, 12, 0),
        )
        self.assertIn('2. ARQUIVOS UTILIZADOS', text)
        self.assertIn('Relatório final mais recente usado na comparação: final.xlsx', text)
        self.assertIn('Planilhas diárias selecionadas: 2', text)
        self.assertIn('Registros adicionados durante a atualização: 12', text)
        self.assertIn('CONFLITO ID 10.', text)
        self.assertIn(REVIEW_NOTICE, text)

    def test_relatorio_rank_exibe_totais_e_regra_do_mapeamento(self):
        config = RankJobConfig(
            mother_file='mae.xlsx', output_dir='.', year=2026, month=10,
            previous_rank='rank-anterior.xlsx',
        )
        counts = {
            'base_rank': 100, 'occurrences': 20, 'vehicle': 8,
            'mapping_technicians': 15, 'mapping_providers': 9,
            'mapped': 19, 'unmapped': 1,
        }
        text = rank_report(
            config, Path('rank.xlsx'), counts,
            ['Técnico: A -> Supervisão: B (fonte: Técnico)'],
            datetime(2026, 11, 14, 12, 0),
        )
        self.assertIn('Ocorrências com supervisão mapeada: 19', text)
        self.assertIn('Ocorrências sem supervisão mapeada: 1', text)
        self.assertIn('A supervisão é localizada primeiro pelo campo Técnico.', text)
        self.assertIn('Supervisor e Prestador', text)
        self.assertIn(REVIEW_NOTICE, text)


if __name__ == '__main__':
    unittest.main()
