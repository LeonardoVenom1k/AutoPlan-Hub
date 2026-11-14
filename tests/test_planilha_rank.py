import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from automacoes.planilha_rank import engine


def record(identifier, provider, stage='Concluído', observation=None, client='Cliente A'):
    return {
        '_id': str(identifier), 'id contrato': str(identifier), 'cliente': client,
        'tecnico': 'Técnico', 'prestador de servico do tecnico': provider,
        'tipo visita': 'Manutenção', 'etapa visita': stage,
        'status': stage, 'observacoes': observation,
        'data inicio': datetime(2026, 10, 1, 8, 0), 'placa': f'ABC{identifier}',
    }


class PlanilhaRankTests(unittest.TestCase):
    def test_painel_supervisao_exibe_supervisor_e_prestadores(self):
        self.assertEqual(
            engine.SUPERVISION_ROW_FIELDS,
            ['Supervisão', 'Prestador de Serviço do Técnico'],
        )

    def test_monta_base_com_concluidas_e_ocorrencias(self):
        daily = [
            record(1, 'Prestador A'),
            record(2, 'Prestador B', 'Cancelada', 'Cancelada pela Maxtrack'),
            record(3, 'Prestador C', 'Cancelada'),
        ]
        inconformities = [
            record(2, 'Prestador B', 'Cancelada', 'No Show(Não Comparecimento do técnico)'),
            record(3, 'Prestador C', 'Cancelada', 'Cancelada pelo cliente'),
        ]
        result = engine.build_base_rank(daily, inconformities)
        self.assertEqual([item['_id'] for item in result], ['1', '2'])
        self.assertEqual(result[1]['etapa visita'], 'No Show(Não Comparecimento do técnico)')

    def test_normaliza_variacoes_de_atraso(self):
        self.assertEqual(
            engine.canonical_penalty('Atrasado - técnico informou'),
            'Atrasado - Técnico informou',
        )
        self.assertEqual(
            engine.canonical_penalty('Atrasado - Técnico Não Informou'),
            'Atrasado - Técnico não informou',
        )

    def test_base_veiculo_usa_somente_as_seis_categorias(self):
        records = [
            record(1, 'A', 'Cancelada', category)
            for category in engine.VEHICLE_CATEGORIES
        ] + [record(9, 'A', 'Cancelada', 'Cancelada pela Maxtrack')]
        result = engine.build_vehicle_base(records)
        self.assertEqual(len(result), 6)
        self.assertEqual({item['observacoes'] for item in result}, set(engine.VEHICLE_CATEGORIES))

    def test_rankings_separam_melhores_e_ocorrencias(self):
        base = [record(index, 'Sem Falhas') for index in range(1, 6)]
        base += [record(10, 'Com Falhas', 'Atrasado - Técnico não informou',
                        'Atrasado - Técnico não informou')]
        _, best, ugly, no_show = engine.summarize_rank(base)
        self.assertEqual(best[0]['provider'], 'Sem Falhas')
        self.assertEqual(ugly[0]['provider'], 'Com Falhas')
        self.assertEqual(no_show, [])

    def test_reaproveita_mapeamento_de_supervisao(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'rank.xlsx'
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = 'Base Supervisão '
            sheet.append(engine.SUPERVISION_HEADERS)
            values = [1, 'Cliente', 'Técnico', 'Prestador A', 'Supervisora A'] + [None] * 11
            sheet.append(values)
            workbook.save(path)
            mapping = engine.load_supervision_mapping(path)
            self.assertEqual(mapping['tecnico::tecnico'], 'Supervisora A')
            self.assertEqual(mapping['prestador::prestador a'], 'Supervisora A')

    def test_supervisao_prioriza_tecnico_quando_prestador_tem_varios(self):
        mapping = {
            'tecnico::tecnico a': 'Supervisora A',
            'tecnico::tecnico b': 'Supervisor B',
        }
        base = engine.build_supervision_base([
            record(1, 'Prestador Único', 'Atrasado - Técnico não informou',
                   'Atrasado - Técnico não informou') | {'tecnico': 'Técnico B'},
        ], mapping)
        self.assertEqual(base[0]['supervisao'], 'Supervisor B')

    def test_log_informa_mapeamento_de_supervisao(self):
        with tempfile.TemporaryDirectory() as folder:
            mother = Path(folder) / 'mae.xlsx'
            previous = Path(folder) / 'rank-anterior.xlsx'
            output = Path(folder) / 'rank.xlsx'

            workbook = Workbook()
            daily = workbook.active
            daily.title = '01-10-2026'
            daily.append(engine.HEADERS)
            inconformities = workbook.create_sheet('Inconformidades')
            inconformities.append(engine.HEADERS)
            inconformities.append([
                1, 'Cliente', 'Técnico A', 'Prestador A', 'Manutenção', 'Cancelada',
                None, 'Atrasado - Técnico não informou', None,
                datetime(2026, 10, 1, 8), None, None, None, 'ABC1D23', None,
            ])
            workbook.save(mother)

            workbook = Workbook()
            base = workbook.active
            base.title = 'Base Supervisão'
            base.append(engine.SUPERVISION_HEADERS)
            base.append([1, 'Cliente', 'Técnico A', 'Prestador A', 'Supervisora A'] + [None] * 11)
            workbook.save(previous)

            report = []
            progress = []
            counts = engine.create_rank_workbook(
                mother, output, previous, report,
                lambda value, message: progress.append((value, message)),
            )

            self.assertEqual(counts['mapped'], 1)
            self.assertEqual(counts['unmapped'], 0)
            self.assertTrue(any('Mapeamento carregado: 1 técnico' in message
                                for _, message in progress))
            self.assertTrue(any('Supervisão aplicada: 1/1' in message
                                for _, message in progress))
            self.assertTrue(any(
                'Técnico: Técnico A -> Supervisão: Supervisora A (fonte: Técnico)' in message
                for _, message in progress
            ))
            self.assertIn('Supervisão aplicada: 1 de 1 ocorrência(s); 0 não mapeada(s).', report)

    def test_cria_estrutura_com_abas_e_tabelas(self):
        with tempfile.TemporaryDirectory() as folder:
            mother = Path(folder) / 'mae.xlsx'
            output = Path(folder) / 'rank.xlsx'
            workbook = Workbook()
            day = workbook.active
            day.title = '01-10-2026'
            day.append(engine.HEADERS)
            day.append([1, 'Cliente', 'Técnico', 'Prestador A', 'Manutenção',
                        'Concluído', None, None, None, datetime(2026, 10, 1, 8),
                        None, None, None, 'ABC1D23', None])
            inc = workbook.create_sheet('Inconformidades')
            inc.append(engine.HEADERS)
            workbook.save(mother)
            engine.create_rank_workbook(mother, output)
            generated = load_workbook(output, data_only=False)
            self.assertEqual(generated.sheetnames, [
                'Painel Inconformidades', 'Inconformidades', 'Base Rank', 'Painel Rank',
                'Rank The Best', 'Rank Ugly', 'Base Veículo Indisponível',
                'Painel Veículo Indisponível', 'Painel Supervisão', 'Base Supervisão',
            ])
            self.assertEqual(generated['Base Rank'].sheet_state, 'hidden')
            self.assertEqual(generated['Rank The Best']['B2'].value, 1)
            self.assertIsNone(generated['Base Rank'].auto_filter.ref)
            self.assertTrue(generated['Rank Ugly']['B1'].alignment.wrap_text)
            generated.close()

    def test_aplica_formatacao_visual_solicitada(self):
        with tempfile.TemporaryDirectory() as folder:
            mother = Path(folder) / 'mae.xlsx'
            output = Path(folder) / 'rank.xlsx'
            workbook = Workbook()
            day = workbook.active
            day.title = '01-10-2026'
            day.append(engine.HEADERS)
            day.append([1, 'Cliente', 'Técnico', 'Prestador A', 'Manutenção',
                        'Concluído', 'Concluído', None, None,
                        datetime(2026, 10, 1, 8), None, None, None, 'ABC1D23', None])
            inc = workbook.create_sheet('Inconformidades')
            inc.append(engine.HEADERS)
            inc.append([2, 'Cliente', 'Técnico', 'Prestador B', 'Manutenção',
                        'Cancelada', 'Cancelada', engine.PENALTY_CATEGORIES[3]])
            inc.append([3, 'Cliente', 'Técnico', 'Prestador C', 'Manutenção',
                        'Cancelada', 'Cancelada', 'Cancelada pela Maxtrack'])
            workbook.save(mother)

            engine.create_rank_workbook(mother, output)

            generated = load_workbook(output, data_only=False)
            rank_ugly = generated['Rank Ugly']
            self.assertEqual(rank_ugly['A1'].value, 'Rótulos de Linha')
            no_show_header = next(
                row for row in range(2, rank_ugly.max_row + 1)
                if rank_ugly.cell(row, 1).value == engine.PENALTY_CATEGORIES[3]
            )
            self.assertGreater(no_show_header, 2)
            self.assertEqual(
                rank_ugly.cell(no_show_header, 2).value.rstrip(),
                engine.PENALTY_CATEGORIES[3],
            )
            self.assertEqual(
                rank_ugly.cell(no_show_header, 1).fill.fgColor.rgb[-6:], engine.RED
            )
            self.assertEqual(rank_ugly.cell(no_show_header + 1, 1).value, 'Prestador B')
            inconformities = generated['Inconformidades']
            self.assertEqual(inconformities['A1'].fill.fgColor.rgb[-6:], engine.BLUE)
            self.assertEqual(inconformities['A1'].font.color.rgb[-6:], engine.WHITE)
            self.assertEqual(
                inconformities['A2'].fill.fgColor.rgb[-6:],
                engine.mother_engine.FILL_INCONFORMITY_A,
            )
            self.assertEqual(
                inconformities['A3'].fill.fgColor.rgb[-6:],
                engine.mother_engine.FILL_INCONFORMITY_B,
            )
            self.assertIsNone(inconformities.tables['TabelaInconformidades'].tableStyleInfo)
            for coordinate in ('A1', 'A2', 'H2'):
                cell = inconformities[coordinate]
                self.assertEqual(cell.border.left.style, 'thin')
                self.assertEqual(cell.border.right.style, 'thin')
                self.assertEqual(cell.border.top.style, 'thin')
                self.assertEqual(cell.border.bottom.style, 'thin')
                self.assertEqual(cell.border.left.color.rgb[-6:], '000000')
            generated.close()


if __name__ == '__main__':
    unittest.main()

