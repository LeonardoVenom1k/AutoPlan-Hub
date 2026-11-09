from datetime import datetime
import unittest

from openpyxl import Workbook

from automacoes.planilha_mae import engine


class InconformidadesTests(unittest.TestCase):
    def test_formata_horario_de_chegada(self):
        self.assertEqual(
            engine.format_arrival(datetime(2026, 10, 26, 12, 35, 48)),
            'Chegou às 12h35',
        )
        self.assertEqual(engine.format_arrival('chegou as 8:5'), 'Chegou às 08h05')
        self.assertEqual(
            engine.format_arrival('Técnico chegou 12 horas e 35 minutos'),
            'Chegou às 12h35',
        )

    def test_motivo_final_vai_para_informacoes(self):
        records = engine.monthly_inconformities([{
            '_id': '123',
            'motivo cancelamento': 'Cancelado pelo cliente',
        }], [])
        self.assertEqual(records[0]['observacoes'], 'Cancelada pelo cliente')
        self.assertEqual(records[0]['informacoes'], 'Cancelado pelo cliente')

    def test_atualiza_somente_a_copia_da_visita_atrasada(self):
        workbook = Workbook()
        daily = workbook.active
        daily.title = '26-10-2026'
        daily.append(engine.STANDARD_HEADERS)
        daily.append([
            '12345', 'Cliente', 'Equipe agendada', 'Prestador', 'Manutenção',
            'Agendada', None, 'Atrasado - Técnico não informou', None,
            datetime(2026, 10, 26, 8, 0), None, None, None, 'ABC1D23', None,
        ])

        final_record = {
            '_id': '12345',
            'id contrato': '12345',
            'tipo visita': 'Manutenção',
            'etapa visita': 'Concluído',
            'data inicio': datetime(2026, 10, 26, 8, 0),
            'data visita inicio': datetime(2026, 10, 26, 8, 26, 1),
            'motivo cancelamento': 'Veículo liberado após contato',
            'placa': 'ABC1D23',
        }

        daily_source = {
            '_id': '12345',
            'id contrato': '12345',
            'tipo visita': 'Manutenção',
            'data inicio': datetime(2026, 10, 26, 8, 0),
            'placa': 'ABC1D23',
            'motivo cancelamento': 'Equipe avisou a central',
        }

        report = []
        engine.rebuild_inconformities(
            workbook, [], [final_record], report, [daily_source]
        )

        # A aba diária é preservada.
        self.assertEqual(daily['F2'].value, 'Agendada')
        self.assertIsNone(daily['I2'].value)

        # Apenas a cópia em Inconformidades recebe o resultado final.
        inconformidades = workbook['Inconformidades']
        self.assertEqual(inconformidades['F2'].value, 'Concluído')
        self.assertEqual(inconformidades['H2'].value, 'Atrasado - Técnico não informou')
        self.assertEqual(
            inconformidades['I2'].value,
            'Chegou às 08h26 | Equipe avisou a central | Veículo liberado após contato',
        )
        self.assertEqual(report, [])


if __name__ == '__main__':
    unittest.main()
