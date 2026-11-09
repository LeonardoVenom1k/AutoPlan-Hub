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
            'placa': 'ABC1D23',
        }

        report = []
        engine.rebuild_inconformities(workbook, [], [final_record], report)

        # A aba diária é preservada.
        self.assertEqual(daily['F2'].value, 'Agendada')
        self.assertIsNone(daily['I2'].value)

        # Apenas a cópia em Inconformidades recebe o resultado final.
        inconformidades = workbook['Inconformidades']
        self.assertEqual(inconformidades['F2'].value, 'Concluído')
        self.assertEqual(inconformidades['H2'].value, 'Atrasado - Técnico não informou')
        self.assertEqual(inconformidades['I2'].value, 'Chegou às 08h26')
        self.assertEqual(report, [])


if __name__ == '__main__':
    unittest.main()
