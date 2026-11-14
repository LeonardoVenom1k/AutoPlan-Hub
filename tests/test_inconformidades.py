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
        self.assertEqual(inconformidades['G2'].value, 'Concluído')
        self.assertEqual(inconformidades['H2'].value, 'Atrasado - Técnico não informou')
        self.assertEqual(
            inconformidades['I2'].value,
            'Veículo liberado após contato',
        )
        self.assertEqual(report, [])

    def test_localiza_resultado_com_horario_e_tecnico_diferentes(self):
        source = {
            '_id': '1748520',
            'data inicio': datetime(2026, 10, 13, 9, 30),
            'placa': 'RTZ5B69',
            'tipo visita': 'Manutenção',
            'tecnico': 'Equipe agendada',
        }
        final = {
            '_id': '1748520',
            'data inicio': datetime(2026, 10, 13, 9, 0),
            'placa': 'RTZ5B69',
            'tipo visita': 'Manutenção',
            'tecnico': 'Diogo de Paula - Jundsar',
            'etapa visita': 'Concluído',
        }
        self.assertIs(engine.find_final_record(source, [final]), final)

    def test_contrato_unico_localiza_resultado_mesmo_com_horario_distante(self):
        source = {
            '_id': '1748520',
            'data inicio': datetime(2026, 10, 13, 9, 0),
            'etapa visita': 'Agendada',
        }
        final = {
            '_id': '1748520',
            'data inicio': datetime(2026, 10, 13, 17, 30),
            'etapa visita': 'Concluído',
        }
        self.assertIsNone(engine.find_final_record(source, [final]))
        self.assertIs(
            engine.find_final_record(source, [final], accept_unique_contract=True),
            final,
        )

    def test_resultado_final_preenche_etapa_e_status_nas_inconformidades(self):
        workbook = Workbook()
        daily = workbook.active
        daily.title = '13-10-2026'
        daily.append(engine.STANDARD_HEADERS)
        daily.append([
            '1748520', 'Cliente', 'Técnico', 'Prestador', 'Manutenção',
            'Agendada', 'Aguardando ação gestor',
            'Atrasado - Técnico informou', None,
            datetime(2026, 10, 13, 9, 0),
        ])
        final = {
            '_id': '1748520', 'id contrato': '1748520',
            'data inicio': datetime(2026, 10, 13, 17, 30),
            'etapa visita': 'Concluído', 'status': None,
            'data visita inicio': datetime(2026, 10, 13, 17, 35),
        }

        engine.rebuild_inconformities(workbook, [], [final], [], [])

        inconformidades = workbook['Inconformidades']
        self.assertEqual(inconformidades['F2'].value, 'Concluído')
        self.assertEqual(inconformidades['G2'].value, 'Concluído')
        self.assertEqual(daily['F2'].value, 'Agendada')
        self.assertEqual(daily['G2'].value, 'Aguardando ação gestor')

    def test_localiza_mesma_visita_entre_planilhas_dos_colegas(self):
        morning = {
            '_id': '12345', '_file': '26-10 MANHÃ.xlsx',
            'data inicio': datetime(2026, 10, 26, 8, 0),
            'placa': 'ABC1D23', 'tipo visita': 'Manutenção',
            'tecnico': 'Equipe inicial', 'etapa visita': 'Agendada',
        }
        afternoon = {
            '_id': '12345', '_file': '26-10 TARDE.xlsx',
            'data inicio': datetime(2026, 10, 26, 10, 30),
            'placa': 'ABC1D23', 'tipo visita': 'Manutenção',
            'tecnico': 'Equipe substituta', 'etapa visita': 'Concluído',
        }
        records = {engine.visit_key(morning): morning}
        self.assertEqual(
            engine.find_similar_visit_key(records, afternoon),
            engine.visit_key(morning),
        )
        report = []
        engine.merge_record(morning, afternoon, report, 26, '12345')
        self.assertEqual(morning['data inicio'], datetime(2026, 10, 26, 10, 30))
        self.assertEqual(morning['tecnico'], 'Equipe substituta')
        self.assertEqual(morning['etapa visita'], 'Concluído')

    def test_nao_mistura_duas_visitas_distantes_no_mesmo_dia(self):
        morning = {
            '_id': '12345', 'data inicio': datetime(2026, 10, 26, 9, 0),
            'placa': 'ABC1D23', 'tipo visita': 'Manutenção',
        }
        afternoon = {
            '_id': '12345', 'data inicio': datetime(2026, 10, 26, 16, 0),
            'placa': 'ABC1D23', 'tipo visita': 'Manutenção',
        }
        self.assertIsNone(engine.find_final_record(afternoon, [morning]))

    def test_motivo_tem_prioridade_sobre_horario(self):
        self.assertEqual(
            engine.compose_information(
                'Chegou às 09h23', datetime(2026, 10, 13, 9, 23),
                'Cancelado pela Maxtrack',
            ),
            'Cancelado pela Maxtrack',
        )

    def test_formata_colunas_de_data_e_hora(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(engine.STANDARD_HEADERS)
        sheet.append([None] * 9 + [datetime(2026, 10, 13, 9, 30)])
        engine.format_datetime_columns(sheet)
        self.assertEqual(sheet['J2'].number_format, 'dd/mm/yyyy hh:mm')
        self.assertGreaterEqual(sheet.column_dimensions['J'].width, 19)


if __name__ == '__main__':
    unittest.main()
