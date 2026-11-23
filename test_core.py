import io, unittest
from pathlib import Path
from unittest.mock import patch
from openpyxl import load_workbook, Workbook
from core import *

def p(d, a):
    return dict(date=d, amount=a, state='contabilizzata', file='pos.csv', ref='riga', time=d, match=None)

def n(d, a, group=False):
    return dict(date=d, amount=a, file='nota.pdf', ref='pagina 1', description='PRIVATI', used=False, group=group, code='', customer='Privati')

def fake(ps, ns, dates=None):
    with patch('core.parse_pos', return_value=ps), patch('core.parse_notes', return_value=(ns, dates or {x['date'] for x in ns}, [])):
        return analyze([dict(role='pos', name='pos.csv', data=b'p'), dict(role='notes', name='nota.pdf', data=b'n')])

class Rules(unittest.TestCase):

    def test_shift(self):
        r = fake([p('2026-05-20', 43000)], [n('2026-05-21', 43000)])
        self.assertEqual(r['rows'][0]['status'], 'shift')
        self.assertEqual(r['daily'][0]['date'], '2026-05-20')
        self.assertEqual(r['daily'][0]['matched'], 43000)

    def test_priority_no_reuse(self):
        r = fake([p('2026-05-20', 5000), p('2026-05-21', 5000)], [n('2026-05-21', 5000)], {'2026-05-20', '2026-05-21'})
        self.assertEqual([x['status'] for x in r['rows']], ['missing', 'ok'])

    def test_sum(self):
        r = fake([p('2026-06-15', 14404), p('2026-06-15', 2600)], [n('2026-06-15', 17004, True)])
        self.assertEqual(r['counts']['sum'], 2)
        self.assertEqual(r['matched_total'], 17004)

    def test_no_sum_normal(self):
        self.assertEqual(fake([p('2026-06-15', 14404), p('2026-06-15', 2600)], [n('2026-06-15', 17004)])['matched_total'], 0)

    def test_ambiguous_sum(self):
        (chosen, err) = subset_unique(list(enumerate([p('2026-01-01', 100), p('2026-01-01', 200), p('2026-01-01', 100), p('2026-01-01', 200)])), 300)
        self.assertIsNone(chosen)
        self.assertIn('Più', err)

    def test_one_day_only(self):
        self.assertEqual(fake([p('2026-05-20', 43000)], [n('2026-05-22', 43000)])['rows'][0]['status'], 'incomplete')

    def test_formats(self):
        self.assertEqual(parse_pos('x.csv', b'Importo;Data;Stato\n1.750;15/06/2026 05:03:16;Contabilizzata\n')[0]['amount'], 175000)
        wb = Workbook()
        ws = wb.active
        ws.append(['Importo', 'Data', 'Stato'])
        ws.append([1750, datetime(2026, 6, 15), 'Contabilizzata'])
        b = io.BytesIO()
        wb.save(b)
        self.assertEqual(parse_pos('x.xlsx', b.getvalue())[0]['amount'], 175000)

    def test_unknown_state(self):
        row = p('2026-05-20', 123)
        row['state'] = 'preautorizzata'
        r = fake([row], [n('2026-05-20', 123)])
        self.assertEqual(r['total'], 0)
        self.assertEqual(r['rows'][0]['status'], 'state')

    def test_export(self):
        note = n('2026-05-20', 43000)
        note['customer'] = '=HYPERLINK("bad")'
        r = fake([p('2026-05-20', 43000)], [note])
        wb = load_workbook(io.BytesIO(export_excel(r)))
        self.assertEqual(wb['Dettaglio']['D2'].data_type, 's')
        self.assertEqual(wb['Dettaglio']['G2'].data_type, 'f')
        self.assertEqual(wb['Riepilogo giornaliero']['D2'].value, 430)
if __name__ == '__main__':
    unittest.main()
