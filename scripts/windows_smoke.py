"""Exercise the built exe on Windows with synthetic data only."""
import base64,io,json,subprocess,tempfile
from pathlib import Path
from reportlab.pdfgen import canvas
from openpyxl import Workbook

def pdf_bytes(lines):
    b=io.BytesIO();c=canvas.Canvas(b)
    for i,line in enumerate(lines):c.drawString(30,800-i*20,line)
    c.save();return b.getvalue()
def file(role,name,data):return dict(role=role,name=name,content=base64.b64encode(data).decode())
with tempfile.TemporaryDirectory() as tmp:
    folder=Path(tmp)
    wb=Workbook();ws=wb.active;ws.append(['Importo','Data','Stato']);ws.append([10,'10/01/2026','Contabilizzata']);ws.append([20,'10/01/2026','Contabilizzata']);b=io.BytesIO();wb.save(b)
    files=[file('pos','pos.csv',b'Importo;Data;Stato\n50;10/01/2026;Contabilizzata\n90;10/01/2026;Negata\n'),file('pos','pos.xlsx',b.getvalue()),file('notes','nota.pdf',pdf_bytes(['AZIENDA DEMO DATA 11/01/2026','020102 800101 50,00 CD 123 PAG A MZ POS 1','AZIENDA DEMO DATA 10/01/2026']))]
    # Separate date per document, matching the app's input format.
    files[2]=file('notes','nota11.pdf',pdf_bytes(['AZIENDA DEMO DATA 11/01/2026','020102 800101 50,00 CD 123 PAG A MZ POS 1']))
    files.append(file('notes','nota10.pdf',pdf_bytes(['AZIENDA DEMO DATA 10/01/2026','800102 010101 30,00 STORNO POS DA INC\\C\\CORR 2','020102 800102 30,00 PRIVATI PAG A MZ POS 3'])))
    fixture=folder/'fixture.json';fixture.write_text(json.dumps({'files':files}),encoding='utf-8');report=folder/'report.json'
    run=subprocess.run([str(Path('dist/RiconciliaPOS/RiconciliaPOS.exe').resolve()),'--smoke-test',str(fixture),str(report)],timeout=150)
    assert run.returncode==0,run.returncode
    result=json.loads(report.read_text(encoding='utf-8'));assert result['ok'],result
    print(json.dumps(result,indent=2))
