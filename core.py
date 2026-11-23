import csv, io, re, hashlib
from datetime import datetime, date, timedelta
from decimal import Decimal, InvalidOperation
from collections import defaultdict
from pypdf import PdfReader
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

class InputError(ValueError): pass

def money(value):
    if isinstance(value,(int,float,Decimal)): d=Decimal(str(value))
    else:
        s=str(value).strip().replace('€','').replace('EUR','').replace('\u00a0','').replace(' ','')
        if ',' in s: s=s.replace('.','').replace(',','.')
        elif re.fullmatch(r'-?\d{1,3}(\.\d{3})+',s): s=s.replace('.','')
        try:d=Decimal(s)
        except InvalidOperation:raise InputError(f'Importo non leggibile: {value}')
    if not d.is_finite() or d*100 != (d*100).to_integral_value():raise InputError(f'Importo non valido al centesimo: {value}')
    return int(d*100)

def day(value):
    if isinstance(value,(date,datetime)): return value.strftime('%Y-%m-%d')
    s=str(value).strip().split(' ')[0].split('T')[0]
    for fmt in ['%d/%m/%Y','%d-%m-%Y','%Y-%m-%d','%d/%m/%y']:
        try:return datetime.strptime(s,fmt).date().isoformat()
        except ValueError:pass
    raise InputError(f'Data non leggibile: {value}')

def nextday(s):return (date.fromisoformat(s)+timedelta(days=1)).isoformat()
def norm(s):return re.sub(r'\s+',' ',str(s or '').strip()).lower()
def pdf(data):
    try:
        r=PdfReader(io.BytesIO(data)); texts=[p.extract_text() or '' for p in r.pages]
    except Exception as e: raise InputError('PDF non leggibile o protetto.') from e
    if not any(t.strip() for t in texts):raise InputError('PDF senza testo: questa versione richiede PDF testuali, non scansioni.')
    return texts

def parse_pos(name,data):
    ext=name.lower().rsplit('.',1)[-1]; tables=[]
    if ext=='csv':
        try:s=data.decode('utf-8-sig')
        except UnicodeDecodeError:s=data.decode('cp1252')
        delimiter=';' if s.count(';')>=s.count(',') else ','
        tables=[('',list(csv.reader(io.StringIO(s),delimiter=delimiter)))]
    elif ext=='xlsx':
        wb=load_workbook(io.BytesIO(data),read_only=True,data_only=True)
        tables=[(ws.title,list(ws.values)) for ws in wb];wb.close()
    elif ext=='pdf':
        pdf(data)
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as doc:
            for i,page in enumerate(doc.pages):
                ts=page.extract_tables() or page.extract_tables({'vertical_strategy':'text','horizontal_strategy':'text'})
                tables.extend((f'pagina {i+1}',t) for t in ts)
    else:raise InputError('Formato POS supportato: CSV, XLSX o PDF testuale.')
    result=[]; recognized=False
    for section,table in tables:
        cols=None
        for lineno,row in enumerate(table,1):
            names=[norm(x) for x in row]
            if 'importo' in names and ('data' in names or 'data transazione' in names):
                cols={k:names.index(k) for k in names if k in ['importo','data','data transazione','stato']};recognized=True;continue
            if cols is None or not any(x is not None and str(x).strip() for x in row):continue
            def get(k):return row[cols[k]] if k in cols and cols[k]<len(row) else None
            if norm(get('importo')) in ['totale','importo']:continue
            try:
                a=money(get('importo'));d=day(get('data') or get('data transazione'))
            except (InputError,TypeError) as e:raise InputError(f'{name}, {section} riga {lineno}: {e}') from e
            status=norm(get('stato'))
            result.append({'date':d,'amount':a,'state':status,'file':name,'ref':f'{section} riga {lineno}'.strip(),'time':str(get('data') or get('data transazione')),'match':None})
    if not recognized or not result:raise InputError(f'{name}: tabella POS non riconosciuta. Servono intestazioni Data e Importo; Stato è consigliato.')
    return result

def parse_notes(name,data):
    if not name.lower().endswith('.pdf'):raise InputError('Per la prima nota carica PDF testuali nel formato degli esempi.')
    texts=pdf(data); dates=set(); rows=[]; stornos=[]
    current=None
    for p,t in enumerate(texts,1):
        m=re.search(r'\bDATA\s+(\d{2}/\d{2}/\d{4})',t)
        if m:current=day(m[1]);dates.add(current)
        if not current:raise InputError(f'{name}: data della prima nota non riconosciuta.')
        for line in t.splitlines():
            line=line.strip(); nums=re.findall(r'(?<!\d)\d[\d.]*,\d{2}',line)
            if not nums:continue
            base={'date':current,'file':name,'ref':f'pagina {p}','description':line,'amount':money(nums[-1]),'used':False,'group':False,'code':'','customer':''}
            if re.match(r'^\d+\s+\d+',line) and 'STORNO POS DA INC\\C\\CORR' in line:
                stornos.append(base);continue
            if not re.match(r'^\d+\s+[\d-]+\s',line):continue
            if 'PAG A MZ POS' not in line and 'RIMB\\CRED\\ A MZ POS' not in line:continue
            cm=re.search(r'\bCD\s+(\d+)\b',line)
            if cm:base['code']=str(int(cm[1]))
            tail=re.split(r'\d[\d.]*,\d{2}',line)[-1].strip()
            base['description']=tail
            base['customer']=re.split(r'PAG A MZ POS|RIMB\\CRED\\',tail)[0].strip() if not cm else ''
            rows.append(base)
    warnings=[]
    for s in stornos:
        linked=[r for r in rows if r['date']==s['date'] and r['amount']==s['amount'] and not r['code'] and re.search(r'PRIVAT',r['description'],re.I) and not r['group']]
        if len(linked)==1:
            linked[0]['group']=True;linked[0]['storno']=s['description']
        elif not linked:
            s['group']=True;s['customer']='Corrispettivi POS';s['storno']=s['description'];rows.append(s)
        else:warnings.append(f'{name}: più righe possibili per lo storno da {s["amount"]/100:.2f}; verificare il documento.')
    if not rows: warnings.append(f'{name}: nessuna riga POS selezionabile.')
    return rows,dates,warnings

def parse_clients(name,data):
    clients=defaultdict(set)
    for page,text in enumerate(pdf(data),1):
        for line in text.splitlines():
            m=re.search(r'\b04\s+01\s+(\d{4,})\s+(.+)',line)
            if not m:continue
            label=re.split(r'\s{2,}',m[2].strip())[0].strip()
            if label:clients[str(int(m[1]))].add(label)
    if not clients:raise InputError(f'{name}: codici clienti non riconosciuti nel PDF.')
    return clients

def subset_unique(items,target,limit=100000):
    # Find at most two index-distinct combinations. Never guess when ambiguous.
    solutions=[];visits=0
    items=[(i,r) for i,r in items if 0<r['amount']<=target]
    items.sort(key=lambda x:x[1]['amount'],reverse=True)
    suffix=[0]*(len(items)+1)
    for j in range(len(items)-1,-1,-1):suffix[j]=suffix[j+1]+items[j][1]['amount']
    def go(k,left,chosen):
        nonlocal visits
        visits+=1
        if visits>limit:raise OverflowError
        if left==0:
            if len(chosen)>1:solutions.append(chosen)
            return
        if k==len(items) or left<0 or suffix[k]<left or len(solutions)>1:return
        i,r=items[k]
        if r['amount']<=left:go(k+1,left-r['amount'],chosen+[i])
        go(k+1,left,chosen)
    try:go(0,target,[])
    except (OverflowError,RecursionError):return None,'Ricerca combinazioni troppo ampia: verifica manuale'
    if len(solutions)==1:return solutions[0],None
    return None,'Più somme possibili: verifica manuale' if solutions else None

def analyze(files):
    ps=[];ns=[];dates=set();warnings=[];clients=defaultdict(set);inventory=[];seen=set()
    for f in files:
        name=f['name'];data=f['data'];role=f['role'];digest=hashlib.sha256(data).hexdigest()
        if (role,digest) in seen:warnings.append(f'{name}: copia identica ignorata.');continue
        seen.add((role,digest));inventory.append({'name':name,'role':role})
        if role=='pos':ps.extend(parse_pos(name,data))
        elif role=='notes':
            rr,dd,ww=parse_notes(name,data);ns.extend(rr);dates.update(dd);warnings.extend(ww)
        elif role=='clients':
            for c,n in parse_clients(name,data).items():clients[c].update(n)
    if not ps or not dates:raise InputError('Carica almeno un file POS e una prima nota.')
    excluded=sum(p['state']=='negata' for p in ps)
    pending=[p for p in ps if p['state'] not in ('negata','contabilizzata')]
    valid=[p for p in ps if p['state']=='contabilizzata']
    if pending:warnings.append(f'{len(pending)} operazioni con stato diverso da Contabilizzata/Negata: da verificare, escluse dai totali riconciliabili.')
    for n in ns:
        if n['code']:
            found=clients.get(n['code'],set())
            n['customer']=next(iter(found)) if len(found)==1 else 'Cliente da verificare'
    valid.sort(key=lambda p:(p['date'],p['time'],p['file'],p['ref']))
    # Reserve exact matches globally, then process groups using only unused POS.
    for off in (0,1):
        for p in valid:
            if p['match'] is not None:continue
            for j,n in enumerate(ns):
                if not n['used'] and n['date']==(p['date'] if off==0 else nextday(p['date'])) and n['amount']==p['amount']:
                    n['used']=True;p['match']=j;p['shift']=off;break
        for j,n in enumerate(ns):
            if n['used'] or not n['group'] or n.get('issue'):continue
            d=n['date'] if off==0 else (date.fromisoformat(n['date'])-timedelta(days=1)).isoformat()
            candidates=[(i,p) for i,p in enumerate(valid) if p['date']==d and p['match'] is None]
            chosen,err=subset_unique(candidates,n['amount'])
            if chosen:
                n['used']=True
                for i in chosen:valid[i]['match']=j;valid[i]['shift']=off;valid[i]['sum']=True
            elif err:n['issue']=err
    output=[]
    for i,p in enumerate(valid):
        n=ns[p['match']] if p['match'] is not None else None
        if n:status='shift' if p.get('shift') else ('sum' if p.get('sum') else 'ok')
        else:status='missing' if p['date'] in dates and nextday(p['date']) in dates else 'incomplete'
        missing=[d for d in [p['date'],nextday(p['date'])] if d not in dates]
        output.append({'id':i,'date':p['date'],'note_date':n['date'] if n else '', 'code':n['code'] if n else '', 'customer':n['customer'] if n else '', 'amount':p['amount'],'allocated':p['amount'] if n else None,'note_total':n['amount'] if n else None,'status':status,'group':f'G{p["match"]+1}' if p.get('sum') else '', 'pos_file':p['file'],'pos_ref':p['ref'],'time':p['time'],'note_file':n['file'] if n else '', 'note_ref':n['ref'] if n else '', 'description':n['description'] if n else '', 'detail':('Somma di transazioni POS. ' if p.get('sum') else '')+('Registrazione il giorno successivo.' if p.get('shift') else '') if n else ('Manca prima nota: '+', '.join(missing) if missing else 'Nessun importo corrispondente nelle due date.'),'client_check':bool(n and n['code'] and n['customer']=='Cliente da verificare')})
    for p in pending:
        output.append({'id':len(output),'date':p['date'],'note_date':'','code':'','customer':'','amount':p['amount'],'allocated':None,'note_total':None,'status':'state','group':'','pos_file':p['file'],'pos_ref':p['ref'],'time':p['time'],'note_file':'','note_ref':'','description':'','detail':f'Stato POS: {p["state"] or "non indicato"}. Verifica necessaria.','client_check':False})
    daily=[]
    for d in sorted({r['date'] for r in output}):
        rr=[r for r in output if r['date']==d and r['status']!='state'];matched=[r for r in rr if r['allocated'] is not None]
        daily.append({'date':d,'count':len(rr),'pos':sum(r['amount'] for r in rr),'matched':sum(r['amount'] for r in matched),'difference':sum(r['amount'] for r in rr if r['allocated'] is None),'shift':sum(r['amount'] for r in matched if r['status']=='shift'),'unmatched':len(rr)-len(matched)})
    unmatched_notes=[{k:v for k,v in n.items() if k!='used'} for n in ns if not n['used']]
    counts=CounterSafe(output)
    return {'rows':output,'daily':daily,'warnings':warnings,'unmatched_notes':unmatched_notes,'files':inventory,'note_dates':sorted(dates),'excluded':excluded,'counts':counts,'total':sum(p['amount'] for p in valid),'matched_total':sum(r['allocated'] or 0 for r in output)}

def CounterSafe(rows):
    counts=defaultdict(int)
    for r in rows:counts[r['status']]+=1
    counts['matched']=sum(counts[k] for k in ['ok','sum','shift']);counts['review']=sum(counts[k] for k in ['missing','incomplete','state']);counts['clients']=sum(r['client_check'] for r in rows)
    return dict(counts)

LABELS={'ok':'OK','shift':'Slittamento +1 giorno','sum':'Corrispondenza per somma','missing':'Senza corrispondenza','incomplete':'Verifica incompleta','state':'Stato POS da verificare'}
def export_excel(result):
    wb=Workbook();ws=wb.active;ws.title='Dettaglio'
    ws.append(['Data POS','Data prima nota','Codice cliente','Cliente','Importo POS','Quota riconciliata','Differenza','Esito','Gruppo somma','Totale gruppo prima nota (non sommare)','File POS','Riferimento POS','File prima nota','Riferimento prima nota','Descrizione','Note verifica','Annotazioni manuali'])
    for idx,r in enumerate(result['rows'],2):
        ws.append([date.fromisoformat(r['date']),date.fromisoformat(r['note_date']) if r['note_date'] else None,r['code'],r['customer'],r['amount']/100,None if r['allocated'] is None else r['allocated']/100,f'=E{idx}-F{idx}',LABELS[r['status']],r['group'],None if r['note_total'] is None else r['note_total']/100,r['pos_file'],r['pos_ref'],r['note_file'],r['note_ref'],r['description'],r['detail']+(' Cliente da verificare.' if r['client_check'] else ''),''])
        for c in [1,2]:ws.cell(idx,c).number_format='dd/mm/yyyy'
        for c in [5,6,7,10]:ws.cell(idx,c).number_format='#,##0.00'
        color={'missing':'FDE8E8','incomplete':'FFF1D6','state':'FFF1D6','shift':'E2EEFF','sum':'DCF4EA','ok':'DCF4EA'}[r['status']]
        ws.cell(idx,8).fill=PatternFill('solid',fgColor=color)
    summary=wb.create_sheet('Riepilogo giornaliero');summary.append(['Data POS','Operazioni valide','Totale POS','Totale riconciliato per data POS','Importo non riconciliato','Di cui slittamenti','Operazioni non riconciliate'])
    for d in result['daily']:summary.append([date.fromisoformat(d['date']),d['count'],d['pos']/100,d['matched']/100,d['difference']/100,d['shift']/100,d['unmatched']])
    for rr in summary.iter_rows(min_row=2):
        rr[0].number_format='dd/mm/yyyy'
        for c in rr[2:6]:c.number_format='#,##0.00'
    review=wb.create_sheet('Da verificare');review.append(['Tipo','Data','Importo','File','Riferimento','Descrizione'])
    for r in result['rows']:
        if r['status'] in ['missing','incomplete','state'] or r['client_check']:review.append([LABELS[r['status']] if not r['client_check'] else 'Cliente da verificare',r['date'],r['amount']/100,r['pos_file'],r['pos_ref'],r['detail']])
    for n in result['unmatched_notes']:review.append(['Prima nota non associata',n['date'],n['amount']/100,n['file'],n['ref'],n.get('issue',n['description'])])
    info=wb.create_sheet('Criteri');info.append(['Voce','Descrizione'])
    for rr in [('Attribuzione','Totali per data POS, inclusi gli slittamenti.'),('Somme','Quota riconciliata ripartita sulle singole transazioni; il totale gruppo è informativo e non va sommato.'),('Negate ignorate',str(result['excluded'])),('Stati da verificare','Preautorizzate e stati mancanti/sconosciuti esclusi dai totali riconciliabili.'),('Modifiche manuali','Riepilogo ed esiti fotografano il confronto esportato. Per ricalcolarli, rieseguire il confronto nell’app. Usare Annotazioni manuali per commenti.'),('Copertura','Prime note: '+', '.join(result['note_dates']))]:info.append(rr)
    for w in result['warnings']:info.append(['Avviso',w])
    for sheet in wb:
        sheet.freeze_panes='A2';sheet.auto_filter.ref=sheet.dimensions
        for cell in sheet[1]:cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='183D39');cell.alignment=Alignment(wrap_text=True)
        sheet.row_dimensions[1].height=42
        for col in sheet.columns:
            letter=col[0].column_letter;sheet.column_dimensions[letter].width=min(65,max(16,max(len(str(c.value or '')) for c in col)+2))
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                # Imported text must never become an Excel formula.
                if isinstance(cell.value,str) and cell.data_type=='f' and not (sheet.title=='Dettaglio' and cell.column==7):cell.data_type='s'
                cell.alignment=Alignment(vertical='top',wrap_text=True)
    out=io.BytesIO();wb.save(out);return out.getvalue()
