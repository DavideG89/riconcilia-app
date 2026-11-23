import argparse,base64,json,secrets,threading,webbrowser
from pathlib import Path
from http.server import ThreadingHTTPServer,BaseHTTPRequestHandler
from urllib.parse import urlparse
from core import analyze,export_excel,InputError
ROOT=Path(__file__).resolve().parent
TOKEN=secrets.token_urlsafe(24)
RESULTS={}
LOCK=threading.Lock()
class Handler(BaseHTTPRequestHandler):
    def send(self,status,data,mime='application/json'):
        if isinstance(data,dict):data=json.dumps(data,ensure_ascii=False).encode()
        self.send_response(status);self.send_header('Content-Type',mime);self.send_header('Content-Length',str(len(data)));self.send_header('Cache-Control','no-store');self.send_header('X-Content-Type-Options','nosniff');self.end_headers();self.wfile.write(data)
    def do_GET(self):
        path=urlparse(self.path).path
        if path=='/api/config':return self.send(200,{'token':TOKEN})
        if path=='/':path='/index.html'
        allowed={'/index.html':'text/html; charset=utf-8','/app.js':'text/javascript; charset=utf-8','/style.css':'text/css; charset=utf-8'}
        if path not in allowed:return self.send(404,{'error':'Non trovato'})
        self.send(200,(ROOT/'static'/path[1:]).read_bytes(),allowed[path])
    def do_POST(self):
        if self.headers.get('X-Local-Token')!=TOKEN:return self.send(403,{'error':'Sessione non valida. Ricarica la pagina.'})
        size=int(self.headers.get('Content-Length','0'))
        if size>50*1024*1024:return self.send(413,{'error':'Limite di caricamento: 35 MB complessivi.'})
        try:
            data=json.loads(self.rfile.read(size));path=urlparse(self.path).path
            if path=='/api/analyze':
                files=[]
                for f in data.get('files',[]):
                    if f['role'] not in ['pos','notes','clients']:raise InputError('Tipo documento non valido.')
                    files.append({'name':Path(f['name']).name,'role':f['role'],'data':base64.b64decode(f['content'],validate=True)})
                with LOCK:
                    result=analyze(files);key=secrets.token_urlsafe(16)
                    if len(RESULTS)>=8:RESULTS.pop(next(iter(RESULTS)))
                    RESULTS[key]=result
                return self.send(200,dict(result,session=key))
            if path=='/api/demo':
                parent=ROOT.parent
                names={'pos':sorted(parent.glob('*Transazioni*.csv')), 'notes':sorted(set(parent.glob('prima nota*.pdf')) | set(parent.glob('pn-*.pdf'))), 'clients':sorted(parent.glob('codici CLIENTI*.pdf'))}
                files=[{'name':path.name,'role':role,'data':path.read_bytes()} for role,paths in names.items() for path in paths if path.is_file()]
                with LOCK:
                    result=analyze(files);key=secrets.token_urlsafe(16)
                    if len(RESULTS)>=8:RESULTS.pop(next(iter(RESULTS)))
                    RESULTS[key]=result
                return self.send(200,dict(result,session=key))
            if path=='/api/export':
                result=RESULTS.get(data.get('session'))
                if result is None:raise InputError('Risultato non più disponibile. Ripeti il confronto.')
                return self.send(200,export_excel(result),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            return self.send(404,{'error':'Non trovato'})
        except (InputError,ValueError,KeyError) as e:self.send(400,{'error':str(e)})
        except Exception as e:
            print('Errore:',type(e).__name__,str(e),flush=True)
            self.send(500,{'error':'Elaborazione non riuscita. Verifica il formato dei documenti e riprova.'})
    def log_message(self,*args):pass
if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--port',type=int,default=8765);ap.add_argument('--open',action='store_true');args=ap.parse_args()
    server=ThreadingHTTPServer(('127.0.0.1',args.port),Handler)
    url=f'http://127.0.0.1:{args.port}';print('Riconcilia POS: '+url,flush=True)
    if args.open:threading.Timer(.7,lambda:webbrowser.open(url)).start()
    try:server.serve_forever()
    except KeyboardInterrupt:server.server_close()
