"""Desktop launcher for the portable Windows build."""
import argparse,base64,io,json,os,sys,threading,urllib.request,webbrowser,zipfile
from pathlib import Path
from http.server import ThreadingHTTPServer
from server import Handler,TOKEN

def start_server():
    server=ThreadingHTTPServer(('127.0.0.1',0),Handler)
    threading.Thread(target=server.serve_forever,daemon=True).start()
    return server,f'http://127.0.0.1:{server.server_address[1]}'

def smoke_test(fixture_path,report_path):
    server,url=start_server()
    try:
        with urllib.request.urlopen(url,timeout=20) as r:assert b'Avvia confronto' in r.read()
        for asset in ['style.css','app.js']:
            with urllib.request.urlopen(url+'/'+asset,timeout=20) as r:assert len(r.read())>100
        fixture=json.loads(Path(fixture_path).read_text(encoding='utf-8'))
        def post(path,body):
            req=urllib.request.Request(url+'/api/'+path,data=json.dumps(body).encode(),headers={'Content-Type':'application/json','X-Local-Token':TOKEN})
            return urllib.request.urlopen(req,timeout=90).read()
        result=json.loads(post('analyze',fixture))
        assert result['counts']['matched']==3,result['counts']
        assert result['counts']['shift']==1,result['counts']
        assert result['counts']['sum']==2,result['counts']
        assert result['excluded']==1
        data=post('export',{'session':result['session']})
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            assert z.testzip() is None
            assert 'xl/workbook.xml' in z.namelist()
        Path(report_path).write_text(json.dumps({'ok':True,'counts':result['counts'],'excel_bytes':len(data)}),encoding='utf-8')
    finally:server.shutdown();server.server_close()

def main():
    if sys.stdout is None:sys.stdout=open(os.devnull,'w')
    if sys.stderr is None:sys.stderr=open(os.devnull,'w')
    parser=argparse.ArgumentParser();parser.add_argument('--smoke-test',nargs=2,metavar=('FIXTURE','REPORT'));args=parser.parse_args()
    if args.smoke_test:
        try:smoke_test(*args.smoke_test)
        except Exception as e:
            Path(args.smoke_test[1]).write_text(json.dumps({'ok':False,'error':str(e)}),encoding='utf-8');raise
        return
    import tkinter as tk
    from tkinter import messagebox
    root=tk.Tk();root.title('Riconcilia POS');root.geometry('460x260');root.resizable(False,False);root.configure(bg='#f4f6f5')
    try:server,url=start_server()
    except OSError as e:
        messagebox.showerror('Avvio non riuscito',f'Impossibile avviare il servizio locale: {e}');root.destroy();return
    tk.Label(root,text='Riconcilia POS',font=('Segoe UI',22,'bold'),bg='#f4f6f5',fg='#153c36').pack(pady=(28,10))
    tk.Label(root,text='Il programma è attivo sul tuo computer.\nLascia aperta questa finestra mentre lavori.',font=('Segoe UI',11),bg='#f4f6f5',fg='#3d5c50').pack()
    tk.Button(root,text='Apri nel browser',command=lambda:webbrowser.open(url),font=('Segoe UI',11),padx=22,pady=8).pack(pady=18)
    def close():server.shutdown();server.server_close();root.destroy()
    tk.Button(root,text='Chiudi il programma',command=close,font=('Segoe UI',10)).pack()
    root.protocol('WM_DELETE_WINDOW',close);root.after(500,lambda:webbrowser.open(url));root.mainloop()
if __name__=='__main__':main()
