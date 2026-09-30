from __future__ import annotations
import json,os,socket,sys,threading,time,urllib.request,webbrowser
from pathlib import Path
import uvicorn
ROOT=Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:sys.path.insert(0,str(ROOT))
CFG=ROOT/"config"/"app.json";RUNTIME=ROOT/"runtime"
def cfg():
    if CFG.exists():return json.loads(CFG.read_text(encoding="utf-8"))
    return {"host":"127.0.0.1","preferred_port":8787,"open_browser":True}
def port(host,p):
    for x in range(p,p+30):
        with socket.socket() as s:
            try:s.bind((host,x));return x
            except OSError:pass
    raise RuntimeError("Не знайдено вільний порт.")
def open_ready(url):
    for _ in range(100):
        try:
            if urllib.request.urlopen(url+"/api/health",timeout=.5).status==200:
                webbrowser.open(url);return
        except:time.sleep(.25)
    webbrowser.open(url)
def main():
    c=cfg();host=c.get("host","127.0.0.1");p=port(host,int(c.get("preferred_port",8787)));RUNTIME.mkdir(parents=True,exist_ok=True)
    (RUNTIME/"active_port.txt").write_text(str(p),encoding="utf-8");url=f"http://{host}:{p}"
    print(f"\nDoctor Eleks Automation: {url}\nLauncher PID: {os.getpid()}\nДля завершення натисніть Ctrl+C.\n",flush=True)
    if c.get("open_browser",True):threading.Thread(target=open_ready,args=(url,),daemon=True).start()
    reload_enabled="--reload" in sys.argv
    if reload_enabled:
        print("Автоперезапуск увімкнено: після зміни Python-коду в консолі з'являться APP STOPPED і APP STARTED.",flush=True)
    uvicorn.run("app.main:app",host=host,port=p,reload=reload_enabled,
                reload_dirs=[str(ROOT/"app"),str(ROOT/"scripts")] if reload_enabled else None,
                access_log=False,log_level="info")
if __name__=="__main__":main()
