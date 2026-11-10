"""Herramientas locales, sin subida de documentos a servicios externos."""
from pathlib import Path
import os,sys,subprocess,tempfile,json
from collections import Counter

QUALITY_LIBS=Path(r'C:\Users\clopez\Documents\ChatGPT\LOCAL CLC\tmp\font_quality\libs')
VISUAL_LIBS=Path(r'C:\Users\clopez\Documents\ChatGPT\LOCAL CLC\tmp\santander_v2\libs')

def font_quality(root=None):
    root=Path(root or Path(__file__).parent)
    files=[root/name for name in ['CLC_Movimientos_Select_Regular.ttf','CLC_Select_Escaneo_Editable.ttf']]
    env=os.environ.copy();env['PYTHONPATH']=str(QUALITY_LIBS);env['PYTHONIOENCODING']='utf-8'
    with tempfile.TemporaryDirectory(prefix='select_font_quality_') as folder:
        reports=[];logs=[];codes=[];counts=Counter()
        # Son familias independientes: no exigir nombres de familia iguales.
        for index,file in enumerate(files):
            report=Path(folder)/f'informe_{index}.json'
            command=[sys.executable,'-m','fontbakery','check-universal',str(file),'--skip-network','--no-progress','--no-colors','--json',str(report)]
            result=subprocess.run(command,env=env,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=180)
            if not report.exists():raise RuntimeError((result.stderr or result.stdout)[-2000:])
            raw=json.loads(report.read_text(encoding='utf-8'));reports.append({'font':file.name,'report':raw});counts.update(raw.get('result',{}));logs.append(file.name+'\n'+result.stdout+'\n'+result.stderr);codes.append(result.returncode)
        return {'exit_code':max(codes),'log':'\n'.join(logs),'report':json.dumps({'result':dict(counts),'fonts':reports},ensure_ascii=False,indent=2).encode('utf-8')}

def compare_visual(reference,candidate,page_reference,page_candidate,area):
    env=os.environ.copy();env['PYTHONPATH']=str(VISUAL_LIBS);env['PYTHONIOENCODING']='utf-8'
    with tempfile.TemporaryDirectory(prefix='select_compare_') as folder:
        folder=Path(folder);(folder/'reference.pdf').write_bytes(reference);(folder/'candidate.pdf').write_bytes(candidate)
        result=subprocess.run([sys.executable,str(Path(__file__).with_name('select_compare_worker.py')),str(folder),str(page_reference),str(page_candidate),area],env=env,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=60)
        if result.returncode:raise RuntimeError(result.stderr[-2000:])
        return {'score':json.loads((folder/'metrics.json').read_text())['ssim'],'reference':(folder/'reference.png').read_bytes(),'candidate':(folder/'candidate.png').read_bytes(),'difference':(folder/'difference.png').read_bytes()}
