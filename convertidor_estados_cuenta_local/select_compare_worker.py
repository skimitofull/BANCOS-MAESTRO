from pathlib import Path
import sys,json,fitz,numpy as np
from PIL import Image
from skimage.metrics import structural_similarity
folder=Path(sys.argv[1]);indices=[int(sys.argv[2])-1,int(sys.argv[3])-1]
areas={'Resumen':fitz.Rect(30,180,582,492),'Movimientos':fitz.Rect(30,132,582,735),'Pagina sin logo ni pie':fitz.Rect(30,95,582,735)}
clip=areas[sys.argv[4]];arrays=[]
for index,name in zip(indices,['reference','candidate']):
    doc=fitz.open(folder/f'{name}.pdf')
    if index<0 or index>=len(doc):raise ValueError('Numero de hoja fuera del PDF')
    if abs(doc[index].rect.width-612)>2 or abs(doc[index].rect.height-792)>2:raise ValueError('Compara dos hojas tamaño carta; no el PDF de comparación lado a lado.')
    pix=doc[index].get_pixmap(matrix=fitz.Matrix(2,2),clip=clip,colorspace=fitz.csGRAY)
    array=np.frombuffer(pix.samples,np.uint8).reshape(pix.height,pix.width).copy();arrays.append(array);Image.fromarray(array).save(folder/f'{name}.png')
score=float(structural_similarity(*arrays,data_range=255));diff=np.abs(arrays[0].astype(float)-arrays[1].astype(float)).astype(np.uint8)
red=np.full_like(diff,255);image=np.stack([red,255-diff,255-diff],axis=2);Image.fromarray(image).save(folder/'difference.png')
(folder/'metrics.json').write_text(json.dumps({'ssim':score,'note':'Incluye diferencias de datos y posición; no demuestra identidad de fuentes.'}),encoding='utf-8')
