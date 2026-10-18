#!/usr/bin/env python3
"""
Genera il grafico HTML "PUN GME - profilo orario giornaliero" (file unico e autonomo).

Uso:
    pip install pandas openpyxl
    python build_pun_dashboard.py CARTELLA_CON_XLSX [-o PUN_profilo_orario.html]

Input : file .xlsx GME con colonne Data (gg/mm/aaaa), Ora, Periodo, PUN in EUR/MWh (dati quartorari).
Output: un solo file .html (usa Chart.js e SheetJS da CDN), apribile in qualsiasi browser
        o caricabile su qualsiasi hosting statico.
Nota  : nella pagina il pulsante "Carica serie PUN" aggiunge altri file senza rigenerare.
"""
import argparse, glob, json, os
import numpy as np, pandas as pd

MESI = ['gen','feb','mar','apr','mag','giu','lug','ago','set','ott','nov','dic']
COL = {10: '#94c4df', 11: '#64a9d3', 12: '#3e8ec4', 1: '#1f6eb3', 2: '#084f99', 3: '#08306b', 4: '#8fbc8f', 5: '#2e8b57', 6: '#feba55', 7: '#fd7435', 8: '#e2191c', 9: '#970026'}
FASE = lambda m: 'inv' if m in (10,11,12,1,2,3) else 'est' if m in (6,7,8,9) else 'mid'

def leggi(path):
    d = pd.read_excel(path).iloc[:, :4]
    d.columns = ['Data', 'Ora', 'Periodo', 'P']
    if not pd.api.types.is_numeric_dtype(d.P):
        d['P'] = d.P.astype(str).str.replace(',', '.').astype(float)
    d['Ora'] = pd.to_numeric(d.Ora, errors='coerce')
    d['dt'] = pd.to_datetime(d.Data, format='%d/%m/%Y', errors='coerce')
    return d.dropna(subset=['dt', 'Ora', 'P'])

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('cartella')
    ap.add_argument('-o', '--output', default='PUN_profilo_orario.html')
    a = ap.parse_args()
    files = sorted(glob.glob(os.path.join(a.cartella, '*.xlsx')))
    if not files:
        raise SystemExit('Nessun file .xlsx trovato in ' + a.cartella)
    d = pd.concat([leggi(f) for f in files])
    d = d[d.Ora <= 24]  # ignora la 25a ora del cambio ora solare
    h = d.groupby(['dt', 'Ora']).P.mean().unstack().reindex(columns=range(1, 25))
    out = []
    for per, s in h.groupby(h.index.to_period('M')):
        righe = [[None if np.isnan(v) else round(float(v), 1) for v in r] for r in s.values]
        ultimo = s.index.max().day
        etichetta = f'{MESI[per.month-1]} {per.year}' + (f' (1-{ultimo})' if ultimo < per.days_in_month else '')
        out.append(dict(k=str(per), l=etichetta, ph=FASE(per.month), c=COL[per.month], d=righe))
    html = TEMPLATE.replace('__DATA__', json.dumps(out, separators=(',', ':')))
    with open(a.output, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'{len(out)} mesi, {len(h)} giorni -> {a.output}')

TEMPLATE = r'''<!DOCTYPE html>
<html lang="it"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>PUN GME – profilo orario giornaliero</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{--bg:#fff;--fg:#1f2933;--mut:#6b7785;--grid:#e3e7ec;--card:#f4f6f8;--bd:#d3d9e0;--acc:#1f2933;--med:#7c3aed;
box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#14181d;--fg:#e8ecf0;--mut:#98a3af;--grid:#2a3139;--card:#1d232a;--bd:#38414b;--acc:#f1f5f9;--med:#a78bfa}}
:root[data-theme="dark"]{--bg:#14181d;--fg:#e8ecf0;--mut:#98a3af;--grid:#2a3139;--card:#1d232a;--bd:#38414b;--acc:#f1f5f9;--med:#a78bfa}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.4 Arial,Helvetica,sans-serif;padding:16px}
h1{font-size:18px;margin:0 0 2px}p.sub{margin:0 0 12px;color:var(--mut)}
.row{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:8px 0}
.lab{color:var(--mut);font-size:12px;margin-right:4px}
button,label.chk{font:inherit;border:1px solid var(--bd);background:var(--card);color:var(--fg);border-radius:16px;padding:5px 11px;cursor:pointer}
button.on{background:var(--fg);color:var(--bg);border-color:var(--fg)}
.chip{display:inline-flex;align-items:center;gap:6px}.chip i{width:10px;height:10px;border-radius:50%;display:inline-block}
.chip:not(.on){opacity:.45}
label.chk{display:inline-flex;gap:6px;align-items:center}
.box{position:relative;height:68vh;min-height:420px;margin-top:8px}
#st{margin-top:10px;color:var(--mut);font-size:13px}
</style></head><body>
<h1>PUN GME – profilo orario giornaliero</h1>
<p class="sub">Prezzo medio orario (media dei 4 quarti d'ora), €/MWh. Ogni curva è un giorno.</p>
<div class="row"><span class="lab">Vista</span><button id="mM" class="on">Mese per mese</button><button id="mA">Aggregato (media e mediana)</button>
<label class="chk"><input type="checkbox" id="cd" checked> Curve giornaliere</label></div>
<div class="row"><span class="lab">Mesi</span><button id="all">Tutti</button><button id="none">Nessuno</button><button id="inv">Inverno</button><button id="est">Estate</button></div>
<div class="row" id="chips"></div>
<div class="row"><span class="lab">Dati</span><button id="up">Carica serie PUN</button><button id="del" disabled>Elimina mesi selezionati</button><button id="rs" disabled>Ripristina originali</button><input type="file" id="fi" accept=".xlsx,.xls,.csv" multiple hidden><span class="lab" id="msg"></span></div>
<div class="box"><canvas id="c"></canvas></div>
<div id="st"></div>
<script>
const D=__DATA__;let sel=new Set(),mode='m',days=true;
const $=id=>document.getElementById(id);
const rgba=(h,a)=>{const n=parseInt(h.slice(1),16);return `rgba(${n>>16},${n>>8&255},${n&255},${a})`};
const css=n=>getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const med=a=>{const s=[...a].sort((x,y)=>x-y),n=s.length;return n?(n%2?s[(n-1)/2]:(s[n/2-1]+s[n/2])/2):null};
const qt=(a,p)=>{if(!a.length)return null;const s=[...a].sort((x,y)=>x-y),k=(s.length-1)*p,f=Math.floor(k);return s[f]+(s[Math.min(f+1,s.length-1)]-s[f])*(k-f)};
const avg=a=>a.length?a.reduce((x,y)=>x+y,0)/a.length:null;
const prof=(rows,f)=>Array.from({length:24},(_,i)=>f(rows.map(r=>r[i]).filter(v=>v!=null)));
const pts=p=>p.map((v,i)=>({x:i+1,y:v}));
const MN=['gen','feb','mar','apr','mag','giu','lug','ago','set','ott','nov','dic'];
const CM={};D.forEach(m=>CM[+m.k.slice(5)]={c:m.c,ph:m.ph});
const ld=k=>{try{return JSON.parse(localStorage.getItem(k)||'null')}catch(e){return null}};
let U=ld('pun_up_v1')||{},R=new Set(ld('pun_rm_v1')||[]);
const save=()=>{try{Object.keys(U).length?localStorage.setItem('pun_up_v1',JSON.stringify(U)):localStorage.removeItem('pun_up_v1');R.size?localStorage.setItem('pun_rm_v1',JSON.stringify([...R])):localStorage.removeItem('pun_rm_v1')}catch(e){}};
const M=()=>{const o={};D.forEach(m=>{if(!R.has(m.k))o[m.k]=m});Object.values(U).forEach(m=>o[m.k]=m);return Object.values(o).sort((a,b)=>a.k<b.k?-1:1)};
sel=new Set(M().map(m=>m.k));
function chips(){$('chips').innerHTML='';M().forEach(m=>{const b=document.createElement('button');b.className='chip on';b.dataset.k=m.k;
b.innerHTML=`<i style="background:${m.c}"></i>${m.l}${U[m.k]?' ＋':''}`;
b.onclick=()=>{sel.has(m.k)?sel.delete(m.k):sel.add(m.k);draw()};$('chips').appendChild(b)})}
function toDate(v){if(v instanceof Date)return isNaN(v)?null:new Date(Date.UTC(v.getFullYear(),v.getMonth(),v.getDate()));
if(typeof v=='number')return v>20000&&v<80000?new Date(Math.round((v-25569)*864e5)):(v>19000101&&v<21001231?new Date(Date.UTC(Math.floor(v/10000),Math.floor(v/100)%100-1,v%100)):null);
const s=String(v).trim();let m=s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/);if(m)return new Date(Date.UTC(+m[3],+m[2]-1,+m[1]));
m=s.match(/^(\d{4})[\/.\-]?(\d{2})[\/.\-]?(\d{2})/);return m?new Date(Date.UTC(+m[1],+m[2]-1,+m[3])):null}
const num=v=>{if(typeof v=='number')return v;if(v==null)return NaN;const s=String(v).trim().replace(/\s/g,'');return /^-?[\d.,]+$/.test(s)?Number(s.includes(',')?s.replace(/\./g,'').replace(',','.'):s):NaN};
function parseRows(rows,acc){let hi=-1,cD=0,cO=1,cP=2,cV=3;
for(let i=0;i<Math.min(rows.length,15);i++){const r=(rows[i]||[]).map(x=>String(x==null?'':x));const d=r.findIndex(x=>/^\s*data/i.test(x));
if(d>=0){hi=i;cD=d;const o=r.findIndex(x=>/^\s*ora\b/i.test(x)),p=r.findIndex(x=>/^\s*period/i.test(x)),v=r.findIndex(x=>/€|mwh|pun|prezzo/i.test(x));cO=o;cP=p;cV=v>=0?v:Math.max(...[d,o,p])+1;break}}
let n=0;
for(let i=hi+1;i<rows.length;i++){const r=rows[i]||[];const dt=toDate(r[cD]);if(!dt||isNaN(dt))continue;
let h=cO>=0?Number(r[cO]):NaN;if(!(h>=1&&h<=24)&&cP>=0){const p=Number(r[cP]);h=p>=1&&p<=96?Math.ceil(p/4):NaN}
const p=num(r[cV]);if(!(h>=1&&h<=24)||isNaN(p))continue;
const k=dt.getUTCFullYear()+'-'+String(dt.getUTCMonth()+1).padStart(2,'0'),dd=dt.getUTCDate();
acc[k]=acc[k]||{};acc[k][dd]=acc[k][dd]||Array.from({length:24},()=>[]);acc[k][dd][h-1].push(p);n++}
return n}
async function readFiles(files){const acc={},bad=[];
if(typeof XLSX=='undefined'){$('msg').textContent='Libreria Excel non caricata: controlla la connessione e ricarica la pagina';return}
for(const f of files){try{const wb=XLSX.read(await f.arrayBuffer(),{type:'array'});let n=0,prev='';
for(const sn of wb.SheetNames){const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:null});if(!prev)prev=JSON.stringify((rows[0]||[]).slice(0,5));n+=parseRows(rows,acc);if(n)break}
if(!n)bad.push(f.name+' (nessuna riga valida; prima riga: '+prev+')')}catch(e){bad.push(f.name+' ('+(e&&e.message||e)+')')}}
const added=[];
Object.keys(acc).sort().forEach(k=>{const ds=Object.keys(acc[k]).map(Number).sort((a,b)=>a-b),mo=+k.slice(5),y=k.slice(0,4),
dim=new Date(Date.UTC(+y,mo,0)).getUTCDate(),last=ds[ds.length-1],cm=CM[mo]||{c:'#888',ph:'mid'};
U[k]={k,l:`${MN[mo-1]} ${y}`+(last<dim?` (1-${last})`:''),ph:cm.ph,c:cm.c,d:ds.map(d=>acc[k][d].map(a=>a.length?Math.round(avg(a)*10)/10:null))};R.delete(k);added.push(k)});
save();added.forEach(k=>sel.add(k));chips();draw();
$('msg').textContent=(added.length?'Caricati: '+added.map(k=>U[k].l).join(', '):'Nessuna serie caricata')+(bad.length?' · problemi: '+bad.join('; '):'')}
$('up').onclick=()=>$('fi').click();
$('fi').onchange=async e=>{const fs=[...e.target.files];e.target.value='';await readFiles(fs)};
let cf=null;const dl=()=>$('del').textContent='Elimina mesi selezionati';
$('del').onclick=()=>{if(!cf){const n=sel.size;$('del').textContent=`Conferma: elimina ${n} ${n==1?'mese':'mesi'}`;cf=setTimeout(()=>{cf=null;dl()},4000);return}
clearTimeout(cf);cf=null;dl();const n=sel.size;
[...sel].forEach(k=>{delete U[k];if(D.some(m=>m.k==k))R.add(k)});sel=new Set();save();chips();draw();$('msg').textContent=`Eliminati ${n} ${n==1?'mese':'mesi'}`};
$('rs').onclick=()=>{U={};R=new Set();save();sel=new Set(M().map(m=>m.k));chips();draw();$('msg').textContent='Ripristinati i dati originali'};
const setSel=f=>{sel=new Set(M().filter(f).map(m=>m.k));draw()};
$('all').onclick=()=>setSel(()=>1);$('none').onclick=()=>setSel(()=>0);
$('inv').onclick=()=>setSel(m=>m.ph=='inv');$('est').onclick=()=>setSel(m=>m.ph=='est');
$('mM').onclick=()=>{mode='m';draw()};$('mA').onclick=()=>{mode='a';draw()};
$('cd').onchange=e=>{days=e.target.checked;draw()};
const ch=new Chart($('c'),{type:'scatter',data:{datasets:[]},options:{animation:false,maintainAspectRatio:false,
interaction:{mode:'index',axis:'x',intersect:false},
scales:{x:{type:'linear',min:1,max:24,ticks:{stepSize:1},title:{display:true,text:'Ora (1–24)'}},
y:{type:'linear',title:{display:true,text:'Prezzo (€/MWh)'}}},
plugins:{legend:{position:'bottom',labels:{generateLabels:c=>Chart.defaults.plugins.legend.labels.generateLabels(c).filter(i=>{const d=c.data.datasets[i.datasetIndex];return d.tip&&d.label!='P25'}).map(i=>{if(c.data.datasets[i.datasetIndex].label=='P75')i.text='Banda P25–P75';return i})}},
tooltip:{filter:i=>i.dataset.tip,callbacks:{title:t=>'Ora '+t[0].raw.x,label:t=>`${t.dataset.label}: ${t.raw.y==null?'n.d.':t.raw.y.toFixed(1)} €/MWh`}}}}});
function draw(){
$('del').disabled=!sel.size;$('rs').disabled=!Object.keys(U).length&&!R.size;
document.querySelectorAll('.chip').forEach(b=>b.classList.toggle('on',sel.has(b.dataset.k)));
$('mM').classList.toggle('on',mode=='m');$('mA').classList.toggle('on',mode=='a');
const ms=M().filter(m=>sel.has(m.k)),ds=[];
const line=(label,p,c,w,dash,tip)=>({label,data:pts(p),borderColor:c,borderWidth:w,borderDash:dash,pointRadius:0,pointHoverRadius:tip?4:0,tension:.25,order:3,tip,showLine:true,spanGaps:false});
if(days)ms.forEach(m=>m.d.forEach(r=>ds.push(line('',r,rgba(m.c,mode=='a'?.13:.3),1,[],false))));
let all=[];
if(mode=='m')ms.forEach(m=>ds.push(line(m.l+' (media)',prof(m.d,avg),m.c,3,[],true)));
else{ms.forEach(m=>all=all.concat(m.d));
if(all.length){if(!days||ms.length==1){const q=p=>prof(all,a=>qt(a,p));
const lo=line('P25',q(.25),'rgba(0,0,0,0)',0,[],true),hi=line('P75',q(.75),'rgba(0,0,0,0)',0,[],true);
hi.fill='-1';hi.backgroundColor=rgba(css('--acc'),.16);lo.order=hi.order=2;lo.pointHoverRadius=hi.pointHoverRadius=0;ds.push(lo,hi)}
const a1=line('Media',prof(all,avg),css('--acc'),3.5,[],true),a2=line('Mediana',prof(all,med),css('--med'),3.5,[8,5],true);
a1.order=a2.order=1;ds.push(a1,a2)}}
if(mode=='m')ms.forEach(m=>all=all.concat(m.d));
const o=ch.options,g=css('--grid'),t=css('--mut');
['x','y'].forEach(a=>{o.scales[a].grid={color:g};o.scales[a].ticks.color=t;o.scales[a].title.color=t});
o.plugins.legend.labels.color=css('--fg');
ch.data.datasets=ds;ch.update('none');
const mp=all.length?prof(all,avg):null;
$('st').textContent=mp?`${all.length} giorni · prezzo medio ${avg(mp).toFixed(1)} €/MWh · ora più cara: ${mp.indexOf(Math.max(...mp))+1} (${Math.max(...mp).toFixed(1)}) · ora più economica: ${mp.indexOf(Math.min(...mp))+1} (${Math.min(...mp).toFixed(1)})`:'Seleziona almeno un mese.';}
matchMedia('(prefers-color-scheme: dark)').addEventListener('change',draw);
chips();draw();
</script></body></html>
'''

if __name__ == '__main__':
    main()
