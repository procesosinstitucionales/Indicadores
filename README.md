<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Ejercicio 1 – Fichas técnicas de indicadores</title>
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<style>
:root{box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px);--bg:#f6f8f4;--card:#fff;--tx:#1d2a1a;--mu:#5d6b58;--bd:#d3dccd;--ac:#5a9a1e;--ac2:#3f7a0f;--pe:#fff3b0;--er:#b3261e}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#141912;--card:#1d241a;--tx:#e8efe3;--mu:#9bab94;--bd:#33402e;--ac:#7ab83a;--ac2:#8fd04c;--pe:#4d4410;--er:#ff8a80}}
:root[data-theme="dark"]{--bg:#141912;--card:#1d241a;--tx:#e8efe3;--mu:#9bab94;--bd:#33402e;--ac:#7ab83a;--ac2:#8fd04c;--pe:#4d4410;--er:#ff8a80}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
*,*:before,*:after{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--tx);font:15px/1.45 system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif}
.w{max-width:980px;margin:0 auto;padding:18px 14px 60px}
h1{font-size:21px;margin:0 0 2px}h2{font-size:16px;margin:0 0 10px;color:var(--ac2)}
.sub{color:var(--mu);margin:0 0 16px;font-size:13px}
.c{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:14px;margin-bottom:14px}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
button,.btn{background:var(--ac);color:#fff;border:0;border-radius:8px;padding:9px 15px;font:600 14px inherit;cursor:pointer}
button.s{background:transparent;color:var(--ac2);border:1px solid var(--ac)}
button:disabled{opacity:.45;cursor:not-allowed}
input[type=text],textarea,select{width:100%;padding:8px;border:1px solid var(--bd);border-radius:7px;background:var(--bg);color:var(--tx);font:14px inherit}
textarea{min-height:58px;resize:vertical}
label{display:block;font-size:12px;font-weight:600;color:var(--mu);margin:8px 0 3px;text-transform:uppercase;letter-spacing:.02em}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:0 12px}.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:0 12px}
@media(max-width:640px){.g2,.g3{grid-template-columns:1fr}}
.pe{background:var(--pe)!important}
.ok{color:var(--ac2);font-weight:600}.er{color:var(--er)}.mu{color:var(--mu);font-size:13px}
.q{display:flex;gap:8px;align-items:center;padding:8px 10px;border:1px solid var(--bd);border-radius:8px;margin-top:6px;cursor:pointer}
.q.on{border-color:var(--ac);outline:1px solid var(--ac)}
.q span.st{margin-left:auto;font-size:12px;color:var(--mu)}
.ind{border-left:4px solid var(--ac)}
details summary{cursor:pointer;font-weight:600}
ul{margin:6px 0 0 18px;padding:0}
</style>
</head>
<body>
<div class="w">
<h1>Ejercicio 1 · Fichas técnicas de indicadores</h1>
<p class="sub">Adjunta la caracterización (PDF) de cada subproceso → valida los 3 indicadores → descarga el Excel diligenciado sobre la plantilla institucional.</p>

<div class="c" id="s0">
  <h2>1. Plantilla institucional de la ficha (.xlsx)</h2>
  <div class="row"><input type="file" id="tpl" accept=".xlsx"><span id="tplst" class="mu"></span></div>
  <p class="mu" style="margin:6px 0 0">Se carga una sola vez y queda guardada en este navegador.</p>
</div>

<div class="c">
  <h2>2. Caracterizaciones de subproceso (PDF)</h2>
  <div class="row"><input type="file" id="pdfs" accept=".pdf" multiple></div>
  <div id="queue"></div>
</div>

<div id="work" hidden>
  <div class="c">
    <div class="row"><button id="bAn">Analizar y proponer indicadores</button><span id="anst" class="mu"></span></div>
  </div>
  <div id="editor" hidden>
    <div class="c">
      <h2>Entendimiento del subproceso</h2>
      <label>Objetivo (extraído)</label><textarea id="f_obj"></textarea>
      <label>Síntesis y puntos críticos</label><textarea id="f_ent" style="min-height:110px"></textarea>
      <details><summary>Vacíos de información detectados</summary><ul id="vac"></ul></details>
    </div>
    <div class="c">
      <h2>Datos generales</h2>
      <div class="g3">
        <div><label>Proceso</label><input type="text" id="f_proc"></div>
        <div><label>Subproceso</label><input type="text" id="f_sub"></div>
        <div><label>Versión</label><input type="text" id="f_ver" value="1"></div>
      </div>
    </div>
    <div id="inds"></div>
    <div class="c">
      <div class="row"><button id="bGen">Aprobar indicadores y generar Excel</button><span id="genst" class="mu"></span></div>
      <p class="mu" style="margin:6px 0 0">Los campos amarillos son [PENDIENTE]: complétalos antes de generar. Todo es editable.</p>
    </div>
  </div>
</div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>
<script>
const $=id=>document.getElementById(id);
const TIPOS=["Eficacia","Eficiencia","Efectividad"],PERS=["Anual","Semestral","Trimestral","Bimestral","Mensual"];
let tplBuf=null,items=[],cur=-1,sampleNS=null,dlNS=null;

/* ---------- plantilla ---------- */
const b64=buf=>{let s="";const u=new Uint8Array(buf);for(let i=0;i<u.length;i+=0x8000)s+=String.fromCharCode.apply(null,u.subarray(i,i+0x8000));return btoa(s)};
const unb64=t=>{const s=atob(t),u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return u.buffer};
try{const t=localStorage.getItem("tpl_ficha");if(t){tplBuf=unb64(t);$("tplst").textContent="Plantilla guardada ✓";$("tplst").className="ok"}}catch(e){}
$("tpl").onchange=async e=>{const f=e.target.files[0];if(!f)return;tplBuf=await f.arrayBuffer();
  try{localStorage.setItem("tpl_ficha",b64(tplBuf))}catch(_){}
  $("tplst").textContent="Plantilla cargada: "+f.name+" ✓";$("tplst").className="ok"};

/* ---------- PDF ---------- */
pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
async function pdfText(buf){
  const pdf=await pdfjsLib.getDocument({data:buf}).promise;let out="";
  for(let p=1;p<=pdf.numPages;p++){
    const tc=await (await pdf.getPage(p)).getTextContent();const lines={};
    tc.items.forEach(i=>{const y=Math.round(i.transform[5]/3);(lines[y]=lines[y]||[]).push([i.transform[4],i.str])});
    Object.keys(lines).map(Number).sort((a,b)=>b-a).forEach(y=>{out+=lines[y].sort((a,b)=>a[0]-b[0]).map(x=>x[1]).join(" ")+"\n"});
    out+="\n";
  }
  return out;
}
$("pdfs").onchange=async e=>{
  for(const f of e.target.files){
    const it={name:f.name,status:"leyendo…",text:"",data:null};items.push(it);renderQ();
    try{it.text=await pdfText(await f.arrayBuffer());it.status=it.text.trim().length>80?"listo":"sin texto (¿escaneado?)"}catch(err){it.status="error de lectura"}
    renderQ();
  }
  if(cur<0&&items.length)sel(0);
  e.target.value="";
};
function renderQ(){
  $("queue").innerHTML="";
  items.forEach((it,i)=>{const d=document.createElement("div");d.className="q"+(i===cur?" on":"");
    const a=document.createElement("b");a.textContent=it.name;const s=document.createElement("span");s.className="st";
    s.textContent=it.data?(it.done?"Excel generado ✓":"propuesta lista"):it.status;d.append(a,s);d.onclick=()=>sel(i);$("queue").append(d)});
}
function sel(i){cur=i;renderQ();$("work").hidden=false;const it=items[i];
  $("editor").hidden=!it.data;$("anst").textContent="";if(it.data)fill(it.data);}

/* ---------- análisis ---------- */
const PROMPT=`Actúa como Ingeniero Senior de Procesos, experto en gestión por procesos e indicadores de gestión en instituciones de educación superior.
REGLA CERO INVENCIÓN: no inventes datos, metas, responsables ni fórmulas que no estén en el texto o no se deduzcan directamente de él. Si falta información crítica escribe exactamente "[PENDIENTE: qué información se requiere]".
Del texto de la caracterización (extraído de un PDF) analiza objetivo, alcance, proveedores, entradas, actividades, responsables, salidas, clientes y los indicadores/riesgos sugeridos. Si el nombre del proceso o subproceso aparece como marcador genérico ("NOMBRE DEL PROCESO", "NOMBRE DEL SUBPROCESO") usa "[PENDIENTE: ...]".
Propón exactamente 3 indicadores de gestión: directamente ligados al objetivo, que midan resultados (no solo ejecución de actividades), no redundantes, medibles con información razonable, con responsable identificable (usa el líder o un cargo citado en el texto; si no, [PENDIENTE]). Prioriza los indicadores mencionados en el documento cuando sean pertinentes.
Responde SOLO un JSON con esta forma:
{"proceso":"","subproceso":"","lider":"","objetivo":"","entendimiento":"síntesis ejecutiva del subproceso, puntos críticos y relación actividades-indicadores (máx 180 palabras)","vacios":["..."],
"indicadores":[{"nombre":"","definicion":"qué mide y para qué (1-2 frases)","responsable":"","premisa":"supuestos y fuente de información","formula":"(Numerador / Denominador) x 100, nombrando las variables","tipo":"Eficacia|Eficiencia|Efectividad","periodicidad":"Anual|Semestral|Trimestral|Bimestral|Mensual","meta":"","unidad":"%, días, etc.","justificacion":"pertinencia frente al objetivo"}]}
Para "tipo" y "periodicidad" usa solo los valores listados. La meta solo si hay base; si no, [PENDIENTE].
TEXTO DE LA CARACTERIZACIÓN:
`;
$("bAn").onclick=async()=>{
  const it=items[cur];if(!it||it.text.trim().length<80){$("anst").textContent="El PDF no tiene texto legible.";return}
  try{
    if(!sampleNS)sampleNS=await claude.use("sample");
    if(!sampleNS){$("anst").textContent="Función de análisis no disponible en esta vista.";return}
    $("bAn").disabled=true;$("anst").textContent="Analizando…";
    const d=await sampleNS.json(PROMPT+it.text.slice(0,30000),{modelTier:"complex",cache:false});
    it.data=d;it.done=false;renderQ();$("editor").hidden=false;fill(d);$("anst").textContent="Propuesta lista: valida y ajusta.";
  }catch(e){$("anst").textContent="Error: "+(e&&e.message||e.code||e)}
  $("bAn").disabled=false;
};

/* ---------- editor ---------- */
const mark=el=>{el.classList.toggle("pe",/\[PENDIENTE/i.test(el.value))};
function bind(el,v){el.value=v||"";mark(el);el.oninput=()=>mark(el)}
function fill(d){
  bind($("f_obj"),d.objetivo);bind($("f_ent"),d.entendimiento);bind($("f_proc"),d.proceso);bind($("f_sub"),d.subproceso);
  $("vac").innerHTML="";(d.vacios||[]).forEach(v=>{const l=document.createElement("li");l.textContent=v;$("vac").append(l)});
  $("inds").innerHTML="";
  (d.indicadores||[]).slice(0,3).forEach((x,i)=>{
    const c=document.createElement("div");c.className="c ind";c.dataset.i=i;
    const opt=(arr,v)=>arr.map(a=>`<option${a===v?" selected":""}>${a}</option>`).join("");
    c.innerHTML=`<h2>Indicador ${i+1}</h2>
    <label>Nombre</label><input type="text" data-k="nombre">
    <label>Definición</label><textarea data-k="definicion"></textarea>
    <div class="g2"><div><label>Responsable</label><input type="text" data-k="responsable"></div>
    <div><label>Premisa / fuente</label><input type="text" data-k="premisa"></div></div>
    <label>Fórmula</label><textarea data-k="formula"></textarea>
    <div class="g3"><div><label>Tipo</label><select data-k="tipo">${opt(TIPOS,x.tipo)}</select></div>
    <div><label>Periodicidad</label><select data-k="periodicidad">${opt(PERS,x.periodicidad)}</select></div>
    <div><label>Unidad de medida</label><input type="text" data-k="unidad"></div></div>
    <label>Meta</label><input type="text" data-k="meta">
    <label>Justificación (no va en la ficha)</label><textarea data-k="justificacion"></textarea>`;
    c.querySelectorAll("[data-k]").forEach(el=>{if(el.tagName!=="SELECT")bind(el,x[el.dataset.k])});
    $("inds").append(c);
  });
}
function collect(){
  const inds=[...document.querySelectorAll(".ind")].map(c=>{const o={};c.querySelectorAll("[data-k]").forEach(e=>o[e.dataset.k]=e.value.trim());return o});
  return{proceso:$("f_proc").value.trim(),subproceso:$("f_sub").value.trim(),ver:$("f_ver").value.trim()||"1",inds};
}

/* ---------- Excel ---------- */
function put(ws,addr,v,h){
  const c=ws.getCell(addr);c.value=v;
  c.alignment=Object.assign({},c.alignment,{wrapText:true,vertical:"middle"});
  if(/\[PENDIENTE/i.test(String(v)))c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFF3B0"}};
  if(h)ws.getRow(+addr.replace(/\D/g,"")).height=h;
}
const hgt=(t,w,min)=>Math.max(min,Math.ceil((t||"").length/w)*15+6);
const slug=s=>(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toUpperCase().replace(/[^A-Z0-9]+/g,"_").replace(/^_|_$/g,"");
$("bGen").onclick=async()=>{
  if(!tplBuf){$("genst").textContent="Primero carga la plantilla (.xlsx).";$("genst").className="er";return}
  $("genst").className="mu";$("genst").textContent="Generando…";
  try{
    const d=collect(),n=new Date(),dd=String(n.getDate()).padStart(2,"0"),mm=String(n.getMonth()+1).padStart(2,"0"),yy=String(n.getFullYear());
    const wb=new ExcelJS.Workbook();await wb.xlsx.load(tplBuf.slice(0));
    const lista=wb.worksheets[0],fs=wb.worksheets.filter(w=>/^IND\s?\d/i.test(w.name));
    put(lista,"C3",d.proceso);
    d.inds.forEach((x,i)=>{
      const r=6+i;put(lista,"B"+r,x.nombre);put(lista,"D"+r,d.subproceso);lista.getRow(r).height=hgt(x.nombre,40,30);
      const ws=fs[i];if(!ws)return;
      put(ws,"F1",d.proceso);put(ws,"F2",d.subproceso);put(ws,"F3","Ficha técnica");
      put(ws,"J2",d.ver);put(ws,"J4",dd);put(ws,"K4",mm);put(ws,"L4",yy);
      put(ws,"E6",x.nombre,hgt(x.nombre,70,24));put(ws,"E7",x.definicion,hgt(x.definicion,70,40));
      put(ws,"D9",x.responsable,hgt(x.responsable,30,32));put(ws,"H9",x.premisa,hgt(x.premisa,55,32));
      put(ws,"D11",x.formula,hgt(x.formula,50,40));put(ws,"J11",x.tipo);
      put(ws,"D12",x.periodicidad);put(ws,"G12",x.meta);put(ws,"J12",x.unidad);
    });
    const buf=await wb.xlsx.writeBuffer();
    const nm=slug(d.subproceso.replace(/\[PENDIENTE[^\]]*\]/i,""))||slug(items[cur].name.replace(/\.pdf$/i,""))||"SUBPROCESO";
    const filename="ficha tecnica de indicador_"+nm+".xlsx";
    if(!dlNS)dlNS=await claude.use("downloads");
    if(dlNS){await dlNS.save({filename,data:buf})}
    else{const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([buf]));a.download=filename;a.click()}
    items[cur].done=true;renderQ();$("genst").className="ok";$("genst").textContent="Excel generado: "+filename;
  }catch(e){$("genst").className="er";$("genst").textContent="Error: "+(e&&(e.message||e.code)||e)}
};
</script>
</body>
</html>
