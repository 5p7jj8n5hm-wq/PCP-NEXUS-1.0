const KEY="pcp_nexus_v11";let db=JSON.parse(localStorage.getItem(KEY)||'{"services":[],"components":[]}');let extracted={};
const $=id=>document.getElementById(id);
document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>{document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));document.querySelectorAll(".panel").forEach(x=>x.classList.remove("active"));b.classList.add("active");$(b.dataset.tab).classList.add("active");render()});
$("excel").onchange=()=>{$("fileStatus").textContent=$("excel").files.length?`Excel seleccionado: ${$("excel").files[0].name}`:"Sin archivos cargados."};
$("photos").onchange=()=>{$("fileStatus").textContent+=` · ${$("photos").files.length} foto(s)`};
$("scanBtn").onclick=async()=>{
 const f=$("excel").files[0]; if(!f){alert("Selecciona primero un Excel.");return}
 if(typeof XLSX==="undefined"){alert("No se cargó el lector de Excel. Verifica conexión a internet.");return}
 const data=await f.arrayBuffer(); const wb=XLSX.read(data,{type:"array"}); let out=[];
 wb.SheetNames.forEach(n=>{const rows=XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,defval:""});out.push({sheet:n,rows:rows.slice(0,80)});});
 extracted={file:f.name,sheets:out}; let text=out.map(s=>s.rows.map(r=>r.join(" | ")).join("\n")).join("\n");
 const m=text.match(/\b(V[- ]?\d{2,4})\b/i); if(m)$("well").value=m[1].toUpperCase().replace(" ","-");
 $("extract").innerHTML=`<div class="item"><b>Reporte procesado</b><div class="muted">${esc(f.name)} · ${out.length} hoja(s)</div><div class="tag">Extracción local completada</div></div>`;
};
$("saveService").onclick=()=>{let s={id:Date.now(),well:$("well").value||"Sin pozo",date:new Date().toISOString().slice(0,10),file:extracted.file||"Sin Excel",photos:[...$("photos").files].map(f=>f.name),failure:$("failure").value,diagnosis:$("diagnosis").value,raw:extracted};db.services.unshift(s);save();alert("Análisis guardado en PCP NEXUS 1.1");};
$("componentForm").onsubmit=e=>{e.preventDefault();db.components.unshift({id:Date.now(),well:$("cwell").value,type:$("ctype").value,serial:$("serial").value,status:$("status").value,obs:$("obs").value});save();e.target.reset();alert("Componente registrado");};
function save(){localStorage.setItem(KEY,JSON.stringify(db));render()}
function render(){ $("wellCount").textContent=new Set(db.services.map(x=>x.well)).size;$("serviceCount").textContent=db.services.length;$("componentCount").textContent=db.components.length;
$("historyList").innerHTML=[...db.services.map(s=>`<div class="item"><b>${esc(s.well)}</b><div class="muted">${esc(s.date)} · ${esc(s.file)}<br>${esc(s.failure||"Sin falla registrada")}<br>${esc(s.diagnosis||"")}</div></div>`),...db.components.map(c=>`<div class="item"><b>${esc(c.well)} · ${esc(c.type)}</b><div class="muted">Serial: ${esc(c.serial)} · ${esc(c.status)}<br>${esc(c.obs||"")}</div></div>`)].join("")||'<div class="box">Aún no hay registros.</div>'}
function esc(v){return String(v??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]))}render();