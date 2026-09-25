const KEY="pcp_nexus_v13";
let old=JSON.parse(localStorage.getItem("pcp_nexus_v12")||'{"services":[],"components":[]}');
let db=JSON.parse(localStorage.getItem(KEY)||"null")||old;
let analysis=null;
const $=id=>document.getElementById(id);
document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>{
 document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));
 document.querySelectorAll(".panel").forEach(x=>x.classList.remove("active"));
 b.classList.add("active");$(b.dataset.tab).classList.add("active");render();
});
$("excel").onchange=()=>{let f=$("excel").files[0];$("fileStatus").textContent=f?`Seleccionado: ${f.name}`:"Sin archivos."};

$("scanBtn").onclick=async()=>{
 const f=$("excel").files[0];
 if(!f){alert("Selecciona un Excel primero.");return}
 if(typeof XLSX==="undefined"){alert("No está disponible el lector de Excel. Comprueba internet y vuelve a intentarlo.");return}
 $("reportResult").innerHTML='<div class="item">Procesando todas las hojas...</div>';
 const wb=XLSX.read(await f.arrayBuffer(),{type:"array"});
 let sheets=[], fullText="";
 wb.SheetNames.forEach(name=>{
   const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:""});
   const text=rows.map(r=>r.map(v=>String(v)).join(" | ")).join("\n");
   sheets.push({name,rows,text}); fullText+="\n"+text;
 });
 const fileWells=extractWells(f.name), contentWells=extractWells(fullText);
 const wells=[...new Set([...fileWells,...contentWells])];
 let selected=chooseWell(fileWells,contentWells);
 let confidence=selected?(fileWells.includes(selected)&&contentWells.includes(selected)?"Alta":"Media"):"Baja";
 $("well").value=selected||"NO DETECTADO";
 $("cwell").value=selected||"";
 const contradictions=wells.filter(x=>x!==selected);
 const labels=extractFields(fullText);
 const comps=extractComponents(fullText);
 const conflict=contradictions.length>0;
 $("failure").value=labels.failure||inferFailure(fullText);
 $("diagnosis").value=buildDiagnosis(labels,comps,selected);
 $("initialCondition").value=extractInitialCondition(fullText);
 $("majorChange").value=detectMajorChange(fullText)?"Sí":"No";
 $("changeType").value=detectChangeType(fullText);
 $("finalDate").value=detectFinalDate(fullText)||"";
 const questions=comps.filter(c=>c.destination==="Pendiente de confirmación");
 $("questions").innerHTML=questions.length?
 `<div class="question"><b>Confirmaciones necesarias</b><br>${questions.map((c,i)=>`
 <div class="small" style="margin-top:8px"><b>${i+1}. ${esc(c.type)} ${esc(c.serial||"sin serial")}</b> — ¿continúa en el pozo?
 <select data-q="${i}" class="destQ"><option>Pendiente de confirmación</option><option>Continúa instalado</option><option>Salió y volvió a ingresar</option><option>Salió y no volvió</option><option>Fue reemplazado</option></select></div>`).join("")}</div>`:"";
 $("evidence").innerHTML=`<div class="item ${conflict?'warning':'notice'}"><b>Identificación</b><br>Pozo: ${esc(selected||"No detectado")} · Confianza: ${confidence}<br>Hojas: ${sheets.length}<br>${conflict?`⚠️ También aparecen: ${contradictions.map(esc).join(", ")}. Verificar antes de guardar.`:"✓ No se detectó conflicto de pozo."}</div>
 <div class="item"><b>Componentes/seriales detectados</b><br>${comps.length?comps.map(c=>`<span class="tag">${esc(c.type)}: ${esc(c.serial||"No visible")}</span>`).join(""):"No se detectaron identificadores con los patrones actuales."}</div>`;
 $("reportResult").innerHTML=`<div class="item"><b>Reporte procesado</b><div class="muted">${esc(f.name)} · ${sheets.length} hoja(s)</div>${sheets.map(s=>`<span class="tag">${esc(s.name)}</span>`).join("")}</div>`;
 analysis={file:f.name,sheets,wells,selected,confidence,contradictions,labels,comps,photos:[...$("photos").files].map(x=>x.name),fullText};
};

$("saveService").onclick=()=>{
 if(!analysis){alert("Primero analiza el reporte.");return}
 if(!analysis.selected||analysis.contradictions.length){alert("No se puede guardar todavía: verifica el pozo y las inconsistencias detectadas.");return}
 const q=[...document.querySelectorAll(".destQ")];
 q.forEach((el,i)=>analysis.comps[i].destination=el.value);
 if(analysis.comps.some(c=>c.destination==="Pendiente de confirmación")){
   alert("Hay componentes inspeccionados cuyo destino sigue pendiente. Confirma si continúan, regresan, salen o fueron reemplazados.");
   return;
 }
 const service={id:Date.now(),well:analysis.selected,date:$("finalDate").value||new Date().toISOString().slice(0,10),
 file:analysis.file,photos:analysis.photos,failure:$("failure").value,majorChange:$("majorChange").value,
 changeType:$("changeType").value,initialCondition:$("initialCondition").value,diagnosis:$("diagnosis").value,
 components:analysis.comps,spacing:readSpacing()};
 db.services.unshift(service);
 analysis.comps.forEach(c=>db.components.unshift({id:Date.now()+Math.random(),well:analysis.selected,type:c.type,serial:c.serial||"No visible",status:c.status,destiny:c.destination,obs:c.obs||"Se conserva trazabilidad NEXUS 1.3."}));
 save();alert("Análisis NEXUS 1.3 guardado.");
};

$("componentForm").onsubmit=e=>{
 e.preventDefault();
 const photos=[...$("cphoto").files].map(x=>x.name);
 db.components.unshift({id:Date.now(),well:$("cwell").value,type:$("ctype").value,serial:$("serial").value||"No visible",
 status:$("status").value,destiny:$("destiny").value,photos,obs:$("obs").value});
 save();e.target.reset();$("cwell").value=analysis?.selected||"";alert("Componente registrado.");
};

$("validateSpacing").onclick=()=>{
 const s=readSpacing();
 if(!Number.isFinite(s.ref)||!Number.isFinite(s.spacing)){alert("Ingrese profundidad de referencia y espaciamiento.");return}
 const diff=Math.abs(s.ref-s.spacing);
 const yOk=s.y>=1.15&&s.y<=1.30;
 let calc=null;
 if([s.P,s.K,s.ref,s.d].every(Number.isFinite)) calc=(s.P*s.ref*s.K/1000)+s.d;
 $("spacingResult").innerHTML=`<div class="item ${diff<=0.5&&yOk?'notice':'warning'}">
 <b>${diff<=0.5&&yOk?'✓ VALIDADO':'⚠ REVISAR'}</b><br>
 Referencia (${esc(s.type)}): ${s.ref.toFixed(2)} ft<br>
 Espaciamiento de varilla: ${s.spacing.toFixed(2)} ft<br>
 Diferencia: ${diff.toFixed(2)} ft (máximo 0,50 ft)<br>
 Punta de tubing: ${Number.isFinite(s.tubing)?s.tubing.toFixed(2)+' ft (referencia independiente)':'No registrada'}<br>
 ${calc!==null?`Y calculado por fórmula: ${calc.toFixed(2)} ft<br>`:""}
 Factor de seguridad: ${Number.isFinite(s.y)?s.y.toFixed(2):"No registrado"} — ${yOk?"dentro de rango":"fuera de rango 1,15–1,30"}
 </div>`;
};

$("exportNexus").onclick=()=>{
 const blob=new Blob([JSON.stringify({version:"1.3",exportedAt:new Date().toISOString(),db},null,2)],{type:"application/json"});
 const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`PCP_NEXUS_1.3_${new Date().toISOString().slice(0,10)}.json`;a.click();URL.revokeObjectURL(a.href);
};

function readSpacing(){return {
 model:$("pumpModel").value,type:$("pumpType").value,ref:Number($("refDepth").value),spacing:Number($("rodSpacing").value),
 tubing:Number($("tubingEnd").value),K:Number($("factorK").value),P:Number($("pressureP").value),d:Number($("spacerD").value),
 y:Number($("safetyY").value),rods:Number($("rodCount").value)
};}
function extractWells(t){const s=String(t).toUpperCase(),m=s.match(/\bV[\s\-–_]*\d{2,4}\b/g)||[];return [...new Set(m.map(x=>x.replace(/[\s–_]+/g,"-")))];}
function chooseWell(f,c){if(f.length===1&&c.includes(f[0]))return f[0];if(c.length===1)return c[0];if(f.length===1&&!c.length)return f[0];return f.find(x=>c.includes(x))||c[0]||f[0]||"";}
function extractFields(t){const s=String(t);const find=ps=>{for(const p of ps){const m=s.match(p);if(m)return m[1].trim()}return""};return {failure:find([/(?:MOTIVO|CAUSA|FALLA|PROBLEMA|DESCRIPCI[ÓO]N DEL SERVICIO)\s*[:\-]\s*([^\n|]+)/i]),diagnosis:find([/(?:DIAGN[ÓO]STICO|CONCLUSI[ÓO]N|OBSERVACIONES)\s*[:\-]\s*([^\n|]+)/i])};}
function extractComponents(t){
 const out=[],s=String(t);
 const patterns=[
 ["Estator",/(?:ESTATOR)[^A-Z0-9]{0,50}(?:SERIAL|S\/N|N°|NO)?[^A-Z0-9]{0,20}([A-Z0-9][A-Z0-9×x\-\/]{4,})/ig],
 ["Rotor",/(?:ROTOR)[^A-Z0-9]{0,50}(?:SERIAL|S\/N|N°|NO)?[^A-Z0-9]{0,20}([A-Z0-9][A-Z0-9×x\-\/]{4,})/ig],
 ["Motor",/(?:MOTOR)[^A-Z0-9]{0,50}(?:SERIAL|S\/N|N°|NO)?[^A-Z0-9]{0,20}([A-Z0-9][A-Z0-9×x\-\/]{4,})/ig],
 ["Cabezal",/(?:CABEZAL|HEAD)[^A-Z0-9]{0,50}(?:SERIAL|S\/N|N°|NO)?[^A-Z0-9]{0,20}([A-Z0-9][A-Z0-9×x\-\/]{4,})/ig],
 ["Check valve",/(?:CHECK\s*VALVE)[^A-Z0-9]{0,50}(?:SERIAL|S\/N|N°|NO)?[^A-Z0-9]{0,20}([A-Z0-9][A-Z0-9×x\-\/]{3,})/ig],
 ["Niplesilla",/(?:NIPLESILLA)[^A-Z0-9]{0,50}(?:SERIAL|S\/N|N°|NO)?[^A-Z0-9]{0,20}([A-Z0-9][A-Z0-9×x\-\/]{3,})/ig]
 ];
 for(const [type,re] of patterns){let m;while((m=re.exec(s))&&out.length<80){let serial=m[1].replace(/[.,;:]$/,"");if(!/^(SERIAL|MODELO|TIPO|OPERATIVO|NUEVO)$/i.test(serial))out.push({type,serial,destination:"Pendiente de confirmación",status:"Detectado",obs:"Extraído del reporte; validar con fotografía."})}}
 return [...new Map(out.map(x=>[x.type+"|"+x.serial,x])).values()];
}
function inferFailure(t){const s=String(t).toLowerCase();const keys=["bajo torque","varilla desconectada","varilla partida","rotor partido","fuga por coupling","tubería desconectada","run life","cambio de bomba","optimización","recañoneo"];return keys.find(k=>s.includes(k))||"No identificado automáticamente; revisar el reporte.";}
function buildDiagnosis(l,c,w){let a=[];if(l.failure)a.push("Hecho extraído: "+l.failure);if(l.diagnosis)a.push("Observación extraída: "+l.diagnosis);if(c.length)a.push("Se detectaron "+c.length+" identificadores; cada componente debe validarse con evidencia y destino.");a.push("Diagnóstico preliminar: verificar contra la evidencia del servicio.");return a.join("\n");}
function extractInitialCondition(t){const s=String(t),m=s.match(/(?:TENSI[ÓO]N|PESO|GIRO|TORQUE|PRESI[ÓO]N)[\s\S]{0,800}/i);return m?m[0].slice(0,1200):"";}
function detectMajorChange(t){return /rotor\s+partido|cambio\s+de\s+bomba|cambio\s+de\s+estator|cambio\s+de\s+rotor|cambio\s+de\s+sarta/i.test(t);}
function detectChangeType(t){const s=String(t).toLowerCase();if(/rotor\s+partido/.test(s))return"Rotor partido";if(/cambio\s+de\s+bomba/.test(s))return"Cambio de bomba";if(/cambio\s+de\s+estator/.test(s))return"Cambio de estator";if(/cambio\s+de\s+rotor/.test(s))return"Cambio de rotor";if(/cambio\s+de\s+sarta/.test(s))return"Cambio de sarta";return"";}
function detectFinalDate(t){const m=String(t).match(/(?:ARRANQUE|FECHA\s+DE\s+ARRANQUE|TERMINACI[ÓO]N)[^0-9]{0,30}(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/i);if(!m)return"";let y=m[3].length===2?"20"+m[3]:m[3];return `${y}-${String(m[2]).padStart(2,"0")}-${String(m[1]).padStart(2,"0")}`;}
function save(){localStorage.setItem(KEY,JSON.stringify(db));render();}
function render(){
 $("wellCount").textContent=new Set(db.services.map(x=>x.well)).size;$("serviceCount").textContent=db.services.length;$("componentCount").textContent=db.components.length;
 $("historyList").innerHTML=[...db.services.map(s=>`<div class="item"><b>${esc(s.well)}</b><div class="muted">${esc(s.date)} · ${esc(s.file)}<br>${esc(s.failure||"Sin falla")}<br>${esc(s.majorChange==="Sí"?"Cambio: "+s.changeType:"Sin cambio importante")}</div></div>`),...db.components.map(c=>`<div class="item"><b>${esc(c.well)} · ${esc(c.type)}</b><div class="muted">Serial: ${esc(c.serial)} · Estado: ${esc(c.status)} · Destino: ${esc(c.destiny)}<br>${esc(c.obs||"")}</div></div>`)].join("")||'<div class="box">Aún no hay registros.</div>';
 $("componentList").innerHTML=db.components.slice(0,30).map(c=>`<div class="item"><b>${esc(c.well)} · ${esc(c.type)}</b> — ${esc(c.serial)}<br><span class="small">Estado: ${esc(c.status)} · Destino: ${esc(c.destiny)}</span></div>`).join("")||"";
}
function esc(v){return String(v??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]))}
render();
