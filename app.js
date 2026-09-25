const KEY="pcp_nexus_v131";
let old=JSON.parse(localStorage.getItem("pcp_nexus_v13")||localStorage.getItem("pcp_nexus_v12")||'{"services":[],"components":[]}');
let db=JSON.parse(localStorage.getItem(KEY)||"null")||old;
let analysis=null;
const $=id=>document.getElementById(id);

document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>{
  document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));
  document.querySelectorAll(".panel").forEach(x=>x.classList.remove("active"));
  b.classList.add("active"); $(b.dataset.tab).classList.add("active"); render();
});

$("excel").onchange=()=>{const f=$("excel").files[0]; $("fileStatus").textContent=f?`Seleccionado: ${f.name}`:"Sin archivos."};

$("scanBtn").onclick=async()=>{
  const f=$("excel").files[0];
  if(!f){alert("Selecciona un Excel primero.");return}
  if(typeof XLSX==="undefined"){alert("No está disponible el lector de Excel. Comprueba internet y vuelve a intentarlo.");return}
  $("reportResult").innerHTML='<div class="item">Procesando hojas con análisis contextual...</div>';
  const wb=XLSX.read(await f.arrayBuffer(),{type:"array",cellDates:true});
  const sheets=[]; let fullText="";
  wb.SheetNames.forEach(name=>{
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:"",raw:false});
    const text=rows.map(r=>r.map(v=>String(v??"")).join(" | ")).join("\n");
    sheets.push({name,rows,text}); fullText+="\n"+text;
  });

  const fileWells=extractWells(f.name), contentWells=extractWells(fullText);
  const wells=[...new Set([...fileWells,...contentWells])];
  const selected=chooseWell(fileWells,contentWells);
  const confidence=selected?(fileWells.includes(selected)&&contentWells.includes(selected)?"Alta":"Media"):"Baja";
  $("well").value=selected||"NO DETECTADO"; $("cwell").value=selected||"";
  const contradictions=wells.filter(x=>x!==selected);
  const labels=extractFields(fullText);
  const events=extractEvents(fullText);
  const comps=extractComponentsFromSheets(sheets);
  const finalDate=detectFinalDate(sheets,fullText);
  const major=detectMajorChange(fullText);
  $("failure").value=labels.failure||inferFailure(fullText);
  $("majorChange").value=major?"Sí":"No";
  $("changeType").value=detectChangeType(fullText);
  $("finalDate").value=finalDate||"";
  $("initialCondition").value=extractInitialCondition(fullText);
  $("diagnosis").value=buildDiagnosis(labels,events,comps,selected);

  const questions=comps.filter(c=>c.destination==="Pendiente de confirmación");
  $("questions").innerHTML=questions.length?
    `<div class="question"><b>Confirmaciones necesarias</b><div class="small muted">Solo se pregunta por componentes con evidencia estructurada suficiente. Los hallazgos narrativos no se convierten automáticamente en componentes.</div>${questions.map((c,i)=>`
      <div class="small" style="margin-top:10px"><b>${i+1}. ${esc(c.type)}${c.serial?` SN ${esc(c.serial)}`:""}</b> — ¿cuál es su destino?
      <select data-q="${i}" class="destQ"><option>Pendiente de confirmación</option><option>Continúa instalado</option><option>Salió y volvió a ingresar</option><option>Salió y no volvió</option><option>Fue reemplazado</option></select></div>`).join("")}</div>` :
    `<div class="item notice"><b>✓ Sin confirmaciones automáticas</b><br>No se detectaron componentes físicos con destino incierto mediante evidencia estructurada.</div>`;

  const compHtml=comps.length?comps.map(c=>`<span class="tag">${esc(c.type)}${c.serial?`: ${esc(c.serial)}`:""}</span>`).join(""):"No se identificaron componentes físicos con evidencia estructurada suficiente.";
  const eventHtml=events.length?events.map(e=>`<span class="tag">${esc(e)}</span>`).join(""):"No se detectaron eventos con los patrones configurados.";
  const conflict=contradictions.length>0;
  $("evidence").innerHTML=`
    <div class="item ${conflict?'warning':'notice'}"><b>Identificación</b><br>Pozo: ${esc(selected||"No detectado")} · Confianza: ${confidence}<br>Hojas: ${sheets.length}<br>${conflict?`⚠️ También aparecen: ${contradictions.map(esc).join(", ")}. Verificar antes de guardar.`:"✓ No se detectó conflicto de pozo."}</div>
    <div class="item"><b>Componentes físicos detectados</b><br>${compHtml}</div>
    <div class="item"><b>Eventos/hallazgos narrativos</b><br>${eventHtml}</div>`;
  $("reportResult").innerHTML=`<div class="item"><b>Reporte procesado con motor contextual 1.3.1</b><div class="muted">${esc(f.name)} · ${sheets.length} hoja(s) · ${comps.length} componente(s) físico(s) · ${events.length} evento(s)</div>${sheets.map(s=>`<span class="tag">${esc(s.name)}</span>`).join("")}</div>`;

  analysis={version:"1.3.1",file:f.name,sheets,wells,selected,confidence,contradictions,labels,events,comps,photos:[...$("photos").files].map(x=>x.name),fullText,finalDate};
};

$("saveService").onclick=()=>{
  if(!analysis){alert("Primero analiza el reporte.");return}
  if(!analysis.selected||analysis.contradictions.length){alert("No se puede guardar todavía: verifica el pozo y las inconsistencias detectadas.");return}
  const q=[...document.querySelectorAll(".destQ")];
  const pending=analysis.comps.filter(c=>c.destination==="Pendiente de confirmación");
  q.forEach((el,i)=>{if(pending[i])pending[i].destination=el.value;});
  if(analysis.comps.some(c=>c.destination==="Pendiente de confirmación")){
    alert("Hay componentes inspeccionados cuyo destino sigue pendiente. Confirma el destino antes de construir el completamiento final.");return;
  }
  const service={id:Date.now(),well:analysis.selected,date:analysis.finalDate||$("finalDate").value||"",
    file:analysis.file,photos:analysis.photos,failure:$("failure").value,majorChange:$("majorChange").value,
    changeType:$("changeType").value,initialCondition:$("initialCondition").value,diagnosis:$("diagnosis").value,
    events:analysis.events,components:analysis.comps,spacing:readSpacing(),generatedReport:null};
  db.services.unshift(service);
  analysis.comps.forEach(c=>db.components.unshift({id:Date.now()+Math.random(),well:analysis.selected,type:c.type,serial:c.serial||"No visible",status:c.status,destiny:c.destination,obs:c.obs||"Se conserva trazabilidad NEXUS 1.3.1.",photos:c.photos||[]}));
  save();alert("Análisis NEXUS 1.3.1 guardado.");
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
  let calc=null;if([s.P,s.K,s.ref,s.d].every(Number.isFinite))calc=(s.P*s.ref*s.K/1000)+s.d;
  const sfOk=!Number.isFinite(s.safetyFactor)||(s.safetyFactor>=1.15&&s.safetyFactor<=1.30);
  $("spacingResult").innerHTML=`<div class="item ${diff<=0.5&&sfOk?'notice':'warning'}"><b>${diff<=0.5&&sfOk?'✓ VALIDADO':'⚠ REVISAR'}</b><br>
  Referencia (${esc(s.type)}): ${s.ref.toFixed(2)} ft<br>Espaciamiento de varilla: ${s.spacing.toFixed(2)} ft<br>Diferencia: ${diff.toFixed(2)} ft (máximo 0,50 ft)<br>
  Punta de tubing: ${Number.isFinite(s.tubing)?s.tubing.toFixed(2)+' ft (referencia independiente)':'No registrada'}<br>
  ${calc!==null?`Y calculado por fórmula: ${calc.toFixed(2)} ft<br>`:""}
  Factor de seguridad: ${Number.isFinite(s.safetyFactor)?s.safetyFactor.toFixed(2):"No registrado"} — ${sfOk?"dentro de rango":"fuera de rango 1,15–1,30"}
  ${diff>0.5?`<br><b>INCONSISTENCIA DE ESPACIAMIENTO. Requiere revisión.</b>`:""}</div>`;
};

$("exportNexus").onclick=()=>{
  const blob=new Blob([JSON.stringify({version:"1.3.1",exportedAt:new Date().toISOString(),db},null,2)],{type:"application/json"});
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`PCP_NEXUS_1.3.1_${new Date().toISOString().slice(0,10)}.json`;a.click();URL.revokeObjectURL(a.href);
};

function readSpacing(){return {model:$("pumpModel").value,type:$("pumpType").value,ref:Number($("refDepth").value),spacing:Number($("rodSpacing").value),tubing:Number($("tubingEnd").value),K:Number($("factorK").value),P:Number($("pressureP").value),d:Number($("spacerD").value),safetyFactor:Number($("safetyFactor").value),rods:Number($("rodCount").value)};}
function extractWells(t){const s=String(t).toUpperCase(),m=s.match(/\bV[\s\-–_]*\d{2,4}\b/g)||[];return [...new Set(m.map(x=>x.replace(/[\s–_]+/g,"-")))];}
function chooseWell(f,c){if(f.length===1&&c.includes(f[0]))return f[0];if(c.length===1)return c[0];if(f.length===1&&!c.length)return f[0];return f.find(x=>c.includes(x))||c[0]||f[0]||"";}
function norm(s){return String(s??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toUpperCase().replace(/\s+/g," ").trim();}
function extractFields(t){const s=String(t);const find=ps=>{for(const p of ps){const m=s.match(p);if(m)return m[1].trim()}return""};return {failure:find([/(?:MOTIVO|FALLA|PROBLEMA|DESCRIPCI[ÓO]N DEL SERVICIO)\s*[:\-]\s*([^\n|]+)/i]),diagnosis:find([/(?:DIAGN[ÓO]STICO|CONCLUSI[ÓO]N)\s*[:\-]\s*([^\n|]+)/i])};}

const COMPONENTS=[
 {type:"Estator saliente",aliases:["ESTATOR","ESTATOR SALIENTE"]},
 {type:"Rotor saliente",aliases:["ROTOR","ROTOR SALIENTE"]},
 {type:"Motor",aliases:["MOTOR ELECTRICO","MOTOR"]},
 {type:"Cabezal",aliases:["CABEZAL","HEAD"]},
 {type:"Anclas",aliases:["ANCLA","ANCLAS","TORQUE STOPPER"]},
 {type:"Separador de gas",aliases:["SEPARADOR DE GAS","GAS SEPARATOR"]},
 {type:"Check valve",aliases:["CHECK VALVE","VALVULA CHECK"]},
 {type:"Niplesilla",aliases:["NIPLESILLA"]},
 {type:"Hollow rod",aliases:["HOLLOW ROD"]},
 {type:"Varillas",aliases:["VARILLAS","S/R","SUCKER RODS"]},
 {type:"Couplings",aliases:["COUPLING","COUPLINGS"]},
 {type:"Barra lisa",aliases:["BARRA LISA","POLISHED ROD"]},
 {type:"Pony rods",aliases:["PONY ROD","PONY RODS"]},
 {type:"Centralizadores",aliases:["CENTRALIZADOR","CENTRALIZADORES"]}
];
function componentForText(s){const n=norm(s);return COMPONENTS.find(c=>c.aliases.some(a=>n.includes(norm(a))));}
function serialFromText(s){const n=String(s??"");
  const m=n.match(/(?:SERIAL|S\/N|S\.N\.|SN|N[°º]?\s*(?:DE\s*)?SERIE|NO\.?\s*DE\s*SERIE)\s*[:#\-]?\s*([A-Z0-9][A-Z0-9._\-/×x]{2,})/i);
  return m?m[1].replace(/[.,;:]$/i,""):"";
}
function explicitSerialEvidence(rowText){return /(?:SERIAL|S\/N|S\.N\.|\bSN\b|N[°º]?\s*(?:DE\s*)?SERIE|NO\.?\s*DE\s*SERIE)/i.test(rowText);}
function extractComponentsFromSheets(sheets){
  const out=[];
  for(const sh of sheets){
    for(let ri=0;ri<sh.rows.length;ri++){
      const cells=sh.rows[ri].map(v=>String(v??"")).filter(Boolean);
      if(!cells.length)continue;
      const rowText=cells.join(" | ");
      const comp=componentForText(rowText);
      if(!comp)continue;
      const serial=serialFromText(rowText);
      // Only create an automatic component when the row carries explicit serial evidence.
      // A narrative phrase such as "rotor partido" is an event, not a physical component.
      if(!explicitSerialEvidence(rowText)||!serial)continue;
      const nearby=cells.join(" | ");
      const status=detectStatus(nearby);
      const key=comp.type+"|"+serial;
      if(!out.some(x=>x.key===key))out.push({key,type:comp.type,serial,status,destination:"Pendiente de confirmación",obs:`Evidencia estructurada en hoja ${sh.name}, fila ${ri+1}.`,sheet:sh.name,row:ri+1,photos:[]});
    }
  }
  return out;
}
function detectStatus(s){const n=norm(s);if(/PARTID[OA]|ROTO|FRACTUR/.test(n))return"Partido";if(/DESCONECTAD/.test(n))return"Desconectado";if(/CORROSI/.test(n))return"Corroído";if(/NO OPERATIVO|NO FUNCIONA/.test(n))return"No operativo";if(/DESGAST/.test(n))return"Desgastado";if(/OPERATIVO|FUNCIONANDO/.test(n))return"Operativo";return"No evaluable";}
function extractEvents(t){const s=String(t).toLowerCase(),keys=["bajo torque","varilla desconectada","varilla partida","rotor partido","fuga por coupling","tubería desconectada","run life","cambio de bomba","cambio de estator","cambio de rotor","cambio de sarta","optimización","recañoneo","elastómero desgarrado"];return [...new Set(keys.filter(k=>s.includes(k)))];}
function inferFailure(t){const s=String(t).toLowerCase();const keys=["bajo torque","varilla desconectada","varilla partida","rotor partido","fuga por coupling","tubería desconectada","run life","cambio de bomba","optimización","recañoneo"];return keys.find(k=>s.includes(k))||"No identificado automáticamente; revisar el reporte.";}
function buildDiagnosis(l,events,c,w){const a=[];if(l.failure)a.push("Motivo/falla inicial extraído: "+l.failure);if(events.length)a.push("Eventos/hallazgos detectados: "+events.join(", ")+".");if(l.diagnosis)a.push("Observación/conclusión explícita: "+l.diagnosis);if(c.length)a.push("Se identificaron "+c.length+" componentes físicos con evidencia estructurada; cada uno requiere estado y destino.");else a.push("No se identificaron componentes físicos mediante evidencia estructurada. No se crearán componentes a partir de palabras narrativas aisladas.");a.push("Diagnóstico preliminar: validar contra evidencia del servicio antes de construir el completamiento final.");return a.join("\n");}
function extractInitialCondition(t){const s=String(t),m=s.match(/(?:TENSI[ÓO]N|PESO|GIRO|TORQUE|PRESI[ÓO]N)[\s\S]{0,800}/i);return m?m[0].slice(0,1200):"";}
function detectMajorChange(t){return /rotor\s+partido|cambio\s+de\s+bomba|cambio\s+de\s+estator|cambio\s+de\s+rotor|cambio\s+de\s+sarta|cambio\s+importante/i.test(t);}
function detectChangeType(t){const s=norm(t);if(/ROTOR\s+PARTIDO/.test(s))return"Rotor partido";if(/CAMBIO\s+DE\s+BOMBA/.test(s))return"Cambio de bomba";if(/CAMBIO\s+DE\s+ESTATOR/.test(s))return"Cambio de estator";if(/CAMBIO\s+DE\s+ROTOR/.test(s))return"Cambio de rotor";if(/CAMBIO\s+DE\s+SARTA/.test(s))return"Cambio de sarta";return"";}
function parseDateAt(m){if(!m)return"";let [a,b,c]=m.slice(1,4);if(String(c).length===4){return `${c}-${String(b).padStart(2,"0")}-${String(a).padStart(2,"0")}`;}return `20${c}-${String(b).padStart(2,"0")}-${String(a).padStart(2,"0")}`;}
function detectFinalDate(sheets,text){
  const patterns=[/(?:FECHA\s+DE\s+ARRANQUE|ARRANQUE|PUESTA\s+EN\s+MARCHA|FECHA\s+FINAL|TERMINACI[ÓO]N\s+EFECTIVA)[^0-9]{0,40}(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/i];
  for(const p of patterns){const m=String(text).match(p);if(m)return parseDateAt(m);}
  // Search row context for the same explicit labels, preserving the source evidence.
  for(const sh of sheets)for(const row of sh.rows){const r=row.map(v=>String(v??"")).join(" | ");if(/FECHA\s+DE\s+ARRANQUE|ARRANQUE|PUESTA\s+EN\s+MARCHA|FECHA\s+FINAL|TERMINACI[ÓO]N\s+EFECTIVA/i.test(r)){const m=r.match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/);if(m)return parseDateAt(m);}}
  return"";
}
function save(){localStorage.setItem(KEY,JSON.stringify(db));render();}
function render(){
  $("wellCount").textContent=new Set(db.services.map(x=>x.well)).size;$("serviceCount").textContent=db.services.length;$("componentCount").textContent=db.components.length;
  $("historyList").innerHTML=[...db.services.map(s=>`<div class="item"><b>${esc(s.well)}</b><div class="muted">${esc(s.date||"Fecha pendiente")} · ${esc(s.file)}<br>${esc(s.failure||"Sin falla")}<br>${esc(s.majorChange==="Sí"?"Cambio: "+s.changeType:"Sin cambio importante")}</div></div>`),...db.components.map(c=>`<div class="item"><b>${esc(c.well)} · ${esc(c.type)}</b><div class="muted">Serial: ${esc(c.serial)} · Estado: ${esc(c.status)} · Destino: ${esc(c.destiny)}<br>${esc(c.obs||"")}</div></div>`)].join("")||'<div class="box">Aún no hay registros.</div>';
  $("componentList").innerHTML=db.components.slice(0,30).map(c=>`<div class="item"><b>${esc(c.well)} · ${esc(c.type)}</b> — ${esc(c.serial)}<br><span class="small">Estado: ${esc(c.status)} · Destino: ${esc(c.destiny)}</span></div>`).join("")||"";
}
function esc(v){return String(v??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]))}
render();
