const KEY="pcp_nexus_v12";
let db=JSON.parse(localStorage.getItem(KEY)||'{"services":[],"components":[]}');
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
 const fileWells=extractWells(f.name);
 const contentWells=extractWells(fullText);
 const wells=[...new Set([...fileWells,...contentWells])];
 let selected=chooseWell(fileWells,contentWells);
 let confidence=selected?(fileWells.includes(selected)&&contentWells.includes(selected)?"Alta":"Media"):"Baja";
 $("well").value=selected||"NO DETECTADO";
 const contradictions=wells.filter(x=>x!==selected);
 const labels=extractFields(fullText);
 const comps=extractComponents(fullText);
 const conflict=contradictions.length>0;
 $("failure").value=labels.failure||inferFailure(fullText);
 $("diagnosis").value=buildDiagnosis(labels,comps,selected);
 $("evidence").innerHTML=`<div class="item ${conflict?'warning':'notice'}"><b>Identificación</b><br>Pozo: ${esc(selected||"No detectado")} · Confianza: ${confidence}<br>Hojas: ${sheets.length}<br>${conflict?`⚠️ También aparecen: ${contradictions.map(esc).join(", ")}. Verificar antes de guardar.`:"✓ No se detectó conflicto de pozo."}</div>
 <div class="item"><b>Componentes/seriales detectados</b><br>${comps.length?comps.map(c=>`<span class="tag">${esc(c.type)}: ${esc(c.serial)}</span>`).join(""):"No se detectaron seriales con los patrones actuales."}</div>`;
 $("reportResult").innerHTML=`<div class="item"><b>Reporte procesado</b><div class="muted">${esc(f.name)} · ${sheets.length} hoja(s)</div>${sheets.map(s=>`<span class="tag">${esc(s.name)}</span>`).join("")}</div>`;
 analysis={file:f.name,sheets,wells,selected,confidence,contradictions,labels,comps,photos:[...$("photos").files].map(x=>x.name)};
};

$("saveService").onclick=()=>{
 if(!analysis){alert("Primero analiza el reporte.");return}
 if(!analysis.selected||analysis.contradictions.length){alert("No se puede guardar todavía: verifica el pozo y las inconsistencias detectadas.");return}
 db.services.unshift({id:Date.now(),well:analysis.selected,date:new Date().toISOString().slice(0,10),file:analysis.file,photos:analysis.photos,failure:$("failure").value,diagnosis:$("diagnosis").value,components:analysis.comps});
 analysis.comps.forEach(c=>db.components.unshift({id:Date.now()+Math.random(),well:analysis.selected,type:c.type,serial:c.serial,status:"Detectado en reporte",obs:"Extraído automáticamente; verificar contra fotografía/reporte."}));
 save();alert("Análisis guardado correctamente."); 
};

$("componentForm").onsubmit=e=>{
 e.preventDefault();db.components.unshift({id:Date.now(),well:$("cwell").value,type:$("ctype").value,serial:$("serial").value,status:$("status").value,obs:$("obs").value});save();e.target.reset();alert("Componente registrado.");
};

function extractWells(t){
 const s=String(t).toUpperCase();
 const m=s.match(/\bV[\s\-–_]*\d{2,4}\b/g)||[];
 return [...new Set(m.map(x=>x.replace(/[\s–_]+/g,"-")))];
}
function chooseWell(fileWells,contentWells){
 if(fileWells.length===1 && contentWells.includes(fileWells[0])) return fileWells[0];
 if(contentWells.length===1) return contentWells[0];
 if(fileWells.length===1 && !contentWells.length) return fileWells[0];
 return fileWells.find(x=>contentWells.includes(x))||contentWells[0]||fileWells[0]||"";
}
function extractFields(t){
 const s=String(t);
 const find=(patterns)=>{for(const p of patterns){const m=s.match(p);if(m)return m[1].trim()}return""};
 return {
 failure:find([/(?:MOTIVO|CAUSA|FALLA|PROBLEMA|DESCRIPCI[ÓO]N DEL SERVICIO)\s*[:\-]\s*([^\n|]+)/i]),
 diagnosis:find([/(?:DIAGN[ÓO]STICO|CONCLUSI[ÓO]N|OBSERVACIONES)\s*[:\-]\s*([^\n|]+)/i])
 };
}
function extractComponents(t){
 const out=[]; const s=String(t);
 const patterns=[
  ["Estator",/(?:ESTATOR)[^A-Z0-9]{0,30}(?:SERIAL|S\/N|N°|NO)?[^A-Z0-9]{0,20}([A-Z0-9][A-Z0-9\-\/]{5,})/ig],
  ["Rotor",/(?:ROTOR)[^A-Z0-9]{0,30}(?:SERIAL|S\/N|N°|NO)?[^A-Z0-9]{0,20}([A-Z0-9][A-Z0-9\-\/]{5,})/ig],
  ["Motor",/(?:MOTOR)[^A-Z0-9]{0,30}(?:SERIAL|S\/N|N°|NO)?[^A-Z0-9]{0,20}([A-Z0-9][A-Z0-9\-\/]{5,})/ig],
  ["Cabezal",/(?:CABEZAL|HEAD)[^A-Z0-9]{0,30}(?:SERIAL|S\/N|N°|NO)?[^A-Z0-9]{0,20}([A-Z0-9][A-Z0-9\-\/]{5,})/ig]
 ];
 for(const [type,re] of patterns){let m;while((m=re.exec(s))&&out.length<40){let serial=m[1].replace(/[.,;:]$/,"");if(!/^(SERIAL|MODELO|TIPO|OPERATIVO)$/i.test(serial))out.push({type,serial})}}
 return [...new Map(out.map(x=>[x.type+"|"+x.serial,x])).values()];
}
function inferFailure(t){
 const s=String(t).toLowerCase();
 const keys=["bajo torque","varilla desconectada","varilla partida","rotor partido","fuga por coupling","tubería desconectada","run life","cambio de bomba","optimización","workover"];
 return keys.find(k=>s.includes(k))||"No identificado automáticamente; revisar el reporte.";
}
function buildDiagnosis(labels,comps,well){
 let a=[];
 if(labels.failure)a.push("Hecho extraído: "+labels.failure);
 if(labels.diagnosis)a.push("Observación extraída: "+labels.diagnosis);
 if(comps.length)a.push("Se detectaron "+comps.length+" identificadores de componentes; validar cada serial con el reporte/fotografía.");
 a.push("El diagnóstico automático es preliminar y debe verificarse contra la evidencia del servicio.");
 return a.join("\n");
}
function save(){localStorage.setItem(KEY,JSON.stringify(db));render()}
function render(){
 $("wellCount").textContent=new Set(db.services.map(x=>x.well)).size;
 $("serviceCount").textContent=db.services.length;
 $("componentCount").textContent=db.components.length;
 $("historyList").innerHTML=[...db.services.map(s=>`<div class="item"><b>${esc(s.well)}</b><div class="muted">${esc(s.date)} · ${esc(s.file)}<br>${esc(s.failure||"Sin falla")}<br>${esc(s.diagnosis||"")}</div></div>`),...db.components.map(c=>`<div class="item"><b>${esc(c.well)} · ${esc(c.type)}</b><div class="muted">Serial: ${esc(c.serial)} · ${esc(c.status)}<br>${esc(c.obs||"")}</div></div>`)].join("")||'<div class="box">Aún no hay registros.</div>';
}
function esc(v){return String(v??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]))}
render();