const KEY="pcp_nexus_v1";let db=JSON.parse(localStorage.getItem(KEY)||'{"services":[],"components":[]}');
const $=id=>document.getElementById(id);
document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>{document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));document.querySelectorAll(".panel").forEach(x=>x.classList.remove("active"));b.classList.add("active");$(b.dataset.tab).classList.add("active");render()});
$("date").value=new Date().toISOString().slice(0,10);
$("serviceForm").onsubmit=e=>{e.preventDefault();let s={id:Date.now(),well:$("well").value,date:$("date").value,type:$("type").value,failure:$("failure").value,diagnosis:$("diagnosis").value,excel:$("excel").files[0]?.name||"",photos:[...$("photos").files].map(f=>f.name)};db.services.unshift(s);save();e.target.reset();$("date").value=new Date().toISOString().slice(0,10);alert("Servicio registrado en PCP NEXUS 1.0");};
$("componentForm").onsubmit=e=>{e.preventDefault();db.components.unshift({id:Date.now(),well:$("cwell").value,type:$("ctype").value,serial:$("serial").value,status:$("status").value,obs:$("obs").value});save();e.target.reset();alert("Componente registrado");};
function save(){localStorage.setItem(KEY,JSON.stringify(db));render()}
function render(){$("wellCount").textContent=new Set(db.services.map(x=>x.well)).size;$("serviceCount").textContent=db.services.length;$("componentCount").textContent=db.components.length;
let h=[...db.services.map(s=>`<div class="item"><b>${esc(s.well)}</b> · ${esc(s.date)}<div class="muted">${esc(s.type||"Servicio")}<br>${esc(s.failure||"Sin falla registrada")}</div></div>`),...db.components.map(c=>`<div class="item"><b>${esc(c.well)}</b> · ${esc(c.type)}</b><div class="muted">Serial: ${esc(c.serial)} · ${esc(c.status)}<br>${esc(c.obs||"")}</div></div>`)].join("");$("historyList").innerHTML=h||'<div class="box">Aún no hay registros.</div>'}
function esc(v){return String(v??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]))}
render();
if("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js");
