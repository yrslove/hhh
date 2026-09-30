const state={
  data:null, functionMap:new Map(), current:null, layout:null, lastCheck:null,
  settings:null, dragged:null, simpleCallback:null, actionTimer:null, actionPending:false, actionTitle:"", cancelPending:false, resume:null
};
const $=s=>document.querySelector(s);
const $$=s=>[...document.querySelectorAll(s)];

function esc(v){return String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[c]));}
function toast(msg,kind="info",timeout=3300){
  const x=document.createElement("div");x.className=`toast ${kind}`;x.textContent=msg;
  $("#toastStack").appendChild(x);setTimeout(()=>x.remove(),timeout);
}
async function api(url,opt={}){
  const r=await fetch(url,opt);let d={};try{d=await r.json()}catch{}
  if(!r.ok)throw new Error(d.detail||d.message||`HTTP ${r.status}`);
  return d;
}
function showView(name){
  $$(".view").forEach(x=>x.classList.remove("is-visible"));$$(".nav").forEach(x=>x.classList.remove("is-active"));
  $(`#view-${name}`).classList.add("is-visible");$(`.nav[data-view="${name}"]`).classList.add("is-active");
  if(name==="logs")loadLogs();
}
$$(".nav").forEach(b=>b.addEventListener("click",()=>showView(b.dataset.view)));

function renderList(el,items){el.innerHTML="";(items||[]).forEach(x=>{const li=document.createElement("li");li.textContent=x;el.appendChild(li);});}

async function loadFunctions(){
  state.data=await api("/api/functions");
  state.layout=await api("/api/layout");
  state.functionMap.clear();
  state.data.groups.forEach(g=>g.functions.forEach(fn=>{fn.group_title=g.title;state.functionMap.set(fn.id,fn)}));
  renderGroups();renderQuick();
}
function renderQuick(){
  const ids=["procedures_create","procedures_batch_sign","procedure_counts","important_000","check_frm_before_etft","check_etft_overlaps"];
  const box=$("#quickFunctions");box.innerHTML="";
  ids.forEach(id=>{
    const fn=state.functionMap.get(id);if(!fn)return;
    const b=document.createElement("button");b.className="quick-card";b.dataset.functionId=id;
    b.innerHTML=`<strong>${esc(fn.title)}</strong><span>${esc(fn.note||fn.short||"")}</span>`;
    b.addEventListener("click",()=>openFunction(id));box.appendChild(b);
  });
}
function renderGroups(){
  const host=$("#groupList");host.innerHTML="";
  state.data.groups.forEach(group=>{
    const wrap=document.createElement("section");wrap.className="group";wrap.dataset.groupId=group.id;
    wrap.innerHTML=`<div class="group-head">
      <div class="group-title"><strong>${esc(group.title)}</strong><span class="muted">${group.functions.length} функц.</span></div>
      <div class="group-tools"><button class="btn btn-secondary btn-small rename-group">Перейменувати</button></div>
    </div><div class="group-body"></div>`;
    const body=wrap.querySelector(".group-body");
    body.addEventListener("dragover",e=>{e.preventDefault();body.classList.add("drop-target");handleDragOver(body,e)});
    body.addEventListener("dragleave",()=>body.classList.remove("drop-target"));
    body.addEventListener("drop",e=>{e.preventDefault();body.classList.remove("drop-target");dropCard(body,e)});
    wrap.querySelector(".rename-group").addEventListener("click",()=>promptSimple("Перейменувати групу","Нова назва",group.title,async value=>{
      await api(`/api/groups/${encodeURIComponent(group.id)}/rename`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({title:value})});
      toast("Групу перейменовано","success");await loadFunctions();
    }));
    group.functions.forEach(fn=>body.appendChild(makeFunctionCard(fn)));
    host.appendChild(wrap);
  });
}
function makeFunctionCard(fn){
  const row=document.createElement("div");row.className="function-card";row.draggable=true;row.dataset.functionId=fn.id;
  row.innerHTML=`<div class="drag-handle" title="Перетягнути">⋮⋮</div>
    <div><div class="function-title">${esc(fn.title)}</div>${fn.note?`<div class="function-note">${esc(fn.note)}</div>`:""}</div>
    <div class="function-short">${esc(fn.short||"")}</div>
    <button class="btn btn-secondary btn-small" data-function-id="${esc(fn.id)}">Відкрити</button>`;
  row.querySelector("button").addEventListener("click",()=>openFunction(fn.id));
  row.addEventListener("dragstart",e=>{state.dragged=row;row.classList.add("dragging");e.dataTransfer.effectAllowed="move";});
  row.addEventListener("dragend",()=>{row.classList.remove("dragging");state.dragged=null;$$(".drop-target").forEach(x=>x.classList.remove("drop-target"));});
  return row;
}
function handleDragOver(container,e){
  const after=getDragAfter(container,e.clientY);
  if(!state.dragged)return;
  if(after==null)container.appendChild(state.dragged);else container.insertBefore(state.dragged,after);
}
function getDragAfter(container,y){
  const els=[...container.querySelectorAll(".function-card:not(.dragging)")];
  return els.reduce((closest,child)=>{
    const box=child.getBoundingClientRect(),offset=y-box.top-box.height/2;
    return offset<0&&offset>closest.offset?{offset,element:child}:closest;
  },{offset:Number.NEGATIVE_INFINITY,element:null}).element;
}
async function dropCard(container,e){
  if(!state.dragged)return;
  if(!container.contains(state.dragged))container.appendChild(state.dragged);
  await saveDomLayout();toast("Порядок функцій збережено","success");
}
async function saveDomLayout(){
  const layout=state.layout||await api("/api/layout");
  const groups=$$(".group").map((g,i)=>({id:g.dataset.groupId,title:g.querySelector(".group-title strong").textContent,order:i}));
  const funcs={...(layout.functions||{})};
  groups.forEach(g=>{
    const body=$(`.group[data-group-id="${CSS.escape(g.id)}"] .group-body`);
    [...body.querySelectorAll(".function-card")].forEach((card,i)=>{
      const fid=card.dataset.functionId;const old=funcs[fid]||{};
      funcs[fid]={title:old.title||state.functionMap.get(fid)?.title||"",note:old.note||state.functionMap.get(fid)?.note||"",group_id:g.id,order:i};
    });
  });
  state.layout={groups,functions:funcs};
  const r=await api("/api/layout",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(state.layout)});
  state.layout=r.layout;
}

function paramInput(p){
  const wrap=document.createElement("div");wrap.className="param-row";
  const label=document.createElement("label");label.textContent=p.label+(p.required?" *":"");wrap.appendChild(label);
  let input;
  if(p.type==="checkbox"){
    input=document.createElement("input");input.type="checkbox";input.checked=!!p.default;
  }else{
    input=document.createElement("input");input.type=["password","month"].includes(p.type)?p.type:"text";input.value=p.default??"";
  }
  input.dataset.paramName=p.name;wrap.appendChild(input);return wrap;
}
function openFunction(id){
  const fn=state.functionMap.get(id);if(!fn)return;state.current=fn;
  state.resume=null;$("#continueFunction").hidden=true;
  $("#drawerGroup").textContent=fn.group_title||"";$("#drawerId").textContent=`ID: ${fn.id}`;
  $("#drawerTitle").textContent=fn.title;$("#drawerShort").textContent=fn.short||"";
  renderList($("#drawerSequence"),fn.sequence);renderList($("#drawerFiles"),fn.files);renderList($("#drawerAccess"),fn.access);renderList($("#drawerManual"),fn.manual);
  const note=fn.note||"";$("#drawerNoteBox").hidden=!note;$("#drawerNote").textContent=note;
  const params=$("#drawerParams");params.innerHTML="";(fn.params||[]).forEach(p=>params.appendChild(paramInput(p)));
  $("#customTitle").value=fn.title;$("#customNote").value=note;$("#noteCounter").textContent=`${note.length}/300`;
  const sel=$("#customGroup");sel.innerHTML="";
  state.data.groups.forEach(g=>{const o=document.createElement("option");o.value=g.id;o.textContent=g.title;if(fn.group_id===g.id)o.selected=true;sel.appendChild(o);});
  $("#actionResult").hidden=true;$("#actionResult").className="result-box";
  $("#functionDrawer").classList.add("is-open");$("#functionDrawer").setAttribute("aria-hidden","false");$("#backdrop").hidden=false;
  toast(`Відкрито функцію: ${fn.title}`);
}
function closeDrawer(){
  $("#functionDrawer").classList.remove("is-open");$("#functionDrawer").setAttribute("aria-hidden","true");$("#backdrop").hidden=true;
}
$("#closeDrawer").onclick=closeDrawer;$("#closeDrawer2").onclick=closeDrawer;$("#backdrop").onclick=closeDrawer;
$("#customNote").addEventListener("input",e=>$("#noteCounter").textContent=`${e.target.value.length}/300`);

$("#saveCustomization").addEventListener("click",async()=>{
  if(!state.current)return;
  try{
    await api(`/api/functions/${encodeURIComponent(state.current.id)}/customize`,{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({title:$("#customTitle").value,note:$("#customNote").value,group_id:$("#customGroup").value})});
    toast("Назву, примітку та групу збережено","success");await loadFunctions();openFunction(state.current.id);
  }catch(e){toast(e.message,"error",5000)}
});

function gatherParams(){
  const out={};
  $("#drawerParams").querySelectorAll("[data-param-name]").forEach(x=>{out[x.dataset.paramName]=x.type==="checkbox"?x.checked:x.value});
  return out;
}
function renderActionStatus(actions=[]){
  const box=$("#actionStatus"),action=actions[0];
  if(!action){
    state.cancelPending=false;$("#cancelActions").disabled=false;$("#cancelActions").textContent="Зупинити";
    if(state.actionPending){
      box.hidden=false;$("#actionStatusText").textContent=`Запуск: ${state.actionTitle||"функції"}`;$("#actionProgress").value=2;
    }else box.hidden=true;
    return;
  }
  box.hidden=false;
  const title=state.functionMap.get(action.function_id)?.title||action.function_id;
  const stopButton=$("#cancelActions");
  if(action.cancel_requested){state.cancelPending=true;stopButton.disabled=true;stopButton.textContent="Зупиняємо…";}
  else if(!state.cancelPending){stopButton.disabled=false;stopButton.textContent="Зупинити";}
  $("#actionStatusText").textContent=action.cancel_requested
    ? `${title}: Зупиняємо…`
    : `${title}: ${action.detail||"виконується"} (${action.progress||0}%)`;
  $("#actionProgress").value=action.progress||0;
}
async function pollActionStatus(){
  try{
    const r=await api("/api/actions/status");renderActionStatus(r.actions||[]);
    if(!state.actionPending&&!(r.actions||[]).length&&state.actionTimer){clearInterval(state.actionTimer);state.actionTimer=null;}
  }catch(e){/* A transient status-poll error must not interrupt the action itself. */}
}
function startActionPolling(title){
  state.actionPending=true;state.actionTitle=title;renderActionStatus();
  if(!state.actionTimer)state.actionTimer=setInterval(pollActionStatus,600);
  pollActionStatus();
}
function finishActionPolling(){
  state.actionPending=false;pollActionStatus();
}
$("#cancelActions").addEventListener("click",async()=>{
  const b=$("#cancelActions");b.disabled=true;b.textContent="Зупиняємо…";
  try{const r=await api("/api/actions/cancel",{method:"POST"});state.cancelPending=r.count>0;toast(r.count?"Зупиняємо виконання…":"Активних функцій немає.","warn",2500);await pollActionStatus();}
  catch(e){state.cancelPending=false;toast(e.message,"error",4200);b.disabled=false;b.textContent="Зупинити";}
});
async function executeCurrent(resume=false){
  const fn=state.current;if(!fn)return;
  const runBtn=$("#runFunction"),continueBtn=$("#continueFunction");runBtn.disabled=true;continueBtn.disabled=true;continueBtn.hidden=true;
  const resultEl=$("#actionResult");resultEl.hidden=false;resultEl.className="result-box";resultEl.textContent="Виконується...";
  toast(`${resume?"Продовження":"Запущено"}: ${fn.title}`);
  startActionPolling(fn.title);
  try{
    const payload=gatherParams();
    if(resume){
      if(!state.resume||state.resume.functionId!==fn.id)throw new Error("Контрольна точка для продовження не знайдена.");
      payload._resume_from_checkpoint=true;payload._resume_key=state.resume.key;
    }
    const result=await api(`/api/functions/${encodeURIComponent(fn.id)}/execute`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({payload})});
    const summary={message:result.message,status:result.status,duration_ms:result.duration_ms,data:result.data};
    const humanLog=(result.data?.human_log||"").trim();
    resultEl.textContent=result.ok&&result.status==="done"&&humanLog?humanLog:(result.message||JSON.stringify(summary,null,2));
    if(!result.ok){resultEl.classList.add("error");toast(result.message||"Помилка","error",5200)}
    else if(result.status==="waiting_user"){toast(result.message||"Потрібна дія користувача","warn",5200)}
    else toast(`Виконано: ${fn.title}`,"success");
    if(result.data && (result.data.procedure_count_issues||result.data.errors||result.data.order_errors||result.data.overlap_errors||result.data.ar3_monoservice)){
      state.lastCheck={function:fn,result};renderResults();showView("results");
    }
    const canContinue=result.status==="waiting_user"&&result.data?.can_continue&&result.data?.resume_key;
    state.resume=canContinue?{functionId:fn.id,key:result.data.resume_key}:null;
    continueBtn.hidden=!canContinue;
    await loadLogs();
  }catch(e){state.resume=null;continueBtn.hidden=true;resultEl.classList.add("error");resultEl.textContent=e.message;toast(e.message,"error",5200)}
  finally{runBtn.disabled=false;continueBtn.disabled=false;finishActionPolling()}
}
$("#runFunction").addEventListener("click",()=>executeCurrent(false));
$("#continueFunction").addEventListener("click",()=>executeCurrent(true));

function displayDate(value){
  const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(value||"");
  return m?`${m[3]}.${m[2]}.${m[1]}`:(value||"");
}
function scheduleIssue(e){
  const x=document.createElement("div");x.className="issue";
  const patient=`${e.patient||""}, ${displayDate(e.date)}`;
  if(e.frm&&e.etft){
    const detail=`${e.frm.cell||""} : ${e.etft.cell||""} | ${e.frm.specialist||""} — ${e.etft.specialist||""}`;
    x.innerHTML=`<strong>${esc(e.type)}</strong><p>${esc(patient)}<br>ФРМ: ${esc(e.frm.interval||"")} · ЕТ/ФТ: ${esc(e.etft.interval||"")}<strong class="issue-detail">${esc(detail)}</strong></p>`;
  }else if(e.first&&e.second){
    const detail=`${e.first.cell||""} : ${e.second.cell||""} | ${e.first.specialist||""} — ${e.second.specialist||""}`;
    x.innerHTML=`<strong>${esc(e.type)}</strong><p>${esc(patient)}<br>Перетин: ${esc(e.first.interval||"")} і ${esc(e.second.interval||"")}<strong class="issue-detail">${esc(detail)}</strong></p>`;
  }else{
    x.innerHTML=`<strong>${esc(e.type||"Помилка")}</strong><p>${esc(e.message||JSON.stringify(e))}</p>`;
  }
  return x;
}
function issueDateKey(value){
  const raw=String(value??"");
  const iso=/^(\d{4})[-/.](\d{2})[-/.](\d{2})$/.exec(raw);
  if(iso)return `${iso[1]}-${iso[2]}-${iso[3]}`;
  const local=/^(\d{2})[./-](\d{2})[./-](\d{4})$/.exec(raw);
  return local?`${local[3]}-${local[2]}-${local[1]}`:raw;
}
function issueMonthKey(value){
  const key=issueDateKey(value),m=/^(\d{4})-(\d{2})-/.exec(key);
  return m?`${m[1]}-${m[2]}`:"";
}
function appendErrorSection(host,title,items){
  if(!items.length)return;
  const heading=document.createElement("h3");heading.className="result-section-title";heading.textContent=`${title}: ${items.length}`;host.appendChild(heading);
  items.forEach(e=>host.appendChild(scheduleIssue(e)));
}
function appendMonthlyErrorSection(host,title,items){
  if(!items.length)return;
  const heading=document.createElement("h3");heading.className="result-section-title";heading.textContent=`${title}: ${items.length}`;host.appendChild(heading);
  const ordered=[...items].sort((a,b)=>issueDateKey(a.date).localeCompare(issueDateKey(b.date)));
  const groups=new Map();
  ordered.forEach(e=>{
    const month=issueMonthKey(e.date)||"unknown";
    if(!groups.has(month))groups.set(month,[]);
    groups.get(month).push(e);
  });
  groups.forEach((monthItems,month)=>{
    const group=document.createElement("section");group.className="result-month-group";group.dataset.month=month;
    const monthHeading=document.createElement("h2");monthHeading.className="result-month-title";
    monthHeading.textContent=month==="unknown"?"МЕСЯЦ НЕ УКАЗАН":`${month.slice(5)} МЕСЯЦ`;
    group.appendChild(monthHeading);
    monthItems.forEach(e=>group.appendChild(scheduleIssue(e)));
    host.appendChild(group);
  });
}
function renderProcedureCountResults(host,data){
  const sum=data.summary||{};
  const scope=sum.month?`за ${sum.month}`:"за весь графік";
  const boxes=[
    ["Перевірено пацієнтів",sum.patients??0],
    ["Пацієнтів із проблемами",sum.patients_with_issues??0],
    ["Порушень",sum.issues??0],
    ["Записів ЕТ/ФТ",sum.etft_records??0],
    ["Записів Псих-Лого",sum.psych_records??0]
  ];
  const summary=document.createElement("div");summary.className="result-summary";
  boxes.forEach(([label,value])=>{
    const box=document.createElement("div");box.className="summary-box";
    box.innerHTML=`<span>${esc(label)}</span><strong>${esc(value)}</strong>`;
    summary.appendChild(box);
  });
  host.appendChild(summary);
  const issues=data.procedure_count_issues||[];
  if(!issues.length){
    const ok=document.createElement("div");ok.className="panel";
    ok.textContent=sum.month&&sum.patients===0?`За ${sum.month} не знайдено пацієнтів для перевірки.`:
      sum.month?`За ${sum.month} порушень кількості процедур не знайдено.`:
      "Порушень кількості процедур не знайдено.";
    host.appendChild(ok);
  }
  const patientKey=item=>`${item.source_sheet||""}!${item.source_cell}`;
  const totals=new Map((data.patient_summaries||[]).map(item=>[patientKey(item),item]));
  const groups=new Map();
  issues.forEach(item=>{
    const key=patientKey(item);
    if(!groups.has(key))groups.set(key,[]);
    groups.get(key).push(item);
  });
  groups.forEach((items,key)=>{
    const patient=items[0].patient||"";
    const patientTotal=totals.get(key);
    const section=document.createElement("section");section.className="schedule-info";
    const heading=document.createElement("h3");
    heading.textContent=`${patient} · ${items[0].source_sheet||""}!${items[0].source_cell} · ${patientTotal?.total??"—"} процедур ${scope} · ${patientTotal?.days??"—"} днів із записами`;
    section.appendChild(heading);
    items.sort((a,b)=>(a.date||"").localeCompare(b.date||""));
    items.forEach(item=>{
      const issue=document.createElement("div");issue.className="issue";
      if(item.kind==="daily_count"){
        const details=(item.locations||[]).map(loc=>
          `<li>${esc(loc.sheet)}!${esc(loc.cell)}${loc.continuation_cell?`–${esc(loc.continuation_cell)}`:""} · ${esc(loc.time)} · ${esc(loc.specialist||"Спеціаліст не вказаний")}</li>`
        ).join("");
        issue.innerHTML=`<strong>${esc(displayDate(item.date))}: ${esc(item.day_total)} процедур за день (потрібно ${esc(item.daily_target??2)})</strong>`+
          `<p>ЕТ/ФТ: ${esc(item.etft_count)} · Псих-Лого: ${esc(item.psych_count)} · Всього ${esc(scope)}: ${esc(item.patient_total)}</p>`+
          (details?`<ul>${details}</ul>`:"<p>Записів на цих вкладках немає.</p>");
      }else{
        const title=item.kind==="course_days"?"Кількість днів":item.kind==="course_total"?"Замало процедур":item.kind==="month_empty"?"Немає процедур за місяць":"Неоднозначний пацієнт";
        issue.innerHTML=`<strong>${esc(title)}</strong><p>${esc(item.message||"")}</p>`;
      }
      section.appendChild(issue);
    });
    host.appendChild(section);
  });
  const warnings=data.structure_warnings||[];
  if(warnings.length){
    const section=document.createElement("section");section.className="schedule-info";
    section.innerHTML=`<h3>Попередження структури графіка: ${esc(warnings.length)}</h3>`;
    warnings.forEach(item=>{
      const line=document.createElement("p");
      line.textContent=`${item.sheet}!${item.cell}: ${item.message}`;
      section.appendChild(line);
    });
    host.appendChild(section);
  }
}
function renderResults(){
  const empty=$("#resultsEmpty"),host=$("#resultsContent");host.innerHTML="";
  if(!state.lastCheck){empty.hidden=false;return}empty.hidden=true;
  if(state.lastCheck.function?.id==="procedure_counts"){
    renderProcedureCountResults(host,state.lastCheck.result.data||{});
    return;
  }
  const d=state.lastCheck.result.data||{},sum=d.summary||{},standalone=d.errors||[];
  const order=d.order_errors||standalone.filter(e=>e.frm&&e.etft);
  const overlaps=d.overlap_errors||standalone.filter(e=>e.first&&e.second);
  const other=standalone.filter(e=>!order.includes(e)&&!overlaps.includes(e));
  const ar3=d.ar3_monoservice||[];
  const boxes=[
    ["Всього помилок",sum.errors??(order.length+overlaps.length+other.length)],
    sum.frm_hourly!==undefined?["Годинні ФРМ",sum.frm_hourly]:null,
    sum.frm_repeat_ignored!==undefined?["Повторних ФРМ — пропущено",sum.frm_repeat_ignored]:null,
    ["Помилки ФРМ → ЕТ/ФТ",sum.order_errors??order.length],
    ["Перетини ЕТ/ФТ",sum.overlap_errors??overlaps.length],
    sum.ar3_monoservice!==undefined?["54 пакет АР3-1 (монопослуга)",sum.ar3_monoservice]:null
  ].filter(Boolean);
  const summary=document.createElement("div");summary.className="result-summary";
  boxes.forEach(([a,b])=>{const x=document.createElement("div");x.className="summary-box";x.innerHTML=`<span>${esc(a)}</span><strong>${esc(b??0)}</strong>`;summary.appendChild(x)});
  host.appendChild(summary);
  appendMonthlyErrorSection(host,"Помилки ФРМ → ЕТ/ФТ",order);
  appendMonthlyErrorSection(host,"Перетини ЕТ/ФТ",overlaps);
  if(other.length)appendErrorSection(host,"Інші помилки",other);
  if(ar3.length){
    const info=document.createElement("section");info.className="schedule-info";
    info.innerHTML=`<h3>54 пакет група АР3-1 (монопослуга): ${esc(ar3.length)}</h3>`;
    ar3.forEach(item=>{
      const row=document.createElement("p");
      row.innerHTML=`<strong>${esc(item.patient||"")}</strong><br>${esc(displayDate(item.first_date))} <code>${esc(item.first_cell||"")}</code> → ${esc(displayDate(item.next_date))} <code>${esc(item.next_cell||"")}</code> · ${esc(item.gap_days??"")} дні`;
      info.appendChild(row);
    });
    host.appendChild(info);
  }
  if(!order.length&&!overlaps.length&&!other.length){
    const ok=document.createElement("div");ok.className="panel";ok.textContent="Помилок не знайдено.";host.appendChild(ok);
  }
}

$("#uploadWorkbook").addEventListener("click",async()=>{
  const f=$("#workbookFile").files[0];if(!f){toast("Спочатку виберіть XLSX/XLSM","warn");return}
  const fd=new FormData();fd.append("file",f);toast("Завантаження Excel...");
  try{
    const r=await api("/api/upload",{method:"POST",body:fd});
    $("#currentWorkbook").textContent=r.path;$("#workbookTop").textContent=f.name;toast("Excel підключено","success");
  }catch(e){toast(e.message,"error",5000)}
});
async function loadCurrentWorkbook(){
  const r=await api("/api/current-workbook");if(r.path){$("#currentWorkbook").textContent=r.path;$("#workbookTop").textContent=r.path.split(/[\\/]/).pop();}
}

async function loadLogs(){
  try{
    const r=await api("/api/logs?limit=400");const body=$("#logsBody");body.innerHTML="";
    r.events.forEach(e=>{
      const tr=document.createElement("tr");
      tr.innerHTML=`<td>${esc((e.created_at||"").replace("T"," ").slice(0,23))}</td><td class="level-${esc(e.level)}">${esc(e.level)}</td>
        <td>${esc(e.action_id||"")}</td><td>${esc(e.event_type||"")}</td><td>${esc(e.message||"")}</td>
        <td><details><summary>відкрити</summary><pre>${esc(JSON.stringify(e.data||{},null,2))}</pre></details></td>`;body.appendChild(tr);
    });
  }catch(e){toast(e.message,"error")}
}
$("#refreshLogs").onclick=loadLogs;

async function loadSettings(){
  state.settings=await api("/api/settings");
  const keys=["graph1_path","coordinators_path","desktop_eleks_path","ehealth_base_url","browser_executable","remote_debug_port","frm_sheet","etft_sheet","distribution_sheet","coordinator_sheet","schedule_year"];
  keys.forEach(k=>{const el=$(`#s_${k}`);if(el)el.value=state.settings[k]??""});
  $("#s_safe_mode").checked=!!state.settings.safe_mode;
  $("#s_specialists").value=JSON.stringify(state.settings.specialists||{},null,2);
}
$("#saveSettings").addEventListener("click",async()=>{
  try{
    let specialists={};try{specialists=JSON.parse($("#s_specialists").value||"{}")}catch{throw new Error("JSON спеціалістів некоректний.")}
    const values={};
    ["graph1_path","coordinators_path","desktop_eleks_path","ehealth_base_url","browser_executable","frm_sheet","etft_sheet","distribution_sheet","coordinator_sheet"].forEach(k=>values[k]=$(`#s_${k}`).value);
    values.remote_debug_port=Number($("#s_remote_debug_port").value||9222);values.schedule_year=Number($("#s_schedule_year").value||2026);
    values.safe_mode=$("#s_safe_mode").checked;values.specialists=specialists;
    const r=await api("/api/settings",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({values})});
    state.settings=r.settings;toast("Налаштування збережено","success");
  }catch(e){toast(e.message,"error",5000)}
});

function promptSimple(title,label,value,callback){
  $("#simpleTitle").textContent=title;$("#simpleLabel").textContent=label;$("#simpleInput").value=value||"";
  state.simpleCallback=callback;$("#simpleModal").hidden=false;$("#simpleInput").focus();
}
$("#simpleCancel").onclick=()=>{$("#simpleModal").hidden=true;state.simpleCallback=null};
$("#simpleOk").onclick=async()=>{
  const cb=state.simpleCallback,val=$("#simpleInput").value.trim();if(!val){toast("Введіть назву","warn");return}
  $("#simpleModal").hidden=true;state.simpleCallback=null;if(cb)await cb(val);
};
$("#createGroup").addEventListener("click",()=>promptSimple("Створити групу","Назва групи","",async value=>{
  try{await api("/api/groups",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({title:value})});toast("Групу створено","success");await loadFunctions()}
  catch(e){toast(e.message,"error")}
}));

(async function init(){
  try{
    await Promise.all([loadFunctions(),loadSettings(),loadCurrentWorkbook(),loadLogs()]);
    await pollActionStatus();
    if(!$("#actionStatus").hidden&&!state.actionTimer)state.actionTimer=setInterval(pollActionStatus,600);
  }
  catch(e){toast(`Помилка запуску UI: ${e.message}`,"error",6000)}
})();
