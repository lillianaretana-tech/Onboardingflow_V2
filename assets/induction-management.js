(function(){'use strict';
let profile,sessions=[];
const client=()=>OnboardAuth.client;
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const statusLabel={scheduled:'Programada',registration_open:'Registro abierto',in_progress:'En curso',closed:'Cerrada',cancelled:'Cancelada'};
const moduleStatusLabel={pending:'Pendiente',in_progress:'En curso',completed:'Completado',approved:'Aprobado',interrupted:'Interrumpido',must_repeat:'Debe repetir',failed:'No aprobado'};
const moduleStatusColor={pending:'var(--warning)',in_progress:'var(--warning)',completed:'var(--success)',approved:'var(--success)',interrupted:'var(--danger)',must_repeat:'var(--danger)',failed:'var(--danger)'};
const attendanceLabel={registered:'Registrado',waiting_room:'En sala de espera',present:'Presente',late:'Tardío',absent:'Ausente',incomplete:'Incompleto'};
const attendanceColor={registered:'var(--warning)',waiting_room:'var(--warning)',present:'var(--success)',late:'var(--success)',absent:'var(--danger)',incomplete:'var(--danger)'};

function form(s={}){
  const d=document.createElement('div');d.className='modal-backdrop';
  d.innerHTML=`<form class="modal"><h2>${s.id?'Editar':'Nueva'} inducción</h2><div class="form-grid"><div class="field"><label class="label">Nombre</label><input class="input" name="title" value="${esc(s.title||'Inducción de primer ingreso')}" required></div><div class="field"><label class="label">Proyecto (opcional)</label><input class="input" name="project_name" value="${esc(s.project_name)}"></div><div class="field"><label class="label">Fecha</label><input class="input" name="session_date" type="date" value="${esc(s.session_date)}" required></div><div class="field"><label class="label">Hora</label><input class="input" name="start_time" type="time" value="${esc((s.start_time||'').slice(0,5))}" required></div><div class="field"><label class="label">Enlace de Teams</label><input class="input" name="meeting_url" type="url" value="${esc(s.meeting_url)}" required></div><div class="field"><label class="label">Cupo</label><input class="input" name="capacity" type="number" min="1" value="${esc(s.capacity||30)}" required></div></div><div class="field"><label class="label">Notas internas</label><textarea class="input" name="notes">${esc(s.notes)}</textarea></div><div class="actions"><button class="btn primary">Guardar</button><button type="button" class="btn secondary" data-close>Cerrar</button></div></form>`;
  document.body.append(d);
  d.querySelector('[data-close]').onclick=()=>d.remove();
  d.querySelector('form').addEventListener('submit',async e=>{
    e.preventDefault();
    const v=Object.fromEntries(new FormData(e.target));
    v.capacity=Number(v.capacity);v.tenant_id=profile.tenant_id;v.created_by=s.created_by||profile.id;v.status=s.status||'scheduled';
    const btn=d.querySelector('button.btn.primary');btn.disabled=true;btn.textContent='Guardando...';
    const q=s.id?client().from('of_sessions').update(v).eq('id',s.id):client().from('of_sessions').insert(v);
    const{error}=await q;
    btn.disabled=false;btn.textContent='Guardar';
    if(error)return alert('No fue posible guardar: '+error.message);
    d.remove();await load();
  });
}

async function setSession(id,status,btn){
  if(btn){btn.disabled=true;btn.dataset.originalText=btn.textContent;btn.textContent='Procesando...'}
  const{error}=await client().rpc('of_hr_set_session_status',{p_session_id:id,p_status:status});
  if(btn){btn.disabled=false;btn.textContent=btn.dataset.originalText}
  if(error){alert('No fue posible cambiar la inducción: '+error.message);return false}
  return true;
}

async function setModule(sessionId,moduleId,status,btn){
  if(btn){btn.disabled=true;btn.dataset.originalText=btn.textContent;btn.textContent='...'}
  const{error}=await client().rpc('of_hr_set_session_module',{p_session_id:sessionId,p_module_id:moduleId,p_status:status});
  if(btn){btn.disabled=false;btn.textContent=btn.dataset.originalText}
  if(error){alert('No fue posible actualizar el módulo: '+error.message);return false}
  return true;
}

async function waitingRoomAction(sessionCandidateId,action,description,btn){
  if(btn){btn.disabled=true;btn.dataset.originalText=btn.textContent;btn.textContent='...'}
  const{error}=await client().rpc('of_supervisor_waiting_room_action',{p_session_candidate_id:sessionCandidateId,p_action:action,p_description:description||null});
  if(btn){btn.disabled=false;btn.textContent=btn.dataset.originalText}
  if(error){alert('No fue posible actualizar la asistencia: '+error.message);return false}
  return true;
}

async function manage(id,existing){
  existing?.remove();
  const session=sessions.find(x=>x.id===id);
  const[mods,progress,people]=await Promise.all([
    client().from('of_modules').select('id,name,module_order,medical_only').eq('active',true).order('module_order'),
    client().from('of_module_progress').select('module_id,status').eq('session_id',id),
    client().from('of_session_candidates').select('id,attendance_status,joined_at,candidate:of_candidates(person:of_people(full_name,document_id))').eq('session_id',id)
  ]);
  if(mods.error||progress.error||people.error)return alert((mods.error||progress.error||people.error).message);
  const progressMap=new Map((progress.data||[]).map(p=>[p.module_id,p.status]));

  const d=document.createElement('div');d.className='modal-backdrop';
  d.innerHTML=`<div class="modal">
    <div class="management-head"><div><h2>${esc(session.title)}</h2><p data-session-meta>${esc(session.session_date)} · ${esc(session.start_time)} · <strong>${esc(statusLabel[session.status]||session.status)}</strong></p></div><button class="btn secondary" data-close>Cerrar</button></div>
    <div class="actions" style="justify-content:flex-start">
      <button class="btn primary" data-session-start>Iniciar inducción</button>
      <button class="btn secondary" data-session-close>Finalizar inducción</button>
      <button class="btn secondary" data-session-cancel>Cancelar inducción</button>
      ${session.meeting_url?`<a class="btn secondary" href="${esc(session.meeting_url)}" target="_blank" rel="noopener">Abrir Teams</a>`:''}
    </div>
    <h3>Avance de módulos</h3>
    <div class="people">${(mods.data||[]).map(m=>{
      const st=progressMap.get(m.id)||'pending';
      return `<article class="person-card" data-module-card="${m.id}">
        <strong>${esc(m.name)}</strong>${m.medical_only?'<small> · Solo proyecto médico</small>':''}
        <div><span class="pill" style="background:${moduleStatusColor[st]||'var(--warning)'};color:#fff" data-module-status="${m.id}">${moduleStatusLabel[st]||st}</span></div>
        <div class="actions" style="justify-content:flex-start">
          <button class="btn primary" data-module-start="${m.id}">Iniciar</button>
          <button class="btn secondary" data-module-complete="${m.id}">Completar</button>
          <button class="btn secondary" data-module-repeat="${m.id}">Repetir</button>
        </div>
      </article>`;
    }).join('')||'<p>No hay módulos activos configurados.</p>'}</div>
    <h3>Participantes</h3>
    <p class="hint" style="color:var(--muted);font-size:13px">Autorice el ingreso o marque ausente aquí mismo, sin cambiar de pantalla.</p>
    <div class="table-wrap"><table><thead><tr><th>Nombre</th><th>Identificación</th><th>Asistencia</th><th>Ingreso</th><th>Acciones</th></tr></thead><tbody>${(people.data||[]).map(pcand=>{
      const pending=['registered','waiting_room'].includes(pcand.attendance_status);
      return `<tr data-participant-row="${pcand.id}">
        <td>${esc(pcand.candidate?.person?.full_name)}</td>
        <td>${esc(pcand.candidate?.person?.document_id)}</td>
        <td><span class="pill" style="background:${attendanceColor[pcand.attendance_status]||'var(--warning)'};color:#fff" data-participant-status="${pcand.id}">${attendanceLabel[pcand.attendance_status]||pcand.attendance_status}</span></td>
        <td>${esc(pcand.joined_at?new Date(pcand.joined_at).toLocaleString('es-CR'):'')}</td>
        <td>${pending?`<button class="btn primary" data-admit="${pcand.id}">Autorizar ingreso</button> <button class="btn secondary" data-absent="${pcand.id}">Marcar ausente</button> `:''}<button class="btn secondary" data-incident="${pcand.id}">Incidencia</button></td>
      </tr>`;
    }).join('')||'<tr><td colspan="5">Sin participantes.</td></tr>'}</tbody></table></div>
  </div>`;
  document.body.append(d);

  d.querySelector('[data-close]').onclick=()=>d.remove();

  d.querySelector('[data-session-start]').addEventListener('click',async ev=>{
    const ok=await setSession(id,'in_progress',ev.currentTarget);
    if(ok)manage(id,d);
  });
  d.querySelector('[data-session-close]').addEventListener('click',async ev=>{
    if(!confirm('¿Finalizar esta inducción?'))return;
    const ok=await setSession(id,'closed',ev.currentTarget);
    if(ok)manage(id,d);
  });
  d.querySelector('[data-session-cancel]').addEventListener('click',async ev=>{
    if(!confirm('¿Confirma que desea cancelar esta inducción y cerrar las invitaciones asociadas?'))return;
    const ok=await setSession(id,'cancelled',ev.currentTarget);
    if(ok)manage(id,d);
  });

  d.querySelectorAll('[data-module-start]').forEach(b=>b.addEventListener('click',async ev=>{
    const mid=ev.currentTarget.dataset.moduleStart;
    const ok=await setModule(id,mid,'in_progress',ev.currentTarget);
    if(ok)updateModulePill(d,mid,'in_progress');
  }));
  d.querySelectorAll('[data-module-complete]').forEach(b=>b.addEventListener('click',async ev=>{
    const mid=ev.currentTarget.dataset.moduleComplete;
    const ok=await setModule(id,mid,'completed',ev.currentTarget);
    if(ok)updateModulePill(d,mid,'completed');
  }));
  d.querySelectorAll('[data-module-repeat]').forEach(b=>b.addEventListener('click',async ev=>{
    const mid=ev.currentTarget.dataset.moduleRepeat;
    if(!confirm('¿Marcar este módulo para repetir?'))return;
    const ok=await setModule(id,mid,'must_repeat',ev.currentTarget);
    if(ok)updateModulePill(d,mid,'must_repeat');
  }));

  d.querySelectorAll('[data-admit]').forEach(b=>b.addEventListener('click',async ev=>{
    const scid=ev.currentTarget.dataset.admit;
    const ok=await waitingRoomAction(scid,'admit',null,ev.currentTarget);
    if(ok)manage(id,d);
  }));
  d.querySelectorAll('[data-absent]').forEach(b=>b.addEventListener('click',async ev=>{
    const scid=ev.currentTarget.dataset.absent;
    if(!confirm('¿Marcar a esta persona como ausente?'))return;
    const ok=await waitingRoomAction(scid,'absent',null,ev.currentTarget);
    if(ok)manage(id,d);
  }));
  d.querySelectorAll('[data-incident]').forEach(b=>b.addEventListener('click',async ev=>{
    const scid=ev.currentTarget.dataset.incident;
    const desc=prompt('Describa brevemente la incidencia:');
    if(desc===null)return;
    const ok=await waitingRoomAction(scid,'incident',desc,ev.currentTarget);
    if(ok)alert('Incidencia registrada.');
  }));
}

function updateModulePill(modal,moduleId,status){
  const pill=modal.querySelector(`[data-module-status="${moduleId}"]`);
  if(!pill)return;
  pill.style.background=moduleStatusColor[status]||'var(--warning)';
  pill.style.color='#fff';
  pill.textContent=moduleStatusLabel[status]||status;
}

async function load(){
  profile=profile||await OnboardAuth.getProfile();
  const{data,error}=await client().from('of_sessions').select('*,participants:of_session_candidates(count)').order('session_date',{ascending:false});
  if(error)return console.error(error);
  sessions=data||[];window.__ofSessions=sessions;render();
}

function render(){
  const root=document.getElementById('induction');
  if(!root)return;
  root.innerHTML=`<div class="grid grid-4">
    <article class="card metric"><strong>${sessions.filter(s=>['scheduled','registration_open'].includes(s.status)).length}</strong><span>Próximas</span></article>
    <article class="card metric"><strong>${sessions.reduce((n,s)=>n+(s.participants?.[0]?.count||0),0)}</strong><span>Asignados</span></article>
    <article class="card metric"><strong>${sessions.filter(s=>s.status==='in_progress').length}</strong><span>En curso</span></article>
    <article class="card metric"><strong>${sessions.filter(s=>s.status==='closed').length}</strong><span>Cerradas</span></article>
  </div>
  <section class="card section">
    <div class="management-head"><div><h2>Calendario de inducciones</h2><p>Inicie la sesión y avance cada módulo desde aquí.</p></div><button class="btn primary" data-new>Nueva inducción</button></div>
    <div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Hora</th><th>Nombre</th><th>Asignados</th><th>Estado</th><th>Acciones</th></tr></thead><tbody>${sessions.map(s=>`<tr><td>${esc(s.session_date)}</td><td>${esc(s.start_time)}</td><td>${esc(s.title)}</td><td>${s.participants?.[0]?.count||0}/${s.capacity||30}</td><td>${esc(statusLabel[s.status]||s.status)}</td><td><button class="btn primary" data-manage="${s.id}">Controlar</button> <button class="btn secondary" data-edit="${s.id}">Editar</button></td></tr>`).join('')||'<tr><td colspan="6">No hay inducciones.</td></tr>'}</tbody></table></div>
  </section>`;
  root.querySelector('[data-new]').onclick=()=>form();
  root.querySelectorAll('[data-manage]').forEach(b=>b.onclick=()=>manage(b.dataset.manage));
  root.querySelectorAll('[data-edit]').forEach(b=>b.onclick=()=>form(sessions.find(s=>s.id===b.dataset.edit)));
}

document.addEventListener('DOMContentLoaded',load);
})();
