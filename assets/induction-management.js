(function(){'use strict';
let profile,sessions=[];
const client=()=>OnboardAuth.client;
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const statusLabel={scheduled:'Programada',registration_open:'Registro abierto',in_progress:'En curso',closed:'Cerrada',cancelled:'Cancelada'};
const moduleStatusLabel={pending:'Pendiente',in_progress:'En curso',completed:'Completado',approved:'Aprobado',interrupted:'Interrumpido',must_repeat:'Debe repetir',failed:'No aprobado'};
const moduleStatusColor={pending:'var(--warning)',in_progress:'var(--warning)',completed:'var(--success)',approved:'var(--success)',interrupted:'var(--danger)',must_repeat:'var(--danger)',failed:'var(--danger)'};
const attendanceLabel={registered:'Registrado',waiting_room:'En sala de espera',present:'Presente',late:'Tardío',absent:'Ausente',incomplete:'Incompleto'};
const attendanceColor={registered:'var(--warning)',waiting_room:'var(--warning)',present:'var(--success)',late:'var(--success)',absent:'var(--danger)',incomplete:'var(--danger)'};
const candidateStatusLabel={pending_hr_review:'Pendiente revisión RH',returned_for_correction:'Devuelto',approved_by_hr:'Aprobado por RH',invited:'Invitado',registered:'Registrado',in_induction:'En inducción',pending_module:'Debe repetir algo',approved_for_hire:'Aprobado para contratar',not_approved:'No aprobado',cancelled:'Cancelado',absent:'Ausente'};

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

async function waitingRoomAction(sessionCandidateId,action,description,btn){
  if(btn){btn.disabled=true;btn.dataset.originalText=btn.textContent;btn.textContent='...'}
  const{error}=await client().rpc('of_supervisor_waiting_room_action',{p_session_candidate_id:sessionCandidateId,p_action:action,p_description:description||null});
  if(btn){btn.disabled=false;btn.textContent=btn.dataset.originalText}
  if(error){alert('No fue posible actualizar la asistencia: '+error.message);return false}
  return true;
}

function moduleForm(sessionCandidateId,mod,currentStatus,currentScore,onDone){
  const d=document.createElement('div');d.className='modal-backdrop';
  const statusOptions=['pending','in_progress','completed','approved','interrupted','must_repeat','failed'];
  d.innerHTML=`<form class="modal"><h2>${esc(mod.name)}</h2>
    <div class="field"><label class="label">Estado</label><select class="select" name="status">${statusOptions.map(s=>`<option value="${s}" ${currentStatus===s?'selected':''}>${moduleStatusLabel[s]}</option>`).join('')}</select></div>
    ${mod.is_exam?`<div class="field"><label class="label">Nota (0-100, mínimo para aprobar: ${mod.passing_score ?? 70})</label><input class="input" name="score" type="number" min="0" max="100" step="0.1" value="${currentScore??''}"></div><p class="hint" style="color:var(--muted);font-size:13px">Si la nota queda por debajo del mínimo y marca "Completado", el sistema lo pasa a "Debe repetir" automáticamente.</p>`:''}
    <div class="actions"><button class="btn primary">Guardar</button><button type="button" class="btn secondary" data-close>Cancelar</button></div>
  </form>`;
  document.body.append(d);
  d.querySelector('[data-close]').onclick=()=>d.remove();
  d.querySelector('form').addEventListener('submit',async e=>{
    e.preventDefault();
    const f=new FormData(e.target);
    const status=f.get('status');
    const score=mod.is_exam&&f.get('score')!==''?Number(f.get('score')):null;
    const btn=d.querySelector('button.btn.primary');btn.disabled=true;btn.textContent='Guardando...';
    const{data,error}=await client().rpc('of_hr_set_candidate_module',{p_session_candidate_id:sessionCandidateId,p_module_id:mod.id,p_status:status,p_score:score});
    btn.disabled=false;btn.textContent='Guardar';
    if(error)return alert('No fue posible actualizar: '+error.message);
    d.remove();
    onDone(data.final_status,score);
  });
}

async function manage(id,existing){
  existing?.remove();
  const session=sessions.find(x=>x.id===id);
  const[mods,scRes]=await Promise.all([
    client().from('of_modules').select('id,name,module_order,medical_only,is_exam,passing_score').eq('active',true).order('module_order'),
    client().from('of_session_candidates').select('id,attendance_status,joined_at,candidate:of_candidates(id,status,person:of_people(full_name,document_id),project:of_projects(name,is_medical))').eq('session_id',id)
  ]);
  if(mods.error||scRes.error)return alert((mods.error||scRes.error).message);
  const modules=mods.data||[],participants=scRes.data||[];
  const scIds=participants.map(p=>p.id);
  let progress=[];
  if(scIds.length){
    const pr=await client().from('of_module_progress').select('session_candidate_id,module_id,status,score').in('session_candidate_id',scIds);
    if(pr.error)return alert(pr.error.message);
    progress=pr.data||[];
  }
  const progressMap=new Map(progress.map(p=>[`${p.session_candidate_id}:${p.module_id}`,p]));

  const d=document.createElement('div');d.className='modal-backdrop';
  d.innerHTML=`<div class="modal">
    <div class="management-head"><div><h2>${esc(session.title)}</h2><p>${esc(session.session_date)} · ${esc(session.start_time)} · <strong>${esc(statusLabel[session.status]||session.status)}</strong></p></div><button class="btn secondary" data-close>Cerrar</button></div>
    <div class="actions" style="justify-content:flex-start">
      <button class="btn primary" data-session-start>Iniciar inducción</button>
      <button class="btn secondary" data-session-close>Finalizar inducción</button>
      <button class="btn secondary" data-session-cancel>Cancelar inducción</button>
      ${session.meeting_url?`<a class="btn secondary" href="${esc(session.meeting_url)}" target="_blank" rel="noopener">Abrir Teams</a>`:''}
    </div>
    <h3>Participantes</h3>
    <p class="hint" style="color:var(--muted);font-size:13px">Autorice el ingreso y registre el avance de cada módulo/examen por persona. Toque cualquier estado para actualizarlo.</p>
    <div class="people">${participants.map(pcand=>{
      const pending=['registered','waiting_room'].includes(pcand.attendance_status);
      const isMedical=!!pcand.candidate?.project?.is_medical;
      const applicable=modules.filter(m=>!m.medical_only||isMedical);
      return `<article class="person-card" data-participant-card="${pcand.id}">
        <div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px">
          <div><strong>${esc(pcand.candidate?.person?.full_name)}</strong><br><small>${esc(pcand.candidate?.person?.document_id)} · ${esc(pcand.candidate?.project?.name||'')}</small></div>
          <div style="text-align:right">
            <span class="pill" style="background:${attendanceColor[pcand.attendance_status]||'var(--warning)'};color:#fff" data-attendance-pill="${pcand.id}">${attendanceLabel[pcand.attendance_status]||pcand.attendance_status}</span><br>
            <small data-candidate-status="${pcand.candidate?.id}">${esc(candidateStatusLabel[pcand.candidate?.status]||pcand.candidate?.status||'')}</small>
          </div>
        </div>
        <div class="actions" style="justify-content:flex-start;margin-top:8px">
          ${pending?`<button class="btn primary" data-admit="${pcand.id}">Autorizar ingreso</button> <button class="btn secondary" data-absent="${pcand.id}">Marcar ausente</button> `:''}<button class="btn secondary" data-incident="${pcand.id}">Incidencia</button>
        </div>
        <div class="table-wrap" style="margin-top:10px"><table><thead><tr><th>Módulo</th><th>Estado</th><th>Nota</th><th></th></tr></thead><tbody>${applicable.map(m=>{
          const prog=progressMap.get(`${pcand.id}:${m.id}`);
          const st=prog?.status||'pending';
          return `<tr><td>${esc(m.name)}${m.is_exam?' <small>(examen)</small>':''}</td><td><span class="pill" style="background:${moduleStatusColor[st]||'var(--warning)'};color:#fff" data-module-pill="${pcand.id}:${m.id}">${moduleStatusLabel[st]||st}</span></td><td data-module-score="${pcand.id}:${m.id}">${prog?.score??''}</td><td><button class="btn secondary" data-module-edit="${pcand.id}:${m.id}">Actualizar</button></td></tr>`;
        }).join('')||'<tr><td colspan="4">No hay módulos activos.</td></tr>'}</tbody></table></div>
      </article>`;
    }).join('')||'<p>No hay participantes en esta inducción.</p>'}</div>
  </div>`;
  document.body.append(d);

  d.querySelector('[data-close]').onclick=()=>d.remove();

  d.querySelector('[data-session-start]').addEventListener('click',async ev=>{
    const ok=await setSession(id,'in_progress',ev.currentTarget);
    if(ok)manage(id,d);
  });
  d.querySelector('[data-session-close]').addEventListener('click',async ev=>{
    if(!confirm('¿Finalizar esta inducción? Se cerrará esta ventana.'))return;
    const ok=await setSession(id,'closed',ev.currentTarget);
    if(ok){alert('Inducción finalizada.');d.remove();await load()}
  });
  d.querySelector('[data-session-cancel]').addEventListener('click',async ev=>{
    if(!confirm('¿Confirma que desea cancelar esta inducción y cerrar las invitaciones asociadas? Se cerrará esta ventana.'))return;
    const ok=await setSession(id,'cancelled',ev.currentTarget);
    if(ok){alert('Inducción cancelada.');d.remove();await load()}
  });

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

  d.querySelectorAll('[data-module-edit]').forEach(b=>b.addEventListener('click',ev=>{
    const[scid,mid]=ev.currentTarget.dataset.moduleEdit.split(':');
    const mod=modules.find(m=>m.id===mid);
    const prog=progressMap.get(`${scid}:${mid}`);
    moduleForm(scid,mod,prog?.status||'pending',prog?.score,(finalStatus,score)=>{
      const pill=d.querySelector(`[data-module-pill="${scid}:${mid}"]`);
      if(pill){pill.style.background=moduleStatusColor[finalStatus]||'var(--warning)';pill.textContent=moduleStatusLabel[finalStatus]||finalStatus}
      const scoreCell=d.querySelector(`[data-module-score="${scid}:${mid}"]`);
      if(scoreCell)scoreCell.textContent=score??'';
      progressMap.set(`${scid}:${mid}`,{session_candidate_id:scid,module_id:mid,status:finalStatus,score});
    });
  }));
}

function repeatsPanel(){
  const d=document.createElement('div');d.className='modal-backdrop';
  d.innerHTML=`<div class="modal"><div class="management-head"><div><h2>Pendientes de repetir</h2><p>De todas las inducciones, no solo la que tenga abierta.</p></div><button class="btn secondary" data-close>Cerrar</button></div><div class="table-wrap"><table><thead><tr><th>Nombre</th><th>Identificación</th><th>Proyecto</th><th>Módulo</th><th>Nota</th><th>Inducción</th></tr></thead><tbody data-repeats-body><tr><td colspan="6">Cargando...</td></tr></tbody></table></div></div>`;
  document.body.append(d);
  d.querySelector('[data-close]').onclick=()=>d.remove();
  client().from('of_v_pending_repeats').select('*').then(({data,error})=>{
    const body=d.querySelector('[data-repeats-body]');
    if(error){body.innerHTML=`<tr><td colspan="6">${esc(error.message)}</td></tr>`;return}
    body.innerHTML=(data||[]).map(r=>`<tr><td>${esc(r.full_name)}</td><td>${esc(r.document_id)}</td><td>${esc(r.project_name)}</td><td>${esc(r.module_name)}</td><td>${r.score??''}</td><td>${esc(r.session_title)} · ${esc(r.session_date)}</td></tr>`).join('')||'<tr><td colspan="6">No hay pendientes de repetir.</td></tr>';
  });
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
    <div class="management-head"><div><h2>Calendario de inducciones</h2><p>Inicie la sesión y avance cada módulo desde aquí.</p></div><div style="display:flex;gap:8px"><button class="btn secondary" data-repeats>Pendientes de repetir</button><button class="btn primary" data-new>Nueva inducción</button></div></div>
    <div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Hora</th><th>Nombre</th><th>Asignados</th><th>Estado</th><th>Acciones</th></tr></thead><tbody>${sessions.map(s=>`<tr><td>${esc(s.session_date)}</td><td>${esc(s.start_time)}</td><td>${esc(s.title)}</td><td>${s.participants?.[0]?.count||0}/${s.capacity||30}</td><td>${esc(statusLabel[s.status]||s.status)}</td><td><button class="btn primary" data-manage="${s.id}">Controlar</button> <button class="btn secondary" data-edit="${s.id}">Editar</button></td></tr>`).join('')||'<tr><td colspan="6">No hay inducciones.</td></tr>'}</tbody></table></div>
  </section>`;
  root.querySelector('[data-new]').onclick=()=>form();
  root.querySelector('[data-repeats]').onclick=()=>repeatsPanel();
  root.querySelectorAll('[data-manage]').forEach(b=>b.onclick=()=>manage(b.dataset.manage));
  root.querySelectorAll('[data-edit]').forEach(b=>b.onclick=()=>form(sessions.find(s=>s.id===b.dataset.edit)));
}

document.addEventListener('DOMContentLoaded',load);
})();
