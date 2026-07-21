(function(){'use strict';
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const candidateStatusLabel={pending_hr_review:'Pendiente revisión RH',returned_for_correction:'Devuelto',approved_by_hr:'Aprobado por RH',invited:'Invitado',registered:'Registrado',in_induction:'En inducción',pending_module:'Debe repetir algo',approved_for_hire:'Aprobado para contratar',not_approved:'No aprobado',cancelled:'Cancelado',absent:'Ausente'};

async function load(){
  const{data,error}=await OnboardAuth.client.from('of_hiring_records').select('*,candidate:of_candidates(status,person:of_people(full_name,document_id),project:of_projects(name),supervisor:of_profiles!of_candidates_interviewing_supervisor_id_fkey(full_name))').order('supervisor_approved_at',{ascending:false});
  if(error){document.getElementById('hiringRecords').innerHTML=`<p>${esc(error.message)}</p>`;return}
  document.getElementById('hiringRecords').innerHTML=(data||[]).map(r=>{
    const ready=r.candidate?.status==='approved_for_hire';
    return `<article class="person"><div class="person-top"><div class="avatar">${(r.candidate?.person?.full_name||'C').split(' ').map(x=>x[0]).slice(0,2).join('')}</div><span class="pill ${r.hr_status==='accepted'?'ok':r.hr_status==='returned'?'bad':'warn'}">${r.hr_status}</span></div><div class="person-name">${esc(r.candidate?.person?.full_name)}</div><div class="person-meta">${esc(r.candidate?.person?.document_id)} · ${esc(r.candidate?.project?.name)} · ${esc(r.position_name)}</div><div class="person-meta">Inducción: ${esc(candidateStatusLabel[r.candidate?.status]||r.candidate?.status||'')}${!ready?' — todavía no lista para aceptar':''}</div><div class="actions"><button class="btn secondary" data-view-hiring="${r.id}">Ver expediente</button>${r.hr_status==='pending'?`<button class="btn primary" data-accept-hiring="${r.id}" ${ready?'':'disabled title="Falta completar módulos/exámenes de inducción"'}>Aceptar ingreso</button><button class="btn secondary" data-return-hiring="${r.id}">Devolver</button>`:''}</div></article>`;
  }).join('')||'<p>No hay expedientes pendientes.</p>';
  window.hiringRows=data||[];
}

function view(r){
  const d=document.createElement('div');d.className='modal-backdrop';
  const items=[['Nombre',r.candidate?.person?.full_name],['Cedula',r.candidate?.person?.document_id],['Seguro social',r.social_security_number],['Nacionalidad',r.nationality],['Direccion',r.address],['Celular',r.mobile_phone],['Contacto de emergencia',r.emergency_contact_name],['Numero de emergencia',r.emergency_contact_phone],['Puesto',r.position_name],['Salario',r.salary],['Horario',r.work_shift],['Capacitacion 1',r.training_date_1],['Capacitacion 2',r.training_date_2],['Capacitacion 3',r.training_date_3],['Supervisor',r.candidate?.supervisor?.full_name],['Estado de inducción',candidateStatusLabel[r.candidate?.status]||r.candidate?.status]];
  d.innerHTML=`<div class="modal"><h2>Expediente de contratación</h2>${items.map(i=>`<div class="kv"><span>${esc(i[0])}</span><strong>${esc(i[1]||'No aplica')}</strong></div>`).join('')}<button class="btn secondary" data-close>Cerrar</button></div>`;
  document.body.append(d);
  d.querySelector('[data-close]').onclick=()=>d.remove();
}

async function status(id,hr_status,notes,btn){
  if(btn){btn.disabled=true;btn.textContent='Procesando...'}
  const{error}=await OnboardAuth.client.rpc('of_hr_review_hiring_record',{p_id:id,p_hr_status:hr_status,p_notes:notes||null});
  if(error){alert(error.message);if(btn){btn.disabled=false;btn.textContent=hr_status==='accepted'?'Aceptar ingreso':'Devolver'}return}
  await load();
}

document.addEventListener('DOMContentLoaded',()=>{
  const tabs=document.querySelector('.tabs'),btn=document.createElement('button');
  btn.className='tab';btn.textContent='Ingresos para RH';btn.onclick=()=>showView('hiring',btn);
  tabs.append(btn);
  const section=document.createElement('section');
  section.id='hiring';section.className='view';
  section.innerHTML='<section class="card section"><h2>Expedientes listos para contratación</h2><p style="color:var(--muted)">Enviados por supervisores después de aprobar exámenes y capacitaciones. Solo se puede aceptar si la inducción quedó en estado "Aprobado para contratar".</p><div id="hiringRecords" class="people"></div></section>';
  document.querySelector('main').append(section);
  load();
});

document.addEventListener('click',e=>{
  if(e.target.dataset.viewHiring)view(window.hiringRows.find(r=>r.id===e.target.dataset.viewHiring));
  if(e.target.dataset.acceptHiring)status(e.target.dataset.acceptHiring,'accepted',null,e.target);
  if(e.target.dataset.returnHiring){
    const n=prompt('Indique lo que debe corregirse:');
    if(n)status(e.target.dataset.returnHiring,'returned',n,e.target);
  }
});
})();
