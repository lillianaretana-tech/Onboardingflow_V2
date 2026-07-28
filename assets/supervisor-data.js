(function(){'use strict';let profile;const client=()=>window.OnboardAuth.client,statusText={pending_hr_review:'Pendiente de RH',returned_for_correction:'Devuelto por RH',approved_by_hr:'Aprobado por RH',invited:'Invitación enviada',registered:'Registrado',in_induction:'En inducción',approved_for_hire:'Aprobado para contratar',not_approved:'No aprobado'};function fields(){const section=document.getElementById('registrationSection'),inputs=section.querySelectorAll('input,select');return{phone:inputs[0].value.trim(),email:inputs[1].value.trim(),project:inputs[2].value,position:inputs[3].value.trim(),date:inputs[4].value}}async function loadMine(){const {data,error}=await client().from('of_candidates').select('id,status,position_name,proposed_induction_date,person:of_people(full_name),project:of_projects(name)').order('created_at',{ascending:false}).limit(50);if(error)throw error;const list=document.querySelector('.section .people')||document.querySelector('.people');if(list)list.innerHTML=(data||[]).map(c=>`<article class="person"><div class="avatar">${(c.person?.full_name||'C').split(' ').map(x=>x[0]).slice(0,2).join('')}</div><div class="person-name">${c.person?.full_name||'Candidato'}</div><div class="person-meta">${c.project?.name||'Sin proyecto'} · ${c.position_name||'Sin puesto'}</div><span class="pill ${c.status==='approved_by_hr'?'pill-ok':'pill-warn'}" style="margin-top:10px">${statusText[c.status]||c.status}</span></article>`).join('')||'<p>No ha registrado candidatos todavía.</p>';const metrics=document.querySelectorAll('.metric strong');if(metrics[0])metrics[0].textContent=data.length;if(metrics[1])metrics[1].textContent=data.filter(c=>c.status==='pending_hr_review').length;if(metrics[2])metrics[2].textContent=data.filter(c=>c.status==='returned_for_correction').length}

const eligLabel={eligible:['pill-ok','Contratable'],review_required:['pill-warn','Revisión RH'],not_eligible:['pill-bad','No contratable']};

window.checkEligibility=async function(){
  const name=document.getElementById('candidateName').value.trim();
  const id=document.getElementById('candidateId').value.trim();
  const box=document.getElementById('eligibilityBox'),title=document.getElementById('eligibilityTitle'),text=document.getElementById('eligibilityText'),reg=document.getElementById('registrationSection');

  if(!id&&!name){alert('Escriba al menos el nombre para verificar.');return}

  box.className='eligibility show';
  try{
    if(id){
      // Verificacion exacta por cedula — se mantiene igual que siempre.
      const{data,error}=await client().rpc('of_check_eligibility',{p_document_id:id});
      if(error)throw error;
      const state=data.status;
      if(state==='not_eligible'){box.classList.add('bad');title.textContent='No elegible para contratación'}
      else if(state==='review_required'){box.classList.add('review');title.textContent='Revisión de Recursos Humanos'}
      else{box.classList.add('ok');title.textContent='Apto para continuar'}
      text.textContent=data.message;
    }else{
      // Verificacion rapida solo por nombre (CV sin cedula) — no registra
      // nada, es solo para decidir si vale la pena entrevistar. No toca
      // el formulario de registro, que es independiente de esto.
      if(name.length<3){alert('Escriba al menos 3 letras del nombre.');box.classList.remove('show');return}
      const{data,error}=await client().rpc('of_check_eligibility_by_name',{p_name:name});
      if(error)throw error;
      const matches=data.matches||[];
      if(!matches.length){
        box.classList.add('ok');
        title.textContent='Sin coincidencias en exfuncionarios';
        text.textContent='No aparece nadie con ese nombre en el historial. Puede continuar con la entrevista; si más adelante consigue la cédula, puede volver a verificar para mayor certeza.';
      }else{
        const worstClass=matches.some(m=>m.eligibility_status==='not_eligible')?'bad':matches.some(m=>m.eligibility_status==='review_required')?'review':'ok';
        box.classList.add(worstClass);
        title.textContent=`${matches.length} coincidencia(s) — revise cuál es la persona antes de decidir`;
        text.innerHTML=matches.map(m=>{
          const[cls,label]=eligLabel[m.eligibility_status]||['pill-warn',m.eligibility_status];
          return `<div style="margin-top:10px;padding-top:10px;border-top:1px solid rgba(0,0,0,.08)">
            <strong>${esc(m.full_name)}</strong> ${m.document_id?`(céd. ${esc(m.document_id)})`:'<em>(sin cédula registrada)</em>'}<br>
            <span class="pill ${cls}">${label}</span>
            ${m.company_name?` · ${esc(m.company_name)}`:''}${m.previous_project&&m.previous_project!==m.company_name?` · ${esc(m.previous_project)}`:''}
            ${m.exit_date?` · Salió: ${esc(m.exit_date)}`:''}${m.exit_reason?` · ${esc(m.exit_reason)}`:''}
          </div>`;
        }).join('');
      }
    }
  }catch(error){
    box.classList.add('bad');title.textContent='No fue posible verificar';text.textContent=error.message;
  }
};

function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}

window.submitCandidate=async function(){const name=document.getElementById('candidateName').value.trim(),documentId=document.getElementById('candidateId').value.trim(),f=fields(),button=document.querySelector('[onclick="submitCandidate()"]'),message=document.getElementById('submitMessage');if(!name||!documentId||!f.project||!f.position){alert('Complete nombre, identificación, proyecto y puesto.');return}button.disabled=true;try{const {data,error}=await client().rpc('of_supervisor_register_candidate',{p_full_name:name,p_document_id:documentId,p_email:f.email||null,p_phone:f.phone||null,p_project_name:f.project,p_position_name:f.position,p_induction_date:f.date||null});if(error)throw error;message.textContent='Candidato enviado correctamente a Recursos Humanos. Estado: Pendiente de revisión.';message.classList.remove('hidden');document.getElementById('registrationSection').style.display='none';await loadMine()}catch(error){alert('No fue posible enviar el candidato: '+error.message)}finally{button.disabled=false}};
document.addEventListener('DOMContentLoaded',async()=>{profile=await window.OnboardAuth.guard(['supervisor']);if(profile)try{await loadMine()}catch(error){console.error(error)}})})();
