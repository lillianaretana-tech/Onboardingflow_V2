(function(){'use strict';
async function closeCandidate(id){const reason=prompt('Motivo del rechazo o cierre del proceso:');if(!reason)return;const{error}=await OnboardAuth.client.rpc('of_hr_close_candidate_process',{p_candidate_id:id,p_final_status:'not_approved',p_reason:reason});if(error)return alert('No fue posible cerrar el proceso: '+error.message);alert('Proceso cerrado. Se revocó el código y se retiró a la persona de invitaciones y sala de espera.');location.reload()}
document.addEventListener('click',e=>{const reject=e.target.closest('[data-reject-id]');if(reject){e.preventDefault();e.stopImmediatePropagation();closeCandidate(reject.dataset.rejectId);return}},true);
// Nota: el boton "Cancelar induccion" que antes se inyectaba aqui via
// MutationObserver ahora vive de forma nativa dentro de induction-management.js
// (data-session-cancel), para evitar la condicion de carrera que impedia
// que Iniciar/Finalizar y los botones de modulo respondieran.
})();
