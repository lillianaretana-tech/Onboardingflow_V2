(function(){'use strict';
let profile,rows=[];
const client=()=>OnboardAuth.client,x=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const ELIG_LABEL={eligible:'Contratable',review_required:'Revision RH',not_eligible:'No contratable'};
const ELIG_VALUE={Contratable:'eligible','Revision RH':'review_required','No contratable':'not_eligible'};

function toRow(r){
  return {
    id:r.id, person_id:r.person_id,
    name:r.person?.full_name||'', document:r.person?.document_id||'', email:r.person?.email||'', phone:r.person?.phone||'',
    company:r.company_name||'', project:r.previous_project||'', position:r.previous_position||'',
    entry_date:r.entry_date||'', exit_date:r.exit_date||'', exit_reason:r.exit_reason||'',
    eligibility:ELIG_LABEL[r.eligibility_status]||'Revision RH', internal_notes:r.internal_notes||'',
    review_date:r.reviewed_at||'', deleted_at:r.deleted_at, deleted_by:r.deleted_by, deletion_reason:r.deletion_reason
  };
}

async function load(){
  profile=profile||await OnboardAuth.getProfile();
  const{data,error}=await client().from('of_former_employees').select('*,person:of_people(full_name,document_id,email,phone)').order('updated_at',{ascending:false});
  if(error)return console.error(error);
  rows=(data||[]).map(toRow);
  render();
}

function render(){
  let q=(document.getElementById('formerSearch')?.value||'').toLowerCase(),f=document.getElementById('eligibilityFilter')?.value,del=document.getElementById('showDeletedFormer')?.checked;
  document.getElementById('formerTable').innerHTML=rows.filter(r=>(del||!r.deleted_at)&&(!f||r.eligibility===f)&&Object.values(r).some(v=>String(v).toLowerCase().includes(q)))
    .map(r=>`<tr class="${r.deleted_at?'row-deleted':''}"><td>${x(r.name)}<br><small>${x(r.email)} ${x(r.phone)}</small></td><td>${x(r.document)}</td><td>${x(r.company)}</td><td>${x(r.project)}<br>${x(r.position)}</td><td>${x(r.exit_date)}<br>${x(r.exit_reason)}</td><td>${x(r.eligibility)}</td><td><button class="btn secondary" data-edit="${r.id}">Editar</button> <button class="btn secondary" data-delete="${r.id}">${r.deleted_at?'Restaurar':'Eliminar'}</button></td></tr>`)
    .join('')||'<tr><td colspan="7">No hay registros.</td></tr>';
}

function form(r={}){
  let d=document.createElement('div');d.className='modal-backdrop';
  let fs=[['name','Nombre completo'],['document','Cedula'],['email','Correo'],['phone','Telefono'],['company','Empresa'],['project','Proyecto anterior'],['position','Puesto anterior'],['entry_date','Fecha ingreso','date'],['exit_date','Fecha salida','date'],['exit_reason','Motivo salida'],['review_date','Fecha revision','date']];
  d.innerHTML=`<form class="modal"><h2>${r.id?'Editar':'Agregar'} exfuncionario</h2><div class="form-grid">${fs.map(a=>`<div class="field"><label class="label">${a[1]}</label><input class="input" name="${a[0]}" type="${a[2]||'text'}" value="${x(r[a[0]])}" ${['name','document','company'].includes(a[0])?'required':''}></div>`).join('')}<div class="field"><label class="label">Elegibilidad</label><select class="select" name="eligibility"><option ${r.eligibility==='Contratable'?'selected':''}>Contratable</option><option ${r.eligibility==='Revision RH'?'selected':''}>Revision RH</option><option ${r.eligibility==='No contratable'?'selected':''}>No contratable</option></select></div></div><div class="field"><label class="label">Observacion interna (solo RH/admin)</label><textarea class="input" name="internal_notes">${x(r.internal_notes)}</textarea></div><div class="actions"><button class="btn primary">Guardar</button> <button type="button" class="btn secondary" data-close>Cerrar</button></div></form>`;
  document.body.append(d);
  d.querySelector('[data-close]').onclick=()=>d.remove();
  d.querySelector('form').onsubmit=async e=>{
    e.preventDefault();
    const f=new FormData(e.target),n=Object.fromEntries(f);
    if(!profile)profile=await OnboardAuth.getProfile();

    const{data:person,error:pErr}=await client().from('of_people').upsert({
      tenant_id:profile.tenant_id, full_name:n.name, document_id:n.document,
      email:n.email||null, phone:n.phone||null, current_relationship:'former_employee'
    },{onConflict:'tenant_id,document_id'}).select().single();
    if(pErr)return alert(pErr.message);

    const payload={
      tenant_id:profile.tenant_id, person_id:person.id, company_name:n.company,
      previous_project:n.project||null, previous_position:n.position||null,
      entry_date:n.entry_date||null, exit_date:n.exit_date||null, exit_reason:n.exit_reason||null,
      eligibility_status:ELIG_VALUE[n.eligibility]||'review_required',
      internal_notes:n.internal_notes||null, reviewed_at:n.review_date||null, updated_by:profile.id
    };
    const q=r.id?client().from('of_former_employees').update(payload).eq('id',r.id):client().from('of_former_employees').insert(payload);
    const{error}=await q;
    if(error)return alert(error.message);
    d.remove();await load();
  };
}

const H=['Nombre completo','Cedula o identificacion','Correo','Telefono','Empresa','Proyecto anterior','Puesto anterior','Fecha de ingreso','Fecha de salida','Motivo de salida','Elegibilidad','Observacion interna','Fecha de revision'];
const K=['name','document','email','phone','company','project','position','entry_date','exit_date','exit_reason','eligibility','internal_notes','review_date'];

function wb(data,n){
  let s=XLSX.utils.aoa_to_sheet([H,...data.map(r=>K.map(k=>r[k]||''))]),w=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(w,s,n);return w;
}

function wbReport(data,title){
  const totalCols=H.length,lastCol=XLSX.utils.encode_col(totalCols-1);
  const generated=new Date().toLocaleString('es-CR');
  const aoa=[
    ['LILLYTECH ONBOARDFLOW',...Array(totalCols-1).fill('')],
    [title,...Array(totalCols-1).fill('')],
    [`Generado: ${generated}`,...Array(totalCols-1).fill('')],
    [], H, ...data.map(r=>K.map(k=>r[k]||''))
  ];
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges']=[`A1:${lastCol}1`,`A2:${lastCol}2`,`A3:${lastCol}3`].map(XLSX.utils.decode_range);
  ws['!cols']=[24,16,24,14,20,20,20,14,14,20,16,26,14].map(w=>({wch:w}));
  ws['!autofilter']={ref:`A5:${lastCol}${Math.max(5,data.length+5)}`};
  ws['!freeze']={xSplit:0,ySplit:5,topLeftCell:'A6',state:'frozen'};
  const border={top:{style:'thin',color:{rgb:'D9D5CB'}},bottom:{style:'thin',color:{rgb:'D9D5CB'}},left:{style:'thin',color:{rgb:'D9D5CB'}},right:{style:'thin',color:{rgb:'D9D5CB'}}};
  function paint(range,fill,color,bold,size,center){
    const r=XLSX.utils.decode_range(range);
    for(let y=r.s.r;y<=r.e.r;y++)for(let x=r.s.c;x<=r.e.c;x++){
      const a=XLSX.utils.encode_cell({r:y,c:x});
      if(!ws[a])ws[a]={t:'s',v:''};
      ws[a].s={font:{name:'Aptos',color:{rgb:color},bold,sz:size},fill:{patternType:'solid',fgColor:{rgb:fill}},alignment:{vertical:'center',horizontal:center?'center':'left',wrapText:true},border};
    }
  }
  paint(`A1:${lastCol}1`,'0B132B','FFFFFF',true,19,true);
  paint(`A2:${lastCol}2`,'C9782A','FFFFFF',true,13,true);
  paint(`A3:${lastCol}3`,'FAF8F2','667085',false,10,false);
  paint(`A5:${lastCol}5`,'0B132B','FFFFFF',true,10,true);
  data.forEach((_,i)=>paint(`A${i+6}:${lastCol}${i+6}`,i%2?'FFFFFF':'F7F4EC','242936',false,10,false));
  const w=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(w,ws,'Exfuncionarios');return w;
}

function formatChooser(onXlsx,onPdf){
  const d=document.createElement('div');d.className='modal-backdrop';
  d.innerHTML=`<div class="modal" style="max-width:360px"><h2>Exportar</h2><p style="color:var(--muted);font-size:14px">¿En qué formato lo prefiere?</p><div class="actions" style="flex-direction:column;gap:10px"><button class="btn primary" data-fmt-xlsx>Excel (.xlsx)</button><button class="btn secondary" data-fmt-pdf>PDF</button><button type="button" class="btn secondary" data-close>Cancelar</button></div></div>`;
  document.body.append(d);
  d.querySelector('[data-close]').onclick=()=>d.remove();
  d.querySelector('[data-fmt-xlsx]').onclick=()=>{d.remove();onXlsx()};
  d.querySelector('[data-fmt-pdf]').onclick=()=>{d.remove();onPdf()};
}

function loadPdfLibs(){
  return new Promise((resolve,reject)=>{
    if(window.jspdf&&window.jspdf.jsPDF&&typeof new window.jspdf.jsPDF().autoTable==='function')return resolve();
    const s1=document.createElement('script');s1.src='https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
    s1.onload=()=>{
      const s2=document.createElement('script');s2.src='https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js';
      s2.onload=()=>resolve();s2.onerror=reject;document.head.append(s2);
    };
    s1.onerror=reject;document.head.append(s1);
  });
}

async function exportFormerPdf(){
  await loadPdfLibs();
  const{jsPDF}=window.jspdf;
  const doc=new jsPDF({orientation:'landscape',unit:'pt',format:'letter'});
  const w=doc.internal.pageSize.getWidth();
  doc.setFillColor(11,19,43);doc.rect(0,0,w,58,'F');
  doc.setTextColor(255,255,255);doc.setFont('times','bold');doc.setFontSize(16);
  doc.text('LillyTech OnboardFlow',30,26);
  doc.setFont('helvetica','normal');doc.setFontSize(10);doc.setTextColor(210,215,225);
  doc.text('Reporte de exfuncionarios',30,44);
  doc.setFillColor(201,120,42);doc.rect(0,58,w,20,'F');
  doc.setTextColor(255,255,255);doc.setFontSize(9);
  doc.text(`${rows.length} registros  ·  Generado: `+new Date().toLocaleString('es-CR'),30,72);
  doc.autoTable({
    startY:92,
    head:[['Nombre','Cédula','Empresa','Proyecto/Puesto','Fecha salida','Motivo','Elegibilidad']],
    body:rows.filter(r=>!r.deleted_at).map(r=>[r.name,r.document,r.company,`${r.project||''} ${r.position||''}`.trim(),r.exit_date||'',r.exit_reason||'',r.eligibility||'']),
    styles:{font:'helvetica',fontSize:8,cellPadding:5,lineColor:[231,228,220],lineWidth:.5},
    headStyles:{fillColor:[11,19,43],textColor:255,fontStyle:'bold'},
    alternateRowStyles:{fillColor:[247,244,236]},
    margin:{left:30,right:30}
  });
  doc.save(`OnboardFlow_Exfuncionarios_${new Date().toISOString().slice(0,10)}.pdf`);
}

function load_file(file){
  let rd=new FileReader;
  rd.onload=e=>{
    let w=XLSX.read(e.target.result,{type:'array',cellDates:true}),a=XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]],{header:1,defval:''});
    a=a.map(row=>row.map(v=>v instanceof Date?v.toISOString().slice(0,10):v));
    // Busca la fila de encabezados entre las primeras filas, por si la
    // plantilla trae filas de título/branding antes de los encabezados.
    let headerIndex=a.findIndex(row=>H.every(h=>row.includes(h)));
    if(headerIndex<0){
      document.getElementById('importPreview').classList.remove('hidden');
      document.getElementById('importValidation').textContent='No se encontró la fila de encabezados esperada. Use la plantilla oficial sin modificarla.';
      document.getElementById('previewTable').innerHTML='<tbody>'+a.slice(0,11).map(row=>'<tr>'+row.map(v=>`<td>${x(v)}</td>`).join('')+'</tr>').join('')+'</tbody>';
      document.querySelector('[data-action="confirm-import"]').disabled=true;
      window.ofPreview=[];
      return;
    }
    window.ofPreview=a.slice(headerIndex+1).filter(row=>row.some(Boolean));
    document.getElementById('importPreview').classList.remove('hidden');
    document.getElementById('importValidation').textContent=window.ofPreview.length+' filas validadas';
    document.getElementById('previewTable').innerHTML='<tbody><tr>'+H.map(h=>`<th>${x(h)}</th>`).join('')+'</tr>'+window.ofPreview.slice(0,10).map(row=>'<tr>'+row.map(v=>`<td>${x(v)}</td>`).join('')+'</tr>').join('')+'</tbody>';
    document.querySelector('[data-action="confirm-import"]').disabled=false;
  };
  rd.readAsArrayBuffer(file);
}

async function confirm_import(){
  if(!profile)profile=await OnboardAuth.getProfile();
  const mode=document.getElementById('duplicateMode').value;
  const payload=(window.ofPreview||[]).map(v=>Object.fromEntries(K.map((k,i)=>[k,v[i]||''])));
  const btn=document.querySelector('[data-action="confirm-import"]');
  btn.disabled=true;btn.textContent='Importando...';
  const{data,error}=await client().rpc('of_hr_import_former_employees',{p_rows:payload,p_mode:mode});
  btn.disabled=false;btn.textContent='Confirmar importación';
  if(error)return alert(error.message);
  let b=document.getElementById('importSummary');
  b.classList.remove('hidden');
  b.textContent=`Agregados: ${data.added} · Actualizados: ${data.updated} · Omitidos: ${data.skipped} · Errores: ${data.errors}`;
  await load();
}

document.addEventListener('click',async e=>{
  let a=e.target.dataset.action;
  if(a==='new-former')form();
  if(a==='template')XLSX.writeFile(wb([],'Plantilla'),'OnboardFlow_Plantilla_Exfuncionarios.xlsx');
  if(a==='export-former')formatChooser(
    ()=>XLSX.writeFile(wbReport(rows.filter(r=>!r.deleted_at),'REPORTE DE EXFUNCIONARIOS'),`OnboardFlow_Exfuncionarios_${new Date().toISOString().slice(0,10)}.xlsx`),
    exportFormerPdf
  );
  if(a==='confirm-import')confirm_import();
  if(e.target.dataset.edit)form(rows.find(r=>r.id===e.target.dataset.edit));
  if(e.target.dataset.delete){
    let r=rows.find(v=>v.id===e.target.dataset.delete);
    if(!profile)profile=await OnboardAuth.getProfile();
    if(r.deleted_at){
      const{error}=await client().from('of_former_employees').update({deleted_at:null,deleted_by:null,deletion_reason:null,updated_by:profile.id}).eq('id',r.id);
      if(error)return alert(error.message);
    }else{
      let reason=prompt('Motivo de eliminacion logica:');
      if(!reason)return;
      const{error}=await client().from('of_former_employees').update({deleted_at:new Date().toISOString(),deleted_by:profile.id,deletion_reason:reason,updated_by:profile.id}).eq('id',r.id);
      if(error)return alert(error.message);
    }
    await load();
  }
});
document.addEventListener('input',e=>{if(['formerSearch','eligibilityFilter','showDeletedFormer'].includes(e.target.id))render()});
document.addEventListener('change',e=>{
  if(e.target.id==='formerImport'&&e.target.files[0])load_file(e.target.files[0]);
  if(['eligibilityFilter','showDeletedFormer'].includes(e.target.id))render();
});
document.addEventListener('DOMContentLoaded',load);
})();
