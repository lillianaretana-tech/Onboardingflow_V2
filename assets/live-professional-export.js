(function(){'use strict';
const candidateStatusLabel={pending_hr_review:'Pendiente revisión RH',returned_for_correction:'Devuelto',approved_by_hr:'Aprobado por RH',invited:'Invitado',registered:'Registrado',in_induction:'En inducción',pending_module:'Debe repetir algo',approved_for_hire:'Aprobado para contratar',not_approved:'No aprobado',cancelled:'Cancelado',absent:'Ausente'};
const attendanceLabel={registered:'Registrado',waiting_room:'En sala de espera',present:'Presente',late:'Tardío',absent:'Ausente',incomplete:'Incompleto'};
const moduleStatusLabel={pending:'Pendiente',in_progress:'En curso',completed:'Completado',approved:'Aprobado',interrupted:'Interrumpido',must_repeat:'Debe repetir',failed:'No aprobado'};
const hiringLabel={pending:'Pendiente',accepted:'Aceptado',returned:'Devuelto'};

function formatChooser(onXlsx,onPdf){
  const d=document.createElement('div');d.className='modal-backdrop';
  d.innerHTML=`<div class="modal" style="max-width:360px"><h2>Exportar reporte</h2><p style="color:var(--muted);font-size:14px">¿En qué formato lo prefiere?</p><div class="actions" style="flex-direction:column;gap:10px"><button class="btn primary" data-fmt-xlsx>Excel (.xlsx)</button><button class="btn secondary" data-fmt-pdf>PDF</button><button type="button" class="btn secondary" data-close>Cancelar</button></div></div>`;
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

async function renderPdfReport(opts){
  await loadPdfLibs();
  const{jsPDF}=window.jspdf;
  const doc=new jsPDF({orientation:'landscape',unit:'pt',format:'letter'});
  const w=doc.internal.pageSize.getWidth();
  doc.setFillColor(11,19,43);doc.rect(0,0,w,58,'F');
  doc.setTextColor(255,255,255);doc.setFont('times','bold');doc.setFontSize(16);
  doc.text('LillyTech OnboardFlow',30,26);
  doc.setFont('helvetica','normal');doc.setFontSize(10);doc.setTextColor(210,215,225);
  doc.text(opts.title,30,44);
  doc.setFillColor(201,120,42);doc.rect(0,58,w,20,'F');
  doc.setTextColor(255,255,255);doc.setFontSize(9);
  doc.text(opts.subtitle+'  ·  Generado: '+new Date().toLocaleString('es-CR'),30,72);
  doc.autoTable({
    startY:92, head:[opts.columns], body:opts.rows,
    styles:{font:'helvetica',fontSize:7.5,cellPadding:4,lineColor:[231,228,220],lineWidth:.5},
    headStyles:{fillColor:[11,19,43],textColor:255,fontStyle:'bold'},
    alternateRowStyles:{fillColor:[247,244,236]},
    margin:{left:24,right:24}
  });
  doc.save(opts.filename);
}

async function fetchCandidateReport(){
  const{data,error}=await OnboardAuth.client.from('of_v_candidate_report').select('*').order('full_name');
  if(error){alert('No fue posible generar el reporte: '+error.message);return null}
  return data||[];
}

function moduleNamesFrom(rows){
  const names=[];
  rows.forEach(r=>(r.modules||[]).forEach(m=>{if(!names.includes(m.module))names.push(m.module)}));
  names.sort((a,b)=>{
    const oa=rows.flatMap(r=>r.modules||[]).find(m=>m.module===a)?.order||0;
    const ob=rows.flatMap(r=>r.modules||[]).find(m=>m.module===b)?.order||0;
    return oa-ob;
  });
  return names;
}

async function exportCandidateXlsx(){
  if(typeof XLSX==='undefined'){alert('No fue posible cargar Excel.');return}
  const rows=await fetchCandidateReport();
  if(!rows)return;
  try{
    const moduleNames=moduleNamesFrom(rows);
    const baseHeaders=['Nombre completo','Cédula','Correo','Teléfono','Proyecto','Puesto','Supervisor','Inducción','Fecha inducción','Asistencia','Estado del candidato','Estado de contratación'];
    const moduleHeaders=moduleNames.flatMap(m=>[`${m} · Estado`,`${m} · Nota`]);
    const headers=[...baseHeaders,...moduleHeaders];

    const values=rows.map(r=>{
      const modMap=new Map((r.modules||[]).map(m=>[m.module,m]));
      const base=[
        r.full_name||'', r.document_id||'', r.email||'', r.phone||'', r.project_name||'', r.position_name||'',
        r.supervisor_name||'', r.session_title||'', r.session_date||'',
        attendanceLabel[r.attendance_status]||r.attendance_status||'',
        candidateStatusLabel[r.candidate_status]||r.candidate_status||'',
        hiringLabel[r.hiring_status]||r.hiring_status||''
      ];
      const moduleCols=moduleNames.flatMap(name=>{
        const m=modMap.get(name);
        return [m?(moduleStatusLabel[m.status]||m.status):'', m?.score??''];
      });
      return[...base,...moduleCols];
    });

    const generated=new Date().toLocaleString('es-CR');
    const totalCols=headers.length;
    const lastCol=XLSX.utils.encode_col(totalCols-1);
    const aoaRows=[
      ['LILLYTECH ONBOARDFLOW',...Array(totalCols-1).fill('')],
      ['REPORTE COMPLETO DE INDUCCIÓN — RESULTADOS Y ESTADO',...Array(totalCols-1).fill('')],
      [`Generado: ${generated}`,...Array(totalCols-1).fill('')],
      [],
      headers,
      ...values
    ];
    const ws=XLSX.utils.aoa_to_sheet(aoaRows);
    ws['!merges']=[`A1:${lastCol}1`,`A2:${lastCol}2`,`A3:${lastCol}3`].map(XLSX.utils.decode_range);
    ws['!cols']=[28,16,26,16,20,20,20,24,16,16,20,18,...moduleNames.flatMap(()=>[16,10])].map(w=>({wch:w}));
    ws['!autofilter']={ref:`A5:${lastCol}${Math.max(5,values.length+5)}`};
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
    values.forEach((_,i)=>paint(`A${i+6}:${lastCol}${i+6}`,i%2?'FFFFFF':'F7F4EC','242936',false,10,false));

    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,'Reporte completo');
    const d=new Date().toISOString().slice(0,10);
    XLSX.writeFile(wb,`LillyTech_Reporte_Induccion_Completo_${d}.xlsx`);
  }catch(error){
    alert('No fue posible exportar el reporte: '+error.message);
  }
}

async function exportCandidatePdf(){
  const rows=await fetchCandidateReport();
  if(!rows)return;
  try{
    const columns=['Nombre','Cédula','Proyecto','Puesto','Supervisor','Asistencia','Estado','Contratación'];
    const body=rows.map(r=>[
      r.full_name||'', r.document_id||'', r.project_name||'', r.position_name||'', r.supervisor_name||'',
      attendanceLabel[r.attendance_status]||r.attendance_status||'',
      candidateStatusLabel[r.candidate_status]||r.candidate_status||'',
      hiringLabel[r.hiring_status]||r.hiring_status||''
    ]);
    await renderPdfReport({
      title:'Reporte de inducción — resultados y estado', subtitle:`${body.length} candidatos`,
      columns, rows:body, filename:`LillyTech_Reporte_Induccion_${new Date().toISOString().slice(0,10)}.pdf`
    });
  }catch(error){
    alert('No fue posible exportar el PDF: '+error.message);
  }
}

window.exportXlsx=function(){formatChooser(exportCandidateXlsx,exportCandidatePdf)};
})();
