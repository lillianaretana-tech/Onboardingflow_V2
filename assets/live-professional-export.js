(function(){'use strict';
const candidateStatusLabel={pending_hr_review:'Pendiente revisión RH',returned_for_correction:'Devuelto',approved_by_hr:'Aprobado por RH',invited:'Invitado',registered:'Registrado',in_induction:'En inducción',pending_module:'Debe repetir algo',approved_for_hire:'Aprobado para contratar',not_approved:'No aprobado',cancelled:'Cancelado',absent:'Ausente'};
const attendanceLabel={registered:'Registrado',waiting_room:'En sala de espera',present:'Presente',late:'Tardío',absent:'Ausente',incomplete:'Incompleto'};
const moduleStatusLabel={pending:'Pendiente',in_progress:'En curso',completed:'Completado',approved:'Aprobado',interrupted:'Interrumpido',must_repeat:'Debe repetir',failed:'No aprobado'};
const hiringLabel={pending:'Pendiente',accepted:'Aceptado',returned:'Devuelto'};

window.exportXlsx=async function(){
  if(typeof XLSX==='undefined'){alert('No fue posible cargar Excel.');return}
  try{
    const{data,error}=await OnboardAuth.client.from('of_v_candidate_report').select('*').order('full_name');
    if(error)throw error;
    const rows=data||[];

    // Columnas de modulos: se arman dinamicamente segun lo que exista hoy,
    // ordenadas por module_order, para que el reporte siempre refleje los
    // modulos/examenes vigentes sin tener que tocar este archivo.
    const moduleNames=[];
    rows.forEach(r=>(r.modules||[]).forEach(m=>{if(!moduleNames.includes(m.module))moduleNames.push(m.module)}));
    moduleNames.sort((a,b)=>{
      const oa=rows.flatMap(r=>r.modules||[]).find(m=>m.module===a)?.order||0;
      const ob=rows.flatMap(r=>r.modules||[]).find(m=>m.module===b)?.order||0;
      return oa-ob;
    });

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
};
})();
