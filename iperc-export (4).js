// ════════════════════════════════════════════════════
// DIFFUSION MODULE
// ════════════════════════════════════════════════════
// ════════════════════════════════════════════════════
// PDF GENERATION — Landscape A4, professional format
// ════════════════════════════════════════════════════
// ════════════════════════════════════════════════════
// EXPORTAR EXCEL — Formato exacto Arca Continental
// ════════════════════════════════════════════════════
async function generateExcel(){
  const btn = document.getElementById('btn-gen-xlsx');
  if(btn){ btn.disabled=true; btn.innerHTML='<span>⏳</span> Generando…'; }
  try{
    if(!window.XLSX){
      await new Promise((res,rej)=>{
        const s=document.createElement('script');
        s.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
        s.onload=res; s.onerror=()=>rej(new Error('No se pudo cargar SheetJS'));
        document.head.appendChild(s);
      });
    }

    // ── Datos del proyecto ──────────────────────────────────────
    const cli     = (_selectedClient&&_selectedClient.nombre)||'—';
    const fecha   = new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'2-digit',year:'numeric'});
    const desc    = (document.getElementById('scope-text')||{}).value||'—';
    const lugar   = (document.getElementById('p_lugar')||{}).value||'—';
    const area    = (document.getElementById('p_area')||{}).value||'—';
    const elab    = (document.getElementById('p_elaboro')||{}).value||'—';
    const rev     = (document.getElementById('p_reviso')||{}).value||'—';
    const apro    = (document.getElementById('p_aprobo')||{}).value||'—';
    const cod     = (document.getElementById('p_codigo')||{}).value||'—';
    const pers    = (document.getElementById('p_personal')||{}).value||'—';
    const puestos = (document.getElementById('p_puestos')||{}).value||'—';
    const vigFin  = (function(){
      var d=new Date(); d.setMonth(d.getMonth()+9);
      return d.toLocaleDateString('es-MX',{day:'2-digit',month:'2-digit',year:'numeric'});
    })();

    function _cn(t){
      if(!t) return t;
      return (t+'')
        .replace(/\bsegún\s+NOM[\w\-\.]+/gi,'').replace(/\baplicar\s+NOM[\w\-\.]+/gi,'')
        .replace(/\by\s+normativa\s+NOM[\w\-\.]+(?:[,\s]+NOM[\w\-\.]+)*/gi,'')
        .replace(/\bNOM[\-\d]+(?:[\-A-Z]+)?([\-\d]+)?/g,'')
        .replace(/\bOSHA\s+\d{4}\.\d+/g,'').replace(/\s{2,}/g,' ').trim();
    }

    // ── Grupos de actividades ───────────────────────────────────
    const rawActs   = window._rawActividades||[];
    const selRisks  = (state&&state.selectedRisks)||{};
    const actOrder  = rawActs.map(function(a){return a.nombre||a.name;});
    const groups    = [];

    actOrder.forEach(function(actName){
      const risks = selRisks[actName]||[];
      if(!risks.length) return;
      const ra = rawActs.find(function(a){return (a.nombre||a.name)===actName;});
      const subsText = (ra&&Array.isArray(ra.subpasos)&&ra.subpasos.length)
        ? ra.subpasos.map(function(s){
            return _cn((s.paso?s.paso+'. ':'')+(s.personal?'['+s.personal+'] ':'')+s.descripcion);
          }).join('\n')
        : '';
      const parts = [
        _cn(ra&&ra.descripcion?ra.descripcion:''),
        ra&&ra.consideraciones?'VERIFICAR: '+_cn(ra.consideraciones):'',
        ra&&ra.nota?'NOTA: '+ra.nota:'',
        subsText
      ].filter(Boolean);
      groups.push({
        actName:  actName,
        paso:     ra&&ra.paso?ra.paso:'',
        fullDesc: (actName+'\n\n'+parts.join('\n')).trim(),
        risks:    risks
      });
    });

    if(!groups.length){ showToast('⚠️ Primero genera el análisis IPERC', 3000); return; }

    // ── Paleta de colores (exacta del archivo Arca) ─────────────
    const C_RED_TITLE  = 'C00000'; // rojo oscuro encabezado EVALUACION
    const C_YELLOW_HDR = 'FFC000'; // amarillo encabezados
    const C_RED_CTRL   = 'FF0000'; // rojo jerarquía controles
    const C_LIGHT      = 'F2F2F2'; // gris claro valores
    const C_CREAM      = 'FFF2CC'; // crema "Reglas que Salvan Vidas"

    function grColor(gr){
      if(gr>400) return 'FF0000';
      if(gr>200) return 'FFC000';
      if(gr>70)  return 'FFFF00';
      if(gr>20)  return '00B050';
      return              '00B0F0';
    }
    function grLabel(gr){
      if(gr>400) return 'INMINENTE'; if(gr>200) return 'ALTO';
      if(gr>70)  return 'NOTABLE';   if(gr>20)  return 'MODERADO';
      return 'ACEPTABLE';
    }

    // ── Helpers de celda ────────────────────────────────────────
    function S(fill, bold, sz, fc, ha, va, wrap, bstyle){
      var s = {};
      if(fill) s.fill = {patternType:'solid',fgColor:{rgb:fill}};
      s.font = {bold:!!bold, sz:sz||10, color:{rgb:fc||'000000'}};
      s.alignment = {horizontal:ha||'left',vertical:va||'center',wrapText:wrap!==false};
      if(bstyle){
        var b = bstyle==='medium'?{style:'medium',color:{rgb:'000000'}}:{style:'thin',color:{rgb:'000000'}};
        s.border = {top:b,bottom:b,left:b,right:b};
      }
      return s;
    }
    function CE(v, s){ return {v:v||'', t:typeof v==='number'?'n':'s', s:s}; }
    function CF(f, v, s){ return {v:v, t:'n', f:f, s:s}; }

    // ── Hoja IPERC ──────────────────────────────────────────────
    var ws = {};
    var merges = [];
    function M(r1,c1,r2,c2){ merges.push({s:{r:r1-1,c:c1-1},e:{r:r2-1,c:c2-1}}); }
    function W(col,row,cellData){ ws[XLSX.utils.encode_cell({r:row-1,c:col-1})]=cellData; }
    // fill empty merge cells
    function FillMerge(r1,c1,r2,c2,fillHex){
      var b={style:'thin',color:{rgb:'000000'}};
      for(var ri=r1;ri<=r2;ri++) for(var ci=c1;ci<=c2;ci++){
        var addr=XLSX.utils.encode_cell({r:ri-1,c:ci-1});
        if(!ws[addr]) ws[addr]={v:'',t:'s',s:{fill:{patternType:'solid',fgColor:{rgb:fillHex}},border:{top:b,bottom:b,left:b,right:b}}};
      }
    }

    // ── Fila 1: Título ──────────────────────────────────────────
    W(1,1,CE('FORMATO DE IDENTIFICACIÓN DE PELIGROS, EVALUACION DE RIESGO Y CONTROL (IPERC)',
      S(null,true,14,'000000','center','center',false,'medium')));
    M(1,1,1,24); FillMerge(1,1,1,24,C_LIGHT);

    // ── Filas 2-6: Info general ─────────────────────────────────
    var sBrk = S(null,false,10,'000000','left','center',false,'thin');
    var sVal = S(C_LIGHT,false,10,'000000','left','center',true,'thin');

    // Fila 2
    W(1,2,CE('Área de Trabajo:',sBrk)); M(1,2,2,2);
    W(3,2,CE(area,sVal)); M(3,2,4,2);
    W(5,2,CE('Descripción del trabajo a realizar:',sBrk));
    W(6,2,CE(desc,S(C_LIGHT,false,10,'000000','left','center',true,'thin'))); M(6,2,11,2);
    W(12,2,CE('Elaborado por:',sBrk));
    W(13,2,CE(elab,sVal)); M(13,2,15,2);
    W(16,2,CE('Firma:',sBrk)); M(16,2,20,2);
    W(21,2,CE('Fecha:',sBrk));
    W(22,2,CE(fecha,sVal)); M(22,2,24,2);

    // Fila 3
    W(1,3,CE('Tipo de Operación:',sBrk)); M(1,3,2,3);
    W(3,3,CE('☐ Rutinaria          ☐ No Rutinaria',S(C_LIGHT,false,9,'000000','left','center',true,'thin'))); M(3,3,4,3);
    W(5,3,CE('Lugar donde se llevará a cabo la actividad:',sBrk));
    W(6,3,CE(lugar,sVal)); M(6,3,11,3);
    W(12,3,CE('Revisado por:',sBrk));
    W(13,3,CE(rev,sVal)); M(13,3,15,3);
    W(16,3,CE('Firma:',sBrk)); M(16,3,20,3);
    W(21,3,CE('Fecha:',sBrk));
    W(22,3,CE(fecha,sVal)); M(22,3,24,3);

    // Fila 4
    W(1,4,CE('Condiciones de\nOperación:',S(null,false,10,'000000','left','center',true,'thin'))); M(1,4,2,6);
    W(3,4,CE('☐ Normal\n☐ Mantenimiento\n☐ Limpieza\n☐ Cámbio de Formato\n☐ Emergencia\n☐ Construcción\n☐ Instalación / Desmantelamiento\n☐ Otros, Especifique:',S(C_LIGHT,false,9,'000000','left','top',true,'thin'))); M(3,4,3,6);
    W(5,4,CE('Personal que realizará la actividad:',sBrk));
    W(6,4,CE(pers,sVal)); M(6,4,11,4);
    W(12,4,CE('Aprobado por:',sBrk));
    W(13,4,CE(apro,sVal)); M(13,4,15,4);
    W(16,4,CE('Firma:',sBrk)); M(16,4,20,4);
    W(21,4,CE('Fecha:',sBrk));
    W(22,4,CE(fecha,sVal)); M(22,4,24,4);

    // Fila 5
    W(5,5,CE('Puesto que realiza la actividad:',sBrk));
    W(6,5,CE(puestos,sVal)); M(6,5,11,5);
    W(12,5,CE('Equipo de Protección Personal General:',S(null,false,10,'000000','left','center',true,'thin'))); M(12,5,12,6);
    W(13,5,CE('Casco, lentes de seguridad, guantes, zapatos con casquillo, chaleco reflectante.',sVal)); M(13,5,20,5);
    W(21,5,CE('PERIODO DE VIGENCIA DEL IPERC',S(null,true,9,'000000','center','center',true,'thin'))); M(21,5,21,6);
    W(22,5,CE('DEL:',sBrk));
    W(23,5,CE(fecha,sVal)); M(23,5,24,5);

    // Fila 6
    W(5,6,CE('IPERC ID:',S(null,true,10,'000000','left','center',false,'thin')));
    W(6,6,CE('CODIGO: '+cod,sVal)); M(6,6,9,6);
    W(10,6,CE('REV: —',sBrk));
    W(13,6,CE('',sVal)); M(13,6,15,6);
    W(16,6,CE('EPP especial según actividades con riesgo específico (altura, soldadura, espacios confinados).',sVal)); M(16,6,20,6);
    W(22,6,CE('AL:',sBrk));
    W(23,6,CE(vigFin,sVal)); M(23,6,24,6);

    // ── Fila 7: Reglas que Salvan Vidas ────────────────────────
    W(1,7,CE('SELECCIONE LAS REGLAS QUE SALVAN VIDAS QUE SEAN APLICABLES AL TRABAJO:',
      S('000000',true,11,'FFFFFF','left','center',false,'thin')));
    M(1,7,24,7); FillMerge(1,7,24,7,'000000');

    // ── Filas 8-9: Imágenes (placeholder) ──────────────────────
    W(1,8,CE('',S('000000',false,9,'000000','center','center',false,'medium'))); M(1,8,3,9);
    W(4,8,CE('',S('FFFFFF',false,9,'888888','center','center',false,'medium'))); M(4,8,24,9);

    // ── Fila 10: Banda EVALUACION ───────────────────────────────
    W(1,10,CE('EVALUACION',S(C_RED_TITLE,true,14,'FFFFFF','center','center',false,'medium')));
    M(1,10,24,10); FillMerge(1,10,24,10,C_RED_TITLE);

    // ── Filas 11-12: Encabezados ────────────────────────────────
    function HY(col,row,text){ W(col,row,CE(text,S(C_YELLOW_HDR,true,9,'000000','center','center',true,'medium'))); }
    function HR(col,row,text){ W(col,row,CE(text,S(C_RED_CTRL,  true,9,'FFFFFF','center','center',true,'medium'))); }

    // Fila 11
    HY(1,11,'Paso No'); M(1,11,1,12);
    HY(2,11,'ACTIVIDADES DEL TRABAJO PASO A PASO'); M(2,11,3,12);
    HY(4,11,'PELIGRO\nIDENTIFICADO'); M(4,11,4,12);
    HY(5,11,'DESCRIPCION DEL RIESGO ASOCIADO AL PELIGRO'); M(5,11,5,12);
    HY(6,11,'¿QUIEN PODRIA RESULTAR LESIONADO?'); M(6,11,6,12);
    HY(7,11,'RIESGO INHERENTE\n(EVALUAR EL RIESGO SIN CONTROLES)'); M(7,11,10,11);
    HY(11,11,'CLASIFICACION\nDEL RIESGO'); M(11,11,11,12);
    HR(12,11,'MEDIDAS DE CONTROL A IMPLEMENTAR PARA REDUCIR EL GRADO DE RIESGO'); M(12,11,16,11);
    HY(17,11,'RIESGO FINAL\n(RIESGO RESIDUAL)'); M(17,11,20,11);
    HY(21,11,'CLASIFICACION\nDEL RIESGO'); M(21,11,21,12);
    HY(22,11,'EFECTIVIDAD DE LOS CONTROLES'); M(22,11,24,11);

    // Fila 12
    HY(7,12,'C'); HY(8,12,'E'); HY(9,12,'P');
    HY(10,12,'GRADO\nDE RIESGO');
    HR(12,12,'ELIMINACION'); HR(13,12,'SUSTITUCION');
    HR(14,12,'CONTROLES DE\nINGENIERÍA'); HR(15,12,'CONTROLES\nADMINISTRATIVOS');
    HR(16,12,'EPP');
    HY(17,12,'C'); HY(18,12,'E'); HY(19,12,'P'); HY(20,12,'GR');
    HY(22,12,'DEFINICION'); HY(23,12,'EJECUCION'); HY(24,12,'EFECTIVIDAD');

    // ── Filas de datos ──────────────────────────────────────────
    var dataRow = 13;
    var rowHeights = {};
    for(var ri=0;ri<=12;ri++) rowHeights[ri]={
      0:36.75,1:32.65,2:32.65,3:32.65,4:33,5:31.15,
      6:28.5,7:66.75,8:141,  // filas 7,8,9 exactas del original
      9:24,10:34.15,11:45
    }[ri]||20;

    groups.forEach(function(g){
      var n  = g.risks.length;
      var r0 = dataRow;
      var r1 = dataRow + n - 1;

      // Altura de filas de datos
      for(var ri2=r0;ri2<=r1;ri2++) rowHeights[ri2-1] = 80;

      // A: Paso No (merge grupo)
      W(1,r0,CE(g.paso||'',S(null,false,12,'000000','center','center',false,'medium')));
      if(n>1) M(1,r0,1,r1);

      // B-C: Actividad (merge grupo)
      W(2,r0,CE(g.fullDesc,S(null,false,10,'000000','left','top',true,'medium')));
      M(2,r0,3,r1);

      // Riesgos
      g.risks.forEach(function(rsk, ri3){
        var dr    = r0 + ri3;
        var grI   = (rsk.c||0)*(rsk.e||0)*(rsk.p||0);
        var grR   = (rsk.c2||0)*(rsk.e2||0)*(rsk.p2||0);
        var cI    = grColor(grI), cR = grColor(grR);
        var lI    = grLabel(grI), lR = grLabel(grR);
        var sData = S(null,false,10,'000000','left','top',true,'thin');
        var sNum  = S(null,false,12,'000000','center','center',false,'thin');
        var sGrI  = S(cI,true,12,'000000','center','center',false,'thin');
        var sGrR  = S(cR,true,12,'000000','center','center',false,'thin');
        var sLblI = S(cI,true,10,'000000','center','center',false,'thin');
        var sLblR = S(cR,true,10,'000000','center','center',false,'thin');

        W(4,dr,CE(rsk.tipo||'',sData));
        W(5,dr,CE(_cn(rsk.riesgo||''),sData));
        W(6,dr,CE(_cn(rsk.consec||''),sData));
        W(7,dr,CE(rsk.c||0,sNum));
        W(8,dr,CE(rsk.e||0,sNum));
        W(9,dr,CE(rsk.p||0,sNum));

        // J: GR inherente fórmula
        var gR=XLSX.utils.encode_cell({r:dr-1,c:6});
        var hR=XLSX.utils.encode_cell({r:dr-1,c:7});
        var iR=XLSX.utils.encode_cell({r:dr-1,c:8});
        W(10,dr,CF('G'+dr+'*H'+dr+'*I'+dr,grI,sGrI));

        // K: Clasificación inherente
        W(11,dr,CE(lI,sLblI));

        // Controles L-P
        var elim = rsk.elim&&rsk.elim!=='N/A'?_cn(rsk.elim):'N/A';
        var sust = rsk.sust&&rsk.sust!=='N/A'?_cn(rsk.sust):'N/A';
        W(12,dr,CE(elim,sData));
        W(13,dr,CE(sust,sData));
        W(14,dr,CE(_cn(rsk.ingenieria||'N/A'),sData));
        W(15,dr,CE(_cn(rsk.admin||''),sData));
        W(16,dr,CE(_cn(rsk.epp||''),sData));

        // Q,R,S: C,E,P residual
        W(17,dr,CE(rsk.c2||0,sNum));
        W(18,dr,CE(rsk.e2||0,sNum));
        W(19,dr,CE(rsk.p2||0,sNum));

        // T: GR residual fórmula
        W(20,dr,CF('Q'+dr+'*R'+dr+'*S'+dr,grR,sGrR));

        // U: Clasificación residual
        W(21,dr,CE(lR,sLblR));

        // V,W: Definición / Ejecución
        W(22,dr,CE(lI,S(cI,true,10,'000000','center','center',false,'thin')));
        W(23,dr,CE(lR,S(cR,true,10,'000000','center','center',false,'thin')));

        // X: Efectividad IF
        var vC='V'+dr, wC='W'+dr;
        W(24,dr,{v:'ALTO',t:'s',
          f:'IF(AND('+vC+'="ALTO",'+wC+'="ALTO"),"ALTO",IF(AND('+vC+'="ALTO",'+wC+'="MODERADO"),"ALTO",IF(AND('+vC+'="MODERADO",'+wC+'="ALTO"),"ALTO","MODERADO")))',
          s:S(null,true,10,'000000','center','center',false,'thin')});
      });

      dataRow = r1 + 1;
    });

    // ── Config de hoja ──────────────────────────────────────────
    ws['!ref']    = XLSX.utils.encode_range({s:{r:0,c:0},e:{r:dataRow,c:23}});
    ws['!merges'] = merges;

    // Anchos exactos del original Arca
    ws['!cols'] = [
      {wch:7.3},{wch:17.8},{wch:67.53},{wch:14},{wch:28},
      {wch:22},{wch:5.5},{wch:5.5},{wch:5.5},{wch:12.5},
      {wch:18},{wch:14},{wch:14},{wch:24},{wch:35},
      {wch:26},{wch:5.5},{wch:5.5},{wch:5.5},{wch:8},
      {wch:18},{wch:13},{wch:13},{wch:15}
    ];

    // Alturas de fila (exactas del original)
    var rHts = [{hpt:36.75},{hpt:32.65},{hpt:32.65},{hpt:32.65},{hpt:33},
                {hpt:31.15},{hpt:28.5},{hpt:66.75},{hpt:141},{hpt:24},
                {hpt:34.15},{hpt:45}];
    for(var di=12; di<dataRow; di++) rHts.push({hpt:80});
    ws['!rows'] = rHts;

    ws['!pageSetup'] = {orientation:'landscape',paperSize:8,fitToPage:true,fitToWidth:1,fitToHeight:0};

    // ── Hoja CLASIFICACION GR ───────────────────────────────────
    var ws2 = {};
    var GR_ROWS = [
      ['Mayor de 400',   'Riesgo Inminente / Muy Alto','FF0000','Detención inmediata de la actividad peligrosa hasta que se reduzca el riesgo.'],
      ['Entre 200 y 400','Riesgo Alto',               'FFC000','Se requiere corrección inmediata. Actividades en suspensión hasta aplicar controles.'],
      ['Entre 70 y 200', 'Riesgo Notable',            'FFFF00','Corrección necesaria urgente. El nivel de riesgo debe ser revisado periódicamente.'],
      ['Entre 20 y 70',  'Riesgo Moderado',           '00B050','Actividades en esta categoría contienen un nivel de riesgo tolerable con controles.'],
      ['Menos de 20',    'Riesgo Aceptable / Bajo',   '00B0F0','El riesgo en este nivel se considera Aceptable / Bajo.'],
    ];
    function WW2(col,row,cellData){ ws2[XLSX.utils.encode_cell({r:row-1,c:col-1})]=cellData; }
    WW2(1,1,CE('CLASIFICACION DEL RIESGO',S(C_RED_TITLE,true,14,'FFFFFF','center','center',false,'medium')));
    ws2['!merges']=[{s:{r:0,c:0},e:{r:0,c:2}}];
    WW2(1,2,CE('GRADO DE RIESGO',        S(C_YELLOW_HDR,true,10,'000000','center','center',false,'medium')));
    WW2(2,2,CE('CLASIFICACIÓN DEL RIESGO',S(C_YELLOW_HDR,true,10,'000000','center','center',false,'medium')));
    WW2(3,2,CE('ACCIONES FRENTE AL RIESGO',S(C_YELLOW_HDR,true,10,'000000','center','center',false,'medium')));
    GR_ROWS.forEach(function(gr,i){
      WW2(1,i+3,CE(gr[0],S(null,false,10,'000000','left','center',false,'thin')));
      WW2(2,i+3,CE(gr[1],S(gr[2],true,10,'000000','center','center',false,'thin')));
      WW2(3,i+3,CE(gr[3],S(null,false,9,'000000','left','top',true,'thin')));
    });
    ws2['!ref']  = 'A1:C7';
    ws2['!cols'] = [{wch:16},{wch:28},{wch:65}];
    ws2['!rows'] = [{hpt:24},{hpt:20},{hpt:36},{hpt:36},{hpt:36},{hpt:36},{hpt:36}];

    // ── Construir libro y descargar ─────────────────────────────
    const WB = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(WB, ws,  'FORMATO IPERC DIGITAL');
    XLSX.utils.book_append_sheet(WB, ws2, 'CLASIFICACION GR');

    var slug = cli.replace(/[^a-zA-Z0-9]/g,'_').substring(0,20);
    var file = 'IPERC_'+slug+'_'+fecha.replace(/\//g,'-')+'.xlsx';
    XLSX.writeFile(WB, file, {cellStyles:true, bookSST:false});
    showToast('✅ Excel descargado: '+file, 3000);

  } catch(err){
    console.error('generateExcel error:', err);
    showToast('❌ Error al generar Excel: '+err.message, 4000);
  } finally {
    if(btn){ btn.disabled=false; btn.innerHTML='<span>📊</span> Exportar Excel'; }
  }
}

// buildIPERCRows legacy
function buildIPERCRows(){
  var rows2=[], rawActs=window._rawActividades||[], selRisks=(state&&state.selectedRisks)||{};
  Object.entries(selRisks).forEach(function([actName,risks]){
    var ra=rawActs.find(function(a){return (a.nombre||a.name)===actName;});
    (risks||[]).forEach(function(r,i){
      rows2.push({paso:i===0?((ra&&ra.paso)||''):'',act:actName,
        desc:i===0?(ra&&ra.descripcion||actName):'',
        num:r.num||i+1,tipo:r.tipo||'',riesgo:r.riesgo||'',consec:r.consec||'',
        c:r.c,e:r.e,p:r.p,admin:r.admin||'',epp:r.epp||'',c2:r.c2,e2:r.e2,p2:r.p2});
    });
  });
  return rows2;
}


async function generatePDF(){
  // ══════════════════════════════════════════════════════════════════
  // PDF formato ARCA CONTINENTAL — colores y estructura del Excel
  // ══════════════════════════════════════════════════════════════════
  const rows=window._rows||[];
  if(!rows.length){showToast('⚠️ Genera el análisis primero (agrega riesgos).');return;}
  const proj=getProj();
  const c=_selectedClient||CLIENT_CONFIG[CLIENT_CONFIG.length-1];
  const btn=document.getElementById('btn-gen-pdf');
  if(btn){btn.disabled=true;btn.innerHTML='<span>⏳</span> Generando...';}

  try{
    const {jsPDF}=window.jspdf;
    // A3 landscape para que quepan todas las columnas igual que el Excel
    const doc=new jsPDF({orientation:'landscape',unit:'mm',format:'a3'});
    // ── Huella digital en metadatos del PDF ───────────────────────
    const _build = (typeof FTS_BUILD!=='undefined') ? FTS_BUILD : 'FTS-IPERC-v1.5-20260313-A7F3';
    doc.setProperties({
      title:   'IPERC FTS — Análisis de Riesgos Industrial',
      subject: 'IPERC generado por FTS DC-3 Suite',
      author:  'SERVICIOS FTS SA DE CV',
      keywords: _build,
      creator: 'FTS DC-3 ' + _build
    });
    const W=doc.internal.pageSize.getWidth();  // 420mm
    const H=doc.internal.pageSize.getHeight(); // 297mm
    const ML=7, MR=7; // márgenes
    const TW=W-ML-MR; // ancho útil

    // ── Colores exactos Arca ───────────────────────────────────────
    const C_RED_TITLE = [192, 0,   0  ]; // C00000 - banda EVALUACION
    const C_YELLOW    = [255,192,  0  ]; // FFC000 - encabezados col
    const C_RED_CTRL  = [255, 0,   0  ]; // FF0000 - jerarquía controles
    const C_BLACK     = [  0, 0,   0  ]; // fila 7 y A8:C9
    const C_WHITE     = [255,255, 255 ];
    const C_LIGHT     = [242,242, 242 ]; // F2F2F2 - valores info

    function grFill(gr){
      if(gr>400) return [255,  0,  0];   // Inminente - Rojo
      if(gr>200) return [255,192,  0];   // Alto      - Ámbar
      if(gr>70)  return [255,255,  0];   // Notable   - Amarillo
      if(gr>20)  return [  0,176, 80];   // Moderado  - Verde
      return              [  0,176,240];  // Aceptable - Azul
    }
    function grLabel(gr){
      if(gr>400) return 'INMINENTE';
      if(gr>200) return 'ALTO';
      if(gr>70)  return 'NOTABLE';
      if(gr>20)  return 'MODERADO';
      return 'ACEPTABLE';
    }
    function grTextColor(gr){
      // Amarillo y verde tienen texto negro; rojo/ámbar/azul texto negro también
      return [0,0,0];
    }

    // ══════════════════════════════════════════════════════════════
    // FUNCIÓN: dibujar header en cada página
    // ══════════════════════════════════════════════════════════════
    function drawPageHeader(pageNum, totalPages){
      const x=ML, y=6;

      // ── Fila 1: Título principal ────────────────────────────────
      doc.setFillColor(255,255,255);
      doc.setDrawColor(0,0,0);
      doc.setLineWidth(0.4);
      doc.rect(x, y, TW, 8, 'FD');
      doc.setFont('helvetica','bold');
      doc.setFontSize(10);
      doc.setTextColor(0,0,0);
      doc.text('FORMATO DE IDENTIFICACIÓN DE PELIGROS, EVALUACION DE RIESGO Y CONTROL (IPERC)',
        x+TW/2, y+5.2, {align:'center'});

      const y2=y+8; // fila 2 start

      // ── Filas 2-6: Info general (6 filas × ~5mm) ───────────────
      const ROW_H = 5; // mm por fila info
      const INFO_ROWS = 5;
      const INFO_H = ROW_H * INFO_ROWS; // 25mm total

      // Definir zonas columnas para info (proporcional al TW)
      // Col A-B label | C-D value | E label | F-K value | L label | M-O value | P-T firma | U-X fecha/vigencia
      const zA=x,          wA=22;   // Área de Trabajo label
      const zC=zA+wA,      wC=22;   // Área value
      const zE=zC+wC,      wE=30;   // Descripción label
      const zF=zE+wE,      wF=60;   // Descripción value (F-K)
      const zL=zF+wF,      wL=22;   // Elaborado/Revisado/Aprobado label
      const zM=zL+wL,      wM=35;   // Nombre value (M-O)
      const zP=zM+wM,      wP=16;   // Firma label (P)
      const zQ=zP+wP,      wQ=30;   // Firma space (Q-T)
      const zU=zQ+wQ,      wU=22;   // Fecha label (U)
      const zV=zU+wU,      wV=TW-(zU-x+wU); // Fecha value (V-X)

      function infoLbl(txt, cx, cy, w, h){
        doc.setFillColor(...C_WHITE);
        doc.setDrawColor(0,0,0); doc.setLineWidth(0.2);
        doc.rect(cx,cy,w,h,'FD');
        doc.setFont('helvetica','normal'); doc.setFontSize(6.5);
        doc.setTextColor(0,0,0);
        doc.text(txt, cx+1.5, cy+h/2+2, {maxWidth:w-2});
      }
      function infoVal(txt, cx, cy, w, h){
        doc.setFillColor(...C_LIGHT);
        doc.setDrawColor(0,0,0); doc.setLineWidth(0.2);
        doc.rect(cx,cy,w,h,'FD');
        doc.setFont('helvetica','normal'); doc.setFontSize(7);
        doc.setTextColor(0,0,0);
        doc.text(String(txt||'—'), cx+1.5, cy+h/2+2, {maxWidth:w-2});
      }

      // Fila 2: Área de Trabajo
      infoLbl('Área de Trabajo:',  zA, y2,       wA, ROW_H);
      infoVal(proj.area||'—',      zC, y2,       wC, ROW_H);
      infoLbl('Descripción del trabajo a realizar:', zE, y2, wE, ROW_H);
      infoVal(proj.trabajo||'—',   zF, y2,       wF, ROW_H);
      infoLbl('Elaborado por:',    zL, y2,       wL, ROW_H);
      infoVal(proj.elaboro||'—',   zM, y2,       wM, ROW_H);
      infoLbl('Firma:',            zP, y2,       wP, ROW_H);
      doc.setFillColor(...C_WHITE); doc.rect(zQ,y2,wQ,ROW_H,'FD');
      infoLbl('Fecha:',            zU, y2,       wU, ROW_H);
      infoVal(proj.fecha||'—',     zV, y2,       wV, ROW_H);

      // Fila 3: Tipo de Operación
      const y3=y2+ROW_H;
      infoLbl('Tipo de Operación:', zA, y3,      wA, ROW_H);
      infoVal('',                   zC, y3,      wC, ROW_H);
      infoLbl('Lugar donde se llevará a cabo la actividad:', zE, y3, wE, ROW_H);
      infoVal(proj.lugar||'—',      zF, y3,      wF, ROW_H);
      infoLbl('Revisado por:',      zL, y3,      wL, ROW_H);
      infoVal(proj.reviso||'—',     zM, y3,      wM, ROW_H);
      infoLbl('Firma:',             zP, y3,      wP, ROW_H);
      doc.setFillColor(...C_WHITE); doc.rect(zQ,y3,wQ,ROW_H,'FD');
      infoLbl('Fecha:',             zU, y3,      wU, ROW_H);
      infoVal(proj.fecha||'—',      zV, y3,      wV, ROW_H);

      // Filas 4-6: Condiciones (merged A-B vertical)
      const y4=y3+ROW_H;
      // Col A-B merged 3 filas
      doc.setFillColor(...C_WHITE); doc.setDrawColor(0,0,0); doc.setLineWidth(0.2);
      doc.rect(zA, y4, wA, ROW_H*3, 'FD');
      doc.setFont('helvetica','normal'); doc.setFontSize(6.5); doc.setTextColor(0,0,0);
      doc.text('Condiciones de\nOperación:', zA+1.5, y4+4, {maxWidth:wA-2});

      // Col C merged 3 filas — checkboxes
      doc.setFillColor(...C_LIGHT); doc.rect(zC, y4, wC, ROW_H*3, 'FD');
      doc.setFont('helvetica','normal'); doc.setFontSize(5.5); doc.setTextColor(0,0,0);
      const cbLines=['[] Normal','[] Mantenimiento','[] Limpieza','[] Cambio de Formato','[] Emergencia','[] Construccion','[] Instalacion/Desmantelamiento','[] Otros:'];
      cbLines.forEach(function(l,li){ doc.text(l, zC+1.5, y4+2+li*1.8, {maxWidth:wC-2}); });

      // Fila 4
      infoLbl('Personal que realizará la actividad:', zE, y4, wE, ROW_H);
      infoVal(proj.personal||'—',  zF, y4, wF, ROW_H);
      infoLbl('Aprobado por:',     zL, y4, wL, ROW_H);
      infoVal(proj.aprobo||'—',    zM, y4, wM, ROW_H);
      infoLbl('Firma:',            zP, y4, wP, ROW_H);
      doc.setFillColor(...C_WHITE); doc.rect(zQ,y4,wQ,ROW_H,'FD');
      infoLbl('Fecha:',            zU, y4, wU, ROW_H);
      infoVal(proj.fecha||'—',     zV, y4, wV, ROW_H);

      // Fila 5
      const y5=y4+ROW_H;
      infoLbl('Puesto que realiza la actividad:', zE, y5, wE, ROW_H);
      infoVal(proj.puesto||'—',    zF, y5, wF, ROW_H);
      // EPP General label (merged L-L filas 5-6)
      doc.setFillColor(...C_WHITE); doc.rect(zL, y5, wL, ROW_H*2,'FD');
      doc.setFont('helvetica','normal'); doc.setFontSize(6); doc.setTextColor(0,0,0);
      doc.text('Equipo de Protección\nPersonal General:', zL+1, y5+3, {maxWidth:wL-1});
      // EPP valor
      doc.setFillColor(...C_LIGHT); doc.rect(zM, y5, wM+wP+wQ, ROW_H, 'FD');
      doc.setFontSize(6.5);
      doc.text('Casco, lentes, guantes, zapatos con casquillo, chaleco reflectante.',
        zM+1.5, y5+3.5, {maxWidth:wM+wP+wQ-2});
      // Vigencia
      doc.setFillColor(...C_WHITE); doc.rect(zU, y5, wU+wV, ROW_H*2,'FD');
      doc.setFont('helvetica','bold'); doc.setFontSize(6);
      doc.text('PERIODO DE VIGENCIA DEL IPERC', zU+1, y5+3, {maxWidth:wU+wV-2});
      doc.setFont('helvetica','normal'); doc.setFontSize(6.5);
      doc.text('DEL: '+proj.fecha+'  AL: '+proj.vigencia, zU+1, y5+7.5, {maxWidth:wU+wV-2});

      // Fila 6
      const y6=y5+ROW_H;
      infoLbl('IPERC ID:',         zE, y6, wE, ROW_H);
      infoVal('CODIGO: '+(proj.codigo||'—'), zF, y6, wF, ROW_H);
      // EPP especial
      doc.setFillColor(...C_LIGHT); doc.rect(zM, y6, wM+wP+wQ, ROW_H, 'FD');
      doc.setFontSize(6);
      doc.text('EPP especial según actividades con riesgo específico (altura, soldadura, espacios confinados).',
        zM+1.5, y6+3.5, {maxWidth:wM+wP+wQ-2});

      const y7=y4+ROW_H*3; // y6+ROW_H

      // ── Fila 7: NEGRA — Reglas que Salvan Vidas ──────────────────
      doc.setFillColor(...C_BLACK);
      doc.rect(x, y7, TW, 5, 'F');
      doc.setFont('helvetica','bold'); doc.setFontSize(7.5);
      doc.setTextColor(255,255,255);
      doc.text('SELECCIONE LAS REGLAS QUE SALVAN VIDAS QUE SEAN APLICABLES AL TRABAJO:',
        x+3, y7+3.5);

      // ── Filas 8-9: Negro (logos) + Blanco (iconos) ───────────────
      const y8=y7+5;
      const wLogoZone=45; // ancho zona negra (col A-C aprox)
      doc.setFillColor(...C_BLACK);
      doc.rect(x, y8, wLogoZone, 12, 'F');
      // Texto "FTS" en blanco en la zona negra
      doc.setFont('helvetica','bold'); doc.setFontSize(11);
      doc.setTextColor(255,255,255);
      doc.text('FTS', x+wLogoZone/2, y8+7, {align:'center'});
      // Zona blanca iconos
      doc.setFillColor(...C_WHITE);
      doc.setDrawColor(0,0,0); doc.setLineWidth(0.2);
      doc.rect(x+wLogoZone, y8, TW-wLogoZone, 12, 'FD');

      // ── Fila 10: EVALUACION rojo oscuro ─────────────────────────
      const y10=y8+12;
      doc.setFillColor(...C_RED_TITLE);
      doc.rect(x, y10, TW, 6, 'F');
      doc.setFont('helvetica','bold'); doc.setFontSize(11);
      doc.setTextColor(255,255,255);
      doc.text('EVALUACION', x+TW/2, y10+4.2, {align:'center'});

      return y10+6; // Y donde empieza la tabla de riesgos
    }

    // ══════════════════════════════════════════════════════════════
    // CONSTRUIR DATOS DE LA TABLA
    // ══════════════════════════════════════════════════════════════
    const rawActsPdf = window._rawActividades || [];
    function _pdfFullDesc(actName){
      function _stripN(s){ return (s||'').replace(/^\d+[\.-]\s*/,'').trim().toLowerCase(); }
      const aS=_stripN(actName);
      const ra=rawActsPdf.find(function(a){
        return a.nombre===actName||a.name===actName
          ||_stripN(a.nombre||'')===aS||_stripN(a.name||'')===aS
          ||aS.includes(_stripN(a.nombre||'').substring(0,20))
          ||_stripN(a.nombre||'').includes(aS.substring(0,20));
      });
      if(!ra) return actName;
      var parts=[];
      parts.push(ra.nombre||actName);
      if(ra.descripcion) parts.push(_cleanNomRefs(ra.descripcion));
      if(ra.consideraciones) parts.push('VERIFICAR: '+_cleanNomRefs(ra.consideraciones));
      if(ra.nota) parts.push('NOTA CRITICA: '+ra.nota);
      if(Array.isArray(ra.subpasos)&&ra.subpasos.length){
        ra.subpasos.forEach(function(s){
          var q=s.personal?'['+s.personal+'] ':'';
          parts.push((s.paso?s.paso+'. ':'')+q+_cleanNomRefs(s.descripcion||''));
        });
      }
      return parts.join('\n');
    }

    var pdfGroups=[], pdfSeen={};
    rows.forEach(function(r){
      if(!pdfSeen[r.act]){pdfSeen[r.act]=true; pdfGroups.push({act:r.act,risks:[]});}
      pdfGroups[pdfGroups.length-1].risks.push(r);
    });

    // ── Construir body ─────────────────────────────────────────────
    const tableBody=[];
    pdfGroups.forEach(function(group, gi){
      var fullDesc=_pdfFullDesc(group.act);
      var span=group.risks.length;
      group.risks.forEach(function(r, ri){
        var grI=Math.round(r.c*r.e*r.p);
        var grR=Math.round((r.c2||r.c)*(r.e2||r.e)*(r.p2||r.p));
        var lI=grLabel(grI), lR=grLabel(grR);
        var cI=grFill(grI),  cR=grFill(grR);
        var tcI=grTextColor(grI), tcR=grTextColor(grR);
        var ctrl=_cleanNomRefs(
          (r.ingenieria&&r.ingenieria!=='N/A'?r.ingenieria+'\n':'N/A\n')
          +'· '+(r.admin||'—'));

        var row=[];
        if(ri===0){
          // Paso (merged)
          row.push({content:String(gi+1), rowSpan:span,
            styles:{valign:'middle',halign:'center',fontStyle:'bold',fontSize:9,
                    fillColor:C_WHITE,textColor:[0,0,0]}});
          // Actividades (merged)
          row.push({content:fullDesc, rowSpan:span,
            styles:{valign:'top',fontSize:6,fillColor:C_WHITE,
                    textColor:[0,0,0],overflow:'linebreak'}});
        }
        // # riesgo
        row.push({content:String(ri+1),styles:{halign:'center',fillColor:C_WHITE,textColor:[0,0,0]}});
        // Tipo peligro
        row.push({content:r.tipo||'—',styles:{halign:'center',fillColor:C_WHITE}});
        // Descripción riesgo
        row.push({content:_cleanNomRefs(r.riesgo||'—'),styles:{fillColor:C_WHITE}});
        // Consecuencia / ¿Quién?
        row.push({content:_cleanNomRefs(r.consec||'—'),styles:{fillColor:C_WHITE}});
        // C E P GR inherente
        row.push({content:String(r.c),styles:{halign:'center',fillColor:C_WHITE}});
        row.push({content:String(r.e),styles:{halign:'center',fillColor:C_WHITE}});
        row.push({content:String(r.p),styles:{halign:'center',fillColor:C_WHITE}});
        row.push({content:String(grI),
          styles:{halign:'center',fontStyle:'bold',fillColor:cI,textColor:tcI}});
        // Clasificación inherente
        row.push({content:lI,
          styles:{halign:'center',fontStyle:'bold',fontSize:6,fillColor:cI,textColor:tcI}});
        // Controles (Eliminación / Sustitución / Ingeniería / Admin / EPP)
        var elim=_cleanNomRefs(r.elim&&r.elim!=='N/A'?r.elim:'N/A');
        var sust=_cleanNomRefs(r.sust&&r.sust!=='N/A'?r.sust:'N/A');
        var ing=_cleanNomRefs(r.ingenieria||'N/A');
        var adm=_cleanNomRefs(r.admin||'—');
        var epp=_cleanNomRefs(r.epp||'—');
        row.push({content:elim,styles:{fillColor:C_WHITE,fontSize:5.5}});
        row.push({content:sust,styles:{fillColor:C_WHITE,fontSize:5.5}});
        row.push({content:ing, styles:{fillColor:C_WHITE,fontSize:5.5}});
        row.push({content:adm, styles:{fillColor:C_WHITE,fontSize:5.5}});
        row.push({content:epp, styles:{fillColor:C_WHITE,fontSize:5.5}});
        // C E P GR residual
        row.push({content:String(r.c2||r.c),styles:{halign:'center',fillColor:C_WHITE}});
        row.push({content:String(r.e2||r.e),styles:{halign:'center',fillColor:C_WHITE}});
        row.push({content:String(r.p2||r.p),styles:{halign:'center',fillColor:C_WHITE}});
        row.push({content:String(grR),
          styles:{halign:'center',fontStyle:'bold',fillColor:cR,textColor:tcR}});
        // Clasificación residual
        row.push({content:lR,
          styles:{halign:'center',fontStyle:'bold',fontSize:6,fillColor:cR,textColor:tcR}});
        // Efectividad
        row.push({content:r.ef||'ALTO',
          styles:{halign:'center',fontStyle:'bold',fillColor:C_WHITE,textColor:[0,0,0]}});
        tableBody.push(row);
      });
    });

    // ══════════════════════════════════════════════════════════════
    // RENDERIZAR PÁGINA 1 CON HEADER
    // ══════════════════════════════════════════════════════════════
    const yTable=drawPageHeader(1,1);

    // ── Encabezados de columna (2 filas — Arca style) ─────────────
    // Fila 11: amarillo — Fila 12: amarillo + rojo para controles
    const hdrAmbar = {fillColor:C_YELLOW, textColor:[0,0,0], fontStyle:'bold', fontSize:6, halign:'center', lineColor:[0,0,0], lineWidth:0.3};
    const hdrRojo  = {fillColor:C_RED_CTRL, textColor:[255,255,255], fontStyle:'bold', fontSize:6, halign:'center', lineColor:[0,0,0], lineWidth:0.3};

    doc.autoTable({
      startY: yTable,
      margin: {left:ML, right:MR},
      styles:{
        fontSize:5.8, cellPadding:1.2,
        lineColor:[0,0,0], lineWidth:0.25,
        font:'helvetica', textColor:[0,0,0],
        overflow:'linebreak', minCellHeight:6
      },
      headStyles:{
        fillColor:C_YELLOW, textColor:[0,0,0],
        fontStyle:'bold', fontSize:6, halign:'center',
        lineColor:[0,0,0], lineWidth:0.3
      },
      head:[
        // Fila 11
        [
          {content:'Paso No',           rowSpan:2, styles:hdrAmbar},
          {content:'ACTIVIDADES DEL TRABAJO\nPASO A PASO', rowSpan:2, styles:hdrAmbar},
          {content:'#',                 rowSpan:2, styles:hdrAmbar},
          {content:'PELIGRO\nIDENTIFICADO', rowSpan:2, styles:hdrAmbar},
          {content:'DESCRIPCION DEL RIESGO\nASOCIADO AL PELIGRO', rowSpan:2, styles:hdrAmbar},
          {content:'¿QUIEN PODRIA\nRESULTAR LESIONADO?', rowSpan:2, styles:hdrAmbar},
          {content:'RIESGO INHERENTE\n(EVALUAR EL RIESGO SIN CONTROLES)', colSpan:4, styles:hdrAmbar},
          {content:'CLASIFICACION\nDEL RIESGO', rowSpan:2, styles:hdrAmbar},
          {content:'MEDIDAS DE CONTROL A IMPLEMENTAR PARA REDUCIR EL GRADO DE RIESGO', colSpan:5, styles:hdrRojo},
          {content:'RIESGO FINAL\n(RIESGO RESIDUAL)', colSpan:4, styles:hdrAmbar},
          {content:'CLASIFICACION\nDEL RIESGO', rowSpan:2, styles:hdrAmbar},
          {content:'EFECTIVIDAD', rowSpan:2, styles:hdrAmbar},
        ],
        // Fila 12
        [
          {content:'C',styles:hdrAmbar},{content:'E',styles:hdrAmbar},
          {content:'P',styles:hdrAmbar},{content:'GRADO\nDE RIESGO',styles:hdrAmbar},
          {content:'ELIMINACION',styles:hdrRojo},{content:'SUSTITUCION',styles:hdrRojo},
          {content:'CONTROLES DE\nINGENIERIA',styles:hdrRojo},
          {content:'CONTROLES\nADMINISTRATIVOS',styles:hdrRojo},
          {content:'EPP',styles:hdrRojo},
          {content:'C',styles:hdrAmbar},{content:'E',styles:hdrAmbar},
          {content:'P',styles:hdrAmbar},{content:'GR',styles:hdrAmbar},
        ]
      ],
      body: tableBody,
      tableWidth: TW,
      columnStyles:{
        0:{cellWidth:8,  halign:'center'},   // Paso No
        1:{cellWidth:75},                    // Actividades
        2:{cellWidth:5,  halign:'center'},   // #
        3:{cellWidth:12, halign:'center'},   // Tipo peligro
        4:{cellWidth:38},                    // Desc riesgo
        5:{cellWidth:28},                    // Quién
        6:{cellWidth:5,  halign:'center'},   // C
        7:{cellWidth:5,  halign:'center'},   // E
        8:{cellWidth:5,  halign:'center'},   // P
        9:{cellWidth:10, halign:'center'},   // GR inherente
        10:{cellWidth:12,halign:'center'},   // Clasif inherente
        11:{cellWidth:18},                   // Eliminación
        12:{cellWidth:18},                   // Sustitución
        13:{cellWidth:34},                   // Ingeniería
        14:{cellWidth:46},                   // Admin
        15:{cellWidth:40},                   // EPP
        16:{cellWidth:5, halign:'center'},   // C2
        17:{cellWidth:5, halign:'center'},   // E2
        18:{cellWidth:5, halign:'center'},   // P2
        19:{cellWidth:10,halign:'center'},   // GR residual
        20:{cellWidth:12,halign:'center'},   // Clasif residual
        21:{cellWidth:10,halign:'center'},   // Efectividad
      },
      theme:'grid',
      // Header en cada página nueva
      didDrawPage: function(data){
        const pg=doc.internal.getCurrentPageInfo().pageNumber;
        if(pg>1){
          // En páginas 2+ solo dibujar la banda EVALUACION y dejar espacio
          const yH=6;
          doc.setFillColor(...C_RED_TITLE);
          doc.rect(ML, yH, TW, 5, 'F');
          doc.setFont('helvetica','bold'); doc.setFontSize(9);
          doc.setTextColor(255,255,255);
          doc.text('EVALUACION — continuación',ML+TW/2, yH+3.5, {align:'center'});
        }
        // Footer en cada página
        const pgN=doc.internal.getCurrentPageInfo().pageNumber;
        const tot=doc.internal.getNumberOfPages();
        doc.setFontSize(5.5); doc.setTextColor(150,150,150);
        doc.setDrawColor(180,180,180); doc.setLineWidth(0.2);
        doc.line(ML, H-7, W-MR, H-7);
        doc.setFont('helvetica','normal');
        doc.text('SERVICIOS FTS SA DE CV  ·  IPERC FORMATO ARCA CONTINENTAL  ·  Método FINE (C×E×P)  ·  Código: '+(proj.codigo||'—'),
          ML, H-4.5);
        doc.text('Generado: '+new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'long',year:'numeric'})
          +'  ·  Pág. '+pgN+' / '+tot, W-MR, H-4.5, {align:'right'});
      }
    });

    // ── Firmas ────────────────────────────────────────────────────
    var ySig=doc.lastAutoTable.finalY+5;
    if(ySig+22>H-10){doc.addPage(); ySig=15;}
    var sigW=(TW)/3;
    [{lbl:'Elaboró · Segurista FTS',     name:proj.elaboro||''},
     {lbl:'Revisó · Supervisor FTS',     name:proj.reviso||''},
     {lbl:`Aprobó · EHS ${c.nombre}`,    name:proj.aprobo||''}
    ].forEach(function(s,i){
      var sx=ML+(i*sigW)+2;
      doc.setDrawColor(120,120,120); doc.setLineWidth(0.3); doc.setLineDash([2,2]);
      doc.line(sx, ySig+9, sx+sigW-8, ySig+9);
      doc.setLineDash([]);
      doc.setFont('helvetica','normal'); doc.setFontSize(6.5); doc.setTextColor(60,60,60);
      doc.text(s.name||'_______________', sx+(sigW-8)/2, ySig+7, {align:'center',maxWidth:sigW-10});
      doc.setFontSize(6); doc.setTextColor(100,100,100);
      doc.text(s.lbl, sx+(sigW-8)/2, ySig+13, {align:'center'});
    });

    // ── Guardar ───────────────────────────────────────────────────
    var slug=(proj.cliente||c.id||'IPERC').replace(/[^a-zA-Z0-9]/g,'_');
    var fdate=proj.fecha||new Date().toISOString().split('T')[0];
    // ── Huella digital invisible en cada página del PDF ───────────
    // Texto en blanco, tamaño 1pt — no visible pero extraíble con lector PDF
    const _totalPg=doc.internal.getNumberOfPages();
    for(let _pi=1;_pi<=_totalPg;_pi++){
      doc.setPage(_pi);
      doc.setTextColor(255,255,255); doc.setFontSize(1);
      doc.setFont('helvetica','normal');
      doc.text(_build+'|'+fdate+'|p'+_pi, ML, H-0.5);
    }
    doc.save('IPERC_'+slug+'_'+fdate+'.pdf');
    const hint=document.getElementById('pdf-hint');
    if(hint) hint.style.display='block';
    showToast('✅ PDF Arca generado');
  }catch(err){
    console.error('PDF error:',err);
    showToast('⚠️ Error al generar PDF: '+err.message);
  }
  if(btn){btn.disabled=false;btn.innerHTML='<span>📄</span> Generar PDF';}
}

function printIPERC(){
  window.print();
}

function generateDiffusionPDF(){
  const rows=window._rows||[];
  if(!rows.length){showToast('⚠️ Genera el IPERC primero.');return;}
  const proj=getProj();
  const c=_selectedClient||CLIENT_CONFIG[CLIENT_CONFIG.length-1];
  const sorted=[...rows].sort((a,b)=>(b.c*b.e*b.p)-(a.c*a.e*a.p));
  const w=window.open('','_blank');
  w.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8">
  <title>Constancia de Difusión — ${proj.cliente||c.nombre}</title>
  <style>
  *{box-sizing:border-box}body{font-family:Arial,sans-serif;font-size:10px;margin:15mm 12mm;color:#111}
  .header{display:flex;align-items:center;gap:12px;border-bottom:3px solid #D83B01;padding-bottom:8px;margin-bottom:10px}
  .fts-badge{background:#D83B01;color:#fff;font-weight:800;font-size:14px;padding:4px 12px;border-radius:4px}
  .header-info h2{font-size:14px;font-weight:700;margin:0}
  .header-info p{font-size:10px;color:#666;margin:2px 0 0}
  table{width:100%;border-collapse:collapse;margin-top:8px;font-size:9.5px}
  th{background:#1e3a5f;color:#fff;padding:4px 6px;text-align:left}
  td{border:1px solid #ddd;padding:4px 6px;vertical-align:top}
  .lbl{font-weight:600;color:#444;background:#f8f9fa;width:80px}
  .gr-badge{font-size:9px;font-weight:700;padding:2px 5px;border-radius:3px;display:inline-block}
  .gr-inim,.gr-alto{background:#991b1b;color:#fff}.gr-not{background:#fef3c7;color:#333}.gr-mod{background:#dcfce7;color:#333}.gr-acep{background:#f5f5f5;color:#777}
  .sig-area{display:grid;grid-template-columns:1fr 1fr 1fr;gap:24px;margin-top:20px}
  .sig-block{text-align:center;padding-top:36px;border-top:1px solid #888}
  .att-row td{height:20px}
  @page{size:A4;margin:15mm 12mm}@media print{body{margin:0}}
  </style></head><body>
  <div class="header">
    <div class="fts-badge">FTS</div>
    <div class="header-info">
      <h2>CONSTANCIA DE DIFUSIÓN DE RIESGOS</h2>
      <p>${c.nombre} · ${c.formato} · Método FINE NOM-004-STPS</p>
    </div>
  </div>
  <table><tr>
    <td class="lbl">Cliente</td><td>${proj.cliente||'—'}</td>
    <td class="lbl">Trabajo</td><td>${proj.trabajo||'—'}</td>
    <td class="lbl">Código</td><td>${proj.codigo||'—'}</td>
  </tr><tr>
    <td class="lbl">Área</td><td>${proj.area||'—'}</td>
    <td class="lbl">Lugar</td><td>${proj.lugar||'—'}</td>
    <td class="lbl">Fecha</td><td>${proj.fecha||'—'}</td>
  </tr></table>
  <br>
  <table><thead><tr><th style="width:20px">#</th><th style="width:80px">Actividad</th><th>Peligro / Riesgo</th><th style="width:60px">GR / Nivel</th><th>Controles Clave</th><th style="width:90px">EPP</th></tr></thead>
  <tbody>${sorted.map((r,i)=>{
    const gr=r.c*r.e*r.p;const lv=grLevel(gr);
    const ctrl=Array.isArray(r.admin)?r.admin.join(' · '):(r.admin||'—');
    return `<tr><td style="text-align:center">${i+1}</td><td>${r.act}</td><td><strong>${r.riesgo}</strong><br><span style="color:#666;font-size:9px">${r.consec||''}</span></td><td style="text-align:center"><span class="gr-badge ${lv.cls}">${gr}</span></td><td style="font-size:9px">${ctrl}</td><td style="font-size:9px">${r.epp||'—'}</td></tr>`;
  }).join('')}</tbody></table>
  <br>
  <strong style="font-size:10px">Lista de Asistencia — Personal que recibió la difusión</strong>
  <table style="margin-top:5px"><thead><tr><th style="width:20px">#</th><th>Nombre Completo</th><th style="width:90px">Puesto</th><th style="width:70px">No. Empleado</th><th style="width:80px">Firma</th></tr></thead>
  <tbody>${Array.from({length:14},(_,i)=>`<tr class="att-row"><td style="text-align:center">${i+1}</td><td></td><td></td><td></td><td></td></tr>`).join('')}</tbody></table>
  <div class="sig-area">
    <div class="sig-block"><div>${proj.elaboro||'_______________'}</div><div style="color:#666;font-size:9px;margin-top:3px">Elaboró · Segurista FTS</div></div>
    <div class="sig-block"><div>${proj.reviso||'_______________'}</div><div style="color:#666;font-size:9px;margin-top:3px">Revisó · Supervisor FTS</div></div>
    <div class="sig-block"><div>${proj.aprobo||'_______________'}</div><div style="color:#666;font-size:9px;margin-top:3px">Aprobó · EHS ${c.nombre}</div></div>
  </div>
  <div style="margin-top:16px;padding-top:6px;border-top:1px solid #eee;font-size:8px;color:#aaa;text-align:center">
    SERVICIOS FTS SA DE CV · Análisis de Riesgos · Método FINE NOM-004-STPS · Generado ${new Date().toLocaleDateString('es-MX')}
  </div>
  <script>window.onload=()=>{setTimeout(()=>{window.print();},400)}<\/script>
  </body></html>`);
  w.document.close();
}

