// dashboard.js - interacción UI, render preview y export
(function(){
  const $ = (id) => document.getElementById(id);
  // 🔒 CONGELADO: Google Sheets integration (en construcción)
  // const SHEETS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbzQBKG9NtwjsSlth2Yk4g9pP9ui61QAGGg1KapdCjQRvR8A7nnyO4uHYnSi4tX3uoWUgw/exec';

  function setStatus(msg){
    $("status").textContent = msg;
  }

  function overlayShow(title, msg){
    $("ovTitle").textContent = title || "Procesando";
    $("ovMsg").textContent = msg || "Procesando...";
    $("overlay").classList.add("show");
  }
  function overlayMsg(msg){
    $("ovMsg").textContent = msg || "Procesando...";
  }
  function overlayHide(){
    $("overlay").classList.remove("show");
  }

  function renderPreview(outAoa, maxRows=25){
    const head = outAoa[0] || [];
    const body = outAoa.slice(1, 1+maxRows);

    console.log(`📄 RenderPreview: ${head.length} columnas, ${body.length} filas`);

    const thead = $("previewTable").querySelector("thead");
    const tbody = $("previewTable").querySelector("tbody");
    thead.innerHTML = "";
    tbody.innerHTML = "";

    const trh = document.createElement("tr");
    head.forEach(h => {
      const th = document.createElement("th");
      th.textContent = String(h).replace(/\n/g, " / ");
      trh.appendChild(th);
    });
    thead.appendChild(trh);

    body.forEach((row, idx) => {
      const tr = document.createElement("tr");
      row.forEach((v, colIdx) => {
        const td = document.createElement("td");
        if (v instanceof Date){
          td.textContent = fmtYYYYMMDD(v);
        } else {
          td.textContent = (v === null || v === undefined) ? "" : String(v);
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    console.log(`✓ Preview renderizada: ${body.length} filas visibles`);
  }

  function resetAll(){
    APP.workbook = null;
    APP.sourceFileName = "";
    APP.parsedRows = [];
    APP.headerRowIndex = null;
    APP.sheetUsedName = "";
    APP.seasons = [];

    $("filePill").textContent = "No has cargado ningún reporte.";
    $("seasonSel").innerHTML = '<option value="">—</option>';
    $("seasonSel").disabled = true;
    $("btnExport").disabled = true;
    $("btnReset").disabled = true;

    setStatus("Listo.");
  }

  // View switching
  function showView(name){
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    const sel = document.getElementById('view-' + name);
    if (sel) sel.classList.add('active');
    document.querySelectorAll('.nav-item').forEach(n => n.classList.toggle('active', n.dataset.view === name));
    // Congelado: Seguimiento y Dashboard en construcción
    // if (name === 'seguimiento') loadSeguimiento();
    // if (name === 'dashboard') loadDashboard();
  }

  const navProcesar = document.getElementById('navProcesar');
  if (navProcesar) navProcesar.addEventListener('click', () => showView('procesar'));
  const navSeguimiento = document.getElementById('navSeguimiento');
  if (navSeguimiento) navSeguimiento.addEventListener('click', () => alert('Seguimiento está en construcción. Por mientras, enfócate en procesar y guardar los datos.'));
  const navDashboard = document.getElementById('navDashboard');
  if (navDashboard) navDashboard.addEventListener('click', () => alert('Dashboard está en construcción. Por mientras, enfócate en procesar y guardar los datos.'));

  // Exponer funciones al scope global para input.js
  window.setStatus = setStatus;
  window.overlayShow = overlayShow;
  window.overlayMsg = overlayMsg;
  window.overlayHide = overlayHide;
  window.renderPreview = renderPreview;

  // Eventos
  $("btnReset").addEventListener("click", () => resetAll());

  // 🔒 CONGELADO: postAoaToSheets (en construcción)
  /*
  async function postAoaToSheets(outAoa){
    ...
  }
  */

  // ----------------------
  // 🔒 CONGELADO: Seguimiento (en construcción)
  // ----------------------
  /*
  async function fetchRowsFromSheets(offset=0, limit=500){
    const payload = { action: 'list', offset, limit };
    const resp = await fetch(SHEETS_ENDPOINT, { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload) });
    return resp.json();
  }

  async function loadSeguimiento(){
    setStatus('Cargando seguimiento...');
    try{
      const res = await fetchRowsFromSheets(0,1000);
      if (!(res && res.ok)){
        setStatus('Error cargando seguimiento');
        return;
      }
      renderTrackTable(res.rows || []);
      setStatus('Seguimiento cargado');
    } catch(err){
      console.error(err);
      setStatus('Error en seguimiento: ' + (err.message || err));
    }
  }

  function renderTrackTable(rows){
    const thead = document.getElementById('trackTable').querySelector('thead');
    const tbody = document.getElementById('trackTable').querySelector('tbody');
    thead.innerHTML = '';
    tbody.innerHTML = '';
    if (!rows.length) { thead.innerHTML = '<tr><th>—</th></tr>'; return; }

    // headers from first row keys
    const headers = Object.keys(rows[0]);
    const trh = document.createElement('tr');
    trh.appendChild(Object.assign(document.createElement('th'), { textContent: 'ID' }));
    headers.forEach(h => {
      if (h === 'ID') return;
      const th = document.createElement('th'); th.textContent = String(h).replace(/\n/g,' / '); trh.appendChild(th);
    });
    trh.appendChild(Object.assign(document.createElement('th'), { textContent: 'REAL DATE' }));
    thead.appendChild(trh);

    rows.forEach((r) => {
      const tr = document.createElement('tr');
      const idCell = document.createElement('td'); idCell.textContent = r['ID'] || r['Id'] || '';
      tr.appendChild(idCell);
      headers.forEach(h => {
        if (h === 'ID') return;
        const td = document.createElement('td');
        const v = r[h];
        td.textContent = (v instanceof Array ? JSON.stringify(v) : (v instanceof Object ? JSON.stringify(v) : (v === null ? '' : v)));
        tr.appendChild(td);
      });
      // REAL DATE input
      const tdReal = document.createElement('td');
      const inp = document.createElement('input'); inp.type = 'date'; inp.className = 'real-date';
      const val = r['REAL DATE'];
      if (val){
        const d = new Date(val);
        if (!isNaN(d.getTime())) inp.value = d.toISOString().slice(0,10);
      }
      inp.dataset.id = r['ID'];
      inp.addEventListener('change', () => { inp.dataset.dirty = '1'; document.getElementById('btnSaveTrack').disabled = false; });
      tdReal.appendChild(inp);
      tr.appendChild(tdReal);
      tbody.appendChild(tr);
    });
  }

  document.getElementById('btnRefreshTrack').addEventListener('click', () => loadSeguimiento());
  document.getElementById('btnSaveTrack').addEventListener('click', async () => {
    const inputs = Array.from(document.querySelectorAll('#trackTable input.real-date')).filter(i => i.dataset.dirty === '1');
    if (!inputs.length) { alert('No hay cambios para guardar'); return; }
    const updates = inputs.map(i => ({ id: Number(i.dataset.id), val: i.value }));
    try{
      setStatus('Guardando cambios...');
      const promises = updates.map(u => fetch(SHEETS_ENDPOINT, { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ action:'update', id: u.id, fields: { 'REAL DATE': u.val } }) }).then(r=>r.json()));
      const results = await Promise.all(promises);
      const failed = results.filter(r => !(r && r.ok));
      if (failed.length) setStatus('Algunos updates fallaron'); else setStatus('Cambios guardados');
      inputs.forEach(i => { delete i.dataset.dirty; });
      document.getElementById('btnSaveTrack').disabled = true;
      loadSeguimiento();
    } catch(err){ console.error(err); setStatus('Error guardando cambios'); }
  });
  */

  // ----------------------
  // 🔒 CONGELADO: Dashboard (en construcción)
  // ----------------------
  /*
  let chartDiff = null, chartPercent = null;
  async function loadDashboard(){
    setStatus('Cargando datos para dashboard...');
    try{
      const res = await fetchRowsFromSheets(0,1000);
      const rows = (res && res.ok) ? (res.rows || []) : [];
      renderDashboard(rows);
      setStatus('Dashboard listo');
    } catch(err){ console.error(err); setStatus('Error cargando dashboard'); }
  }

  function renderDashboard(rows){
    const diffs = [];
    const byOp = new Map();
    rows.forEach(r => {
      const top = r['TOP DATE'];
      const real = r['REAL DATE'];
      const op = String(r['OP'] || '—');
      const dTop = top ? new Date(top) : null;
      const dReal = real ? new Date(real) : null;
      let diff = null;
      if (dTop && dReal && !isNaN(dTop.getTime()) && !isNaN(dReal.getTime())) diff = Math.round((dReal - dTop)/(1000*60*60*24));
      if (diff !== null) diffs.push(diff);
      const entry = byOp.get(op) || { total:0, done:0 };
      entry.total += 1;
      if (dReal && !isNaN(dReal.getTime())) entry.done += 1;
      byOp.set(op, entry);
    });

    const labels = ['<-10','-10..-1','0','1..7','8..15','>15'];
    const buckets = [0,0,0,0,0,0];
    diffs.forEach(v => {
      if (v < -10) buckets[0]++;
      else if (v < 0) buckets[1]++;
      else if (v === 0) buckets[2]++;
      else if (v <=7) buckets[3]++;
      else if (v <=15) buckets[4]++;
      else buckets[5]++;
    });
    const ctx = document.getElementById('chartDiff').getContext('2d');
    if (chartDiff) chartDiff.destroy();
    chartDiff = new Chart(ctx, { type:'bar', data:{ labels, datasets:[{ label:'Registros', data: buckets, backgroundColor:'#1f6feb' }] }, options:{ responsive:true } });

    const ops = Array.from(byOp.entries()).map(([op, v]) => ({ op, pct: v.total ? Math.round((v.done / v.total)*100) : 0, total: v.total }));
    ops.sort((a,b)=> b.total - a.total);
    const topOps = ops.slice(0,8);
    const ctx2 = document.getElementById('chartPercent').getContext('2d');
    if (chartPercent) chartPercent.destroy();
    chartPercent = new Chart(ctx2, { type:'bar', data:{ labels: topOps.map(x=>x.op), datasets:[{ label:'% Cumplimiento', data: topOps.map(x=>x.pct), backgroundColor:'#10b981' }] }, options:{ responsive:true, scales:{ y:{ beginAtZero:true, max:100 } } } });
  }
  */

  $("btnExport").addEventListener("click", async () => {
    if (!APP.parsedRows.length){ alert("Primero carga el reporte."); return; }
    const season = $("seasonSel").value;
    if (!season){ alert("Selecciona Season."); return; }

    console.log(`🚀 INICIANDO PROCESAMIENTO`);
    console.log(`   Season: ${season}`);
    console.log(`   Total filas parseadas: ${APP.parsedRows.length}`);

    overlayShow("Procesando", "Filtrando Producción + Season...");
    setStatus("Procesando...");
    $("btnExport").disabled = true;
    $("btnPick").disabled = true;
    $("seasonSel").disabled = true;

    try{
      const H = (name) => normHeader(name);
      const hStatusSeason = H("STATUS\nSeason");
      const hSeason = H("Season");
      const hHodLLL = H("HOD\nLULULEMON");
      const hHodCOF = H("HOD\nCOFACO\nPCP");
      const hOP = H("OP");
      const hStyle = H("Style");
      const hColorName = H("Color Name");
      const hColorCode = H("Color Code");

      console.log(`   Headers: [${hStatusSeason}] [${hSeason}] [${hHodLLL}]`);

      overlayMsg("Aplicando filtros...");
      let filtered = APP.parsedRows.filter(r => {
        const ss = normKey(r[hStatusSeason]);
        const okProd = (ss === "produccion" || ss === "producción");
        const okSeason = String(r[hSeason] ?? "").trim() === season;
        return okProd && okSeason;
      });

      console.log(`✓ Después de filtrar por Producción + Season: ${filtered.length} filas`);

      if (!filtered.length) throw new Error("No hay filas para Producción y esa Season.");

      overlayMsg("Normalizando fechas y preparando agrupación...");
      filtered = filtered.map(r => {
        const hod = parseExcelDate(r[hHodLLL]);
        const cof = parseExcelDate(r[hHodCOF]);
        return { ...r, __hod: hod, __cof: cof, __hodTime: hod ? hod.getTime() : null };
      }).filter(r => r.__hodTime !== null);

      console.log(`✓ Después de validar HOD: ${filtered.length} filas`);

      if (!filtered.length) throw new Error("Todas las filas filtradas están sin fecha HOD LULULEMON.");

      overlayMsg("Escogiendo 1er despacho: min(HOD LULULEMON) por Style + Color...");
      const minMap = new Map();
      for (const r of filtered){
        const key = `${String(r[hStyle] ?? "").trim()}||${String(r[hColorName] ?? "").trim()}||${String(r[hColorCode] ?? "").trim()}`;
        const t = r.__hodTime;
        if (t === null) continue;
        if (!minMap.has(key) || t < minMap.get(key)) minMap.set(key, t);
      }

      let filteredMin = filtered.filter(r => {
        const key = `${String(r[hStyle] ?? "").trim()}||${String(r[hColorName] ?? "").trim()}||${String(r[hColorCode] ?? "").trim()}`;
        return r.__hodTime === minMap.get(key);
      });

      console.log(`✓ Después de min por Style + Color: ${filteredMin.length} filas`);

      overlayMsg("Deduplicando por Style + Color (mantener el primero por fecha)...");
      filteredMin.sort((a,b) => (a.__hodTime - b.__hodTime));

      const seen = new Set();
      const finalRows = [];
      for (const r of filteredMin){
        const key = `${String(r[hStyle] ?? "").trim()}||${String(r[hColorName] ?? "").trim()}||${String(r[hColorCode] ?? "").trim()}`;
        if (seen.has(key)) continue;
        seen.add(key);
        finalRows.push(r);
      }

      console.log(`✓ FINAL - Filas únicas: ${finalRows.length}`);

      overlayMsg("Calculando TOP DATE (15 días hábiles, sin sábados, domingos ni miércoles)...");
      const outAoa = [];
      outAoa.push(OUT_COLUMNS);

      for (const r of finalRows){
        const hod = r.__hod;
        const top = adjustNoWednesday(subtractBusinessDays(hod, 15));

        const poRaw = r[H("PO #")];
        const ccRaw = r[H("Color Code")];

        const poVal = toSafeIntOrBlank(poRaw);
        const ccVal = toSafeIntOrBlank(ccRaw);

        // Usar fechas como Date para que Excel las reconozca como fecha (no texto)
        const cofacoPcp = r.__cof instanceof Date ? r.__cof : parseExcelDate(r[H("HOD\nCOFACO\nPCP")]);
        const cofacoPcpVal = cofacoPcp instanceof Date ? cofacoPcp : "";
        const hodVal = hod instanceof Date ? hod : "";
        const topVal = top instanceof Date ? top : "";

        const row = [
          r[H("STATUS\nSeason")] ?? "",
          r[H("Season")] ?? "",
          poVal,
          r[H("OP")] ?? "",
          cofacoPcpVal,
          hodVal,
          r[H("Style")] ?? "",
          r[H("Style Name")] ?? "",
          ccVal,
          r[H("Color Description")] ?? "",
          r[H("Color Name")] ?? "",
          topVal,
          ""  // REAL DATE vacía para que el usuario la llene
        ];
        outAoa.push(row);
      }

      console.log(`✓ AOA generada: ${outAoa.length} filas (1 header + ${finalRows.length} datos)`);

      // expose last exported AOA so it can be downloaded
      APP.lastOutAoa = outAoa;
      
      console.log(`✓ AOA generada: ${outAoa.length} filas (1 header + ${finalRows.length} datos)`);

      const filas = finalRows.length;
      const stylesU = new Set(finalRows.map(r => String(r[normHeader("Style")] ?? "").trim()).filter(Boolean)).size;
      const colorsU = new Set(finalRows.map(r => String(r[normHeader("Color Code")] ?? "").trim()).filter(Boolean)).size;
      const now = new Date();
      const nowStr = `${fmtYYYYMMDD(now)} ${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;

      const _sum = $("summary");
      if (_sum) {
        _sum.textContent =
          `Season: ${season}\n` +
          `STATUS Season: Producción\n` +
          `Filas procesadas: ${filas}\n` +
          `Styles únicos: ${stylesU}\n` +
          `Color Code únicos: ${colorsU}\n` +
          `Hoja leída: ${APP.sheetUsedName}\n` +
          `Encabezado detectado: fila ${APP.headerRowIndex+1}\n` +
          `✅ Listo para descargar\n` +
          `Fecha/Hora: ${nowStr}`;
      }

      console.log(`📊 RESULTADOS FINALES:`);
      console.log(`   Filas: ${filas}, Styles: ${stylesU}, Colors: ${colorsU}`);
      console.log(`   ✓ Listo para descargar`);

      setStatus('✅ Procesado. Descargando Excel...');
      
      // Descargar automáticamente
      setTimeout(() => downloadExcel(), 500);
      
      overlayHide();

    } catch(err){
      overlayHide();
      console.error('🚨 Error general:', err);
      alert("Ocurrió un error:\n" + (err?.message || err));
      setStatus("Error en procesamiento.");
    } finally {
      $("btnExport").disabled = false;
      $("btnPick").disabled = false;
      $("seasonSel").disabled = false;
    }
  });

  // ----------------------
  // ----------------------
  // Descargar Excel con 2 hojas: Tops y Dashboard
  // ----------------------
  function generateDashboardSheet(outAoa) {
    // Generar tablas dinámicas que referencia los datos de la hoja Tops
    // Esto permite que cuando se rellenen datos en Tops, los gráficos se actualicen automáticamente
    
    const dashboardAoa = [];
    const hdrMap = {};
    outAoa[0].forEach((h, i) => { hdrMap[h] = i; });
    
    const totalDataRows = outAoa.length - 1;  // Excluir header
    const topRow = 2;  // Primera fila de datos en Tops (1-indexed, row 2 es la primera data)
    const lastRow = totalDataRows + 1;  // Última fila de datos
    
    // =====================================================
    // SECCIÓN 1: RESUMEN GENERAL
    // =====================================================
    dashboardAoa.push(['RESUMEN GENERAL']);
    dashboardAoa.push(['Métrica', 'Cantidad']);
    dashboardAoa.push(['Total Filas', totalDataRows]);
    dashboardAoa.push(['Total OPs Únicos', `=SUMPRODUCT(1/COUNTIF(Tops!D${topRow}:D${lastRow},Tops!D${topRow}:D${lastRow}&""))`]);
    dashboardAoa.push(['Total Styles', `=SUMPRODUCT(1/COUNTIF(Tops!G${topRow}:G${lastRow},Tops!G${topRow}:G${lastRow}&""))`]);
    dashboardAoa.push(['Total Colors', `=SUMPRODUCT(1/COUNTIF(Tops!I${topRow}:I${lastRow},Tops!I${topRow}:I${lastRow}&""))`]);
    dashboardAoa.push(['Seasons únicos', `=SUMPRODUCT(1/COUNTIF(Tops!B${topRow}:B${lastRow},Tops!B${topRow}:B${lastRow}&""))`]);
    dashboardAoa.push(['']);
    
    // =====================================================
    // SECCIÓN 2: FILAS POR OP (Tabla dinámica)
    // =====================================================
    dashboardAoa.push(['FILAS POR OP']);
    dashboardAoa.push(['OP', 'Cantidad']);
    
    const opMap = new Map();
    for (let i = 1; i < outAoa.length; i++) {
      const op = outAoa[i][hdrMap["OP"]] || 'SIN OP';
      opMap.set(op, (opMap.get(op) || 0) + 1);
    }
    
    // Agregar en orden descendente por cantidad
    for (const [op, count] of Array.from(opMap.entries()).sort((a, b) => b[1] - a[1])) {
      dashboardAoa.push([op, count]);
    }
    
    dashboardAoa.push(['']);
    
    // =====================================================
    // SECCIÓN 3: FILAS POR SEASON
    // =====================================================
    dashboardAoa.push(['FILAS POR SEASON']);
    dashboardAoa.push(['Season', 'Cantidad']);
    
    const seasonMap = new Map();
    for (let i = 1; i < outAoa.length; i++) {
      const season = outAoa[i][hdrMap["Season"]] || 'SIN SEASON';
      seasonMap.set(season, (seasonMap.get(season) || 0) + 1);
    }
    
    for (const [season, count] of Array.from(seasonMap.entries()).sort((a, b) => b[1] - a[1])) {
      dashboardAoa.push([season, count]);
    }
    
    dashboardAoa.push(['']);
    
    // =====================================================
    // SECCIÓN 4: STATUS SEASON
    // =====================================================
    dashboardAoa.push(['FILAS POR STATUS SEASON']);
    dashboardAoa.push(['Status', 'Cantidad']);
    
    const statusMap = new Map();
    for (let i = 1; i < outAoa.length; i++) {
      const status = outAoa[i][hdrMap["STATUS\nSeason"]] || 'SIN STATUS';
      statusMap.set(status, (statusMap.get(status) || 0) + 1);
    }
    
    for (const [status, count] of Array.from(statusMap.entries()).sort((a, b) => b[1] - a[1])) {
      dashboardAoa.push([status, count]);
    }
    
    dashboardAoa.push(['']);
    
    // =====================================================
    // SECCIÓN 5: DIFERENCIA DÍAS (con referencias a Tops)
    // =====================================================
    dashboardAoa.push(['DIFERENCIA DÍAS (TOP DATE - REAL DATE)']);
    dashboardAoa.push(['OP', 'Top Date', 'Real Date', 'Diferencia (días)']);
    
    // Las fórmulas referencian la hoja Tops y se actualizan cuando se edita REAL DATE
    for (let i = 1; i < outAoa.length; i++) {
      const op = outAoa[i][hdrMap["OP"]] || '';
      const excelRow = i + 1;  // +1 porque row 1 es header de Tops, row i+1 es el dato
      
      dashboardAoa.push([
        op,
        { f: `Tops!L${excelRow}` },  // TOP DATE
        { f: `Tops!M${excelRow}` },  // REAL DATE
        { f: `IF(AND(Tops!L${excelRow}<>"",Tops!M${excelRow}<>""),Tops!M${excelRow}-Tops!L${excelRow},"")` }
      ]);
    }
    
    return dashboardAoa;
  }

  function downloadExcel() {
    try {
      if (!APP.lastOutAoa || APP.lastOutAoa.length < 2) {
        alert('No hay datos para descargar. Primero procesa los datos.');
        return;
      }

      console.log('📥 Descargando Excel...');
      
      // Crear workbook con 2 hojas
      const wb = XLSX.utils.book_new();
      
      // Hoja 1: Tops (los datos procesados)
      console.log('  Creando hoja Tops...');
      const topsSheet = XLSX.utils.aoa_to_sheet(APP.lastOutAoa);
      
      // Ajustar ancho de columnas
      topsSheet['!cols'] = computeColWidths(APP.lastOutAoa);
      
      // Agregar filtros automáticos (AutoFilter)
      const lastRow = APP.lastOutAoa.length;  // Incluir header y todas las filas
      const colCount = OUT_COLUMNS.length;
      const lastCol = String.fromCharCode(65 + colCount - 1);  // Convertir número a letra (A, B, C, ... M)
      topsSheet['!autofilter'] = { ref: `A1:${lastCol}${lastRow}` };
      
      XLSX.utils.book_append_sheet(wb, topsSheet, 'Tops');
      console.log('  Hoja Tops creada ✓');
      
      // Hoja 2: Dashboard (tablas y resúmenes)
      console.log('  Creando hoja Dashboard...');
      const dashboardData = generateDashboardSheet(APP.lastOutAoa);
      console.log(`  Dashboard data: ${dashboardData.length} filas`);
      const dashboardSheet = XLSX.utils.aoa_to_sheet(dashboardData);
      dashboardSheet['!cols'] = [
        { wch: 25 },
        { wch: 15 },
        { wch: 15 },
        { wch: 20 }
      ];
      
      XLSX.utils.book_append_sheet(wb, dashboardSheet, 'Dashboard');
      console.log('  Hoja Dashboard creada ✓');
      
      // Descargar archivo
      const now = new Date();
      const timestamp = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;
      const filename = `TOPs_${timestamp}.xlsx`;
      
      console.log(`  Escribiendo archivo: ${filename}...`);
      XLSX.writeFile(wb, filename, { cellDates: true, dateNF: "dd/mm/yyyy" });
      console.log(`✅ Excel descargado: ${filename}`);
      setStatus(`✓ Excel descargado: ${filename}`);
    } catch(err) {
      console.error('🚨 Error en downloadExcel:', err);
      alert('Error al descargar Excel:\n' + (err?.message || err));
      setStatus('Error al descargar Excel.');
    }
  }

  // Inicializar UI
  resetAll();

})();

