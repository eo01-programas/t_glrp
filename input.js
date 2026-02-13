// input.js - manejo de carga de archivo y parsing inicial
(function(){
  const $id = (id) => document.getElementById(id);

  $id("btnPick").addEventListener("click", () => $id("fileInput").click());

  $id("fileInput").addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    if (!window.XLSX){
      alert("No se pudo cargar la librería XLSX (SheetJS). Revisa tu conexión a internet.");
      return;
    }

    if (window.overlayShow) overlayShow("Cargando", "Leyendo Excel y detectando encabezados...");
    if (window.setStatus) setStatus("Cargando reporte...");
    $id("btnPick").disabled = true;

    try{
      APP.sourceFileName = file.name;

      const data = await file.arrayBuffer();
      APP.workbook = XLSX.read(data, {
        type: "array",
        cellDates: true,
        cellNF: true,
        cellText: false
      });

      APP.sheetUsedName = APP.workbook.SheetNames.includes("LLL GR.") ? "LLL GR." : APP.workbook.SheetNames[0];
      const ws = APP.workbook.Sheets[APP.sheetUsedName];

      if (window.overlayMsg) overlayMsg("Convirtiendo hoja a tabla (leyendo documento completo sin filtros)...");
      
      // IMPORTANTE: Leer TODAS las filas incluyendo las OCULTAS/FILTRADAS
      // sheet_to_json respeta los filtros, así que lo hacemos manualmente
      const aoa = [];
      
      // Obtener el rango completo del worksheet
      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
      
      // Reconstruir la AOA fila por fila, sin respetar filtros
      for (let row = range.s.r; row <= range.e.r; row++) {
        const rowData = [];
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = ws[cellAddress];
          if (cell && cell.v !== undefined && cell.v !== null) {
            rowData.push(cell.v);
          } else {
            rowData.push("");
          }
        }
        aoa.push(rowData);
      }

      console.log(`📊 Filas cargadas: ${aoa.length} (ignorando filtros)`);


      APP.headerRowIndex = detectHeaderRow(aoa);

      if (window.overlayMsg) overlayMsg(`Validando columnas requeridas (encabezado detectado en fila ${APP.headerRowIndex+1})...`);
      APP.parsedRows = aoaToObjects(aoa, APP.headerRowIndex);
      
      console.log(`✅ Encabezado en fila ${APP.headerRowIndex}, Filas parseadas: ${APP.parsedRows.length}`);

      // Seasons
      const seasonHeader = normHeader("Season");
      const uniq = new Set();
      APP.parsedRows.forEach(r => {
        const v = r[seasonHeader];
        if (v !== null && v !== undefined && String(v).trim() !== ""){
          uniq.add(String(v).trim());
        }
      });
      APP.seasons = Array.from(uniq).sort((a,b) => a.localeCompare(b, "es"));

      $id("seasonSel").innerHTML = '<option value="">—</option>' + APP.seasons.map(s => `<option value="${s}">${s}</option>`).join("");
      $id("seasonSel").disabled = false;

      $id("btnExport").disabled = false;
      $id("btnReset").disabled = false;

      $id("filePill").textContent = `Archivo: ${APP.sourceFileName}  |  Hoja: ${APP.sheetUsedName}  |  Encabezado: fila ${APP.headerRowIndex+1}`;
      const _s = $id("summary"); if (_s) _s.textContent = "—";
      $id("previewWrap").style.display = "none";
      $id("previewNote").style.display = "none";

      if (window.setStatus) setStatus(`Listo. Seasons detectadas: ${APP.seasons.length}`);
      if (window.overlayHide) overlayHide();
    } catch(err){
      if (window.overlayHide) overlayHide();
      console.error(err);
      alert("No se pudo cargar el Excel:\n" + (err?.message || err));
      if (window.setStatus) setStatus("Error al cargar.");
    } finally {
      $id("btnPick").disabled = false;
      e.target.value = "";
    }
  });

})();
