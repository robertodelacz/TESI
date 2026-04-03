/**
 * =====================================================
 * DASHBOARD TESI - FINANCIERA CUALLI
 * Archivo: tesi-functions.gs
 * =====================================================
 */

/**
 * Sirve el dashboard como aplicación web embebible
 */
function doGet(e) {
  var template = HtmlService.createHtmlOutputFromFile('dashboard-tesi');
  return template
    .setTitle('Dashboard TESI - Financiera Cualli')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Función auxiliar para obtener la URL de implementación
 */
function obtenerURLDashboard() {
  var url = ScriptApp.getService().getUrl();
  Logger.log('URL del Dashboard: ' + url);
  return url;
}


// =====================================================
// FUNCIÓN OPTIMIZADA - AÑOS Y MESES EN UNA SOLA LECTURA
// =====================================================

/**
 * Obtiene todos los años y meses disponibles en UNA SOLA lectura
 * @returns {Object} { "2025": [{numero: 1, nombre: "Enero"}, ...], "2024": [...] }
 */
function obtenerAniosYMesesDisponibles() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TESI");
  const datos = hoja.getDataRange().getValues();
  const idxFecha = datos[0].indexOf("Fecha de Firma");

  const monthNames = [
    '', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];

  const resultado = {};

  for (let i = 1; i < datos.length; i++) {
    const fecha = new Date(datos[i][idxFecha]);
    if (isValidDate(fecha)) {
      const anio = fecha.getFullYear();
      const mes = fecha.getMonth() + 1;
      if (!resultado[anio]) resultado[anio] = new Set();
      resultado[anio].add(mes);
    }
  }

  const resultadoFinal = {};
  Object.keys(resultado).forEach(anio => {
    resultadoFinal[anio] = Array.from(resultado[anio])
      .sort((a, b) => a - b)
      .map(mes => ({ numero: mes, nombre: monthNames[mes] }));
  });

  return resultadoFinal;
}


// =====================================================
// FUNCIONES ORIGINALES (FALLBACK)
// =====================================================

function obtenerAniosDisponibles() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TESI");
  const datos = hoja.getDataRange().getValues();
  const idxFecha = datos[0].indexOf("Fecha de Firma");
  const years = new Set();
  for (let i = 1; i < datos.length; i++) {
    const fecha = new Date(datos[i][idxFecha]);
    if (isValidDate(fecha)) years.add(fecha.getFullYear());
  }
  return Array.from(years);
}

function obtenerMesesDisponiblesPorAnio(anio) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TESI");
  const datos = hoja.getDataRange().getValues();
  const idxFecha = datos[0].indexOf("Fecha de Firma");
  const meses = new Set();
  const monthNames = [
    {numero:1,nombre:"Enero"},{numero:2,nombre:"Febrero"},{numero:3,nombre:"Marzo"},
    {numero:4,nombre:"Abril"},{numero:5,nombre:"Mayo"},{numero:6,nombre:"Junio"},
    {numero:7,nombre:"Julio"},{numero:8,nombre:"Agosto"},{numero:9,nombre:"Septiembre"},
    {numero:10,nombre:"Octubre"},{numero:11,nombre:"Noviembre"},{numero:12,nombre:"Diciembre"}
  ];
  for (let i = 1; i < datos.length; i++) {
    const fecha = new Date(datos[i][idxFecha]);
    if (isValidDate(fecha) && fecha.getFullYear() == anio) meses.add(fecha.getMonth() + 1);
  }
  return monthNames.filter(month => meses.has(month.numero));
}


// =====================================================
// UTILIDAD: VALIDACIÓN DE FECHA
// =====================================================

function isValidDate(date) {
  return date instanceof Date && !isNaN(date.getTime());
}


// =====================================================
// FUNCIÓN PRINCIPAL: OBTENER DATOS TESI POR PERÍODO
// =====================================================

/**
 * Retorna datos TESI para un mes/año específico
 * @param {string} periodo - Formato "YYYY-MM" (ej: "2025-11")
 */
function obtenerTESIPorPeriodo(periodo) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TESI");
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0];
  const filas = datos.slice(1);

  const idxID               = encabezados.indexOf("ID Oportunidad");
  const idxFecha            = encabezados.indexOf("Fecha de Firma");
  const idxTipo             = encabezados.indexOf("Tipo");
  const idxEstatus          = encabezados.indexOf("Estatus");
  const idxAreas            = encabezados.indexOf("Áreas con Incidencias");
  const idxTotalIncidencias = encabezados.indexOf("Total Incidencias");
  const idxDetalle          = encabezados.indexOf("Detalle incidencias"); // col K
  const idxEtapaCierre      = encabezados.indexOf("Etapa Cierre");        // col Ba

  const [anio, mes] = periodo.split("-").map(Number);

  let total          = 0;
  let sinIncidencias = 0;
  let conIncidencias = 0;
  let revision       = 0;
  let totalTESI      = 0;
  let tesiCompletos  = 0;

  // ── Conteo por tipo ──
  const conteoPorTipo = {
    'Normal':       { total: 0, conIncidencias: 0 },
    'Colocación 0': { total: 0, conIncidencias: 0 },
    'Renovación':   { total: 0, conIncidencias: 0 },
    'Reestructura': { total: 0, conIncidencias: 0 },
    'Ampliación':   { total: 0, conIncidencias: 0 }
  };

  // ── Conteo etapas de cierre (solo completos) ──
  const etapasCierre = {
    '1a Revision': 0,
    '2a Revision': 0,
    'Final':       0
  };

  const listaIncidencias       = [];
  const listaTESIuniverso      = [];
  const listaCompletosUniverso = [];

  filas.forEach(fila => {
    const fechaRaw   = fila[idxFecha];
    const estatusRaw = (fila[idxEstatus] || "").toString().trim().toLowerCase();
    const tipoRaw    = (fila[idxTipo]    || "").toString().trim();
    const areasRaw   = (fila[idxAreas]   || "").toString().trim().toLowerCase();

    const fecha = new Date(fechaRaw);
    if (isNaN(fecha)) return;

    if (fecha.getFullYear() === anio && fecha.getMonth() + 1 === mes) {
      total++;

      const fechaTexto = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");

      if (estatusRaw === "completo") {
        sinIncidencias++;

        // Etapa de cierre
        const etapaRaw = (fila[idxEtapaCierre] || "").toString().trim().toLowerCase();
        if      (etapaRaw.includes("1a") || etapaRaw.includes("primera")) etapasCierre['1a Revision']++;
        else if (etapaRaw.includes("2a") || etapaRaw.includes("segunda")) etapasCierre['2a Revision']++;
        else if (etapaRaw.includes("final"))                               etapasCierre['Final']++;

        listaCompletosUniverso.push({ id: fila[idxID], fecha: fechaTexto });

      } else if (estatusRaw === "incidencias") {
        conIncidencias++;
        if (!areasRaw.includes("sin incidencias")) {
          listaIncidencias.push({
            id:               fila[idxID],
            fecha:            fechaTexto,
            tipo:             normalizarTipo(tipoRaw),
            areas:            fila[idxAreas],
            totalIncidencias: fila[idxTotalIncidencias] || 0,
            documentos:       parsearDetalleIncidencias((fila[idxDetalle] || "").toString())
          });
        }

      } else if (estatusRaw === "revision") {
        revision++;
      }

      // Conteo por tipo (todos los estatus)
      const tipoNorm = normalizarTipo(tipoRaw);
      if (conteoPorTipo[tipoNorm]) {
        conteoPorTipo[tipoNorm].total++;
        if (estatusRaw === "incidencias") conteoPorTipo[tipoNorm].conIncidencias++;
      }

      // TESI
      const esTESI = tipoRaw.toLowerCase() === "normal" ||
                     tipoRaw.toLowerCase() === "colocación cero" ||
                     tipoRaw.toLowerCase() === "colocacion cero";
      if (esTESI) {
        totalTESI++;
        listaTESIuniverso.push({ id: fila[idxID], fecha: fechaTexto, estatus: estatusRaw });
        if (estatusRaw === "completo") tesiCompletos++;
      }
    }
  });

  const tesi = totalTESI > 0 ? (tesiCompletos / totalTESI) * 100 : 0;

  let bono = "Sin bono";
  if      (tesi >= 95) bono = "Bono máximo / 15 días";
  else if (tesi >= 90) bono = "Bono medio / 10 días";
  else if (tesi >= 85) bono = "Bono mínimo / 5 días";

  return {
    total,
    sinIncidencias,
    conIncidencias,
    revision,
    conteoPorTipo,        // ← nuevo
    etapasCierre,         // ← nuevo
    tesiTotal:            totalTESI,
    tesiCompletos:        tesiCompletos,
    tesi:                 Math.round(tesi),
    bono,
    incidencias:          listaIncidencias,
    tesiCalificados:      listaTESIuniverso,
    completosUniverso:    listaCompletosUniverso
  };
}


// =====================================================
// FUNCIÓN PARA EL DASHBOARD: TODOS LOS PERÍODOS
// =====================================================

/**
 * Devuelve todos los períodos con sus datos.
 * Llamado desde el HTML via google.script.run.getDatosTESI()
 */
function getDatosTESI() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TESI");
  const datos = hoja.getDataRange().getValues();
  const idxFecha = datos[0].indexOf("Fecha de Firma");

  const periodos = new Set();
  for (let i = 1; i < datos.length; i++) {
    const fecha = new Date(datos[i][idxFecha]);
    if (isValidDate(fecha)) {
      const p = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM");
      periodos.add(p);
    }
  }

  const resultado = {};
  [...periodos].sort().forEach(p => {
    try {
      resultado[p] = obtenerTESIPorPeriodo(p);
    } catch (e) {
      console.error(`Error en período ${p}:`, e);
    }
  });

  return resultado;
}

/**
 * Carga datos de UN solo año — más rápido que cargar todo
 * Llamado desde HTML via google.script.run.getDatosTESIAnio(anio)
 */
function getDatosTESIAnio(anio) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TESI");
  const datos = hoja.getDataRange().getValues();
  const idxFecha = datos[0].indexOf("Fecha de Firma");

  const periodos = new Set();
  for (let i = 1; i < datos.length; i++) {
    const fecha = new Date(datos[i][idxFecha]);
    if (isValidDate(fecha) && fecha.getFullYear() === parseInt(anio)) {
      const p = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM");
      periodos.add(p);
    }
  }

  const resultado = {};
  [...periodos].sort().forEach(p => {
    try { resultado[p] = obtenerTESIPorPeriodo(p); }
    catch(e) { console.error(`Error período ${p}:`, e); }
  });

  return resultado;
}


// =====================================================
// HELPERS
// =====================================================

/**
 * Normaliza el tipo de expediente a categorías estándar
 */
function normalizarTipo(tipo) {
  const t = (tipo || "").toLowerCase().trim();
  if (t === "normal")                                     return "Normal";
  if (t === "colocación cero" || t === "colocacion cero") return "Colocación 0";
  if (t.includes("renov"))                                return "Renovación";
  if (t.includes("reestruc"))                             return "Reestructura";
  if (t.includes("amplia"))                               return "Ampliación";
  return "Normal";
}

/**
 * Utilidad: Obtener nombre del mes
 */
function obtenerNombreMes(numeroMes) {
  const meses = [
    '', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];
  return meses[numeroMes] || '';
}


function obtenerColorArea(area) {
  const colores = {
    'Comercial': { bg:'#fff7ed', border:'#fed7aa', texto:'#c2410c' },
    'Cartera':   { bg:'#fdf2f8', border:'#fbcfe8', texto:'#be185d' },
    'Crédito':   { bg:'#ecfeff', border:'#a5f3fc', texto:'#0e7490' },
    'Jurídico':  { bg:'#f7fee7', border:'#d9f99d', texto:'#3f6212' },
    'PLD':       { bg:'#faf5ff', border:'#e9d5ff', texto:'#7c3aed' }
  };
  const key = Object.keys(colores).find(k =>
    area.toLowerCase().replace('í','i').includes(k.toLowerCase().replace('í','i'))
  );
  return colores[key] || { bg:'#f8fafc', border:'#e2e8f0', texto:'#475569' };
}

// =====================================================
// TEST
// =====================================================

function testRendimiento() {
  console.log('=== TEST DE RENDIMIENTO ===\n');

  console.log('1. Probando obtenerAniosYMesesDisponibles()...');
  let inicio = new Date();
  const resultadoOptimizado = obtenerAniosYMesesDisponibles();
  console.log(`   Tiempo: ${new Date() - inicio} ms`);
  console.log(`   Años encontrados: ${Object.keys(resultadoOptimizado).join(', ')}`);

  console.log('\n2. Probando getDatosTESI()...');
  inicio = new Date();
  const todos = getDatosTESI();
  console.log(`   Tiempo: ${new Date() - inicio} ms`);
  console.log(`   Períodos cargados: ${Object.keys(todos).join(', ')}`);

  // Mostrar muestra del primer período
  const primerPeriodo = Object.keys(todos)[0];
  if (primerPeriodo) {
    const d = todos[primerPeriodo];
    console.log(`\n3. Muestra período ${primerPeriodo}:`);
    console.log(`   Total: ${d.total} | Completos: ${d.sinIncidencias} | Con inc: ${d.conIncidencias}`);
    console.log(`   Tipos: ${JSON.stringify(d.conteoPorTipo)}`);
    console.log(`   Etapas cierre: ${JSON.stringify(d.etapasCierre)}`);
    if (d.incidencias.length > 0) {
      console.log(`   Primer expediente con incidencias:`);
      console.log(`   - ID: ${d.incidencias[0].id}`);
      console.log(`   - Tipo: ${d.incidencias[0].tipo}`);
      console.log(`   - Documentos: ${d.incidencias[0].documentos.length}`);
      if (d.incidencias[0].documentos.length > 0) {
        console.log(`   - Primer doc: ${JSON.stringify(d.incidencias[0].documentos[0])}`);
      }
    }
  }
}


// =====================================================
// SISTEMA DE EMAIL INFOGRÁFICO
// =====================================================

function generarHTMLEmailInfografia(datos, periodo) {
  const [anio, mes] = periodo.split("-");
  const nombreMes = obtenerNombreMes(parseInt(mes));

  let html = HtmlService.createHtmlOutputFromFile('EmailInfografia').getContent();

  const hora = new Date().getHours();
  let saludo;
  if      (hora >= 5  && hora < 12) saludo = "Estimado equipo, buenos días:";
  else if (hora >= 12 && hora < 19) saludo = "Estimado equipo, buenas tardes:";
  else                               saludo = "Estimado equipo, buenas noches:";

  const ICONOS = {
    logo:      'https://cualli.mx/wp-content/uploads/2022/07/cualli-bl@3x.png',
    comercial: 'https://cdn-icons-png.freepik.com/512/14598/14598062.png',
    cartera:   'https://cdn-icons-png.freepik.com/512/8376/8376410.png',
    credito:   'https://cdn-icons-png.freepik.com/512/1674/1674640.png',
    juridico:  'https://cdn-icons-png.freepik.com/512/4252/4252349.png',
    pld:       'https://cdn-icons-png.freepik.com/512/3359/3359520.png'
  };

  const ICONOS_KPI = {
    total:       'https://cdn-icons-png.freepik.com/512/4371/4371079.png',
    completos:   'https://cdn-icons-png.freepik.com/512/2724/2724742.png',
    incidencias: 'https://cdn-icons-png.freepik.com/512/7100/7100197.png'
  };

  const total        = datos.total || 1;
  const completos    = datos.sinIncidencias || 0;
  const incidencias  = datos.conIncidencias || 0;
  const pctCompletos   = Math.round((completos   / total) * 100);
  const pctIncidencias = Math.round((incidencias / total) * 100);

  const graficoBarras               = generarGraficoBarrasPorAreaEjecutivo(datos, nombreMes, anio, ICONOS);
  const seccionExpedientesIncidencias = generarSeccionExpedientesIncidenciasEjecutivo(datos);
  const seccionDetalleAreas         = generarSeccionDetalleAreasEjecutivo(datos, ICONOS);

  const fechaGen = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  html = html
    .replace(/{{LOGO_URL}}/g,                      ICONOS.logo)
    .replace(/{{ICONO_KPI_TOTAL}}/g,               ICONOS_KPI.total)
    .replace(/{{ICONO_KPI_COMPLETOS}}/g,            ICONOS_KPI.completos)
    .replace(/{{ICONO_KPI_INCIDENCIAS}}/g,          ICONOS_KPI.incidencias)
    .replace(/{{MES}}/g,                            nombreMes.toUpperCase())
    .replace(/{{MES_COMPLETO}}/g,                   nombreMes)
    .replace(/{{ANIO}}/g,                           anio)
    .replace(/{{SALUDO}}/g,                         saludo)
    .replace(/{{TOTAL_EVALUADOS}}/g,                total)
    .replace(/{{COMPLETOS}}/g,                      completos)
    .replace(/{{INCIDENCIAS}}/g,                    incidencias)
    .replace(/{{PCT_COMPLETOS}}/g,                  pctCompletos)
    .replace(/{{PCT_INCIDENCIAS}}/g,                pctIncidencias)
    .replace(/{{SECCION_GRAFICO_BARRAS}}/g,         graficoBarras)
    .replace(/{{SECCION_EXPEDIENTES_INCIDENCIAS}}/g, seccionExpedientesIncidencias)
    .replace(/{{SECCION_DETALLE_AREAS}}/g,          seccionDetalleAreas)
    .replace(/{{FECHA_GENERACION}}/g,               fechaGen);

  return html;
}


function generarGraficoBarrasPorAreaEjecutivo(datos, nombreMes, anio, ICONOS) {
  const areas = {
    'Comercial': { color:'#f97316', colorClaro:'#fb923c', colorOscuro:'#ea580c', icono:ICONOS.comercial, label:'Comercial' },
    'Cartera':   { color:'#ec4899', colorClaro:'#f472b6', colorOscuro:'#db2777', icono:ICONOS.cartera,   label:'Cartera'   },
    'Crédito':   { color:'#06b6d4', colorClaro:'#22d3ee', colorOscuro:'#0891b2', icono:ICONOS.credito,   label:'Crédito'   },
    'Jurídico':  { color:'#84cc16', colorClaro:'#C7FF6C', colorOscuro:'#0DA31A', icono:ICONOS.juridico,  label:'Jurídico'  },
    'PLD':       { color:'#a855f7', colorClaro:'#c084fc', colorOscuro:'#9333ea', icono:ICONOS.pld,       label:'PLD'       }
  };

  const totalesPorArea = {};
  Object.keys(areas).forEach(area => totalesPorArea[area] = 0);

  if (datos.incidencias && datos.incidencias.length > 0) {
    datos.incidencias.forEach(row => {
      const texto = row.areas || '';
      Object.keys(areas).forEach(nombreArea => {
        const regex = new RegExp(`${nombreArea.replace('í','[íi]')}\\s*:\\s*(\\d+)`, 'i');
        const match = texto.match(regex);
        if (match) totalesPorArea[nombreArea] += parseInt(match[1]) || 0;
      });
    });
  }

  const totalIncidencias = Object.values(totalesPorArea).reduce((a, b) => a + b, 0);

  if (totalIncidencias === 0) {
    return `
      <tr><td style="padding: 0 60px 40px;">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0"
               style="background:linear-gradient(135deg,#f0fdf4,#dcfce7);border-radius:16px;border:1px solid #86efac;">
          <tr><td style="padding:40px;">
            <table role="presentation" width="100%" cellspacing="0" cellpadding="0"><tr>
              <td width="80" align="center" valign="middle" style="padding-right:25px;">
                <table role="presentation" cellspacing="0" cellpadding="0"
                       style="width:80px;height:80px;background:linear-gradient(135deg,#059669,#10b981);border-radius:50%;">
                  <tr><td align="center" valign="middle">
                    <img src="https://cdn-icons-png.freepik.com/512/294/294432.png" width="42" height="42" style="display:block;"/>
                  </td></tr>
                </table>
              </td>
              <td align="left" valign="middle">
                <h3 style="margin:0 0 8px;font-size:22px;font-weight:800;color:#065f46;font-family:'IBM Plex Sans',sans-serif;">
                  Desempeño Excelente
                </h3>
                <p style="margin:0;font-size:15px;color:#047857;font-weight:500;">
                  No se registraron incidencias en este período
                </p>
              </td>
            </tr></table>
          </td></tr>
        </table>
      </td></tr>`;
  }

  let filasHTML = '';
  let index = 0;
  const areasConIncidencias = Object.keys(areas).filter(a => totalesPorArea[a] > 0);

  Object.keys(areas).forEach(nombreArea => {
    const areaInfo = areas[nombreArea];
    const valor    = totalesPorArea[nombreArea];
    if (valor > 0) {
      const porcentaje  = Math.round((valor / totalIncidencias) * 100);
      const bgColor     = index % 2 === 0 ? '#ffffff' : '#f8fafc';
      const borderColor = index === areasConIncidencias.length - 1 ? 'transparent' : '#e2e8f0';
      filasHTML += `
        <tr>
          <td style="padding:14px 20px;border-bottom:1px solid ${borderColor};background:${bgColor};">
            <table role="presentation" cellspacing="0" cellpadding="0"><tr>
              <td style="padding-right:12px;">
                <img src="${areaInfo.icono}" alt="${areaInfo.label}" style="width:36px;height:36px;display:block;"/>
              </td>
              <td><div style="font-size:14px;font-weight:700;color:#1a202c;font-family:'IBM Plex Sans',sans-serif;">${areaInfo.label}</div></td>
            </tr></table>
          </td>
          <td style="padding:14px 20px;border-bottom:1px solid ${borderColor};background:${bgColor};text-align:center;">
            <table role="presentation" cellspacing="0" cellpadding="0" align="center"
                   style="width:44px;height:44px;background:linear-gradient(135deg,${areaInfo.colorClaro},${areaInfo.color});border-radius:50%;border:5px solid ${areaInfo.colorOscuro};">
              <tr><td align="center" valign="middle"
                      style="color:#fff;font-size:18px;font-weight:900;font-family:'IBM Plex Sans',sans-serif;text-shadow:0 1px 2px rgba(0,0,0,0.1);">
                ${valor}
              </td></tr>
            </table>
          </td>
          <td style="padding:14px 20px;border-bottom:1px solid ${borderColor};background:${bgColor};">
            <div style="color:#4a5568;font-size:13px;font-weight:500;line-height:1.5;">
              ${valor} incidencia${valor!==1?'s':''} registrada${valor!==1?'s':''} en el período, representando el ${porcentaje}% del total
            </div>
          </td>
        </tr>`;
      index++;
    }
  });

  return `
    <tr><td style="padding:0 60px 50px;">
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
        <tr><td align="center" style="padding-bottom:32px;">
          <h2 style="margin:0 0 8px;font-size:22px;font-weight:800;color:#1a202c;font-family:'IBM Plex Sans',sans-serif;letter-spacing:-0.5px;">
            Distribución de Incidencias por Área
          </h2>
          <p style="margin:0;font-size:13px;color:#64748b;font-weight:500;">
            Período: ${nombreMes} ${anio} • Total: ${totalIncidencias} incidencias
          </p>
        </td></tr>
      </table>
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0"
             style="background:#fff;border:1px solid #e2e8f0;border-radius:6px;">
        <thead><tr style="background:#f8fafc;">
          <th style="padding:14px 20px;text-align:left;color:#4a5568;font-size:11px;text-transform:uppercase;letter-spacing:1px;font-weight:700;font-family:'Inter',sans-serif;border-bottom:2px solid #cbd5e1;width:30%;">Área</th>
          <th style="padding:14px 20px;text-align:center;color:#4a5568;font-size:11px;text-transform:uppercase;letter-spacing:1px;font-weight:700;font-family:'Inter',sans-serif;border-bottom:2px solid #cbd5e1;width:15%;">Incidencias</th>
          <th style="padding:14px 20px;text-align:left;color:#4a5568;font-size:11px;text-transform:uppercase;letter-spacing:1px;font-weight:700;font-family:'Inter',sans-serif;border-bottom:2px solid #cbd5e1;width:55%;">Descripción</th>
        </tr></thead>
        <tbody>${filasHTML}</tbody>
      </table>
    </td></tr>`;
}


function generarSeccionExpedientesIncidenciasEjecutivo(datos) {
  if (!datos.incidencias || datos.incidencias.length === 0) return '';

  const expedientes = datos.incidencias.sort((a, b) => b.totalIncidencias - a.totalIncidencias);
  const mostrar     = expedientes.slice(0, 15);

  let filasHTML = mostrar.map((exp, idx) => {
    const bgColor     = idx % 2 === 0 ? '#ffffff' : '#f8fafc';
    const borderColor = idx === mostrar.length - 1 ? 'transparent' : '#e2e8f0';
    return `
      <tr>
        <td style="padding:14px 20px;border-bottom:1px solid ${borderColor};background:${bgColor};">
          <div style="font-weight:600;color:#1a202c;font-size:13px;font-family:'Inter',sans-serif;">${exp.id}</div>
        </td>
        <td style="padding:14px 20px;border-bottom:1px solid ${borderColor};background:${bgColor};">
          <div style="color:#718096;font-size:13px;font-weight:500;">${exp.fecha}</div>
        </td>
        <td style="padding:14px 20px;border-bottom:1px solid ${borderColor};background:${bgColor};text-align:center;">
          <table role="presentation" cellspacing="0" cellpadding="0" align="center"
                 style="background:#fff;border:2px solid #dc2626;padding:6px 14px;border-radius:6px;">
            <tr><td style="color:#dc2626;font-weight:800;font-size:15px;font-family:'IBM Plex Sans',sans-serif;">${exp.totalIncidencias}</td></tr>
          </table>
        </td>
      </tr>`;
  }).join('');

  return `
    <tr><td style="padding:0 60px 40px;">
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
        <tr><td align="center" style="padding-bottom:32px;">
          <h2 style="margin:0 0 8px;font-size:22px;font-weight:800;color:#1a202c;font-family:'IBM Plex Sans',sans-serif;letter-spacing:-0.5px;">
            Expedientes con Incidencias
          </h2>
          <p style="margin:0;font-size:13px;color:#64748b;font-weight:500;">
            Total: ${expedientes.length} expediente${expedientes.length!==1?'s':''}
            ${mostrar.length < expedientes.length ? ` (mostrando top ${mostrar.length})` : ''}
          </p>
        </td></tr>
      </table>
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0"
             style="background:#fff;border:1px solid #e2e8f0;border-radius:6px;">
        <thead><tr style="background:#f8fafc;">
          <th style="padding:14px 20px;text-align:left;color:#4a5568;font-size:11px;text-transform:uppercase;letter-spacing:1px;font-weight:700;font-family:'Inter',sans-serif;border-bottom:2px solid #cbd5e1;width:45%;">ID Oportunidad</th>
          <th style="padding:14px 20px;text-align:left;color:#4a5568;font-size:11px;text-transform:uppercase;letter-spacing:1px;font-weight:700;font-family:'Inter',sans-serif;border-bottom:2px solid #cbd5e1;width:35%;">Fecha de Firma</th>
          <th style="padding:14px 20px;text-align:center;color:#4a5568;font-size:11px;text-transform:uppercase;letter-spacing:1px;font-weight:700;font-family:'Inter',sans-serif;border-bottom:2px solid #cbd5e1;width:20%;">Incidencias</th>
        </tr></thead>
        <tbody>${filasHTML}</tbody>
      </table>
    </td></tr>`;
}


function generarSeccionDetalleAreasEjecutivo(datos, ICONOS) {
  const areas = {
    'Comercial': { icono:ICONOS.comercial, color:'#f97316', colorClaro:'#FFA459', colorOscuro:'#ea580c', colorBg:'#fff7ed', colorBorder:'#fed7aa' },
    'Cartera':   { icono:ICONOS.cartera,   color:'#ec4899', colorClaro:'#FF8CC8', colorOscuro:'#db2777', colorBg:'#fdf2f8', colorBorder:'#fbcfe8' },
    'Crédito':   { icono:ICONOS.credito,   color:'#06b6d4', colorClaro:'#78EDFF', colorOscuro:'#0891b2', colorBg:'#ecfeff', colorBorder:'#a5f3fc' },
    'Jurídico':  { icono:ICONOS.juridico,  color:'#84cc16', colorClaro:'#C7FF6C', colorOscuro:'#0DA31A', colorBg:'#f7fee7', colorBorder:'#d9f99d' },
    'PLD':       { icono:ICONOS.pld,       color:'#a855f7', colorClaro:'#D3A7FF', colorOscuro:'#9333ea', colorBg:'#faf5ff', colorBorder:'#e9d5ff' }
  };

  let seccionHTML = `
    <tr><td style="padding:0 60px 40px;background:#fff;">
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
        <tr><td align="center" style="padding-bottom:32px;">
          <h2 style="margin:0 0 8px;font-size:22px;font-weight:800;color:#1a202c;font-family:'IBM Plex Sans',sans-serif;letter-spacing:-0.5px;">
            Detalle por Área
          </h2>
          <p style="margin:0;font-size:13px;color:#64748b;font-weight:500;">
            Desglose de incidencias por departamento
          </p>
        </td></tr>
      </table>`;

  Object.keys(areas).forEach(nombreArea => {
    const areaInfo       = areas[nombreArea];
    const expedientesArea = [];
    let totalIncidenciasArea = 0;

    if (datos.incidencias) {
      datos.incidencias.forEach(row => {
        const texto = row.areas || '';
        const regex = new RegExp(`${nombreArea.replace('í','[íi]')}\\s*:\\s*(\\d+)`, 'i');
        const match = texto.match(regex);
        if (match) {
          const cantidad = parseInt(match[1]) || 0;
          if (cantidad > 0) {
            expedientesArea.push({ id: row.id, fecha: row.fecha, cantidad });
            totalIncidenciasArea += cantidad;
          }
        }
      });
    }

    if (expedientesArea.length > 0) {
      expedientesArea.sort((a, b) => b.cantidad - a.cantidad);

      const filasAreaExpedientes = expedientesArea.map((exp, idx) => {
        const bgColor     = idx % 2 === 0 ? '#ffffff' : '#f8fafc';
        const borderColor = idx === expedientesArea.length - 1 ? 'transparent' : '#e2e8f0';
        return `
          <tr>
            <td style="padding:12px 18px;border-bottom:1px solid ${borderColor};background:${bgColor};">
              <div style="font-weight:600;color:#4a5568;font-size:12px;font-family:'Inter',sans-serif;">${exp.id}</div>
            </td>
            <td style="padding:12px 18px;border-bottom:1px solid ${borderColor};background:${bgColor};">
              <div style="color:#718096;font-size:12px;font-weight:500;">${exp.fecha}</div>
            </td>
            <td style="padding:12px 18px;border-bottom:1px solid ${borderColor};background:${bgColor};text-align:center;">
              <table role="presentation" cellspacing="0" cellpadding="0" align="center"
                     style="background:#fff;border:2px solid ${areaInfo.color};padding:6px 14px;border-radius:6px;">
                <tr><td style="color:${areaInfo.color};font-weight:800;font-size:15px;font-family:'IBM Plex Sans',sans-serif;">${exp.cantidad}</td></tr>
              </table>
            </td>
          </tr>`;
      }).join('');

      seccionHTML += `
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0"
               style="margin-bottom:20px;background:#fff;border:1px solid ${areaInfo.colorBorder};border-radius:6px;">
          <tr>
            <td colspan="3" style="background:${areaInfo.colorBg};padding:18px 22px;border-bottom:1px solid ${areaInfo.colorBorder};">
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0"><tr>
                <td width="48" valign="middle" style="padding-right:14px;">
                  <img src="${areaInfo.icono}" alt="${nombreArea}" style="width:40px;height:40px;display:block;"/>
                </td>
                <td valign="middle">
                  <h3 style="margin:0 0 3px;font-size:16px;font-weight:700;color:#1a202c;font-family:'IBM Plex Sans',sans-serif;">
                    ${nombreArea}
                  </h3>
                  <p style="margin:0;font-size:11px;color:#718096;font-weight:600;">
                    ${expedientesArea.length} expediente${expedientesArea.length!==1?'s':''} · ${totalIncidenciasArea} incidencia${totalIncidenciasArea!==1?'s':''}
                  </p>
                </td>
                <td width="70" align="right" valign="middle">
                  <table role="presentation" cellspacing="0" cellpadding="0"
                         style="width:40px;height:40px;background:linear-gradient(135deg,${areaInfo.colorClaro},${areaInfo.color});border-radius:50%;border:5px solid ${areaInfo.colorOscuro};">
                    <tr><td align="center" valign="middle"
                            style="color:#fff;font-size:18px;font-weight:900;font-family:'IBM Plex Sans',sans-serif;text-shadow:0 1px 2px rgba(0,0,0,0.1);">
                      ${totalIncidenciasArea}
                    </td></tr>
                  </table>
                </td>
              </tr></table>
            </td>
          </tr>
          <tr style="background:#f8fafc;">
            <th style="padding:12px 18px;text-align:left;color:#4a5568;font-size:10px;text-transform:uppercase;letter-spacing:0.8px;font-weight:700;font-family:'Inter',sans-serif;border-bottom:1px solid #e2e8f0;width:45%;">ID Oportunidad</th>
            <th style="padding:12px 18px;text-align:left;color:#4a5568;font-size:10px;text-transform:uppercase;letter-spacing:0.8px;font-weight:700;font-family:'Inter',sans-serif;border-bottom:1px solid #e2e8f0;width:40%;">Fecha de Firma</th>
            <th style="padding:12px 18px;text-align:center;color:#4a5568;font-size:10px;text-transform:uppercase;letter-spacing:0.8px;font-weight:700;font-family:'Inter',sans-serif;border-bottom:1px solid #e2e8f0;width:15%;">Cantidad</th>
          </tr>
          ${filasAreaExpedientes}
        </table>`;

    } else {
      seccionHTML += `
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0"
               style="margin-bottom:16px;background:#f0fdf4;border-radius:6px;border:1px solid #86efac;">
          <tr><td style="padding:18px 22px;">
            <table role="presentation" width="100%" cellspacing="0" cellpadding="0"><tr>
              <td width="48" valign="middle" style="padding-right:14px;">
                <img src="${areaInfo.icono}" alt="${nombreArea}" style="width:40px;height:40px;display:block;"/>
              </td>
              <td valign="middle">
                <h3 style="margin:0 0 3px;font-size:16px;font-weight:700;color:#1a202c;font-family:'IBM Plex Sans',sans-serif;">${nombreArea}</h3>
                <p style="margin:0;font-size:13px;color:#047857;font-weight:600;">✓ Sin incidencias registradas</p>
              </td>
            </tr></table>
          </td></tr>
        </table>`;
    }
  });

  seccionHTML += `</td></tr>`;
  return seccionHTML;
}


// =====================================================
// FUNCIONES DE EMAIL (se mantienen intactas)
// =====================================================

function enviarEmailTESIInfografia(periodo, destinatarios) {
  try {
    const datos    = obtenerTESIPorPeriodo(periodo);
    const htmlEmail = generarHTMLEmailInfografia(datos, periodo);
    const [anio, mes] = periodo.split("-");
    const nombreMes   = obtenerNombreMes(parseInt(mes));
    const asunto = `📊 EXPEDIENTES CUALLI/Cierre ${nombreMes} ${anio} / Estatus de integración Expedientes Crédito`;
    MailApp.sendEmail({ to: destinatarios, subject: asunto, htmlBody: htmlEmail, name: "Control de Calidad Cualli" });
    return { success: true, mensaje: `Email ejecutivo enviado a: ${destinatarios}`, tesi: datos.tesi };
  } catch (error) {
    console.error("Error al enviar email ejecutivo:", error);
    return { success: false, error: error.toString() };
  }
}

function testEmailEjecutivo() {
  const resultado = enviarEmailTESIInfografia("2025-01", "tu.email@example.com");
  console.log(resultado);
}

function abrirEmailTESI() {
  const html = HtmlService.createHtmlOutputFromFile('PreviewEmail.html').setTitle('📧 Email TESI');
  SpreadsheetApp.getUi().showSidebar(html);
}

function abrirPreviewModal(periodo) {
  try {
    const datos     = obtenerTESIPorPeriodo(periodo);
    const htmlEmail = generarHTMLEmailInfografia(datos, periodo);
    const output    = HtmlService.createHtmlOutput(htmlEmail).setWidth(1200).setHeight(800);
    const [anio, mes] = periodo.split("-");
    const nombreMes   = obtenerNombreMes(parseInt(mes));
    SpreadsheetApp.getUi().showModalDialog(output, `📧 Vista Previa - Email TESI ${nombreMes} ${anio}`);
    return { success: true };
  } catch (error) {
    Logger.log("Error en abrirPreviewModal: " + error);
    return { success: false, error: error.toString() };
  }
}




/**
 * Utilidad: Obtener nombre del mes (si no la tienes ya)
 */
function obtenerNombreMes(numeroMes) {
  const meses = [
    '', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];
  return meses[numeroMes] || '';
}


/**
 * Limpia una línea de documento de la col K
 * Entrada:  "• Comercial | Checklist | 01. Dictamen de la empresa /Nombre de la Operación | Falta"
 * Salida:   { area, documento, comentario }
 *           documento  → "Dictamen de la empresa"
 *           comentario → "Falta"
 */
function limpiarDocumento(linea) {
  const limpia = linea.replace(/^[•\-·]\s*/, '').trim();
  const partes = limpia.split(' | ').map(p => p.trim());

  const area       = partes[0] || '';
  const origen     = partes[1] || '';
  const docRaw     = partes[2] || '';
  const comentario = partes[3] || '';

  const documento = docRaw
    .replace(/^\*+\s*/,  '')      
    .replace(/^\d+\.\s*/, '')      
    .replace(/^\*+\s*/,  '')       
    .replace(/\s*\/.*$/, '')     
    .trim();

  return {
    area,
    origen,
    documento,
    comentario,
    textoLimpio: documento + (comentario && comentario !== '—' ? ` — ${comentario}` : '')
  };
}
/**
 * Parsea toda la columna K de un expediente
 * Devuelve array de { area, documento, comentario, textoLimpio }
 */
function parsearDetalleIncidencias(texto) {
  if (!texto) return [];
  const raw = String(texto).trim();
  if (!raw || raw === '—') return [];

  return raw.split('\n')
    .map(l => l.trim())
    .filter(l => /^[•\-·]/.test(l))
    .map(linea => limpiarDocumento(linea))
    .filter(d => d.documento); // descartar líneas vacías
}

// =====================================================
// CALENDARIO PROCESOS AUTOMATIZADOS
// Agregar este bloque al final de tesi-functions.gs
// =====================================================

/**
 * Configuración del archivo de Proyectos
 */
const PROYECTOS_CONFIG = {
  SPREADSHEET_ID: '1Cl7t-x6TSQN8Zf3UeDq0tBV2yh0etDxrPAD4GSBiXZ8',
  HOJA_NOMBRE: 'proyectos'
};

/**
 * Obtiene todos los datos de proyectos y calcula KPIs agregados.
 * Se llama desde el frontend via google.script.run.getDatosProyectos()
 * 
 * @returns {Object} Objeto con proyectos individuales y KPIs calculados
 */
function getDatosProyectos() {
  try {
    const ss = SpreadsheetApp.openById(PROYECTOS_CONFIG.SPREADSHEET_ID);
    const hoja = ss.getSheetByName(PROYECTOS_CONFIG.HOJA_NOMBRE);
    
    if (!hoja) {
      throw new Error('No se encontró la hoja "' + PROYECTOS_CONFIG.HOJA_NOMBRE + '"');
    }
    
    const datos = hoja.getDataRange().getValues();
    const encabezados = datos[0];
    const filas = datos.slice(1).filter(function(fila) {
      return fila[1]; // Columna B = Proyecto (debe existir)
    });
    
    // ══════════════════════════════════════════════════════
    // ÍNDICES DE COLUMNAS ACTUALIZADOS (0-indexed)
    // A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, L=11, M=12
    // ══════════════════════════════════════════════════════
    const COL = {
      NO: 0,                    // A - No.
      PROYECTO: 1,              // B - Proyecto
      DESCRIPCION: 2,           // C - Descripción
      AREA_SOLICITANTE: 3,      // D - Área Solicitante
      TIPO: 4,                  // E - Tipo (Tren de Crédito / Cumplimiento Normativo)
      RESPONSABLE_CONTRALORIA: 5, // F - Responsable Contraloría
      SOLICITANTE: 6,           // G - Solicitante
      DURACION_DIAS: 7,         // H - Duración (días)
      FECHA_INICIO: 8,          // I - Fecha Inicio
      FECHA_COMPROMISO: 9,      // J - Fecha Compromiso
      ESTADO: 10,               // K - Estado
      NIVEL_RIESGO: 11,         // L - Nivel de Riesgo
      PORCENTAJE_AVANCE: 12     // M - % Avance
    };
    
    // ── Parsear cada proyecto ──
    const proyectos = filas.map(function(fila, idx) {
      return {
        id: fila[COL.NO] || idx + 1,
        proyecto: fila[COL.PROYECTO] || '',
        descripcion: fila[COL.DESCRIPCION] || '',
        areaSolicitante: normalizarAreaProyecto(fila[COL.AREA_SOLICITANTE]),
        tipo: normalizarTipoProyecto(fila[COL.TIPO]),
        responsableContraloria: fila[COL.RESPONSABLE_CONTRALORIA] || '',
        solicitante: fila[COL.SOLICITANTE] || '',
        duracionDias: parsearNumeroProyecto(fila[COL.DURACION_DIAS]),
        fechaInicio: formatearFechaProyecto(fila[COL.FECHA_INICIO]),
        fechaCompromiso: formatearFechaProyecto(fila[COL.FECHA_COMPROMISO]),
        estado: normalizarEstadoProyecto(fila[COL.ESTADO]),
        nivelRiesgo: normalizarRiesgoProyecto(fila[COL.NIVEL_RIESGO]),
        avance: parsearPorcentajeProyecto(fila[COL.PORCENTAJE_AVANCE])
      };
    });
    
    // ── Calcular KPIs ──
    var kpis = calcularKPIsProyectos(proyectos);
    
    return {
      proyectos: proyectos,
      kpis: kpis,
      actualizadoEn: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('Error en getDatosProyectos:', error);
    throw new Error('No se pudieron cargar los proyectos: ' + error.message);
  }
}

/**
 * Calcula todos los KPIs agregados de los proyectos
 */
function calcularKPIsProyectos(proyectos) {
  var total = proyectos.length;
  
  // ── Conteo por Estado ──
  var porEstado = {};
  proyectos.forEach(function(p) {
    var estado = p.estado || 'Sin estado';
    porEstado[estado] = (porEstado[estado] || 0) + 1;
  });
  
  // ── Conteo por Área Solicitante ──
  var porArea = {};
  proyectos.forEach(function(p) {
    var area = p.areaSolicitante || 'Sin área';
    if (!porArea[area]) {
      porArea[area] = { total: 0, enCurso: 0, riesgoAlto: 0, avancePromedio: 0, sumAvance: 0 };
    }
    porArea[area].total++;
    porArea[area].sumAvance += p.avance || 0;
    if (p.nivelRiesgo === 'Alto') porArea[area].riesgoAlto++;
    
    // Contar en curso
    var estadosEnCurso = ['Solicitud inicial', 'Levantamiento de información', 'En desarrollo', 'En pruebas'];
    if (estadosEnCurso.some(function(e) { return p.estado && p.estado.toLowerCase().includes(e.toLowerCase()); })) {
      porArea[area].enCurso++;
    }
  });
  
  // Calcular promedio de avance por área
  Object.keys(porArea).forEach(function(area) {
    porArea[area].avancePromedio = porArea[area].total > 0 
      ? Math.round(porArea[area].sumAvance / porArea[area].total) 
      : 0;
    delete porArea[area].sumAvance;
  });
  
  // ── Conteo por Tipo (NUEVO) ──
  var porTipo = {
    'Tren de Crédito': 0,
    'Cumplimiento Normativo': 0,
    'Otro': 0
  };
  proyectos.forEach(function(p) {
    var tipo = p.tipo || 'Otro';
    if (tipo.toLowerCase().includes('crédito') || tipo.toLowerCase().includes('credito')) {
      porTipo['Tren de Crédito']++;
    } else if (tipo.toLowerCase().includes('cumplimiento') || tipo.toLowerCase().includes('normativo')) {
      porTipo['Cumplimiento Normativo']++;
    } else {
      porTipo['Otro']++;
    }
  });
  
  // ── Conteo por Nivel de Riesgo ──
  var porRiesgo = { 'Alto': 0, 'Medio': 0, 'Bajo': 0 };
  proyectos.forEach(function(p) {
    var riesgo = p.nivelRiesgo || 'Sin definir';
    if (porRiesgo.hasOwnProperty(riesgo)) {
      porRiesgo[riesgo]++;
    }
  });
  
  // ── Proyectos en curso vs completados ──
  var estadosEnCurso = ['Solicitud inicial', 'Levantamiento de información', 'En pruebas', 'En desarrollo'];
  var estadosCompletados = ['Completado', 'En producción', 'Finalizado'];
  
  var enCurso = 0;
  var completados = 0;
  var pausados = 0;
  var cancelados = 0;
  
  proyectos.forEach(function(p) {
    var estado = (p.estado || '').toLowerCase();
    if (estadosEnCurso.some(function(e) { return estado.includes(e.toLowerCase()); })) {
      enCurso++;
    } else if (estadosCompletados.some(function(e) { return estado.includes(e.toLowerCase()); })) {
      completados++;
    } else if (estado.includes('pausado')) {
      pausados++;
    } else if (estado.includes('cancelado')) {
      cancelados++;
    }
  });
  
  // ── Avance promedio general ──
  var sumAvance = 0;
  var countAvance = 0;
  proyectos.forEach(function(p) {
    if (p.avance !== null && p.avance !== undefined) {
      sumAvance += p.avance;
      countAvance++;
    }
  });
  var avancePromedio = countAvance > 0 ? Math.round(sumAvance / countAvance) : 0;
  
  // ── Duración promedio ──
  var sumDuracion = 0;
  var countDuracion = 0;
  proyectos.forEach(function(p) {
    if (p.duracionDias && p.duracionDias > 0) {
      sumDuracion += p.duracionDias;
      countDuracion++;
    }
  });
  var duracionPromedio = countDuracion > 0 ? Math.round(sumDuracion / countDuracion) : 0;
  
  // ── Score de riesgo (0-100, mayor = más riesgo) ──
  var riesgoScore = total > 0 
    ? Math.round(((porRiesgo['Alto'] * 100) + (porRiesgo['Medio'] * 50) + (porRiesgo['Bajo'] * 10)) / total)
    : 0;
  
  return {
    total: total,
    enCurso: enCurso,
    completados: completados,
    pausados: pausados,
    cancelados: cancelados,
    avancePromedio: avancePromedio,
    duracionPromedio: duracionPromedio,
    porEstado: porEstado,
    porArea: porArea,
    porTipo: porTipo,
    porRiesgo: porRiesgo,
    riesgoScore: riesgoScore,
    riesgoAlto: porRiesgo['Alto'],
    riesgoMedio: porRiesgo['Medio'],
    riesgoBajo: porRiesgo['Bajo']
  };
}


// =====================================================
// HELPERS PARA PROYECTOS
// (Con sufijo "Proyecto" para evitar conflictos)
// =====================================================

/**
 * Normaliza el nombre del área solicitante
 */
function normalizarTipoProyecto(tipo) {
  if (!tipo) return '';
  var t = String(tipo).trim();
  
  // Normalizar variantes comunes
  if (t.toLowerCase().includes('crédito') || t.toLowerCase().includes('credito') || t.toLowerCase().includes('tren')) {
    return 'Tren de Crédito';
  }
  if (t.toLowerCase().includes('cumplimiento') || t.toLowerCase().includes('normativo') || t.toLowerCase().includes('regulatorio')) {
    return 'Cumplimiento Normativo';
  }
  
  return t; // Devolver el valor original si no coincide
}
 
/**
 * Normaliza el área solicitante del proyecto
 */
function normalizarAreaProyecto(area) {
  if (!area) return 'Sin área';
  var a = String(area).trim();
  
  // Lista de áreas válidas - ajusta según tus áreas reales
  var areasValidas = [
    'Comercial', 'Cartera', 'Crédito', 'Jurídico', 'PLD',
    'Operaciones', 'Sistemas', 'Contraloría', 'Dirección General',
    'Recursos Humanos', 'Administración', 'Cobranza', 'Mesa de Control'
  ];
  
  // Buscar coincidencia
  for (var i = 0; i < areasValidas.length; i++) {
    if (a.toLowerCase().includes(areasValidas[i].toLowerCase())) {
      return areasValidas[i];
    }
  }
  
  return a; // Devolver el valor original si no hay coincidencia
}
 
/**
 * Normaliza el estado del proyecto
 */
function normalizarEstadoProyecto(estado) {
  if (!estado) return 'Sin estado';
  var e = String(estado).trim().toLowerCase();
  
  if (e.includes('solicitud')) return 'Solicitud inicial';
  if (e.includes('levantamiento')) return 'Levantamiento de información';
  if (e.includes('desarrollo')) return 'En desarrollo';
  if (e.includes('prueba')) return 'En pruebas';
  if (e.includes('producción') || e.includes('produccion')) return 'En producción';
  if (e.includes('completado') || e.includes('finalizado')) return 'Completado';
  if (e.includes('pausado')) return 'Pausado';
  if (e.includes('cancelado')) return 'Cancelado';
  
  return String(estado).trim();
}
 
/**
 * Normaliza el nivel de riesgo
 */
function normalizarRiesgoProyecto(riesgo) {
  if (!riesgo) return '';
  var r = String(riesgo).trim().toLowerCase();
  
  if (r.includes('alto') || r === 'high' || r === '3') return 'Alto';
  if (r.includes('medio') || r === 'medium' || r === '2') return 'Medio';
  if (r.includes('bajo') || r === 'low' || r === '1') return 'Bajo';
  
  return String(riesgo).trim();
}
 
/**
 * Parsea un número de la hoja
 */
function parsearNumeroProyecto(valor) {
  if (!valor && valor !== 0) return 0;
  var num = parseInt(String(valor).replace(/[^\d.-]/g, ''), 10);
  return isNaN(num) ? 0 : num;
}
 
/**
 * Parsea un porcentaje de la hoja
 */
function parsearPorcentajeProyecto(valor) {
  if (!valor && valor !== 0) return 0;
  
  // Si es número decimal (0.85 = 85%)
  if (typeof valor === 'number' && valor <= 1) {
    return Math.round(valor * 100);
  }
  
  // Si es string con %
  var str = String(valor).replace('%', '').trim();
  var num = parseFloat(str);
  
  if (isNaN(num)) return 0;
  
  // Si parece ser decimal
  if (num <= 1 && num > 0) {
    return Math.round(num * 100);
  }
  
  return Math.round(num);
}
 
/**
 * Formatea una fecha para mostrar
 */
function formatearFechaProyecto(fecha) {
  if (!fecha) return '';
  
  try {
    var d = new Date(fecha);
    if (isNaN(d.getTime())) return String(fecha);
    
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch (e) {
    return String(fecha);
  }
}

/**
 * Función de prueba para verificar que todo funciona
 */
function testGetDatosProyectos() {
  var datos = getDatosProyectos();
  console.log('=== TEST PROYECTOS ===');
  console.log('Total proyectos:', datos.kpis.total);
  console.log('En curso:', datos.kpis.enCurso);
  console.log('Completados:', datos.kpis.completados);
  console.log('Riesgo Alto:', datos.kpis.riesgoAlto);
  console.log('Vencidos:', datos.kpis.vencidos);
  console.log('Próximos a vencer:', datos.kpis.proximosVencer);
  console.log('Por área:', JSON.stringify(datos.kpis.porArea, null, 2));
  console.log('Por estado:', JSON.stringify(datos.kpis.porEstado, null, 2));
  if (datos.proyectos.length > 0) {
    console.log('Primer proyecto:', JSON.stringify(datos.proyectos[0], null, 2));
  }
}
