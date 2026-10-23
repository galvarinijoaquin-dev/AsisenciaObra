// ===================== CONFIGURACION =====================
// Pegar aca el ID de la planilla (lo que esta entre /d/ y /edit en su direccion)
const ID_PLANILLA = 'PEGAR_ID_DE_LA_PLANILLA';

const ZONA = 'America/Argentina/San_Juan';
// ID de la carpeta de Drive donde van los Excel (lo que esta despues de /folders/ en su direccion).
// Si se deja vacio, busca la carpeta por nombre dentro de "Mi unidad".
const ID_CARPETA = 'PEGAR_ID_DE_LA_CARPETA';
const CARPETA_NOMBRE = 'Asistencia Obra';
const PREFIJO_EXCEL = 'Asistencia_Obra_';

const HOJA_NOMBRES = 'Nombres';
const HOJA_ASISTENCIA = 'Asistencia';

// Columnas de la hoja Asistencia (empezando en 0)
const COL = { ID: 0, NOMBRE: 1, ENTRADA: 2, SALIDA: 3, EXPORTADO: 4 };

// ===================== API =====================
// La pagina (index.html) esta en GitHub Pages y le pide los datos a esta URL /exec.
// Asi nadie abre Apps Script directo y no hay problemas de cuentas de Google.
function doGet(e) {
  return json_(estado());
}

function doPost(e) {
  let datos = {};
  try { datos = JSON.parse(e.postData.contents); } catch (err) {}
  const res = marcar(datos.nombre, datos.tipo);
  // Devuelve tambien los contadores actualizados para que la pagina no tenga que pedirlos aparte
  res.estado = estado();
  return json_(res);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Ejecutar UNA sola vez desde el editor para crear las hojas y la actualizacion automatica
function setup() {
  const ss = libro_();
  ss.setSpreadsheetTimeZone(ZONA);
  crearHoja_(ss, HOJA_NOMBRES, ['Nombre'], [['Empleado de Ejemplo']]);
  crearHoja_(ss, HOJA_ASISTENCIA, ['ID', 'Nombre', 'Entrada', 'Salida', 'Exportado'], []);
  ss.getSheetByName(HOJA_ASISTENCIA).getRange('C:D').setNumberFormat('yyyy-mm-dd hh:mm:ss');

  // Los cuadros y el Excel del mes se actualizan solos cada 10 minutos (solo si hubo marcados nuevos)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'generarExcel')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('generarExcel').timeBased().everyMinutes(10).create();
}

function crearHoja_(ss, nombre, cabecera, ejemplo) {
  if (ss.getSheetByName(nombre)) return;
  const hoja = ss.insertSheet(nombre);
  hoja.appendRow(cabecera);
  hoja.getRange(1, 1, 1, cabecera.length).setFontWeight('bold');
  hoja.setFrozenRows(1);
  if (ejemplo.length) hoja.getRange(2, 1, ejemplo.length, ejemplo[0].length).setValues(ejemplo);
}

// ===================== AYUDANTES =====================
function libro_() {
  return SpreadsheetApp.openById(ID_PLANILLA);
}

function hoja_(nombre) {
  return libro_().getSheetByName(nombre);
}

function lista_(nombreHoja) {
  const hoja = hoja_(nombreHoja);
  if (hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(String);
}

function empleados_() {
  return lista_(HOJA_NOMBRES).sort((a, b) => a.localeCompare(b));
}

function registros_() {
  const hoja = hoja_(HOJA_ASISTENCIA);
  if (hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, 5).getValues();
}

// Ultimo registro de cada empleado: { nombre: { fila, r } }
function ultimos_(regs) {
  const ult = {};
  regs.forEach((r, i) => { ult[r[COL.NOMBRE]] = { fila: i + 2, r: r }; });
  return ult;
}

function dia_(fecha) { return Utilities.formatDate(fecha, ZONA, 'yyyy-MM-dd'); }
function mes_(fecha) { return Utilities.formatDate(fecha, ZONA, 'yyyy-MM'); }
function hora_(fecha) { return Utilities.formatDate(fecha, ZONA, 'HH:mm'); }
function fechaHora_(fecha) { return Utilities.formatDate(fecha, ZONA, 'yyyy-MM-dd HH:mm:ss'); }

function abierta_(r) { return !(r[COL.SALIDA] instanceof Date); }

function duracion_(segs) {
  if (!(segs > 0)) return '0h 0m';
  return Math.floor(segs / 3600) + 'h ' + Math.floor((segs % 3600) / 60) + 'm';
}

// ===================== PANTALLA =====================
function estado() {
  const nombres = empleados_();
  const ult = ultimos_(registros_());
  const hoy = dia_(new Date());

  const lista = [];
  const enSala = [];
  const finalizados = [];

  nombres.forEach(nombre => {
    const u = ult[nombre];
    let adentro = false;
    if (u) {
      const r = u.r;
      if (abierta_(r)) {
        adentro = true;
        enSala.push({ nombre: nombre, entrada: hora_(r[COL.ENTRADA]) });
      } else if (dia_(r[COL.SALIDA]) === hoy) {
        finalizados.push({ nombre: nombre, entrada: hora_(r[COL.ENTRADA]), salida: hora_(r[COL.SALIDA]) });
      }
    }
    lista.push({ nombre: nombre, adentro: adentro });
  });

  return {
    lista: lista,
    en_sala: enSala.length,
    ya_salieron: finalizados.length,
    total: nombres.length,
    en_sala_detalle: enSala,
    finalizados_detalle: finalizados
  };
}

function marcar(nombre, tipo) {
  if (!nombre || nombre === 'none' || empleados_().indexOf(nombre) === -1) {
    return { status: 'error', msg: 'Debe seleccionar un nombre del listado' };
  }

  // Evita que dos marcados al mismo tiempo se pisen
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { status: 'error', msg: 'Sistema ocupado, intenta de nuevo' };

  try {
    const hoja = hoja_(HOJA_ASISTENCIA);
    const u = ultimos_(registros_())[nombre];
    const abierta = u && abierta_(u.r);

    // Entrada sin salida de un dia anterior: no se cierra sola, la corrige el tecnico
    if (abierta && dia_(u.r[COL.ENTRADA]) !== dia_(new Date())) {
      return { status: 'error', msg: 'AVISA AL TECNICO' };
    }

    if (tipo === 'entrada') {
      if (abierta) return { status: 'error', msg: 'ERROR YA MARCASTE ENTRADA' };
      hoja.appendRow([hoja.getLastRow(), nombre, new Date(), '', 0]);
      return { status: 'success', msg: 'Marcado de entrada correcto' };
    }

    if (tipo === 'salida') {
      if (!abierta) return { status: 'error', msg: 'ERROR YA MARCASTE SALIDA' };
      hoja.getRange(u.fila, COL.SALIDA + 1, 1, 2).setValues([[new Date(), 0]]);
      return { status: 'success', msg: 'Marcado de salida correcto' };
    }

    return { status: 'error', msg: 'Tipo de marcado invalido' };
  } finally {
    lock.releaseLock();
  }
}

// ===================== EXCEL =====================
function carpeta_() {
  if (ID_CARPETA && ID_CARPETA.indexOf('PEGAR') !== 0) return DriveApp.getFolderById(ID_CARPETA);
  const it = DriveApp.getRootFolder().getFoldersByName(CARPETA_NOMBRE);
  return it.hasNext() ? it.next() : null;
}

// La ejecuta el reloj automatico cada 10 minutos. Si algo falla queda en rojo en "Ejecuciones".
function generarExcel() {
  const msg = generarExcel_();
  console.log(msg);
  if (msg.indexOf('ATENCION') === 0) throw new Error(msg);
  return msg;
}

function generarExcel_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return 'ATENCION: Sistema ocupado, intenta de nuevo en unos segundos.';

  try {
    const regs = registros_();
    const pendientes = regs.filter(r => Number(r[COL.EXPORTADO]) !== 1);
    const meses = [...new Set(pendientes.map(r => mes_(r[COL.ENTRADA])))].sort();

    if (!meses.length) return 'No hay asistencias nuevas para exportar. Ya está todo al día.';

    const nombres = empleados_();

    // Cuadros del mes actual dentro de esta misma planilla
    const mesActual = mes_(new Date());
    pintarMes_(hojaLimpia_('Registro Diario'), hojaLimpia_('Total Mensual'), mesActual,
      regs.filter(r => mes_(r[COL.ENTRADA]) === mesActual), nombres);

    const carpeta = carpeta_();
    if (!carpeta) return 'ATENCION: No se encontro la carpeta "' + CARPETA_NOMBRE + '" en Mi unidad.';

    const archivos = meses.map(mes =>
      crearExcel_(mes, regs.filter(r => mes_(r[COL.ENTRADA]) === mes), nombres, carpeta));

    hoja_(HOJA_ASISTENCIA).getRange(2, COL.EXPORTADO + 1, regs.length, 1).setValues(regs.map(() => [1]));
    return '¡Éxito! Planilla exacta y completa generada: ' + archivos.join(', ');
  } catch (e) {
    return 'ATENCION: Error al generar el Excel: ' + e.message;
  } finally {
    lock.releaseLock();
  }
}

// Dibuja los cuadros del mes: uno por dia en ws1 y el total por empleado en ws2
function pintarMes_(ws1, ws2, mes, filas, nombres) {
  filas = filas.slice().sort((a, b) => a[COL.ENTRADA] - b[COL.ENTRADA]);

  // --- Hoja 1: Registro Diario ---
  [220, 150, 150, 140].forEach((ancho, i) => ws1.setColumnWidth(i + 1, ancho));

  const totales = {};
  const dias = [...new Set(filas.map(r => dia_(r[COL.ENTRADA])))];
  let fila = 1;

  dias.forEach(dia => {
    ws1.getRange(fila, 1, 1, 4).merge().setValue('▶ FECHA: ' + dia)
      .setFontWeight('bold').setBackground('#D9E1F2').setFontSize(11).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, null, null);
    fila++;

    ws1.getRange(fila, 1, 1, 4)
      .setValues([['Empleado', 'Hora Entrada', 'Hora Salida', 'Tiempo Trabajado']])
      .setFontWeight('bold').setBackground('#4B5563').setFontColor('white')
      .setBorder(true, true, true, true, true, true);
    fila++;

    const datos = filas.filter(r => dia_(r[COL.ENTRADA]) === dia).map(r => {
      const ent = r[COL.ENTRADA];
      const sal = r[COL.SALIDA];
      const segs = sal instanceof Date ? (sal - ent) / 1000 : 0;
      totales[r[COL.NOMBRE]] = (totales[r[COL.NOMBRE]] || 0) + segs;
      return [r[COL.NOMBRE], fechaHora_(ent),
              sal instanceof Date ? fechaHora_(sal) : 'Sin salida', duracion_(segs)];
    });

    // Formato texto para que Sheets no convierta las horas en fechas
    ws1.getRange(fila, 1, datos.length, 4).setNumberFormat('@').setValues(datos)
      .setBorder(true, true, true, true, true, true);
    fila += datos.length + 1;
  });

  // --- Hoja 2: Total Mensual ---
  ws2.setColumnWidth(1, 40);
  ws2.setColumnWidth(2, 280);
  ws2.setColumnWidth(3, 160);

  ws2.getRange(3, 2, 1, 2).merge().setValue('RESUMEN DE HORAS - MES: ' + mes)
    .setFontWeight('bold').setBackground('#DF3A2C').setFontColor('white').setFontSize(12)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null);

  ws2.getRange(4, 2, 1, 2).setValues([['Empleado', 'Total Exacto']])
    .setFontWeight('bold').setBackground('#FCE4D6').setFontColor('#3E2723')
    .setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);

  if (nombres.length) {
    const resumen = nombres.map(n => [n, duracion_(totales[n] || 0)]);
    ws2.getRange(5, 2, resumen.length, 2).setNumberFormat('@').setValues(resumen)
      .setBorder(true, true, true, true, true, true);
    ws2.getRange(5, 3, resumen.length, 1)
      .setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#367C2B');
  }
}

// Devuelve la hoja (creandola si no existe) completamente vacia
function hojaLimpia_(nombre) {
  const ss = libro_();
  const hoja = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  hoja.getRange(1, 1, hoja.getMaxRows(), hoja.getMaxColumns()).breakApart();
  hoja.clear();
  return hoja;
}

function crearExcel_(mes, filas, nombres, carpeta) {
  // Se arma en una planilla temporal de Google y se exporta como .xlsx
  const tmp = SpreadsheetApp.create('tmp_' + PREFIJO_EXCEL + mes);
  tmp.setSpreadsheetTimeZone(ZONA);

  try {
    pintarMes_(tmp.getSheets()[0].setName('Registro Diario'), tmp.insertSheet('Total Mensual'), mes, filas, nombres);
    SpreadsheetApp.flush();

    // --- Exportar a .xlsx y reemplazar el anterior en la carpeta ---
    const nombreArchivo = PREFIJO_EXCEL + mes + '.xlsx';
    const url = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
    const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
      .getBlob().setName(nombreArchivo);

    const viejos = carpeta.getFilesByName(nombreArchivo);
    while (viejos.hasNext()) viejos.next().setTrashed(true);
    carpeta.createFile(blob);

    return nombreArchivo;
  } finally {
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
  }
}
