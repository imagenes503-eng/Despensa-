// =====================================================================
//  DESPENSA ECONÓMICA — Google Apps Script  (versión SYNC BIDIRECCIONAL)
//  Pega TODO este código en script.google.com → Proyecto nuevo
//  Luego: Implementar → Nueva implementación → App web
//         Ejecutar como: Yo
//         Acceso: Cualquiera
//  Copia la URL /exec y pégala en la app (botón ⚙️ Sheets)
// =====================================================================

// OPCIONAL: pega aquí el ID de tu Google Sheet específico
// Si lo dejas vacío, el script usa la hoja activa o crea una nueva
const SPREADSHEET_ID = '';

const HOJA_VENTAS         = 'Ventas';
const HOJA_PRODUCTOS      = 'Inventario';
const HOJA_VENTAS_DIARIAS = 'VentasDiarias';

const CABECERA_VENTAS         = ['ID', 'Fecha', 'Total $', 'Pago $', 'Vuelto $', 'Productos'];
const CABECERA_PRODUCTOS      = ['ID', 'Nombre', 'Categoría', 'Precio Compra', 'Precio Venta', 'Stock', 'Stock Mínimo'];
const CABECERA_VENTAS_DIARIAS = ['Fecha', 'Monto', 'Nota'];

// ─────────────────────────────────────────────────────────────────────
//  POST: escribe datos desde la app
// ─────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const accion = body.accion;
    if (accion === 'VENTA')          return responder(registrarVenta(body));
    if (accion === 'PRODUCTOS')      return responder(sincronizarProductos(body));
    if (accion === 'VENTAS_DIARIAS') return responder(sincronizarVentasDiarias(body));
    return responder({ ok: false, error: 'Accion desconocida: ' + accion });
  } catch (err) {
    return responder({ ok: false, error: err.message });
  }
}

// ─────────────────────────────────────────────────────────────────────
//  GET: devuelve todos los datos para importar en otro teléfono
//  Uso: fetch(URL + "?accion=EXPORTAR") → JSON completo
// ─────────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const accion = (e.parameter && e.parameter.accion) ? e.parameter.accion : 'PING';

    if (accion === 'PING') {
      return responder({ ok: true, msg: 'Despensa Economica API activa' });
    }

    if (accion === 'EXPORTAR') {
      return responder(exportarTodo());
    }

    return responder({ ok: false, error: 'Accion GET desconocida: ' + accion });
  } catch (err) {
    return responder({ ok: false, error: err.message });
  }
}

// ─────────────────────────────────────────────────────────────────────
//  EXPORTAR TODO: lee las tres hojas y devuelve JSON
// ─────────────────────────────────────────────────────────────────────
function exportarTodo() {
  const ss = obtenerSS();

  const productos      = leerHoja(ss, HOJA_PRODUCTOS,      CABECERA_PRODUCTOS);
  const ventas         = leerHoja(ss, HOJA_VENTAS,         CABECERA_VENTAS);
  const ventasDiarias  = leerHoja(ss, HOJA_VENTAS_DIARIAS, CABECERA_VENTAS_DIARIAS);

  return {
    ok: true,
    ts: new Date().toISOString(),
    productos:     productos,
    ventas:        ventas,
    ventasDiarias: ventasDiarias
  };
}

// Lee una hoja y devuelve array de objetos {columna: valor}
function leerHoja(ss, nombre, cabeceras) {
  const hoja = ss.getSheetByName(nombre);
  if (!hoja) return [];
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  const filas = hoja.getRange(2, 1, lastRow - 1, cabeceras.length).getValues();
  return filas
    .filter(f => f.some(c => c !== '' && c !== null))
    .map(f => {
      const obj = {};
      cabeceras.forEach((cab, i) => { obj[cab] = f[i]; });
      return obj;
    });
}

// ─────────────────────────────────────────────────────────────────────
//  REGISTRAR VENTA individual
// ─────────────────────────────────────────────────────────────────────
function registrarVenta(data) {
  const hoja = obtenerHoja(HOJA_VENTAS, CABECERA_VENTAS);
  hoja.appendRow([
    data.id     || '',
    data.fecha  || new Date().toISOString(),
    data.total  || 0,
    data.pago   || 0,
    data.vuelto || 0,
    data.items  || ''
  ]);
  return { ok: true, msg: 'Venta registrada' };
}

// ─────────────────────────────────────────────────────────────────────
//  SINCRONIZAR PRODUCTOS (reemplaza todo el inventario)
// ─────────────────────────────────────────────────────────────────────
function sincronizarProductos(data) {
  const hoja = obtenerHoja(HOJA_PRODUCTOS, CABECERA_PRODUCTOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila > 1) {
    hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).clearContent();
  }
  const filas = data.filas || [];
  if (filas.length > 0) {
    hoja.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
  }
  return { ok: true, msg: filas.length + ' productos sincronizados' };
}

// ─────────────────────────────────────────────────────────────────────
//  SINCRONIZAR VENTAS DIARIAS (reemplaza todo)
// ─────────────────────────────────────────────────────────────────────
function sincronizarVentasDiarias(data) {
  const hoja = obtenerHoja(HOJA_VENTAS_DIARIAS, CABECERA_VENTAS_DIARIAS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila > 1) {
    hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).clearContent();
  }
  const filas = data.filas || [];
  if (filas.length > 0) {
    hoja.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
  }
  return { ok: true, msg: filas.length + ' ventas diarias sincronizadas' };
}

// ─────────────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────────────
function obtenerSS() {
  try {
    return SPREADSHEET_ID
      ? SpreadsheetApp.openById(SPREADSHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
  } catch(e) {
    return SpreadsheetApp.create('Despensa Economica — Datos');
  }
}

function obtenerHoja(nombre, cabeceras) {
  const ss = obtenerSS();
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    const r = hoja.getRange(1, 1, 1, cabeceras.length);
    r.setValues([cabeceras]);
    r.setFontWeight('bold');
    r.setBackground('#d1fae5');
    r.setFontColor('#064e3b');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
