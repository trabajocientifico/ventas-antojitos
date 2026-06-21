/**
 * BACKEND DE PEDIDOS — Antojitos y algo más
 * Pedidos con cuenta abierta. Numeración diaria (#1, #2, … reinicia cada día).
 * Estados: abierto / pagado / anulado.
 * Los dos celulares comparten las cuentas abiertas (se sincronizan vía polling).
 */

// ===== CONFIGURACIÓN =====
const NOMBRE_HOJA  = 'Pedidos';
const ZONA_HORARIA = 'America/Bogota';

const ENCABEZADOS = [
  'Fecha apertura', 'Fecha cobro', '# Pedido', 'Cliente/Mesa',
  'Vendedor', 'Producto', 'Cantidad', 'Precio unit.',
  'Subtotal', 'Pago', 'Estado', 'ID pedido'
];

const COL = {
  fechaApertura: 0, fechaCobro: 1, numPedido: 2, cliente: 3,
  vendedor: 4, producto: 5, cantidad: 6, precio: 7,
  subtotal: 8, pago: 9, estado: 10, idPedido: 11
};

// ===== ROUTER =====
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    let res;
    if (d.action === 'guardar')     res = guardar_(d, false);
    else if (d.action === 'cobrar') res = guardar_(d, true);
    else if (d.action === 'anular') res = anular_(d);
    else throw new Error('Acción desconocida: ' + d.action);
    return responder_(Object.assign({ ok: true }, res));
  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Devuelve total del día, # de pedidos cobrados y la lista de cuentas abiertas. */
function doGet() {
  const hoja = obtenerHoja_();
  const datos = hoja.getDataRange().getValues();
  const hoy = formatoHoy_();

  let totalHoy = 0, efectivo = 0, transferencia = 0;
  const pagadasHoy = {};
  const abiertosMap = {};

  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    const estado = String(f[COL.estado] || '').toLowerCase();
    const sub = Number(f[COL.subtotal]) || 0;
    const idP = f[COL.idPedido];

    if (estado === 'abierto') {
      const fa = aFecha_(f[COL.fechaApertura]);
      if (!fa) continue;
      if (!abiertosMap[idP]) {
        abiertosMap[idP] = {
          idPedido: idP,
          pedido: Number(f[COL.numPedido]) || 0,
          cliente: f[COL.cliente] || '',
          vendedor: f[COL.vendedor] || '',
          abierta: fa.getTime(),
          items: [],
          total: 0
        };
      }
      abiertosMap[idP].items.push({
        producto: f[COL.producto],
        cantidad: Number(f[COL.cantidad]) || 0,
        precio: Number(f[COL.precio]) || 0
      });
      abiertosMap[idP].total += sub;
    } else if (estado === 'pagado') {
      const fc = aFecha_(f[COL.fechaCobro]);
      if (!fc) continue;
      if (Utilities.formatDate(fc, ZONA_HORARIA, 'yyyy-MM-dd') !== hoy) continue;
      totalHoy += sub;
      if (String(f[COL.pago] || '').toLowerCase().indexOf('efect') >= 0) efectivo += sub;
      else transferencia += sub;
      if (idP) pagadasHoy[idP] = true;
    }
  }

  const abiertos = Object.keys(abiertosMap)
    .map(function (k) { return abiertosMap[k]; })
    .sort(function (a, b) { return a.pedido - b.pedido; });

  return responder_({
    ok: true,
    totalHoy: totalHoy,
    efectivo: efectivo,
    transferencia: transferencia,
    ventas: Object.keys(pagadasHoy).length,
    abiertos: abiertos
  });
}

// ===== ACCIONES =====

/**
 * Crea o actualiza un pedido. Si cobrando=true, queda pagado en una sola operación.
 * Reescribe las filas "abiertas" del pedido con la lista actual de items.
 */
function guardar_(d, cobrando) {
  const items = (d.items || []).filter(function (it) { return Number(it.cantidad) > 0; });
  if (items.length === 0) throw new Error('No hay productos en el pedido');

  const hoja = obtenerHoja_();
  const datos = hoja.getDataRange().getValues();

  let idPedido = d.idPedido || '';
  let numPedido = null;
  let fechaApertura = null;
  const filasABorrar = [];

  if (idPedido) {
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][COL.idPedido] === idPedido &&
          String(datos[i][COL.estado]).toLowerCase() === 'abierto') {
        if (!fechaApertura) {
          fechaApertura = datos[i][COL.fechaApertura];
          numPedido = Number(datos[i][COL.numPedido]) || 0;
        }
        filasABorrar.push(i + 1);
      }
    }
    // Si vino idPedido pero ya no hay filas abiertas, otro dispositivo lo cobró/anuló.
    if (!fechaApertura) {
      throw new Error('Este pedido ya fue cobrado o anulado. Recarga la lista.');
    }
  } else {
    // Pedido nuevo
    idPedido = nuevoIdPedido_();
    numPedido = siguienteNumeroPedido_(datos);
    fechaApertura = new Date();
  }

  // Borra filas viejas (de abajo hacia arriba para no descuadrar índices)
  filasABorrar.sort(function (a, b) { return b - a; }).forEach(function (r) { hoja.deleteRow(r); });

  const ahora = new Date();
  const fechaCobro = cobrando ? ahora : '';
  const estado = cobrando ? 'pagado' : 'abierto';
  const pago = cobrando ? (d.pago || '') : '';

  let total = 0;
  const filas = items.map(function (it) {
    const cant = Number(it.cantidad) || 0;
    const pre  = Number(it.precio)   || 0;
    const sub  = cant * pre;
    total += sub;
    return [
      fechaApertura, fechaCobro, numPedido, d.cliente || '',
      d.vendedor || '', it.producto, cant, pre,
      sub, pago, estado, idPedido
    ];
  });

  const filaInicio = hoja.getLastRow() + 1;
  hoja.getRange(filaInicio, 1, filas.length, ENCABEZADOS.length).setValues(filas);

  // Asegura que las columnas de fecha se guarden con formato de fecha
  // (de lo contrario doGet puede recibirlas como string y descartarlas).
  hoja.getRange(filaInicio, COL.fechaApertura + 1, filas.length, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');
  hoja.getRange(filaInicio, COL.fechaCobro + 1, filas.length, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');

  return { pedido: numPedido, idPedido: idPedido, total: total };
}

/** Marca las filas abiertas del pedido como anuladas. No toca filas ya pagadas. */
function anular_(d) {
  if (!d.idPedido) throw new Error('Falta idPedido');
  const hoja = obtenerHoja_();
  const datos = hoja.getDataRange().getValues();
  let cambiadas = 0;
  let yaPagado = false;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][COL.idPedido] === d.idPedido) {
      const e = String(datos[i][COL.estado]).toLowerCase();
      if (e === 'abierto') {
        hoja.getRange(i + 1, COL.estado + 1).setValue('anulado');
        cambiadas++;
      } else if (e === 'pagado') {
        yaPagado = true;
      }
    }
  }
  if (cambiadas === 0) {
    throw new Error(yaPagado
      ? 'Este pedido ya fue cobrado, no se puede anular.'
      : 'Pedido no encontrado');
  }
  return {};
}

// ===== Auxiliares =====
/**
 * Siguiente # de pedido para hoy. Considera los pedidos de hoy Y cualquier
 * cuenta abierta rezagada (de días anteriores), para que no haya dos "#5" a la vez.
 */
function siguienteNumeroPedido_(datos) {
  const hoy = formatoHoy_();
  let max = 0;
  for (let i = 1; i < datos.length; i++) {
    const fa = aFecha_(datos[i][COL.fechaApertura]);
    if (!fa) continue;
    const esHoy = Utilities.formatDate(fa, ZONA_HORARIA, 'yyyy-MM-dd') === hoy;
    const esAbierto = String(datos[i][COL.estado] || '').toLowerCase() === 'abierto';
    if (!esHoy && !esAbierto) continue;
    const n = Number(datos[i][COL.numPedido]) || 0;
    if (n > max) max = n;
  }
  return max + 1;
}

/**
 * Convierte a Date un valor leído de la hoja, sea cual sea su formato:
 * Date nativo, string ISO o legible, o número serial de Sheets/Excel.
 * Devuelve null si no se puede interpretar.
 */
function aFecha_(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === 'number' && isFinite(v)) {
    // Serial de Sheets: días desde 1899-12-30
    const ms = Math.round((v - 25569) * 86400000);
    const d = new Date(ms);
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof v === 'string') {
    const s = v.trim();
    if (!s) return null;
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function nuevoIdPedido_() {
  return Date.now() + '-' + Math.random().toString(36).slice(2, 8);
}

function formatoHoy_() {
  return Utilities.formatDate(new Date(), ZONA_HORARIA, 'yyyy-MM-dd');
}

function obtenerHoja_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
    hoja.appendRow(ENCABEZADOS);
    hoja.setFrozenRows(1);
  } else if (hoja.getLastRow() === 0) {
    hoja.appendRow(ENCABEZADOS);
    hoja.setFrozenRows(1);
  }
  // Garantiza que las columnas de fecha NO queden formateadas como texto.
  hoja.getRange(2, COL.fechaApertura + 1, hoja.getMaxRows() - 1, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');
  hoja.getRange(2, COL.fechaCobro + 1, hoja.getMaxRows() - 1, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');
  return hoja;
}

function responder_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
