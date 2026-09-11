/**
 * Backend de "Recordatorio de Mago" (Google Apps Script).
 *
 * Puesta en marcha (una sola vez):
 *  1. Crea una Hoja de Google nueva > Extensiones > Apps Script.
 *  2. Pega este archivo completo y guarda.
 *  3. Ejecuta la función configurarHorario() (menú Ejecutar). Autoriza los
 *     permisos pedidos. Esto crea la pestaña "Horario" con las 81 dosis de
 *     Danna y Miguel ya calculadas.
 *  4. Implementar > Nueva implementación > tipo "Aplicación web":
 *       - Ejecutar como: Yo
 *       - Quién tiene acceso: Cualquier usuario
 *  5. Copia la URL que termina en /exec y pégala como SCRIPT_URL en index.html.
 *
 * Columnas de la pestaña "Horario":
 *   A fila_id | B fecha | C hora | D bebe | E medicamento | F dosis
 *   G confirmada | H confirmada_en
 */

const HOJA = 'Horario';

function doGet(e) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(HOJA);
  const datos = sheet.getDataRange().getValues();
  const ahora = new Date();
  const alarmas = [];

  for (let i = 1; i < datos.length; i++) {
    const [id, fecha, hora, bebe, medicamento, dosis, confirmada] = datos[i];
    if (confirmada === true) continue;

    const programada = combinarFechaHora(fecha, hora);
    const minutosRetraso = Math.floor((ahora - programada) / 60000);
    if (minutosRetraso < 0) continue;

    alarmas.push({
      id: id,
      fila: i + 1,
      imagen_url: '',
      mensaje_voz: `Dale a ${bebe} su ${medicamento}`,
      dosis: dosis,
      minutos_retraso: minutosRetraso
    });
  }

  return ContentService.createTextOutput(JSON.stringify({ alarmas }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const fila = Number(body.fila);
  const sheet = SpreadsheetApp.getActive().getSheetByName(HOJA);
  sheet.getRange(fila, 7).setValue(true);
  sheet.getRange(fila, 8).setValue(new Date());
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function combinarFechaHora(fecha, hora) {
  const f = new Date(fecha);
  const h = new Date(hora);
  return new Date(f.getFullYear(), f.getMonth(), f.getDate(), h.getHours(), h.getMinutes());
}

function configurarHorario() {
  const ss = SpreadsheetApp.getActive();
  const existente = ss.getSheetByName(HOJA);
  if (existente) ss.deleteSheet(existente);
  const sheet = ss.insertSheet(HOJA);
  sheet.appendRow(['fila_id', 'fecha', 'hora', 'bebe', 'medicamento', 'dosis', 'confirmada', 'confirmada_en']);

  const filas = generarDosis();
  filas.forEach((d, idx) => {
    sheet.appendRow([idx + 1, d.fecha, d.hora, d.bebe, d.medicamento, d.dosis, false, '']);
  });

  sheet.getRange(2, 2, filas.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 3, filas.length, 1).setNumberFormat('hh:mm');
}

function generarDosis() {
  function horasDeTratamiento(inicio, intervaloH, dias) {
    const total = (dias * 24) / intervaloH;
    const lista = [];
    for (let k = 0; k < total; k++) {
      lista.push(new Date(inicio.getTime() + k * intervaloH * 3600000));
    }
    return lista;
  }

  const filas = [];
  function agregar(bebe, medicamento, dosisTexto, inicio, intervaloH, dias) {
    horasDeTratamiento(inicio, intervaloH, dias).forEach(t => {
      filas.push({
        bebe, medicamento, dosis: dosisTexto,
        fecha: new Date(t.getFullYear(), t.getMonth(), t.getDate()),
        hora: new Date(1899, 11, 30, t.getHours(), t.getMinutes())
      });
    });
  }

  // Todos inician el 11 de septiembre de 2026
  agregar('Danna', 'Bactrim suspensión', '3.5 ml', new Date(2026, 8, 11, 20, 0), 12, 7);
  agregar('Miguel', 'Bactrim suspensión', '2 ml', new Date(2026, 8, 11, 20, 0), 12, 5);

  agregar('Danna', 'Sensizone infantil', '2.5 ml', new Date(2026, 8, 11, 21, 0), 12, 6);
  agregar('Miguel', 'Sensizone infantil', '0.5 ml', new Date(2026, 8, 11, 21, 0), 12, 6);

  agregar('Danna', 'Cardomicin infantil', '2.5 ml', new Date(2026, 8, 11, 20, 30), 8, 6);
  agregar('Miguel', 'Mucovibrol gotas', '7 gotas', new Date(2026, 8, 11, 20, 30), 8, 5);

  filas.sort((a, b) => {
    const ta = new Date(a.fecha.getFullYear(), a.fecha.getMonth(), a.fecha.getDate(), a.hora.getHours(), a.hora.getMinutes());
    const tb = new Date(b.fecha.getFullYear(), b.fecha.getMonth(), b.fecha.getDate(), b.hora.getHours(), b.hora.getMinutes());
    return ta - tb;
  });

  return filas;
}
