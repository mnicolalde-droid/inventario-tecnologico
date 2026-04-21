const ACTAS_CONFIG = {
  entrega: {
    nombre: 'Acta de entrega de equipos tecnológicos',
    plantillaId: '1rwVSGlBWJx6HtA07eH4e3S65GD_2O4c1WXw-DFbK-Vw',
    carpetaId: '1V7rjK0V2naOJyKbALvi9KjGffU5LVD4p',
    nombreArchivo: 'Acta_Entrega'
  },
  prestamo: {
    nombre: 'Acta de préstamo de equipos tecnológicos',
    plantillaId: '1QIthsAonlVmHjOtLAAE_wuEuLu008lO73mT9zX32n-8',
    carpetaId: '1KlYOfg-PVEZv79i0C09DdTpawLuAXVIh',
    nombreArchivo: 'Acta_Prestamo'
  },
  recepcion: {
    nombre: 'Acta de recepción de equipos tecnológicos',
    plantillaId: '1_rCGYPUBWtebLP9TOsr-aMaMp97-g3oq82p1rr-i33g',
    carpetaId: '1H3VRDw_dmz3N1BG2CwH7HoYlnBLh7tmp',
    nombreArchivo: 'Acta_Recepcion'
  },
  perifericos:{
    nombre: 'Acta de entrega de periféricos tecnológicos',
    plantillaId: '1Hq-6KNdEtN_Gm3CIE3fILEWRMHW9pfxcfucFD2siZnc',
    carpetaId: '1JH0noJvGIdDDHNLqwYhnQRDqB3tq1O6H',
    nombreArchivo: 'Acta_Perifericos'
  },
  devolucion:{
    nombre: 'Acta de devolución de equipos tecnológicos',
    plantillaId: '1moW9ePndeRl5Otaj2SxXCBY4WGj1JvKW9OA_-39WrtU',
    carpetaId: '1vHE122Y36QnSC2YTgk48B44U-zEW5Eiq',
    nombreArchivo: 'Acta_Devolucion'
  },
  cambio:{
    nombre: 'Acta de cambio de equipos tecnológicos',
    plantillaId: '1ypLcZwEvNa-KM9HrUcdp9JqZZ5uLxsCuw2INXlSIr1U',
    carpetaId: '1Fz2E036rqs1xOdugRjQQsEiAf874sKdz',
    nombreArchivo: 'Acta_Cambio'
  }

};

function obtenerActasDisponibles() {
  return Object.keys(ACTAS_CONFIG).map(key => ({
    id: key,
    nombre: ACTAS_CONFIG[key].nombre
  }));
}

function obtenerColaboradores() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Colaboradores');

  return hoja.getDataRange().getValues()
    .slice(1)
    .map((fila, i) => ({
      id: fila[1] ? fila[1].toString() : '',
      nombre: fila[0],
      cargo: fila[2],
      estado: fila[3] ? fila[3].toString().toLowerCase().trim() : '',
      rowIndex: i + 2
    }))
    .filter(colaborador => colaborador.estado === 'activo');
}

function obtenerEquiposDisponibles(tipoActa, colaboradorId) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Equipos');
  const datos = hoja.getDataRange().getValues().slice(1);
  const idBuscado = colaboradorId ? colaboradorId.toString().trim() : '';

  return datos
    .map((fila, i) => ({
      id: i + 2,
      tipo_equipo: fila[2],
      marca: fila[3],
      modelo: fila[4],
      modelo_equipo: fila[4],
      numero_serie: fila[5],
      nombre_equipo: fila[7],
      memoria_ram: fila[8],
      estado_equipo: fila[11] ? fila[11].toString().toLowerCase().trim() : '',
      identificacion_fk: fila[0] ? fila[0].toString() : ''
    }))
    .filter(equipo => {
      if (tipoActa === 'recepcion') {
        return equipo.identificacion_fk === idBuscado;
      }
      return equipo.identificacion_fk === '1111111111' && equipo.estado_equipo === 'bueno';
    });
}

function obtenerAccesoriosDisponibles(tipoActa, colaboradorId) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accesorios');
  const datos = hoja.getDataRange().getValues().slice(1);
  const idBuscado = colaboradorId ? colaboradorId.toString().trim() : '';
 
  return datos
    .map((fila, i) => ({
      id: i + 2,
      identificador: fila[0],
      nombre_accesorio: fila[1],
      marca_accesorio: fila[2],
      estado_accesorio: fila[3] ? fila[3].toString().toLowerCase().trim() : '',
      disponibilidad: fila[4] ? fila[4].toString().toLowerCase().trim() : '',
      colaborador_asignado: fila[6] ? fila[6].toString() : ''
    }))
    .filter(accesorio => {
      if (tipoActa === 'recepcion') {
        return accesorio.colaborador_asignado === idBuscado;
      }
      return accesorio.disponibilidad === 'disponible';
    });
}

function generarActa(data) {
  const config = ACTAS_CONFIG[data.tipoActa];
  if (!config) {
    throw new Error('Tipo de acta no válido.');
  }

  const copia = DriveApp.getFileById(config.plantillaId).makeCopy();
  const doc = DocumentApp.openById(copia.getId());
  const body = doc.getBody();
  const fecha = obtenerFecha();

  const equiposTexto = (data.equipos || []).map(e => {
    const modelo = e.modelo_equipo || e.modelo || '';
    const memoria = e.memoria_ram ? `${e.memoria_ram} RAM` : '';
    return `·  ${e.tipo_equipo || 'Equipo'} ${e.marca || ''} modelo ${modelo}${memoria ? ' - ' + memoria : ''}.`;
  }).join('\n') || 'Ninguno';

  const accesoriosTexto = (data.accesorios || []).map(a =>
    `·  ${a.nombre_accesorio || 'Accesorio'} marca ${a.marca_accesorio || ''}.`
  ).join('\n') || 'Ninguno';

  const equipoPrimero = Array.isArray(data.equipos) && data.equipos.length ? data.equipos[0] : {};
  const accesorioPrimero = Array.isArray(data.accesorios) && data.accesorios.length ? data.accesorios[0] : {};

  body.replaceText('{{nombre}}', data.usuario.nombre || '');
  body.replaceText('{{identificacion}}', data.usuario.id || '');
  body.replaceText('{{cargo}}', data.usuario.cargo || '');
  body.replaceText('{{tipo_equipo}}', equipoPrimero.tipo_equipo || '');
  body.replaceText('{{marca_equipo}}', equipoPrimero.marca || '');
  body.replaceText('{{modelo}}', equipoPrimero.modelo_equipo || equipoPrimero.modelo || '');
  body.replaceText('{{modelo_equipo}}', equipoPrimero.modelo_equipo || equipoPrimero.modelo || '');
  body.replaceText('{{memoria_ram}}', equipoPrimero.memoria_ram || '');
  body.replaceText('{{nombre_accesorio}}', accesorioPrimero.nombre_accesorio || '');
  body.replaceText('{{marca_accesorio}}', accesorioPrimero.marca_accesorio || '');
  body.replaceText('{{equipos}}', equiposTexto);
  body.replaceText('{{accesorios}}', accesoriosTexto);
  body.replaceText('{{dia}}', fecha.dia);
  body.replaceText('{{mes}}', fecha.mes);
  body.replaceText('{{anio}}', fecha.anio);

  doc.saveAndClose();

  const archivo = DriveApp.getFileById(copia.getId());
  const pdf = archivo.getAs('application/pdf');
  const fechaNombre = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  DriveApp.getFolderById(config.carpetaId)
    .createFile(pdf.setName(`${config.nombreArchivo}_${data.usuario.nombre}_${fechaNombre}.pdf`));
  archivo.setTrashed(true);

  actualizarInventario(data);
}

//actualizar inventario automáticamente
function actualizarInventario(data) {
  const hojaEquipos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Equipos');
  const hojaAccesorios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accesorios');
  const estadoEquipo = data.tipoActa === 'prestamo' ? 'Prestado' : 'Asignado';
  const disponibilidadAccesorio = data.tipoActa === 'prestamo' ? 'Prestado' : 'No disponible';

  (data.equipos || []).forEach(eq => {
    if (!eq.rowIndex) return;
    if (data.tipoActa === 'recepcion') {
      hojaEquipos.getRange(eq.rowIndex, 1).setValue('1111111111');
      hojaEquipos.getRange(eq.rowIndex, 12).setValue('Disponible');
      return;
    }
    hojaEquipos.getRange(eq.rowIndex, 1).setValue(data.usuario.id);
    hojaEquipos.getRange(eq.rowIndex, 12).setValue(estadoEquipo);
  });

  (data.accesorios || []).forEach(ac => {
    if (!ac.rowIndex) return;
    if (data.tipoActa === 'recepcion') {
      hojaAccesorios.getRange(ac.rowIndex, 7).setValue('');
      hojaAccesorios.getRange(ac.rowIndex, 5).setValue('Disponible');
      return;
    }
    hojaAccesorios.getRange(ac.rowIndex, 7).setValue(data.usuario.id);
    hojaAccesorios.getRange(ac.rowIndex, 5).setValue(disponibilidadAccesorio);
  });
}

function obtenerFecha() {
  const fecha = new Date();

  const meses = [
    "enero","febrero","marzo","abril","mayo","junio",
    "julio","agosto","septiembre","octubre","noviembre","diciembre"
  ];

  return {
    dia: fecha.getDate().toString(),
    mes: meses[fecha.getMonth()],
    anio: fecha.getFullYear().toString()
  };
}