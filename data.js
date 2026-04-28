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

const SISTEMAS_ADMIN_ID = '1111111111';

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

function normalizarTexto(valor) {
  return valor ? valor.toString().toLowerCase().trim() : '';
}

function obtenerMapaColaboradores() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Colaboradores');
  const datos = hoja.getDataRange().getValues().slice(1);
  const mapa = {};

  datos.forEach(fila => {
    const identificacion = fila[1] ? fila[1].toString().trim() : '';
    if (identificacion) {
      mapa[identificacion] = fila[0] || '';
    }
  });

  return mapa;
}

function mapearEquipo(fila, indice, mapaColaboradores = {}) {
  const identificacion = fila[0] ? fila[0].toString().trim() : '';

  return {
    rowIndex: indice + 2,
    tipo_equipo: fila[2],
    marca: fila[3],
    modelo: fila[4],
    modelo_equipo: fila[4],
    numero_serie: fila[5],
    nombre_equipo: fila[7],
    memoria_ram: fila[8],
    almacenamiento: fila[9],
    sistema_operativo: fila[10],
    estado_equipo: normalizarTexto(fila[11]),
    disponibilidad_equipo: normalizarTexto(fila[12]),
    identificacion_fk: identificacion,
    observaciones_equipo: fila[14] ? fila[14].toString() : '',
    proceso: fila[16] ? fila[16].toString() : '',
    nombre_colaborador: mapaColaboradores[identificacion] || ''
  };
}

function mapearAccesorio(fila, indice, mapaColaboradores = {}) {
  const colaboradorAsignado = fila[7] ? fila[7].toString().trim() : '';

  return {
    rowIndex: indice + 2,
    identificador: fila[0],
    nombre_accesorio: fila[1],
    marca_accesorio: fila[2],
    estado_accesorio: normalizarTexto(fila[3]),
    disponibilidad_accesorio: normalizarTexto(fila[4]),
    observaciones_accesorio: fila[5] ? fila[5].toString() : '',
    colaborador_asignado: colaboradorAsignado,
    proceso: fila[9] ? fila[9].toString() : '',
    nombre_colaborador: mapaColaboradores[colaboradorAsignado] || ''
  };
}

function obtenerEquiposDisponibles(tipoActa, colaboradorId) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Equipos');
  const datos = hoja.getDataRange().getValues().slice(1);
  const idBuscado = colaboradorId ? colaboradorId.toString().trim() : '';

  return datos
    .map((fila, i) => mapearEquipo(fila, i))
    .filter(equipo => {
      if (tipoActa === 'recepcion') {
        return equipo.identificacion_fk === idBuscado && equipo.disponibilidad_equipo === 'ocupado';
      }
      if (tipoActa === 'cambio') {
        return (equipo.identificacion_fk === '' || equipo.identificacion_fk === SISTEMAS_ADMIN_ID) &&
          equipo.disponibilidad_equipo === 'disponible' &&
          equipo.estado_equipo === 'bueno';
      }
      return (equipo.identificacion_fk === '' || equipo.identificacion_fk === SISTEMAS_ADMIN_ID) &&
             (equipo.disponibilidad_equipo === 'disponible' || equipo.estado_equipo === 'bueno');
    });
}

function obtenerAccesoriosDisponibles(tipoActa, colaboradorId) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accesorios');
  const datos = hoja.getDataRange().getValues().slice(1);
  const idBuscado = colaboradorId ? colaboradorId.toString().trim() : '';
 
  return datos
    .map((fila, i) => mapearAccesorio(fila, i))
    .filter(accesorio => {
      if (tipoActa === 'recepcion') {
        return accesorio.colaborador_asignado === idBuscado && accesorio.disponibilidad_accesorio === 'ocupado';
      }
      if (tipoActa === 'cambio') {
        return accesorio.disponibilidad_accesorio === 'disponible' && accesorio.estado_accesorio === 'bueno';
      }
      if (tipoActa === 'perifericos') {
        return accesorio.disponibilidad_accesorio === 'disponible' && (accesorio.estado_accesorio === 'bueno' || accesorio.estado_accesorio === '');
      }
      return accesorio.disponibilidad_accesorio === 'disponible';
    });
}

function obtenerEquiposPrestados() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Equipos');
  const mapaColaboradores = obtenerMapaColaboradores();
  const datos = hoja.getDataRange().getValues().slice(1);

  return datos
    .map((fila, i) => mapearEquipo(fila, i, mapaColaboradores))
    .filter(equipo => equipo.disponibilidad_equipo === 'prestado');
}

function obtenerAccesoriosPrestados() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accesorios');
  const mapaColaboradores = obtenerMapaColaboradores();
  const datos = hoja.getDataRange().getValues().slice(1);

  return datos
    .map((fila, i) => mapearAccesorio(fila, i, mapaColaboradores))
    .filter(accesorio => accesorio.disponibilidad_accesorio === 'prestado');
}

function obtenerEquiposOcupados(colaboradorId) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Equipos');
  const mapaColaboradores = obtenerMapaColaboradores();
  const datos = hoja.getDataRange().getValues().slice(1);
  const idBuscado = colaboradorId ? colaboradorId.toString().trim() : '';

  return datos
    .map((fila, i) => mapearEquipo(fila, i, mapaColaboradores))
    .filter(equipo =>
      equipo.identificacion_fk === idBuscado &&
      (equipo.disponibilidad_equipo === 'ocupado' || equipo.disponibilidad_equipo === 'prestado')
    );
}

function obtenerAccesoriosOcupados(colaboradorId) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accesorios');
  const mapaColaboradores = obtenerMapaColaboradores();
  const datos = hoja.getDataRange().getValues().slice(1);
  const idBuscado = colaboradorId ? colaboradorId.toString().trim() : '';

  return datos
    .map((fila, i) => mapearAccesorio(fila, i, mapaColaboradores))
    .filter(accesorio =>
      accesorio.colaborador_asignado === idBuscado &&
      (accesorio.disponibilidad_accesorio === 'ocupado' || accesorio.disponibilidad_accesorio === 'prestado')
    );
}

function escaparRegex(texto) {
  return texto.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function reemplazarTexto(body, marcador, valor) {
  body.replaceText(escaparRegex(marcador), valor || '');
}

function reemplazarPrimeraCoincidencia(body, marcador, valor) {
  const coincidencia = body.findText(escaparRegex(marcador));
  if (!coincidencia) return false;

  const texto = coincidencia.getElement().asText();
  texto.deleteText(coincidencia.getStartOffset(), coincidencia.getEndOffsetInclusive());
  texto.insertText(coincidencia.getStartOffset(), valor || '');
  return true;
}

function reemplazarPrimeraCoincidenciaConEstiloVecino(body, marcador, valor) {
  const coincidencia = body.findText(escaparRegex(marcador));
  if (!coincidencia) return false;

  const texto = coincidencia.getElement().asText();
  const inicio = coincidencia.getStartOffset();
  const fin = coincidencia.getEndOffsetInclusive();
  const contenido = texto.getText();
  let atributos = null;

  if (fin + 1 < contenido.length) {
    atributos = texto.getAttributes(fin + 1);
  } else if (inicio > 0) {
    atributos = texto.getAttributes(inicio - 1);
  } else if (contenido.length) {
    atributos = texto.getAttributes(inicio);
  }

  texto.deleteText(inicio, fin);
  texto.insertText(inicio, valor || '');

  if (atributos && valor) {
    texto.setAttributes(inicio, inicio + valor.length - 1, atributos);
  }

  return true;
}

function guardarPdfActa(copiaId, doc, config, nombreUsuario) {
  doc.saveAndClose();

  const archivo = DriveApp.getFileById(copiaId);
  const pdf = archivo.getAs('application/pdf');
  const fechaNombre = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const nombreArchivo = `${config.nombreArchivo}_${nombreUsuario}_${fechaNombre}.pdf`;
  const archivoPdf = DriveApp.getFolderById(config.carpetaId)
    .createFile(pdf.setName(nombreArchivo));

  archivo.setTrashed(true);

  return {
    fileId: archivoPdf.getId(),
    fileName: nombreArchivo,
    previewUrl: `https://drive.google.com/file/d/${archivoPdf.getId()}/view`
  };
}

function generarActa(data) {
  const config = ACTAS_CONFIG[data.tipoActa];
  if (!config) {
    throw new Error('Tipo de acta no válido.');
  }

  if (data.tipoActa === 'cambio') {
    const resultadoCambio = generarActaCambio(data, config);
    actualizarInventario(data);
    return resultadoCambio;
  }

  const copia = DriveApp.getFileById(config.plantillaId).makeCopy();
  const doc = DocumentApp.openById(copia.getId());
  const body = doc.getBody();
  const fecha = obtenerFecha();

  const equiposTexto = (data.equipos || []).map(e => {
    const modelo = e.modelo_equipo || e.modelo || '';
    const memoria = e.memoria_ram ? `${e.memoria_ram} RAM` : '';
    return `- ${e.tipo_equipo || 'Equipo'} ${e.marca || ''} modelo ${modelo}${memoria ? ' - ' + memoria : ''}.`;
  }).join('\n') || 'Ninguno';

  const accesoriosTexto = (data.accesorios || []).map(a =>
    `- ${a.nombre_accesorio || 'Accesorio'} marca ${a.marca_accesorio || ''}.`
  ).join('\n') || 'Ninguno';

  const equipoPrimero = Array.isArray(data.equipos) && data.equipos.length ? data.equipos[0] : {};
  const accesorioPrimero = Array.isArray(data.accesorios) && data.accesorios.length ? data.accesorios[0] : {};

  // Para devolución, obtener el proceso del equipo si no viene en el payload
  let procesoFinal = data.proceso || '';
  if (data.tipoActa === 'devolucion' && !procesoFinal && equipoPrimero.proceso) {
    procesoFinal = equipoPrimero.proceso;
  }

  // Para recepción, usar el estado seleccionado; para otros, usar el del inventario
  let estadoEquipoFinal = data.tipoActa === 'recepcion' ? (data.estadoEquipo || '') : (equipoPrimero.estado_equipo || '');

  reemplazarTexto(body, '{{nombre}}', data.usuario.nombre || '');
  reemplazarTexto(body, '{{identificacion}}', data.usuario.id || '');
  reemplazarTexto(body, '{{cargo}}', data.usuario.cargo || '');
  reemplazarTexto(body, '{{tipo_equipo}}', equipoPrimero.tipo_equipo || '');
  reemplazarTexto(body, '{{marca_equipo}}', equipoPrimero.marca || '');
  reemplazarTexto(body, '{{modelo}}', equipoPrimero.modelo_equipo || equipoPrimero.modelo || '');
  reemplazarTexto(body, '{{modelo_equipo}}', equipoPrimero.modelo_equipo || equipoPrimero.modelo || '');
  reemplazarTexto(body, '{{memoria_ram}}', equipoPrimero.memoria_ram || '');
  reemplazarTexto(body, '{{almacenamiento}}', equipoPrimero.almacenamiento || '');
  reemplazarTexto(body, '{{sistema_operativo}}', equipoPrimero.sistema_operativo || '');
  reemplazarTexto(body, '{{estado_equipo}}', estadoEquipoFinal);
  reemplazarTexto(body, '{{observaciones_equipo}}', data.observacionesEquipo || '');
  reemplazarTexto(body, '{{nombre_accesorio}}', accesorioPrimero.nombre_accesorio || '');
  reemplazarTexto(body, '{{marca_accesorio}}', accesorioPrimero.marca_accesorio || '');
  reemplazarTexto(body, '{{equipos}}', equiposTexto);
  reemplazarTexto(body, '{{accesorios}}', accesoriosTexto);
  reemplazarTexto(body, '{{observaciones_accesorio}}', data.observacionesAccesorio || '');
  reemplazarTexto(body, '{{proceso}}', procesoFinal);
  reemplazarTexto(body, '{{dia}}', fecha.dia);
  reemplazarTexto(body, '{{mes}}', fecha.mes);
  reemplazarTexto(body, '{{anio}}', fecha.anio);

  const resultado = guardarPdfActa(copia.getId(), doc, config, data.usuario.nombre || 'Sin_nombre');

  actualizarInventario(data);
  return resultado;
}

function generarActaCambio(data, config) {
  if (!data.equipoAnterior || !data.equipoNuevo) {
    throw new Error('Para el acta de cambio se requiere un equipo anterior y un nuevo equipo.');
  }

  const copia = DriveApp.getFileById(config.plantillaId).makeCopy();
  const doc = DocumentApp.openById(copia.getId());
  const body = doc.getBody();
  const fecha = obtenerFecha();
  const accesoriosCambio = []
    .concat(data.accesoriosConservados || [])
    .concat(data.accesoriosDevueltos || [])
    .concat(data.accesoriosAsignados || []);

  const accesoriosTexto = accesoriosCambio.map(a =>
    `${a.nombre_accesorio || 'Accesorio'} marca ${a.marca_accesorio || ''}.`
  ).join('\n') || 'Ninguno';

  const estadosTexto = accesoriosCambio.map(a => a.estado_acta || '').join('\n') || 'Ninguno';
  const accesoriosConEstadoTexto = accesoriosCambio.map(a =>
    `${a.nombre_accesorio || 'Accesorio'} marca ${a.marca_accesorio || ''}. - ${a.estado_acta || ''}`
  ).join('\n') || 'Ninguno';

  reemplazarTexto(body, '{{nombre}}', data.usuario.nombre || '');
  reemplazarTexto(body, '{{identificacion}}', data.usuario.id || '');
  reemplazarTexto(body, '{{cargo}}', data.usuario.cargo || '');
  reemplazarTexto(body, '{{motivo}}', data.motivo || '');
  reemplazarTexto(body, '{{estado_equipo}}', data.estadoEquipoAnterior || data.equipoAnterior.estado_equipo || '');
  reemplazarTexto(body, '{{dia}}', fecha.dia);
  reemplazarTexto(body, '{{mes}}', fecha.mes);
  reemplazarTexto(body, '{{anio}}', fecha.anio);

  const bloqueAccesoriosReemplazado = reemplazarPrimeraCoincidenciaConEstiloVecino(
    body,
    '{{accesorios}} - {{estado}}',
    accesoriosConEstadoTexto
  );

  if (!bloqueAccesoriosReemplazado) {
    reemplazarTexto(body, '{{accesorios}}', accesoriosTexto);
    reemplazarTexto(body, '{{estado}}', estadosTexto);
  } else {
    reemplazarTexto(body, '{{accesorios}}', '');
    reemplazarTexto(body, '{{estado}}', '');
  }

  reemplazarPrimeraCoincidenciaConEstiloVecino(body, '{{tipo_equipo}}', data.equipoAnterior.tipo_equipo || '');
  reemplazarPrimeraCoincidenciaConEstiloVecino(body, '{{tipo_equipo}}', data.equipoNuevo.tipo_equipo || '');
  reemplazarTexto(body, '{{tipo_equipo}}', data.equipoNuevo.tipo_equipo || '');
  reemplazarPrimeraCoincidenciaConEstiloVecino(body, '{{marca_equipo}}', data.equipoAnterior.marca || '');
  reemplazarPrimeraCoincidenciaConEstiloVecino(body, '{{marca_equipo}}', data.equipoNuevo.marca || '');
  reemplazarTexto(body, '{{marca_equipo}}', data.equipoNuevo.marca || '');
  reemplazarPrimeraCoincidenciaConEstiloVecino(body, '{{modelo_equipo}}', data.equipoAnterior.modelo_equipo || data.equipoAnterior.modelo || '');
  reemplazarPrimeraCoincidenciaConEstiloVecino(body, '{{modelo_equipo}}', data.equipoNuevo.modelo_equipo || data.equipoNuevo.modelo || '');
  reemplazarTexto(body, '{{modelo_equipo}}', data.equipoNuevo.modelo_equipo || data.equipoNuevo.modelo || '');
  reemplazarPrimeraCoincidenciaConEstiloVecino(body, '{{modelo}}', data.equipoAnterior.modelo_equipo || data.equipoAnterior.modelo || '');
  reemplazarPrimeraCoincidenciaConEstiloVecino(body, '{{modelo}}', data.equipoNuevo.modelo_equipo || data.equipoNuevo.modelo || '');
  reemplazarTexto(body, '{{modelo}}', data.equipoNuevo.modelo_equipo || data.equipoNuevo.modelo || '');
  reemplazarPrimeraCoincidenciaConEstiloVecino(body, '{{memoria_ram}}', data.equipoAnterior.memoria_ram || '');
  reemplazarPrimeraCoincidenciaConEstiloVecino(body, '{{memoria_ram}}', data.equipoNuevo.memoria_ram || '');
  reemplazarTexto(body, '{{memoria_ram}}', data.equipoNuevo.memoria_ram || '');

  return guardarPdfActa(copia.getId(), doc, config, data.usuario.nombre || 'Sin_nombre');
}

function actualizarInventario(data) {
  const hojaEquipos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Equipos');
  const hojaAccesorios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accesorios');

  if (data.tipoActa === 'cambio') {
    actualizarInventarioCambio(data, hojaEquipos, hojaAccesorios);
    return;
  }

  (data.equipos || []).forEach(eq => {
    if (!eq.rowIndex) return;
    if (data.tipoActa === 'recepcion') {
      hojaEquipos.getRange(eq.rowIndex, 1).setValue(SISTEMAS_ADMIN_ID); // identificacion_fk a bodega
      hojaEquipos.getRange(eq.rowIndex, 12).setValue(data.estadoEquipo || ''); // estado_equipo
      hojaEquipos.getRange(eq.rowIndex, 13).setValue('Disponible'); // disponibilidad_equipo
      hojaEquipos.getRange(eq.rowIndex, 15).setValue(data.observacionesEquipo || 'Ninguna'); // observaciones_equipo
      return;
    }
    if (data.tipoActa === 'devolucion') {
      hojaEquipos.getRange(eq.rowIndex, 1).setValue(SISTEMAS_ADMIN_ID); // identificacion_fk a bodega
      hojaEquipos.getRange(eq.rowIndex, 13).setValue('Disponible'); // disponibilidad_equipo
      hojaEquipos.getRange(eq.rowIndex, 17).setValue(''); // limpiar proceso
      return;
    }
    hojaEquipos.getRange(eq.rowIndex, 1).setValue(data.usuario.id); // identificacion_fk al colaborador
    hojaEquipos.getRange(eq.rowIndex, 13).setValue(data.tipoActa === 'prestamo' ? 'Prestado' : 'Ocupado'); // disponibilidad_equipo
    // Guardar proceso si es préstamo
    if (data.tipoActa === 'prestamo' && data.proceso) {
      hojaEquipos.getRange(eq.rowIndex, 17).setValue(data.proceso);
    }
  });

  (data.accesorios || []).forEach(ac => {
    if (!ac.rowIndex) return;
    if (data.tipoActa === 'recepcion') {
      hojaAccesorios.getRange(ac.rowIndex, 8).setValue(SISTEMAS_ADMIN_ID); // colaborador_asignado a bodega
      hojaAccesorios.getRange(ac.rowIndex, 5).setValue('Disponible'); // disponibilidad_accesorio
      hojaAccesorios.getRange(ac.rowIndex, 6).setValue(data.observacionesAccesorio || 'Ninguna'); // observaciones_accesorio
      return;
    }
    if (data.tipoActa === 'devolucion') {
      hojaAccesorios.getRange(ac.rowIndex, 8).setValue(SISTEMAS_ADMIN_ID); // colaborador_asignado a bodega
      hojaAccesorios.getRange(ac.rowIndex, 5).setValue('Disponible'); // disponibilidad_accesorio
      hojaAccesorios.getRange(ac.rowIndex, 10).setValue(''); // limpiar proceso
      return;
    }
    hojaAccesorios.getRange(ac.rowIndex, 8).setValue(data.usuario.id); // colaborador_asignado al colaborador
    hojaAccesorios.getRange(ac.rowIndex, 5).setValue(data.tipoActa === 'prestamo' ? 'Prestado' : 'Ocupado'); // disponibilidad_accesorio
    // Guardar proceso si es préstamo
    if (data.tipoActa === 'prestamo' && data.proceso) {
      hojaAccesorios.getRange(ac.rowIndex, 10).setValue(data.proceso);
    }
  });
}

function actualizarInventarioCambio(data, hojaEquipos, hojaAccesorios) {
  const equipoAnterior = data.equipoAnterior || {};
  const equipoNuevo = data.equipoNuevo || {};
  const accesoriosConservados = data.accesoriosConservados || [];
  const accesoriosDevueltos = data.accesoriosDevueltos || [];
  const accesoriosAsignados = data.accesoriosAsignados || [];

  if (equipoAnterior.rowIndex) {
    hojaEquipos.getRange(equipoAnterior.rowIndex, 1).setValue(SISTEMAS_ADMIN_ID);
    hojaEquipos.getRange(equipoAnterior.rowIndex, 12).setValue(data.estadoEquipoAnterior || '');
    hojaEquipos.getRange(equipoAnterior.rowIndex, 13).setValue('Disponible');
    hojaEquipos.getRange(equipoAnterior.rowIndex, 17).setValue('');
  }

  if (equipoNuevo.rowIndex) {
    hojaEquipos.getRange(equipoNuevo.rowIndex, 1).setValue(data.usuario.id);
    hojaEquipos.getRange(equipoNuevo.rowIndex, 13).setValue('Ocupado');
    hojaEquipos.getRange(equipoNuevo.rowIndex, 17).setValue('');
  }

  accesoriosConservados.forEach(accesorio => {
    if (!accesorio.rowIndex) return;
    hojaAccesorios.getRange(accesorio.rowIndex, 8).setValue(data.usuario.id);
    hojaAccesorios.getRange(accesorio.rowIndex, 5).setValue('Ocupado');
    hojaAccesorios.getRange(accesorio.rowIndex, 10).setValue('');
  });

  accesoriosDevueltos.forEach(accesorio => {
    if (!accesorio.rowIndex) return;
    hojaAccesorios.getRange(accesorio.rowIndex, 8).setValue(SISTEMAS_ADMIN_ID);
    hojaAccesorios.getRange(accesorio.rowIndex, 5).setValue('Disponible');
    hojaAccesorios.getRange(accesorio.rowIndex, 10).setValue('');
  });

  accesoriosAsignados.forEach(accesorio => {
    if (!accesorio.rowIndex) return;
    hojaAccesorios.getRange(accesorio.rowIndex, 8).setValue(data.usuario.id);
    hojaAccesorios.getRange(accesorio.rowIndex, 5).setValue('Ocupado');
    hojaAccesorios.getRange(accesorio.rowIndex, 10).setValue('');
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
