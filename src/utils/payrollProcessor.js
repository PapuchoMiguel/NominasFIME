// Lógica central de procesamiento de nóminas (100% en minutos internamente)

export const SHIFT_CONFIG = {
  DIURNO: {
    nombre: 'Diurno',
    entradaOficialMin: 8 * 60, // 08:00 a. m. = 480 mins
    salidaOficialMin: 16 * 60 + 45, // 04:45 p. m. = 1005 mins
    limiteSemanalMin: 48 * 60, // 48 horas = 2880 mins
    horarioTexto: '08:00 a. m. - 04:45 p. m. (Límite 48h)',
  },
  NOCTURNO: {
    nombre: 'Nocturno',
    entradaOficialMin: 18 * 60, // 06:00 p. m. = 1080 mins
    salidaOficialMin: 1 * 60 + 45, // 01:45 a. m. = 105 mins (+1440 al cruzar medianoche = 1545 mins)
    salidaOficialComparacionMin: 24 * 60 + 1 * 60 + 45, // 1545 mins
    limiteSemanalMin: 42 * 60, // 42 horas = 2520 mins
    horarioTexto: '06:00 p. m. - 01:45 a. m. (Límite 42h)',
  },
};

/**
 * Convierte un valor de hora (String "08:00 a. m.", "13:45", Date o fracción de Excel)
 * a minutos desde las 00:00 (0 a 1439). Retorna null si está vacío o inválido.
 */
export function parseTimeToMinutes(value) {
  if (value === null || value === undefined) return null;

  // Si viene como número serial de Excel (fracción de día 0.0 a 1.0)
  if (typeof value === 'number' && !Number.isNaN(value)) {
    const fraction = value - Math.floor(value);
    return Math.round(fraction * 1440) % 1440;
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.getHours() * 60 + value.getMinutes();
  }

  const raw = String(value).trim();
  if (!raw) return null;

  // Normalizar formato español/inglés: "08:00 a. m.", "08:00 a.m.", "08:00 am", "16:30"
  const clean = raw
    .toLowerCase()
    .replace(/\./g, '')
    .replace(/\s+/g, ' ')
    .trim();

  const match = clean.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(am|pm|a m|p m)?$/i);
  if (!match) return null;

  let hours = parseInt(match[1], 10);
  const minutes = parseInt(match[2], 10);
  const meridiem = match[3] ? match[3].replace(/\s+/g, '') : null;

  if (minutes < 0 || minutes > 59 || hours < 0 || hours > 24) return null;

  if (meridiem === 'am') {
    if (hours === 12) hours = 0;
  } else if (meridiem === 'pm') {
    if (hours !== 12) hours += 12;
  }

  return hours * 60 + minutes;
}

/**
 * Formatea minutos totales a formato HH:MM (ej. 2880 -> "48:00", 75 -> "01:15")
 */
export function formatMinutesToHHMM(minutes) {
  const safeMins = Math.max(0, Math.round(minutes || 0));
  const hh = Math.floor(safeMins / 60);
  const mm = safeMins % 60;
  return `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
}

/**
 * Formatea minutos del reloj (0-1439+) a formato legible 12h (ej. 480 -> "08:00 a. m.")
 */
export function formatClockMinutes(minutes) {
  if (minutes === null || minutes === undefined) return '—';
  const normalized = ((Math.round(minutes) % 1440) + 1440) % 1440;
  const h24 = Math.floor(normalized / 60);
  const mm = String(normalized % 60).padStart(2, '0');
  const period = h24 >= 12 ? 'p. m.' : 'a. m.';
  const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
  return `${String(h12).padStart(2, '0')}:${mm} ${period}`;
}

/**
 * Parsea la fecha en formato dd/mm/yyyy, Date o serial de Excel.
 */
export function parseDateInfo(value) {
  if (!value && value !== 0) {
    return { ddmm: 'S/F', fullDate: '', isSaturday: false, dayOfWeekName: '', timestamp: 0 };
  }

  let day, month, year;

  if (typeof value === 'number') {
    // Número serial de fecha de Excel (base 1899-12-30)
    const utcDays = Math.floor(value - 25569);
    const utcValue = utcDays * 86400 * 1000;
    const dateObj = new Date(utcValue);
    day = dateObj.getUTCDate();
    month = dateObj.getUTCMonth() + 1;
    year = dateObj.getUTCFullYear();
  } else if (value instanceof Date && !Number.isNaN(value.getTime())) {
    day = value.getDate();
    month = value.getMonth() + 1;
    year = value.getFullYear();
  } else {
    const str = String(value).trim();
    const parts = str.split(/[/-]/);
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        // yyyy-mm-dd
        year = parseInt(parts[0], 10);
        month = parseInt(parts[1], 10);
        day = parseInt(parts[2], 10);
      } else {
        // dd/mm/yyyy
        day = parseInt(parts[0], 10);
        month = parseInt(parts[1], 10);
        year = parseInt(parts[2], 10);
      }
    }
  }

  if (!day || !month || !year) {
    const rawStr = String(value).trim();
    return { ddmm: rawStr.slice(0, 5), fullDate: rawStr, isSaturday: false, dayOfWeekName: '', timestamp: 0 };
  }

  const dateObj = new Date(year, month - 1, day);
  const dayOfWeek = dateObj.getDay(); // 0 = Dom, 6 = Sáb
  const daysNames = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
  const dd = String(day).padStart(2, '0');
  const mm = String(month).padStart(2, '0');

  return {
    ddmm: `${dd}/${mm}`,
    fullDate: `${dd}/${mm}/${year}`,
    isSaturday: dayOfWeek === 6,
    dayOfWeekName: daysNames[dayOfWeek] || '',
    timestamp: dateObj.getTime(),
  };
}

/**
 * Deduce el turno de un empleado según sus primeras checadas promedio de la semana.
 * Regla: "Deduce el turno si las primeras checadas promedio de la semana son
 * antes de las 4:00 a. m. (< 240 mins) o después de las 3:00 p. m. (> 900 mins) (es nocturno)."
 */
export function deduceEmployeeShift(rows) {
  const firstPunches = [];

  for (const row of rows) {
    const tipo = String(row['Tipo de Día'] ?? row['Tipo de Dia'] ?? '').trim().toLowerCase();
    if (['vacaciones', 'asueto', 'festivo'].includes(tipo)) continue;

    const entradaMin = parseTimeToMinutes(row['Entrada']);
    if (entradaMin !== null) {
      firstPunches.push(entradaMin);
    }
  }

  if (firstPunches.length === 0) {
    return {
      shiftKey: 'DIURNO',
      shift: SHIFT_CONFIG.DIURNO,
      avgFirstPunchMin: 480,
    };
  }

  // Si hay checadas nocturnas mezcladas antes y después de medianoche (ej. 23:45 y 00:15),
  // verificamos primero el promedio directo:
  const sumDirect = firstPunches.reduce((acc, v) => acc + v, 0);
  const avgDirect = sumDirect / firstPunches.length;

  // También verificamos si todas las checadas caen en ventana nocturna (< 4:00 a.m. o > 3:00 p.m.)
  const allInNightWindow = firstPunches.every((m) => m < 240 || m > 900);

  const isNocturno = avgDirect < 240 || avgDirect > 900 || allInNightWindow;

  return {
    shiftKey: isNocturno ? 'NOCTURNO' : 'DIURNO',
    shift: isNocturno ? SHIFT_CONFIG.NOCTURNO : SHIFT_CONFIG.DIURNO,
    avgFirstPunchMin: Math.round(avgDirect),
  };
}

/**
 * Procesa todos los registros del Excel y devuelve el resumen por empleado + desglose diario.
 */
export function processPayrollData(rawRows, options = {}) {
  const {
    mealExcessMode = 'excess', // 'excess' (comida - 45) | 'total' (comida total)
    incompleteMealAsMissingPunch = false, // true si comida incompleta también anota "Falta checada"
  } = options;

  // Agrupar filas por empleado manteniendo el orden de aparición
  const employeeMap = new Map();

  for (const row of rawRows) {
    const empNameRaw = row['Empleado'];
    if (!empNameRaw || !String(empNameRaw).trim()) continue;
    const empName = String(empNameRaw).trim();

    if (!employeeMap.has(empName)) {
      employeeMap.set(empName, []);
    }
    employeeMap.get(empName).push(row);
  }

  const results = [];

  for (const [empleado, empRows] of employeeMap.entries()) {
    // Ordenar cronológicamente por fecha si es posible
    const sortedRows = [...empRows].sort((a, b) => {
      const da = parseDateInfo(a['Fecha']).timestamp;
      const db = parseDateInfo(b['Fecha']).timestamp;
      return da - db;
    });

    const { shiftKey, shift, avgFirstPunchMin } = deduceEmployeeShift(sortedRows);
    const isArturoAguilar = empleado.toLowerCase() === 'arturo aguilar';

    let totalMinutosSemana = 0;
    let retardoTotalSemana = 0;
    const incidenciasDiarias = [];
    const desgloseDiario = [];

    for (const row of sortedRows) {
      const dateInfo = parseDateInfo(row['Fecha']);
      const { ddmm, fullDate, isSaturday, dayOfWeekName } = dateInfo;

      const tipoDiaRaw = String(row['Tipo de Día'] ?? row['Tipo de Dia'] ?? 'Regular').trim();
      const tipoDiaLower = tipoDiaRaw.toLowerCase();
      const esDiaEspecial = ['vacaciones', 'asueto', 'festivo'].includes(tipoDiaLower);

      const rawEntrada = parseTimeToMinutes(row['Entrada']);
      const rawSalidaComida = parseTimeToMinutes(row['Salida Comida']);
      const rawRegresoComida = parseTimeToMinutes(row['Regreso Comida']);
      const rawSalida = parseTimeToMinutes(row['Salida']);

      const dayIncidencias = [];

      // 1. Días Especiales: VACACIONES, Asueto, Festivo -> Ignora checadas y suma 09:00 hrs (540 mins)
      if (esDiaEspecial) {
        const efectivoDia = 540;
        totalMinutosSemana += efectivoDia;
        desgloseDiario.push({
          fecha: fullDate,
          ddmm,
          dayOfWeekName,
          isSaturday,
          tipoDia: tipoDiaRaw,
          entradaRaw: row['Entrada'] || '—',
          salidaComidaRaw: row['Salida Comida'] || '—',
          regresoComidaRaw: row['Regreso Comida'] || '—',
          salidaRaw: row['Salida'] || '—',
          comidaMin: 0,
          retardoMin: 0,
          efectivoMin: efectivoDia,
          efectivoHHMM: formatMinutesToHHMM(efectivoDia),
          incidencias: [`Día especial (${tipoDiaRaw}): +09:00 hrs`],
          estado: 'especial',
        });
        continue;
      }

      // 2. Verificar si todo el día está vacío (Falta)
      const sinRegistros =
        rawEntrada === null &&
        rawSalidaComida === null &&
        rawRegresoComida === null &&
        rawSalida === null;

      if (sinRegistros) {
        const inc = `- ${ddmm}: Falta`;
        incidenciasDiarias.push(inc);
        dayIncidencias.push(inc);
        desgloseDiario.push({
          fecha: fullDate,
          ddmm,
          dayOfWeekName,
          isSaturday,
          tipoDia: tipoDiaRaw,
          entradaRaw: '—',
          salidaComidaRaw: '—',
          regresoComidaRaw: '—',
          salidaRaw: '—',
          comidaMin: 0,
          retardoMin: 0,
          efectivoMin: 0,
          efectivoHHMM: '00:00',
          incidencias: dayIncidencias,
          estado: 'falta',
        });
        continue;
      }

      // 3. Injerencias (Inferencia de Entrada o Salida faltante)
      let entradaMin = rawEntrada;
      let salidaMin = rawSalida;
      let faltaChecada = false;
      let entradaInferida = false;
      let salidaInferida = false;

      if (entradaMin === null && (salidaMin !== null || rawSalidaComida !== null || rawRegresoComida !== null)) {
        entradaMin = shift.entradaOficialMin;
        faltaChecada = true;
        entradaInferida = true;
      }

      if (salidaMin === null && (entradaMin !== null || rawSalidaComida !== null || rawRegresoComida !== null)) {
        salidaMin = shift.salidaOficialMin;
        faltaChecada = true;
        salidaInferida = true;
      }

      const comidaIncompleta =
        (rawSalidaComida !== null && rawRegresoComida === null) ||
        (rawSalidaComida === null && rawRegresoComida !== null);

      if (incompleteMealAsMissingPunch && comidaIncompleta) {
        faltaChecada = true;
      }

      // 4. Tolerancia y Retardos (Los sábados no se cuenta retardo)
      let retardoDia = 0;
      if (!isSaturday && !entradaInferida && entradaMin !== null) {
        if (shiftKey === 'DIURNO') {
          if (entradaMin > shift.entradaOficialMin) {
            retardoDia = entradaMin - shift.entradaOficialMin;
          }
        } else {
          // Turno Nocturno: Entrada oficial 06:00 p. m. (1080 mins).
          // Si llegó después de medianoche (< 240 mins, ej. 12:30 a. m.), sumamos 1440 para comparar.
          const entradaComparacion = entradaMin < 240 ? entradaMin + 1440 : entradaMin;
          if (entradaComparacion > shift.entradaOficialMin) {
            retardoDia = entradaComparacion - shift.entradaOficialMin;
          }
        }
      }
      retardoTotalSemana += retardoDia;

      // 5. Cruces de Medianoche: Si Salida < Entrada, asumimos cruce de medianoche y sumamos +1440 mins (+24h)
      let salidaAjustada = salidaMin;
      if (salidaAjustada < entradaMin) {
        salidaAjustada += 1440;
      }

      // 6. Verificar Salida Anticipada (comparando última salida del día con salida oficial del turno)
      let salidaAnticipada = false;
      if (!salidaInferida) {
        if (shiftKey === 'DIURNO') {
          if (salidaAjustada < shift.salidaOficialMin) {
            salidaAnticipada = true;
          }
        } else {
          // En turno nocturno la salida oficial es 01:45 a. m. del día siguiente (1545 mins).
          // Si el empleado entró después de medianoche (< 240 mins) y salió esa misma madrugada (< 1440),
          // ajustamos +1440 para comparar contra la línea de tiempo del turno nocturno.
          const salidaComparacionNocturna =
            entradaMin < 240 && salidaAjustada < 1440 ? salidaAjustada + 1440 : salidaAjustada;
          if (salidaComparacionNocturna < shift.salidaOficialComparacionMin) {
            salidaAnticipada = true;
          }
        }
      }

      if (salidaAnticipada) {
        const inc = `- ${ddmm}: Salida anticipada`;
        incidenciasDiarias.push(inc);
        dayIncidencias.push(inc);
      }

      // 7. Regla de Comida y Excepción "Arturo Aguilar"
      let comidaMin = 0;
      if (isArturoAguilar) {
        comidaMin = 60; // Siempre se le restan exactamente 60 minutos
      } else if (rawSalidaComida !== null && rawRegresoComida !== null) {
        let regresoAjustado = rawRegresoComida;
        if (regresoAjustado < rawSalidaComida) {
          regresoAjustado += 1440;
        }
        comidaMin = regresoAjustado - rawSalidaComida;

        if (comidaMin > 45) {
          const minsReportados = mealExcessMode === 'total' ? comidaMin : comidaMin - 45;
          const inc = `- ${ddmm}: Exceso de comida por ${minsReportados} minutos`;
          incidenciasDiarias.push(inc);
          dayIncidencias.push(inc);
        }
      } else {
        comidaMin = 0;
      }

      // 8. Registrar Falta checada si hubo inferencia
      if (faltaChecada) {
        const inc = `- ${ddmm}: Falta checada`;
        incidenciasDiarias.push(inc);
        dayIncidencias.push(inc);
      }

      // 9. Tiempo efectivo del día (en minutos)
      const tiempoBrutoDia = salidaAjustada - entradaMin;
      const efectivoDia = Math.max(0, tiempoBrutoDia - comidaMin);
      totalMinutosSemana += efectivoDia;

      desgloseDiario.push({
        fecha: fullDate,
        ddmm,
        dayOfWeekName,
        isSaturday,
        tipoDia: tipoDiaRaw,
        entradaRaw: row['Entrada'] || (entradaInferida ? `${formatClockMinutes(entradaMin)} (Inf.)` : '—'),
        salidaComidaRaw: row['Salida Comida'] || '—',
        regresoComidaRaw: row['Regreso Comida'] || '—',
        salidaRaw: row['Salida'] || (salidaInferida ? `${formatClockMinutes(salidaMin)} (Inf.)` : '—'),
        entradaInferida,
        salidaInferida,
        comidaIncompleta,
        comidaMin,
        retardoMin: retardoDia,
        efectivoMin: efectivoDia,
        efectivoHHMM: formatMinutesToHHMM(efectivoDia),
        incidencias: dayIncidencias,
        estado: dayIncidencias.length > 0 || retardoDia > 0 ? 'incidencia' : 'ok',
      });
    }

    // Cálculo de totales semanales según límites del turno
    const limiteTurnoMin = shift.limiteSemanalMin;
    const minutosTrabajados = Math.min(totalMinutosSemana, limiteTurnoMin);
    const minutosExtras = totalMinutosSemana > limiteTurnoMin ? totalMinutosSemana - limiteTurnoMin : 0;
    const minutosDebidos = totalMinutosSemana < limiteTurnoMin ? limiteTurnoMin - totalMinutosSemana : 0;
    const minutosTotales = minutosTrabajados + minutosExtras; // Igual a totalMinutosSemana

    // Construcción de Observaciones finales
    const observacionesLista = [];
    if (retardoTotalSemana > 10) {
      observacionesLista.push(`- Retardo total en la semana: ${retardoTotalSemana} minutos`);
    }
    observacionesLista.push(...incidenciasDiarias);

    const observacionesTexto =
      observacionesLista.length > 0 ? observacionesLista.join('\n') : 'Sin incidencias';

    results.push({
      empleado,
      turno: shift.nombre,
      limiteTurnoHHMM: formatMinutesToHHMM(limiteTurnoMin),
      avgFirstPunchFormatted: formatClockMinutes(avgFirstPunchMin),
      minutosTrabajados,
       minutosExtras,
      minutosDebidos,
      minutosTotales,
      retardoTotalSemana,
      horasTrabajadas: formatMinutesToHHMM(minutosTrabajados),
      horasExtras: formatMinutesToHHMM(minutosExtras),
      horasDebidas: formatMinutesToHHMM(minutosDebidos),
      horasTotales: formatMinutesToHHMM(minutosTotales),
      observaciones: observacionesTexto,
      observacionesLista,
      tieneIncidencias: observacionesLista.length > 0,
      desgloseDiario,
    });
  }

  return results;
}
