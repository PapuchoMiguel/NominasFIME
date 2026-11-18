import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload,
  FileSpreadsheet,
  Play,
  Download,
  Clock,
  Users,
  AlertTriangle,
  CheckCircle2,
  Sun,
  Moon,
  Search,
  ChevronDown,
  ChevronUp,
  Sparkles,
  RefreshCw,
  Info,
  SlidersHorizontal,
  Table as TableIcon,
  FileCheck,
} from 'lucide-react';
import { processPayrollData, formatMinutesToHHMM } from './utils/payrollProcessor';

export default function App() {
  const [rawRows, setRawRows] = useState([]);
  const [fileName, setFileName] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessed, setIsProcessed] = useState(false);
  const [activeTab, setActiveTab] = useState('resultados'); // 'resultados' | 'entrada' | 'reglas'
  const [searchQuery, setSearchQuery] = useState('');
  const [shiftFilter, setShiftFilter] = useState('ALL'); // 'ALL' | 'Diurno' | 'Nocturno'
  const [statusFilter, setStatusFilter] = useState('ALL'); // 'ALL' | 'INCIDENCIAS' | 'LIMPIOS' | 'EXTRAS' | 'DEBIDAS'
  const [expandedEmployee, setExpandedEmployee] = useState(null);
  const [errorMsg, setErrorMsg] = useState('');
  const [loadingSample, setLoadingSample] = useState(false);

  // Opciones avanzadas de interpretación
  const [mealExcessMode, setMealExcessMode] = useState('excess'); // 'excess' (comida - 45) | 'total'
  const [incompleteMealAsMissingPunch, setIncompleteMealAsMissingPunch] = useState(false);

  const fileInputRef = useRef(null);

  // Leer un ArrayBuffer de Excel usando SheetJS
  const parseExcelBuffer = (buffer, name) => {
    try {
      setErrorMsg('');
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      if (!firstSheetName) {
        throw new Error('El archivo Excel no contiene hojas válidas.');
      }
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

      if (!jsonData || jsonData.length === 0) {
        throw new Error('El archivo Excel está vacío.');
      }

      // Validar columnas esperadas
      const firstRowKeys = Object.keys(jsonData[0]).map((k) => k.trim().toLowerCase());
      const hasEmpleado = firstRowKeys.some((k) => k.includes('empleado'));
      const hasFecha = firstRowKeys.some((k) => k.includes('fecha'));

      if (!hasEmpleado || !hasFecha) {
        throw new Error(
          'El archivo no contiene las columnas requeridas (Empleado, Fecha, Entrada, Salida Comida, Regreso Comida, Salida, Tipo de Día).'
        );
      }

      setRawRows(jsonData);
      setFileName(name);
      setIsProcessed(false);
      setExpandedEmployee(null);
    } catch (err) {
      setErrorMsg(err.message || 'Error al leer el archivo Excel.');
    }
  };

  // Manejo de subida de archivo desde input o drag-and-drop
  const handleFileChange = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      parseExcelBuffer(evt.target.result, file.name);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (!file) return;
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      setErrorMsg('Por favor sube un archivo Excel válido (.xlsx).');
      return;
    }
    const reader = new FileReader();
    reader.onload = (evt) => {
      parseExcelBuffer(evt.target.result, file.name);
    };
    reader.readAsArrayBuffer(file);
  };

  // Cargar el archivo de referencia Registros_Semanales_Caoticos_Testing.xlsx
  const handleLoadReferenceFile = async (autoProcess = false) => {
    try {
      setLoadingSample(true);
      setErrorMsg('');
      const response = await fetch('/Registros_Semanales_Caoticos_Testing.xlsx');
      if (!response.ok) {
        throw new Error('No se pudo cargar el archivo de referencia.');
      }
      const buffer = await response.arrayBuffer();
      parseExcelBuffer(buffer, 'Registros_Semanales_Caoticos_Testing.xlsx');
      if (autoProcess) {
        setIsProcessed(true);
        setActiveTab('resultados');
      }
    } catch (err) {
      setErrorMsg(err.message || 'Error al cargar archivo de prueba.');
    } finally {
      setLoadingSample(false);
    }
  };

  // Acción del botón "Procesar Nómina"
  const handleProcessPayroll = () => {
    if (rawRows.length === 0) {
      setErrorMsg('Primero selecciona o sube un archivo Excel (.xlsx).');
      return;
    }
    setErrorMsg('');
    setIsProcessed(true);
    setActiveTab('resultados');
  };

  // Resultados calculados (en minutos internamente)
  const processedResults = useMemo(() => {
    if (!isProcessed || rawRows.length === 0) return [];
    return processPayrollData(rawRows, {
      mealExcessMode,
      incompleteMealAsMissingPunch,
    });
  }, [rawRows, isProcessed, mealExcessMode, incompleteMealAsMissingPunch]);

  // Filtrado de la tabla
  const filteredResults = useMemo(() => {
    return processedResults.filter((item) => {
      const matchesSearch =
        !searchQuery.trim() ||
        item.empleado.toLowerCase().includes(searchQuery.trim().toLowerCase()) ||
        item.observaciones.toLowerCase().includes(searchQuery.trim().toLowerCase());

      const matchesShift = shiftFilter === 'ALL' || item.turno === shiftFilter;

      let matchesStatus = true;
      if (statusFilter === 'INCIDENCIAS') matchesStatus = item.tieneIncidencias;
      if (statusFilter === 'LIMPIOS') matchesStatus = !item.tieneIncidencias;
      if (statusFilter === 'EXTRAS') matchesStatus = item.minutosExtras > 0;
      if (statusFilter === 'DEBIDAS') matchesStatus = item.minutosDebidos > 0;

      return matchesSearch && matchesShift && matchesStatus;
    });
  }, [processedResults, searchQuery, shiftFilter, statusFilter]);

  // Métricas resumen (KPIs)
  const stats = useMemo(() => {
    if (processedResults.length === 0) {
      return {
        totalEmpleados: 0,
        diurnos: 0,
        nocturnos: 0,
        totalTrabajadasHHMM: '00:00',
        totalExtrasHHMM: '00:00',
        totalDebidasHHMM: '00:00',
        conIncidencias: 0,
      };
    }

    let minTrab = 0;
    let minExt = 0;
    let minDeb = 0;
    let diurnos = 0;
    let nocturnos = 0;
    let conIncidencias = 0;

    for (const r of processedResults) {
      minTrab += r.minutosTrabajados;
      minExt += r.minutosExtras;
      minDeb += r.minutosDebidos;
      if (r.turno === 'Diurno') diurnos++;
      else nocturnos++;
      if (r.tieneIncidencias) conIncidencias++;
    }

    return {
      totalEmpleados: processedResults.length,
      diurnos,
      nocturnos,
      totalTrabajadasHHMM: formatMinutesToHHMM(minTrab),
      totalExtrasHHMM: formatMinutesToHHMM(minExt),
      totalDebidasHHMM: formatMinutesToHHMM(minDeb),
      conIncidencias,
    };
  }, [processedResults]);

  // Exportar tabla resultante a un nuevo archivo Excel (.xlsx)
  const handleDownloadExcel = () => {
    if (processedResults.length === 0) return;

    // Estructura exacta solicitada para la tabla de salida
    const exportRows = processedResults.map((row) => ({
      Empleado: row.empleado,
      'Horas Trabajadas': row.horasTrabajadas,
      'Horas Extras': row.horasExtras,
      'Horas Debidas': row.horasDebidas,
      'Horas Totales': row.horasTotales,
      Observaciones: row.observaciones,
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportRows);

    // Ajustar anchos de columna para lectura limpia en Excel
    worksheet['!cols'] = [
      { wch: 24 }, // Empleado
      { wch: 18 }, // Horas Trabajadas
      { wch: 16 }, // Horas Extras
      { wch: 16 }, // Horas Debidas
      { wch: 16 }, // Horas Totales
      { wch: 52 }, // Observaciones
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Nomina_Procesada');

    // Hoja adicional opcional de auditoría con el desglose diario
    const auditRows = [];
    for (const emp of processedResults) {
      for (const d of emp.desgloseDiario) {
        auditRows.push({
          Empleado: emp.empleado,
          Turno: emp.turno,
          Fecha: d.fecha,
          Día: d.dayOfWeekName,
          'Tipo de Día': d.tipoDia,
          Entrada: d.entradaRaw,
          'Salida Comida': d.salidaComidaRaw,
          'Regreso Comida': d.regresoComidaRaw,
          Salida: d.salidaRaw,
          'Comida (min)': d.comidaMin,
          'Retardo (min)': d.retardoMin,
          'Tiempo Efectivo': d.efectivoHHMM,
          Incidencias: d.incidencias.join(' | ') || 'Sin incidencias',
        });
      }
    }
    const auditSheet = XLSX.utils.json_to_sheet(auditRows);
    auditSheet['!cols'] = [
      { wch: 22 },
      { wch: 12 },
      { wch: 14 },
      { wch: 8 },
      { wch: 14 },
      { wch: 16 },
      { wch: 16 },
      { wch: 16 },
      { wch: 16 },
      { wch: 14 },
      { wch: 14 },
      { wch: 16 },
      { wch: 45 },
    ];
    XLSX.utils.book_append_sheet(workbook, auditSheet, 'Desglose_Auditoria');

    const baseName = fileName ? fileName.replace(/\.(xlsx|xls)$/i, '') : 'Nomina';
    XLSX.writeFile(workbook, `Resultado_${baseName}.xlsx`);
  };

  const uniqueEmployeesCount = useMemo(() => {
    const set = new Set();
    for (const r of rawRows) {
      if (r['Empleado']) set.add(String(r['Empleado']).trim());
    }
    return set.size;
  }, [rawRows]);

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col">
      {/* Top Navbar */}
      <header className="sticky top-0 z-30 border-b border-slate-800/80 bg-slate-950/90 backdrop-blur-md">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-indigo-500 to-violet-600 flex items-center justify-center shadow-lg shadow-indigo-500/25">
              <Clock className="w-5 h-5 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-lg font-bold tracking-tight text-white">NóminaPro</h1>
                <span className="text-[11px] font-semibold uppercase tracking-wider px-2 py-0.5 rounded-full bg-indigo-500/15 text-indigo-300 border border-indigo-500/30">
                  Motor en Minutos
                </span>
              </div>
              <p className="text-xs text-slate-400 hidden sm:block">
                Procesamiento automatizado de asistencia, turnos, horas extras e incidencias
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2.5">
            <button
              type="button"
              onClick={() => handleLoadReferenceFile(true)}
              disabled={loadingSample}
              className="inline-flex items-center gap-2 px-3.5 py-2 rounded-lg text-xs font-medium bg-slate-800/90 hover:bg-slate-800 text-slate-200 border border-slate-700/80 transition cursor-pointer disabled:opacity-50"
              title="Cargar y procesar Registros_Semanales_Caoticos_Testing.xlsx"
            >
              <Sparkles className="w-3.5 h-3.5 text-amber-400" />
              <span>Cargar Excel de Referencia (Testing)</span>
            </button>

            {isProcessed && processedResults.length > 0 && (
              <button
                type="button"
                onClick={handleDownloadExcel}
                className="inline-flex items-center gap-2 px-4 py-2 rounded-lg text-xs font-semibold bg-emerald-600 hover:bg-emerald-500 text-white shadow-lg shadow-emerald-600/25 transition cursor-pointer"
              >
                <Download className="w-4 h-4" />
                <span>Descargar Excel (.xlsx)</span>
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">
        {/* Panel Superior: Carga de Archivo + Botón Procesar Nómina */}
        <section className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          {/* Dropzone de Carga Excel */}
          <div className="lg:col-span-8 bg-slate-900/70 border border-slate-800/90 rounded-2xl p-5 shadow-xl">
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="w-4 h-4 text-indigo-400" />
                <h2 className="text-sm font-semibold text-slate-200">
                  1. Archivo de Entrada (.xlsx)
                </h2>
              </div>
              {fileName && (
                <span className="inline-flex items-center gap-1.5 text-xs text-emerald-400 bg-emerald-500/10 border border-emerald-500/20 px-2.5 py-0.5 rounded-full">
                  <FileCheck className="w-3.5 h-3.5" />
                  Archivo cargado
                </span>
              )}
            </div>

            <div
              onDragOver={(e) => {
                e.preventDefault();
                setIsDragging(true);
              }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              className={`relative border-2 border-dashed rounded-xl p-6 text-center transition cursor-pointer ${
                isDragging
                  ? 'border-indigo-400 bg-indigo-500/10'
                  : fileName
                  ? 'border-emerald-500/40 bg-emerald-500/5 hover:border-emerald-500/60'
                  : 'border-slate-700/80 hover:border-indigo-500/60 bg-slate-950/50 hover:bg-slate-900/80'
              }`}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileChange}
                className="hidden"
              />

              <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
                <div
                  className={`w-12 h-12 rounded-2xl flex items-center justify-center shrink-0 ${
                    fileName
                      ? 'bg-emerald-500/15 text-emerald-400 border border-emerald-500/30'
                      : 'bg-indigo-500/15 text-indigo-400 border border-indigo-500/30'
                  }`}
                >
                  {fileName ? (
                    <FileSpreadsheet className="w-6 h-6" />
                  ) : (
                    <Upload className="w-6 h-6" />
                  )}
                </div>

                <div className="text-center sm:text-left">
                  {fileName ? (
                    <>
                      <p className="text-sm font-semibold text-white break-all">{fileName}</p>
                      <p className="text-xs text-slate-400 mt-0.5">
                        <span className="text-emerald-400 font-medium">{rawRows.length}</span>{' '}
                        registros diarios detectados ·{' '}
                        <span className="text-indigo-400 font-medium">{uniqueEmployeesCount}</span>{' '}
                        empleados únicos. Haz clic para cambiar de archivo.
                      </p>
                    </>
                  ) : (
                    <>
                      <p className="text-sm font-medium text-slate-200">
                        Arrastra tu archivo Excel aquí o{' '}
                        <span className="text-indigo-400 underline">explora en tu equipo</span>
                      </p>
                      <p className="text-xs text-slate-400 mt-1">
                        Columnas requeridas: Empleado, Fecha, Entrada, Salida Comida, Regreso Comida, Salida, Tipo de Día
                      </p>
                    </>
                  )}
                </div>
              </div>
            </div>

            {errorMsg && (
              <div className="mt-3 flex items-center gap-2 text-xs text-rose-300 bg-rose-500/10 border border-rose-500/30 rounded-lg px-3.5 py-2.5">
                <AlertTriangle className="w-4 h-4 text-rose-400 shrink-0" />
                <span>{errorMsg}</span>
              </div>
            )}
          </div>

          {/* Panel de Acción: Procesar y Descargar */}
          <div className="lg:col-span-4 bg-slate-900/70 border border-slate-800/90 rounded-2xl p-5 shadow-xl flex flex-col justify-between">
            <div>
              <div className="flex items-center gap-2 mb-2">
                <Play className="w-4 h-4 text-indigo-400" />
                <h2 className="text-sm font-semibold text-slate-200">
                  2. Ejecución y Exportación
                </h2>
              </div>
              <p className="text-xs text-slate-400 leading-relaxed">
                Calcula internamente en minutos los turnos Diurno (48h) y Nocturno (42h), cruces de
                medianoche, regla especial de Arturo Aguilar y genera incidencias semanales.
              </p>
            </div>

            <div className="space-y-2.5 mt-4">
              <button
                type="button"
                onClick={handleProcessPayroll}
                disabled={rawRows.length === 0}
                className="w-full py-3 px-4 rounded-xl font-semibold text-sm flex items-center justify-center gap-2.5 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white shadow-lg shadow-indigo-600/25 transition cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed"
              >
                <Play className="w-4 h-4 fill-current" />
                <span>Procesar Nómina</span>
              </button>

              <button
                type="button"
                onClick={handleDownloadExcel}
                disabled={!isProcessed || processedResults.length === 0}
                className="w-full py-2.5 px-4 rounded-xl font-semibold text-xs flex items-center justify-center gap-2 bg-emerald-600/90 hover:bg-emerald-500 text-white border border-emerald-500/40 shadow-md transition cursor-pointer disabled:opacity-35 disabled:cursor-not-allowed"
              >
                <Download className="w-4 h-4" />
                <span>Descargar Tabla Resultante (.xlsx)</span>
              </button>
            </div>
          </div>
        </section>

        {/* Tarjetas de Resumen KPI cuando ya se procesó la nómina */}
        {isProcessed && processedResults.length > 0 && (
          <section className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3.5">
            <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
              <div className="flex items-center justify-between text-slate-400 text-xs mb-1">
                <span>Empleados</span>
                <Users className="w-4 h-4 text-indigo-400" />
              </div>
              <div className="text-2xl font-bold text-white font-mono-num">
                {stats.totalEmpleados}
              </div>
              <div className="text-[11px] text-slate-400 mt-1 flex items-center gap-2">
                <span className="text-amber-300">{stats.diurnos} Diurnos</span>
                <span>·</span>
                <span className="text-sky-300">{stats.nocturnos} Nocturnos</span>
              </div>
            </div>

            <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
              <div className="flex items-center justify-between text-slate-400 text-xs mb-1">
                <span>Hrs. Trabajadas</span>
                <Clock className="w-4 h-4 text-emerald-400" />
              </div>
              <div className="text-2xl font-bold text-emerald-400 font-mono-num">
                {stats.totalTrabajadasHHMM}
              </div>
              <div className="text-[11px] text-slate-400 mt-1">Topadas al límite de turno</div>
            </div>

            <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
              <div className="flex items-center justify-between text-slate-400 text-xs mb-1">
                <span>Hrs. Extras</span>
                <Sparkles className="w-4 h-4 text-indigo-400" />
              </div>
              <div className="text-2xl font-bold text-indigo-400 font-mono-num">
                {stats.totalExtrasHHMM}
              </div>
              <div className="text-[11px] text-slate-400 mt-1">Tiempo sobre el límite</div>
            </div>

            <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
              <div className="flex items-center justify-between text-slate-400 text-xs mb-1">
                <span>Hrs. Debidas</span>
                <AlertTriangle className="w-4 h-4 text-amber-400" />
              </div>
              <div className="text-2xl font-bold text-amber-400 font-mono-num">
                {stats.totalDebidasHHMM}
              </div>
              <div className="text-[11px] text-slate-400 mt-1">Faltante para el límite</div>
            </div>

            <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
              <div className="flex items-center justify-between text-slate-400 text-xs mb-1">
                <span>Con Incidencias</span>
                <AlertTriangle className="w-4 h-4 text-rose-400" />
              </div>
              <div className="text-2xl font-bold text-rose-400 font-mono-num">
                {stats.conIncidencias}
              </div>
              <div className="text-[11px] text-slate-400 mt-1">Empleados con observaciones</div>
            </div>

            <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
              <div className="flex items-center justify-between text-slate-400 text-xs mb-1">
                <span>Sin Incidencias</span>
                <CheckCircle2 className="w-4 h-4 text-teal-400" />
              </div>
              <div className="text-2xl font-bold text-teal-400 font-mono-num">
                {stats.totalEmpleados - stats.conIncidencias}
              </div>
              <div className="text-[11px] text-slate-400 mt-1">Asistencia limpia</div>
            </div>
          </section>
        )}

        {/* Navegación de Pestañas y Controles */}
        <section className="bg-slate-900/70 border border-slate-800/90 rounded-2xl shadow-xl overflow-hidden">
          <div className="border-b border-slate-800 px-5 pt-4 pb-3 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
            {/* Tabs */}
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => setActiveTab('resultados')}
                className={`inline-flex items-center gap-2 px-3.5 py-2 rounded-lg text-xs font-semibold transition cursor-pointer ${
                  activeTab === 'resultados'
                    ? 'bg-indigo-600 text-white shadow-sm'
                    : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
                }`}
              >
                <TableIcon className="w-3.5 h-3.5" />
                <span>Tabla de Resultados ({processedResults.length})</span>
              </button>

              <button
                type="button"
                onClick={() => setActiveTab('entrada')}
                className={`inline-flex items-center gap-2 px-3.5 py-2 rounded-lg text-xs font-semibold transition cursor-pointer ${
                  activeTab === 'entrada'
                    ? 'bg-indigo-600 text-white shadow-sm'
                    : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
                }`}
              >
                <FileSpreadsheet className="w-3.5 h-3.5" />
                <span>Datos de Entrada ({rawRows.length})</span>
              </button>

              <button
                type="button"
                onClick={() => setActiveTab('reglas')}
                className={`inline-flex items-center gap-2 px-3.5 py-2 rounded-lg text-xs font-semibold transition cursor-pointer ${
                  activeTab === 'reglas'
                    ? 'bg-indigo-600 text-white shadow-sm'
                    : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
                }`}
              >
                <SlidersHorizontal className="w-3.5 h-3.5" />
                <span>Reglas Aplicadas</span>
              </button>
            </div>

            {/* Filtros de búsqueda cuando estamos en Resultados */}
            {activeTab === 'resultados' && isProcessed && (
              <div className="flex flex-wrap items-center gap-2.5">
                <div className="relative">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="Buscar empleado o incidencia..."
                    className="pl-8 pr-3 py-1.5 bg-slate-950 border border-slate-800 rounded-lg text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500 w-56"
                  />
                </div>

                <select
                  value={shiftFilter}
                  onChange={(e) => setShiftFilter(e.target.value)}
                  className="bg-slate-950 border border-slate-800 rounded-lg px-2.5 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-indigo-500"
                >
                  <option value="ALL">Todos los turnos</option>
                  <option value="Diurno">Turno Diurno (48h)</option>
                  <option value="Nocturno">Turno Nocturno (42h)</option>
                </select>

                <select
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                  className="bg-slate-950 border border-slate-800 rounded-lg px-2.5 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-indigo-500"
                >
                  <option value="ALL">Todos los estados</option>
                  <option value="INCIDENCIAS">Con observaciones</option>
                  <option value="LIMPIOS">Sin incidencias</option>
                  <option value="EXTRAS">Con Horas Extras</option>
                  <option value="DEBIDAS">Con Horas Debidas</option>
                </select>
              </div>
            )}
          </div>

          {/* Contenido de la pestaña activa */}
          {activeTab === 'resultados' && (
            <div>
              {!isProcessed ? (
                <div className="py-16 px-6 text-center max-w-lg mx-auto">
                  <div className="w-14 h-14 rounded-2xl bg-indigo-500/10 border border-indigo-500/20 flex items-center justify-center mx-auto mb-4 text-indigo-400">
                    <FileSpreadsheet className="w-7 h-7" />
                  </div>
                  <h3 className="text-base font-semibold text-white mb-1">
                    {rawRows.length > 0
                      ? 'Archivo listo para procesar'
                      : 'Sube un archivo Excel o carga el archivo de prueba'}
                  </h3>
                  <p className="text-xs text-slate-400 mb-6 leading-relaxed">
                    {rawRows.length > 0
                      ? `Se cargaron ${rawRows.length} filas de ${uniqueEmployeesCount} empleados. Presiona "Procesar Nómina" para ejecutar las reglas de negocio.`
                      : 'Puedes subir cualquier archivo .xlsx con la estructura de asistencia o probar de inmediato con Registros_Semanales_Caoticos_Testing.xlsx.'}
                  </p>
                  <div className="flex flex-wrap items-center justify-center gap-3">
                    {rawRows.length > 0 ? (
                      <button
                        type="button"
                        onClick={handleProcessPayroll}
                        className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl text-xs font-semibold bg-indigo-600 hover:bg-indigo-500 text-white shadow-lg shadow-indigo-600/25 transition cursor-pointer"
                      >
                        <Play className="w-4 h-4 fill-current" />
                        <span>Procesar Nómina Ahora</span>
                      </button>
                    ) : (
                      <button
                        type="button"
                        onClick={() => handleLoadReferenceFile(true)}
                        className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl text-xs font-semibold bg-indigo-600 hover:bg-indigo-500 text-white shadow-lg shadow-indigo-600/25 transition cursor-pointer"
                      >
                        <Sparkles className="w-4 h-4" />
                        <span>Cargar y Procesar Archivo de Referencia</span>
                      </button>
                    )}
                  </div>
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-slate-800 bg-slate-950/60 text-[11px] font-semibold uppercase tracking-wider text-slate-400">
                        <th className="py-3.5 pl-5 pr-3">Empleado</th>
                        <th className="py-3.5 px-3 text-center">Horas Trabajadas</th>
                        <th className="py-3.5 px-3 text-center">Horas Extras</th>
                        <th className="py-3.5 px-3 text-center">Horas Debidas</th>
                        <th className="py-3.5 px-3 text-center">Horas Totales</th>
                        <th className="py-3.5 pl-3 pr-5">Observaciones</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/70 text-sm">
                      {filteredResults.map((row) => {
                        const isExpanded = expandedEmployee === row.empleado;
                        return (
                          <React.Fragment key={row.empleado}>
                            <tr
                              onClick={() =>
                                setExpandedEmployee(isExpanded ? null : row.empleado)
                              }
                              className={`group transition cursor-pointer ${
                                isExpanded
                                  ? 'bg-indigo-950/25'
                                  : 'hover:bg-slate-800/40'
                              }`}
                            >
                              {/* Empleado + Badge de Turno */}
                              <td className="py-4 pl-5 pr-3 align-top">
                                <div className="flex items-start gap-2.5">
                                  <button
                                    type="button"
                                    className="mt-0.5 p-1 rounded-md bg-slate-800/80 text-slate-400 group-hover:text-white transition"
                                    title="Ver desglose diario"
                                  >
                                    {isExpanded ? (
                                      <ChevronUp className="w-3.5 h-3.5" />
                                    ) : (
                                      <ChevronDown className="w-3.5 h-3.5" />
                                    )}
                                  </button>
                                  <div>
                                    <div className="font-semibold text-white flex items-center gap-2">
                                      <span>{row.empleado}</span>
                                      {row.empleado.toLowerCase() === 'arturo aguilar' && (
                                        <span className="text-[10px] font-medium px-1.5 py-0.5 rounded bg-purple-500/20 text-purple-300 border border-purple-500/30">
                                          Comida Fija 60m
                                        </span>
                                      )}
                                    </div>
                                    <div className="mt-1 flex items-center gap-2">
                                      {row.turno === 'Diurno' ? (
                                        <span className="inline-flex items-center gap-1 text-[11px] font-medium px-2 py-0.5 rounded-md bg-amber-500/10 text-amber-300 border border-amber-500/25">
                                          <Sun className="w-3 h-3" />
                                          Diurno (Límite 48:00)
                                        </span>
                                      ) : (
                                        <span className="inline-flex items-center gap-1 text-[11px] font-medium px-2 py-0.5 rounded-md bg-sky-500/10 text-sky-300 border border-sky-500/25">
                                          <Moon className="w-3 h-3" />
                                          Nocturno (Límite 42:00)
                                        </span>
                                      )}
                                    </div>
                                  </div>
                                </div>
                              </td>

                              {/* Horas Trabajadas */}
                              <td className="py-4 px-3 text-center align-top">
                                <span className="inline-block font-mono-num font-semibold text-sm px-2.5 py-1 rounded-lg bg-slate-800/90 text-slate-100 border border-slate-700/70">
                                  {row.horasTrabajadas}
                                </span>
                              </td>

                              {/* Horas Extras */}
                              <td className="py-4 px-3 text-center align-top">
                                <span
                                  className={`inline-block font-mono-num font-semibold text-sm px-2.5 py-1 rounded-lg border ${
                                    row.minutosExtras > 0
                                      ? 'bg-indigo-500/15 text-indigo-300 border-indigo-500/35'
                                      : 'bg-slate-900/60 text-slate-500 border-slate-800'
                                  }`}
                                >
                                  {row.horasExtras}
                                </span>
                              </td>

                              {/* Horas Debidas */}
                              <td className="py-4 px-3 text-center align-top">
                                <span
                                  className={`inline-block font-mono-num font-semibold text-sm px-2.5 py-1 rounded-lg border ${
                                    row.minutosDebidos > 0
                                      ? 'bg-amber-500/15 text-amber-300 border-amber-500/35'
                                      : 'bg-slate-900/60 text-slate-500 border-slate-800'
                                  }`}
                                >
                                  {row.horasDebidas}
                                </span>
                              </td>

                              {/* Horas Totales */}
                              <td className="py-4 px-3 text-center align-top">
                                <span className="inline-block font-mono-num font-bold text-sm px-2.5 py-1 rounded-lg bg-emerald-500/10 text-emerald-300 border border-emerald-500/25">
                                  {row.horasTotales}
                                </span>
                              </td>

                              {/* Observaciones */}
                              <td className="py-4 pl-3 pr-5 align-top">
                                {row.tieneIncidencias ? (
                                  <div className="space-y-1 text-xs font-mono-num text-rose-200 bg-rose-500/10 border border-rose-500/25 rounded-xl p-2.5 whitespace-pre-line">
                                    {row.observaciones}
                                  </div>
                                ) : (
                                  <span className="inline-flex items-center gap-1.5 text-xs font-medium px-2.5 py-1 rounded-lg bg-teal-500/10 text-teal-300 border border-teal-500/25">
                                    <CheckCircle2 className="w-3.5 h-3.5" />
                                    Sin incidencias
                                  </span>
                                )}
                              </td>
                            </tr>

                            {/* Fila Expandible: Auditoría Diaria del Empleado */}
                            {isExpanded && (
                              <tr className="bg-slate-950/90">
                                <td colSpan={6} className="px-6 py-4 border-b border-slate-800">
                                  <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
                                    <div className="flex items-center gap-2 text-xs text-slate-300">
                                      <Info className="w-4 h-4 text-indigo-400" />
                                      <span className="font-semibold">
                                        Desglose Diario de {row.empleado}
                                      </span>
                                      <span className="text-slate-500">|</span>
                                      <span>
                                        1ra checada promedio:{' '}
                                        <strong className="text-slate-200">
                                          {row.avgFirstPunchFormatted}
                                        </strong>
                                      </span>
                                      <span className="text-slate-500">|</span>
                                      <span>
                                        Retardo semanal acumulado:{' '}
                                        <strong
                                          className={
                                            row.retardoTotalSemana > 10
                                              ? 'text-rose-400'
                                              : 'text-emerald-400'
                                          }
                                        >
                                          {row.retardoTotalSemana} min
                                        </strong>{' '}
                                        {row.retardoTotalSemana <= 10 && row.retardoTotalSemana > 0
                                          ? '(Dentro de tolerancia <= 10m)'
                                          : ''}
                                      </span>
                                    </div>
                                  </div>

                                  <div className="overflow-x-auto rounded-xl border border-slate-800">
                                    <table className="w-full text-xs text-left">
                                      <thead className="bg-slate-900 text-slate-400 uppercase text-[10px]">
                                        <tr>
                                          <th className="py-2 px-3">Fecha</th>
                                          <th className="py-2 px-3">Tipo de Día</th>
                                          <th className="py-2 px-3">Entrada</th>
                                          <th className="py-2 px-3">Salida Comida</th>
                                          <th className="py-2 px-3">Regreso Comida</th>
                                          <th className="py-2 px-3">Salida</th>
                                          <th className="py-2 px-3 text-center">Comida</th>
                                          <th className="py-2 px-3 text-center">Retardo</th>
                                          <th className="py-2 px-3 text-center">Efectivo Día</th>
                                          <th className="py-2 px-3">Incidencias del Día</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-800/60 font-mono-num">
                                        {row.desgloseDiario.map((d, idx) => (
                                          <tr
                                            key={idx}
                                            className="hover:bg-slate-900/50 text-slate-300"
                                          >
                                            <td className="py-2 px-3 font-medium text-white">
                                              {d.fecha}{' '}
                                              <span className="text-slate-500 font-sans">
                                                ({d.dayOfWeekName})
                                              </span>
                                            </td>
                                            <td className="py-2 px-3 font-sans">
                                              <span
                                                className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                                                  d.estado === 'especial'
                                                    ? 'bg-violet-500/20 text-violet-300 border border-violet-500/30'
                                                    : 'bg-slate-800 text-slate-300'
                                                }`}
                                              >
                                                {d.tipoDia}
                                              </span>
                                            </td>
                                            <td
                                              className={`py-2 px-3 ${
                                                d.entradaInferida ? 'text-amber-300 font-semibold' : ''
                                              }`}
                                            >
                                              {d.entradaRaw}
                                            </td>
                                            <td className="py-2 px-3">{d.salidaComidaRaw}</td>
                                            <td className="py-2 px-3">{d.regresoComidaRaw}</td>
                                            <td
                                              className={`py-2 px-3 ${
                                                d.salidaInferida ? 'text-amber-300 font-semibold' : ''
                                              }`}
                                            >
                                              {d.salidaRaw}
                                            </td>
                                            <td className="py-2 px-3 text-center">
                                              {d.comidaMin}m
                                            </td>
                                            <td className="py-2 px-3 text-center">
                                              {d.retardoMin > 0 ? (
                                                <span className="text-rose-400 font-semibold">
                                                  +{d.retardoMin}m
                                                </span>
                                              ) : (
                                                <span className="text-slate-500">0m</span>
                                              )}
                                            </td>
                                            <td className="py-2 px-3 text-center font-bold text-emerald-300">
                                              {d.efectivoHHMM} ({d.efectivoMin}m)
                                            </td>
                                            <td className="py-2 px-3 font-sans text-[11px]">
                                              {d.incidencias.length > 0 ? (
                                                <span className="text-amber-300">
                                                  {d.incidencias.join(', ')}
                                                </span>
                                              ) : (
                                                <span className="text-slate-500">—</span>
                                              )}
                                            </td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {/* Pestaña: Vista Previa de Datos de Entrada */}
          {activeTab === 'entrada' && (
            <div className="overflow-x-auto max-h-[540px]">
              {rawRows.length === 0 ? (
                <div className="p-12 text-center text-xs text-slate-400">
                  No hay datos cargados todavía.
                </div>
              ) : (
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="sticky top-0 bg-slate-950 border-b border-slate-800 text-slate-400 uppercase text-[10px]">
                    <tr>
                      <th className="py-3 px-4">#</th>
                      <th className="py-3 px-4">Empleado</th>
                      <th className="py-3 px-4">Fecha</th>
                      <th className="py-3 px-4">Entrada</th>
                      <th className="py-3 px-4">Salida Comida</th>
                      <th className="py-3 px-4">Regreso Comida</th>
                      <th className="py-3 px-4">Salida</th>
                      <th className="py-3 px-4">Tipo de Día</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/60 font-mono-num">
                    {rawRows.map((r, i) => (
                      <tr key={i} className="hover:bg-slate-800/40">
                        <td className="py-2.5 px-4 text-slate-500">{i + 1}</td>
                        <td className="py-2.5 px-4 font-sans font-medium text-white">
                          {r['Empleado']}
                        </td>
                        <td className="py-2.5 px-4">{String(r['Fecha'] ?? '')}</td>
                        <td className="py-2.5 px-4">{String(r['Entrada'] || '—')}</td>
                        <td className="py-2.5 px-4">{String(r['Salida Comida'] || '—')}</td>
                        <td className="py-2.5 px-4">{String(r['Regreso Comida'] || '—')}</td>
                        <td className="py-2.5 px-4">{String(r['Salida'] || '—')}</td>
                        <td className="py-2.5 px-4 font-sans">{String(r['Tipo de Día'] ?? r['Tipo de Dia'] ?? '')}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}

          {/* Pestaña: Reglas de Negocio y Parámetros */}
          {activeTab === 'reglas' && (
            <div className="p-6 space-y-6 text-xs text-slate-300">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="bg-slate-950/70 border border-slate-800 rounded-xl p-4 space-y-2">
                  <h3 className="text-sm font-semibold text-white flex items-center gap-2">
                    <Sun className="w-4 h-4 text-amber-400" />
                    Identificación de Turnos y Límites
                  </h3>
                  <ul className="space-y-1.5 text-slate-400 list-disc list-inside">
                    <li>
                      <strong className="text-slate-200">Turno Diurno:</strong> 08:00 a. m. a 04:45
                      p. m. — Límite semanal: <span className="text-amber-300">48:00 hrs</span>{' '}
                      (2880 mins).
                    </li>
                    <li>
                      <strong className="text-slate-200">Turno Nocturno:</strong> 06:00 p. m. a
                      01:45 a. m. — Límite semanal: <span className="text-sky-300">42:00 hrs</span>{' '}
                      (2520 mins).
                    </li>
                    <li>
                      Deducción automática: Es nocturno si el promedio semanal de primeras checadas es
                      antes de las 04:00 a. m. (&lt; 240m) o después de las 03:00 p. m. (&gt; 900m).
                    </li>
                  </ul>
                </div>

                <div className="bg-slate-950/70 border border-slate-800 rounded-xl p-4 space-y-2">
                  <h3 className="text-sm font-semibold text-white flex items-center gap-2">
                    <Clock className="w-4 h-4 text-indigo-400" />
                    Comidas, Medianoche y Días Especiales
                  </h3>
                  <ul className="space-y-1.5 text-slate-400 list-disc list-inside">
                    <li>
                      <strong className="text-slate-200">Cruce de Medianoche:</strong> Si{' '}
                      <code>Salida &lt; Entrada</code>, se suma +24 hrs (+1440 mins) sin cortar a
                      las 11:59 p. m.
                    </li>
                    <li>
                      <strong className="text-slate-200">Excepción Arturo Aguilar:</strong> Ignora
                      sus checadas de comida reales, resta siempre <strong>60 mins</strong> y omite
                      exceso de comida.
                    </li>
                    <li>
                      <strong className="text-slate-200">VACACIONES, Asueto o Festivo:</strong>{' '}
                      Ignora checadas y suma <strong>09:00 hrs</strong> (540 mins).
                    </li>
                  </ul>
                </div>
              </div>

              {/* Opciones ajustables de criterio */}
              <div className="bg-slate-950/70 border border-slate-800 rounded-xl p-4 space-y-3">
                <h3 className="text-sm font-semibold text-white">
                  Parámetros de Formato de Incidencias
                </h3>
                <div className="flex flex-col sm:flex-row gap-6">
                  <label className="flex items-center gap-2.5 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={mealExcessMode === 'total'}
                      onChange={(e) =>
                        setMealExcessMode(e.target.checked ? 'total' : 'excess')
                      }
                      className="rounded border-slate-700 bg-slate-900 text-indigo-600 focus:ring-indigo-500"
                    />
                    <span>
                      En <code>Exceso de comida por X minutos</code>, mostrar minutos totales de
                      comida en vez de solo el excedente sobre 45 mins (actual:{' '}
                      <strong className="text-indigo-400">
                        {mealExcessMode === 'excess' ? 'Excedente (Comida - 45m)' : 'Total Comida'}
                      </strong>
                      )
                    </span>
                  </label>

                  <label className="flex items-center gap-2.5 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={incompleteMealAsMissingPunch}
                      onChange={(e) => setIncompleteMealAsMissingPunch(e.target.checked)}
                      className="rounded border-slate-700 bg-slate-900 text-indigo-600 focus:ring-indigo-500"
                    />
                    <span>
                      Incluir checada de comida incompleta (solo salida o solo regreso) como{' '}
                      <code>Falta checada</code>
                    </span>
                  </label>
                </div>
              </div>
            </div>
          )}
        </section>
      </main>
    </div>
  );
}
