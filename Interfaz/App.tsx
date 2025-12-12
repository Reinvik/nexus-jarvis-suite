import React, { useState, useEffect, useRef } from 'react';
import { db, storage } from './firebaseConfig';
import { collection, addDoc, query, orderBy, limit, onSnapshot, serverTimestamp, updateDoc, doc } from 'firebase/firestore';
import { ref, uploadBytes, getDownloadURL } from 'firebase/storage';
import { BotDefinition, BotType, Order } from './types';
import {
  Bot,
  FileSpreadsheet,
  Truck,
  Eye,
  ShieldCheck,
  ArrowRightLeft,
  Calculator,
  Activity,
  CheckCircle2,
  XCircle,
  Clock,
  Loader2,
  UploadCloud,
  Zap,
  BarChart3,
  Cpu,
  LayoutGrid,
  Calendar,
  Search,
  ChevronRight,
  Terminal,
  FileText,
  User,
  Settings,
  Save,
  RotateCcw,
  X,
  Mail,
  FolderSync
} from 'lucide-react';

// --- Configuración Inicial (Default) ---
const DEFAULT_BOTS: BotDefinition[] = [
  {
    id: 'MIGO',
    name: 'Ingesta Masiva (MIGO)',
    description: 'Motor de automatización para movimientos de mercancía masivos en SAP.',
    icon: 'FileSpreadsheet',
    requiresFile: true,
    fileType: 'excel',
    supportsOpenMode: true
  },
  {
    id: 'PALLET',
    name: 'Optimizador de Altura',
    description: 'Análisis de volumetría y sincronización de datos maestros (LX02).',
    icon: 'Bot',
    requiresFile: true,
    fileType: 'excel',
    supportsOpenMode: true
  },
  {
    id: 'TRANSPORTE',
    name: 'Monitor Logístico',
    description: 'Trazabilidad en tiempo real de transportes y rutas (VT11/VT03N).',
    icon: 'Truck',
    requiresFile: false,
    requiresParam: true,
    paramLabel: 'Rango de Fechas'
  },
  {
    id: 'VISION',
    name: 'Visión Operacional',
    description: 'Digitalización de pizarras de operaciones mediante IA.',
    icon: 'Eye',
    requiresFile: true,
    fileType: 'image',
    supportsOpenMode: true
  },
  {
    id: 'AUDITOR',
    name: 'Guardián de Stock',
    description: 'Algoritmo de detección de anomalías y auditoría de inventario.',
    icon: 'ShieldCheck',
    requiresFile: false,
    requiresParam: true,
    paramLabel: 'Código de Almacén (Ej: SGVT)'
  },
  {
    id: 'LT01',
    name: 'Orquestador de Traspasos',
    description: 'Agente de transferencias automáticas ubicación a ubicación.',
    icon: 'ArrowRightLeft',
    requiresFile: true,
    fileType: 'excel',
    supportsOpenMode: true
  },
  {
    id: 'UMV',
    name: 'Sincronizador Maestros',
    description: 'Extracción y reconciliación de factores de conversión (UMV).',
    icon: 'Calculator',
    requiresFile: true,
    fileType: 'excel',
    supportsOpenMode: true
  },
  {
    id: 'CONCILIACION_EMAIL',
    name: 'Traspaso Perdida de Vacío',
    description: 'Automatización de traspasos por pérdida de vacío desde correos de Outlook.',
    icon: 'Mail',
    requiresFile: false
  },
  {
    id: 'ZONALES',
    name: 'Consolidación Zonales',
    description: 'Consolidador automático de reportes zonales (Faltantes, Sobrantes, Daño Mecánico).',
    icon: 'FolderSync',
    requiresFile: false
  },
  {
    id: 'ANALISIS_ZONALES',
    name: 'Analizador de Transporte',
    description: 'Auditoría avanzada de transportes: Detección de productos cambiados y cuadratura.',
    icon: 'BarChart3',
    requiresFile: false
  },
];

const IconMap: Record<string, React.ElementType> = {
  FileSpreadsheet, Truck, Eye, ShieldCheck, ArrowRightLeft, Calculator, Bot, Mail, FolderSync, BarChart3
};

interface AppConfig {
  userName: string;
  bots: BotDefinition[];
}

export default function App() {
  // --- Config State ---
  const [config, setConfig] = useState<AppConfig>({
    userName: 'Usuario SAP',
    bots: DEFAULT_BOTS
  });

  // Temp state for the settings modal
  const [tempConfig, setTempConfig] = useState<AppConfig | null>(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);

  // --- App Logic State ---
  const [selectedBotId, setSelectedBotId] = useState<string | null>(null);
  const [orders, setOrders] = useState<Order[]>([]);
  const [file, setFile] = useState<File | null>(null);
  const [paramValue, setParamValue] = useState('');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [stats, setStats] = useState({ total: 0, success: 0, pending: 0 });
  const [useOpenMode, setUseOpenMode] = useState(() => {
    const saved = localStorage.getItem('nexus_useOpenMode');
    return saved ? JSON.parse(saved) : false;
  });
  const [openFileName, setOpenFileName] = useState(() => {
    return localStorage.getItem('nexus_openFileName') || '';
  });
  const [sendEmail, setSendEmail] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('Enviando...');

  // Persist preferences
  useEffect(() => {
    localStorage.setItem('nexus_useOpenMode', JSON.stringify(useOpenMode));
    localStorage.setItem('nexus_openFileName', openFileName);
  }, [useOpenMode, openFileName]);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const terminalEndRef = useRef<HTMLDivElement>(null);

  // Helper to find current selected bot object from dynamic config
  const selectedBot = config.bots.find(b => b.id === selectedBotId) || null;

  // --- Load/Save Config ---
  useEffect(() => {
    const savedConfig = localStorage.getItem('nexus_config');
    if (savedConfig) {
      try {
        const parsed = JSON.parse(savedConfig);
        // Siempre usar DEFAULT_BOTS como fuente de verdad para los bots disponibles
        // Solo preservar nombres personalizados por el usuario
        const mergedBots = DEFAULT_BOTS.map(defaultBot => {
          const savedBot = parsed.bots?.find((b: BotDefinition) => b.id === defaultBot.id);
          if (savedBot && savedBot.name !== defaultBot.name) {
            // Preservar nombre personalizado
            return { ...defaultBot, name: savedBot.name };
          }
          return defaultBot;
        });
        setConfig({
          userName: parsed.userName || 'Usuario SAP',
          bots: mergedBots
        });
      } catch (e) {
        console.error("Error loading config", e);
      }
    }
  }, []);

  const saveConfig = (newConfig: AppConfig) => {
    setConfig(newConfig);
    localStorage.setItem('nexus_config', JSON.stringify(newConfig));
    setIsSettingsOpen(false);
  };

  const restoreDefaults = () => {
    const defaultConfig = {
      userName: 'Usuario SAP',
      bots: DEFAULT_BOTS
    };
    setTempConfig(defaultConfig);
  };

  const openSettings = () => {
    setTempConfig(JSON.parse(JSON.stringify(config))); // Deep copy
    setIsSettingsOpen(true);
  };

  // --- Firestore Logic ---
  useEffect(() => {
    const q = query(
      collection(db, 'ordenes_bot'),
      orderBy('fecha_creacion', 'desc'),
      limit(500)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const newOrders: Order[] = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      } as Order));

      setOrders(newOrders);
      const total = newOrders.length;
      const success = newOrders.filter(o => o.status === 'success').length;
      const pending = newOrders.filter(o => o.status === 'pending' || o.status === 'running').length;
      setStats({ total, success, pending });
    });
    return () => unsubscribe();
  }, []);

  const handleBotClick = (botId: string) => {
    setSelectedBotId(botId);
    setFile(null);
    setParamValue('');
    setStartDate('');
    setEndDate('');
    setIsDragging(false);
    // Activar modo abierto por defecto para VISION y PALLET
    setUseOpenMode(botId === 'VISION' || botId === 'PALLET');

    // No limpiar el nombre del archivo para "recordar" el último usado
    // setOpenFileName('');
    setSendEmail(false);
  };

  // --- Drag & Drop ---
  const handleDragOver = (e: React.DragEvent) => { e.preventDefault(); setIsDragging(true); };
  const handleDragLeave = (e: React.DragEvent) => { e.preventDefault(); setIsDragging(false); };
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0 && selectedBot) {
      const droppedFile = e.dataTransfer.files[0];
      if (selectedBot.fileType === 'excel' && !droppedFile.name.match(/\.(xls|xlsx|xlsm)$/)) {
        alert("Por favor ingrese un archivo Excel válido."); return;
      }
      if (selectedBot.fileType === 'image' && !droppedFile.type.startsWith('image/')) {
        alert("Por favor ingrese un archivo de imagen válido."); return;
      }
      setFile(droppedFile);
    }
  };

  const handleSubmit = async () => {
    if (!selectedBot) return;

    // Validación condicional según modo
    if (selectedBot.requiresFile) {
      if (useOpenMode) {
        if (!openFileName) { alert("Debe ingresar el nombre del archivo."); return; }
      } else {
        if (!file) { alert("Debe cargar un archivo para continuar."); return; }
      }
    }

    if (selectedBot.id === 'AUDITOR' && !paramValue) { alert("Debe ingresar el código de almacén."); return; }
    if (selectedBot.id === 'TRANSPORTE' && (!startDate || !endDate)) { alert("Debe seleccionar ambas fechas."); return; }

    setIsSubmitting(true);
    setSubmitStatus('Iniciando...'); // Nuevo estado para mostrar progreso

    try {
      let downloadURL = '';
      let originalName = '';

      if (selectedBot.requiresFile) {
        if (useOpenMode) {
          // Modo Archivo Abierto
          originalName = openFileName;
          if (selectedBot.fileType === 'excel' && !originalName.match(/\.(xls|xlsx|xlsm)$/i)) {
            originalName += '.xlsx';
          }
        } else if (file) {
          // Modo Upload Normal
          setSubmitStatus('Subiendo archivo...');
          // Sanear nombre de archivo
          const sanitizedName = file.name.replace(/[^a-zA-Z0-9._-]/g, '_');
          const storageRef = ref(storage, `uploads/${Date.now()}_${sanitizedName}`);

          const snapshot = await uploadBytes(storageRef, file);
          setSubmitStatus('Obteniendo URL...');
          downloadURL = await getDownloadURL(snapshot.ref);
          originalName = file.name;
        }
      }

      setSubmitStatus('Creando orden...');
      const parametros: Record<string, any> = {};
      if (selectedBot.id === 'AUDITOR') parametros.almacen = paramValue;
      if (selectedBot.id === 'TRANSPORTE') {
        const format = (d: string) => { if (!d) return ''; const [y, m, dDay] = d.split('-'); return `${dDay}.${m}.${y}`; };
        parametros.fechas = `${format(startDate)}-${format(endDate)}`;
        parametros.sendEmail = sendEmail;
      }

      const newOrder = {
        tipo_bot: selectedBot.id,
        status: 'pending',
        fecha_creacion: serverTimestamp(),
        worker: 'En Cola',
        ruta_archivo: downloadURL, // Será vacío si es modo abierto
        nombre_archivo_original: originalName,
        parametros: parametros,
        execution_logs: []
      };

      await addDoc(collection(db, 'ordenes_bot'), newOrder);

      // Reset UI after order is successfully created
      setFile(null);
      setParamValue('');
      setStartDate('');
      setEndDate('');
      setOpenFileName('');
      setSendEmail(false);
      setIsSubmitting(false);
      setSubmitStatus('Enviando...');

      alert("✅ Orden enviada exitosamente.\n\nPuedes ver el progreso en la Consola de Eventos abajo.");
    } catch (error) {
      console.error("Error creating order:", error);
      setIsSubmitting(false);
      setSubmitStatus('Enviando...');
      alert(`❌ Error al enviar la orden:\n${error instanceof Error ? error.message : 'Error desconocido'}`);
    }
  };

  const handleCancelOrder = async (orderId: string) => {
    if (!confirm("¿Estás seguro de que deseas cancelar esta orden?")) return;
    try {
      await updateDoc(doc(db, 'ordenes_bot', orderId), {
        status: 'cancelled',
        execution_logs: [...(orders.find(o => o.id === orderId)?.execution_logs || []), "⚠️ Cancelado por el usuario"]
      });
    } catch (error) {
      console.error("Error cancelling order:", error);
      alert("Error al cancelar la orden.");
    }
  };

  return (
    <div className="h-full flex bg-dark-bg text-slate-200 font-sans overflow-hidden">

      {/* --- SIDEBAR --- */}
      <aside className="w-64 bg-dark-panel border-r border-white/5 flex flex-col flex-shrink-0 z-20 shadow-2xl">
        <div className="h-14 flex items-center gap-2 px-4 border-b border-white/5 bg-gradient-to-r from-nexus-900/20 to-transparent">
          <div className="p-1.5 bg-nexus-500/10 rounded-lg border border-nexus-500/20">
            <Cpu className="w-4 h-4 text-nexus-400" />
          </div>
          <div>
            <h1 className="font-bold text-white tracking-tight text-xs">Nexus Jarvis Automation Suite</h1>
            <span className="text-[9px] text-slate-500 font-medium tracking-wider">SUITE LOGÍSTICA</span>
          </div>
        </div>

        <div className="flex-1 overflow-y-auto custom-scrollbar p-3 space-y-1.5">
          <p className="text-[9px] uppercase font-bold text-slate-500 mb-2 px-1.5 tracking-widest">Catálogo</p>
          {config.bots.map((bot) => {
            const Icon = IconMap[bot.icon];
            const isActive = selectedBotId === bot.id;
            return (
              <button
                key={bot.id}
                onClick={() => handleBotClick(bot.id)}
                className={`w-full group relative flex items-center gap-2 p-2 rounded-lg transition-all duration-300 text-left border ${isActive
                  ? 'bg-nexus-600 text-white border-nexus-500 shadow-lg shadow-nexus-900/50'
                  : 'bg-transparent text-slate-400 border-transparent hover:bg-white/5 hover:text-slate-200'
                  }`}
              >
                <Icon className={`w-4 h-4 ${isActive ? 'text-white' : 'text-slate-500 group-hover:text-nexus-400'}`} />
                <div className="flex-1 min-w-0">
                  <span className="block text-xs font-semibold truncate">{bot.name}</span>
                </div>
                {isActive && <ChevronRight className="w-3 h-3 text-white/50" />}
              </button>
            );
          })}
        </div>

        {/* Sidebar Footer: Settings & Stats */}
        <div className="p-3 border-t border-white/5 bg-black/20 space-y-2">
          <div className="flex gap-2">
            <button
              onClick={openSettings}
              className="flex-1 flex items-center justify-center gap-2 p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-white/5 transition-colors text-xs font-medium border border-transparent hover:border-white/5"
              title="Configuración"
            >
              <Settings className="w-3 h-3" />
              Config
            </button>
            <button
              onClick={async () => {
                if (!confirm("¿Reiniciar todo el sistema? Esto cerrará los bots y recargará la interfaz.")) return;
                try {
                  await addDoc(collection(db, 'ordenes_bot'), {
                    tipo_bot: 'SYSTEM_RESTART',
                    status: 'pending',
                    fecha_creacion: serverTimestamp(),
                    worker: 'En Cola',
                    parametros: {}
                  });
                  alert("🔄 Reinicio solicitado. El sistema se cerrará y volverá a abrir en unos segundos.");
                } catch (e) {
                  alert("Error solicitando reinicio");
                }
              }}
              className="flex-1 flex items-center justify-center gap-2 p-1.5 rounded-lg text-amber-500/80 hover:text-amber-400 hover:bg-amber-500/10 transition-colors text-xs font-medium border border-amber-500/20"
              title="Reiniciar Sistema"
            >
              <RotateCcw className="w-3 h-3" />
              Reiniciar
            </button>
          </div>

          <div className="grid grid-cols-2 gap-1.5">
            <div className="bg-white/5 rounded-lg p-1.5 border border-white/5">
              <span className="block text-[9px] text-slate-500 uppercase">Activas</span>
              <span className="block text-base font-bold text-emerald-400">{stats.pending}</span>
            </div>
            <div className="bg-white/5 rounded-lg p-1.5 border border-white/5">
              <span className="block text-[9px] text-slate-500 uppercase">Total</span>
              <span className="block text-base font-bold text-slate-300">{stats.total}</span>
            </div>
          </div>
        </div>
      </aside>

      {/* --- MAIN AREA --- */}
      <main className="flex-1 flex flex-col min-w-0 bg-[radial-gradient(ellipse_at_top_right,_var(--tw-gradient-stops))] from-slate-900 via-dark-bg to-black relative">

        {/* Header */}
        <div className="h-12 border-b border-white/5 flex items-center justify-between px-4 bg-dark-bg/50 backdrop-blur-md z-10">
          <div className="flex items-center gap-1.5">
            <Activity className="w-3 h-3 text-emerald-500 animate-pulse" />
            <span className="text-[10px] font-mono text-emerald-400">SISTEMA CONECTADO</span>
          </div>
          <div className="flex items-center gap-2">
            <div className="text-right">
              <span className="block text-[9px] text-slate-500 uppercase tracking-wider">Bienvenido,</span>
              <span className="block text-xs font-bold text-white tracking-wide">{config.userName}</span>
            </div>
            <div className="w-7 h-7 rounded-full bg-slate-800 flex items-center justify-center border border-white/10 shadow-lg shadow-black/50">
              <User className="w-4 h-4 text-slate-400" />
            </div>
          </div>
        </div>

        {/* Workspace */}
        <div className="flex-1 p-4 overflow-y-auto custom-scrollbar">
          {!selectedBot ? (
            <div className="h-full flex flex-col items-center justify-center text-center space-y-4 opacity-40">
              <div className="p-4 rounded-full bg-white/5 border border-white/10">
                <LayoutGrid className="w-12 h-12 text-slate-500" />
              </div>
              <div>
                <h2 className="text-xl font-bold text-white mb-1">Espacio de Trabajo Vacío</h2>
                <p className="text-sm text-slate-400 max-w-md mx-auto">Seleccione un módulo de automatización para comenzar.</p>
              </div>
            </div>
          ) : (
            <div className="max-w-5xl mx-auto space-y-3 animate-in fade-in slide-in-from-bottom-4 duration-500">
              <div className="flex items-start gap-3">
                <div className="p-2 rounded-xl bg-nexus-500/20 border border-nexus-500/30 text-nexus-400 shadow-lg shadow-nexus-900/20">
                  {React.createElement(IconMap[selectedBot.icon], { className: "w-6 h-6" })}
                </div>
                <div>
                  <h2 className="text-xl font-bold text-white mb-0.5">{selectedBot.name}</h2>
                  <p className="text-sm text-slate-400 leading-snug">{selectedBot.description}</p>
                </div>
              </div>
              <div className="h-px bg-white/10 w-full"></div>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                <div className="space-y-3">
                  {selectedBot.requiresFile && (
                    <div className="space-y-2">
                      <div className="flex items-center justify-between">
                        <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest flex items-center gap-1.5">
                          <FileText className="w-3 h-3" />
                          Archivo Fuente ({selectedBot.fileType === 'excel' ? 'Excel' : 'Imagen'})
                        </label>
                        {selectedBot.supportsOpenMode && (
                          <div className="flex items-center gap-2">
                            <span className="text-[10px] text-slate-500">
                              {selectedBot.fileType === 'image' ? 'Usar Archivo Local' : 'Usar Excel Abierto'}
                            </span>
                            <button
                              onClick={() => setUseOpenMode(!useOpenMode)}
                              className={`w-8 h-4 rounded-full transition-colors relative ${useOpenMode ? 'bg-nexus-500' : 'bg-slate-700'}`}
                            >
                              <div className={`absolute top-0.5 left-0.5 w-3 h-3 bg-white rounded-full transition-transform ${useOpenMode ? 'translate-x-4' : 'translate-x-0'}`} />
                            </button>
                          </div>
                        )}
                      </div>

                      {useOpenMode ? (
                        <div className="bg-white/5 border border-white/10 rounded-xl p-4 space-y-2 animate-in fade-in">
                          <label className="text-xs text-slate-400 block">
                            {selectedBot.fileType === 'image' ? 'Nombre del archivo de imagen:' : 'Nombre del archivo Excel abierto:'}
                          </label>
                          <input
                            type="text"
                            value={openFileName}
                            onChange={(e) => setOpenFileName(e.target.value)}
                            placeholder={selectedBot.fileType === 'image' ? "Ej: pizarra_semana_48.jpg" : "Ej: carga_migo.xlsx"}
                            className="w-full bg-dark-bg border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-nexus-500/50 text-sm"
                          />
                          <p className="text-[10px] text-amber-500/80 flex items-center gap-1">
                            <Zap className="w-3 h-3" />
                            {selectedBot.fileType === 'image'
                              ? "El bot buscará este archivo en la carpeta de OneDrive."
                              : "El bot buscará este archivo en tus libros de Excel abiertos."}
                          </p>
                        </div>
                      ) : (
                        <div
                          onDragOver={handleDragOver} onDragLeave={handleDragLeave} onDrop={handleDrop}
                          onClick={() => fileInputRef.current?.click()}
                          className={`relative group h-32 border-2 border-dashed rounded-xl flex flex-col items-center justify-center transition-all cursor-pointer overflow-hidden ${isDragging ? 'border-nexus-400 bg-nexus-500/10' : file ? 'border-emerald-500/50 bg-emerald-500/5' : 'border-slate-700 hover:border-nexus-500/50 hover:bg-white/[0.02]'}`}
                        >
                          <input
                            type="file"
                            ref={fileInputRef}
                            className="hidden"
                            accept={selectedBot.fileType === 'image' ? "image/*" : ".xlsx,.xls,.xlsm"}
                            onChange={(e) => {
                              const selectedFile = e.target.files?.[0];
                              if (selectedFile) {
                                setFile(selectedFile);
                              }
                            }}
                          />
                          {file ? (
                            <div className="text-center z-10 animate-in zoom-in duration-300">
                              <div className="w-8 h-8 rounded-full bg-emerald-500/20 flex items-center justify-center mx-auto mb-1.5">
                                <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                              </div>
                              <p className="font-semibold text-sm text-emerald-300 truncate max-w-[200px] px-2">{file.name}</p>
                              <p className="text-[10px] text-emerald-500/70 mt-0.5">Listo para procesar</p>
                              <button onClick={(e) => { e.stopPropagation(); setFile(null); }} className="mt-1.5 text-[9px] text-slate-400 hover:text-white underline">Cambiar archivo</button>
                            </div>
                          ) : (
                            <div className="text-center z-10 pointer-events-none">
                              <div className={`w-8 h-8 rounded-full flex items-center justify-center mx-auto mb-1.5 transition-colors ${isDragging ? 'bg-nexus-500 text-white' : 'bg-slate-800 text-slate-400 group-hover:bg-nexus-500/20 group-hover:text-nexus-400'}`}>
                                <UploadCloud className="w-4 h-4" />
                              </div>
                              <p className="text-sm font-medium text-slate-300 group-hover:text-white transition-colors">{isDragging ? "Suelta el archivo aquí" : "Arrastra tu archivo aquí"}</p>
                              <p className="text-[10px] text-slate-500 mt-0.5">o haz clic para explorar</p>
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  )}
                  {selectedBot.requiresParam && selectedBot.id === 'AUDITOR' && (
                    <div className="space-y-2">
                      <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest flex items-center gap-1.5"><Search className="w-3 h-3" />{selectedBot.paramLabel}</label>
                      <div className="relative">
                        <select
                          value={paramValue}
                          onChange={(e) => setParamValue(e.target.value)}
                          className="w-full bg-dark-panel border border-slate-700 rounded-xl px-3 py-3 text-white text-base font-mono focus:outline-none focus:ring-2 focus:ring-nexus-500/50 transition-all appearance-none cursor-pointer"
                        >
                          <option value="" disabled className="bg-slate-800 text-slate-400">Seleccione Almacén</option>
                          <option value="SGVT" className="bg-slate-800 text-white">SGVT</option>
                          <option value="SDIF" className="bg-slate-800 text-white">SDIF</option>
                          <option value="SGTR" className="bg-slate-800 text-white">SGTR</option>
                          <option value="SGSD" className="bg-slate-800 text-white">SGSD</option>
                          <option value="CDNW" className="bg-slate-800 text-white">CDNW</option>
                          <option value="SGEN" className="bg-slate-800 text-white">SGEN</option>
                          <option value="AVAS" className="bg-slate-800 text-white">AVAS</option>
                          <option value="SGBC" className="bg-slate-800 text-white">SGBC</option>
                          <option value="TAVI" className="bg-slate-800 text-white">TAVI</option>
                        </select>
                        <ChevronRight className="w-4 h-4 text-slate-500 absolute right-3 top-1/2 transform -translate-y-1/2 rotate-90 pointer-events-none" />
                      </div>
                    </div>
                  )}
                  {selectedBot.id === 'TRANSPORTE' && (
                    <div className="space-y-3">
                      <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest flex items-center gap-1.5"><Calendar className="w-3 h-3" />Intervalo de Análisis</label>
                      <div className="grid grid-cols-2 gap-3">
                        <div className="space-y-1"><span className="text-[9px] text-slate-500 ml-1 font-bold">DESDE</span><input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} className="w-full bg-dark-panel border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-nexus-500/50 transition-all text-xs" /></div>
                        <div className="space-y-1"><span className="text-[9px] text-slate-500 ml-1 font-bold">HASTA</span><input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} className="w-full bg-dark-panel border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-nexus-500/50 transition-all text-xs" /></div>
                      </div>
                      <div className="flex items-center gap-2 bg-white/5 border border-white/10 rounded-lg p-3">
                        <input
                          type="checkbox"
                          id="sendEmail"
                          checked={sendEmail}
                          onChange={(e) => setSendEmail(e.target.checked)}
                          className="w-4 h-4 rounded border-slate-700 bg-dark-panel text-nexus-500 focus:ring-2 focus:ring-nexus-500/50 cursor-pointer"
                        />
                        <label htmlFor="sendEmail" className="text-sm text-slate-300 cursor-pointer select-none">Enviar correo al finalizar</label>
                      </div>
                    </div>
                  )}
                </div>
                <div className="flex flex-col justify-between space-y-3">
                  <div className="bg-white/5 border border-white/5 rounded-xl p-3">
                    <h4 className="text-xs font-bold text-white mb-2">Resumen de Ejecución</h4>
                    <ul className="space-y-2 text-xs text-slate-400">
                      <li className="flex justify-between"><span>Bot:</span><span className="text-nexus-300 font-medium truncate ml-2">{selectedBot.name}</span></li>
                      <li className="flex justify-between"><span>Estado Worker:</span><span className="text-emerald-400 font-medium flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>Online</span></li>
                    </ul>
                  </div>
                  <button onClick={handleSubmit} disabled={isSubmitting} className={`w-full py-3.5 px-5 rounded-xl font-bold text-base text-white flex items-center justify-center gap-2.5 transition-all transform active:scale-[0.98] shadow-xl ${isSubmitting ? 'bg-slate-700 cursor-not-allowed grayscale' : 'bg-gradient-to-r from-nexus-600 to-indigo-600 hover:from-nexus-500 hover:to-indigo-500 shadow-nexus-600/20'}`}>
                    {isSubmitting ? <><Loader2 className="w-5 h-5 animate-spin" />{submitStatus}</> : <><Zap className="w-5 h-5 fill-current" />EJECUTAR FLUJO</>}
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Terminal */}
        <div className="h-36 border-t border-white/10 bg-[#0d1117] flex flex-col flex-shrink-0 z-20 shadow-[0_-5px_15px_rgba(0,0,0,0.5)]">
          <div className="h-8 flex items-center justify-between px-3 border-b border-white/5 bg-white/[0.02]">
            <div className="flex items-center gap-1.5"><Terminal className="w-3 h-3 text-emerald-500" /><span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Consola de Eventos</span></div>
            <div className="flex items-center gap-1.5"><div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></div><span className="text-[9px] text-emerald-500 font-mono">LIVE</span></div>
          </div>
          <div className="flex-1 overflow-y-auto p-2 space-y-0.5 font-mono text-xs custom-scrollbar">
            {orders.length === 0 ? <div className="text-slate-700 text-center mt-6 italic text-xs">-- Sin eventos --</div> : orders.map((order) => (
              <div key={order.id} className="mb-1">
                <div className="flex gap-3 hover:bg-white/[0.02] p-0.5 rounded transition-colors group">
                  <span className="text-slate-600 w-16 flex-shrink-0 text-[10px] pt-0.5">{order.fecha_creacion?.toDate().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' })}</span>
                  <div className="flex-1">
                    <div className="flex items-center gap-1.5">
                      <StatusIndicator status={order.status} />
                      <span className={`font-semibold text-[11px] ${order.status === 'error' ? 'text-red-400' : 'text-slate-300'}`}>[{config.bots.find(b => b.id === order.tipo_bot)?.name || order.tipo_bot}]</span>
                      {(order.status === 'pending' || order.status === 'running') && (
                        <button
                          onClick={(e) => { e.stopPropagation(); handleCancelOrder(order.id); }}
                          className="ml-2 px-1.5 py-0.5 bg-red-500/20 hover:bg-red-500/40 text-red-400 text-[9px] rounded border border-red-500/30 transition-colors"
                          title="Cancelar Orden"
                        >
                          CANCELAR
                        </button>
                      )}
                    </div>
                    {order.execution_logs && order.execution_logs.length > 0 && (
                      <div className="ml-4 mt-1 bg-black/40 rounded p-1.5 border-l-2 border-emerald-500/50">
                        {order.execution_logs.slice(-2).map((log, i) => <div key={i} className="text-[10px] text-emerald-400/90 font-mono whitespace-pre-wrap">$ {log}</div>)}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            ))}
            <div ref={terminalEndRef} />
          </div>
        </div>
      </main>

      {/* --- SETTINGS MODAL --- */}
      {isSettingsOpen && tempConfig && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="w-full max-w-2xl bg-dark-surface border border-white/10 rounded-2xl shadow-2xl flex flex-col max-h-[90vh]">

            {/* Modal Header */}
            <div className="flex items-center justify-between p-6 border-b border-white/5 bg-white/[0.02]">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-nexus-500/20 rounded-lg text-nexus-400"><Settings className="w-6 h-6" /></div>
                <div>
                  <h3 className="text-xl font-bold text-white">Configuración del Sistema</h3>
                  <p className="text-sm text-slate-400">Personaliza la apariencia y nombres de la suite.</p>
                </div>
              </div>
              <button onClick={() => setIsSettingsOpen(false)} className="text-slate-500 hover:text-white transition-colors"><X className="w-6 h-6" /></button>
            </div>

            {/* Modal Content */}
            <div className="flex-1 overflow-y-auto p-6 space-y-8 custom-scrollbar">

              {/* Sección Perfil */}
              <div className="space-y-4">
                <h4 className="text-sm font-bold text-emerald-400 uppercase tracking-widest border-b border-white/5 pb-2">Perfil de Usuario</h4>
                <div className="grid grid-cols-1 gap-4">
                  <div>
                    <label className="block text-xs font-medium text-slate-400 mb-1">Nombre para Mostrar</label>
                    <input
                      type="text"
                      value={tempConfig.userName}
                      onChange={(e) => setTempConfig({ ...tempConfig, userName: e.target.value })}
                      className="w-full bg-dark-bg border border-slate-700 rounded-lg px-4 py-2 text-white focus:border-nexus-500 focus:outline-none"
                    />
                  </div>
                </div>
              </div>

              {/* Sección Bots */}
              <div className="space-y-4">
                <div className="flex items-center justify-between border-b border-white/5 pb-2">
                  <h4 className="text-sm font-bold text-nexus-400 uppercase tracking-widest">Personalización de Módulos</h4>
                  <button onClick={restoreDefaults} className="text-xs text-nexus-400 hover:text-nexus-300 flex items-center gap-1"><RotateCcw className="w-3 h-3" /> Restaurar Nombres</button>
                </div>

                <div className="space-y-3">
                  {tempConfig.bots.map((bot, idx) => (
                    <div key={bot.id} className="flex items-center gap-4 bg-white/[0.02] p-3 rounded-lg border border-white/5">
                      <div className="p-2 bg-slate-800 rounded text-slate-400">
                        {React.createElement(IconMap[bot.icon], { className: "w-4 h-4" })}
                      </div>
                      <div className="flex-1">
                        <label className="text-[10px] text-slate-500 uppercase font-bold block mb-1">Nombre del Módulo (ID: {bot.id})</label>
                        <input
                          type="text"
                          value={bot.name}
                          onChange={(e) => {
                            const newBots = [...tempConfig.bots];
                            newBots[idx] = { ...newBots[idx], name: e.target.value };
                            setTempConfig({ ...tempConfig, bots: newBots });
                          }}
                          className="w-full bg-transparent border-b border-slate-700 text-sm text-white focus:border-nexus-500 focus:outline-none pb-1"
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* Modal Footer */}
            <div className="p-6 border-t border-white/5 bg-black/20 flex justify-end gap-3 rounded-b-2xl">
              <button
                onClick={() => setIsSettingsOpen(false)}
                className="px-4 py-2 rounded-lg text-slate-400 hover:text-white hover:bg-white/5 transition-colors text-sm font-medium"
              >
                Cancelar
              </button>
              <button
                onClick={() => tempConfig && saveConfig(tempConfig)}
                className="px-6 py-2 rounded-lg bg-nexus-600 hover:bg-nexus-500 text-white font-bold text-sm shadow-lg shadow-nexus-900/20 flex items-center gap-2"
              >
                <Save className="w-4 h-4" />
                Guardar Cambios
              </button>
            </div>

          </div>
        </div>
      )}
    </div>
  );
}

const StatusIndicator = ({ status }: { status: string }) => {
  switch (status) {
    case 'pending': return <span className="text-amber-500 font-bold text-xs">[EN COLA]</span>;
    case 'running': return <span className="text-blue-500 font-bold text-xs animate-pulse">[PROCESANDO]</span>;
    case 'success': return <span className="text-emerald-500 font-bold text-xs">[EXITOSO]</span>;
    case 'error': return <span className="text-red-500 font-bold text-xs">[FALLIDO]</span>;
    default: return <span className="text-slate-500 text-xs">[UNK]</span>;
  }
};