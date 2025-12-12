
export type BotType = 'MIGO' | 'PALLET' | 'TRANSPORTE' | 'AUDITOR' | 'LT01' | 'UMV' | 'VISION' | 'CONCILIACION_EMAIL' | 'ZONALES' | 'ANALISIS_ZONALES' | 'SYSTEM_RESTART';

export interface BotDefinition {
  id: BotType;
  name: string;
  description: string;
  icon: string;
  requiresFile: boolean;
  fileType?: 'excel' | 'image';
  requiresParam?: boolean;
  paramLabel?: string;
  supportsOpenMode?: boolean;
}

export interface Order {
  id: string;
  tipo_bot: BotType;
  status: 'pending' | 'running' | 'success' | 'error';
  fecha_creacion: any; // Firestore Timestamp
  worker?: string;
  mensaje?: string;
  error?: string;
  ruta_archivo?: string; // URL in Storage
  nombre_archivo_original?: string;
  parametros?: {
    almacen?: string;
    fechas?: string;
    sendEmail?: boolean;
  };
  execution_logs?: string[]; // Array of log strings from the Python worker
}
