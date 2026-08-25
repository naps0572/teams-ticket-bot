// Tipos compartidos en todo el bot.

/** Prioridades soportadas para un ticket. Ajusta a los valores que use tu ITSM. */
export type Prioridad = 'baja' | 'media' | 'alta' | 'critica';

/** Tipo de ticket según el catálogo de categorización (columnas IN / SR). */
export type TipoTicket = 'incidente' | 'solicitud';

/** Borrador de ticket que el modelo va completando durante la conversación. */
export interface DraftTicket {
  asunto: string;
  descripcion: string;
  /** "incidente" (algo falla) o "solicitud" (una petición). */
  tipo?: TipoTicket;
  /** Servicio del catálogo (Categorizacion.xlsx). */
  servicio?: string;
  /** Oferta de servicio de Ivanti; determina el SvcReqTmplLink del SR. */
  oferta?: string;
  /** Categoría del catálogo, dentro del servicio. */
  categoria?: string;
  /** Subcategoría opcional. */
  subcategoria?: string;
  prioridad?: Prioridad;
}

/** Datos de quien solicita el ticket (vienen de la identidad de Teams). */
export interface Requester {
  name?: string;
  email?: string;
  aadObjectId?: string;
}

/** Resultado de crear un ticket en el ITSM. */
export interface CreatedTicket {
  id: string;
  url?: string;
  /** Tipo con el que se creó (para el mensaje/tarjeta). */
  tipo?: TipoTicket;
}

/** Un turno de la conversación con el modelo. */
export interface ChatTurn {
  role: 'user' | 'assistant';
  content: string;
}

/** Lo que el modelo devuelve en cada turno (JSON estructurado). */
export interface AssistantResult {
  /** Texto que se le muestra al usuario. */
  respuesta: string;
  /** true cuando ya hay suficiente información para crear el ticket. */
  listo: boolean;
  /** Datos del ticket, presentes solo cuando `listo` es true. */
  ticket: DraftTicket | null;
}
