import tree from './categorizacion.tree.json';
import plantillas from './sr_plantillas.json';
import { TipoTicket } from './types';

// ============================================================================
//  CATÁLOGO DE CATEGORIZACIÓN (fuente: Categorizacion.xlsx)
// ----------------------------------------------------------------------------
//  Jerarquía Servicio -> Categoría -> Subcategoría. Cada fila del Excel indica
//  si corresponde a un Incidente (IN) o a una Solicitud (SR).
//
//  Este módulo:
//   - arma un resumen compacto del catálogo para inyectarlo al prompt del LLM,
//   - normaliza el servicio/categoría que elige el LLM (sin acentos/mayúsculas),
//   - resuelve la plantilla de Solicitud (SvcReqTmplLink) por servicio.
// ============================================================================

interface CatCategoria {
  nombre: string;
  descripcion: string;
  tipos: string[]; // ['IN'], ['SR'] o ambos
  subcategorias: { nombre: string; tipo: string }[];
}
interface CatServicio {
  nombre: string;
  palabrasClave: string;
  categorias: CatCategoria[];
}

const SERVICIOS = (tree as { servicios: CatServicio[] }).servicios;
const PLANTILLAS = plantillas as {
  porServicio: Record<string, { recId: string; plantilla: string }>;
};

/** Normaliza para comparar: minúsculas, sin acentos, sin espacios extra. */
function norm(s: string | undefined): string {
  return (s ?? '')
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .toLowerCase()
    .trim();
}

/**
 * Resumen del catálogo para el prompt del LLM. Formato compacto:
 *   ## Servicio  (palabras clave: ...)
 *   - Categoría [IN|SR]
 * Ronda ~800 tokens; suficiente para que el modelo elija servicio + categoría.
 */
let promptCache: string | null = null;
export function catalogoParaPrompt(): string {
  if (promptCache) return promptCache;
  const bloques = SERVICIOS.map((s) => {
    const pk = s.palabrasClave ? `  (palabras clave: ${s.palabrasClave.replace(/\s+/g, ' ').slice(0, 160)})` : '';
    const cats = s.categorias
      .map((c) => `  - ${c.nombre} [${c.tipos.join('/') || '—'}]`)
      .join('\n');
    return `## ${s.nombre}${pk}\n${cats}`;
  });
  promptCache = bloques.join('\n');
  return promptCache;
}

export interface Clasificacion {
  servicio?: string; // nombre canónico
  categoria?: string; // nombre canónico
  tiposCategoria: TipoTicket[]; // qué permite el catálogo para esa categoría
}

const TIPO_MAP: Record<string, TipoTicket> = { IN: 'incidente', SR: 'solicitud' };

/**
 * Normaliza el servicio/categoría elegidos por el LLM a los nombres canónicos
 * del catálogo y devuelve qué tipos (incidente/solicitud) admite esa categoría.
 */
export function clasificar(servicio?: string, categoria?: string): Clasificacion {
  const svc = SERVICIOS.find((s) => norm(s.nombre) === norm(servicio));
  if (!svc) return { tiposCategoria: [] };
  const cat = svc.categorias.find((c) => norm(c.nombre) === norm(categoria));
  const tipos = (cat?.tipos ?? [])
    .map((t) => TIPO_MAP[t])
    .filter((t): t is TipoTicket => Boolean(t));
  return { servicio: svc.nombre, categoria: cat?.nombre, tiposCategoria: tipos };
}

/** RecId de la plantilla de Solicitud (SvcReqTmplLink) para un servicio. */
export function plantillaSR(servicio?: string): string | undefined {
  if (!servicio) return undefined;
  // match exacto y luego insensible a acentos/mayúsculas
  const directo = PLANTILLAS.porServicio[servicio];
  if (directo) return directo.recId;
  const key = Object.keys(PLANTILLAS.porServicio).find((k) => norm(k) === norm(servicio));
  return key ? PLANTILLAS.porServicio[key].recId : undefined;
}
