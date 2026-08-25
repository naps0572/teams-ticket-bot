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
//   - resuelve la plantilla de Solicitud (SvcReqTmplLink) por oferta/servicio.
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
  plantillas: Array<{ recId: string; nombre: string; servicio: string }>;
};

export interface OfertaServicio {
  recId: string;
  nombre: string;
  servicio: string;
}

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

/** Ofertas de servicio publicadas que el modelo puede seleccionar para un SR. */
export function ofertasParaPrompt(): string {
  const agrupadas = new Map<string, string[]>();
  for (const oferta of PLANTILLAS.plantillas) {
    if (!oferta.servicio) continue;
    const nombres = agrupadas.get(oferta.servicio) ?? [];
    if (!nombres.some((nombre) => norm(nombre) === norm(oferta.nombre))) nombres.push(oferta.nombre);
    agrupadas.set(oferta.servicio, nombres);
  }
  return [...agrupadas.entries()]
    .map(([servicio, ofertas]) => `- ${servicio}: ${ofertas.join(' | ')}`)
    .join('\n');
}

/** Ofertas publicadas asociadas a un servicio, sin duplicados. */
export function ofertasSR(servicio?: string): OfertaServicio[] {
  if (!servicio) return [];
  return PLANTILLAS.plantillas.filter(
    (oferta, index, todas) =>
      norm(oferta.servicio) === norm(servicio) &&
      todas.findIndex(
        (item) =>
          norm(item.servicio) === norm(oferta.servicio) && norm(item.nombre) === norm(oferta.nombre),
      ) === index,
  );
}

/** Devuelve la oferta canónica seleccionada dentro del servicio. */
export function ofertaSR(servicio?: string, oferta?: string): OfertaServicio | undefined {
  if (!oferta) return undefined;
  return ofertasSR(servicio).find((item) => norm(item.nombre) === norm(oferta));
}

/** Subcategorías de Solicitud válidas para un servicio/categoría. */
export function subcategoriasSR(servicio?: string, categoria?: string): string[] {
  const svc = SERVICIOS.find((item) => norm(item.nombre) === norm(servicio));
  const cat = svc?.categorias.find((item) => norm(item.nombre) === norm(categoria));
  if (!cat) return [];
  return cat.subcategorias
    .filter((item) => item.tipo === 'SR')
    .map((item) => item.nombre)
    .filter((nombre, index, todas) => todas.findIndex((item) => norm(item) === norm(nombre)) === index);
}

/** Devuelve el nombre canónico de una subcategoría de Solicitud. */
export function subcategoriaSR(
  servicio?: string,
  categoria?: string,
  subcategoria?: string,
): string | undefined {
  if (!subcategoria) return undefined;
  return subcategoriasSR(servicio, categoria).find((item) => norm(item) === norm(subcategoria));
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

/** RecId de la plantilla de Solicitud (SvcReqTmplLink) para una oferta. */
export function plantillaSR(servicio?: string, oferta?: string): string | undefined {
  if (!servicio) return undefined;
  if (oferta) {
    return ofertaSR(servicio, oferta)?.recId;
  }
  // match exacto y luego insensible a acentos/mayúsculas
  const directo = PLANTILLAS.porServicio[servicio];
  if (directo) return directo.recId;
  const key = Object.keys(PLANTILLAS.porServicio).find((k) => norm(k) === norm(servicio));
  return key ? PLANTILLAS.porServicio[key].recId : undefined;
}
