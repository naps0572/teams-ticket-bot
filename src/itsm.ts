import { clasificar, plantillaSR } from './categorias';
import { CreatedTicket, DraftTicket, Prioridad, Requester, TipoTicket } from './types';

// ============================================================================
//  CONECTOR DEL ITSM — Ivanti Service Manager (Neurons / ISM Cloud)
// ----------------------------------------------------------------------------
//  El resto del bot solo llama a `createTicket(...)`. Según el catálogo
//  (Categorizacion.xlsx) el ticket puede ser:
//    - Incidente  -> POST /businessobject/incidents
//    - Solicitud  -> POST /businessobject/servicereqs  (requiere plantilla)
//
//  Auth: header  Authorization: rest_api_key=<KEY>   (NO Bearer)
//
//  Campos obligatorios (verificado con sondas al API):
//    Incidente : Service, Category, Team(OwnerTeam), Customer(ProfileLink)
//    Solicitud : Service, Contact(ProfileLink), SvcReqTmplLink_RecID(plantilla)
//  Priority en incidentes lo calcula Ivanti desde Impact + Urgency.
//
//  MODO DRY-RUN: sin ITSM_BASE_URL devuelve un ticket simulado (LOCAL-xxxx).
//
//  Variables de entorno:
//    ITSM_BASE_URL, ITSM_API_KEY
//    ITSM_DEFAULT_CATEGORY / ITSM_DEFAULT_SERVICE / ITSM_DEFAULT_TEAM
//    ITSM_DEFAULT_PROFILE_LINK   Customer fallback si el correo no existe
//    ITSM_SR_TEMPLATE_RECID      plantilla SR por defecto si el servicio no tiene
// ============================================================================

type IvantiObject = 'incidents' | 'servicereqs';

/** Impact/Urgency de Ivanti a partir de la prioridad interna. */
const NIVEL_IVANTI: Record<Prioridad, string> = {
  baja: 'Low',
  media: 'Medium',
  alta: 'High',
  critica: 'Critical',
};

export async function createTicket(
  ticket: DraftTicket,
  requester: Requester,
): Promise<CreatedTicket> {
  const baseUrl = process.env.ITSM_BASE_URL;
  const tipo: TipoTicket = ticket.tipo === 'solicitud' ? 'solicitud' : 'incidente';

  // --- Modo dry-run (sin ITSM configurado) --------------------------------
  if (!baseUrl) {
    const fakeId = `LOCAL-${Date.now().toString().slice(-6)}`;
    // eslint-disable-next-line no-console
    console.warn(`[itsm] ITSM_BASE_URL no configurado: dry-run, ticket ${fakeId} (${tipo})`);
    return { id: fakeId, tipo };
  }

  const apiKey = process.env.ITSM_API_KEY;
  if (!apiKey) {
    throw new Error('[itsm] Falta ITSM_API_KEY para autenticar contra Ivanti.');
  }

  const root = baseUrl.replace(/\/$/, '');
  const objectType: IvantiObject = tipo === 'solicitud' ? 'servicereqs' : 'incidents';

  // Contacto/Customer: RecId del empleado resuelto por su correo.
  const profileLink = await resolveCustomer(root, apiKey, requester.email);

  const body =
    objectType === 'servicereqs'
      ? buildSolicitud(ticket, profileLink)
      : buildIncidente(ticket, profileLink);

  const data = await ivantiPost(root, apiKey, objectType, body);

  const id =
    asString(data.IncidentNumber) ??
    asString(data.ServiceReqNumber) ??
    asString(data.RecId) ??
    'desconocido';
  const url = buildTicketUrl(root, objectType, asString(data.RecId));

  return { id, url, tipo };
}

/** POST autenticado a un business object de Ivanti. */
async function ivantiPost(
  root: string,
  apiKey: string,
  objectType: IvantiObject,
  body: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const res = await fetch(`${root}/api/odata/businessobject/${objectType}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `rest_api_key=${apiKey}`,
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const detail = await res.text().catch(() => '');
    throw new Error(`Ivanti respondió ${res.status}: ${detail}`);
  }
  return (await res.json()) as Record<string, unknown>;
}

/**
 * Resuelve el contacto buscando el empleado por su correo en `employees`.
 * Devuelve su RecId (usado como ProfileLink). Fallback: ITSM_DEFAULT_PROFILE_LINK.
 */
async function resolveCustomer(
  root: string,
  apiKey: string,
  email?: string,
): Promise<string | undefined> {
  const fallback = process.env.ITSM_DEFAULT_PROFILE_LINK || undefined;
  if (!email) return fallback;
  const filter = encodeURIComponent(`PrimaryEmail eq '${email.replace(/'/g, "''")}'`);
  const url = `${root}/api/odata/businessobject/employees?$top=1&$select=RecId&$filter=${filter}`;
  try {
    const res = await fetch(url, { headers: { Authorization: `rest_api_key=${apiKey}` } });
    if (!res.ok) return fallback;
    const data = (await res.json()) as { value?: Array<{ RecId?: string }> };
    return asString(data.value?.[0]?.RecId) ?? fallback;
  } catch {
    return fallback;
  }
}

/** Cuerpo para un Incidente. */
function buildIncidente(ticket: DraftTicket, profileLink?: string): Record<string, unknown> {
  const nivel = NIVEL_IVANTI[ticket.prioridad ?? 'media'];
  const { servicio, categoria } = catalogoValido(ticket);

  const body: Record<string, unknown> = {
    Subject: ticket.asunto,
    Symptom: buildDetalle(ticket),
    Impact: nivel,
    Urgency: nivel, // Priority lo calcula Ivanti desde Impact + Urgency
    Status: 'Assigned', // "Asignado" (no "En Progreso")
    Source: 'Self Service',
    Service: servicio,
    Category: categoria,
    OwnerTeam: equipo(),
  };
  if (profileLink) body.ProfileLink = profileLink;
  return body;
}

/** Cuerpo para una Solicitud de Servicio (requiere plantilla). */
function buildSolicitud(ticket: DraftTicket, profileLink?: string): Record<string, unknown> {
  const { servicio } = catalogoValido(ticket);
  const template = plantillaSR(servicio) ?? process.env.ITSM_SR_TEMPLATE_RECID;
  if (!template) {
    throw new Error(
      `No hay plantilla de Solicitud (SvcReqTmplLink) para el servicio "${servicio}". ` +
        'Configura ITSM_SR_TEMPLATE_RECID o mapea el servicio en sr_plantillas.json.',
    );
  }

  const body: Record<string, unknown> = {
    Subject: ticket.asunto,
    Symptom: buildDetalle(ticket),
    Source: 'Self Service',
    Service: servicio,
    SvcReqTmplLink_RecID: template,
    Status: 'Assigned', // "Asignado"
    OwnerTeam: equipo(),
  };
  if (profileLink) body.ProfileLink = profileLink;
  return body;
}

/** Equipo (cola) al que se asignan todos los tickets. */
function equipo(): string {
  return process.env.ITSM_DEFAULT_TEAM ?? 'Mesa de Servicio';
}

/**
 * Normaliza servicio/categoría contra el catálogo. Si el LLM eligió algo fuera
 * de catálogo, cae a los defaults de entorno para no romper la creación.
 */
function catalogoValido(ticket: DraftTicket): { servicio: string; categoria: string } {
  const c = clasificar(ticket.servicio, ticket.categoria);
  return {
    servicio: c.servicio ?? process.env.ITSM_DEFAULT_SERVICE ?? 'General',
    categoria: c.categoria ?? ticket.categoria ?? process.env.ITSM_DEFAULT_CATEGORY ?? 'General',
  };
}

/** Detalle del ticket, anteponiendo categoría/subcategoría del usuario. */
function buildDetalle(ticket: DraftTicket): string {
  const meta: string[] = [];
  if (ticket.categoria) meta.push(`Categoría: ${ticket.categoria}`);
  if (ticket.subcategoria) meta.push(`Subcategoría: ${ticket.subcategoria}`);
  return meta.length ? `${meta.join(' | ')}\n\n${ticket.descripcion}` : ticket.descripcion;
}

/** Enlace de mejor esfuerzo al registro REST en Ivanti (si hay RecId). */
function buildTicketUrl(
  root: string,
  objectType: IvantiObject,
  recId?: string,
): string | undefined {
  if (!recId) return undefined;
  return `${root}/api/odata/businessobject/${objectType}('${recId}')`;
}

function asString(value: unknown): string | undefined {
  if (value === null || value === undefined) return undefined;
  return String(value);
}
