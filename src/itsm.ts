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

  const recId = asString(data.RecId);
  const id =
    asString(data.IncidentNumber) ??
    asString(data.ServiceReqNumber) ??
    recId ??
    'desconocido';
  const url = buildTicketUrl(root, objectType, recId, id);

  // Algunas plantillas de SR reemplazan campos durante el POST. Reaplicamos
  // resumen y descripción una vez que Ivanti ya creó e inicializó el registro.
  // Si falla, no lanzamos error: el SR ya existe y reintentar podría duplicarlo.
  if (objectType === 'servicereqs' && recId) {
    try {
      await ivantiPatch(root, apiKey, objectType, recId, serviceRequestContent(ticket));
    } catch (err) {
      // eslint-disable-next-line no-console
      console.warn(
        `[itsm] SR ${id} fue creado, pero no se pudieron reafirmar resumen/descripción:`,
        err,
      );
    }
  }

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

/** PATCH autenticado a un registro ya creado. */
async function ivantiPatch(
  root: string,
  apiKey: string,
  objectType: IvantiObject,
  recId: string,
  body: Record<string, unknown>,
): Promise<void> {
  const res = await fetch(
    `${root}/api/odata/businessobject/${objectType}('${encodeURIComponent(recId)}')`,
    {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `rest_api_key=${apiKey}`,
      },
      body: JSON.stringify(body),
    },
  );
  if (!res.ok) {
    const detail = await res.text().catch(() => '');
    throw new Error(`Ivanti rechazó la actualización del SR (${res.status}): ${detail}`);
  }
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
    ...serviceRequestContent(ticket),
    Source: 'Self Service',
    Service: servicio,
    SvcReqTmplLink_RecID: template,
    Status: 'Assigned', // "Asignado"
    OwnerTeam: equipo(),
  };
  if (profileLink) body.ProfileLink = profileLink;
  return body;
}

/**
 * Campos que alimentan el resumen y la descripción visibles del SR.
 * Los nombres se pueden adaptar si otro ambiente de Ivanti usa campos propios.
 */
export function serviceRequestContent(ticket: DraftTicket): Record<string, string> {
  const summaryField = process.env.ITSM_SR_SUMMARY_FIELD || 'AA_Resumen';
  const descriptionField = process.env.ITSM_SR_DESCRIPTION_FIELD || 'Symptom';
  return {
    [summaryField]: ticket.asunto,
    [descriptionField]: buildDetalle(ticket),
  };
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

/**
 * Enlace a la interfaz Self Service de Ivanti, no al endpoint OData.
 * ITSM_TICKET_URL_TEMPLATE permite adaptar SSO o una UI personalizada con:
 * {baseUrl}, {recId}, {id} y {type}.
 */
export function buildTicketUrl(
  root: string,
  objectType: IvantiObject,
  recId?: string,
  id?: string,
): string | undefined {
  if (!recId) return undefined;
  const type = objectType === 'servicereqs' ? 'ServiceReq' : 'Incident';
  const template = process.env.ITSM_TICKET_URL_TEMPLATE;

  if (template) {
    return template
      .replaceAll('{baseUrl}', root)
      .replaceAll('{recId}', encodeURIComponent(recId))
      .replaceAll('{id}', encodeURIComponent(id ?? recId))
      .replaceAll('{type}', type);
  }

  // Formato oficial para abrir el registro en My Items (Self Service UI V2/V3).
  return `${root}/Modules/SelfService/#myItems/view/${encodeURIComponent(recId)}`;
}

function asString(value: unknown): string | undefined {
  if (value === null || value === undefined) return undefined;
  return String(value);
}
