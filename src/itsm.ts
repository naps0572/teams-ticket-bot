import { clasificar, ofertasSR, plantillaSR } from './categorias';
import { CreatedTicket, DraftTicket, Prioridad, Requester, TipoTicket } from './types';

// ============================================================================
//  CONECTOR DEL ITSM — Ivanti Service Manager (Neurons / ISM Cloud)
// ----------------------------------------------------------------------------
//  El resto del bot solo llama a `createTicket(...)`. Según el catálogo
//  (Categorizacion.xlsx) el ticket puede ser:
//    - Incidente  -> POST /businessobject/incidents
//    - Solicitud  -> POST /api/rest/ServiceRequest/new (oferta + parámetros)
//
//  Auth: header  Authorization: rest_api_key=<KEY>   (NO Bearer)
//
//  Campos obligatorios (verificado con sondas al API):
//    Incidente : Service, Category, Team(OwnerTeam), Customer(ProfileLink)
//    Solicitud : subscriptionId, strUserId y parameters de la oferta
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

interface IvantiRequesterProfile {
  recId?: string;
  displayName?: string;
  country?: string;
  location?: string;
  orgName?: string;
}

interface ServiceRequestTemplateParam {
  RecId?: string;
  Name?: string;
  DisplayName?: string;
  DisplayType?: string;
  RequiredExpression?: string;
}

interface ServiceRequestSubscription {
  RecId?: string;
  OrgUnitLink_RecID?: string;
}

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

  // Contacto/Customer: empleado resuelto por correo o perfil fallback.
  const profile = await resolveCustomer(root, apiKey, requester);

  let recId: string | undefined;
  let id: string;
  if (objectType === 'servicereqs') {
    const created = await createSolicitud(root, apiKey, ticket, requester, profile);
    recId = created.recId;
    id = created.id;
  } else {
    const body = buildIncidente(ticket, profile.recId);
    const data = await ivantiPost(root, apiKey, objectType, body);
    recId = asString(data.RecId);
    id = asString(data.IncidentNumber) ?? recId ?? 'desconocido';
  }
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
 * Resuelve el empleado y los datos necesarios para seleccionar la publicación
 * de la oferta. Si Teams no entrega correo, consulta el perfil fallback por RecId.
 */
async function resolveCustomer(
  root: string,
  apiKey: string,
  requester: Requester,
): Promise<IvantiRequesterProfile> {
  const fallback = process.env.ITSM_DEFAULT_PROFILE_LINK || undefined;
  const defaultProfile: IvantiRequesterProfile = {
    recId: fallback,
    displayName: requester.name,
    country: process.env.ITSM_DEFAULT_COUNTRY ?? 'El Salvador',
    orgName: process.env.ITSM_SR_SUBSCRIPTION_ORG_NAME,
  };
  const rawFilter = requester.email
    ? `PrimaryEmail eq '${requester.email.replace(/'/g, "''")}'`
    : fallback
      ? `RecId eq '${fallback.replace(/'/g, "''")}'`
      : undefined;
  if (!rawFilter) return defaultProfile;
  const url = `${root}/api/odata/businessobject/employees?$top=1&$filter=${encodeURIComponent(rawFilter)}`;
  try {
    const res = await fetch(url, { headers: { Authorization: `rest_api_key=${apiKey}` } });
    if (!res.ok) return defaultProfile;
    const data = (await res.json()) as {
      value?: Array<{
        RecId?: string;
        DisplayName?: string;
        Address1Country?: string;
        Address1City?: string;
        AA_OrgPadre?: string;
      }>;
    };
    const employee = data.value?.[0];
    if (!employee) return defaultProfile;
    return {
      recId: asString(employee.RecId) ?? fallback,
      displayName: asString(employee.DisplayName) ?? requester.name,
      country:
        asString(employee.Address1Country) ?? process.env.ITSM_DEFAULT_COUNTRY ?? 'El Salvador',
      location: asString(employee.Address1City) ?? asString(employee.Address1Country),
      orgName: process.env.ITSM_SR_SUBSCRIPTION_ORG_NAME ?? asString(employee.AA_OrgPadre),
    };
  } catch {
    return defaultProfile;
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

/** Crea una Solicitud usando la API de ofertas, que instancia ServiceReqParam. */
async function createSolicitud(
  root: string,
  apiKey: string,
  ticket: DraftTicket,
  requester: Requester,
  profile: IvantiRequesterProfile,
): Promise<{ id: string; recId?: string }> {
  const { servicio, categoria } = catalogoValido(ticket);
  const disponibles = ofertasSR(servicio);
  if (!ticket.oferta && disponibles.length > 1) {
    throw new Error(
      `El servicio "${servicio}" tiene varias ofertas; se debe seleccionar una antes de crear el SR.`,
    );
  }
  const oferta = ticket.oferta ?? disponibles[0]?.nombre;
  const mappedTemplate = oferta ? plantillaSR(servicio, oferta) : plantillaSR(servicio);
  const template = mappedTemplate ?? (!oferta ? process.env.ITSM_SR_TEMPLATE_RECID : undefined);
  if (!template) {
    throw new Error(
      `No hay plantilla de Solicitud (SvcReqTmplLink) para la oferta ` +
        `"${ticket.oferta ?? 'predeterminada'}" del servicio "${servicio}". ` +
        'Configura la oferta en sr_plantillas.json.',
    );
  }
  if (!profile.recId) {
    throw new Error(
      'No se pudo resolver el RecId del solicitante. Configura ITSM_DEFAULT_PROFILE_LINK.',
    );
  }

  const subscriptionId = await resolveSubscription(root, apiKey, template, profile.orgName);
  const definitions = await odataList<ServiceRequestTemplateParam>(
    root,
    apiKey,
    'servicereqtemplateparams',
    `ParentLink_RecID eq '${template}'`,
  );
  const country = profile.country ?? process.env.ITSM_DEFAULT_COUNTRY ?? 'El Salvador';
  const parameters = await buildServiceRequestParameters(
    root,
    apiKey,
    definitions,
    ticket,
    servicio,
    categoria,
    country,
    profile,
    requester,
  );

  const serviceReqData: Record<string, unknown> = {
    Subject: ticket.asunto,
    ...serviceRequestContent(ticket),
    Service: servicio,
    Urgency: NIVEL_IVANTI[ticket.prioridad ?? 'media'],
    Source: process.env.ITSM_SR_SOURCE ?? 'Direct Support',
    OwnerTeam: equipo(),
    AA_Category: categoria,
    AA_Subcategory: ticket.subcategoria ?? categoria,
    AA_Pais: country,
  };

  const body = {
    attachmentsToDelete: [],
    attachmentsToUpload: [],
    parameters,
    delayedFulfill: false,
    formName: process.env.ITSM_SR_FORM_NAME ?? 'ServiceReq.ResponsiveAnalyst.DefaultLayout',
    saveReqState: false,
    serviceReqData,
    strCustomerLocation: profile.location ?? country,
    strUserId: profile.recId,
    subscriptionId,
    localOffset: Number(process.env.ITSM_LOCAL_OFFSET ?? -360),
  };

  const res = await fetch(`${root}/api/rest/ServiceRequest/new`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `rest_api_key=${apiKey}`,
    },
    body: JSON.stringify(body),
  });
  const detail = await res.text().catch(() => '');
  if (!res.ok) {
    throw new Error(`Ivanti rechazó la creación de la solicitud (${res.status}): ${detail}`);
  }
  const data = detail ? (JSON.parse(detail) as Record<string, unknown>) : {};
  if (data.IsSuccess === false) {
    throw new Error(`Ivanti no creó la solicitud: ${asString(data.ErrorText) ?? 'error desconocido'}`);
  }
  const requests = data.ServiceRequests as Array<Record<string, unknown>> | undefined;
  const created = requests?.[0];
  const recId = asString(created?.strRequestRecId) ?? asString(created?.RecId);
  const id =
    asString(created?.strRequestNum) ?? asString(created?.ServiceReqNumber) ?? recId ?? 'desconocido';
  if (!created || id === 'desconocido') {
    throw new Error(`Ivanti no devolvió la solicitud creada: ${detail.slice(0, 800)}`);
  }
  return { id, recId };
}

/** Resuelve la publicación (ServiceReqSubscription) correspondiente a la oferta. */
async function resolveSubscription(
  root: string,
  apiKey: string,
  templateRecId: string,
  orgName?: string,
): Promise<string> {
  const configured = process.env.ITSM_SR_SUBSCRIPTION_RECID;
  if (configured) return configured;
  const subscriptions = await odataList<ServiceRequestSubscription>(
    root,
    apiKey,
    'servicereqsubscriptions',
    `SvcReqTmplLink_RecID eq '${templateRecId}'`,
  );
  const valid = subscriptions.filter((item) => asString(item.RecId));
  if (valid.length === 0) {
    throw new Error(`La plantilla ${templateRecId} no tiene una oferta publicada en Ivanti.`);
  }

  const configuredOrg = process.env.ITSM_SR_SUBSCRIPTION_ORG_RECID;
  if (configuredOrg) {
    const selected = valid.find((item) => item.OrgUnitLink_RecID === configuredOrg);
    if (selected?.RecId) return selected.RecId;
  }

  if (orgName) {
    for (const subscription of valid) {
      if (!subscription.OrgUnitLink_RecID) continue;
      const units = await odataList<{ RecId?: string; Name?: string }>(
        root,
        apiKey,
        'organizationalunits',
        `RecId eq '${subscription.OrgUnitLink_RecID}'`,
      );
      if (normalizeIvanti(units[0]?.Name) === normalizeIvanti(orgName) && subscription.RecId) {
        return subscription.RecId;
      }
    }
  }

  if (valid.length === 1 && valid[0].RecId) return valid[0].RecId;
  throw new Error(
    `La plantilla ${templateRecId} tiene varias publicaciones. ` +
      'Configura ITSM_SR_SUBSCRIPTION_ORG_NAME o ITSM_SR_SUBSCRIPTION_ORG_RECID.',
  );
}

/** Convierte los campos del borrador en parámetros par-{RecId} de Ivanti. */
async function buildServiceRequestParameters(
  root: string,
  apiKey: string,
  definitions: ServiceRequestTemplateParam[],
  ticket: DraftTicket,
  service: string,
  category: string,
  country: string,
  profile: IvantiRequesterProfile,
  requester: Requester,
): Promise<Record<string, string>> {
  const parameters: Record<string, string> = {};
  for (const definition of definitions) {
    const recId = asString(definition.RecId);
    if (!recId) continue;
    if (['category', 'rowaligner'].includes((definition.DisplayType ?? '').toLowerCase())) continue;
    const mapped = mapTemplateParameter(
      definition.Name,
      ticket,
      service,
      category,
      country,
      profile,
      requester,
    );
    if (!mapped) {
      if (/\$\(\s*true\s*\)/i.test(definition.RequiredExpression ?? '')) {
        throw new Error(
          `La oferta exige el parámetro "${definition.DisplayName ?? definition.Name ?? recId}" ` +
            'y el bot todavía no tiene un valor para completarlo.',
        );
      }
      continue;
    }
    const key = `par-${recId}`;
    parameters[key] = mapped.value;
    const validationRecId =
      definition.DisplayType === 'combo' && mapped.value
        ? (await resolveValidationValue(root, apiKey, recId, mapped.value)) ??
          (await resolveDependentValidationValue(
            root,
            apiKey,
            definition.Name,
            mapped.value,
            service,
            category,
          ))
        : undefined;
    const optionRecId = validationRecId ?? mapped.recId;
    if (optionRecId) parameters[`${key}-recId`] = optionRecId;
  }
  return parameters;
}

function mapTemplateParameter(
  name: string | undefined,
  ticket: DraftTicket,
  service: string,
  category: string,
  country: string,
  profile: IvantiRequesterProfile,
  requester: Requester,
): { value: string; recId?: string } | undefined {
  const field = normalizeIvanti(name).replace(/_/g, '');
  const requesterName = profile.displayName ?? requester.name ?? requester.email ?? '';
  if (field.includes('adjunt')) return { value: '' };
  if (field.includes('subcategor')) return { value: ticket.subcategoria ?? category };
  if (field.includes('categor')) return { value: category };
  if (field.includes('servicio') || field.includes('service')) return { value: service };
  if (field.includes('pais') || field.includes('country')) return { value: country };
  if (field.includes('resumen') || field.includes('titulo') || field.includes('subject')) {
    return { value: ticket.asunto };
  }
  if (field.includes('descripcion') || field.includes('detalle') || field.includes('symptom')) {
    return { value: ticket.descripcion };
  }
  if (field.includes('recid')) return profile.recId ? { value: profile.recId } : undefined;
  if (
    field.includes('usuario') ||
    field.includes('solicitante') ||
    field.includes('reporta') ||
    field.includes('requiere')
  ) {
    return requesterName
      ? { value: requesterName, ...(profile.recId ? { recId: profile.recId } : {}) }
      : undefined;
  }
  return undefined;
}

/** RecId de un valor de combo. Algunas listas dependientes devuelven vacío; el texto sigue siendo válido. */
async function resolveValidationValue(
  root: string,
  apiKey: string,
  parameterRecId: string,
  expected: string,
): Promise<string | undefined> {
  try {
    const res = await fetch(
      `${root}/api/rest/ServiceRequest/${encodeURIComponent(parameterRecId)}/ValidationList`,
      { headers: { Authorization: `rest_api_key=${apiKey}` } },
    );
    if (!res.ok) return undefined;
    const rows = (await res.json()) as unknown[][];
    const match = rows.find(
      (row) =>
        normalizeIvanti(asString(row[1])) === normalizeIvanti(expected) ||
        normalizeIvanti(asString(row[2])) === normalizeIvanti(expected),
    );
    return asString(match?.[0]);
  } catch {
    return undefined;
  }
}

/** Resuelve combos dependientes que ValidationList no devuelve sin contexto previo. */
async function resolveDependentValidationValue(
  root: string,
  apiKey: string,
  parameterName: string | undefined,
  expected: string,
  service: string,
  category: string,
): Promise<string | undefined> {
  const field = normalizeIvanti(parameterName).replace(/_/g, '');
  const serviceValue = escapeOData(service);
  const categoryValue = escapeOData(category);
  if (field.includes('subcategor')) {
    const rows = await odataList<{ RecId?: string }>(
      root,
      apiKey,
      'aa_subcategoryservicereqs',
      `Service eq '${serviceValue}' and Category eq '${categoryValue}' and ` +
        `Subcategory eq '${escapeOData(expected)}'`,
    );
    return asString(rows[0]?.RecId);
  }
  if (field.includes('categor')) {
    const rows = await odataList<{ RecId?: string }>(
      root,
      apiKey,
      'aa_categoryservicereqs',
      `Service eq '${serviceValue}' and Category eq '${escapeOData(expected)}'`,
    );
    return asString(rows[0]?.RecId);
  }
  return undefined;
}

async function odataList<T>(
  root: string,
  apiKey: string,
  objectName: string,
  filter: string,
): Promise<T[]> {
  const url = `${root}/api/odata/businessobject/${objectName}?$filter=${encodeURIComponent(filter)}`;
  const res = await fetch(url, { headers: { Authorization: `rest_api_key=${apiKey}` } });
  if (res.status === 204) return [];
  const detail = await res.text().catch(() => '');
  if (!res.ok) {
    throw new Error(`Ivanti rechazó la consulta de ${objectName} (${res.status}): ${detail}`);
  }
  const data = detail ? (JSON.parse(detail) as { value?: T[] }) : {};
  return data.value ?? [];
}

function normalizeIvanti(value: string | undefined): string {
  return (value ?? '')
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .toLowerCase()
    .trim();
}

function escapeOData(value: string): string {
  return value.replace(/'/g, "''");
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
    [descriptionField]: ticket.descripcion,
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
