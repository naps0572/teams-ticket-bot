import { catalogoParaPrompt, ofertasParaPrompt } from './categorias';
import { AssistantResult, ChatTurn } from './types';

// Cliente de OpenRouter. La API es compatible con el formato de OpenAI, así
// que solo necesitamos un fetch al endpoint de chat completions.

const OPENROUTER_URL = 'https://openrouter.ai/api/v1/chat/completions';

const SYSTEM_PROMPT = `Eres un asistente de soporte técnico dentro de Microsoft Teams.
Tu única función es ayudar al usuario a levantar un ticket de soporte de forma clara y rápida.

Debes recopilar, mediante una conversación natural y breve, estos datos:
- asunto: título corto del problema.
- descripcion: explicación con el detalle suficiente para que un técnico entienda qué pasa.
- tipo: "incidente" si algo FALLA o dejó de funcionar; "solicitud" si el usuario PIDE algo (acceso, alta, instalación, información).
- servicio: elígelo EXACTAMENTE de la lista del catálogo de abajo (usa el nombre tal cual).
- oferta: para una solicitud, elígela EXACTAMENTE de las ofertas publicadas para ese servicio.
- categoria: elígela EXACTAMENTE de las categorías de ese servicio en el catálogo.
- prioridad: uno de estos valores exactos: "baja", "media", "alta", "critica".

Catálogo oficial de categorización (Servicio → Categorías). Cada categoría indica
entre corchetes si aplica a Incidente [IN], Solicitud [SR] o ambos [IN/SR]:
${catalogoParaPrompt()}

Ofertas de servicio publicadas en Ivanti (Servicio → Ofertas):
${ofertasParaPrompt()}

Reglas:
1. Haz solo las preguntas necesarias. Si el usuario ya dio un dato, no lo vuelvas a pedir.
2. Elige servicio y categoria SOLO del catálogo; no inventes valores. Usa las palabras clave para acertar.
   Para una solicitud, elige también una oferta de la lista correspondiente al servicio.
3. El "tipo" debe ser coherente con la categoría: si la categoría es solo [IN] usa "incidente"; si es solo [SR] usa "solicitud"; si es [IN/SR], decide según sea una falla o una petición.
4. Si falta información, pregunta de forma concreta y amable por lo que falte (una o dos cosas a la vez).
5. Cuando tengas asunto, descripcion, tipo, servicio, categoria y prioridad, marca "listo": true. Si el tipo es "solicitud", también debes tener la oferta. Esto significa solamente que el BORRADOR está completo; el backend pedirá confirmación al usuario antes de crear el ticket.
6. Si el usuario no especifica prioridad, propón una razonable según el impacto y confírmala.
7. Responde SIEMPRE en español y de forma concisa.
8. NUNCA afirmes que el ticket fue creado, registrado o enviado. Solo el backend puede decirlo después de recibir una respuesta exitosa de Ivanti.

Responde EXCLUSIVAMENTE con un objeto JSON válido, sin texto adicional ni bloques de código, con esta forma:
{
  "respuesta": "<lo que le dirás al usuario>",
  "listo": <true|false>,
  "ticket": { "asunto": "", "descripcion": "", "tipo": "incidente|solicitud", "servicio": "", "oferta": "", "categoria": "", "subcategoria": "", "prioridad": "" } | null
}
Cuando "listo" sea false, "ticket" debe ser null. Cuando sea true, "ticket" debe llevar todos los campos.`;

interface OpenRouterMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

function buildMessages(history: ChatTurn[]): OpenRouterMessage[] {
  return [
    { role: 'system', content: SYSTEM_PROMPT },
    ...history.map((t) => ({ role: t.role, content: t.content })),
  ];
}

/** Quita posibles ```json ... ``` que algunos modelos añaden. */
function stripFences(text: string): string {
  return text.replace(/```json/gi, '').replace(/```/g, '').trim();
}

function parseResult(raw: string): AssistantResult {
  const clean = stripFences(raw);
  const parsed = JSON.parse(clean) as Partial<AssistantResult>;
  return {
    respuesta: parsed.respuesta ?? 'Perdona, no entendí bien. ¿Podrías repetirlo?',
    listo: parsed.listo === true,
    ticket: parsed.listo === true ? (parsed.ticket ?? null) : null,
  };
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Reintentos ante saturación temporal (429) o errores transitorios del
 *  proveedor (5xx). Útil sobre todo con modelos `:free`, que se rate-limitean
 *  aguas arriba con frecuencia. Configurable con OPENROUTER_MAX_RETRIES. */
const MAX_RETRIES = Number(process.env.OPENROUTER_MAX_RETRIES ?? 3);

export async function runTicketAssistant(history: ChatTurn[]): Promise<AssistantResult> {
  const apiKey = process.env.OPENROUTER_API_KEY;
  if (!apiKey) {
    throw new Error('Falta la variable de entorno OPENROUTER_API_KEY.');
  }

  const model = process.env.OPENROUTER_MODEL || 'anthropic/claude-3.5-sonnet';

  let res!: Response;
  let lastDetail = '';
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    res = await fetch(OPENROUTER_URL, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
        // Cabeceras opcionales que OpenRouter usa para atribución (rankings).
        ...(process.env.OPENROUTER_SITE_URL
          ? { 'HTTP-Referer': process.env.OPENROUTER_SITE_URL }
          : {}),
        ...(process.env.OPENROUTER_APP_NAME
          ? { 'X-Title': process.env.OPENROUTER_APP_NAME }
          : {}),
      },
      body: JSON.stringify({
        model,
        messages: buildMessages(history),
        temperature: 0.3,
        response_format: { type: 'json_object' },
      }),
    });

    // 429 (rate limit) y 5xx suelen ser transitorios: reintentamos con backoff.
    if ((res.status === 429 || res.status >= 500) && attempt < MAX_RETRIES) {
      lastDetail = await res.text().catch(() => '');
      const backoffMs = 800 * 2 ** attempt; // 0.8s, 1.6s, 3.2s...
      await sleep(backoffMs);
      continue;
    }
    break;
  }

  if (!res.ok) {
    const detail = (await res.text().catch(() => '')) || lastDetail;
    throw new Error(`OpenRouter respondió ${res.status}: ${detail}`);
  }

  const data = (await res.json()) as {
    choices?: Array<{ message?: { content?: string } }>;
  };
  const content = data.choices?.[0]?.message?.content;
  if (!content) {
    throw new Error('OpenRouter no devolvió contenido en la respuesta.');
  }

  try {
    return parseResult(content);
  } catch {
    // Si el modelo no devolvió JSON parseable, tratamos el texto como una
    // respuesta conversacional normal (sin cerrar el ticket).
    return { respuesta: stripFences(content), listo: false, ticket: null };
  }
}
