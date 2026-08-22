import { ChatTurn, DraftTicket } from './types';

// Estado de la conversación por cada chat de Teams.
//
// IMPORTANTE: esto vive en memoria, así que se pierde al reiniciar el proceso
// y no se comparte entre instancias. Para producción, reemplaza este Map por
// un almacenamiento persistente (Redis, Cosmos DB, tabla SQL, etc.) usando la
// misma interfaz (getHistory / append / reset).

interface ConversationState {
  history: ChatTurn[];
  pendingTicket?: DraftTicket;
}

const store = new Map<string, ConversationState>();

/** Número máximo de turnos que conservamos para no crecer sin límite. */
const MAX_TURNS = 20;

export function getHistory(conversationId: string): ChatTurn[] {
  return store.get(conversationId)?.history ?? [];
}

export function append(conversationId: string, turn: ChatTurn): ChatTurn[] {
  const history = getHistory(conversationId);
  history.push(turn);
  // Conservamos solo los últimos MAX_TURNS turnos.
  const trimmed = history.slice(-MAX_TURNS);
  const current = store.get(conversationId);
  store.set(conversationId, { ...current, history: trimmed });
  return trimmed;
}

export function getPendingTicket(conversationId: string): DraftTicket | undefined {
  return store.get(conversationId)?.pendingTicket;
}

export function setPendingTicket(conversationId: string, ticket: DraftTicket): void {
  const current = store.get(conversationId);
  store.set(conversationId, { history: current?.history ?? [], pendingTicket: ticket });
}

export function clearPendingTicket(conversationId: string): void {
  const current = store.get(conversationId);
  if (!current) return;
  store.set(conversationId, { history: current.history });
}

export function reset(conversationId: string): void {
  store.delete(conversationId);
}

// Serializa los mensajes de una conversación para impedir que dos
// confirmaciones simultáneas creen el mismo ticket dos veces.
const locks = new Map<string, Promise<void>>();

export async function withLock<T>(
  conversationId: string,
  fn: () => Promise<T>,
): Promise<T> {
  const previous = locks.get(conversationId) ?? Promise.resolve();
  const result = previous.then(fn, fn);
  const marker = result.then(
    () => undefined,
    () => undefined,
  );
  locks.set(conversationId, marker);
  void marker.finally(() => {
    if (locks.get(conversationId) === marker) locks.delete(conversationId);
  });
  return result;
}
