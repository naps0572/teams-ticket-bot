import { ChatTurn } from './types';

// Estado de la conversación por cada chat de Teams.
//
// IMPORTANTE: esto vive en memoria, así que se pierde al reiniciar el proceso
// y no se comparte entre instancias. Para producción, reemplaza este Map por
// un almacenamiento persistente (Redis, Cosmos DB, tabla SQL, etc.) usando la
// misma interfaz (getHistory / append / reset).

const store = new Map<string, ChatTurn[]>();

/** Número máximo de turnos que conservamos para no crecer sin límite. */
const MAX_TURNS = 20;

export function getHistory(conversationId: string): ChatTurn[] {
  return store.get(conversationId) ?? [];
}

export function append(conversationId: string, turn: ChatTurn): ChatTurn[] {
  const history = getHistory(conversationId);
  history.push(turn);
  // Conservamos solo los últimos MAX_TURNS turnos.
  const trimmed = history.slice(-MAX_TURNS);
  store.set(conversationId, trimmed);
  return trimmed;
}

export function reset(conversationId: string): void {
  store.delete(conversationId);
}
