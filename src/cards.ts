import {
  AdaptiveCard,
  Fact,
  FactSet,
  OpenUrlAction,
  TextBlock,
} from '@microsoft/teams.cards';
import { CreatedTicket, DraftTicket, Requester } from './types';

/** Construye la tarjeta de confirmación que se muestra al crear el ticket. */
export function buildTicketCard(
  created: CreatedTicket,
  ticket: DraftTicket,
  requester: Requester,
): AdaptiveCard {
  const tipoLabel = created.tipo === 'solicitud' ? 'Solicitud' : 'Incidente';
  const facts: Fact[] = [
    new Fact(tipoLabel, created.id),
    new Fact('Asunto', ticket.asunto),
    new Fact('Servicio', ticket.servicio ?? '—'),
    new Fact('Categoría', ticket.categoria ?? 'general'),
    new Fact('Prioridad', ticket.prioridad ?? 'media'),
    new Fact('Solicitante', requester.name ?? requester.email ?? '—'),
  ];

  const card = new AdaptiveCard(
    new TextBlock(`✅ ${tipoLabel} creado`, { weight: 'Bolder', size: 'Large', wrap: true }),
    new TextBlock(ticket.descripcion, { wrap: true, isSubtle: true }),
    new FactSet(...facts),
  );

  if (created.url) {
    card.withActions(new OpenUrlAction(created.url, { title: 'Ver ticket' }));
  }

  return card;
}
