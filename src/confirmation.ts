import { DraftTicket } from './types';

function normalize(text: string): string {
  return text
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Solo acepta respuestas inequívocas; una frase con cambios nunca confirma. */
export function isExplicitConfirmation(text: string): boolean {
  const value = normalize(text);
  return [
    'si',
    'si confirmo',
    'confirmo',
    'confirmar',
    'crear ticket',
    'crear el ticket',
    'si crear ticket',
    'si crear el ticket',
    'de acuerdo',
    'correcto',
    'adelante',
  ].includes(value);
}

export function buildConfirmationSummary(ticket: DraftTicket): string {
  const tipo = ticket.tipo === 'solicitud' ? 'Solicitud de servicio' : 'Incidente';
  return [
    'Tengo listo este borrador:',
    '',
    `• Tipo: ${tipo}`,
    `• Asunto: ${ticket.asunto}`,
    `• Descripción: ${ticket.descripcion}`,
    `• Servicio: ${ticket.servicio ?? '—'}`,
    ...(ticket.tipo === 'solicitud' ? [`• Oferta: ${ticket.oferta ?? '—'}`] : []),
    `• Categoría: ${ticket.categoria ?? '—'}`,
    ...(ticket.subcategoria ? [`• Subcategoría: ${ticket.subcategoria}`] : []),
    `• Prioridad: ${ticket.prioridad ?? 'media'}`,
    '',
    'El ticket todavía NO ha sido creado.',
    'Responde “Confirmo” para crearlo, o indica qué deseas modificar. También puedes responder “Cancelar”.',
  ].join('\n');
}
