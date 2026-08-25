import { App } from '@microsoft/teams.apps';
import { MessageActivity } from '@microsoft/teams.api';
import { ConsoleLogger } from '@microsoft/teams.common';
import { DevtoolsPlugin } from '@microsoft/teams.dev';

import { buildTicketCard } from './cards';
import { ofertaSR, ofertasSR, subcategoriaSR, subcategoriasSR } from './categorias';
import { buildConfirmationSummary, isExplicitConfirmation } from './confirmation';
import * as store from './conversation';
import { createTicket } from './itsm';
import { runTicketAssistant } from './openrouter';
import { Requester } from './types';

const isDev = process.env.NODE_ENV !== 'production';

const app = new App({
  logger: new ConsoleLogger('@app/ticket-bot', { level: isDev ? 'debug' : 'info' }),
  // DevTools SOLO en desarrollo: da una interfaz web de pruebas sin auth.
  // No debe habilitarse en producción.
  plugins: isDev ? [new DevtoolsPlugin()] : [],
});

app.on('message', async ({ activity, send, log }) => {
  const conversationId = activity.conversation.id;
  const text = (activity.text ?? '').trim();

  if (!text) {
    await send('Cuéntame qué problema tienes y te ayudo a levantar un ticket. 🙂');
    return;
  }

  // Comando rápido para empezar de nuevo.
  if (/^(reiniciar|reset|cancelar|nuevo)$/i.test(text)) {
    store.reset(conversationId);
    await send('Listo, empecemos de nuevo. ¿Cuál es el problema?');
    return;
  }

  await store.withLock(conversationId, async () => {
    // Identidad del solicitante (viene de Teams).
    const requester: Requester = {
      name: activity.from?.name,
      aadObjectId: activity.from?.aadObjectId,
      // El correo suele venir de Graph; aquí lo dejamos abierto para enriquecer
      // con `api.user.*` o Graph si lo necesitas.
    };

    const pendingTicket = store.getPendingTicket(conversationId);
    if (pendingTicket && isExplicitConfirmation(text)) {
      // Consumimos la confirmación antes del POST. El lock impide que otra
      // confirmación de la misma conversación cree un duplicado.
      store.clearPendingTicket(conversationId);
      await send({ type: 'typing' });
      try {
        const created = await createTicket(pendingTicket, requester);
        const card = buildTicketCard(created, pendingTicket, requester);
        const tipoLabel = created.tipo === 'solicitud' ? 'Tu solicitud' : 'Tu incidente';
        await send(
          new MessageActivity(`${tipoLabel} ${created.id} fue creado.`).addCard('adaptive', card),
        );
        store.reset(conversationId);
      } catch (err) {
        // Conservamos el borrador para que el usuario pueda reintentar.
        store.setPendingTicket(conversationId, pendingTicket);
        log.error(err);
        await send(
          'Confirmaste el borrador, pero no pude crear el ticket en el sistema. ' +
            'Responde “Confirmo” para reintentar o “Cancelar” para descartarlo.',
        );
      }
      return;
    }

    // Si había un borrador y la respuesta no fue una confirmación inequívoca,
    // se interpreta como una corrección y se vuelve a consultar al agente.
    if (pendingTicket) store.clearPendingTicket(conversationId);

    const history = store.append(conversationId, { role: 'user', content: text });
    await send({ type: 'typing' });

    let result;
    try {
      result = await runTicketAssistant(history);
    } catch (err) {
      log.error(err);
      await send(
        'Uy, tuve un problema al procesar tu mensaje. Inténtalo de nuevo en un momento.',
      );
      return;
    }

    if (!result.listo || !result.ticket) {
      store.append(conversationId, { role: 'assistant', content: result.respuesta });
      await send(result.respuesta);
      return;
    }

    // Una Solicitud debe llevar una oferta explícita. Si el servicio solo tiene
    // una, la completamos; si tiene varias, el usuario debe escoger antes de
    // llegar a la confirmación para evitar una plantilla predeterminada errónea.
    if (result.ticket.tipo === 'solicitud') {
      const disponibles = ofertasSR(result.ticket.servicio);
      const seleccionada = ofertaSR(result.ticket.servicio, result.ticket.oferta);
      if (seleccionada) {
        result.ticket.oferta = seleccionada.nombre;
      } else if (disponibles.length === 1) {
        result.ticket.oferta = disponibles[0].nombre;
      } else if (disponibles.length > 1) {
        const opciones = disponibles.map((oferta) => `• ${oferta.nombre}`).join('\n');
        const respuesta = result.ticket.oferta
          ? `La oferta “${result.ticket.oferta}” no corresponde al servicio ` +
            `“${result.ticket.servicio}”. Elige una de estas opciones:\n${opciones}`
          : `Antes de preparar la solicitud, elige la oferta de servicio:\n${opciones}`;
        store.append(conversationId, { role: 'assistant', content: respuesta });
        await send(respuesta);
        return;
      }

      const subcategorias = subcategoriasSR(result.ticket.servicio, result.ticket.categoria);
      const subcategoria = subcategoriaSR(
        result.ticket.servicio,
        result.ticket.categoria,
        result.ticket.subcategoria,
      );
      if (subcategoria) {
        result.ticket.subcategoria = subcategoria;
      } else if (subcategorias.length === 1) {
        result.ticket.subcategoria = subcategorias[0];
      } else if (subcategorias.length > 1) {
        const opciones = subcategorias.map((item) => `• ${item}`).join('\n');
        const respuesta = result.ticket.subcategoria
          ? `La subcategoría “${result.ticket.subcategoria}” no es válida. Elige una:\n${opciones}`
          : `Para completar los parámetros, elige la subcategoría:\n${opciones}`;
        store.append(conversationId, { role: 'assistant', content: respuesta });
        await send(respuesta);
        return;
      }
    }

    // Un borrador completo nunca se crea en este mismo turno. Primero se
    // guarda y se muestra un resumen controlado por el backend.
    const summary = buildConfirmationSummary(result.ticket);
    store.setPendingTicket(conversationId, result.ticket);
    store.append(conversationId, { role: 'assistant', content: summary });
    await send(summary);
  });
});

app
  .start(Number(process.env.PORT) || 3978)
  .then(() => {
    // eslint-disable-next-line no-console
    console.log('Bot de tickets iniciado 🚀');
  })
  .catch((err) => {
    // eslint-disable-next-line no-console
    console.error('No se pudo iniciar el bot:', err);
    process.exit(1);
  });
