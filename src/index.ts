import { App } from '@microsoft/teams.apps';
import { MessageActivity } from '@microsoft/teams.api';
import { ConsoleLogger } from '@microsoft/teams.common';
import { DevtoolsPlugin } from '@microsoft/teams.dev';

import { buildTicketCard } from './cards';
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

  // Identidad del solicitante (viene de Teams).
  const requester: Requester = {
    name: activity.from?.name,
    aadObjectId: activity.from?.aadObjectId,
    // El correo suele venir de Graph; aquí lo dejamos abierto para enriquecer
    // con `api.user.*` o Graph si lo necesitas.
  };

  const history = store.append(conversationId, { role: 'user', content: text });

  // Indicador de "escribiendo..." mientras consultamos al modelo.
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

  // Guardamos la respuesta del asistente en el historial.
  store.append(conversationId, { role: 'assistant', content: result.respuesta });

  // Todavía falta información: seguimos conversando.
  if (!result.listo || !result.ticket) {
    await send(result.respuesta);
    return;
  }

  // Ya hay datos suficientes: creamos el ticket en el ITSM.
  try {
    const created = await createTicket(result.ticket, requester);
    const card = buildTicketCard(created, result.ticket, requester);
    const tipoLabel = created.tipo === 'solicitud' ? 'Tu solicitud' : 'Tu incidente';
    await send(
      new MessageActivity(`${tipoLabel} ${created.id} fue creado.`).addCard('adaptive', card),
    );
    // Cerramos la conversación para el siguiente ticket.
    store.reset(conversationId);
  } catch (err) {
    log.error(err);
    await send(
      'Tengo todos los datos, pero no pude crear el ticket en el sistema. ' +
        'Vuelve a intentarlo o avisa a soporte si el problema persiste.',
    );
  }
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
