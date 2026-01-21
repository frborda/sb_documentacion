import { ClientError, endpoint } from '../../../_core'
import config from '../../../config'

/**
 * @swagger
 * /v2/webhook:
 *   post:
 *     summary: webhook
 *     description: Webhook de entrada que procesa la información que recibe del controlador de Recinto
 *     responses: {}
 *     security:
 *       - WebhookApiKey: []
 *     tags:
 *       - Webhook
*/
export const webhook = endpoint(async (req, res) => {
  if (req.headers['x-api-key'] !== config.service.webhookApiKey) {
    throw new ClientError('Autenticación inválida o vencida', 401)
  }

  const recinto = req.app.locals.recinto

  if (recinto.emulacion) {
    // ignoramos si estamos emulando
    console.log('webhook nec: apagar emulacion')
    return res.send()
  }

  const noop = typeof req.query.noop !== 'undefined'
  if (noop) return res.sendStatus(200)

  console.log('webhook nec', req.body)
  res.sendStatus(200)
})
