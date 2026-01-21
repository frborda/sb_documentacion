import { ClientError, endpoint } from '../../../_core'
import config from '../../../config'

/**
 * @swagger
 * /v1/webhook:
 *   post:
 *     summary: webhook
 *     description: Webhook de entrada que procesa la información que recibe desde BlackProxy
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
    return res.send()
  }

  if (!req.body.basculas || !req.body.identificaciones) {
    throw new ClientError('Se requieren basculas e identificaciones')
  }

  const bancasTotales = recinto.bancas.length

  const estadoBasculas = req.body.basculas.split(';')
  const estadoIdentificaciones = req.body.identificaciones.split(';')
  if (estadoBasculas.length !== bancasTotales || estadoIdentificaciones.length !== bancasTotales) {
    throw new ClientError(`Se recibieron ${estadoBasculas.length} basculas y ${estadoIdentificaciones.length} identificaciones, se esperan ${bancasTotales}`)
  }

  const resultado = {
    totalProcesados: bancasTotales,
    totalExitosos: 0,
    totalFallidos: 0,
    fallidos: []
  }

  for (let i = 0; i < bancasTotales; i++) {
    const bascula = estadoBasculas[i]
    const id = estadoIdentificaciones[i]

    try {
      // estado de bascula
      if (bascula === '1') {
        recinto.activarBascula(i)
      } else {
        recinto.desactivarBascula(i)
      }
      // estado de identificacion por huella
      if (id !== '0') {
        recinto.identificar(i, id)
      } else {
        recinto.desidentificar(i)
      }

      resultado.totalExitosos++
    } catch (err) {
      resultado.totalFallidos++
      resultado.fallidos.push({ error: err.message, banca: recinto.bancas[i], input: { bascula, id } })
    }
  }

  res.json(resultado)
})

/**
 * @swagger
 * /v1/webhook/debug:
 *   post:
 *     summary: webhook.debug
 *     description: Webhook de salida que permite comprobar el funcionamiento del replier
 *     responses: {}
 *     security:
 *       - WebhookApiKey: []
 *     tags:
 *       - Webhook
*/
export const recintoWebhookDebug = endpoint(async (req, res) => {
  if (req.headers['x-api-key'] !== config.service.webhookApiKey) {
    throw new ClientError('Autenticación inválida o vencida', 401)
  }

  const noop = typeof req.query.noop !== 'undefined'
  if (noop) return res.sendStatus(200)

  console.log('webhook', req.body)
  res.sendStatus(200)
})
