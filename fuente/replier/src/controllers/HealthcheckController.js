import { endpoint } from '../_core'
import config from '../config'

// TODO agregar estado INICIANDO y OK

/**
 * @swagger
 * /healthcheck:
 *   get:
 *     summary: healthcheck
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Status
*/
export const healthcheck = endpoint(async (req, res) => {
  const recinto = req.app.locals.recinto
  const replier = req.app.locals.replier
  res.json({
    status: 'ok',
    env: config.env,
    debug: config.debug,
    recinto: {
      nroBancas: recinto.bancas.length,
      presidentePorDefecto: recinto.presidentePorDefecto,
      emulacion: recinto.emulacion,
      webhook: config.recinto.webhook.url,
      diputados: config.recinto.diputados.url,
      replyMilliseconds: replier.replyMilliseconds,
      replyPaused: replier.isPaused()
    },
    votaciones: {
      webhook: config.votaciones.webhook.url
    }
  })
})
