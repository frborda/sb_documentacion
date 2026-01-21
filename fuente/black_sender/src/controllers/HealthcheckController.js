import { endpoint } from '../_core'
import config from '../config'

// TODO agregar estado INICIANDO, OK y RECONECTANDO A BASE
// TODO chequear el estado de conexion con la base de datos

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
  res.json({
    status: 'ok',
    env: config.env,
    debug: config.debug,
    replier: {
      webhook: config.webhook.url
    },
    database: {
      host: config.database.host,
      port: config.database.port,
      name: config.database.name,
      user: config.database.user
    }
  })
})
