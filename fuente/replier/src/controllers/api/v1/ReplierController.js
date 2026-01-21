import { ClientError, endpoint } from '../../../_core'
import config from '../../../config'

/**
 * @swagger
 * /v1/replier:pause:
 *   post:
 *     summary: replier.pause
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Replier
 *
 * /v1/replier:unpause:
 *   post:
 *     summary: replier.unpause
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Replier
*/
export const replierActionSwitch = endpoint(async (req, res, next) => {
  const replier = req.app.locals.replier

  switch (req.params.action) {
    case 'pause':
      {
        if (config.debug === false) throw new ClientError('Solo se puede pausar replier en modo debug')
        const autoUnpause = typeof (req.body.auto_unpause) === 'boolean' ? req.body.auto_unpause : false
        await replier.pause(autoUnpause)
      }
      break

    case 'unpause':
      await replier.unpause()
      break

    default:
      return next()
  }
  res.sendStatus(200)
})
