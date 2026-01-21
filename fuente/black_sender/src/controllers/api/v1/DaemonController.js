import { endpoint } from '../../../_core'

/**
 * @swagger
 * /v1/daemon:
 *   get:
 *     summary: daemon.estado
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Daemon
*/
export const daemonEstado = endpoint(async (req, res) => {
  const bcr = req.app.locals.bcr
  res.json({
    data: bcr
  })
})

/**
 * @swagger
 * /v1/daemon:regenerar-cuiles:
 *   post:
 *     summary: daemon.regenerarCuiles
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Daemon
*/
export const daemonActionSwitch = endpoint(async (req, res, next) => {
  const bcr = req.app.locals.bcr

  switch (req.params.action) {
    case 'regenerar-cuiles':
      await bcr.regenerarCuiles()

      break
    default:
      return next()
  }
  res.sendStatus(200)
})
