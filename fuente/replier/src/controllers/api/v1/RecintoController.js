import { ClientError, endpoint } from '../../../_core'

/**
 * @swagger
 * /v1/recinto:sync:
 *   get:
 *     summary: recinto.obtenerSync
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Recinto
*/
export const recintoQuerySwitch = endpoint(async (req, res, next) => {
  const recinto = req.app.locals.recinto

  switch (req.params.action) {
    case 'sync':
      res.json(recinto.serializeToSync())
      return

    default:
      return next()
  }
})

/**
 * @swagger
 * /v1/recinto:configurar:
 *   post:
 *     summary: recinto.configurar
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Recinto
 *
 * /v1/recinto:limpiar-identificaciones:
 *   post:
 *     summary: recinto.limpiarIdentificaciones
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Recinto
*/
export const recintoActionSwitch = endpoint(async (req, res, next) => {
  const recinto = req.app.locals.recinto

  switch (req.params.action) {
    case 'configurar':
      if (typeof (req.body.emulacion) === 'undefined') {
        throw new ClientError('emulacion es requerido')
      }
      if (req.body.emulacion) recinto.activarEmulacion()
      else recinto.desactivarEmulacion()

      break

    case 'limpiar-identificaciones':
      recinto.limpiarIdentificaciones()

      break
    default:
      return next()
  }
  res.sendStatus(200)
})

/**
 * @swagger
 * /v1/recinto:
 *   get:
 *     summary: recinto.obtener
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Recinto
*/
export const obtenerRecinto = endpoint(async (req, res, next) => {
  const recinto = req.app.locals.recinto

  res.json({
    totalBancas: recinto.nroBancas,
    totalBasculasActivas: recinto.nroBasculasActivas,
    totalIdentificados: recinto.nroIdentificados,
    lastModified: recinto.lastModified,
    data: recinto
  })
})
