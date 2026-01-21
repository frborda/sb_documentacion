import { Votacion } from '../../../domain/Votacion'
import { ClientError, endpoint } from '../../../_core'

function validarNormalizarIDVotacion (idVotacion) {
  if (typeof (idVotacion) === 'string') {
    idVotacion = parseFloat(idVotacion)
  }
  if (typeof (idVotacion) !== 'number' || Number.isInteger(idVotacion) === false) {
    throw new ClientError('idVotacion debe ser un entero')
  }
  if (idVotacion < 1) {
    throw new ClientError('idVotacion debe ser mayor a 0')
  }
  return idVotacion
}

function validarNormalizarCuil (cuil) {
  if (typeof (cuil) === 'string') {
    cuil = parseFloat(cuil)
  }
  if (typeof (cuil) !== 'number' || Number.isInteger(cuil) === false) {
    throw new ClientError('cuil debe ser un entero')
  }
  if (cuil < 1) {
    throw new ClientError('cuil debe ser mayor a 0')
  }
  return cuil
}

/**
 * @swagger
 * /v1/votaciones:
 *   get:
 *     summary: votaciones.listar
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Votaciones
*/
export const listarVotaciones = endpoint(async (req, res) => {
  const votaciones = req.app.locals.votaciones
  res.json({
    totalVotaciones: Object.keys(votaciones).length,
    data: Object.values(votaciones)
  })
})

/**
 * @swagger
 * /v1/votaciones:
 *   delete:
 *     summary: votaciones.limpiar
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Votaciones
*/
export const limpiarVotaciones = endpoint(async (req, res) => {
  const votaciones = req.app.locals.votaciones

  for (const key in votaciones) {
    if (Object.prototype.hasOwnProperty.call(votaciones, key)) {
      const votacion = votaciones[key]
      if (votacion.estado !== 'INICIADA') {
        // solo eliminamos las votaciones que no estan iniciadas
        delete votaciones[key]
      }
    }
  }

  res.sendStatus(200)
})

/**
 * @swagger
 * /v1/votaciones/{idVotacion}:
 *   get:
 *     summary: votaciones.obtener
 *     parameters:
 *       - name: idVotacion
 *         in: path
 *         required: true
 *         description: ID de votación.
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Votaciones
*/
export const obtenerVotacion = endpoint(async (req, res) => {
  const idVotacion = validarNormalizarIDVotacion(req.params.idVotacion)
  const votacion = req.app.locals.votaciones[idVotacion]
  if (!votacion) {
    throw new ClientError('votación no existe', 404)
  }

  res.json({
    data: votacion
  })
})

/**
 * @swagger
 * /v1/votaciones/{idVotacion}:sync:
 *   get:
 *     summary: votaciones.obtenerSync
 *     parameters:
 *       - name: idVotacion
 *         in: path
 *         required: true
 *         description: ID de votación.
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Votaciones
*/
export const votacionQuerySwitch = endpoint(async (req, res, next) => {
  const idVotacion = validarNormalizarIDVotacion(req.params.idVotacion)
  const votacion = req.app.locals.votaciones[idVotacion]
  if (!votacion) {
    throw new ClientError('votación no existe', 404)
  }

  switch (req.params.action) {
    case 'sync':
      res.json(votacion.serializeToSync())
      return

    default:
      return next()
  }
})

/**
 * @swagger
 * /v1/votaciones/{idVotacion}:iniciar:
 *   post:
 *     summary: votacion.iniciar
 *     parameters:
 *       - name: idVotacion
 *         in: path
 *         required: true
 *         description: ID de votación.
 *     requestBody:
 *       description: "Info de la votación"
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             required:
 *               - duracionEnSegundos
 *             properties:
 *               duracionEnSegundos:
 *                 type: integer
 *                 minimum: 1
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Votaciones
 *
 * /v1/votaciones/{idVotacion}:cerrar:
 *   post:
 *     summary: votacion.cerrar
 *     parameters:
 *       - name: idVotacion
 *         in: path
 *         required: true
 *         description: ID de votación.
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Votaciones
 *
 * /v1/votaciones/{idVotacion}:cancelar:
 *   post:
 *     summary: votacion.cancelar
 *     parameters:
 *       - name: idVotacion
 *         in: path
 *         required: true
 *         description: ID de votación.
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Votaciones
*/
export const votacionActionSwitch = endpoint(async (req, res, next) => {
  const idVotacion = validarNormalizarIDVotacion(req.params.idVotacion)
  let votacion = req.app.locals.votaciones[idVotacion]
  const replier = req.app.locals.replier
  const commander = req.app.locals.commander
  if (req.params.action === 'iniciar' && !votacion) {
    // generamos la instancia de la votacion
    votacion = new Votacion(idVotacion, req.app.locals.votacionesPublisher, replier, commander)
    req.app.locals.votaciones[idVotacion] = votacion
  }
  if (!votacion) {
    throw new ClientError('votación no existe', 404)
  }

  switch (req.params.action) {
    case 'iniciar':
      if (typeof (req.body.duracionEnSegundos) === 'undefined') {
        throw new ClientError('duracionEnSegundos es requerido')
      }
      votacion.iniciarVotacion(req.body.duracionEnSegundos)
      break

    case 'cerrar':
      votacion.cerrarVotacion()
      break

    case 'cancelar':
      votacion.cancelarVotacion()
      break

    case 'unsetVoto':
      if (typeof (req.body.cuil) === 'undefined') {
        throw new ClientError('cuil es requerido')
      }
      votacion.unsetVoto(req.body.cuil)
      break

    default:
      return next()
  }
  res.sendStatus(200)
})

/**
 * @swagger
 * /v1/votaciones/{idVotacion}/votos/{cuil}:
 *   get:
 *     summary: votaciones.obtenerVoto
 *     parameters:
 *       - name: idVotacion
 *         in: path
 *         required: true
 *         description: ID de votación.
 *       - name: cuil
 *         in: path
 *         required: true
 *         description: Cuil del diputado.
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Votaciones
*/
export const obtenerVoto = endpoint(async (req, res) => {
  const idVotacion = validarNormalizarIDVotacion(req.params.idVotacion)
  const votacion = req.app.locals.votaciones[idVotacion]
  if (!votacion) {
    throw new ClientError('votación no existe', 404)
  }

  const cuil = validarNormalizarCuil(req.params.cuil)

  res.json({
    data: {
      voto: votacion.obtenerVoto(cuil)
    }
  })
})

/**
 * @swagger
 * /v1/votaciones/{idVotacion}/votos/{cuil}:
 *   put:
 *     summary: votaciones.upsertVoto
 *     parameters:
 *       - name: idVotacion
 *         in: path
 *         required: true
 *         description: ID de votación.
 *       - name: cuil
 *         in: path
 *         required: true
 *         description: Cuil del diputado.
 *     requestBody:
 *       description: "Voto emitido por diputado: afirmativo (0), negativo (1) ó abstención (2)"
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             required:
 *               - voto
 *             properties:
 *               voto:
 *                 type: integer
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Votaciones
*/
export const upsertVoto = endpoint(async (req, res) => {
  const idVotacion = validarNormalizarIDVotacion(req.params.idVotacion)
  const votacion = req.app.locals.votaciones[idVotacion]
  if (!votacion) {
    throw new ClientError('votación no existe', 404)
  }

  const cuil = validarNormalizarCuil(req.params.cuil)

  if (typeof (req.body.voto) === 'undefined') {
    throw new ClientError('voto es requerido')
  }
  votacion.setVoto(cuil, req.body.voto)
  res.sendStatus(200)
})

/**
 * @swagger
 * /v1/votaciones/{idVotacion}/votos/{cuil}:
 *   delete:
 *     summary: votaciones.deleteVoto
 *     parameters:
 *       - name: idVotacion
 *         in: path
 *         required: true
 *         description: ID de votación.
 *       - name: cuil
 *         in: path
 *         required: true
 *         description: Cuil.
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Votaciones
*/
export const deleteVoto = endpoint(async (req, res) => {
  const idVotacion = validarNormalizarIDVotacion(req.params.idVotacion)
  const votacion = req.app.locals.votaciones[idVotacion]
  if (!votacion) {
    throw new ClientError('votación no existe', 404)
  }

  const cuil = validarNormalizarCuil(req.params.cuil)

  votacion.unsetVoto(cuil)
  res.sendStatus(200)
})
