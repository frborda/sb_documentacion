import { ClientError, endpoint } from '../../../_core'

/**
 * @swagger
 * /v1/bancas:
 *   get:
 *     summary: bancas.listar
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Bancas
*/
export const listarBancas = endpoint(async (req, res) => {
  const recinto = req.app.locals.recinto

  res.json({
    totalBancas: recinto.nroBancas,
    totalBasculasActivas: recinto.nroBasculasActivas,
    totalIdentificados: recinto.nroIdentificados,
    lastModified: recinto.lastModified,
    data: recinto.bancas.map(b => b.toJSON())
  })
})

/**
 * @swagger
 * /v1/bancas/{banca}:
 *   get:
 *     summary: bancas.obtener
 *     parameters:
 *       - name: banca
 *         in: path
 *         required: true
 *         description: Número de banca.
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Bancas
*/
export const obtenerBanca = endpoint(async (req, res) => {
  const recinto = req.app.locals.recinto
  const bancaNro = req.params.banca
  const banca = recinto.obtenerBanca(bancaNro)

  res.json({
    lastModified: banca.lastModified,
    data: banca
  })
})

/**
 * @swagger
 * /v1/bancas/{banca}:
 *   put:
 *     summary: bancas.actualizar
 *     parameters:
 *       - name: banca
 *         in: path
 *         required: true
 *         description: Número de banca.
 *     requestBody:
 *       description: "Estado de banca"
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               bascula:
 *                 type: boolean
 *               identificacion:
 *                 type: integer
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Bancas
*/
export const actualizarBanca = endpoint(async (req, res) => {
  const recinto = req.app.locals.recinto
  if (recinto.emulacion === false) {
    throw new ClientError('La emulación de recinto no se encuentra activa')
  }
  const bancaNro = req.params.banca
  const banca = recinto.obtenerBanca(bancaNro)

  if (typeof (req.body.bascula) !== 'undefined') {
    if (req.body.bascula === true) {
      recinto.activarBascula(bancaNro)
    } else {
      recinto.desactivarBascula(bancaNro)
    }
  }

  if (typeof (req.body.identificacion) !== 'undefined') {
    if (req.body.identificacion) {
      recinto.identificar(bancaNro, req.body.identificacion)
    } else {
      recinto.desidentificar(bancaNro)
    }
  }

  res.json({
    lastModified: banca.lastModified,
    data: banca
  })
})
