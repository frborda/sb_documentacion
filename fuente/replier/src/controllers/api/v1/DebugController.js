import { isArray } from 'underscore'
import { ClientError, endpoint, http } from '../../../_core'
import config from '../../../config'

/**
 * @swagger
 * /v1/debug/commander:start:
 *   post:
 *     summary: debug.commander.start
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Debug
 *
 * /v1/debug/commander:stop:
 *   post:
 *     summary: debug.commander.stop
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Debug
 *
 * /v1/debug/commander:start-voter:
 *   post:
 *     summary: debug.commander.startVoter
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Debug
 *
 * /v1/debug/commander:stop-voter:
 *   post:
 *     summary: debug.commander.stopVoter
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Debug
 *
 * /v1/debug/commander:load-voter:
 *   post:
 *     summary: debug.commander.loadVoter
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Debug
*/
export const commanderActionSwitch = endpoint(async (req, res, next) => {
  const commander = req.app.locals.commander

  switch (req.params.action) {
    case 'start':
      await commander.start()
      break

    case 'stop':
      commander.stop()
      break

    case 'start-voter':
      await commander.activarEmulacionVotacion()
      break

    case 'stop-voter':
      commander.desactivarEmulacionVotacion()
      break

    case 'load-voter':
      if (!isArray(req.body) || req.body.length === 0) {
        throw new ClientError('se requiere arreglo de resultados con al menos un resultado')
      }
      await commander.cargarResultadosAEmular(req.body)
      break

    default:
      return next()
  }
  res.sendStatus(200)
})

/**
 * @swagger
 * /v1/debug/commander:
 *   get:
 *     summary: commander.obtener
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Debug
*/
export const obtenerCommander = endpoint(async (req, res, next) => {
  const commander = req.app.locals.commander

  res.json({
    totalCuiles: commander.cuiles.length,
    cuilPresidente: config.recinto.presidentePorDefecto,
    data: commander
  })
})

/**
 * @swagger
 * /v1/debug/load-testing:
 *   post:
 *     summary: debug.loadTesting
 *     responses: {}
 *     security:
 *       - ServiceApiKey: []
 *     tags:
 *       - Debug
*/
export const loadTesting = endpoint(async (req, res, next) => {
  const recinto = req.app.locals.recinto
  if (!recinto.emulacion) throw new ClientError('Se requiere recinto en modo emulacion')

  // 1 ~ 40s
  // 90 ~ 1h
  // 2160 ~ 24h
  let times = 1;
  if (typeof (req.body.times) === 'number') {
    times = req.body.times;
  }
  if (times < 1) {
    times = 1;
  } else if (times > 4000) {
    times = 4000;
  }

  const startTime = Date.now()
  console.log('load testing: start')

  const cuiles = []
  if (config.recinto.diputados.url) {
    console.log('load testing: obteniendo cuiles de diputados')

    const diputados = await http.get(config.recinto.diputados.url, {
      timeout: 3000, // 3s
      headers: {
        'Content-Type': 'application/json',
        'X-API-KEY': config.recinto.diputados.apiKey
      }
    })

    diputados.data.forEach(d => {
      const cuil = parseInt(d.CUIL)
      if (!cuil || cuil === config.recinto.presidentePorDefecto) return
      cuiles.push(cuil)
    })

    console.log(`load testing: se obtuvieron ${cuiles.length} cuiles`)
  }

  const total = recinto.nroBancas

  res.sendStatus(200)

  let times_executed = 0;
  const sm_wait = 10;
  const md_wait = 1000;
  const lg_wait = 1200;
  while (recinto.emulacion && times_executed < times) {
    try {
      console.log('load testing: identificar a todos, levantar')
      for (let i = 1; i < total; i++) {
        recinto.activarBascula(i)
        if (i < cuiles.length) recinto.identificar(i, cuiles[i])
        await sleep(sm_wait)
      }
      await sleep(md_wait)
      for (let i = 1; i < total; i++) {
        recinto.desactivarBascula(i)
        await sleep(sm_wait)
      }

      console.log('load testing: wait seconds', lg_wait / 1000)
      await sleep(lg_wait)

      console.log('load testing: identificar random, desidentificar, levantar')
      for (let i = 1; i < total; i++) {
        recinto.activarBascula(i)
        if (cuiles.length > 0) recinto.identificar(i, randomCuil(cuiles))
        await sleep(sm_wait)
      }
      await sleep(md_wait)
      for (let i = 1; i < total; i++) {
        recinto.desidentificar(i)
        await sleep(sm_wait)
      }
      await sleep(md_wait)
      for (let i = 1; i < total; i++) {
        recinto.desactivarBascula(i)
        await sleep(sm_wait)
      }

      if (cuiles.length > 0) {
        console.log('load testing: wait seconds', lg_wait / 1000)
      await sleep(lg_wait)

        console.log('load testing: party mode')
        for (let i = 1; i < total; i++) {
          recinto.activarBascula(i)
          if (i < cuiles.length) recinto.identificar(i, cuiles[i])
          await sleep(sm_wait)
        }

        const totalPartyMoves = 4
        for (let i = 0; i < totalPartyMoves; i++) {
          shuffle(cuiles)
          for (let i = 1; i < total; i++) {
            if (i < cuiles.length) recinto.identificar(i, cuiles[i])
            await sleep(sm_wait)
          }
          await sleep(md_wait)
        }

        for (let i = 1; i < total; i++) {
          recinto.desactivarBascula(i)
          await sleep(sm_wait)
        }
      }

      console.log('load testing: end', (Date.now() - startTime) / 1000)
    } catch (error) {
      console.log('load testing: error:', error)
      break;
    }
    times_executed++;
  }
})

function sleep (ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms)
  })
}

function randomFloat (min, max) {
  return Math.random() * (max - min) + min
}

function randomInt (min, max) {
  return Math.floor(randomFloat(min, max))
}

function randomCuil (cuiles) {
  return cuiles[randomInt(0, cuiles.length)]
}

function shuffle (array) {
  let currentIndex = array.length; let randomIndex

  // While there remain elements to shuffle.
  while (currentIndex !== 0) {
    // Pick a remaining element.
    randomIndex = Math.floor(Math.random() * currentIndex)
    currentIndex--;

    // And swap it with the current element.
    [array[currentIndex], array[randomIndex]] = [
      array[randomIndex], array[currentIndex]]
  }

  return array
}
