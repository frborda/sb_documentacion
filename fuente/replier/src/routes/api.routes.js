import { Router } from 'express'
import { requireServiceApiKey } from '../_core'
import * as apiV1 from '../controllers/api/v1'
import * as apiV2 from '../controllers/api/v2'
import * as ctrl from '../controllers/HealthcheckController'

const router = Router()

router.get('/healthcheck', requireServiceApiKey, ctrl.healthcheck)

router.post('/v1/replier::action', requireServiceApiKey, apiV1.replierActionSwitch)

router.get('/v1/recinto', requireServiceApiKey, apiV1.obtenerRecinto)
router.get('/v1/recinto::action', requireServiceApiKey, apiV1.recintoQuerySwitch)
router.post('/v1/recinto::action', requireServiceApiKey, apiV1.recintoActionSwitch)

router.get('/v1/votaciones', requireServiceApiKey, apiV1.listarVotaciones)
router.delete('/v1/votaciones', requireServiceApiKey, apiV1.limpiarVotaciones)
router.get('/v1/votaciones/:idVotacion::action', requireServiceApiKey, apiV1.votacionQuerySwitch)
router.get('/v1/votaciones/:idVotacion', requireServiceApiKey, apiV1.obtenerVotacion)
router.post('/v1/votaciones/:idVotacion::action', requireServiceApiKey, apiV1.votacionActionSwitch)

router.get('/v1/votaciones/:idVotacion/votos/:cuil', requireServiceApiKey, apiV1.obtenerVoto)
router.put('/v1/votaciones/:idVotacion/votos/:cuil', requireServiceApiKey, apiV1.upsertVoto)
router.delete('/v1/votaciones/:idVotacion/votos/:cuil', requireServiceApiKey, apiV1.deleteVoto)

router.get('/v1/bancas', requireServiceApiKey, apiV1.listarBancas)
router.get('/v1/bancas/:banca', requireServiceApiKey, apiV1.obtenerBanca)
router.put('/v1/bancas/:banca', requireServiceApiKey, apiV1.actualizarBanca)

router.post('/v1/webhook', apiV1.webhook) // webhook de entrada
router.post('/v1/webhook/debug', apiV1.recintoWebhookDebug) // webhook de salida para debug
router.post('/v2/webhook', apiV2.webhook) // webhook de entrada

router.get('/v1/debug/commander', requireServiceApiKey, apiV1.obtenerCommander)
router.post('/v1/debug/commander::action', requireServiceApiKey, apiV1.commanderActionSwitch)
router.post('/v1/debug/load-testing', requireServiceApiKey, apiV1.loadTesting)

export default router
