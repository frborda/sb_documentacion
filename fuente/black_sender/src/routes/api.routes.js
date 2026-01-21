import { Router } from 'express'
import { requireServiceApiKey } from '../_core'
import * as ctrl from '../controllers/HealthcheckController'
import * as apiV1 from '../controllers/api/v1'

const router = Router()

router.get('/healthcheck', requireServiceApiKey, ctrl.healthcheck)

router.get('/v1/daemon', requireServiceApiKey, apiV1.daemonEstado)
router.post('/v1/daemon::action', requireServiceApiKey, apiV1.daemonActionSwitch)

export default router
