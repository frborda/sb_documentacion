import { Router } from 'express'
import * as ctrl from '../controllers/HemicicloController'

const router = Router()

router.get('/hemiciclo', ctrl.hemiciclo)

export default router
