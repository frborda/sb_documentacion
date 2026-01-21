import express from 'express'
import https from 'https'
import fs from 'fs'
import morgan from 'morgan'
import helmet from 'helmet'
import config from './config'
import swaggerUI from 'swagger-ui-express'
import apiRoutes from './routes/api.routes'
import { EventEmitter, ClientError, Database, swaggerSpec } from './_core'
import { BancasConsumerReplier } from './daemons/BancasConsumerReplier'

console.log('Iniciando servicio en entorno', config.env)
console.log('Modo debug', config.debug ? 'activo' : 'inactivo')

// Inicializar aplicacion
const app = express()

app.use(helmet())
if (config.debug) app.use(morgan('dev'))
app.use(express.json())
app.use(express.urlencoded({ extended: false }))

// Inicializar estados globales
app.locals.events = new EventEmitter()

// Declarar las rutas
app.use('/api/docs', swaggerUI.serve, swaggerUI.setup(swaggerSpec))
app.use('/api', apiRoutes)

app.use((req, res, next) => {
  // si no se encontro la ruta, respondemos con error 404
  next(new ClientError(`Recurso ${req.originalUrl} no encontrado`, 404))
})

// Gestion de errores
app.use((err, req, res, next) => {
  const code = err.statusCode ?? 500
  const metadata = err.metadata ?? {}
  if (config.debug) {
    metadata.stack = err.stack
  }
  const payload = {
    error: {
      code,
      message: err.message,
      metadata
    }
  }
  res.status(code).json(payload)
})

// Ejecutar aplicacion
const key = fs.readFileSync(config.service.ssl.key)
let cert = fs.readFileSync(config.service.ssl.cert)
if (config.service.ssl.bundle) {
  cert += fs.readFileSync(config.service.ssl.bundle)
}

const server = https.createServer({ key, cert }, app)
server.listen(config.service.port, config.service.host, () => {
  console.log(`Server listening on ${config.service.host}:${config.service.port}`)
})

// Ejecutar daemon
const db = new Database(config)
app.locals.bcr = new BancasConsumerReplier(db)
app.locals.bcr.to(config.webhook.url, config.webhook.apiKey)
app.locals.bcr.run()
