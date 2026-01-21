import express from 'express'
import https from 'https'
import fs from 'fs'
import path from 'path'
import morgan from 'morgan'
import helmet from 'helmet'
import swaggerUI from 'swagger-ui-express'
import { throttle } from 'underscore'
import config from './config'
import appRoutes from './routes/app.routes'
import apiRoutes from './routes/api.routes'
import { Recinto } from './domain/Recinto'
import { EventEmitter, ClientError, swaggerSpec, websocket, WebhookPublisher } from './_core'
import { Replier } from './daemons/Replier'
// eslint-disable-next-line no-unused-vars
import { twig } from 'twig'
import { Commander } from './daemons/Commander'

console.log('Iniciando servicio en entorno', config.env)
console.log('Modo debug', config.debug ? 'activo' : 'inactivo')

// Inicializar aplicacion
const app = express()

// Configurar render engine
app.set('view engine', 'twig')
app.set('views', path.join(__dirname, 'views'))

app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'", "'unsafe-inline'"],
      scriptSrc: ["'self'", "'unsafe-inline'"],
      imgSrc: ["'self'", 'data:'],
      connectSrc: ['*']
    }
  }
}))

app.use('/assets', express.static(path.join(__dirname, 'public')))

if (config.service.logRequests) app.use(morgan('dev'))
app.use(express.json())
app.use(express.urlencoded({ extended: false }))

// Inicializar estados globales
app.locals.events = new EventEmitter()
app.locals.recinto = new Recinto(app.locals.events, config.recinto)
app.locals.votaciones = {} // idreunion_idvotacion => votacion

const votacionesPublisher = new WebhookPublisher('votaciones')
votacionesPublisher.subscribe(config.votaciones.webhook.url, config.votaciones.webhook.apiKey)
app.locals.votacionesPublisher = votacionesPublisher

const replier = new Replier(app.locals.events, app.locals.recinto, config.recinto.webhook.replicarCadaXMilisegundos)
replier.to(config.recinto.webhook.url, config.recinto.webhook.apiKey)
app.locals.replier = replier

const commander = new Commander(app.locals.recinto)
app.locals.commander = commander

// Declarar las rutas
app.use('/api/docs', swaggerUI.serve, swaggerUI.setup(swaggerSpec))
app.use('/api', apiRoutes)
app.use('/', appRoutes)

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
  console.log(`Servicio escuchando en ${config.service.host}:${config.service.port}`)
})

// Websockets
const websocketServer = websocket(server)
const broadcastRecintoState = throttle(() => {
  websocketServer.broadcast('hemiciclo', {
    bancas: app.locals.recinto.toJSON(),
    totalBasculasActivas: app.locals.recinto.nroBasculasActivas,
    totalIdentificaciones: app.locals.recinto.nroIdentificados
  })
}, 1000)
app.locals.events.subscribe(e => {
  broadcastRecintoState()
})

// Ejecutar replier daemon
replier.run()
