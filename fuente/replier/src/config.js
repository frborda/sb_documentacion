import { config } from 'dotenv'

// leer las variables de entorno
config()

// exportar configuracion de la app

// TODO validar configuracion

const env = process.env.ENV ?? 'desarrollo'
const debug = process.env.DEBUG ? process.env.DEBUG === 'true' : env !== 'produccion'
const host = process.env.SERVICE_HOST ? process.env.SERVICE_HOST : 'localhost'
const port = process.env.SERVICE_PORT ? parseInt(process.env.SERVICE_PORT) : 17001

export default {
  env,
  debug,
  service: {
    url: process.env.SERVICE_URL ? process.env.SERVICE_URL : `https://${host}:${port}`,
    host,
    port,
    ssl: {
      key: process.env.SERVICE_SSL_KEY ? process.env.SERVICE_SSL_KEY : '../_devenv/localhost.key',
      cert: process.env.SERVICE_SSL_CERT ? process.env.SERVICE_SSL_CERT : '../_devenv/localhost.crt',
      bundle: process.env.SERVICE_SSL_BUNDLE ? process.env.SERVICE_SSL_BUNDLE : null
    },
    apiKey: process.env.SERVICE_API_KEY,
    webhookApiKey: process.env.WEBHOOK_API_KEY,
    logRequests: process.env.SERVICE_LOG_REQUESTS === 'true'
  },
  recinto: {
    nroBancas: process.env.RECINTO_NRO_BANCAS ? parseInt(process.env.RECINTO_NRO_BANCAS) : 257,
    presidentePorDefecto: parseInt(process.env.RECINTO_PRESIDENTE_POR_DEFECTO),
    emulacionPorDefecto: process.env.RECINTO_EMULACION_POR_DEFECTO ? process.env.RECINTO_EMULACION_POR_DEFECTO === 'true' : debug,
    webhook: {
      url: process.env.RECINTO_WEBHOOK_URL,
      apiKey: process.env.RECINTO_WEBHOOK_API_KEY,
      // si 0 o no definido, envia solo cuando sucede un evento
      replicarCadaXMilisegundos: process.env.RECINTO_WEBHOOK_REPLY_MS ? parseInt(process.env.RECINTO_WEBHOOK_REPLY_MS) : 0
    },
    diputados: {
      url: process.env.RECINTO_DIPUTADOS_URL,
      apiKey: process.env.RECINTO_DIPUTADOS_API_KEY
    }
  },
  votaciones: {
    webhook: {
      url: process.env.VOTACIONES_WEBHOOK_URL,
      apiKey: process.env.VOTACIONES_WEBHOOK_API_KEY
    }
  }
}
