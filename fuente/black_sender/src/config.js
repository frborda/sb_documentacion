import { config } from 'dotenv'

// leer las variables de entorno
config()

const env = process.env.ENV ?? 'desarrollo'
const debug = process.env.DEBUG ? process.env.DEBUG === 'true' : env !== 'produccion'
const host = process.env.SERVICE_HOST ? process.env.SERVICE_HOST : 'localhost'
const port = process.env.SERVICE_PORT ? parseInt(process.env.SERVICE_PORT) : 17000

// exportar configuracion de la app
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
    apiKey: process.env.SERVICE_API_KEY
  },
  webhook: {
    url: process.env.SENDER_WEBHOOK_URL,
    apiKey: process.env.SENDER_WEBHOOK_API_KEY
  },
  database: {
    host: process.env.DATABASE_HOST,
    port: process.env.DATABASE_PORT ? parseInt(process.env.DATABASE_PORT) : 1433,
    name: process.env.DATABASE_NAME,
    user: process.env.DATABASE_USER,
    pass: process.env.DATABASE_PASS
  }
}
