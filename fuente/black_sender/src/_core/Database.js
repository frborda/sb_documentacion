import sql from 'mssql'
import { ServerError } from './errors'

export class Database {
  constructor (config) {
    this.config = {
      server: config.database.host,
      port: config.database.port,
      database: config.database.name,
      user: config.database.user,
      // password: config.database.pass,
      authentication: { type: 'default', options: { userName: config.database.user, password: config.database.pass } },
      options: {
        encrypt: false
        // trustServerCertificate: config.env !== 'produccion'
      }
    }
  }

  async conectar () {
    try {
      await sql.connect(this.config)
      await this.ping()
      console.log('conectado a la base de datos')
    } catch (err) {
      console.log('error al conectar', err.message)
    }
  }

  async ping () {
    try {
      await sql.query`select 1 as t`
    } catch (err) {
      throw ServerError('no se pudo realizar ping:', err.message)
    }
    return true
  }

  async query (q) {
    const result = await sql.query(q)
    return result.recordset
  }
}
