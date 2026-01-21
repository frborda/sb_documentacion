import { throttle } from 'underscore'
import { ServerError, http } from '../_core'

export class BancasConsumerReplier {
  constructor (db) {
    this.db = db
    this.webhooks = []
    this.pollingMilliseconds = 1000 // 1s

    this.cuilPorId = {}
    this.vectorPresenciaAnterior = ''
    this.vectorIdentificacionAnterior = ''

    // por ahora mantenemos replicando los mensajes porque no hay forma
    // externa de acceder a este servicio para solicitarle los datos
    this.keepReplying = true
  }

  async run () {
    try {
      await this.db.conectar()
      await this.generarMapaCuilPorId()
      this.consume()
    } catch (err) {
      console.log('no se pudo ejecutar consumer de bancas:', err.message)
    }
  }

  async generarMapaCuilPorId () {
    await this.chequearCuilesDiputados()

    const rows = await this.db.query(`
      select id, cuil
      from dbo.DiputadosCuil
    `)

    for (const diputado of rows) {
      this.cuilPorId[diputado.id] = diputado.cuil
    }

    console.log('mapa cuil por id generado:', rows.length, 'registros')
  }

  async chequearCuilesDiputados () {
    const rows = await this.db.query(`
      select nombre, apellido
      from dbo.legisladores_activos
      where id not in (
        select id from DiputadosCuil
      )
    `)
    if (rows.length > 0) {
      throw new ServerError('Uno o más diputados activos sin cuil establecido', { diputados: rows })
    }
  }

  consume () {
    this.interval = setInterval(async () => {
      const [vector] = await this.db.query(`
        select vector_presencia, vector_identificacion
        from vector
      `)

      // la banca 0 siempre viene con sit_on
      vector.vector_presencia = '1' + vector.vector_presencia.substring(1)

      let dirty = false
      if (vector.vector_presencia !== this.vectorPresenciaAnterior) {
        dirty = true
        this.vectorPresenciaAnterior = vector.vector_presencia
        this.cacheBasculas = vector.vector_presencia
      }
      if (vector.vector_identificacion !== this.vectorIdentificacionAnterior) {
        dirty = true
        this.vectorIdentificacionAnterior = vector.vector_identificacion
        this.cacheIdentificaciones = this.parseIdentificaciones(vector.vector_identificacion)
      }

      if (dirty || this.keepReplying) {
        this.reply()
      }
    }, this.pollingMilliseconds)
  }

  parseIdentificaciones (identificaciones) {
    return identificaciones
      .split(';')
      .map(i => i === '0' ? '0' : this.cuilPorId[i])
      .join(';')
  }

  to (url, apiKey) {
    this.webhooks.push({ url, apiKey })
    console.log(`Replier suscribed webhook: ${url}`)
  }

  reply = throttle(() => {
    for (const w of this.webhooks) {
      http.post(w.url, {
        basculas: this.cacheBasculas,
        identificaciones: this.cacheIdentificaciones
      }, {
        timeout: 5000,
        headers: {
          'Content-Type': 'application/json',
          'X-API-KEY': w.apiKey
        }
      })
        .then(res => {
          const json = res.data
          if (json.totalFallidos > 0) {
            console.log(`Reply with ${json.totalFallidos} fail/s to ${w.url}:`, json.fallidos)
          }
        })
        .catch(err => {
          const message = err.response?.data?.error?.message ?? err.message
          console.log(`Error replying to ${w.url}:`, message)
        })
    }
  }, this.throttleMilliseconds)

  async regenerarCuiles () {
    await this.generarMapaCuilPorId()
  }

  toJSON () {
    return {
      pollingMilliseconds: this.pollingMilliseconds,
      vectorBasculas: this.cacheBasculas,
      vectorIdentificaciones: this.cacheIdentificaciones,
      cuilPorId: this.cuilPorId
    }
  }
}
