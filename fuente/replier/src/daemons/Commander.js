import config from '../config'
import { ClientError, http } from '../_core'

export class Commander {
  constructor (recinto) {
    this.recinto = recinto
    this.interval = null
    this.waitMilliseconds = 2000
    this.identificar = true
    this.cuiles = []
    this.bancasMetadata = []
    this.nroAcciones = 3
    if (this.recinto.nroBancas < this.nroAcciones) {
      this.nroAcciones = this.recinto.nroBancas
    }
    //
    this.stats = {}
    //
    this.emulacionVotacion = false
    this.indexResultadoAEmular = 0
    this.resultadosAEmular = []
  }

  toJSON () {
    let modo = 'offline'
    if (this.interval) {
      modo = 'actioner'
    } else if (this.emulacionVotacion) {
      modo = 'voter'
    }

    const json = { modo }

    switch (modo) {
      case 'actioner':
        json.waitMilliseconds = this.waitMilliseconds
        json.stats = Object.assign({}, this.stats)
        json.cuiles = [...this.cuiles]
        break
      case 'voter':
        json.resultadosAEmular = this.resultadosAEmular.length > 0 ? [...this.resultadosAEmular] : 'random'
        json.cuiles = [...this.cuiles]
        break
    }

    return json
  }

  async activarEmulacionVotacion () {
    if (!this.recinto.emulacion) throw new ClientError('Se requiere recinto en modo emulacion')
    if (this.emulacionVotacion === true) return false

    // cargar data de diputados
    if (this.identificar && this.cuiles.length === 0) await this.cargarDiputados()
    if (this.identificar === false) throw new ClientError('No se puede emular votacion sin data de diputados')

    // parar la emision de comandos
    this.stop()

    this.emulacionVotacion = true
    console.log('commander: emulacion votacion: activa')
    return true
  }

  cargarResultadosAEmular (resultados) {
    if (this.emulacionVotacion === false) throw new ClientError('Se requiere commander en modo emulacion votacion')

    const resultadosAEmular = []
    for (let i = 0; i < resultados.length; i++) {
      const r = resultados[i]

      const shuffle = r.shuffle === true

      const afirmativos = parseInt(r.afirmativos ?? 0)
      if (afirmativos < 0) throw new ClientError('Valor para afirmativos invalido:', r.afirmativos)

      const negativos = parseInt(r.negativos ?? 0)
      if (negativos < 0) throw new ClientError('Valor para negativos invalido:', r.negativos)

      const abstenciones = parseInt(r.abstenciones ?? 0)
      if (abstenciones < 0) throw new ClientError('Valor para abstenciones invalido:', r.abstenciones)

      const presentesSinVotar = parseInt(r.presentesSinVotar ?? 0)
      if (presentesSinVotar < 0) throw new ClientError('Valor para presentesSinVotar invalido:', r.presentesSinVotar)

      const presentes = afirmativos + negativos + abstenciones + presentesSinVotar
      if (presentes > 256) throw new ClientError('La suma de afirmativos, negativos, abstenciones y presentes sin votar no puede exceder 256 (al presidente lo suma commander como un presente sin votar)')

      resultadosAEmular.push({ shuffle, afirmativos, negativos, abstenciones, presentesSinVotar, ausentes: 256 - presentes })
    }

    this.resultadosAEmular = resultadosAEmular
    this.indexResultadoAEmular = 0
    console.log('commander: se cargaron resultados de votacion para emular:', this.resultadosAEmular.length)

    return true
  }

  _limpiarResultadosAEmular () {
    if (this.resultadosAEmular.length === 0) return false

    this.resultadosAEmular = []

    console.log('commander: se limpiaron los resultados de votacion a emular')
    return true
  }

  desactivarEmulacionVotacion () {
    if (this.emulacionVotacion === false) return false
    this.emulacionVotacion = false
    console.log('commander: emulacion votacion: inactiva')
    this._limpiarResultadosAEmular()
    return true
  }

  estaControlandoLasVotaciones () {
    if (this.emulacionVotacion === false) return false
    if (!this.recinto.emulacion) {
      this.desactivarEmulacionVotacion()
      return false
    }
    return true
  }

  _obtenerSiguienteResultadoAEmular () {
    if (this.resultadosAEmular.length > 0) {
      const resultado = this.resultadosAEmular[this.indexResultadoAEmular]
      this.indexResultadoAEmular = (this.indexResultadoAEmular + 1) % this.resultadosAEmular.length
      return resultado
    } else {
      let max = this.cuiles.length // se excluye siempre al presidente asi que como maximo 256

      const afirmativos = Math.ceil(Math.random() * max)
      max -= afirmativos
      if (max < 0) max = 0

      const negativos = Math.ceil(Math.random() * max)
      max -= negativos
      if (max < 0) max = 0

      const abstenciones = Math.ceil(Math.random() * max)
      max -= abstenciones
      if (max < 0) max = 0

      const presentesSinVotar = Math.ceil(Math.random() * max)
      max -= presentesSinVotar
      if (max < 0) max = 0

      return { afirmativos, negativos, abstenciones, presentesSinVotar, shuffle: true }
    }
  }

  async _sleep (ms) {
    return await new Promise(resolve => setTimeout(resolve, ms))
  }

  async controlarVotacion (votacion) {
    console.log('commander: controla la votacion', votacion.idVotacion)

    const resultado = this._obtenerSiguienteResultadoAEmular()

    const cuiles = [...this.cuiles]
    if (resultado.shuffle) {
      this.shuffle(cuiles)
    }
    let cuilIndex = 0

    for (let i = 0; i < resultado.afirmativos && cuilIndex < cuiles.length; i++) {
      votacion.setVoto(cuiles[cuilIndex], 0, true) // afirmativo
      cuilIndex++
    }
    for (let i = 0; i < resultado.negativos && cuilIndex < cuiles.length; i++) {
      votacion.setVoto(cuiles[cuilIndex], 1, true) // negativo
      cuilIndex++
    }
    for (let i = 0; i < resultado.abstenciones && cuilIndex < cuiles.length; i++) {
      votacion.setVoto(cuiles[cuilIndex], 2, true) // abstencion
      cuilIndex++
    }
    for (let i = 0; i < resultado.presentesSinVotar && cuilIndex < cuiles.length; i++) {
      votacion.setVoto(cuiles[cuilIndex], 4, true) // presente sin votar
      cuilIndex++
    }
    while (cuilIndex < cuiles.length) {
      votacion.setVoto(cuiles[cuilIndex], 3, true) // ausente
      cuilIndex++
    }

    // presidente siempre como presente sin votar
    votacion.setVoto(config.recinto.presidentePorDefecto, 4, true) // presente sin votar

    // esperar para que recinto termine de iniciar la votacion
    await this._sleep(2000) // 2 seconds

    console.log('commander: cerrara la votacion', votacion.idVotacion)
    votacion.cerrarVotacion()
  }

  action () {
    if (!this.recinto.emulacion) {
      this.stop()
      return
    }

    const bancaMetadataIndex = this.stats.comandos % this.bancasMetadata.length
    if (bancaMetadataIndex < this.nroAcciones) {
      this.shuffle(this.bancasMetadata)
    }
    for (let i = 0; i < this.nroAcciones && bancaMetadataIndex + i < this.bancasMetadata.length; i++) {
      const bm = this.bancasMetadata[bancaMetadataIndex + i]
      const b = this.recinto.obtenerBanca(bm.numero)

      if (!b.bascula) {
        // 1. sentar
        this.recinto.activarBascula(bm.numero)
      } else if (bm.diputadoAsignado && !b.identificacion) {
        // 2. identificar
        this.recinto.identificar(bm.numero, bm.diputadoAsignado)
      } else if (bm.numero > 14) {
        // 3. pararse
        // evitamos que las primeras 14 cambien tras haberse identificado
        this.recinto.desactivarBascula(bm.numero)
      }
      this.stats.comandos++
    }
  }

  async start () {
    if (!this.recinto.emulacion) throw new ClientError('Se requiere recinto en modo emulacion')
    if (this.interval) return

    this.desactivarEmulacionVotacion()

    if (this.identificar && this.cuiles.length === 0) await this.cargarDiputados()
    this.prepararBancasMetadata()

    console.log('commander: start')
    this.stats.comandos = 0
    this.stats.startTime = Date.now()
    this.interval = setInterval(() => {
      this.action()
    }, this.waitMilliseconds)
  }

  stop () {
    if (!this.interval) return
    clearInterval(this.interval)
    this.interval = null
    this.stats.endTime = Date.now()
    console.log('commander: stop')
    console.log('commander: stats:', this.stats)
  }

  shuffle (array) {
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

  prepararBancasMetadata () {
    this.bancasMetadata = []
    const cuiles = [...this.cuiles]
    this.shuffle(cuiles)
    for (let i = 1; i < this.recinto.nroBancas; i++) {
      this.bancasMetadata.push({
        numero: i,
        diputadoAsignado: cuiles.pop() ?? null
      })
    }
  }

  async cargarDiputados () {
    if (!config.recinto.diputados.url) {
      this.identificar = false
    }
    console.log('commander: obteniendo cuiles de diputados')

    const diputados = await http.get(config.recinto.diputados.url, {
      timeout: 1000, // 1s
      headers: {
        'Content-Type': 'application/json',
        'X-API-KEY': config.recinto.diputados.apiKey
      }
    })

    diputados.data.forEach(d => {
      const cuil = parseInt(d.CUIL)
      if (!cuil || cuil === config.recinto.presidentePorDefecto) return
      this.cuiles.push(cuil)
    })
    this.identificar = true

    console.log(`commander: se obtuvieron ${this.cuiles.length} cuiles`)
  }
}
