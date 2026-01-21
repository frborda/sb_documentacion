import { ClientError } from '../_core'

export class Votacion {
  constructor (idVotacion, votacionesPublisher, replier, commander) {
    this.idVotacion = idVotacion
    this.estado = 'CREADA' // CREADA / INICIADA / CERRADA / CANCELADA
    this.duracionEnSegundos = 0
    this.timestampInicio = null
    this.timestampFin = null
    this.votoPorCuil = {} // cuil => voto
    this._interval = null

    this.votacionesPublisher = votacionesPublisher
    this.replier = replier
    this.commander = commander
  }

  _log (mensaje) {
    console.log(`ts:${new Date().toISOString()} votacion:${this.idVotacion} estado:${this.estado}${mensaje ? ` ${mensaje}` : ''}`)
  }

  _resetInterval () {
    clearInterval(this._interval)
    this._interval = null
  }

  iniciarVotacion (duracionEnSegundos) {
    // validaciones
    if (this.estado === 'INICIADA') throw new ClientError('votacion ya se encuentra iniciada')
    if (this.estado === 'CERRADA') throw new ClientError('no se puede iniciar una votacion cerrada')
    if (duracionEnSegundos <= 0) throw new ClientError('duracion en segundos debe ser mayor a 0')

    // iniciar votacion
    this.estado = 'INICIADA'
    this.duracionEnSegundos = duracionEnSegundos
    this.timestampInicio = Date.now()
    this.timestampFin = this.timestampInicio + this.duracionEnSegundos * 1000
    this.votoPorCuil = {}

    if (this.commander.estaControlandoLasVotaciones()) {
      this._log('using commander')
      this.commander.controlarVotacion(this)
    } else {
      // iniciar timer
      this._interval = setInterval(() => {
        if (Date.now() < this.timestampFin) return
        this.cerrarVotacion()
      }, 1000)

      this._log()
    }
  }

  _publish (type, payload) {
    // TODO chequear si viene algun error o status code que tengamos que procesar
    this.votacionesPublisher.publish(type, payload)
  }

  cancelarVotacion () {
    // validaciones
    if (this.estado === 'CERRADA') throw new ClientError('no se puede cancelar una votacion cerrada')
    if (this.estado === 'CANCELADA') throw new ClientError('votacion se encuentra cancelada')

    // cancelar votacion
    this.estado = 'CANCELADA'
    this._resetInterval()
    this.votoPorCuil = {}

    this._log()
  }

  cerrarVotacion () {
    // validaciones
    if (this.estado !== 'INICIADA') throw new ClientError('solo se puede cerrar una votacion iniciada')

    // cerrar votacion
    this.estado = 'CERRADA'
    this._resetInterval()

    // antes de cerrar pausamos el replier pasa asegurar no haya cambios de asistencia concurrentes
    this.replier.pause(true)

    this._publish('CERRAR_VOTACION', this.serializeToSync())

    this._log()
  }

  obtenerVoto (cuil) {
    return this.votoPorCuil[cuil] ?? null
  }

  setVoto (cuil, voto, permiteTiposExtra = false) {
    // validaciones
    if (this.estado !== 'INICIADA') throw new ClientError('votacion no se encuentra iniciada')
    if (typeof (voto) !== 'number') {
      throw new ClientError('voto debe ser de tipo entero')
    }
    if (!permiteTiposExtra && voto !== 0 && voto !== 1 && voto !== 2) {
      throw new ClientError('voto debe ser: 0, 1 ó 2')
    }

    // setear voto
    this.votoPorCuil[cuil] = voto
  }

  unsetVoto (cuil) {
    // validaciones
    if (this.estado !== 'INICIADA') throw new ClientError('votacion no se encuentra iniciada')

    // borrar voto
    delete this.votoPorCuil[cuil]
  }

  toJSON () {
    const json = {
      idVotacion: this.idVotacion,
      estado: this.estado
    }
    if (this.estado === 'CREADA' || this.estado === 'CANCELADA') return json

    json.duracionEnSegundos = this.duracionEnSegundos
    const dateInicio = new Date(this.timestampInicio)
    json.timestampInicio = dateInicio.toISOString()
    const dateFin = new Date(this.timestampFin)
    json.timestampFin = dateFin.toISOString()

    const segundosRestantes = (this.timestampFin - Date.now()) / 1000
    json.segundosRestantes = segundosRestantes < 0 ? 0 : segundosRestantes

    json.votoPorCuil = Object.assign({}, this.votoPorCuil)

    return json
  }

  serializeToSync () {
    return {
      id: this.idVotacion,
      st: this.estado,
      vpc: Object.assign({}, this.votoPorCuil)
    }
  }
}
