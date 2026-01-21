import { Banca } from './Banca'
import config from '../config'
import { ClientError } from '../_core'

export class Recinto {
  constructor (events, { nroBancas, presidentePorDefecto, emulacionPorDefecto }) {
    if (nroBancas < 1) throw new Error('nroBancas debe ser mayor o igual a 1')
    this.events = events
    this.nroBancas = nroBancas
    this.nroIdentificados = 0
    this.nroBasculasActivas = 0
    this.bancas = new Array(nroBancas)
    for (let i = 0; i < nroBancas; i++) {
      this.bancas[i] = new Banca(i)
    }
    this.bancaPorCuil = {} // cuil identificado => nro banca

    console.log(`Recinto con ${nroBancas} bancas`)

    this.emulacion = false
    if (emulacionPorDefecto) {
      try {
        this.activarEmulacion()
      } catch (err) {
        console.log('Emulacion inactiva:', err.message)
      }
    } else {
      this.desactivarEmulacion()
    }

    this.presidentePorDefecto = presidentePorDefecto

    // Bascula 0 siempre esta activa y no puede quedar sin identificacion
    this.activarBascula(0)
    this.identificar(0, this.presidentePorDefecto)

    this._updateLastModified()
  }

  activarEmulacion () {
    if (config.debug === false) throw new ClientError('Solo se puede emular en modo debug')
    if (this.emulation === true) return false
    this.emulacion = true
    console.log('Emulacion activa')
    return true
  }

  desactivarEmulacion () {
    if (this.emulacion === false) return false
    this.emulacion = false
    console.log('Emulacion inactiva')
    return true
  }

  obtenerBanca (bancaNro) {
    bancaNro = parseInt(bancaNro)
    if (bancaNro < 0 || bancaNro >= this.bancas.length) {
      throw new ClientError('Banca no encontrada', 404)
    }
    return this.bancas[bancaNro]
  }

  activarBascula (bancaNro) {
    const banca = this.obtenerBanca(bancaNro)

    const ok = banca.activarBascula()
    if (ok === false) return false

    this.nroBasculasActivas++
    this.events.emit('banca.bascula.activa', { numero: banca.numero })

    this._updateLastModified()
    return true
  }

  desactivarBascula (bancaNro) {
    const banca = this.obtenerBanca(bancaNro)

    if (banca.numero === 0) {
      // throw new ClientError('No se puede desactivar la báscula 0')
      return false
    }

    // mantener identificacion en null mientras bascula esta inactiva
    this.desidentificar(banca.numero)

    const ok = banca.desactivarBascula()
    if (ok === false) return false

    this.nroBasculasActivas--
    this.events.emit('banca.bascula.inactiva', { numero: banca.numero })

    this._updateLastModified()
    return true
  }

  identificar (bancaNro, cuil) {
    const banca = this.obtenerBanca(bancaNro)
    cuil = parseInt(cuil)

    if (!cuil) throw new ClientError('Cuil inválido')
    if (!banca.bascula) throw new ClientError(`Bascula '${banca.numero}' debe estar activa para poder asignar identificación`)

    if (banca.identificacion === cuil) {
      // cuil ya estaba identificado en esa banca
      return false
    }

    // validación del presidente
    if (cuil === this.presidentePorDefecto && banca.numero !== 0) {
      const bancaPresidente = this.bancas[0]
      if (bancaPresidente.identificacion === this.presidentePorDefecto) {
        throw new ClientError('Presidente solo se puede identificar en otra banca si alguien ocupa la banca 0')
      }
    }

    // en caso de que este identificado en otra banca, lo desidentificamos
    this.desidentificarCuil(cuil)

    const cuilPrevio = banca.identificacion

    const ok = banca.identificar(cuil)
    if (ok === false) return false

    this.bancaPorCuil[cuil] = banca.numero

    if (cuilPrevio) {
      delete this.bancaPorCuil[cuilPrevio]
    } else {
      this.nroIdentificados++
    }

    this.events.emit('banca.identificacion', { numero: banca.numero, identificacion: cuil })

    this._updateLastModified()
    return true
  }

  desidentificar (bancaNro) {
    const banca = this.obtenerBanca(bancaNro)

    // validación del presidente
    if (banca.numero === 0) {
      throw new ClientError('No se puede desidentificar al presidente')
    }

    const cuilDesidentificado = banca.desidentificar()
    if (cuilDesidentificado === false) return false

    delete this.bancaPorCuil[cuilDesidentificado]
    this.nroIdentificados--
    this.events.emit('banca.desidentificacion', { numero: banca.numero, identificacion: cuilDesidentificado })

    this._updateLastModified()
    return cuilDesidentificado
  }

  desidentificarCuil = (cuil) => {
    cuil = parseInt(cuil)
    const nroBanca = this.bancaPorCuil[cuil] ?? -1
    if (nroBanca < 0) return false

    // validación del presidente
    if (nroBanca === 0) {
      if (cuil === this.presidentePorDefecto) {
        throw new ClientError('Presidente no se puede desidentificar al menos que alguien ocupe la banca 0')
      }

      // identificamos a presidente por defecto en la banca 0
      this.identificar(nroBanca, this.presidentePorDefecto)
      return nroBanca
    }

    this.desidentificar(nroBanca)
    return nroBanca
  }

  limpiarIdentificaciones () {
    if (this.emulacion === false) throw new ClientError('Solo se puede limpiar las identificaciones en modo emulación')
    // ignoramos la banca del presidente
    for (let b = 1; b < this.nroBancas; b++) {
      this.desidentificar(b)
    }
  }

  toJSON () {
    return {
      bancas: this.bancas.map(b => b.toJSON())
    }
  }

  serializeToSync () {
    return this.bancas.map(b => ({
      b: b.bascula ? 1 : 0,
      i: b.identificacion ? b.identificacion : 0
    }))
  }

  _updateLastModified () {
    this.lastModified = new Date().toISOString()
  }
}
